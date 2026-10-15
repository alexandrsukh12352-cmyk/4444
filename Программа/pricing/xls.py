"""Чтение старого формата Excel (.xls, BIFF8) без сторонних модулей.

Поддерживается то, что встречается в прайсах и выгрузках: текст (общая таблица строк
и LABEL), числа (NUMBER, RK, MULRK), закэшированные результаты формул, логические
значения и объединённые ячейки. Оформление, даты как даты и диаграммы не читаются.
"""

from __future__ import annotations

import struct
from pathlib import Path

SIGNATURE = bytes.fromhex("d0cf11e0a1b11ae1")
END_OF_CHAIN = 0xFFFFFFFE
FREE = 0xFFFFFFFF


class XlsError(Exception):
    pass


# ---------------------------------------------------------------- контейнер OLE2


def _ole_stream(data: bytes, names=("Workbook", "Book")) -> bytes:
    if data[:8] != SIGNATURE:
        raise XlsError("это не файл Excel 97–2003")
    sector_size = 1 << struct.unpack_from("<H", data, 0x1E)[0]
    mini_size = 1 << struct.unpack_from("<H", data, 0x20)[0]
    first_dir, = struct.unpack_from("<I", data, 0x30)
    mini_cutoff, first_minifat, n_minifat, first_difat, n_difat = struct.unpack_from("<IIIII", data, 0x38)

    def sector(sid: int) -> bytes:
        off = (sid + 1) * sector_size
        return data[off : off + sector_size]

    difat = list(struct.unpack_from("<109I", data, 0x4C))
    sid = first_difat
    for _ in range(n_difat):
        if sid in (END_OF_CHAIN, FREE):
            break
        block = struct.unpack(f"<{sector_size // 4}I", sector(sid))
        difat += block[:-1]
        sid = block[-1]
    fat: list[int] = []
    for s in difat:
        if s in (END_OF_CHAIN, FREE):
            continue
        fat += struct.unpack(f"<{sector_size // 4}I", sector(s))

    def chain(start: int, table: list[int]) -> list[int]:
        out, seen = [], set()
        while start not in (END_OF_CHAIN, FREE) and start < len(table) and start not in seen:
            seen.add(start)
            out.append(start)
            start = table[start]
        return out

    directory = b"".join(sector(s) for s in chain(first_dir, fat))
    entries = []
    for off in range(0, len(directory), 128):
        e = directory[off : off + 128]
        name_len, = struct.unpack_from("<H", e, 64)
        name = e[: max(0, name_len - 2)].decode("utf-16-le", "ignore")
        etype = e[66]
        start, size = struct.unpack_from("<II", e, 116)
        entries.append((name, etype, start, size))
    root = next((e for e in entries if e[1] == 5), None)

    for wanted in names:
        for name, etype, start, size in entries:
            if etype != 2 or name != wanted:
                continue
            if size >= mini_cutoff or root is None:
                return b"".join(sector(s) for s in chain(start, fat))[:size]
            minifat: list[int] = []
            for s in chain(first_minifat, fat):
                minifat += struct.unpack(f"<{sector_size // 4}I", sector(s))
            mini_stream = b"".join(sector(s) for s in chain(root[2], fat))
            return b"".join(mini_stream[s * mini_size : (s + 1) * mini_size] for s in chain(start, minifat))[:size]
    raise XlsError("в файле нет книги Excel")


# ---------------------------------------------------------------- записи BIFF


def _records(stream: bytes):
    pos, n = 0, len(stream)
    while pos + 4 <= n:
        rid, length = struct.unpack_from("<HH", stream, pos)
        yield pos, rid, stream[pos + 4 : pos + 4 + length]
        pos += 4 + length


def _rk(value: int) -> float:
    if value & 2:
        num = float(value >> 2 if value < 0x80000000 else (value >> 2) - (1 << 30))
    else:
        num = struct.unpack("<d", struct.pack("<Q", (value & 0xFFFFFFFC) << 32))[0]
    return num / 100 if value & 1 else num


def _short_string(buf: bytes, pos: int, encoding: str, biff8: bool, wide_len: bool) -> str:
    if wide_len:
        cch, = struct.unpack_from("<H", buf, pos)
        pos += 2
    else:
        cch = buf[pos]
        pos += 1
    if not biff8:
        return buf[pos : pos + cch].decode(encoding, "replace")
    flags = buf[pos]
    pos += 1
    if flags & 0x01:
        return buf[pos : pos + 2 * cch].decode("utf-16-le", "replace")
    return buf[pos : pos + cch].decode("latin-1")


def _read_sst(parts: list[bytes]) -> list[str]:
    """Общая таблица строк: строки могут переходить из записи SST в CONTINUE посреди символов."""
    strings: list[str] = []
    part_idx, buf = 0, parts[0]
    total, unique = struct.unpack_from("<II", buf, 0)
    pos = 8

    def next_part():
        nonlocal part_idx, buf, pos
        part_idx += 1
        if part_idx >= len(parts):
            raise IndexError
        buf, pos = parts[part_idx], 0

    try:
        for _ in range(unique):
            if pos >= len(buf):
                next_part()
            cch, = struct.unpack_from("<H", buf, pos)
            flags = buf[pos + 2]
            pos += 3
            runs = ext = 0
            if flags & 0x08:
                runs, = struct.unpack_from("<H", buf, pos)
                pos += 2
            if flags & 0x04:
                ext, = struct.unpack_from("<I", buf, pos)
                pos += 4
            wide = flags & 0x01
            chunks, left = [], cch
            while left:
                if pos >= len(buf):
                    next_part()
                    wide = buf[0] & 0x01
                    pos = 1
                avail = (len(buf) - pos) // (2 if wide else 1)
                take = min(left, avail)
                raw = buf[pos : pos + take * (2 if wide else 1)]
                chunks.append(raw.decode("utf-16-le", "replace") if wide else raw.decode("latin-1"))
                pos += take * (2 if wide else 1)
                left -= take
            strings.append("".join(chunks))
            skip = runs * 4 + ext
            while skip:
                if pos >= len(buf):
                    next_part()
                step = min(skip, len(buf) - pos)
                pos += step
                skip -= step
    except (IndexError, struct.error):
        pass
    return strings


CODEPAGES = {1251: "cp1251", 1252: "cp1252", 1200: "utf-16-le", 65001: "utf-8", 866: "cp866"}


def read_xls(path: Path) -> list[tuple[str, list[list], list[tuple[int, int, int, int]]]]:
    """Листы книги: (имя, строки значений, объединённые диапазоны (r1, r2, c1, c2) с нуля)."""
    stream = _ole_stream(Path(path).read_bytes())
    records = list(_records(stream))
    if not records or records[0][1] not in (0x0809, 0x0409, 0x0209, 0x0009):
        raise XlsError("неизвестный формат книги")
    version, = struct.unpack_from("<H", records[0][2], 0)
    biff8 = version == 0x0600
    encoding = "cp1251"

    sheets_meta: list[tuple[int, str]] = []
    sst: list[str] = []
    for i, (pos, rid, body) in enumerate(records):
        if rid == 0x0042 and len(body) >= 2:
            encoding = CODEPAGES.get(struct.unpack_from("<H", body, 0)[0], encoding)
        elif rid == 0x0085:
            offset, = struct.unpack_from("<I", body, 0)
            kind = body[5]
            name = _short_string(body, 6, encoding, biff8, wide_len=False)
            if kind == 0:
                sheets_meta.append((offset, name))
        elif rid == 0x00FC:
            parts = [body]
            for _, rid2, body2 in records[i + 1 :]:
                if rid2 != 0x003C:
                    break
                parts.append(body2)
            sst = _read_sst(parts)
        elif rid == 0x000A:
            break

    by_offset = {pos: idx for idx, (pos, _, _) in enumerate(records)}
    result = []
    for offset, name in sheets_meta:
        start = by_offset.get(offset)
        if start is None:
            continue
        cells: dict[tuple[int, int], object] = {}
        merges: list[tuple[int, int, int, int]] = []
        pending_formula: tuple[int, int] | None = None
        for _, rid, body in records[start + 1 :]:
            if rid == 0x000A:
                break
            try:
                if rid == 0x00FD:
                    r, c, _, idx = struct.unpack_from("<HHHI", body, 0)
                    cells[(r, c)] = sst[idx] if idx < len(sst) else ""
                elif rid == 0x0203:
                    r, c, _, v = struct.unpack_from("<HHHd", body, 0)
                    cells[(r, c)] = v
                elif rid == 0x027E:
                    r, c, _, v = struct.unpack_from("<HHHI", body, 0)
                    cells[(r, c)] = _rk(v)
                elif rid == 0x00BD:
                    r, c = struct.unpack_from("<HH", body, 0)
                    count = (len(body) - 6) // 6
                    for k in range(count):
                        _, v = struct.unpack_from("<HI", body, 4 + 6 * k)
                        cells[(r, c + k)] = _rk(v)
                elif rid in (0x0204, 0x0004):
                    r, c = struct.unpack_from("<HH", body, 0)
                    cells[(r, c)] = _short_string(body, 6, encoding, biff8, wide_len=True)
                elif rid == 0x00D6:  # RSTRING (BIFF5)
                    r, c = struct.unpack_from("<HH", body, 0)
                    cells[(r, c)] = _short_string(body, 6, encoding, biff8, wide_len=True)
                elif rid in (0x0006, 0x0406, 0x0206):
                    r, c = struct.unpack_from("<HH", body, 0)
                    res = body[6:14]
                    if res[6:8] == b"\xff\xff":
                        if res[0] == 0:
                            pending_formula = (r, c)
                        elif res[0] == 1:
                            cells[(r, c)] = bool(res[2])
                    else:
                        cells[(r, c)] = struct.unpack("<d", res)[0]
                elif rid == 0x0207 and pending_formula:
                    cells[pending_formula] = _short_string(body, 0, encoding, biff8, wide_len=True)
                    pending_formula = None
                elif rid == 0x0205:
                    r, c, _, v, is_err = struct.unpack_from("<HHHBB", body, 0)
                    if not is_err:
                        cells[(r, c)] = bool(v)
                elif rid == 0x00E5:
                    count, = struct.unpack_from("<H", body, 0)
                    for k in range(count):
                        merges.append(struct.unpack_from("<HHHH", body, 2 + 8 * k))
            except struct.error:
                continue
        if cells:
            n_rows = max(r for r, _ in cells) + 1
            n_cols = max(c for _, c in cells) + 1
            grid = [[None] * n_cols for _ in range(n_rows)]
            for (r, c), v in cells.items():
                grid[r][c] = v
        else:
            grid = []
        result.append((name, grid, merges))
    return result
