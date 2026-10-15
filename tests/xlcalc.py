"""Вычислитель формул отчёта: ссылки между листами, IF/AND, SUMIFS, COUNTIFS и AGGREGATE.

Нужен, чтобы проверить книгу без Excel: закупка, Кат. А, минимум цены конкурента и дашборд
должны совпасть с тем, что посчитала программа.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from openpyxl.utils import column_index_from_string

CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


class XLErr(Exception):
    def __init__(self, code: str):
        self.code = code
        super().__init__(code)


def _num(v):
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, (int, float)):
        return float(v)
    if v is None:
        return 0.0
    if isinstance(v, str):
        try:
            return float(v.replace(",", ".").replace(" ", ""))
        except ValueError:
            raise XLErr("#VALUE!") from None
    if isinstance(v, XLErr):
        raise v
    raise XLErr("#VALUE!")


def _truth(v) -> bool:
    if isinstance(v, XLErr):
        raise v
    if isinstance(v, str):
        if v.upper() == "TRUE":
            return True
        if v.upper() == "FALSE" or v == "":
            return False
        raise XLErr("#VALUE!")
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    raise XLErr("#VALUE!")


def _is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _same(a, b) -> bool:
    if isinstance(a, str) and isinstance(b, str):
        return a.casefold() == b.casefold()
    if _is_number(a) and _is_number(b):
        return float(a) == float(b)
    if a is None and (b is None or b == "" or b == 0):
        return b in (None, "", 0)
    if b is None and (a is None or a == "" or a == 0):
        return a in (None, "", 0)
    if a is None or b is None:
        return False
    return a == b


def _cmp(op: str, a, b) -> bool:
    if isinstance(a, XLErr):
        raise a
    if isinstance(b, XLErr):
        raise b
    if op in ("=", "<>"):
        eq = _same(a, b)
        return eq if op == "=" else not eq
    if isinstance(a, str) or isinstance(b, str) or a is None or b is None:
        # пустая ячейка в сравнении «больше/меньше» ведёт себя как 0
        if a is None:
            a = 0
        if b is None:
            b = 0
        if isinstance(a, str) or isinstance(b, str):
            sa, sb = str(a).casefold(), str(b).casefold()
            order = (sa > sb) - (sa < sb)
        else:
            order = (_num(a) > _num(b)) - (_num(a) < _num(b))
    else:
        order = (_num(a) > _num(b)) - (_num(a) < _num(b))
    return {">": order > 0, "<": order < 0, ">=": order >= 0, "<=": order <= 0}[op]


def _match(value, crit) -> bool:
    """Критерий COUNTIFS/SUMIFS: числовое условие смотрит только на числа, текст и пустые не подходят."""
    if isinstance(value, XLErr):
        return False
    if isinstance(crit, str):
        for op in (">=", "<=", "<>", ">", "<", "="):
            if crit.startswith(op) and len(crit) > len(op):
                rest = crit[len(op):]
                try:
                    number = float(rest)
                except ValueError:
                    return _cmp(op, "" if value is None else value, rest)
                if not _is_number(value):
                    return False
                return _cmp(op, value, number)
        if crit == "":
            return value in (None, "")
        return isinstance(value, str) and value.casefold() == crit.casefold()
    if _is_number(crit):
        return _is_number(value) and float(value) == float(crit)
    if isinstance(crit, bool):
        return value is crit
    return _same(value, crit)


@dataclass
class Cell:
    sheet: str | None
    col: int
    row: int


@dataclass
class Range:
    sheet: str | None
    c1: int
    r1: int
    c2: int
    r2: int

    def cells(self):
        for r in range(min(self.r1, self.r2), max(self.r1, self.r2) + 1):
            for c in range(min(self.c1, self.c2), max(self.c1, self.c2) + 1):
                yield Cell(self.sheet, c, r)


class Book:
    def __init__(self, wb):
        self.wb = wb
        self.cache: dict[tuple, object] = {}
        self.stack: set[tuple] = set()

    def get(self, sheet: str | None, col: int, row: int, current: str):
        name = sheet or current
        key = (name, col, row)
        if key in self.cache:
            cached = self.cache[key]
            if cached is ...:
                raise XLErr("#REF!")
            return cached
        if key in self.stack:
            raise XLErr("#REF!")
        ws = self.wb[name]
        raw = ws.cell(row, col).value
        self.stack.add(key)
        self.cache[key] = ...
        try:
            value = self._raw(raw, name)
        except XLErr as e:
            value = e
        self.stack.discard(key)
        self.cache[key] = value
        return value

    def _raw(self, raw, sheet: str):
        if isinstance(raw, str) and raw.startswith("="):
            return eval_formula(raw[1:], self, sheet)
        return raw

    def values(self, ref: Range | Cell, current: str) -> list:
        if isinstance(ref, Range) and ref.r1 == 0:
            name = ref.sheet or current
            last = self.wb[name].max_row or 1
            ref = Range(ref.sheet, ref.c1, 1, ref.c2, last)
        cells = ref.cells() if isinstance(ref, Range) else [ref]
        out = []
        for cell in cells:
            value = self.get(cell.sheet, cell.col, cell.row, current)
            out.append(value)
        return out


# ---------------------------------------------------------------- разбор


class Tokens:
    def __init__(self, text: str):
        self.toks = list(_tokenize(text))
        self.i = 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else None

    def take(self):
        t = self.peek()
        self.i += 1
        return t

    def want(self, kind, value=None):
        t = self.take()
        if t is None or t[0] != kind or (value is not None and t[1] != value):
            raise XLErr("#VALUE!")
        return t


def _tokenize(text: str):
    i, n = 0, len(text)
    while i < n:
        c = text[i]
        if c.isspace():
            i += 1
            continue
        if c == '"':
            i += 1
            buf = []
            while i < n:
                if text[i] == '"':
                    if i + 1 < n and text[i + 1] == '"':
                        buf.append('"')
                        i += 2
                    else:
                        i += 1
                        break
                else:
                    buf.append(text[i])
                    i += 1
            yield ("str", "".join(buf))
            continue
        if c == "'":
            i += 1
            buf = []
            while i < n:
                if text[i] == "'":
                    if i + 1 < n and text[i + 1] == "'":
                        buf.append("'")
                        i += 2
                    else:
                        i += 1
                        break
                else:
                    buf.append(text[i])
                    i += 1
            if i < n and text[i] == "!":
                i += 1
            yield ("sheet", "".join(buf))
            continue
        if c in "+-*/^=&(),:!":
            yield ("op", c)
            i += 1
            continue
        if c in "<>":
            if i + 1 < n and text[i + 1] == "=" or c == "<" and i + 1 < n and text[i + 1] == ">":
                yield ("op", text[i:i + 2])
                i += 2
            else:
                yield ("op", c)
                i += 1
            continue
        if c.isdigit() or c == "." and i + 1 < n and text[i + 1].isdigit():
            j = i
            while j < n and (text[j].isdigit() or text[j] in ".eE"):
                if text[j] in "eE" and j + 1 < n and text[j + 1] in "+-":
                    j += 2
                else:
                    j += 1
            yield ("num", float(text[i:j]))
            i = j
            continue
        if c.isalpha() or c in "_$":
            j = i
            while j < n and (text[j].isalnum() or text[j] in "$_."):
                j += 1
            yield ("word", text[i:j])
            i = j
            continue
        raise XLErr("#VALUE!")


def _column(word: str) -> int | None:
    m = re.fullmatch(r"\$?([A-Za-z]{1,3})", word)
    if not m:
        return None
    return column_index_from_string(m.group(1))


def _cell(word: str) -> tuple[int, int] | None:
    m = CELL_RE.match(word)
    if not m:
        return None
    return column_index_from_string(m.group(1)), int(m.group(2))


def parse(text: str):
    ts = Tokens(text)
    node = _parse_cmp(ts)
    if ts.peek() is not None:
        raise XLErr("#VALUE!")
    return node


def _parse_cmp(ts: Tokens):
    node = _parse_concat(ts)
    while (p := ts.peek()) and p[0] == "op" and p[1] in ("=", "<>", "<", ">", "<=", ">="):
        op = ts.take()[1]
        node = ("cmp", op, node, _parse_concat(ts))
    return node


def _parse_concat(ts: Tokens):
    node = _parse_add(ts)
    while (p := ts.peek()) and p == ("op", "&"):
        ts.take()
        node = ("bin", "&", node, _parse_add(ts))
    return node


def _parse_add(ts: Tokens):
    node = _parse_mul(ts)
    while (p := ts.peek()) and p[0] == "op" and p[1] in "+-":
        op = ts.take()[1]
        node = ("bin", op, node, _parse_mul(ts))
    return node


def _parse_mul(ts: Tokens):
    node = _parse_unary(ts)
    while (p := ts.peek()) and p[0] == "op" and p[1] in "*/":
        op = ts.take()[1]
        node = ("bin", op, node, _parse_unary(ts))
    return node


def _parse_unary(ts: Tokens):
    if (p := ts.peek()) and p[0] == "op" and p[1] in "+-":
        op = ts.take()[1]
        return ("unary", op, _parse_unary(ts))
    return _parse_primary(ts)


def _parse_primary(ts: Tokens):
    sheet = None
    p = ts.peek()
    if p and p[0] == "sheet":
        sheet = ts.take()[1]
    elif p and p[0] == "word" and ts.i + 1 < len(ts.toks) and ts.toks[ts.i + 1] == ("op", "!"):
        sheet = ts.take()[1]
        ts.take()
    p = ts.peek()
    if p is None:
        raise XLErr("#VALUE!")
    if p[0] == "num":
        ts.take()
        return ("num", p[1])
    if p[0] == "str":
        ts.take()
        return ("str", p[1])
    if p == ("op", "("):
        ts.take()
        node = _parse_cmp(ts)
        ts.want("op", ")")
        return node
    if p[0] == "word":
        word = ts.take()[1]
        if ts.peek() == ("op", "("):
            ts.take()
            args = []
            if ts.peek() != ("op", ")"):
                while True:
                    args.append(_parse_cmp(ts))
                    if ts.peek() == ("op", ","):
                        ts.take()
                        continue
                    break
            ts.want("op", ")")
            return ("fn", word.upper(), args)
        coord = _cell(word)
        col_only = _column(word)
        if coord is None and col_only and ts.peek() == ("op", ":"):
            ts.take()
            second = ts.want("word")[1]
            col2 = _column(second)
            if not col2:
                raise XLErr("#VALUE!")
            return Range(sheet, col_only, 0, col2, 0)
        if not coord:
            raise XLErr("#NAME?")
        if ts.peek() == ("op", ":"):
            ts.take()
            second = ts.want("word")[1]
            coord2 = _cell(second)
            if not coord2:
                raise XLErr("#VALUE!")
            return Range(sheet, coord[0], coord[1], coord2[0], coord2[1])
        return Cell(sheet, coord[0], coord[1])
    raise XLErr("#VALUE!")


# ---------------------------------------------------------------- вычисление


def eval_formula(text: str, book: Book, sheet: str):
    return _unwrap(_ev(parse(text), book, sheet), book, sheet)


def _unwrap(v, book: Book, sheet: str):
    if isinstance(v, Cell):
        return book.values(v, sheet)[0]
    if isinstance(v, Range):
        vals = book.values(v, sheet)
        return vals[0] if len(vals) == 1 else vals
    return v


def _flatten(v):
    if isinstance(v, list):
        for x in v:
            yield from _flatten(x)
    else:
        yield v


def _expand(v, book: Book, sheet: str):
    """Одна ячейка — скаляр, диапазон — массив (так array-формулы AGGREGATE работают построчно)."""
    if isinstance(v, Cell):
        return book.values(v, sheet)[0]
    if isinstance(v, Range):
        return book.values(v, sheet)
    return v


def _broadcast(op, a, b):
    a_list, b_list = isinstance(a, list), isinstance(b, list)
    if not a_list and not b_list:
        return op(a, b)
    if not a_list:
        a = [a] * len(b)
    if not b_list:
        b = [b] * len(a)
    if len(a) != len(b):
        raise XLErr("#VALUE!")
    return [_safe(op, x, y) for x, y in zip(a, b)]


def _safe(op, a, b):
    try:
        if isinstance(a, XLErr):
            return a
        if isinstance(b, XLErr):
            return b
        return op(a, b)
    except XLErr as e:
        return e


def _bin(op, a, b):
    if op == "&":
        return _as_text(a) + _as_text(b)
    if isinstance(a, XLErr):
        raise a
    if isinstance(b, XLErr):
        raise b
    if op == "+":
        return _num(a) + _num(b)
    if op == "-":
        return _num(a) - _num(b)
    if op == "*":
        return _num(a) * _num(b)
    if op == "/":
        d = _num(b)
        if d == 0:
            raise XLErr("#DIV/0!")
        return _num(a) / d
    raise XLErr("#VALUE!")


def _as_text(v) -> str:
    if isinstance(v, XLErr):
        raise v
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        text = format(v, ".12g")
        return text
    return str(v)


def _ev(node, book: Book, sheet: str):
    if isinstance(node, (Cell, Range)):
        return node
    kind = node[0]
    if kind == "num":
        return node[1]
    if kind == "str":
        return node[1]
    if kind == "unary":
        v = _expand(_ev(node[2], book, sheet), book, sheet)
        def one(x):
            return _num(x) if node[1] == "+" else -_num(x)
        if isinstance(v, list):
            return [_safe(lambda x, _y: one(x), x, 0) for x in v]
        return one(v)
    if kind == "bin":
        a = _expand(_ev(node[2], book, sheet), book, sheet)
        b = _expand(_ev(node[3], book, sheet), book, sheet)
        return _broadcast(lambda x, y: _bin(node[1], x, y), a, b)
    if kind == "cmp":
        a = _expand(_ev(node[2], book, sheet), book, sheet)
        b = _expand(_ev(node[3], book, sheet), book, sheet)
        return _broadcast(lambda x, y: _cmp(node[1], x, y), a, b)
    if kind == "fn":
        return _fn(node[1], node[2], book, sheet)
    raise XLErr("#VALUE!")


def _args(nodes, book, sheet):
    return [_ev(n, book, sheet) for n in nodes]


def _numbers(values):
    for v in values:
        if _is_number(v):
            yield float(v)


def _fn(name: str, nodes, book: Book, sheet: str):
    name = name.split(".")[-1]
    if name == "IF":
        cond = _ev(nodes[0], book, sheet)
        if isinstance(cond, (Cell, Range)):
            cond = book.values(cond, sheet)
            cond = cond[0] if isinstance(cond, list) and len(cond) == 1 else cond
        return _ev(nodes[1] if _truth(cond) else nodes[2], book, sheet)
    if name == "IFERROR":
        try:
            v = _ev(nodes[0], book, sheet)
            if isinstance(v, (Cell, Range)):
                vals = book.values(v, sheet)
                v = vals[0] if len(vals) == 1 else vals
            if isinstance(v, XLErr):
                raise v
            return v
        except XLErr:
            return _ev(nodes[1], book, sheet)
    if name == "AND":
        return all(_truth(_scalar(_ev(n, book, sheet), book, sheet)) for n in nodes)
    if name == "OR":
        return any(_truth(_scalar(_ev(n, book, sheet), book, sheet)) for n in nodes)
    if name == "NOT":
        return not _truth(_scalar(_ev(nodes[0], book, sheet), book, sheet))
    if name == "ISNUMBER":
        v = _scalar(_ev(nodes[0], book, sheet), book, sheet)
        return _is_number(v)
    if name in ("SUM", "COUNT", "COUNTA", "MIN", "MAX", "AVERAGE"):
        vals = []
        for n in nodes:
            v = _ev(n, book, sheet)
            vals.extend(book.values(v, sheet) if isinstance(v, (Cell, Range)) else _flatten(_expand(v, book, sheet)))
        if name == "SUM":
            return sum(_numbers(vals))
        if name == "COUNT":
            return sum(1 for v in vals if _is_number(v))
        if name == "COUNTA":
            return sum(1 for v in vals if v not in (None, ""))
        nums = list(_numbers(vals))
        if not nums:
            raise XLErr("#VALUE!")
        if name == "MIN":
            return min(nums)
        if name == "MAX":
            return max(nums)
        return sum(nums) / len(nums)
    if name == "ABS":
        return abs(_num(_scalar(_ev(nodes[0], book, sheet), book, sheet)))
    if name == "ROUND":
        n = _num(_scalar(_ev(nodes[0], book, sheet), book, sheet))
        digits = int(_num(_scalar(_ev(nodes[1], book, sheet), book, sheet)))
        return round(n, digits)
    if name == "SUBTOTAL":
        fn = int(_num(_scalar(_ev(nodes[0], book, sheet), book, sheet)))
        vals = []
        for n in nodes[1:]:
            v = _ev(n, book, sheet)
            vals.extend(book.values(v, sheet) if isinstance(v, (Cell, Range)) else list(_flatten(v)))
        if fn in (9, 109):
            return sum(_numbers(vals))
        if fn in (2, 102):
            return sum(1 for v in vals if _is_number(v))
        if fn in (3, 103):
            return sum(1 for v in vals if v not in (None, ""))
        raise XLErr("#VALUE!")
    if name in ("SUMIFS", "COUNTIFS"):
        return _ifs(name, nodes, book, sheet)
    if name == "AGGREGATE":
        return _aggregate(nodes, book, sheet)
    if name == "MATCH":
        lookup = _scalar(_ev(nodes[0], book, sheet), book, sheet)
        rng = _ev(nodes[1], book, sheet)
        if not isinstance(rng, Range):
            raise XLErr("#VALUE!")
        for i, value in enumerate(book.values(rng, sheet), 1):
            if _same(value, lookup):
                return i
        raise XLErr("#N/A")
    if name == "INDEX":
        rng = _ev(nodes[0], book, sheet)
        if not isinstance(rng, Range):
            raise XLErr("#VALUE!")
        pos = int(_num(_scalar(_ev(nodes[1], book, sheet), book, sheet)))
        vals = book.values(rng, sheet)
        if not 1 <= pos <= len(vals):
            raise XLErr("#REF!")
        return vals[pos - 1]
    raise XLErr("#NAME?")


def _scalar(v, book, sheet):
    if isinstance(v, (Cell, Range)):
        vals = book.values(v, sheet)
        if len(vals) != 1:
            raise XLErr("#VALUE!")
        v = vals[0]
    if isinstance(v, XLErr):
        raise v
    return v


def _ifs(name, nodes, book, sheet):
    idx = 0
    if name == "SUMIFS":
        sum_range = _ev(nodes[0], book, sheet)
        if not isinstance(sum_range, Range):
            raise XLErr("#VALUE!")
        sums = book.values(sum_range, sheet)
        idx = 1
    else:
        sums = None
    pairs = []
    while idx < len(nodes):
        rng = _ev(nodes[idx], book, sheet)
        crit = _scalar(_ev(nodes[idx + 1], book, sheet), book, sheet)
        if not isinstance(rng, Range):
            raise XLErr("#VALUE!")
        pairs.append((book.values(rng, sheet), crit))
        idx += 2
    n = len(pairs[0][0])
    total = count = 0
    for i in range(n):
        if all(_match(vals[i], crit) for vals, crit in pairs):
            count += 1
            if sums is not None and _is_number(sums[i]):
                total += float(sums[i])
    return total if name == "SUMIFS" else count


def _aggregate(nodes, book, sheet):
    fn = int(_num(_scalar(_ev(nodes[0], book, sheet), book, sheet)))
    option = int(_num(_scalar(_ev(nodes[1], book, sheet), book, sheet)))
    arr = _expand(_ev(nodes[2], book, sheet), book, sheet)
    k = int(_num(_scalar(_ev(nodes[3], book, sheet), book, sheet)))
    values = []
    for v in _flatten(arr if isinstance(arr, list) else [arr]):
        if option == 6 and isinstance(v, XLErr):
            continue
        if isinstance(v, XLErr):
            raise v
        if _is_number(v):
            values.append(float(v))
    if fn == 15:  # SMALL
        values.sort()
        if not 1 <= k <= len(values):
            raise XLErr("#NUM!")
        return values[k - 1]
    if fn == 14:  # LARGE
        values.sort(reverse=True)
        if not 1 <= k <= len(values):
            raise XLErr("#NUM!")
        return values[k - 1]
    if fn == 5:  # MIN
        if not values:
            raise XLErr("#NUM!")
        return min(values)
    raise XLErr("#VALUE!")


def cell(wb, sheet: str, coord: str):
    """Значение ячейки книги openpyxl, с вычислением формулы."""
    book = Book(wb)
    col, row = _cell(coord.replace("$", ""))
    value = book.get(sheet, col, row, sheet)
    if isinstance(value, XLErr):
        raise value
    return value
