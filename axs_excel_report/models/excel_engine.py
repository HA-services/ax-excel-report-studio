"""Excel template engine of AXSync Excel Report Studio.

A template is a normal .xlsx workbook whose cells contain placeholders:

* ``{{ object.partner_id.name }}``: a value of the printed record (``object``), of the loop record, of
  ``company``, ``user``, or ``today`` / ``now`` / ``index`` / ``count``.
* ``{{ for line in object.order_line }}``: repeats the row (or the rows up to ``{{ endfor }}``) once per record.
  ``{{ for rec in records }}`` loops over all the printed records (list reports).
* ``{{ sum:object.order_line.price_subtotal }}``: aggregates (sum, avg, min, max, count).
* Filters after ``|``: ``date:%d/%m/%Y``, ``datetime:%d/%m/%Y %H:%M``, ``upper``, ``lower``, ``title``,
  ``round:2``, ``default:text``, ``image`` or ``image:80`` (picture height in pixels).

Formulas are kept: their references follow the rows that move, and a range covering a repeated row grows with it.
Charts, conditional formatting, data validation, filters, frozen panes and the print area follow the rows too.
"""
import base64
import datetime
import io
import re
from copy import copy, deepcopy

import pytz

from odoo.tools import html2plaintext

from .compat import user_has_groups

try:
    import openpyxl
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.cell_range import MultiCellRange
except ImportError:  # pragma: no cover - declared in the manifest
    openpyxl = None

PLACEHOLDER = re.compile(r'\{\{\s*(.*?)\s*\}\}')
LOOP = re.compile(r'^for\s+([a-z_][a-z0-9_]*)\s+in\s+([a-z_][\w.]*)$', re.I)
ENDLOOP = re.compile(r'^endfor$', re.I)
PATH = re.compile(r'^[a-z_][a-z0-9_]*(\.[a-z_][a-z0-9_]*)*$', re.I)
AGGREGATES = ('sum', 'avg', 'min', 'max', 'count')
TOTAL_FORMULA = re.compile(r"^=(SUM|AVERAGE|MIN|MAX|SUBTOTAL\(\d+,)\(?\$?([A-Z]{1,3}\$?\d+)(:\$?[A-Z]{1,3}\$?\d+)?\)$", re.I)
RANGE_REF = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?$")
SHEET_REF = re.compile(r"^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$")
CHART_REFS = ('val', 'cat', 'tx', 'xVal', 'yVal', 'bubbleSize')
CELL_REF = re.compile(r"(?<![A-Za-z_!'\]])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])")
EXCEL_DATE = {'%d': 'dd', '%m': 'mm', '%Y': 'yyyy', '%y': 'yy', '%b': 'mmm', '%B': 'mmmm', '%a': 'ddd',
              '%A': 'dddd', '%H': 'hh', '%M': 'mm', '%S': 'ss', '%I': 'hh', '%p': 'AM/PM', '%e': 'd'}


class TemplateError(Exception):
    """A placeholder or a loop of the template cannot be used."""


def excel_format(strftime_format):
    """Excel number format for a Python date format."""
    fmt = strftime_format or '%m/%d/%Y'
    for py, xl in EXCEL_DATE.items():
        fmt = fmt.replace(py, xl)
    return fmt


class Value:
    """A resolved placeholder: the raw value plus what is needed to write it in a cell."""
    __slots__ = ('value', 'kind', 'number_format', 'text')

    def __init__(self, value, kind='text', number_format=None, text=None):
        self.value = value
        self.kind = kind
        self.number_format = number_format
        self.text = text


class Renderer:
    """Renders an openpyxl workbook for records of one model."""

    def __init__(self, env, records, lang=None):
        self.env = env
        self.records = records
        lang_code = lang or env.lang or 'en_US'
        Lang = env['res.lang']
        self.lang = Lang._lang_get(lang_code) if hasattr(Lang, '_lang_get') else Lang.search(
            [('code', '=', lang_code)], limit=1)
        self.date_format = (self.lang and self.lang.date_format) or '%m/%d/%Y'
        self.time_format = (self.lang and self.lang.time_format) or '%H:%M:%S'
        self.tz = pytz.timezone(env.context.get('tz') or env.user.tz or 'UTC')
        self.today = datetime.datetime.now(self.tz).date()
        self.now = datetime.datetime.now(self.tz).replace(tzinfo=None)
        self._images = []

    # ------------------------------------------------------------------ values
    def _roots(self, scope):
        roots = {
            'records': self.records,
            'company': self.env.company,
            'user': self.env.user,
        }
        roots.update(scope)
        return roots

    def resolve(self, expression, scope):
        """Evaluate one placeholder expression in ``scope``; returns a :class:`Value`."""
        parts = [p.strip() for p in expression.split('|')]
        expr, filters = parts[0], parts[1:]
        aggregate = None
        if ':' in expr and expr.split(':', 1)[0].strip().lower() in AGGREGATES:
            aggregate, expr = [p.strip() for p in expr.split(':', 1)]
            aggregate = aggregate.lower()
        if expr in ('today', 'now', 'index', 'count'):
            value = self._special(expr, scope)
        else:
            if not PATH.match(expr):
                raise TemplateError("'%s' is not a valid placeholder" % expression)
            value = self._path(expr, scope, aggregate)
        for flt in filters:
            value = self._filter(value, flt)
        return value

    def _special(self, name, scope):
        if name == 'today':
            return Value(self.today, 'date', excel_format(self.date_format))
        if name == 'now':
            return Value(self.now, 'datetime', excel_format('%s %s' % (self.date_format, self.time_format[:5])))
        if name == 'index':
            return Value(scope.get('__index__', 1), 'number', '0')
        return Value(len(scope.get('__items__', self.records)), 'number', '0')

    def _path(self, expr, scope, aggregate=None):
        names = expr.split('.')
        roots = self._roots(scope)
        if names[0] not in roots:
            raise TemplateError("'%s' is unknown, use object, a loop variable, records, company or user"
                                % names[0])
        current = roots[names[0]]
        field = None
        holder = None
        for name in names[1:]:
            if not hasattr(current, '_fields'):
                raise TemplateError("'%s' cannot be followed in %s" % (name, expr))
            if name not in current._fields:
                raise TemplateError("The field '%s' does not exist on %s" % (name, current._name))
            field = current._fields[name]
            if field.groups and not user_has_groups(self.env, field.groups):
                raise TemplateError("You are not allowed to read the field '%s' of %s" % (name, current._name))
            holder = current
            if len(current) > 1 or aggregate:
                current = current.mapped(name)
            else:
                current = current[name] if current else (current.browse() if field.relational else False)
        return self._typed(current, field, holder, aggregate)

    def _typed(self, value, field, holder, aggregate=None):
        if aggregate:
            if hasattr(value, '_name'):
                values = value.ids
            else:
                values = value if isinstance(value, list) else [value]
            if aggregate == 'count':
                return Value(len(values), 'number', '0')
            numbers = [v for v in values if isinstance(v, (int, float)) and not isinstance(v, bool)]
            if not numbers:
                result = 0
            elif aggregate == 'sum':
                result = sum(numbers)
            elif aggregate == 'avg':
                result = sum(numbers) / len(numbers)
            elif aggregate == 'min':
                result = min(numbers)
            else:
                result = max(numbers)
            return Value(result, 'number', self._number_format(field, holder))
        if hasattr(value, '_name'):  # recordset
            if field is None:  # a root such as object or company
                return Value(value.display_name if len(value) == 1 else ', '.join(value.mapped('display_name')))
            if field.type in ('binary',):
                return Value(False)
            names = [r.display_name or '' for r in value]
            return Value(', '.join(n for n in names if n))
        if isinstance(value, list):  # mapped over several records
            if field is not None and field.type == 'binary':
                value = value[0] if value else False
            else:
                texts = [self._text(Value(v, *self._kind(v, field, holder))) for v in value if v not in (False, None)]
                return Value(', '.join(texts))
        if field is None:
            return Value(value)
        if field.type == 'binary':
            return Value(value, 'binary')
        if field.type == 'selection':
            selection = dict(field._description_selection(self.env))
            return Value(selection.get(value, value or ''))
        if field.type == 'boolean':
            return Value(bool(value), 'bool')
        if field.type == 'html':
            return Value(html2plaintext(value or '').strip())
        if value is False or value is None:
            if field.type in ('integer', 'float', 'monetary'):
                return Value(0, 'number', self._number_format(field, holder))
            return Value('')
        kind, number_format = self._kind(value, field, holder)
        return Value(value, kind, number_format)

    def _kind(self, value, field, holder):
        if field is not None and field.type == 'datetime' and isinstance(value, datetime.datetime):
            return 'datetime', excel_format('%s %s' % (self.date_format, self.time_format[:5]))
        if isinstance(value, datetime.datetime):
            return 'datetime', excel_format('%s %s' % (self.date_format, self.time_format[:5]))
        if isinstance(value, datetime.date):
            return 'date', excel_format(self.date_format)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return 'number', self._number_format(field, holder)
        return 'text', None

    def _number_format(self, field, holder):
        if field is None:
            return None
        if field.type == 'integer':
            return '#,##0'
        digits = 2
        try:
            field_digits = field.get_digits(self.env)
            if field_digits:
                digits = field_digits[1]
        except Exception:  # noqa: BLE001 - digits are a nicety
            pass
        base = '#,##0' + ('.' + '0' * digits if digits else '')
        if field.type == 'monetary' and holder is not None and len(holder) == 1:
            currency_field = field.get_currency_field(holder) if hasattr(field, 'get_currency_field') \
                else getattr(field, 'currency_field', None) or 'currency_id'
            currency = holder[currency_field] if currency_field in holder._fields else None
            if currency:
                digits = currency.decimal_places
                base = '#,##0' + ('.' + '0' * digits if digits else '')
                symbol = currency.symbol.replace('"', '')
                return '"%s"\\ %s' % (symbol, base) if currency.position == 'before' else '%s\\ "%s"' % (base, symbol)
        return base

    def _filter(self, value, flt):
        name, _sep, arg = flt.partition(':')
        name = name.strip().lower()
        arg = arg.strip()
        if name in ('date', 'datetime'):
            if isinstance(value.value, (datetime.date, datetime.datetime)):
                fmt = arg or (self.date_format if name == 'date' else '%s %s' % (self.date_format, self.time_format))
                value.text = value.value.strftime(fmt)
                value.number_format = excel_format(fmt)
            return value
        if name in ('upper', 'lower', 'title'):
            text = self._text(value)
            return Value(getattr(text, name)())
        if name == 'round':
            if value.kind == 'number':
                digits = int(arg or 0)
                value.value = round(value.value, digits)
                value.number_format = '#,##0' + ('.' + '0' * digits if digits else '')
            return value
        if name == 'default':
            if value.value in (False, None, '') or (value.kind == 'text' and not value.value):
                return Value(arg)
            return value
        if name == 'image':
            value.kind = 'image'
            value.text = arg or ''
            return value
        raise TemplateError("Unknown filter '%s'" % flt)

    def _text(self, value):
        """Text of a value inside a sentence."""
        if value.text is not None and value.kind != 'image':
            return value.text
        v = value.value
        if value.kind in ('date', 'datetime') and v:
            fmt = self.date_format if value.kind == 'date' else '%s %s' % (self.date_format, self.time_format[:5])
            return v.strftime(fmt)
        if value.kind == 'number':
            return self._format_number(v, value.number_format)
        if value.kind == 'bool':
            return 'Yes' if v else 'No'
        if v in (False, None) or value.kind in ('image', 'binary'):
            return ''
        return str(v)

    def _format_number(self, number, number_format):
        decimals = 0
        if number_format and '.' in number_format:
            decimals = len(re.findall(r'\.(0+)', number_format)[0]) if re.findall(r'\.(0+)', number_format) else 0
        text = ('{:,.%df}' % decimals).format(number)
        sep = (self.lang and self.lang.thousands_sep) or ','
        point = (self.lang and self.lang.decimal_point) or '.'
        text = text.replace(',', '\x00').replace('.', point).replace('\x00', sep)
        symbol = re.findall(r'"([^"]+)"', number_format or '')
        if symbol:
            text = '%s %s' % (symbol[0], text) if number_format.startswith('"') else '%s %s' % (text, symbol[0])
        return text

    # ------------------------------------------------------------------ cells
    def render_text(self, text, scope):
        """Render a string with placeholders as plain text (file names, titles)."""
        return PLACEHOLDER.sub(lambda m: self._text(self.resolve(m.group(1), scope)), text or '')

    def _render_cell(self, ws, target, source_value, scope):
        if not isinstance(source_value, str) or '{{' not in source_value:
            return source_value
        matches = list(PLACEHOLDER.finditer(source_value))
        whole = len(matches) == 1 and matches[0].group(0) == source_value.strip()
        if whole:
            value = self.resolve(matches[0].group(1), scope)
            if value.kind == 'image' or (value.kind == 'binary' and value.value):
                self._add_image(ws, target, value)
                return None
            if value.kind in ('number', 'date', 'datetime', 'bool'):
                if value.number_format and target.number_format in ('General', None):
                    target.number_format = value.number_format
                if value.kind == 'date' and isinstance(value.value, datetime.datetime):
                    return value.value.date()
                if value.text is not None:
                    return value.text if value.kind not in ('date', 'datetime') else value.value
                return value.value
            return self._text(value)
        return PLACEHOLDER.sub(lambda m: self._text(self.resolve(m.group(1), scope)), source_value)

    def _add_image(self, ws, cell, value):
        data = value.value
        if not data:
            return
        try:
            raw = data.content if hasattr(data, 'content') else (base64.b64decode(data) if isinstance(data, (str, bytes)) else bytes(data))
        except Exception:  # noqa: BLE001
            return
        try:
            from PIL import Image as PILImage
            pil = PILImage.open(io.BytesIO(raw))
            pil.load()
        except Exception:  # noqa: BLE001 - not an image (for example a PDF in a binary field)
            return
        height = int(value.text or 0) or 60
        ratio = pil.width / float(pil.height or 1)
        if pil.mode not in ('RGB', 'RGBA'):
            pil = pil.convert('RGBA')
        buffer = io.BytesIO()
        pil.save(buffer, format='PNG')
        buffer.seek(0)
        img = XLImage(buffer)
        img.height = height
        img.width = int(height * ratio)
        ws.add_image(img, cell.coordinate)
        self._images.append(buffer)
        dim = ws.row_dimensions[cell.row]
        needed = height * 0.75 + 4
        if not dim.height or dim.height < needed:
            dim.height = needed

    # ------------------------------------------------------------------ sheets
    def _blocks(self, ws):
        """Loop blocks of a sheet: dicts with start, end, var and source.

        A loop runs to the next ``{{ endfor }}`` when one comes before the next loop, otherwise it covers its row.
        """
        markers = []
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or '{{' not in cell.value:
                    continue
                for m in PLACEHOLDER.finditer(cell.value):
                    expr = m.group(1).strip()
                    loop = LOOP.match(expr)
                    if loop:
                        markers.append(('for', cell.row, loop.group(1), loop.group(2)))
                    elif ENDLOOP.match(expr):
                        markers.append(('end', cell.row, None, None))
        blocks = []
        i = 0
        while i < len(markers):
            kind, row, var, source = markers[i]
            if kind == 'end':
                raise TemplateError("{{ endfor }} without a loop (row %s)" % row)
            if i + 1 < len(markers) and markers[i + 1][0] == 'end':
                blocks.append({'start': row, 'end': markers[i + 1][1], 'var': var, 'source': source})
                i += 2
            else:
                blocks.append({'start': row, 'end': row, 'var': var, 'source': source})
                i += 1
        for a, b in zip(blocks, blocks[1:]):
            if b['start'] <= a['end']:
                raise TemplateError("Loops cannot be nested or share a row (rows %s and %s)" % (a['start'], b['start']))
        return blocks

    @staticmethod
    def _strip_markers(text):
        if not isinstance(text, str):
            return text
        stripped = PLACEHOLDER.sub(lambda m: '' if (LOOP.match(m.group(1).strip()) or ENDLOOP.match(m.group(1).strip()))
                                   else m.group(0), text)
        return stripped if stripped.strip() or not text.strip() else None

    def render_sheet(self, ws, scope, source_title=None):
        """Render one worksheet in place.

        :param source_title: title of the sheet in the template, used by chart references
        """
        blocks = self._blocks(ws)
        max_row = ws.max_row
        max_col = ws.max_column
        # snapshot of the template
        cells = {}
        for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
            for cell in row:
                if cell.value is not None or cell.has_style:
                    cells[(cell.row, cell.column)] = (cell.value, copy(cell._style) if cell.has_style else None,
                                                      cell.hyperlink, cell.comment)
        heights = {r: copy(d) for r, d in ws.row_dimensions.items() if r <= max_row}
        merges = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in ws.merged_cells.ranges]
        images = list(ws._images)
        source_title = source_title or ws.title
        for m in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(m))
        ws._cells.clear()
        for r in list(ws.row_dimensions):
            del ws.row_dimensions[r]
        ws._images = []

        # items of each loop
        plan = []  # list of (template_start, template_end, items or None)
        row = 1
        for block in blocks:
            if row < block['start']:
                plan.append((row, block['start'] - 1, None, None))
            source = self._path_records(block['source'], scope)
            plan.append((block['start'], block['end'], block, source))
            row = block['end'] + 1
        if row <= max_row:
            plan.append((row, max_row, None, None))

        # row mapping: template row -> list of output rows (one per copy)
        mapping = {}
        formulas = []
        out = 1
        for start, end, block, items in plan:
            copies = 1 if block is None else max(1, len(items))
            for k in range(copies):
                for r in range(start, end + 1):
                    mapping.setdefault(r, []).append(out)
                    out += 1

        def rewrite(formula, copy_index, block_rows):
            def repl(m):
                col_abs, col, row_abs, row_s = m.groups()
                r = int(row_s)
                targets = mapping.get(r)
                if not targets:
                    return m.group(0)
                if block_rows and r in block_rows and copy_index is not None:
                    new = targets[min(copy_index, len(targets) - 1)]
                else:
                    new = targets[0]
                return '%s%s%s%s' % (col_abs, col, row_abs, new)

            def range_repl(m):
                left, right = m.group(1), m.group(2)
                lm, rm = CELL_REF.fullmatch(left), CELL_REF.fullmatch(right)
                if not lm or not rm:
                    return m.group(0)
                lr, rr = int(lm.group(4)), int(rm.group(4))
                inside = copy_index is not None and block_rows and lr in block_rows and rr in block_rows
                if inside:
                    return '%s:%s' % (CELL_REF.sub(repl, left), CELL_REF.sub(repl, right))
                start_rows = mapping.get(lr) or [lr]
                end_rows = mapping.get(rr) or [rr]
                return '%s%s%s%s:%s%s%s%s' % (lm.group(1), lm.group(2), lm.group(3), start_rows[0],
                                              rm.group(1), rm.group(2), rm.group(3), end_rows[-1])
            # protect strings
            parts = re.split(r'("(?:[^"]|"")*")', formula)
            for i in range(0, len(parts), 2):
                segment = re.sub(r'(\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)', range_repl, parts[i])
                # single references that are not part of a range already rewritten
                tokens = re.split(r'(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)', segment)
                for j in range(0, len(tokens), 2):
                    tokens[j] = CELL_REF.sub(repl, tokens[j])
                parts[i] = ''.join(tokens)
            return ''.join(parts)

        # write the output
        out = 1
        for start, end, block, items in plan:
            block_rows = set(range(start, end + 1)) if block else set()
            copies = [None] if block is None else (list(items) or [None])
            for k, item in enumerate(copies):
                local = dict(scope)
                if block is not None:
                    local[block['var']] = item if item is not None else items.browse() if hasattr(items, 'browse') else None
                    local['__index__'] = k + 1
                    local['__items__'] = items
                for r in range(start, end + 1):
                    if r in heights:
                        dim = copy(heights[r])
                        dim.index = out
                        ws.row_dimensions[out] = dim
                    for (cr, cc), (value, style, link, comment) in cells.items():
                        if cr != r:
                            continue
                        target = ws.cell(row=out, column=cc)
                        if style is not None:
                            target._style = copy(style)
                        if block is not None and isinstance(value, str):
                            value = self._strip_markers(value)
                        if block is not None and item is None and isinstance(value, str) and '{{' in value:
                            value = None  # empty loop: one blank row keeps formulas valid
                        if isinstance(value, str) and value.startswith('='):
                            target.value = rewrite(value, k if block is not None else None, block_rows)
                            formulas.append(target)
                        else:
                            target.value = self._render_cell(ws, target, value, local)
                        if link:
                            target.hyperlink = copy(link)
                        if comment:
                            target.comment = copy(comment)
                    # merged cells of this row
                    for (r1, c1, r2, c2) in merges:
                        if r1 == r and r2 <= end:
                            ws.merge_cells(start_row=out, start_column=c1, end_row=out + r2 - r1, end_column=c2)
                    out += 1
        # totals without a number format take the format of the values they add up,
        # so a sum of amounts shows the currency like the amounts do
        for target in formulas:
            m = TOTAL_FORMULA.match(target.value)
            if not m or target.number_format != 'General':
                continue
            first = ws[m.group(2).replace('$', '')]
            if first.number_format != 'General':
                target.number_format = first.number_format
        # template pictures (logos) keep their anchor row, shifted like the rows
        for img in images:
            try:
                marker = img.anchor._from
                rows = mapping.get(marker.row + 1)
                if rows:
                    marker.row = rows[0] - 1
            except AttributeError:
                pass
            ws._images.append(img)
        self._move_sheet_features(ws, mapping, max_row, out - 1, source_title)
        return ws

    # ------------------------------------------------------------------ charts, formats, validation
    @staticmethod
    def _move_sheet_features(ws, mapping, max_row, last_row, source_title):
        """Charts, conditional formats, validations, filters, panes and print area follow the rows."""
        shift = last_row - max_row

        def row_of(r, last=False):
            rows = mapping.get(r)
            if rows:
                return rows[-1] if last else rows[0]
            return r + shift if r > max_row else r

        def move_range(ref, expand=True):
            """A cell or a range after the rows moved; a single cell of a repeated row covers every copy."""
            m = RANGE_REF.match(ref.strip())
            if not m:
                return ref
            a1, c1, a2, r1, b1, c2, b2, r2 = m.groups()
            if r2 is None:
                first, last = row_of(int(r1)), row_of(int(r1), last=True)
                if expand and last != first:
                    return '%s%s%s%s:%s%s%s%s' % (a1, c1, a2, first, a1, c1, a2, last)
                return '%s%s%s%s' % (a1, c1, a2, first)
            return '%s%s%s%s:%s%s%s%s' % (a1, c1, a2, row_of(int(r1)), b1, c2, b2, row_of(int(r2), last=True))

        def move_ranges(sqref):
            return ' '.join(move_range(part) for part in str(sqref).split())

        def move_formula(formula):
            m = SHEET_REF.match(formula or '')
            if not m:
                return formula
            sheet = (m.group(1) or '').replace("''", "'") or m.group(2)
            if sheet != source_title:
                return formula
            return "'%s'!%s" % (ws.title.replace("'", "''"), move_range(m.group(3)))

        if shift == 0 and all(len(v) == 1 and v[0] == k for k, v in mapping.items()):
            moved = False
        else:
            moved = True
        # conditional formatting
        if moved:
            from openpyxl.formatting.formatting import ConditionalFormattingList
            rules = [(cf.sqref, list(cf.rules)) for cf in ws.conditional_formatting]
            ws.conditional_formatting = ConditionalFormattingList()
            for sqref, cf_rules in rules:
                for rule in cf_rules:
                    ws.conditional_formatting.add(move_ranges(sqref), rule)
            for dv in ws.data_validations.dataValidation:
                dv.sqref = MultiCellRange(move_ranges(dv.sqref))
            if ws.auto_filter.ref:
                ws.auto_filter.ref = move_range(ws.auto_filter.ref)
            if ws.freeze_panes:
                ws.freeze_panes = move_range(ws.freeze_panes, expand=False)
            area = ws.print_area
            if area:
                areas = [a.split('!')[-1] for a in (area if isinstance(area, (list, tuple)) else area.split(','))]
                ws.print_area = [move_range(a) for a in areas]
        # charts: anchor and data ranges; cached values are dropped so that the chart shows the new data
        for chart in ws._charts:
            anchor = getattr(chart, 'anchor', None)
            start = getattr(anchor, '_from', None)
            if moved and start is not None:
                new = row_of(start.row + 1) - 1
                end = getattr(anchor, 'to', None)
                if end is not None:
                    end.row += new - start.row
                start.row = new
            for series in getattr(chart, 'series', []) or []:
                for name in CHART_REFS:
                    holder = getattr(series, name, None)
                    if holder is None:
                        continue
                    for kind, cache in (('numRef', 'numCache'), ('strRef', 'strCache'), ('multiLvlStrRef', 'multiLvlStrCache')):
                        ref = getattr(holder, kind, None)
                        if ref is not None and ref.f:
                            ref.f = move_formula(ref.f)
                            setattr(ref, cache, None)

    def _path_records(self, source, scope):
        names = source.split('.')
        roots = self._roots(scope)
        if names[0] not in roots:
            raise TemplateError("'%s' is unknown in the loop %s" % (names[0], source))
        current = roots[names[0]]
        for name in names[1:]:
            if not hasattr(current, '_fields') or name not in current._fields:
                raise TemplateError("The field '%s' cannot be used in the loop %s" % (name, source))
            current = current.mapped(name)
        if not hasattr(current, '_name'):
            raise TemplateError("The loop %s must go over records (a one2many or many2many field)" % source)
        return current

    # ------------------------------------------------------------------ workbook
    def render_workbook(self, template_bytes, mode='single'):
        """Render the template for ``self.records``.

        :param mode: ``single`` (one workbook, ``object`` is the first record, ``records`` all of them),
            ``sheets`` (the template sheets are repeated for each record)
        :returns: the rendered workbook as bytes
        """
        wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
        if mode == 'sheets' and len(self.records) > 1:
            originals = list(wb.worksheets)
            # pictures of the template, read once (openpyxl closes the source after reading it)
            pictures = {id(ws): [(img, img._data()) for img in ws._images] for ws in originals}
            charts = {id(ws): [deepcopy(chart) for chart in ws._charts] for ws in originals}
            jobs = []
            for index, record in enumerate(self.records):
                for original in originals:
                    ws = original if index == 0 else wb.copy_worksheet(original)
                    ws._images = [self._clone_image(img, raw) for img, raw in pictures[id(original)]]
                    jobs.append((ws, original, record, index))
            source_titles = {id(ws): ws.title for ws in originals}
            for ws, original, record, index in jobs:
                if index:
                    ws._charts = [deepcopy(chart) for chart in charts[id(original)]]
            for ws, original, record, index in jobs:
                title = self.render_text('{{ object.display_name }}', {'object': record})[:25] or str(index + 1)
                ws.title = self._unique_title(wb, ('%s %s' % (source_titles[id(original)], title))
                                              if len(originals) > 1 else title, ws)
                self.render_sheet(ws, {'object': record}, source_title=source_titles[id(original)])
        else:
            scope = {'object': self.records[:1]}
            for ws in wb.worksheets:
                self.render_sheet(ws, scope)
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    def _clone_image(self, img, raw):
        buffer = io.BytesIO(raw)
        self._images.append(buffer)
        clone = XLImage(buffer)
        clone.width, clone.height = img.width, img.height
        clone.anchor = copy(img.anchor)
        return clone

    @staticmethod
    def _unique_title(wb, title, ws):
        title = re.sub(r'[\[\]:*?/\\]', ' ', title).strip()[:31] or 'Sheet'
        existing = {s.title for s in wb.worksheets if s is not ws}
        base, n = title, 2
        while title in existing:
            suffix = ' (%s)' % n
            title = base[:31 - len(suffix)] + suffix
            n += 1
        return title


def preview_html(xlsx_bytes, max_rows=60, max_cols=20, texts=None):
    """A small HTML table of the first sheet, for the print dialog preview.

    :param texts: translated notes, keys ``more`` (with %s), ``pictures`` (with %s) and ``formulas``
    """
    texts = dict({'more': '%s more rows in the file.', 'pictures': '%s picture(s) in the sheet.',
                  'formulas': 'Formulas, shown in grey, are calculated when the file opens.'}, **(texts or {}))
    from markupsafe import escape
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=False)
    ws = wb.worksheets[0]
    rows = []
    formulas = False
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_rows), max_col=min(ws.max_column, max_cols)):
        cells = []
        for cell in row:
            value = cell.value
            style = []
            if isinstance(value, datetime.datetime):
                value = value.strftime('%Y-%m-%d %H:%M')
            elif isinstance(value, datetime.date):
                value = value.isoformat()
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                decimals = len(cell.number_format.split('.')[1].split(';')[0].rstrip('_)"%')) \
                    if '.' in (cell.number_format or '') else (2 if value % 1 else 0)
                value = ('{:,.%df}' % min(decimals, 6)).format(value)
                style.append('text-align:right')
            elif isinstance(value, str) and value.startswith('='):
                style.append('color:#6b7280;font-style:italic;text-align:right')
                formulas = True
            if cell.font is not None and cell.font.b:
                style.append('font-weight:600')
            cells.append('<td%s>%s</td>' % (' style="%s"' % ';'.join(style) if style else '',
                                             escape('' if value is None else value)))
        rows.append('<tr>%s</tr>' % ''.join(cells))
    more = ''
    if ws.max_row > max_rows:
        more = '<p class="text-muted small mb-0">%s</p>' % escape(texts['more'] % (ws.max_row - max_rows))
    images = len(ws._images)
    note = '<p class="text-muted small mb-0">%s</p>' % escape(texts['pictures'] % images) if images else ''
    if formulas:
        note += '<p class="text-muted small mb-0">%s</p>' % escape(texts['formulas'])
    return ('<div class="o_axs_xlsx_preview"><table class="table table-sm table-bordered mb-1">%s</table>%s%s</div>'
            % (''.join(rows), more, note))


def column_letter(index):
    return get_column_letter(index)
