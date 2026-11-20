"""识别同类 Excel 账表，保留表头与原始行的位置。"""
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import math
import posixpath
import re
import unicodedata
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.utils.datetime import from_excel


class 输入错误(ValueError):
    pass


def 规范(value):
    text = unicodedata.normalize('NFKC', str(value or '')).casefold().replace('帐','账')
    return ''.join(c for c in text if c.isalnum())


def 文本(value):
    return str(value).strip() if value is not None else ''


def 有值(value):
    return 文本(value) not in ('', '0', '-', '--', '—', '缺省', '默认', '#N/A', '#REF!', '#VALUE!')


字段 = {
    '组合': '账户组合|账户组合编码|账户编码组合|会计账户组合|账户编号|账户编码|科目组合|账户组合编号|account combination|code combination|account number',
    '描述': '账户描述|账户说明|账户名称|组合说明|组合描述|账户组合描述|account description|combination description',
    '科目编码': '科目|科目编码|会计科目|会计科目编码|科目编号|科目段编码|natural account',
    '科目名称': '科目段|科目名称|会计科目名称|科目说明|natural account description',
    '日期': '有效日期|记账日期|凭证日期|会计日期|业务日期|日期|effective date|accounting date|journal date',
    '凭证号': '凭证号|凭证号码|凭证编号|凭证序号|journal number|voucher number',
    '摘要': '行摘要|分录摘要|摘要|分录说明|凭证摘要|line description|journal line description',
    '凭证类型': '凭证类型|凭证类别|凭证字|journal category|voucher type',
    '行号': '行号|分录行号|凭证行号|line number',
    '期间': '期间|会计期间|记账期间|会计月份|period|accounting period',
    '公司': '公司段|公司名称|记账公司|公司|company description',
    '部门': '部门段|部门名称|部门|成本中心名称|cost center description|department',
    '子科目': '子科目段|子科目名称|子科目|subaccount description',
    '产品': '产品段|产品名称|product description',
    '项目': '项目段|项目名称|project description',
    '往来': '往来段|往来单位|往来对象|往来名称|往来段名称|counterparty description',
    '银行': '银行段|银行名称|bank name',
    '借方': '本位币借方金额|本位币借方|本期借方|本期借方发生额|本期发生额借方|本年发生额借方|累计发生额借方|累计借方发生额|借方发生额|借方金额|借方|accounted debit|debit',
    '贷方': '本位币贷方金额|本位币贷方|本期贷方|本期贷方发生额|本期发生额贷方|本年发生额贷方|累计发生额贷方|累计贷方发生额|贷方发生额|贷方金额|贷方|accounted credit|credit',
    '期初': '期初|期初余额|期初净额|开始余额|opening balance|beginning balance',
    '期末': '期末|期末余额|期末净额|结束余额|ending balance|closing balance',
    '期初借方': '期初借方|期初余额借方|期初余额借方金额|期初借方余额|opening debit',
    '期初贷方': '期初贷方|期初余额贷方|期初余额贷方金额|期初贷方余额|opening credit',
    '期末借方': '期末借方|期末余额借方|期末余额借方金额|期末借方余额|closing debit',
    '期末贷方': '期末贷方|期末余额贷方|期末余额贷方金额|期末贷方余额|closing credit',
}
字段 = {key: {规范(x) for x in names.split('|')} for key, names in 字段.items()}
公共 = {'组合', '借方', '贷方'}
日期级别 = {'有效日期': 3, '会计日期': 2, '记账日期': 2, '凭证日期': 1, '业务日期': -2, '日期': -1}
金额字段 = {'借方', '贷方', '期初', '期末', '期初借方', '期初贷方', '期末借方', '期末贷方'}
单位 = {'元', '本位币', '记账本位币', '人民币', '金额', '人民币元', 'cny', 'rmb', '单位元'}


@dataclass
class 位置:
    文件: str
    页: str
    行: int

    def __str__(self):
        return f'{self.文件} / {self.页} 第{self.行}行'


@dataclass
class 行:
    位置: 位置
    编码: tuple
    名称: dict
    金额: tuple
    日期: object = None
    凭证: str = ''
    摘要: str = ''
    类型: str = ''
    期间: str = ''
    年份: str = ''
    行号: str = ''


@dataclass
class 表:
    类别: str
    文件: str
    页: str
    表头: int
    字段: dict
    行: list
    合计: list
    期间: str
    警示: list


def 工作表部件(path):
    with ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        links = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        parts = {el.get('Id'): el.get('Target') for el in links}
    for el in workbook.iter():
        if el.tag.rsplit('}', 1)[-1] == 'sheet':
            key = next((v for k, v in el.attrib.items() if k.endswith('}id')), None)
            target = parts.get(key)
            if target:
                yield el.get('name'), (
                    target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
                )


def 读取布局(path, part):
    merges, empty_formulas = [], set()
    with ZipFile(path) as archive, archive.open(part) as stream:
        parents = []
        for event, node in ET.iterparse(stream, events=('start', 'end')):
            if event == 'start':
                parents.append(node)
                continue
            tag = node.tag.rsplit('}', 1)[-1]
            if tag == 'mergeCell':
                merges.append(range_boundaries(node.get('ref')))
            if tag == 'c' and any(child.tag.endswith('}f') for child in node):
                raw = next((child.text for child in node if child.tag.endswith('}v')), None)
                if raw is None:
                    empty_formulas.add(node.get('r'))
            # 每行读完就释放，避免大账表占用过多内存。
            if len(parents) > 1 and tag in ('row', 'mergeCell'):
                parents[-2].remove(node)
            parents.pop()
    return merges, empty_formulas


def 展开表头(raw, merges):
    height = len(raw)
    width = max((len(r) for r in raw), default=0)
    grid = [list(r) + [None] * (width - len(r)) for r in raw]
    for left, top, right, bottom in merges:
        if top > height or left > width:
            continue
        value = grid[top-1][left-1]
        for r in range(top-1, min(bottom, height)):
            for col in range(left-1, min(right, width)):
                grid[r][col] = value
    return grid


def 列候选(kind, grid, extra):
    hits = defaultdict(list)
    for col, levels in enumerate(zip(*grid)):
        names = list(dict.fromkeys(v for v in levels if v))
        names = [val for val in names if val not in 单位 or val in ('本位币', '人民币')]
        if not names:
            continue
        meaningful = [n for n in names if n not in 单位]
        if not meaningful:
            continue
        leaf = meaningful[-1]
        joined = ''.join(meaningful)
        path = ''.join(names)
        if any(term in path for term in ('外币', '原币', '交易币', 'entered', 'foreigncurrency')):
            foreign = True
        else:
            foreign = False
        end = any(x in leaf or x in joined for x in ('期初', '期末', 'opening', 'closing', 'beginning', 'ending'))
        for key, builtins in 字段.items():
            targets = builtins | extra.get(key, set())
            score = 6 if leaf in targets else 0
            if joined in targets and joined != leaf:
                score = max(score, 7)
            if leaf.removesuffix('元') in targets:
                score = max(score, 5)
            if not score:
                continue
            if key in 金额字段:
                if foreign or (key in ('借方', '贷方') and end):
                    continue
                if key in ('期初', '期末') and leaf in 字段['借方'] | 字段['贷方']:
                    continue
                if key in ('借方', '贷方') and kind == '余额表' and '本年' not in path and '累计' not in path and '本月' in path:
                    continue
                if '本位币' in path or 'accounted' in path:
                    score += 4
                elif '人民币' in path:
                    score += 2
                if kind == '余额表' and any(n in path for n in ('本期', '本年', '累计')):
                    score += 1
            if key == '日期':
                score += 日期级别.get(leaf, 0)
            hits[key].append((score, col))
    return hits


def 匹配表头(kind, grid, extra, limits):
    best = None
    combo_aliases = 字段['组合'] | extra.get('组合', set())
    for end in range(1, min(len(grid), limits) + 1):
        for start in range(max(0, end-4), end):
            block = grid[start:end]
            if not any(v in combo_aliases for row in block for v in row):
                continue
            matches = 列候选(kind, block, extra)
            if '组合' not in matches:
                continue
            required = 公共 | ({'日期', '凭证号'} if kind == '序时账' else set())
            if kind == '余额表':
                for base in ('期初', '期末'):
                    paired = {base+'借方', base+'贷方'}
                    if paired <= matches.keys():
                        if base in matches:
                            continue
                        required |= paired
                    else:
                        required.add(base)
            if not required <= matches.keys():
                continue
            selected = {}
            for field, choices in matches.items():
                high = max(score for score, _ in choices)
                selected[field] = (high, [col for score, col in choices if score == high])
            relevant = (set(字段) - {'期初','期末','期初借方','期初贷方','期末借方','期末贷方'}) if kind == '序时账' else (
                set(字段) - {'日期','凭证号','期间','行号','摘要','凭证类型'}
            )
            score = (sum(selected[key][0] for key in required), len(relevant & selected.keys()), -end, start)
            if best is None or score > best[0]:
                best = (score, end, selected, required)
    return best


def 定位表(path, book, kind, forced, limit, extra):
    parts = dict(工作表部件(path))
    candidates = []
    for sheet in book.worksheets:
        if forced and sheet.title != forced:
            continue
        part = parts.get(sheet.title)
        if not part:
            continue
        sheet.reset_dimensions()
        raw = list(sheet.iter_rows(min_row=1, max_row=limit, values_only=True))
        # 先找账户组合别名，避开封面和纯空表。
        if not any(规范(v) in 字段['组合'] | extra.get('组合', set()) for row in raw for v in row):
            continue
        merges, formulas = 读取布局(path, part)
        grid = [[规范(v) for v in row] for row in 展开表头(raw, merges)]
        for option in ([kind] if kind else ['序时账', '余额表']):
            hit = 匹配表头(option, grid, extra, limit)
            if hit:
                candidates.append((option, sheet, hit, merges, formulas, raw))
    if not candidates:
        sheets = '、'.join(book.sheetnames)
        raise 输入错误(f'{path.name} 未找到完整的{kind or "序时账或余额表"}表头。'
                  f'已查看前{limit}行，工作表：{sheets}。请确认账户组合、借贷本位币列和对应的日期凭证或期初期末列。')
    if len(candidates) != 1:
        options = '；'.join(f'{sheet.title} 第{hit[1]}行（{option}）' for option, sheet, hit, *_ in candidates)
        raise 输入错误(f'{path.name} 有多个表或格式候选：{options}。请确认工作表和表类型。')
    return candidates[0]


组合格式 = re.compile(r'[0-9A-Za-z_/-]+(?:[.．][0-9A-Za-z_/-]+){9}\Z')


def 解析描述(description, parts, explicit):
    body = 文本(description)
    if not body:
        return {}
    pieces = re.split(r'(?<!\\)[.．]', body)
    names = {}
    active = [i for i,c in enumerate(parts) if i in (0, 2, 3, 4, 6) or c != '0']
    if len(pieces) == len(active) and len(active) < 10:
        for j, i in enumerate(active):
            end = active[j+1] if j+1 < len(active) else 10
            count = end-i-1
            suffix = '-' * count
            if count and not pieces[j].endswith(suffix):
                break
            value = pieces[j][:-count] if count else pieces[j]
            if parts[i] == '0' and 有值(value):
                break
            names[i+1] = value.strip().replace('\\.', '.')
        else:
            return names
    if len(pieces) > 10:
        # 旧导出在合同利率中也使用普通点号，只有这能还原整十段时才采用。
        decimal_safe = re.split(r'(?<!\\)(?<!\d)[.．]|(?<!\\)[.．](?!\d)', body)
        if len(decimal_safe) == 10 and all(
            有值(value) == (parts[i] != '0') for i,value in enumerate(decimal_safe) if i != 2
        ):
            pieces = decimal_safe
    if len(pieces) == 10:
        return {i+1: p.strip().replace('\\.', '.') for i,p in enumerate(pieces)}
    if not 10 < len(pieces) <= 35:
        return {}
    # 全点号描述可能把利率、小数或公司名称中的点写成普通点号。只采用能唯一分段的结果。
    results = []

    def search(seg, offset, record):
        if len(results) > 1:
            return
        if seg == 10:
            if offset == len(pieces):
                results.append(record)
            return
        end = len(pieces) - (9-seg)
        for stop in range(offset+1, end+1):
            name = '.'.join(pieces[offset:stop]).strip().replace('\\.', '.')
            if (parts[seg] == '0') != (not 有值(name)) and seg != 2:
                continue
            anchor = explicit.get(seg+1)
            if 有值(anchor) and 规范(name) != 规范(anchor):
                continue
            search(seg+1, stop, {**record, seg+1: name})

    search(0, 0, {})
    return results[0] if len(results) == 1 else {}


def 金额(value, position, column, precision):
    label = f'{position} {get_column_letter(column+1)}{position.行}'
    if value is None or (isinstance(value, str) and not value.strip()):
        return 0
    raw = unicodedata.normalize('NFKC', str(value)).strip().replace('−', '-')
    if raw.startswith('(') and raw.endswith(')'):
        raw = '-' + raw[1:-1].strip()
    if raw in ('-', '—', '–'):
        return 0
    if ',' in raw:
        if not re.fullmatch(r'[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?', raw):
            raise 输入错误(f'{label} 金额分组格式无法确认：{value}')
        raw = raw.replace(',', '')
    try:
        result = Decimal(raw)
        scaled = result.scaleb(precision)
        integer = scaled.to_integral_value()
    except InvalidOperation as exc:
        raise 输入错误(f'{label} 无法读取金额：{value}') from exc
    if not result.is_finite():
        raise 输入错误(f'{label} 无法读取金额：{value}')
    tolerance = Decimal(0)
    if isinstance(value, float):
        tolerance = Decimal(str(math.ulp(value) * 2)).scaleb(precision)
    if abs(scaled - integer) > tolerance:
        raise 输入错误(f'{label} 有超过{precision}位的有效小数，请确认金额精度：{value}')
    if abs(integer) >= 10**15:
        raise 输入错误(f'{label} 金额超出本次Excel精度范围。')
    return int(integer)


def 日期(value, epoch, loc):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raw = 文本(value)
    if isinstance(value, (int, float)) and 0 < value < 100000:
        interpreted = from_excel(value, epoch)
        if isinstance(interpreted, datetime):
            return interpreted.date()
    raw = re.sub(r'[年月./]', '-', raw).replace('日', '')
    try:
        if re.fullmatch(r'\d{8}', raw):
            return datetime.strptime(raw, '%Y%m%d').date()
        if re.fullmatch(r'\d{4}-\d{1,2}-\d{1,2}(?:[ T].*)?', raw):
            return date(*map(int, re.split(r'[- T]', raw, maxsplit=3)[:3]))
    except (ValueError, OverflowError):
        pass
    raise 输入错误(f'{loc} 日期无法识别：{value}')


def 标识(cell, loc):
    value = cell.value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if int(value) != value or abs(value) >= 10**15:
            raise 输入错误(f'{loc} 编号不是可精确读取的整数，请从源系统以文本导出。')
        fmt = cell.number_format or ''
        return str(int(value)).zfill(len(fmt)) if re.fullmatch(r'0{2,}', fmt) else str(int(value))
    return 文本(value)


def 读表(path, kind=None, forced=None, limit=120, extra=None, assigned=None, precision=2):
    book = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        return _读表(path, book, kind, forced, limit, extra or {}, assigned or {}, precision)
    finally:
        book.close()


def _读表(path, book, kind, forced, limit, extra, assigned, precision):
    kind, sheet, hit, merges, formulas, raw = 定位表(path, book, kind, forced, limit, extra)
    _, end, matches, required = hit
    columns = {}
    assigned = assigned.get(kind, {})
    for field,col in assigned.items():
        if field not in matches or not 0 <= col < max(map(len,raw)):
            raise 输入错误(f'{path.name} / {sheet.title}：「{field}」指定列超出表头范围或字段尚未识别。')
    for field, (_, choices) in matches.items():
        if field in assigned:
            columns[field] = assigned[field]
            continue
        if len(choices) != 1:
            if field in required or field in ('描述','科目名称','科目编码','公司','部门','子科目','项目','往来','银行','摘要','凭证类型','期间'):
                letters = '、'.join(f'{get_column_letter(c+1)}列' for c in choices)
                raise 输入错误(f'{path.name} / {sheet.title} 表头第{end}行，「{field}」匹配到{letters}，请确认使用哪列。')
            continue
        columns[field] = choices[0]
    if len(set(columns.values())) != len(columns):
        raise 输入错误(f'{path.name} / {sheet.title}：多个字段匹配到了同一列，请核对表头或指定列。')
    for base in ('期初', '期末'):
        if kind == '余额表' and base in columns and {base+'借方',base+'贷方'} <= columns.keys():
            raise 输入错误(f'{path.name} / {sheet.title} 同时有带符号{base}与{base}借贷两列，请确认使用哪组。')
    direct = {'公司':1, '部门':3, '科目名称':4, '子科目':5, '产品':6, '项目':7, '往来':8, '银行':9}
    width = max(max(columns.values())+1, min(12, max(map(len, raw))))
    rows, totals, warnings = [], [], []
    span_start = defaultdict(list)
    active = {}
    for left, top, right, bottom in merges:
        if top <= end:
            continue
        affected = [name for name,col in columns.items() if left <= col+1 <= right]
        if not affected:
            continue
        if left != right or (affected[0] in 金额字段 and top != bottom):
            raise 输入错误(f'{path.name} / {sheet.title} 第{top}行：数据区合并跨越{affected[0]}列，请拆分后确认原值。')
        if top != bottom:
            span_start[top].append((left-1,bottom))
    amount_keys = ['借方','贷方'] if kind == '序时账' else ['期初','借方','贷方','期末']
    used_amounts = [key for key in columns if key in 金额字段 and (kind == '余额表' or key in amount_keys)]
    top_text = ' '.join(文本(v) for row in raw[:end] for v in row if isinstance(v, str))
    amount_headers = ' '.join(文本(row[columns[key]]) for row in raw[:end]
                              for key in used_amounts if columns[key] < len(row))
    if re.search(r'单位\s*[:：]?\s*(?:人民币\s*)?[万千百]元', top_text) or re.search(r'[万千百]元', amount_headers):
        raise 输入错误(f'{path.name} / {sheet.title} 的报表金额单位不是元，请确认如何换算后再处理。')
    period = ''
    if kind == '余额表':
        match = re.search(r'((?:19|20)\d{2})年\s*(\d{1,2})月?\s*(?:至|到|[-—~～])\s*(?:((?:19|20)\d{2})年)?\s*(\d{1,2})月', top_text)
        if match:
            a, b, c, d = match.groups()
            period = f'{a}-{int(b):02d}~{c or a}-{int(d):02d}'
        else:
            match = re.search(r'((?:19|20)\d{2})[-/年](\d{1,2})月?\s*(?:至|到|[~～])\s*((?:19|20)\d{2})[-/年](\d{1,2})', top_text)
            if match:
                a,b,c,d = match.groups()
                period = f'{a}-{int(b):02d}~{c}-{int(d):02d}'
    for number, cells in enumerate(sheet.iter_rows(min_row=end+1, max_col=width), end+1):
        pos = 位置(path.name, sheet.title, number)
        values = [cell.value for cell in cells]
        for col,bottom in span_start.get(number, []):
            active[col] = (bottom, cells[col])
        for col,(bottom,cell) in tuple(active.items()):
            if number > bottom:
                del active[col]
            else:
                values[col] = cell.value
        def val(key):
            return values[columns[key]] if key in columns else None

        for key in required | set(used_amounts) | ({'科目名称','描述','科目编码'} & columns.keys()):
            col = columns[key]
            if f'{get_column_letter(col+1)}{number}' in formulas:
                raise 输入错误(f'{pos} {get_column_letter(col+1)}{number} 的{key}公式没有已保存的计算结果。请在Excel重新计算并保存源文件。')
        combo = 文本(val('组合'))
        if 规范(combo) in 字段['组合'] | extra.get('组合', set()):
            for key in required | set(used_amounts):
                col = columns[key]
                names = 字段[key] | extra.get(key, set())
                found = [i for i,cell in enumerate(values) if 规范(cell) in names]
                if found and col not in found:
                    raise 输入错误(f'{pos} 重复表头的「{key}」列发生变化，请确认是否分开清洗。')
            continue
        money_values = [val(key) for key in used_amounts]
        if not combo and all(v is None or not 文本(v) or 规范(v) in 单位 for v in money_values):
            if val('日期') or val('凭证号'):
                raise 输入错误(f'{pos} 有凭证信息，账户组合为空。')
            continue

        def amount(key):
            if key in columns:
                return 金额(val(key), pos, columns[key], precision)
            return amount(key+'借方') - amount(key+'贷方')

        if not 组合格式.fullmatch(combo):
            rollup = next((文本(x) for x in values if re.fullmatch(r'(?:总计|合计|小计|累计|本页合计|本期合计|本年合计|grand total|total)\s*[:：]?', 文本(x), re.I)), '')
            if rollup and not (val('日期') or val('凭证号')):
                if '小计' not in rollup and '累' not in rollup and '本页' not in rollup:
                    totals.append((pos, tuple(amount(key) for key in amount_keys)))
                continue
            raise 输入错误(f'{pos} 的账户组合无法识别，预期为10段编码：{combo or "空白"}。')
        parts = tuple(re.split('[.．]', combo))
        amounts = tuple(amount(key) for key in amount_keys)
        other_code = 文本(val('科目编码'))
        if other_code and other_code != parts[3]:
            raise 输入错误(f'{pos} 科目列为 {other_code}，账户组合第4段为 {parts[3]}，请确认口径。')
        explicit = {segment:文本(val(field)) for field,segment in direct.items() if 有值(val(field))}
        names = 解析描述(val('描述'), parts, explicit)
        for seg, name in explicit.items():
            alternative = names.get(seg)
            if 有值(alternative) and 规范(alternative) != 规范(name):
                warnings.append((pos, '分列名称与组合说明不同', f'第{seg}段采用分列：{name}，描述：{alternative}'))
            names[seg] = name
        row = 行(pos, parts, names, amounts)
        if kind == '序时账':
            row.日期 = 日期(val('日期'), book.epoch, pos)
            col = columns['凭证号']
            row.凭证 = 标识(active[col][1] if col in active else cells[col], pos)
            if not 有值(row.凭证):
                raise 输入错误(f'{pos} 凭证号为空或无效。')
            row.摘要 = 文本(val('摘要'))
            row.类型 = 文本(val('凭证类型'))
            row.期间 = 文本(val('期间')) or row.日期.strftime('%Y-%m')
            row.年份 = str(row.日期.year)
            row.行号 = 文本(val('行号'))
        else:
            row.年份 = period[:4] if period and period[:4] == period.split('~')[1][:4] else ''
        rows.append(row)
    if not rows:
        raise 输入错误(f'{path.name} / {sheet.title} 找到表头，但没有可转换的账户明细。')
    if kind == '序时账':
        periods = sorted({row.日期.strftime('%Y-%m') for row in rows})
        period = f'{periods[0]}~{periods[-1]}'
    return 表(kind, path.name, sheet.title, end, columns, rows, totals, period, warnings)
