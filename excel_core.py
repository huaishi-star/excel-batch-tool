"""Excel 批量处理核心逻辑（GUI 无关，可独立测试）。

提供两大能力：
1. 批量重命名：按"单元格组合"规则生成新文件名，复制到输出文件夹（原文件不动）
2. 模板填充：把源文件的单元格值按映射关系填入模板对应单元格，保留模板格式

注意：
- 单元格定位一律使用文件/模板的第一个工作表（wb.worksheets[0]）
- 读取使用 data_only=True, read_only=True；模板填充写入使用完整模式以保留格式
"""

from __future__ import annotations

import re
import shutil
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from threading import Event
from typing import Callable, Iterable, Optional

import xlrd
import xlwt
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

# ---------------------------------------------------------------------------
# 常量与正则
# ---------------------------------------------------------------------------

# Windows 文件名非法字符
ILLEGAL_CHARS_RE = re.compile(r'[<>:"/\\|?*]')
# 控制字符（换行、制表等）
CONTROL_CHARS_RE = re.compile(r'[\x00-\x1f]')
# Windows 保留设备名
RESERVED_NAMES = (
    {'CON', 'PRN', 'AUX', 'NUL'}
    | {f'COM{i}' for i in range(1, 10)}
    | {f'LPT{i}' for i in range(1, 10)}
)
# 单元格坐标，如 A1、AB12
CELL_RE = re.compile(r'^[A-Za-z]{1,3}[0-9]{1,7}$')
# 列字母，如 A、B、AA
COL_RE = re.compile(r'^[A-Za-z]{1,3}$')
# 坐标匹配：列字母 + 行号
_COORD_MATCH = re.compile(r'^([A-Za-z]+)([0-9]+)$')

# 整列读取的行数上限（安全阀）
MAX_COL_ROWS = 5000


def _col_letter_to_idx(col: str) -> int:
    """列字母 A、B、AA → 0 基列号。"""
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch.upper()) - 64)
    return idx - 1


def _coord_to_rc(coord: str) -> tuple:
    """'A1' → (0 基行号, 0 基列号)。"""
    m = _COORD_MATCH.match(coord)
    return int(m.group(2)) - 1, _col_letter_to_idx(m.group(1))


def _openpyxl_merged_anchor(ws, coord: str) -> str:
    """若 coord 位于某个合并区域内，返回其左上角(锚点)坐标；否则原样返回。

    仅用于"单元格→单元格"映射：向合并区写入时把值放到锚点，保证显示。
    """
    col_letter, row = coordinate_from_string(coord)
    col = column_index_from_string(col_letter)
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
            return f'{get_column_letter(mr.min_col)}{mr.min_row}'
    return coord


def _openpyxl_merged_range(ws, coord: str):
    """若 coord 位于某个合并区域内，返回该区域对象；否则返回 None。"""
    col_letter, row = coordinate_from_string(coord)
    col = column_index_from_string(col_letter)
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
            return mr
    return None


def _shift_openpyxl_merges(ws, from_row: int, amount: int):
    """插入行后，把 from_row 及以下的合并区域整体下移 amount 行（跨插入点的区域向下扩展）。"""
    if amount <= 0:
        return
    for mr in ws.merged_cells.ranges:
        if mr.min_row >= from_row:
            mr.shift(0, amount)
        elif mr.min_row < from_row <= mr.max_row:
            mr.max_row += amount


def _xls_merged_anchor(merged: list, r: int, c: int) -> tuple:
    """若 (r,c) 位于某个合并区域内（0 基），返回其左上角(锚点)坐标；否则原样返回。"""
    for (rlo, rhi, clo, chi) in merged:
        if rlo <= r < rhi and clo <= c < chi:
            return rlo, clo
    return r, c


def _xlrd_cell_value(book, sheet, r: int, c: int):
    """读取 xlrd 单元格为 Python 值：日期→datetime、布尔→bool、空→None。

    越界（r/c 超出工作表使用范围）按空单元格处理，
    避免 xlrd 对越界访问抛 IndexError（openpyxl 则默认返回空）。
    """
    if r >= sheet.nrows or c >= sheet.ncols:
        return None
    cell = sheet.cell(r, c)
    if cell.ctype == xlrd.XL_CELL_DATE:
        return xlrd.xldate.xldate_as_datetime(cell.value, book.datemode)
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return bool(cell.value)
    if cell.ctype == xlrd.XL_CELL_EMPTY:
        return None
    return cell.value

# 文件名主干最长字符数（Windows 单段建议 ≤255，留余量）
MAX_STEM = 150
# 完整路径最长字符数（避开 Windows MAX_PATH=260，留余量）
MAX_PATH = 250

LogFn = Callable[[str], None]
ProgressFn = Callable[[int, int], None]


# ---------------------------------------------------------------------------
# 数据类
# ---------------------------------------------------------------------------

@dataclass
class PlanItem:
    """重命名计划中的一项。"""
    src: Path
    dst: Path
    name: str  # dst 的文件名，供预览


@dataclass
class FillItem:
    """模板填充计划中的一项。"""
    src: Path
    dst: Path


# ---------------------------------------------------------------------------
# 单元格值 → 文本
# ---------------------------------------------------------------------------

def value_to_text(value, default: str = "") -> str:
    """把单元格值转换成适合用于文件名的显示文本。

    - None             → default（通常为空串，拼接时会被跳过）
    - datetime / date  → %Y%m%d（如 20240115，避免出现 / 或 - 等非法/歧义字符）
    - bool             → TRUE / FALSE（与 Excel 显示一致）
    - 整型浮点          → 去掉 .0（98.0 → '98'；98.5 → '98.5'）
    - 其余              → str(value)
    """
    if value is None:
        return default
    if isinstance(value, (datetime, date)):
        return value.strftime('%Y%m%d')
    if isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def read_cell_text(path, coord: str) -> str:
    """读取文件第一个工作表中 coord 单元格的显示文本（供重命名用），支持 .xls/.xlsx/.xlsm。"""
    if Path(path).suffix.lower() == '.xls':
        book = xlrd.open_workbook(str(path))
        try:
            r, c = _coord_to_rc(coord)
            return value_to_text(_xlrd_cell_value(book, book.sheet_by_index(0), r, c))
        finally:
            book.release_resources()
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        return value_to_text(wb.worksheets[0][coord].value)
    finally:
        wb.close()


def read_cell_raw(path, coord):
    """读取文件第一个工作表中 coord 单元格的原始 Python 值（供模板填充用）。

    None 会转成空串，以便写入时清空模板对应单元格。支持 .xls/.xlsx/.xlsm。
    """
    if Path(path).suffix.lower() == '.xls':
        book = xlrd.open_workbook(str(path))
        try:
            r, c = _coord_to_rc(coord)
            v = _xlrd_cell_value(book, book.sheet_by_index(0), r, c)
            return '' if v is None else v
        finally:
            book.release_resources()
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        v = wb.worksheets[0][coord].value
        return '' if v is None else v
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# 文件名安全处理
# ---------------------------------------------------------------------------

def sanitize_name_part(part: str) -> str:
    """净化文件名片段：
    - 去除控制字符
    - 非法字符替换为下划线
    - 去掉首尾空格与点号
    - 若是 Windows 保留设备名，加下划线前缀
    """
    s = CONTROL_CHARS_RE.sub('', part)
    s = ILLEGAL_CHARS_RE.sub('_', s)
    s = s.strip(' .')
    if not s:
        return s
    if s.split('.')[0].upper() in RESERVED_NAMES:
        s = '_' + s
    return s


def truncate_stem(stem: str, max_stem: int = MAX_STEM) -> str:
    """把主文件名截断到 max_stem 字符，并保证结尾不是点或空格。"""
    if len(stem) <= max_stem:
        return stem
    return stem[:max_stem].rstrip(' .')


def unique_path(dst: Path, used: set[str]) -> Path:
    """目标路径若已存在（磁盘上或本次任务中已用过），自动追加 (1)、(2)…"""
    if dst.name.lower() not in used and not dst.exists():
        return dst
    stem, ext = dst.stem, dst.suffix
    n = 1
    while True:
        candidate = dst.with_name(f'{stem} ({n}){ext}')
        if candidate.name.lower() not in used and not candidate.exists():
            return candidate
        n += 1


def build_dest_name(stem: str, ext: str, out_dir, used: set[str],
                    max_stem: int = MAX_STEM, max_path: int = MAX_PATH) -> Path:
    """根据文件名主干和扩展名构造输出路径。

    依次处理：净化 → 截断 → 保证完整路径长度 → 重名去重。
    """
    stem = truncate_stem(sanitize_name_part(stem), max_stem)
    if not stem:
        stem = '未命名'
    dst = Path(out_dir) / f'{stem}{ext}'
    # 完整路径过长时，逐段缩短主干
    while len(str(dst)) > max_path and len(stem) > 5:
        stem = stem[:-5].rstrip(' .')
        dst = Path(out_dir) / f'{stem}{ext}'
    return unique_path(dst, used)


# ---------------------------------------------------------------------------
# 批量重命名
# ---------------------------------------------------------------------------

def build_name(parts: list[tuple], separator: str) -> str:
    """按规则元素拼接文件名主干。

    parts: [(kind, value)]，kind ∈ {'text', 'data'}。
    - 'text'：手打文字，原样保留；
    - 'data'：单元格内容或原文件名，仅当非空时加入；且只在相邻两个 data 之间插入分隔符。
    """
    out: list[str] = []
    prev_data = False
    for kind, value in parts:
        if kind == 'text':
            out.append(value)
            prev_data = False
        else:  # data
            if value:
                if prev_data and separator:
                    out.append(separator)
                out.append(value)
                prev_data = True
    return ''.join(out)


def plan_rename(files: Iterable, rule: list, separator: str, out_dir,
                ) -> tuple[list[PlanItem], list[tuple]]:
    """生成重命名计划（不执行）。

    rule 是规则元素列表，每个元素为 (kind, value)：
      ('cell', 'A1')     → 读取文件第一个工作表该单元格的显示文本
      ('text', '文字')   → 原样保留的手打文字
      ('filename', '')   → 原文件名（不含扩展名）
    元素按列表顺序拼接；分隔符只插在相邻两个"单元格/原文件名"之间，文字原样保留。
    所有内容都为空时回退到原文件名。返回 (计划列表, 错误列表)。
    """
    used: set[str] = set()
    plan: list[PlanItem] = []
    errors: list[tuple] = []
    for f in files:
        f = Path(f)
        try:
            parts: list[tuple] = []
            for kind, value in rule:
                if kind == 'text':
                    parts.append(('text', value))
                elif kind == 'filename':
                    parts.append(('data', f.stem))
                else:  # cell
                    parts.append(('data', read_cell_text(f, value)))
        except Exception as e:  # 文件损坏 / 被占用等
            errors.append((f, str(e)))
            continue
        stem = build_name(parts, separator)
        if not stem:
            stem = f.stem  # 全部为空时回退原文件名
        ext = f.suffix.lower()
        dst = build_dest_name(stem, ext, out_dir, used)
        used.add(dst.name.lower())
        plan.append(PlanItem(src=f, dst=dst, name=dst.name))
    return plan, errors


def execute_rename(plan: list[PlanItem], log: LogFn, progress: Optional[ProgressFn] = None,
                   cancel_event: Optional[Event] = None) -> dict:
    """执行重命名计划：把原文件复制到目标路径（shutil.copy2 保留 .xlsm 宏）。

    单文件失败不中断批次。返回 {'ok': n, 'skip': n, 'fail': n}。
    """
    if not plan:
        return {'ok': 0, 'skip': 0, 'fail': 0}
    plan[0].dst.parent.mkdir(parents=True, exist_ok=True)
    total = len(plan)
    ok = skip = fail = 0
    for i, item in enumerate(plan):
        if cancel_event is not None and cancel_event.is_set():
            log('[取消] 用户取消操作')
            break
        try:
            shutil.copy2(item.src, item.dst)
            ok += 1
            log(f'[OK] {item.src.name} → {item.dst.name}')
        except PermissionError:
            skip += 1
            log(f'[跳过] {item.src.name}：文件被占用，已跳过')
        except Exception as e:
            fail += 1
            log(f'[失败] {item.src.name}：{e}')
        if progress is not None:
            progress(i + 1, total)
    return {'ok': ok, 'skip': skip, 'fail': fail}


# ---------------------------------------------------------------------------
# 模板填充
# ---------------------------------------------------------------------------

def plan_fill(src_files: Iterable, template_path, mapping: list[tuple], out_dir,
              ) -> tuple[list[FillItem], list[tuple]]:
    """生成模板填充计划（不执行）。

    输出文件名 = 模板名 + '_' + 源文件名（保持模板扩展名），自动去重。
    返回 (计划列表, 错误列表)。
    """
    template_path = Path(template_path)
    tstem = template_path.stem
    text = template_path.suffix.lower()
    used: set[str] = set()
    plan: list[FillItem] = []
    errors: list[tuple] = []
    for f in src_files:
        f = Path(f)
        try:
            dst = build_dest_name(f'{tstem}_{f.stem}', text, out_dir, used)
        except Exception as e:
            errors.append((f, str(e)))
            continue
        used.add(dst.name.lower())
        plan.append(FillItem(src=f, dst=dst))
    return plan, errors


def _copy_row_style(ws, src_row: int, dst_row: int):
    """把 src_row 的样式（字体/边框/填充/数字格式/对齐）复制到 dst_row。

    用于让插入的空白行延续下方数据行的表格边框。
    """
    for col_idx in range(1, ws.max_column + 1):
        s = ws.cell(row=src_row, column=col_idx)
        d = ws.cell(row=dst_row, column=col_idx)
        d.font = copy(s.font)
        d.border = copy(s.border)
        d.fill = copy(s.fill)
        d.number_format = s.number_format
        d.alignment = copy(s.alignment)


def _fill_column(ws, values: list, col: str, start_row: int) -> int:
    """把 values 写入模板 ws 第 col 列，从 start_row 行开始。

    先在目标列内找可用空白（自 start_row 向下，遇到目标列已有内容或合并单元格为止）；
    若源数据超出空白，在边界处一次性批量插入差额空白行（把边界及下方内容、合并区域
    整体下移），绝不覆盖、也绝不因其他列有数据而影响对齐。
    返回插入的总行数。
    """
    # 1) 计算目标列可用空白与边界行（合并单元格视为已占用）
    boundary = None
    available = 0
    row = start_row
    while row <= ws.max_row and available < len(values):
        coord = f'{col}{row}'
        cell = ws[coord]
        if (cell.value is not None and str(cell.value).strip() != '') \
                or _openpyxl_merged_range(ws, coord) is not None:
            boundary = row
            break
        available += 1
        row += 1

    # 2) 源数据超出空白时，在边界处批量插入差额空白行
    shortfall = max(0, len(values) - available) if boundary is not None else 0
    if shortfall:
        ws.insert_rows(boundary, shortfall)
        _shift_openpyxl_merges(ws, boundary, shortfall)
        # 让插入的空行继承下方(边界)行的样式，表格边框连续
        for i in range(shortfall):
            _copy_row_style(ws, boundary + shortfall, boundary + i)

    # 3) 直接写入目标单元格（合并区已在扫描中作为边界排除，避免覆盖合并内容）
    for i, v in enumerate(values):
        r = start_row + i
        ws[f'{col}{r}'] = v
    return shortfall


# ---------------------------------------------------------------------------
# .xls（旧版 Excel）支持：读取 / 模板填充（含插入空白行）
# ---------------------------------------------------------------------------

class _XlsSource:
    """.xls 源文件读取器（xlrd）。"""

    def __init__(self, path):
        self.book = xlrd.open_workbook(str(path))
        self.sheet = self.book.sheet_by_index(0)

    def cell(self, coord):
        r, c = _coord_to_rc(coord)
        v = _xlrd_cell_value(self.book, self.sheet, r, c)
        return '' if v is None else v

    def column(self, col, start_row):
        c = _col_letter_to_idx(col)
        r = start_row - 1
        values = []
        while r < start_row - 1 + MAX_COL_ROWS:
            v = _xlrd_cell_value(self.book, self.sheet, r, c)
            if v is None or (isinstance(v, str) and v == ''):
                break
            values.append(v)
            r += 1
        return values

    def close(self):
        self.book.release_resources()


class _OpenpyxlSource:
    """.xlsx/.xlsm 源文件读取器（openpyxl）。"""

    def __init__(self, path):
        self.wb = load_workbook(path, data_only=True, read_only=True)
        self.ws = self.wb.worksheets[0]

    def cell(self, coord):
        v = self.ws[coord].value
        return '' if v is None else v

    def column(self, col, start_row):
        values = []
        r = start_row
        while r < start_row + MAX_COL_ROWS:
            v = self.ws[f'{col}{r}'].value
            if v is None or (isinstance(v, str) and v == ''):
                break
            values.append(v)
            r += 1
        return values

    def close(self):
        self.wb.close()


def _open_source(path):
    if Path(path).suffix.lower() == '.xls':
        return _XlsSource(path)
    return _OpenpyxlSource(path)


# Excel 内置数字格式串（xlrd 对部分内置格式返回 format_str=None，需自行补齐）
_BUILTIN_FORMATS = {
    1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
    5: '$#,##0_);($#,##0)', 6: '$#,##0_);[Red]($#,##0)',
    7: '$#,##0.00_);($#,##0.00)', 8: '$#,##0.00_);[Red]($#,##0.00)',
    9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
    14: 'm/d/yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
    18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
    22: 'm/d/yy h:mm', 49: '@',
}


def _resolve_num_format(book, format_key):
    """取数字格式串；格式串为 None 时回退到常用内置格式，避免 xlwt 对 None 报错。"""
    fmt = book.format_map.get(format_key)
    if fmt is not None and fmt.format_str:
        return fmt.format_str
    return _BUILTIN_FORMATS.get(format_key, '')


def _make_xlwt_style(book, xf_index):
    """把 xlrd 的 XF 格式转换为 xlwt 样式，用于写入时保留原格式。"""
    style = xlwt.XFStyle()
    if xf_index is None:
        return style
    xf = book.xf_list[xf_index]

    font_rec = book.font_list[xf.font_index]
    f = xlwt.Font()
    f.name = font_rec.name
    f.bold = bool(font_rec.bold)
    f.italic = bool(font_rec.italic)
    f.height = font_rec.height
    f.colour_index = font_rec.colour_index
    style.font = f

    b = xlwt.Borders()
    b.left = xf.border.left_line_style
    b.right = xf.border.right_line_style
    b.top = xf.border.top_line_style
    b.bottom = xf.border.bottom_line_style
    style.borders = b

    pat = xlwt.Pattern()
    pat.pattern = xf.background.fill_pattern
    pat.pattern_fore_colour = xf.background.pattern_colour_index
    pat.pattern_back_colour = xf.background.background_colour_index
    style.pattern = pat

    al = xlwt.Alignment()
    al.horz = xf.alignment.hor_align
    al.vert = xf.alignment.vert_align
    al.wrap = 1 if xf.alignment.text_wrapped else 0
    style.alignment = al

    style.num_format_str = _resolve_num_format(book, xf.format_key) or 'General'
    return style


def _shift_rows_down(cells: dict, styles: dict, merged: list, row: int):
    """把 cells/styles 中所有行号 >= row 的单元格下移一行（0 基），并同步合并单元格。"""
    keys = sorted([k for k in cells if k[0] >= row], key=lambda k: k[0], reverse=True)
    for k in keys:
        cells[(k[0] + 1, k[1])] = cells.pop(k)
        styles[(k[0] + 1, k[1])] = styles.pop(k)
    new_merged = []
    for m in merged:
        rlo, rhi, clo, chi = m
        if rlo >= row:
            new_merged.append((rlo + 1, rhi + 1, clo, chi))
        elif rhi > row:  # 合并范围跨过插入点 → 向下扩展
            new_merged.append((rlo, rhi + 1, clo, chi))
        else:
            new_merged.append((rlo, rhi, clo, chi))
    merged[:] = new_merged


def _fill_xls_template(template_path, cell_values: dict, column_jobs: list, output_path):
    """把单元格/整列数据填入 .xls 模板并另存为 .xls。

    .xls 为旧二进制格式，无插入行能力，因此采用"整表重建"：
    用 xlrd 读取原表全部内容与格式 → 计算插入行与写入 → 用 xlwt 重建保存。
    目标单元格已有内容时插入空白行下移，绝不覆盖。
    """
    book = xlrd.open_workbook(str(template_path), formatting_info=True)
    src = book.sheet_by_index(0)

    # 载入原表：cells[(r,c)]=值，styles[(r,c)]=xf_index（保留有值或有格式的单元格）
    cells: dict = {}
    styles: dict = {}
    for r in range(src.nrows):
        for c in range(src.ncols):
            v = src.cell_value(r, c)
            xf = src.cell_xf_index(r, c)
            if v != '' or xf != 0:
                cells[(r, c)] = v
                styles[(r, c)] = xf
    merged = [list(m) for m in src.merged_cells]

    # 单元格映射（目标在合并区内时重定向到锚点，保证值可见）
    for tpl_coord, value in cell_values.items():
        r, c = _coord_to_rc(tpl_coord)
        r, c = _xls_merged_anchor(merged, r, c)
        cells[(r, c)] = value
        styles.setdefault((r, c), None)

    # 整列映射（源数据超出空白时在边界处批量插入，不因其他列有数据而影响对齐）
    for tpl_col, tpl_start, values in column_jobs:
        c = _col_letter_to_idx(tpl_col)
        row = tpl_start - 1
        # 计算目标列可用空白与边界行（合并单元格视为已占用）
        boundary = None
        available = 0
        probe = row
        while available < len(values):
            occupied = False
            v = cells.get((probe, c))
            if v is not None and str(v).strip() != '':
                occupied = True
            else:
                for (rlo, rhi, clo, chi) in merged:
                    if rlo <= probe < rhi and clo <= c < chi:
                        occupied = True
                        break
            if occupied:
                boundary = probe
                break
            available += 1
            probe += 1
        # 在边界处批量下移差额行（_shift_rows_down 会同步移动合并区域）
        shortfall = max(0, len(values) - available) if boundary is not None else 0
        for _ in range(shortfall):
            _shift_rows_down(cells, styles, merged, boundary)
        # 直接写入目标单元格（不重定向到合并锚点，避免覆盖合并内容）
        for v in values:
            cells[(row, c)] = v
            if (row, c) not in styles:
                styles[(row, c)] = styles.get((row - 1, c))
            row += 1

    # 重建 .xls
    wb = xlwt.Workbook()
    ws = wb.add_sheet(src.name or 'Sheet1')
    # 合并区域内的单元格（含左上角）统一交给 write_merge 处理，
    # 逐格写入时跳过，以免 xlwt 抛 "Attempt to overwrite cell"
    covered = set()
    for (rlo, rhi, clo, chi) in merged:
        for r in range(rlo, rhi):
            for c in range(clo, chi):
                covered.add((r, c))
    for (r, c), v in cells.items():
        if (r, c) in covered:
            continue
        ws.write(r, c, v, _make_xlwt_style(book, styles.get((r, c))))
    for (rlo, rhi, clo, chi) in merged:
        ws.write_merge(rlo, rhi - 1, clo, chi - 1, cells.get((rlo, clo), ''),
                       _make_xlwt_style(book, styles.get((rlo, clo))))
    for col_idx, colinfo in src.colinfo_map.items():
        ws.col(col_idx).width = colinfo.width
    for r, rowinfo in src.rowinfo_map.items():
        if rowinfo.height:
            ws.row(r).height = rowinfo.height
    wb.save(str(output_path))
    book.release_resources()


def execute_fill(template_path, plan: list[FillItem], mapping: list[tuple],
                 log: LogFn, progress: Optional[ProgressFn] = None,
                 cancel_event: Optional[Event] = None) -> dict:
    """执行模板填充计划。

    映射条目支持两种（均指第一个工作表）：
      ('cell', 'A1', 'B2')      → 源单元格内容写入模板单元格（None 写空串以清空）
      ('column', 'C', 2, 'B', 6)→ 源文件 C 列自第 2 行向下（到空单元格为止）的数据，
                                 写入模板 B 列自第 6 行开始；目标已有内容时插入空白行下移，不覆盖

    每个源文件：重新加载模板副本（避免输出互相污染），写入后保存到目标路径。
    - .xlsx/.xlsm 模板：openpyxl 往返保留格式；整列映射用插入空白行扩展。
    - .xls 模板：整表重建（xlrd 读 + xlwt 写），同样支持插入空白行。
    返回 {'ok': n, 'skip': n, 'fail': n}。
    """
    if not plan:
        return {'ok': 0, 'skip': 0, 'fail': 0}
    template_path = Path(template_path)
    is_xls_tpl = template_path.suffix.lower() == '.xls'
    keep_vba = template_path.suffix.lower() == '.xlsm'
    plan[0].dst.parent.mkdir(parents=True, exist_ok=True)
    total = len(plan)
    ok = skip = fail = 0
    for i, item in enumerate(plan):
        if cancel_event is not None and cancel_event.is_set():
            log('[取消] 用户取消操作')
            break
        wb = None
        src = None
        try:
            # 1) 读取源文件数据（每个源文件只打开一次，.xls/.xlsx/.xlsm 均可）
            cell_values: dict = {}   # 模板坐标 → 值
            column_jobs: list = []   # (模板列, 起始行, 数据列表)
            src = _open_source(item.src)
            for m in mapping:
                if m[0] == 'cell':
                    cell_values[m[2]] = src.cell(m[1])
                else:  # column
                    column_jobs.append((m[3], m[4], src.column(m[1], m[2])))
            src.close()
            src = None

            # 2) 写入模板副本（先写单元格映射，再写整列映射，可能插入行）
            if is_xls_tpl:
                _fill_xls_template(template_path, cell_values, column_jobs, item.dst)
            else:
                wb = load_workbook(template_path, keep_vba=keep_vba)
                ws = wb.worksheets[0]
                for tpl_coord, v in cell_values.items():
                    # 目标在合并区内时重定向到锚点，保证值在合并单元格可见
                    ws[_openpyxl_merged_anchor(ws, tpl_coord)] = v
                for tpl_col, tpl_start, values in column_jobs:
                    _fill_column(ws, values, tpl_col, tpl_start)
                wb.save(item.dst)
                wb.close()
                wb = None
            ok += 1
            log(f'[OK] {item.src.name} → {item.dst.name}')
        except PermissionError:
            skip += 1
            log(f'[跳过] {item.src.name}：文件被占用，已跳过')
        except Exception as e:
            fail += 1
            log(f'[失败] {item.src.name}：{e}')
        finally:
            for _w in (wb, src):
                if _w is not None:
                    try:
                        _w.close()
                    except Exception:
                        pass
        if progress is not None:
            progress(i + 1, total)
    return {'ok': ok, 'skip': skip, 'fail': fail}
