"""月次ヒアリング表の整形ツール。

使い方: python3 -I hearing_fix.py <元データのフォルダ> <出力フォルダ> <設定JSON>

設定JSONの例（氏名を含むためリポジトリには置かない）:
{
 "ym": "202610", "month_end": "2026-10-31",
 "companies": {
  "<元ファイル名の先頭>": {
    "company": "株式会社ヤマサORM", "name_col": "B", "yomi_col": "C", "date_col": "F",
    "answer_cols": "GHIJK", "vn_cols": "GH", "non_vn": ["氏名の一部", ...], "last_col": "M",
    "question_cols": "GHIJK",  # 省略時は answer_cols
    "red": ["I20"], "yomi": {"氏名の一部": "呼名"}, "status": {"氏名の一部": "帰国中（産休・育休中）"},
    "fixtext": {"H3": ["誤", "正"]}, "dedupe": ["G7"]
  }
 }
}
ルールは hearing/README.md を参照。
"""
import copy, datetime, glob, json, os, re, sys
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl.worksheet.pagebreak import RowBreak, ColBreak
from openpyxl.worksheet.properties import PageSetupProperties

IN, OUT, CONF = sys.argv[1], sys.argv[2], sys.argv[3]
os.makedirs(OUT, exist_ok=True)
conf = json.load(open(CONF, encoding='utf-8'))
YM = conf['ym']
MONTH_END = datetime.datetime.strptime(conf['month_end'], '%Y-%m-%d')
NOREPLY_IN = ('返答なし', '返信なし', 'リプライなし', 'リプライありません')
NOREPLY = conf.get('noreply_text', '（期日までの相談報告事項なし）')

log = []


def note(company, msg):
    log.append(f'{company}\t{msg}')


def clean(v):
    if not isinstance(v, str):
        return v
    s = v.strip(' 　\n\r\t')
    s = re.sub(r'^[.．:：]+', '', s).strip(' 　\n')
    return s


def recolor(cell, rgb='FFFF0000', bold=None):
    f = cell.font
    cell.font = Font(name=f.name, sz=f.sz, b=f.b if bold is None else bold, i=f.i,
                     u=f.u, strike=f.strike, color=rgb, vertAlign=f.vertAlign)


def diag(cell):
    b = cell.border
    cell.border = Border(left=b.left, right=b.right, top=b.top, bottom=b.bottom,
                         diagonal=Side(style='thin', color='FF000000'), diagonalDown=True, diagonalUp=False)


def col_widths(ws, last):
    from openpyxl.utils import get_column_letter, column_index_from_string
    w = {}
    for dim in ws.column_dimensions.values():
        if dim.width and dim.min and dim.max:
            for i in range(dim.min, dim.max + 1):
                w[get_column_letter(i)] = dim.width
    default = ws.sheet_format.defaultColWidth or 8.43
    return {get_column_letter(i): w.get(get_column_letter(i), default)
            for i in range(1, column_index_from_string(last) + 1)}


def text_lines(text, width, sz):
    import math, unicodedata
    cap = max(width * 11.0 / sz, 1)  # width units per line at this font size
    n = 0
    for line in str(text).split('\n'):
        units = sum(2 if unicodedata.east_asian_width(ch) in 'WFA' else 1 for ch in line)
        n += max(1, math.ceil(units * 1.2 / cap))
    return n


def autofit_rows(ws, r1, r2, last):
    widths = col_widths(ws, last)
    for r in range(r1, r2 + 1):
        need = 0
        for c, w in widths.items():
            cell = ws[f'{c}{r}']
            if cell.value is None or isinstance(cell.value, (int, float, datetime.datetime)):
                continue
            sz = cell.font.sz or 12
            need = max(need, text_lines(cell.value, w, sz) * sz * 1.4 + 8)
        cur = ws.row_dimensions[r].height or 15
        if need > cur:
            ws.row_dimensions[r].height = round(need, 1)


def set_noreply(ws, r, d, qcols, skip):
    # 未返信: リプライ日は空欄にし、各質問欄に「（期日までの相談報告事項なし）」と記載
    d.value = None
    for c in qcols:
        if c in skip:
            continue
        cell = ws[f'{c}{r}']
        cell.value = NOREPLY
        recolor(cell, rgb='FF000000', bold=False)


for prefix, extra in conf['companies'].items():
    company, NC, YC, DC = extra['company'], extra['name_col'], extra['yomi_col'], extra['date_col']
    ANS, VN, NONVN, LAST = extra['answer_cols'], extra.get('vn_cols', ''), extra.get('non_vn', []), extra['last_col']
    QCOLS = extra.get('question_cols', ANS)  # 未返信コメントを入れる質問欄
    hits = glob.glob(os.path.join(IN, prefix + '*'))
    if not hits:
        print('MISSING', company)
        continue
    wb = openpyxl.load_workbook(hits[0])
    ws = wb.worksheets[0]
    ws.title = f'{YM}{company}様 定期ヒアリング'
    for ref, (a, b) in extra.get('fixtext', {}).items():
        if isinstance(ws[ref].value, str) and a in ws[ref].value:
            ws[ref] = ws[ref].value.replace(a, b)
            note(company, f'{ref} 見出しの誤記「{a}」→「{b}」')
    biko = next(r for r in range(4, ws.max_row + 1) if str(ws[f'A{r}'].value or '').strip() == '備考')

    for r in range(4, biko):
        name = str(ws[f'{NC}{r}'].value or '')
        if not name.strip():
            continue
        a = ws[f'A{r}'].value
        if isinstance(a, float) and a.is_integer():
            ws[f'A{r}'] = int(a)
        if isinstance(ws[f'D{r}'].value, str) and ws[f'D{r}'].number_format != 'General':
            ws[f'D{r}'].number_format = 'General'
        for c in ANS:
            cell = ws[f'{c}{r}']
            nv = clean(cell.value)
            if nv != cell.value:
                cell.value = nv if nv != '' else None
        for key, y in extra.get('yomi', {}).items():
            if key in name and not ws[f'{YC}{r}'].value:
                ws[f'{YC}{r}'] = y
                note(company, f'{name.strip()}：呼名「{y}」を補完')
        d = ws[f'{DC}{r}']
        has_ans = any(ws[f'{c}{r}'].value not in (None, '') for c in ANS)
        status = next((v for k, v in extra.get('status', {}).items() if k in name), None)
        skip = VN if (VN and any(k in name for k in NONVN)) else ''
        if isinstance(d.value, str) and d.value.strip() in NOREPLY_IN:
            set_noreply(ws, r, d, QCOLS, skip)
            note(company, f'{name.strip()}：未返信 →「{NOREPLY}」')
        elif d.value in (None, ''):
            if has_ans:
                d.value = MONTH_END
                d.number_format = 'm/d'
                note(company, f'{name.strip()}：リプライ日の記載漏れ → {MONTH_END:%-m/%-d}')
            elif status:
                d.value = status
                note(company, f'{name.strip()}：回答なし →「{status}」（前月と同じ扱い、要確認）')
            else:
                set_noreply(ws, r, d, QCOLS, skip)
                note(company, f'{name.strip()}：未返信 →「{NOREPLY}」')
        if VN and any(k in name for k in NONVN):
            for c in VN:
                if ws[f'{c}{r}'].value in (None, ''):
                    diag(ws[f'{c}{r}'])
    for ref in extra.get('dedupe', []):
        v = ws[ref].value
        if isinstance(v, str):
            parts = [p for p in re.split(r'(?<=。)', v) if p]
            out = []
            for p in parts:
                if p not in out:
                    out.append(p)
            if ''.join(out) != v:
                ws[ref] = ''.join(out)
                note(company, f'{ref} 同じ文が2回入っていたので1回に整理')
    for ref in extra.get('red', []):
        recolor(ws[ref])
    autofit_rows(ws, 3, biko - 1, LAST)

    ws.row_breaks = RowBreak()
    ws.col_breaks = ColBreak()
    ws.print_area = f'A1:{LAST}{biko}'
    ws.print_title_rows = '1:3'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'landscape'
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    wb.save(os.path.join(OUT, f'{YM}{company}様 定期ヒアリング.xlsx'))

open(os.path.join(OUT, '..', 'changes.tsv'), 'w').write('\n'.join(log) + '\n')
print('\n'.join(log))
