"""LibreOffice(UNO)でExcel(.xls/.xlsx)のセルを書き換え、Excel形式とPDFで保存する。
図形（性別欄の○など）の移動にも対応。openpyxlでは図形が消えるため、図形入りの様式はこちらを使う。

使い方:
    python uno_cells.py in.xls out.xls out.pdf edits.json
edits.json:
  {"values": [[行, 列, 数値], ...],          # 1始まりの行・列
   "strings": [[行, 列, "文字列"], ...],
   "move_shapes": [[図形番号, dx(1/100mm), dy], ...]}
"""
import json, subprocess, sys, time
import uno
from com.sun.star.beans import PropertyValue

def P(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p

def main(src, dst_xls, dst_pdf, edits_path):
    edits = json.load(open(edits_path, encoding='utf-8'))
    proc = subprocess.Popen(['soffice', '--headless', '--norestore',
                             '--accept=socket,host=localhost,port=2002;urp;'])
    try:
        ctx = None
        for _ in range(60):
            try:
                local = uno.getComponentContext()
                res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
                ctx = res.resolve('uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext'); break
            except Exception:
                time.sleep(1)
        desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
        doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(src), '_blank', 0, (P('Hidden', True),))
        sh = doc.Sheets.getByIndex(0)
        cell = lambda r, c: sh.getCellByPosition(c - 1, r - 1)
        for r, c, v in edits.get('values', []):
            if cell(r, c).getFormula().startswith('='):
                print('数式のため変更せず', r, c); continue
            cell(r, c).setValue(v)
        for r, c, v in edits.get('strings', []):
            cell(r, c).setString(v)
        dp = sh.DrawPage
        for i, dx, dy in edits.get('move_shapes', []):
            s = dp.getByIndex(i); p = s.Position; p.X += dx; p.Y += dy; s.Position = p
        doc.calculateAll()
        filt = 'MS Excel 97' if dst_xls.endswith('.xls') else 'Calc MS Excel 2007 XML'
        doc.storeToURL(uno.systemPathToFileUrl(dst_xls), (P('FilterName', filt),))
        doc.storeToURL(uno.systemPathToFileUrl(dst_pdf), (P('FilterName', 'calc_pdf_Export'),))
        doc.close(True)
    finally:
        proc.terminate()

if __name__ == '__main__':
    main(*sys.argv[1:5])
