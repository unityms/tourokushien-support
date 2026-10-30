"""参考PDFの1ページの文字を差し替え、旧テキストが残らないよう画像化して書き出す。

使い方（Pythonから）:
    from pdf_overlay import overlay_page
    overlay_page(src_doc, page_no, out_doc,
                 items=[((x0,y0,x1,y1), '新しい文字', x, baseline_y, size, fontfile)])
- items の矩形を白塗りし、指定位置に文字を書いてから300dpiで画像化して out_doc に追加する。
- PyMuPDF の apply_redactions は一部のPDF（PDFium/Quartz出力）でレイアウトが崩れるため使わない。
"""
import pymupdf

def overlay_page(src, pno, out, items, extra=None, dpi=300):
    out.insert_pdf(src, from_page=pno, to_page=pno)
    page = out[-1]
    for r, *_ in items:
        page.draw_rect(pymupdf.Rect(r), color=None, fill=(1, 1, 1), overlay=True)
    for i, (r, text, x, y, size, ff) in enumerate(items):
        page.insert_text((x, y), text, fontfile=ff, fontname=f'f{i}', fontsize=size)
    if extra:
        extra(page)
    rect = page.rect
    pix = page.get_pixmap(dpi=dpi)
    idx = len(out) - 1
    out.delete_page(idx)
    np_ = out.new_page(width=rect.width, height=rect.height)
    np_.insert_image(rect, stream=pix.tobytes('png'))
    return np_
