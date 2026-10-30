"""在留カード表裏の画像を、印刷用PDF（A4）にする。
1ページ目: 原寸（横85.6mm）、2ページ目: 150%拡大。

使い方:
    python card_copy.py front.jpg back.jpg out.pdf [front_crop] [back_crop]
crop は "x0,y0,x1,y1"（画像のピクセル座標、カードの外の余白を切る場合）
"""
import io, sys
import pymupdf
from PIL import Image

MM = 72 / 25.4

def load(path, crop=None):
    im = Image.open(path).convert('RGB')
    if crop:
        im = im.crop(tuple(int(v) for v in crop.split(',')))
    b = io.BytesIO(); im.save(b, 'JPEG', quality=95)
    return b.getvalue(), im.size

def main(front, back, out, fcrop=None, bcrop=None):
    a4 = pymupdf.paper_rect('a4')
    doc = pymupdf.open()
    imgs = [load(front, fcrop), load(back, bcrop)]
    for scale, gap in [(1.0, 40), (1.5, 24)]:
        pg = doc.new_page(width=a4.width, height=a4.height)
        y = 90
        for data, (w, h) in imgs:
            W = 85.6 * MM * scale; H = W * h / w; x0 = (a4.width - W) / 2
            pg.insert_image(pymupdf.Rect(x0, y, x0 + W, y + H), stream=data)
            y += H + gap
    doc.save(out)

if __name__ == '__main__':
    main(*sys.argv[1:6])
