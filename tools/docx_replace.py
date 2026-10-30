"""Word(.docx)の段落内テキストを、ラン（書式の区切り）をまたいで置換する。

使い方:
    python docx_replace.py in.docx out.docx replacements.json
replacements.json: [["置換前", "置換後"], ...]
"""
import json, sys
import docx
from docx.oxml.ns import qn

XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

def replace_in_paragraph(p, old, new):
    ts = list(p.iter(qn('w:t')))
    count = 0
    while True:
        full = ''.join(t.text or '' for t in ts)
        i = full.find(old)
        if i < 0:
            return count
        pos = 0; start = end = None
        for k, t in enumerate(ts):
            L = len(t.text or '')
            if start is None and i < pos + L:
                start = (k, i - pos)
            if start is not None and i + len(old) <= pos + L:
                end = (k, i + len(old) - pos); break
            pos += L
        (ks, os_), (ke, oe) = start, end
        if ks == ke:
            ts[ks].text = ts[ks].text[:os_] + new + ts[ks].text[oe:]
        else:
            ts[ks].text = ts[ks].text[:os_] + new
            for k in range(ks + 1, ke):
                ts[k].text = ''
            ts[ke].text = ts[ke].text[oe:]
        for t in ts:
            t.set(XML_SPACE, 'preserve')
        count += 1
        if old in new:
            return count

def main(src, dst, rep_path):
    d = docx.Document(src)
    reps = json.load(open(rep_path, encoding='utf-8'))
    paras = list(d.element.body.iter(qn('w:p')))
    for old, new in reps:
        n = sum(replace_in_paragraph(p, old, new) for p in paras)
        print(f'{n:3d}  {old!r} -> {new!r}')
    d.save(dst)

if __name__ == '__main__':
    main(*sys.argv[1:4])
