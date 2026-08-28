# -*- coding: utf-8 -*-
"""渲染验收：目录下全部 pptx -> LibreOffice PDF -> PyMuPDF 检查。

用法:
    python qa_render.py <输出目录>
检查: 转换是否成功 / 页数 / 空白页（文本<5字符）/ 汇总页数。
前置: LibreOffice (soffice 在 PATH 或 Program Files)、PyMuPDF (pip install pymupdf)。
"""
import sys, io, os, glob, subprocess, tempfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

SOFFICE_CANDIDATES = [
    'soffice',
    r'C:\Program Files\LibreOffice\program\soffice.exe',
    r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
]

def soffice():
    for c in SOFFICE_CANDIDATES:
        p = subprocess.run(['where', c] if c == 'soffice' else ['test', '-f', c],
                           capture_output=True)
        if c != 'soffice' and os.path.exists(c):
            return c
        if c == 'soffice' and p.returncode == 0:
            return c
    raise SystemExit('未找到 LibreOffice (soffice)，请安装或加入 PATH')

def main():
    base = sys.argv[1]
    files = sorted(glob.glob(os.path.join(base, '*', '*.pptx')) + glob.glob(os.path.join(base, '*.pptx')))
    if not files:
        raise SystemExit('目录下没有 pptx: ' + base)
    outdir = tempfile.mkdtemp(prefix='qa_render_')
    cmd = [soffice(), '--headless', '--convert-to', 'pdf', '--outdir', outdir] + files
    r = subprocess.run(cmd, capture_output=True, timeout=600)
    pdfs = sorted(glob.glob(os.path.join(outdir, '*.pdf')))
    import fitz
    print('pptx %d 个, pdf 转出 %d 个' % (len(files), len(pdfs)))
    if len(pdfs) != len(files):
        converted = {os.path.splitext(os.path.basename(p))[0] for p in pdfs}
        miss = [os.path.basename(f) for f in files
                if os.path.splitext(os.path.basename(f))[0] not in converted]
        print('!! 转换失败（多半文件损坏）:', miss)
    total = 0
    problems = []
    for p in pdfs:
        d = fitz.open(p)
        total += len(d)
        def _blank(pg):
            if len(pg.get_text().strip()) >= 5:
                return False
            # 纯图片页（如官方日程海报页）不算空白：有覆盖过半页面的图片即视为有内容
            for img in pg.get_image_info():
                bb = img.get('bbox')
                if bb and (bb[2]-bb[0])*(bb[3]-bb[1]) >= 0.5*pg.rect.width*pg.rect.height:
                    return False
            return True
        blanks = [i + 1 for i in range(len(d)) if _blank(d[i])]
        if blanks:
            problems.append('%s 空白页 %s' % (os.path.basename(p), blanks))
        print('  %-30s %d页%s' % (os.path.basename(p)[:30], len(d), ' 空白页!' + str(blanks) if blanks else ''))
    print('合计 %d 页' % total)
    print('问题:', problems if problems else '无')
    sys.exit(1 if problems or len(pdfs) != len(files) else 0)

if __name__ == '__main__':
    main()
