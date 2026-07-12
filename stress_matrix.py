# stress_matrix.py — systematic hardening harness for the PDF engine.
# Generates synthetic PDFs across fonts x structures, runs each through the
# REAL engine (pdf_collect + build) with translation-length stressors, and
# reports placed/shrunk/failed per cell so weak points surface before users.
# LOCAL only (imports app.py directly — no HTTP, no keys).
import io
import sys

import fitz

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass
sys.path.insert(0, '.')
from app import pdf_collect, MIN_FONT, _font_class, BULLET_RE, BULLET_ONLY_RE  # noqa

# ── translation stressors: how the replacement text differs from the source ──
# Real translations swing in length; non-latin scripts stress font+shaping.
EXPANDERS = {
    'identity': lambda t: t,
    'de_plus30': lambda t: t + ' ' + t[: max(1, len(t) // 3)],   # German-ish +30%
    'verbose_x2': lambda t: t + ' — ' + t,                        # doubles length
    'ukrainian': lambda t: 'Переклад: ' + t,                      # Cyrillic prefix
    'arabic': lambda t: 'ترجمة: ' + t,                            # RTL prefix
    'thai': lambda t: 'การแปล: ' + t,                             # Thai (no spaces)
    'cjk': lambda t: '翻译内容 ' + t,                              # CJK prefix
    'longword': lambda t: t + ' Kraftfahrzeughaftpflichtversicherung',  # unbreakable
}

FONTS = [('helv', 'Helvetica'), ('times', 'Times-Roman'), ('cour', 'Courier')]


def _page(doc, draw):
    p = doc.new_page(width=595, height=842)
    draw(p)
    return p


def make_doc(structure, fontname):
    """Build a one-page PDF exercising a structure with a given base font."""
    doc = fitz.open()

    def heading_para(p):
        p.insert_text((60, 80), 'Section Heading', fontsize=20, fontname=fontname)
        body = ('This is a normal paragraph of running text that should wrap '
                'across several lines inside a reasonably wide column on the page.')
        p.insert_textbox(fitz.Rect(60, 110, 500, 240), body, fontsize=11, fontname=fontname)

    def bullets(p):
        p.insert_text((60, 80), 'Bullet List', fontsize=16, fontname=fontname)
        y = 120
        for i in range(6):
            # realistic: bullet glyph + text in one box (one block after collect)
            p.insert_textbox(fitz.Rect(60, y - 10, 500, y + 6),
                             '%s List item number %d with some descriptive text after it'
                             % (chr(0x2022), i + 1), fontsize=11, fontname=fontname)
            y += 26

    def numbered(p):
        p.insert_text((60, 80), 'Numbered Steps', fontsize=16, fontname=fontname)
        y = 120
        for i in range(5):
            p.insert_textbox(fitz.Rect(60, y, 500, y + 16),
                             '%d. Step description that runs to a moderate length here' % (i + 1),
                             fontsize=11, fontname=fontname)
            y += 24

    def two_column(p):
        for cx in (60, 320):
            p.insert_textbox(fitz.Rect(cx, 100, cx + 210, 400),
                             'Column text block that needs to stay inside its own '
                             'narrow column and not bleed into the neighbour column.',
                             fontsize=10, fontname=fontname)

    def tight_boxes(p):
        # small labels in tight cells (table-ish) — the classic overflow trap
        y = 100
        for label in ('Name', 'Position', 'Start date', 'Weekly hours'):
            p.insert_textbox(fitz.Rect(60, y, 180, y + 16), label, fontsize=10, fontname=fontname)
            p.insert_textbox(fitz.Rect(190, y, 500, y + 16), 'Value goes here',
                             fontsize=10, fontname=fontname)
            y += 22

    def mixed_sizes(p):
        p.insert_text((60, 80), 'BIG TITLE', fontsize=24, fontname=fontname)
        p.insert_text((60, 130), 'Medium subtitle', fontsize=15, fontname=fontname)
        p.insert_textbox(fitz.Rect(60, 160, 500, 300),
                         'Small body copy underneath the two larger headings above it.',
                         fontsize=9, fontname=fontname)

    structures = {
        'heading_para': heading_para, 'bullets': bullets, 'numbered': numbered,
        'two_column': two_column, 'tight_boxes': tight_boxes, 'mixed_sizes': mixed_sizes,
    }
    _page(doc, structures[structure])
    return doc


def build_with(doc, tr):
    """Mirror app.pdf_build's insert loop (page-scope off) and return report."""
    jobs = pdf_collect(doc)
    report = {'blocks': len(jobs), 'placed': 0, 'shrunk': 0, 'failed': 0, 'fails': []}
    missing = [j['id'] for j in jobs if not isinstance(tr.get(j['id']), str) or not tr[j['id']].strip()]
    if missing:
        return {'error': 'missing', 'ids': missing}
    for page in doc:
        pj = [j for j in jobs if j['page'] == page.number]
        for j in pj:
            page.add_redact_annot(j['rect'])
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
        for j in pj:
            text = tr[j['id']]
            esc = (text.replace('&', '&amp;').replace('<', '&lt;')
                       .replace('>', '&gt;').replace('\n', '<br>'))
            color = '#%06x' % j['color']
            size = max(MIN_FONT, j['size'])
            fam = {'serif': 'Times, serif', 'mono': 'Courier, monospace'}.get(
                j.get('font', 'sans'), 'Helvetica, Arial, sans-serif')
            base = 'font-size:%.1fpx;color:%s;font-family:%s' % (size, color, fam)
            html = ('<b style="%s">%s</b>' % (base, esc)) if j['bold'] else ('<span style="%s">%s</span>' % (base, esc))
            rect = fitz.Rect(j['rect'].x0, j['rect'].y0 - 0.15 * j['size'],
                             j['rect'].x1 + 0.35 * j['rect'].width,
                             j['rect'].y1 + 0.45 * j['size'])
            rect.x1 = min(rect.x1, page.rect.width - 8)
            low = min(1.0, MIN_FONT / size)
            spare, scale = page.insert_htmlbox(rect, html, scale_low=low)
            if spare < 0:
                report['failed'] += 1
                report['fails'].append((j['size'], round(rect.width), round(rect.height), text[:40]))
            elif scale < 1:
                report['shrunk'] += 1
            else:
                report['placed'] += 1
    return report


def run():
    structures = ['heading_para', 'bullets', 'numbered', 'two_column', 'tight_boxes', 'mixed_sizes']
    grand = {'placed': 0, 'shrunk': 0, 'failed': 0}
    worst = []
    print('%-14s %-10s %-11s  P/S/F' % ('structure', 'font', 'expander'))
    print('-' * 60)
    for structure in structures:
        for fk, fname in FONTS:
            for ek, exp in EXPANDERS.items():
                doc = make_doc(structure, fname)
                jobs = pdf_collect(doc)
                tr = {j['id']: exp(j['text']) for j in jobs}
                doc2 = make_doc(structure, fname)  # fresh (build mutates)
                rep = build_with(doc2, tr)
                if 'error' in rep:
                    print('%-14s %-10s %-11s  ERR %s' % (structure, fk, ek, rep['error']))
                    continue
                for k in ('placed', 'shrunk', 'failed'):
                    grand[k] += rep[k]
                flag = '  <-- FAIL' if rep['failed'] else ('  (shrunk)' if rep['shrunk'] else '')
                print('%-14s %-10s %-11s  %d/%d/%d%s' % (
                    structure, fk, ek, rep['placed'], rep['shrunk'], rep['failed'], flag))
                if rep['failed']:
                    worst.append((structure, fk, ek, rep['fails']))
    print('-' * 60)
    tot = sum(grand.values())
    print('TOTAL blocks=%d  placed=%d (%.0f%%)  shrunk=%d  failed=%d' % (
        tot, grand['placed'], 100 * grand['placed'] / max(1, tot), grand['shrunk'], grand['failed']))
    if worst:
        print('\nWORST CELLS (structure/font/expander -> sample fails):')
        for st, fk, ek, fails in worst[:12]:
            print('  %s/%s/%s:' % (st, fk, ek))
            for sz, w, h, txt in fails[:3]:
                print('    size=%.1f box=%dx%d :: %r' % (sz, w, h, txt))


if __name__ == '__main__':
    run()
