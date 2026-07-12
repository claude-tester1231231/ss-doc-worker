# ss-doc-worker — document extract/rebuild engine (PyMuPDF + python-docx).
#
# This service deliberately contains NO translation logic, NO prompts, NO API
# keys and NO business rules. It does exactly two things per format:
#   extract: document in  -> ordered text blocks (id + text) out
#   build:   document + {id: replacement-text} in -> rebuilt document out
# The caller owns everything else. That separation is the point: the whole
# PyMuPDF-touching surface lives in this small AGPL-licensed service.
#
# Auth: shared bearer token from WORKER_TOKEN env (config, not code).
import io
import json
import os
import re

import fitz  # PyMuPDF (AGPL)
import docx  # python-docx (MIT)
from docx.oxml.ns import qn
from flask import Flask, abort, jsonify, request, send_file

app = Flask(__name__)
MAX_BYTES = 25 * 1024 * 1024
MIN_FONT = 8.0

MARK = {(True, False): 'B', (False, True): 'I', (True, True): 'BI'}
M_RE = re.compile(r'\[\[(/?)(BI|B|I)\]\]')


def _auth():
    want = os.environ.get('WORKER_TOKEN', '')
    got = (request.headers.get('Authorization') or '').replace('Bearer ', '', 1)
    if not want or got != want:
        abort(401)


def _file_bytes(field='file'):
    f = request.files.get(field)
    if f is None:
        abort(400, 'missing file')
    data = f.read()
    if not data or len(data) > MAX_BYTES:
        abort(413 if data else 400)
    return data


# ── PDF lane ────────────────────────────────────────────────────────────────

def pdf_collect(doc):
    """The document's OWN text blocks -> paragraph-merged jobs (no heuristics
    beyond joining adjacent same-style blocks). Union rect, bold flag, color."""
    jobs = []
    for pno, page in enumerate(doc):
        raw = []
        for b in page.get_text('dict')['blocks']:
            if b.get('type') != 0:
                continue
            spans = [s for l in b.get('lines', []) for s in l.get('spans', [])]
            text = ' '.join(s['text'] for s in spans).strip()
            if not text:
                continue
            raw.append({
                'page': pno, 'rect': fitz.Rect(b['bbox']), 'text': text,
                'size': max(s['size'] for s in spans),
                'color': max(set(s['color'] for s in spans),
                             key=[s['color'] for s in spans].count),
                'bold': any('bold' in s.get('font', '').lower() for s in spans),
            })
        raw.sort(key=lambda j: (j['rect'].y0, j['rect'].x0))
        merged = []
        for j in raw:
            m = merged[-1] if merged else None
            if m and m['page'] == j['page']:
                vgap = j['rect'].y0 - m['rect'].y1
                xover = min(m['rect'].x1, j['rect'].x1) - max(m['rect'].x0, j['rect'].x0)
                same_style = m['bold'] == j['bold'] and abs(m['size'] - j['size']) < 0.6
                if vgap < 0.9 * j['size'] and xover > 0 and same_style:
                    m['rect'] |= j['rect']
                    m['text'] += ' ' + j['text']
                    continue
            merged.append(dict(j))
        jobs.extend(merged)
    for i, j in enumerate(jobs):
        j['id'] = 'b%d' % i
    return jobs


@app.post('/v1/pdf/extract')
def pdf_extract():
    _auth()
    doc = fitz.open(stream=_file_bytes(), filetype='pdf')
    jobs = pdf_collect(doc)
    # 'page' (1-based) rides along so callers can pipeline per page (lazy view)
    return jsonify({
        'engine': 'pdf', 'pages': len(doc),
        'blocks': [{'id': j['id'], 'text': j['text'], 'page': j['page'] + 1} for j in jobs],
    })


@app.post('/v1/pdf/build')
def pdf_build():
    _auth()
    tr = json.loads(request.form.get('translations') or '{}')
    if not isinstance(tr, dict) or not tr:
        abort(400, 'missing translations')
    doc = fitz.open(stream=_file_bytes(), filetype='pdf')
    jobs = pdf_collect(doc)
    # optional page scope (1-based): build ONLY that page and return a 1-page
    # PDF — the lazy per-page pipeline. Ids stay the global b<N> ids.
    only = request.form.get('page')
    only_idx = None
    if only is not None and str(only).strip() != '':
        try:
            only_idx = int(only) - 1
        except ValueError:
            abort(400, 'bad page')
        if only_idx < 0 or only_idx >= len(doc):
            abort(400, 'bad page')
        jobs = [j for j in jobs if j['page'] == only_idx]
    # hard 1:1 contract: every block IN SCOPE must have a non-empty replacement
    missing = [j['id'] for j in jobs if not isinstance(tr.get(j['id']), str) or not tr[j['id']].strip()]
    if missing:
        abort(422, 'missing ids: ' + ','.join(missing[:20]))
    report = {'blocks': len(jobs), 'placed': 0, 'shrunk': 0, 'failed': 0}
    for page in doc:
        if only_idx is not None and page.number != only_idx:
            continue
        pj = [j for j in jobs if j['page'] == page.number]
        if not pj:
            continue
        for j in pj:
            page.add_redact_annot(j['rect'])
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
        for j in pj:
            text = tr[j['id']]
            # insert_htmlbox: full Unicode (Cyrillic/Arabic/CJK, shaping, RTL) —
            # Base14 'helv' turned every non-Latin char into '?' (live 12/7).
            # The replacement text is DATA: escape it, never interpret as HTML.
            esc = (text.replace('&', '&amp;').replace('<', '&lt;')
                       .replace('>', '&gt;').replace('\n', '<br>'))
            color = '#%06x' % j['color']
            size = max(MIN_FONT, j['size'])
            if j['bold']:
                html = '<b style="font-size:%.1fpx;color:%s">%s</b>' % (size, color, esc)
            else:
                html = '<span style="font-size:%.1fpx;color:%s">%s</span>' % (size, color, esc)
            rect = fitz.Rect(j['rect'].x0, j['rect'].y0 - 0.15 * j['size'],
                             j['rect'].x1 + 0.35 * j['rect'].width,
                             j['rect'].y1 + 0.45 * j['size'])
            rect.x1 = min(rect.x1, page.rect.width - 8)
            # scale_low: auto-shrink down to the 8pt-equivalent floor, never below
            low = min(1.0, MIN_FONT / size)
            spare, scale = page.insert_htmlbox(rect, html, scale_low=low)
            if spare < 0:
                report['failed'] += 1  # fail-don't-clip: htmlbox keeps content visible
            elif scale < 1:
                report['shrunk'] += 1
            else:
                report['placed'] += 1
    if only_idx is not None:
        doc.select([only_idx])
    out = io.BytesIO(doc.tobytes())
    resp = send_file(out, mimetype='application/pdf', as_attachment=True,
                     download_name='translated.pdf')
    resp.headers['X-Build-Report'] = json.dumps(report)
    return resp


# ── DOCX lane ───────────────────────────────────────────────────────────────

def _eff(run, attr):
    v = getattr(run.font, attr)
    return bool(v) if v is not None else False


def docx_pars(d):
    def walk_tables(tables):
        for t in tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p
                    yield from walk_tables(cell.tables)
    for p in d.paragraphs:
        yield p
    yield from walk_tables(d.tables)
    for s in d.sections:
        for part in (s.header, s.footer):
            for p in part.paragraphs:
                yield p
            yield from walk_tables(part.tables)


def docx_segments(par):
    segs = []
    for r in par.runs:
        if r._element.findall(qn('w:drawing')):
            continue
        if not r.text:
            continue
        key = (_eff(r, 'bold') or par.style.name.startswith('Heading'), _eff(r, 'italic'))
        if segs and segs[-1][1] == key:
            segs[-1][0] += r.text
        else:
            segs.append([r.text, key, r])
    return segs


def docx_marked(par):
    segs = docx_segments(par)
    if not segs or not ''.join(s[0] for s in segs).strip():
        return None, None
    base = max(segs, key=lambda s: len(s[0]))
    out = []
    for s in segs:
        if s[1] == base[1]:
            out.append(s[0])
        else:
            m = MARK.get(s[1])
            out.append('[[%s]]%s[[/%s]]' % (m, s[0], m) if m else s[0])
    return ''.join(out), base


@app.post('/v1/docx/extract')
def docx_extract():
    _auth()
    d = docx.Document(io.BytesIO(_file_bytes()))
    blocks = []
    for i, p in enumerate(docx_pars(d)):
        txt, _ = docx_marked(p)
        if txt is not None:
            blocks.append({'id': 'p%d' % i, 'text': txt})
    return jsonify({'engine': 'docx', 'blocks': blocks})


@app.post('/v1/docx/build')
def docx_build():
    _auth()
    tr = json.loads(request.form.get('translations') or '{}')
    if not isinstance(tr, dict) or not tr:
        abort(400, 'missing translations')
    d = docx.Document(io.BytesIO(_file_bytes()))
    pars = list(docx_pars(d))
    report = {'pars': 0, 'marker_fail': 0}
    for i, p in enumerate(pars):
        txt, base = docx_marked(p)
        if txt is None:
            continue
        pid = 'p%d' % i
        new = tr.get(pid)
        if not isinstance(new, str) or not new.strip():
            abort(422, 'missing id: ' + pid)
        for m in ('B', 'I', 'BI'):
            if txt.count('[[%s]]' % m) != new.count('[[%s]]' % m) or \
               txt.count('[[/%s]]' % m) != new.count('[[/%s]]' % m):
                report['marker_fail'] += 1
                new = M_RE.sub('', new)  # fail-safe: base format, never lose text
                break
        bfont = base[2].font
        bb, bi = base[1]
        for r in list(p.runs):
            if r._element.findall(qn('w:drawing')):
                continue
            r._element.getparent().remove(r._element)

        def add_run(seg_text, flags):
            add = p.add_run(seg_text)
            add.bold, add.italic = flags
            add.font.size = bfont.size
            add.font.name = bfont.name
            if bfont.color and bfont.color.rgb is not None:
                add.font.color.rgb = bfont.color.rgb

        pos = 0
        stack = []
        for m in M_RE.finditer(new):
            if m.start() > pos:
                add_run(new[pos:m.start()], stack[-1] if stack else (bb, bi))
            pos = m.end()
            if m.group(1):
                if stack:
                    stack.pop()
            else:
                k = m.group(2)
                stack.append((k in ('B', 'BI'), k in ('I', 'BI')))
        if pos < len(new):
            add_run(new[pos:], (bb, bi))
        report['pars'] += 1
    out = io.BytesIO()
    d.save(out)
    out.seek(0)
    resp = send_file(
        out, as_attachment=True, download_name='translated.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    resp.headers['X-Build-Report'] = json.dumps(report)
    return resp


@app.get('/healthz')
def healthz():
    return jsonify({'ok': True, 'pymupdf': fitz.__doc__.split(':')[0].strip()})


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=8093)
