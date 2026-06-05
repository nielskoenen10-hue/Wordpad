import os
import uuid
import json
import sqlite3
import re
from datetime import datetime
from pathlib import Path

from flask import Flask, request, jsonify, send_from_directory, render_template_string
from PIL import Image
import io

try:
    from docx import Document as DocxDocument
    DOCX_SUPPORT = True
except ImportError:
    DOCX_SUPPORT = False

app = Flask(__name__)

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "logboek.db"
FOTOS_DIR = BASE_DIR / "fotos"
FOTOS_DIR.mkdir(exist_ok=True)

MAX_IMAGE_WIDTH = 1600
JPEG_QUALITY = 82


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS instellingen (
                sleutel TEXT PRIMARY KEY,
                waarde  TEXT
            );
            CREATE TABLE IF NOT EXISTS notities (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                aangemaakt  TEXT NOT NULL,
                vergrendeld INTEGER NOT NULL DEFAULT 0,
                html        TEXT NOT NULL DEFAULT '',
                platte_tekst TEXT NOT NULL DEFAULT ''
            );
        """)


init_db()


def html_to_plain(html: str) -> str:
    text = re.sub(r'<img[^>]*>', ' [afbeelding] ', html)
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'&nbsp;', ' ', text)
    text = re.sub(r'&lt;', '<', text)
    text = re.sub(r'&gt;', '>', text)
    text = re.sub(r'&amp;', '&', text)
    return re.sub(r'\s+', ' ', text).strip()


def compress_image(data: bytes, mime: str) -> tuple[bytes, str]:
    img = Image.open(io.BytesIO(data))
    if img.mode not in ('RGB', 'L'):
        img = img.convert('RGB')
    if img.width > MAX_IMAGE_WIDTH:
        ratio = MAX_IMAGE_WIDTH / img.width
        img = img.resize((MAX_IMAGE_WIDTH, int(img.height * ratio)), Image.LANCZOS)
    out = io.BytesIO()
    img.save(out, format='JPEG', quality=JPEG_QUALITY, optimize=True)
    return out.getvalue(), 'image/jpeg'


# ── Routes ──────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/api/instellingen', methods=['GET'])
def get_instellingen():
    with get_db() as conn:
        rows = conn.execute("SELECT sleutel, waarde FROM instellingen").fetchall()
    return jsonify({r['sleutel']: r['waarde'] for r in rows})


@app.route('/api/instellingen', methods=['POST'])
def set_instellingen():
    data = request.get_json(force=True)
    with get_db() as conn:
        for k, v in data.items():
            conn.execute(
                "INSERT INTO instellingen(sleutel,waarde) VALUES(?,?) "
                "ON CONFLICT(sleutel) DO UPDATE SET waarde=excluded.waarde",
                (k, v)
            )
    return jsonify(ok=True)


@app.route('/api/notities', methods=['GET'])
def get_notities():
    with get_db() as conn:
        rows = conn.execute(
            "SELECT id, aangemaakt, vergrendeld, html FROM notities ORDER BY id"
        ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route('/api/notities', methods=['POST'])
def maak_notitie():
    data = request.get_json(force=True)
    html = data.get('html', '')
    aangemaakt = datetime.now().isoformat(timespec='seconds')
    plain = html_to_plain(html)
    with get_db() as conn:
        cur = conn.execute(
            "INSERT INTO notities(aangemaakt, html, platte_tekst) VALUES(?,?,?)",
            (aangemaakt, html, plain)
        )
        nid = cur.lastrowid
    return jsonify(id=nid, aangemaakt=aangemaakt)


@app.route('/api/notities/<int:nid>', methods=['PUT'])
def update_notitie(nid):
    with get_db() as conn:
        row = conn.execute("SELECT vergrendeld FROM notities WHERE id=?", (nid,)).fetchone()
        if not row:
            return jsonify(error='niet gevonden'), 404
        if row['vergrendeld']:
            return jsonify(error='vergrendeld'), 403
        data = request.get_json(force=True)
        html = data.get('html', '')
        plain = html_to_plain(html)
        conn.execute(
            "UPDATE notities SET html=?, platte_tekst=? WHERE id=?",
            (html, plain, nid)
        )
    return jsonify(ok=True)


@app.route('/api/notities/<int:nid>/vergrendel', methods=['POST'])
def vergrendel(nid):
    with get_db() as conn:
        conn.execute("UPDATE notities SET vergrendeld=1 WHERE id=?", (nid,))
    return jsonify(ok=True)


@app.route('/api/zoek')
def zoek():
    q = request.args.get('q', '').strip()
    if not q:
        return jsonify([])
    like = f'%{q}%'
    with get_db() as conn:
        rows = conn.execute(
            "SELECT id, aangemaakt, vergrendeld, html FROM notities "
            "WHERE platte_tekst LIKE ? ORDER BY id DESC",
            (like,)
        ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route('/api/foto', methods=['POST'])
def upload_foto():
    if 'file' not in request.files:
        return jsonify(error='geen bestand'), 400
    f = request.files['file']
    raw = f.read()
    data, mime = compress_image(raw, f.content_type or 'image/jpeg')
    naam = f"{uuid.uuid4().hex}.jpg"
    (FOTOS_DIR / naam).write_bytes(data)
    return jsonify(url=f'/fotos/{naam}')


@app.route('/fotos/<path:naam>')
def serve_foto(naam):
    return send_from_directory(FOTOS_DIR, naam)


@app.route('/api/open', methods=['POST'])
def open_bestand():
    """Open een .txt of .docx bestand en geef HTML terug."""
    if 'file' not in request.files:
        return jsonify(error='geen bestand'), 400
    f = request.files['file']
    naam = f.filename or ''
    ext = Path(naam).suffix.lower()

    if ext == '.txt':
        tekst = f.read().decode('utf-8', errors='replace')
        regels = tekst.splitlines()
        html_delen = []
        for regel in regels:
            escaped = (regel
                       .replace('&', '&amp;')
                       .replace('<', '&lt;')
                       .replace('>', '&gt;'))
            html_delen.append(f'<p>{escaped if escaped else "<br>"}</p>')
        return jsonify(html=''.join(html_delen), naam=naam)

    if ext == '.docx':
        if not DOCX_SUPPORT:
            return jsonify(error='python-docx niet geïnstalleerd'), 501
        raw = f.read()
        doc = DocxDocument(io.BytesIO(raw))
        html_delen = []
        for para in doc.paragraphs:
            stijl = para.style.name.lower() if para.style else ''
            tag = 'p'
            if 'heading 1' in stijl:
                tag = 'h1'
            elif 'heading 2' in stijl:
                tag = 'h2'
            elif 'heading 3' in stijl:
                tag = 'h3'

            inhoud = ''
            for run in para.runs:
                tekst = (run.text
                         .replace('&', '&amp;')
                         .replace('<', '&lt;')
                         .replace('>', '&gt;'))
                if run.bold:
                    tekst = f'<strong>{tekst}</strong>'
                if run.italic:
                    tekst = f'<em>{tekst}</em>'
                if run.underline:
                    tekst = f'<u>{tekst}</u>'
                inhoud += tekst

            html_delen.append(f'<{tag}>{inhoud if inhoud else "<br>"}</{tag}>')
        return jsonify(html=''.join(html_delen), naam=naam)

    return jsonify(error=f'Bestandstype {ext} niet ondersteund'), 415


@app.route('/api/export')
def export_alle():
    with get_db() as conn:
        rows = conn.execute(
            "SELECT aangemaakt, vergrendeld, html FROM notities ORDER BY id"
        ).fetchall()
    items = [dict(r) for r in rows]
    export_tijd = datetime.now().isoformat(timespec='seconds')
    return jsonify(export=export_tijd, notities=items)


# ── HTML + JS (inline voor één-bestand-deployment) ──────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Logboek</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
    background: #e8e8e8;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
  }

  /* ── toolbar ── */
  #toolbar {
    position: sticky; top: 0; z-index: 100;
    background: #f3f3f3;
    border-bottom: 1px solid #c8c8c8;
    padding: 4px 8px;
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
    align-items: center;
    box-shadow: 0 1px 3px rgba(0,0,0,.15);
  }

  #toolbar button, #toolbar select, #toolbar input[type=color] {
    height: 28px;
    border: 1px solid #bbb;
    border-radius: 3px;
    background: white;
    cursor: pointer;
    font-size: 13px;
    padding: 0 6px;
  }
  #toolbar button:hover { background: #dde; }
  #toolbar button.actief { background: #c5d8f5; border-color: #7aabf0; }

  #toolbar select { padding: 0 2px; }
  #toolbar input[type=color] { width: 28px; padding: 2px; }

  .toolbar-sep {
    width: 1px; height: 22px;
    background: #c0c0c0;
    margin: 0 2px;
  }

  /* ── zoek + acties rechts ── */
  #toolbar-rechts {
    margin-left: auto;
    display: flex; gap: 6px; align-items: center;
  }
  #zoek-input {
    height: 28px; border: 1px solid #bbb; border-radius: 3px;
    padding: 0 8px; font-size: 13px; width: 180px;
  }

  /* ── hoofdgebied ── */
  #hoofd {
    flex: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 24px 16px 80px;
    gap: 20px;
  }

  /* ── notitie-kaart ── */
  .notitie-kaart {
    width: 100%;
    max-width: 860px;
    background: white;
    border-radius: 2px;
    box-shadow: 0 1px 4px rgba(0,0,0,.18);
    overflow: hidden;
  }
  .notitie-header {
    display: flex;
    align-items: center;
    padding: 6px 12px;
    background: #f8f8f8;
    border-bottom: 1px solid #e0e0e0;
    font-size: 12px;
    color: #666;
    gap: 8px;
  }
  .notitie-header .slot {
    font-weight: 600;
    color: #333;
  }
  .vergrendeld-badge {
    margin-left: auto;
    font-size: 11px;
    background: #ffe082;
    color: #7a5500;
    border-radius: 3px;
    padding: 1px 6px;
  }
  .notitie-editor {
    min-height: 120px;
    padding: 16px 20px;
    outline: none;
    font-size: 15px;
    line-height: 1.65;
    color: #1a1a1a;
  }
  .notitie-editor img {
    max-width: 100%;
    height: auto;
    display: block;
    margin: 8px 0;
    border-radius: 2px;
  }
  .notitie-editor[contenteditable=false] {
    background: #fafafa;
    color: #333;
  }
  .notitie-footer {
    display: flex;
    justify-content: flex-end;
    padding: 6px 12px;
    gap: 8px;
    border-top: 1px solid #efefef;
  }
  .notitie-footer button {
    font-size: 12px;
    padding: 3px 10px;
    border: 1px solid #bbb;
    border-radius: 3px;
    background: white;
    cursor: pointer;
  }
  .notitie-footer button:hover { background: #eef; }
  .notitie-footer .btn-vergrendel {
    background: #fff3cd;
    border-color: #f0c040;
    color: #7a5500;
  }

  /* ── nieuwe-notitie knop ── */
  #btn-nieuw {
    width: 100%;
    max-width: 860px;
    padding: 12px;
    background: white;
    border: 2px dashed #bbb;
    border-radius: 2px;
    font-size: 15px;
    color: #888;
    cursor: pointer;
    transition: border-color .15s, color .15s;
  }
  #btn-nieuw:hover { border-color: #7aabf0; color: #3366cc; }

  /* ── naammodal ── */
  #naam-modal {
    display: none;
    position: fixed; inset: 0;
    background: rgba(0,0,0,.45);
    z-index: 200;
    align-items: center;
    justify-content: center;
  }
  #naam-modal.zichtbaar { display: flex; }
  #naam-modal .doos {
    background: white;
    border-radius: 6px;
    padding: 32px 40px;
    max-width: 360px;
    width: 90%;
    box-shadow: 0 8px 32px rgba(0,0,0,.25);
  }
  #naam-modal h2 { margin-bottom: 12px; font-size: 18px; }
  #naam-modal p  { font-size: 13px; color: #555; margin-bottom: 20px; }
  #naam-input {
    width: 100%; padding: 8px 12px;
    border: 1px solid #bbb; border-radius: 4px;
    font-size: 15px; margin-bottom: 16px;
  }
  #naam-modal button {
    width: 100%; padding: 10px;
    background: #3366cc; color: white;
    border: none; border-radius: 4px;
    font-size: 15px; cursor: pointer;
  }
  #naam-modal button:hover { background: #254fa8; }

  /* ── drag-over overlay ── */
  body.drag-over::after {
    content: 'Laat los om afbeelding in te voegen';
    position: fixed; inset: 0;
    background: rgba(50,100,220,.18);
    border: 4px dashed #3366cc;
    z-index: 300;
    display: flex; align-items: center; justify-content: center;
    font-size: 22px; color: #3366cc; pointer-events: none;
  }
</style>
</head>
<body>

<!-- naam-modal -->
<div id="naam-modal">
  <div class="doos">
    <h2>Welkom bij Logboek</h2>
    <p>Vul je naam in. Deze wordt automatisch bij elke notitie gezet en onthouden.</p>
    <input id="naam-input" type="text" placeholder="Jouw naam" autocomplete="name">
    <button onclick="slaaNaamOp()">Opslaan &amp; beginnen</button>
  </div>
</div>

<!-- toolbar -->
<div id="toolbar">
  <!-- opmaak -->
  <button title="Vet (Ctrl+B)" onclick="fmt('bold')"><b>B</b></button>
  <button title="Cursief (Ctrl+I)" onclick="fmt('italic')"><i>I</i></button>
  <button title="Onderstrepen (Ctrl+U)" onclick="fmt('underline')"><u>U</u></button>
  <div class="toolbar-sep"></div>

  <!-- lettertype -->
  <select id="sel-font" title="Lettertype" onchange="fmtFont(this.value)">
    <option value="'Segoe UI',sans-serif">Segoe UI</option>
    <option value="Arial,sans-serif">Arial</option>
    <option value="'Times New Roman',serif">Times New Roman</option>
    <option value="'Courier New',monospace">Courier New</option>
    <option value="Georgia,serif">Georgia</option>
    <option value="Verdana,sans-serif">Verdana</option>
  </select>

  <select id="sel-size" title="Lettergrootte" onchange="fmtSize(this.value)">
    <option value="10">10</option>
    <option value="11">11</option>
    <option value="12" selected>12</option>
    <option value="14">14</option>
    <option value="16">16</option>
    <option value="18">18</option>
    <option value="20">20</option>
    <option value="24">24</option>
    <option value="28">28</option>
    <option value="36">36</option>
  </select>

  <input type="color" id="kleur-kiezer" title="Tekstkleur" value="#000000" onchange="fmtKleur(this.value)">
  <div class="toolbar-sep"></div>

  <!-- lijsten -->
  <button title="Opsomming" onclick="fmt('insertUnorderedList')">&#8226; lijst</button>
  <button title="Genummerde lijst" onclick="fmt('insertOrderedList')">1. lijst</button>
  <div class="toolbar-sep"></div>

  <!-- uitlijning -->
  <button title="Links" onclick="fmt('justifyLeft')">&#8676;</button>
  <button title="Midden" onclick="fmt('justifyCenter')">&#8596;</button>
  <button title="Rechts" onclick="fmt('justifyRight')">&#8677;</button>
  <div class="toolbar-sep"></div>

  <!-- afbeelding invoegen -->
  <button title="Afbeelding invoegen" onclick="kiesFoto()">&#128247; Foto</button>
  <input type="file" id="foto-input" accept="image/*" style="display:none" onchange="uploadFoto(this)">

  <!-- bestand openen -->
  <button title="Bestand openen (.txt of .docx)" onclick="kiesBestand()">&#128194; Openen</button>
  <input type="file" id="bestand-input" accept=".txt,.docx" style="display:none" onchange="openBestand(this)">

  <!-- rechts: zoek + export -->
  <div id="toolbar-rechts">
    <input id="zoek-input" type="search" placeholder="Zoeken…" oninput="zoek(this.value)">
    <button title="Exporteer alle notities als JSON" onclick="exporteer()">&#128229; Export</button>
  </div>
</div>

<!-- hoofdgebied -->
<div id="hoofd">
  <button id="btn-nieuw" onclick="nieuweNotitie()">+ Nieuwe notitie</button>
</div>

<script>
'use strict';

let gebruikersnaam = '';
let huidigeFocusId = null;   // id van de notitie die nu focus heeft
let bewaarTimer = null;
const notities = {};          // id -> { html, vergrendeld }

// ── initialisatie ────────────────────────────────────────────────────────

async function init() {
  const inst = await api('GET', '/api/instellingen');
  if (inst.naam) {
    gebruikersnaam = inst.naam;
  } else {
    document.getElementById('naam-modal').classList.add('zichtbaar');
    return;
  }
  await laadAlles();
}

async function slaaNaamOp() {
  const v = document.getElementById('naam-input').value.trim();
  if (!v) return;
  gebruikersnaam = v;
  await api('POST', '/api/instellingen', { naam: v });
  document.getElementById('naam-modal').classList.remove('zichtbaar');
  await laadAlles();
}

document.getElementById('naam-input').addEventListener('keydown', e => {
  if (e.key === 'Enter') slaaNaamOp();
});

async function laadAlles() {
  const lijst = await api('GET', '/api/notities');
  const hoofd = document.getElementById('hoofd');
  const btnNieuw = document.getElementById('btn-nieuw');
  // verwijder bestaande kaarten
  hoofd.querySelectorAll('.notitie-kaart').forEach(el => el.remove());
  for (const n of lijst) {
    notities[n.id] = { html: n.html, vergrendeld: n.vergrendeld };
    const kaart = maakKaart(n.id, n.aangemaakt, n.html, !!n.vergrendeld);
    hoofd.insertBefore(kaart, btnNieuw);
  }
  btnNieuw.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

// ── notitie kaart ─────────────────────────────────────────────────────────

function maakKaart(id, aangemaakt, html, vergrendeld) {
  const kaart = document.createElement('div');
  kaart.className = 'notitie-kaart';
  kaart.dataset.id = id;

  const header = document.createElement('div');
  header.className = 'notitie-header';
  const dt = new Date(aangemaakt);
  const datumStr = dt.toLocaleDateString('nl-NL', { weekday:'long', year:'numeric', month:'long', day:'numeric' });
  const tijdStr  = dt.toLocaleTimeString('nl-NL', { hour:'2-digit', minute:'2-digit' });
  header.innerHTML = `<span class="slot">${datumStr} – ${tijdStr}</span>
    <span>${escHtml(gebruikersnaam)}</span>
    ${vergrendeld ? '<span class="vergrendeld-badge">&#128274; vergrendeld</span>' : ''}`;

  const editor = document.createElement('div');
  editor.className = 'notitie-editor';
  editor.contentEditable = vergrendeld ? 'false' : 'true';
  editor.innerHTML = html || '<p><br></p>';
  editor.dataset.id = id;

  if (!vergrendeld) {
    editor.addEventListener('focus', () => { huidigeFocusId = id; });
    editor.addEventListener('blur',  () => { if (huidigeFocusId === id) huidigeFocusId = null; });
    editor.addEventListener('input', () => bewaarDebounced(id, editor));
    editor.addEventListener('paste', e => verwerkPlak(e, editor));
    editor.addEventListener('drop',  e => verwerkDrop(e, editor));
    editor.addEventListener('dragover', e => e.preventDefault());
  }

  const footer = document.createElement('div');
  footer.className = 'notitie-footer';
  if (!vergrendeld) {
    const btnOpslaan = document.createElement('button');
    btnOpslaan.textContent = '💾 Opslaan';
    btnOpslaan.onclick = () => bewaar(id, editor);

    const btnVergrendel = document.createElement('button');
    btnVergrendel.className = 'btn-vergrendel';
    btnVergrendel.textContent = '🔒 Vergrendelen';
    btnVergrendel.onclick = () => vergrendel(id, kaart, editor, header);

    footer.append(btnOpslaan, btnVergrendel);
  }

  kaart.append(header, editor, footer);
  return kaart;
}

// ── bewaren ───────────────────────────────────────────────────────────────

function bewaarDebounced(id, editor) {
  clearTimeout(bewaarTimer);
  bewaarTimer = setTimeout(() => bewaar(id, editor), 1200);
}

async function bewaar(id, editor) {
  clearTimeout(bewaarTimer);
  const html = editor.innerHTML;
  await api('PUT', `/api/notities/${id}`, { html });
  notities[id].html = html;
}

async function vergrendel(id, kaart, editor, header) {
  if (!confirm('Notitie vergrendelen? Dit kan niet ongedaan gemaakt worden.')) return;
  await bewaar(id, editor);
  await api('POST', `/api/notities/${id}/vergrendel`);
  notities[id].vergrendeld = true;
  editor.contentEditable = 'false';
  const badge = document.createElement('span');
  badge.className = 'vergrendeld-badge';
  badge.textContent = '🔒 vergrendeld';
  header.appendChild(badge);
  kaart.querySelector('.notitie-footer').innerHTML = '';
}

// ── nieuwe notitie ────────────────────────────────────────────────────────

async function nieuweNotitie(initieleHtml) {
  const html = initieleHtml || '<p><br></p>';
  const res = await api('POST', '/api/notities', { html });
  notities[res.id] = { html, vergrendeld: false };
  const kaart = maakKaart(res.id, res.aangemaakt, html, false);
  const btnNieuw = document.getElementById('btn-nieuw');
  btnNieuw.parentNode.insertBefore(kaart, btnNieuw);
  // focus op editor
  const editor = kaart.querySelector('.notitie-editor');
  editor.focus();
  // zet cursor aan einde
  const range = document.createRange();
  range.selectNodeContents(editor);
  range.collapse(false);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  kaart.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

// ── opmaak ────────────────────────────────────────────────────────────────

function actieveEditor() {
  if (huidigeFocusId !== null) {
    return document.querySelector(`.notitie-editor[data-id="${huidigeFocusId}"]`);
  }
  return null;
}

function fmt(commando) {
  const ed = actieveEditor();
  if (ed) ed.focus();
  document.execCommand(commando, false, null);
}

function fmtFont(familie) {
  const ed = actieveEditor();
  if (ed) ed.focus();
  document.execCommand('fontName', false, familie);
}

function fmtSize(pt) {
  // execCommand fontSize werkt met 1-7, we gebruiken een CSS-truc via fontsize + span
  const ed = actieveEditor();
  if (ed) ed.focus();
  document.execCommand('fontSize', false, '7');
  const spans = (ed || document).querySelectorAll('font[size="7"]');
  spans.forEach(s => {
    s.removeAttribute('size');
    s.style.fontSize = pt + 'pt';
  });
}

function fmtKleur(kleur) {
  const ed = actieveEditor();
  if (ed) ed.focus();
  document.execCommand('foreColor', false, kleur);
}

// ── foto ──────────────────────────────────────────────────────────────────

function kiesFoto() {
  document.getElementById('foto-input').click();
}

async function uploadFoto(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = '';
  await inserteerFotoBestand(file);
}

async function inserteerFotoBestand(file) {
  const form = new FormData();
  form.append('file', file);
  const res = await fetch('/api/foto', { method: 'POST', body: form });
  const data = await res.json();
  if (data.url) inserteerAfbeelding(data.url);
}

function inserteerAfbeelding(url) {
  let ed = actieveEditor();
  if (!ed) {
    // maak nieuwe notitie aan en wacht tot editor klaar is
    nieuweNotitie().then(() => {
      ed = actieveEditor();
      if (ed) _platsAfbeelding(ed, url);
    });
    return;
  }
  _platsAfbeelding(ed, url);
}

function _platsAfbeelding(ed, url) {
  ed.focus();
  document.execCommand('insertImage', false, url);
  bewaarDebounced(parseInt(ed.dataset.id), ed);
}

// ── plakken (Ctrl+V) ──────────────────────────────────────────────────────

async function verwerkPlak(e, editor) {
  const items = (e.clipboardData || window.clipboardData).items;
  for (const item of items) {
    if (item.type.startsWith('image/')) {
      e.preventDefault();
      const blob = item.getAsFile();
      await inserteerFotoBestand(blob);
      return;
    }
  }
}

// ── drag & drop ───────────────────────────────────────────────────────────

document.addEventListener('dragover', e => {
  if ([...e.dataTransfer.items].some(i => i.kind === 'file' && i.type.startsWith('image/'))) {
    e.preventDefault();
    document.body.classList.add('drag-over');
  }
});
document.addEventListener('dragleave', () => document.body.classList.remove('drag-over'));
document.addEventListener('drop', async e => {
  document.body.classList.remove('drag-over');
  if (!e.dataTransfer.files.length) return;
  for (const file of e.dataTransfer.files) {
    if (file.type.startsWith('image/')) {
      e.preventDefault();
      await inserteerFotoBestand(file);
    }
  }
});

async function verwerkDrop(e, editor) {
  // afgehandeld door de body-handler hierboven; zorg dat editor focus krijgt
  huidigeFocusId = parseInt(editor.dataset.id);
}

// ── bestand openen ────────────────────────────────────────────────────────

function kiesBestand() {
  document.getElementById('bestand-input').click();
}

async function openBestand(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = '';
  const form = new FormData();
  form.append('file', file);
  const res = await fetch('/api/open', { method: 'POST', body: form });
  if (!res.ok) {
    const err = await res.json();
    alert('Fout bij openen: ' + (err.error || res.statusText));
    return;
  }
  const data = await res.json();
  await nieuweNotitie(data.html);
}

// ── zoeken ────────────────────────────────────────────────────────────────

let zoekTimer = null;
async function zoek(q) {
  clearTimeout(zoekTimer);
  zoekTimer = setTimeout(async () => {
    const hoofd = document.getElementById('hoofd');
    const kaarten = hoofd.querySelectorAll('.notitie-kaart');
    if (!q.trim()) {
      kaarten.forEach(k => k.style.display = '');
      return;
    }
    const res = await api('GET', `/api/zoek?q=${encodeURIComponent(q)}`);
    const gevondenIds = new Set(res.map(r => String(r.id)));
    kaarten.forEach(k => {
      k.style.display = gevondenIds.has(k.dataset.id) ? '' : 'none';
    });
  }, 300);
}

// ── export ────────────────────────────────────────────────────────────────

async function exporteer() {
  const data = await api('GET', '/api/export');
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `logboek-export-${new Date().toISOString().slice(0,10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

// ── hulpfuncties ──────────────────────────────────────────────────────────

async function api(methode, pad, body) {
  const opts = { method: methode, headers: {} };
  if (body) {
    opts.headers['Content-Type'] = 'application/json';
    opts.body = JSON.stringify(body);
  }
  const res = await fetch(pad, opts);
  return res.json();
}

function escHtml(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ── keyboard shortcuts ────────────────────────────────────────────────────
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && !e.shiftKey) {
    if (e.key === 'b') { e.preventDefault(); fmt('bold'); }
    if (e.key === 'i') { e.preventDefault(); fmt('italic'); }
    if (e.key === 'u') { e.preventDefault(); fmt('underline'); }
  }
});

init();
</script>
</body>
</html>
"""

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
