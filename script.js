const $ = id => document.getElementById(id);
const drop = $('drop'), input = $('file'), panel = $('panel'), status = $('status'), prog = $('prog'), go = $('go'), dl = $('dl');
/* Paste the address of your converter server here (see the converter-server folder).
   Example: const CONVERTER_URL = 'https://your-converter.onrender.com';
   Leave it empty to use the browser-only Word to PDF (page count can differ from Word). */
const CONVERTER_URL = 'https://pdf-word-converter-3z6s.onrender.com/';

let file = null, dir = 'pdf2docx';

pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const setStatus = (t, c) => { status.textContent = t; status.className = c || ''; };
const setProgress = p => { prog.style.width = p + '%'; };
const baseName = n => n.replace(/\.[^.]+$/, '');

/* ---------- theme ---------- */
const themeBtn = $('theme');
function applyTheme(t){
  document.documentElement.dataset.theme = t;
  themeBtn.textContent = t === 'dark' ? '☀️' : '🌙';
  themeBtn.title = t === 'dark' ? 'Switch to light theme' : 'Switch to dark theme';
  try { localStorage.setItem('theme', t); } catch (e) {}
}
applyTheme(document.documentElement.dataset.theme === 'light' ? 'light' : 'dark');
themeBtn.onclick = () => applyTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark');

/* ---------- download button ---------- */
function showDownload(blob, name){
  if(dl.href) URL.revokeObjectURL(dl.href);
  dl.href = URL.createObjectURL(blob); dl.download = name;
  dl.textContent = 'Download ' + name; dl.hidden = false;
}
const hideDownload = () => { dl.hidden = true; };

/* ---------- direction tabs ---------- */
document.querySelectorAll('.tab').forEach(t => t.onclick = () => {
  dir = t.dataset.dir;
  document.querySelectorAll('.tab').forEach(x => { x.classList.toggle('on', x === t); x.setAttribute('aria-selected', x === t); });
  const toPdf = dir === 'docx2pdf';
  input.accept = toPdf ? '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document' : 'application/pdf,.pdf';
  $('dropTitle').textContent = toPdf ? 'Choose a Word file or drop it here' : 'Choose a PDF or drop it here';
  $('dropHint').textContent = toPdf ? 'Only .docx files. Spacing and images are kept' : 'Spacing, images and layout are kept';
  $('modes').style.display = toPdf ? 'none' : 'grid';
  $('sub').textContent = !toPdf ? 'Converted in your browser. Your file is never uploaded.'
    : CONVERTER_URL ? 'Word files are sent to your converter server and deleted right after conversion.'
    : 'Converted in your browser. Page count may differ from Word.';
  reset();
});

/* ---------- file picking ---------- */
function pick(f){
  if(!f) return;
  panel.style.display = 'block'; $('fname').textContent = f.name; hideDownload();
  const ok = dir === 'docx2pdf' ? /\.docx$/i.test(f.name) : /\.pdf$/i.test(f.name);
  if(!ok){ file = null; go.disabled = true; setStatus(dir === 'docx2pdf' ? 'Choose a .docx file.' : 'Choose a .pdf file.', 'err'); return; }
  file = f; go.disabled = false; setProgress(0); setStatus('Ready to convert.');
}
function reset(){ hideDownload(); input.value = ''; file = null; panel.style.display = 'none'; setProgress(0); }
input.addEventListener('change', e => pick(e.target.files[0]));
drop.addEventListener('keydown', e => { if(e.key === 'Enter' || e.key === ' '){ e.preventDefault(); input.click(); } });
['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => pick(e.dataTransfer.files[0]));
$('reset').onclick = reset;

/* ---------- PDF -> DOCX: exact look ----------
   Each page is drawn to an image and PINNED to the top-left corner of a page of the
   same size as the PDF page. Pinned (floating) images cannot drift or push content
   onto an extra page, so every page stays identical to the PDF. */
async function pdfToDocxExact(pdf){
  const D = docxLib, sections = [], SCALE = 2;
  for(let i = 1; i <= pdf.numPages; i++){
    setStatus('Converting page ' + i + ' of ' + pdf.numPages + '...');
    const page = await pdf.getPage(i);
    const base = page.getViewport({scale: 1}), vp = page.getViewport({scale: SCALE});
    const c = document.createElement('canvas');
    c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
    const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
    await page.render({canvasContext: ctx, viewport: vp}).promise;
    const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.92));
    const data = new Uint8Array(await blob.arrayBuffer());
    sections.push({
      properties: {page: {size: {width: Math.round(base.width * 20), height: Math.round(base.height * 20)},
        margin: {top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0}}},
      children: [new D.Paragraph({
        spacing: {before: 0, after: 0, line: 20, lineRule: D.LineRuleType.EXACT},
        children: [new D.ImageRun({
          data,
          transformation: {width: Math.round(base.width * 96 / 72), height: Math.round(base.height * 96 / 72)},
          floating: {
            horizontalPosition: {relative: D.HorizontalPositionRelativeFrom.PAGE, offset: 0},
            verticalPosition: {relative: D.VerticalPositionRelativeFrom.PAGE, offset: 0},
            behindDocument: true, allowOverlap: true, lockAnchor: true
          }
        })]
      })]
    });
    c.width = c.height = 0; setProgress(5 + 85 * i / pdf.numPages);
  }
  return new D.Document({sections});
}

/* ---------- PDF -> DOCX: editable text ---------- */
function pageToParagraphs(items){
  const rows = [];
  items.forEach(it => {
    if(!it.str.trim()) return;
    const y = it.transform[5], x = it.transform[4], h = Math.abs(it.transform[3]) || 10;
    let row = rows.find(r => Math.abs(r.y - y) < h * 0.5);
    if(!row){ row = {y, h, parts: []}; rows.push(row); }
    row.parts.push({x, s: it.str});
  });
  rows.sort((a, b) => b.y - a.y);
  const paras = []; let cur = [], prev = null;
  rows.forEach(r => {
    r.parts.sort((a, b) => a.x - b.x);
    const t = r.parts.map(p => p.s).join(' ').replace(/\s+/g, ' ').trim();
    if(!t) return;
    if(prev && (prev.y - r.y) > prev.h * 1.8){ paras.push(cur.join(' ')); cur = []; }
    cur.push(t); prev = r;
  });
  if(cur.length) paras.push(cur.join(' '));
  return paras;
}
async function pdfToDocxText(pdf){
  const D = docxLib, children = []; let chars = 0;
  for(let i = 1; i <= pdf.numPages; i++){
    setStatus('Converting page ' + i + ' of ' + pdf.numPages + '...');
    const tc = await (await pdf.getPage(i)).getTextContent();
    pageToParagraphs(tc.items).forEach((t, k) => {
      chars += t.length;
      children.push(new D.Paragraph({children: [new D.TextRun({text: t, size: 22})], spacing: {after: 160}, pageBreakBefore: i > 1 && k === 0}));
    });
    setProgress(5 + 85 * i / pdf.numPages);
  }
  if(chars < 5) throw new Error('No text found. This PDF may be scanned images. Use "Exact look" instead.');
  return new D.Document({sections: [{children}]});
}

/* ---------- DOCX -> PDF ----------
   Word's own saved page breaks are respected, every image is fully loaded before capture,
   and each page is captured at its exact size so nothing shifts between pages. */
/* Word pages that rendered taller than a real page are split into proper pages.
   Content moves to the next page when it no longer fits, and every page is fixed
   to the real page height, so the page count and page size match the Word file. */
function paginate(host){
  [...host.querySelectorAll('section.docx')].forEach(sec => {
    const cs = getComputedStyle(sec);
    const pageH = parseFloat(cs.minHeight) || 0;
    if(!pageH) return;
    const padT = parseFloat(cs.paddingTop) || 0, padB = parseFloat(cs.paddingBottom) || 0;
    const avail = pageH - padT - padB;
    const art = sec.querySelector(':scope > article');
    const fix = s => { s.style.height = s.style.minHeight = pageH + 'px'; s.style.overflow = 'hidden'; };
    if(!art || sec.offsetHeight <= pageH + 2){ fix(sec); return; }
    const secTop = sec.getBoundingClientRect().top;
    const kids = [...art.children].map(c => { const r = c.getBoundingClientRect(); return {c, top: r.top - secTop, bottom: r.bottom - secTop}; });
    const groups = []; let cur = [], start = kids.length ? kids[0].top : 0;
    kids.forEach(k => {
      if(cur.length && k.bottom - start > avail){ groups.push(cur); cur = []; start = k.top; }
      cur.push(k.c);
    });
    if(cur.length) groups.push(cur);
    const header = sec.querySelector(':scope > header'), footer = sec.querySelector(':scope > footer');
    let after = sec;
    groups.slice(1).forEach(g => {
      const page = sec.cloneNode(false), a = art.cloneNode(false);
      if(header) page.appendChild(header.cloneNode(true));
      g.forEach(n => a.appendChild(n)); page.appendChild(a);
      if(footer) page.appendChild(footer.cloneNode(true));
      fix(page); after.after(page); after = page;
    });
    fix(sec);
  });
}

async function docxToPdfServer(f){
  setStatus('Sending file to the converter... the first request can take up to a minute.'); setProgress(15);
  const fd = new FormData(); fd.append('file', f, f.name);
  let res;
  try{ res = await fetch(CONVERTER_URL.replace(/\/$/, '') + '/convert', {method: 'POST', body: fd}); }
  catch(e){ throw new Error('Could not reach the converter server. Check CONVERTER_URL and that the server is running.'); }
  setProgress(80); setStatus('Building the PDF...');
  if(!res.ok){ let m = ''; try{ m = (await res.json()).error; }catch(e){} throw new Error(m || 'The converter server could not convert this file.'); }
  return await res.blob();
}

async function docxToPdf(f){
  if(CONVERTER_URL) return docxToPdfServer(f);
  if(!window.docxPreview || !docxPreview.renderAsync) throw new Error('Could not load the Word reader. Check your internet connection and reload the page.');
  const host = $('render'); host.innerHTML = '';
  setStatus('Reading Word file...'); setProgress(5);
  await docxPreview.renderAsync(await f.arrayBuffer(), host, null, {
    className: 'docx', inWrapper: true, breakPages: true, ignoreWidth: false, ignoreHeight: false,
    ignoreLastRenderedPageBreak: false, renderHeaders: true, renderFooters: true, useBase64URL: true
  });
  if(document.fonts && document.fonts.ready) await document.fonts.ready;
  await Promise.all([...host.querySelectorAll('img')].map(im => im.decode ? im.decode().catch(() => {}) : null));
  await new Promise(r => setTimeout(r, 300));
  paginate(host);
  const pages = [...host.querySelectorAll('section.docx')];
  if(!pages.length) throw new Error('This Word file has no pages to convert.');
  let pdf = null;
  for(let i = 0; i < pages.length; i++){
    setStatus('Converting page ' + (i + 1) + ' of ' + pages.length + '...');
    const el = pages[i], w = el.offsetWidth, h = el.offsetHeight;
    const canvas = await html2canvas(el, {
      scale: 2, backgroundColor: '#ffffff', useCORS: true, scrollX: 0, scrollY: 0,
      windowWidth: Math.max(w, innerWidth) + 60,
      onclone: d => { const r = d.getElementById('render'); if(r){ r.style.left = '0'; r.style.top = '0'; } }
    });
    const pw = w * 0.75, ph = h * 0.75, o = pw > ph ? 'l' : 'p';
    if(!pdf) pdf = new jspdf.jsPDF({unit: 'pt', format: [pw, ph], orientation: o}); else pdf.addPage([pw, ph], o);
    pdf.addImage(canvas.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, pw, ph);
    setProgress(10 + 85 * (i + 1) / pages.length);
  }
  host.innerHTML = '';
  return pdf.output('blob');
}

go.onclick = async () => {
  if(!file) return;
  go.disabled = true; hideDownload(); setProgress(3); setStatus('Working...');
  try{
    if(dir === 'docx2pdf'){
      showDownload(await docxToPdf(file), baseName(file.name) + '.pdf');
    }else{
      const pdf = await pdfjsLib.getDocument({data: await file.arrayBuffer()}).promise;
      const mode = document.querySelector('input[name=mode]:checked').value;
      const doc = mode === 'exact' ? await pdfToDocxExact(pdf) : await pdfToDocxText(pdf);
      setStatus('Creating Word file...');
      showDownload(await docxLib.Packer.toBlob(doc), baseName(file.name) + '.docx');
    }
    setProgress(100); setStatus('Done. Click the button to save your file.', 'ok');
  }catch(e){
    setProgress(0); setStatus(e.message || 'Could not convert this file. It may be damaged or protected.', 'err');
  }
  go.disabled = false;
};
