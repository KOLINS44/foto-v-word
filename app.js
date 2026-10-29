/* Страница модуля: загрузка фото, рамка листа, распознавание, предпросмотр и скачивание docx. */
(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const S = { src: null, img: null, name: 'документ', quad: null, worker: null, cache: null, doc: null, bytes: null, busy: false };
  const MAX_SIDE = 3200;
  const tick = () => new Promise((r) => setTimeout(r, 30));

  // ---------- загрузка ----------
  const drop = $('drop'), file = $('file'), cam = $('cam');
  const coarse = window.matchMedia && matchMedia('(pointer: coarse)').matches;
  drop.addEventListener('click', (e) => { if (!coarse && !e.target.closest('button')) file.click(); });
  $('camBtn').addEventListener('click', () => cam.click());
  $('galBtn').addEventListener('click', () => file.click());
  cam.addEventListener('change', () => { if (cam.files[0]) load(cam.files[0], 'снимок'); cam.value = ''; });
  drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); file.click(); } });
  file.addEventListener('change', () => { if (file.files[0]) load(file.files[0]); file.value = ''; });
  for (const ev of ['dragenter', 'dragover']) drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); });
  for (const ev of ['dragleave', 'drop']) drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); });
  drop.addEventListener('drop', (e) => {
    const f = [...e.dataTransfer.files].find((x) => x.type.startsWith('image/'));
    if (f) load(f); else status('Это не картинка. Нужно фото в формате JPG, PNG или WebP.', true);
  });
  window.addEventListener('paste', (e) => {
    const it = [...(e.clipboardData ? e.clipboardData.items : [])].find((i) => i.type.startsWith('image/'));
    if (it) load(it.getAsFile(), 'снимок');
  });

  async function decodeToCanvas(blob) {
    const bmp = await createImageBitmap(blob, { imageOrientation: 'from-image' });
    const k = Math.min(1, MAX_SIDE / Math.max(bmp.width, bmp.height));
    const cv = document.createElement('canvas');
    cv.width = Math.round(bmp.width * k); cv.height = Math.round(bmp.height * k);
    const ctx = cv.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cv.width, cv.height); // прозрачный фон PNG — белый лист, а не чёрный
    ctx.drawImage(bmp, 0, 0, cv.width, cv.height);
    return cv;
  }

  async function load(blob, name) {
    try {
      setSource(await decodeToCanvas(blob));
      S.name = ((name || blob.name || 'документ').replace(/\.[^.]+$/, '') || 'документ').trim();
      drop.classList.add('hidden');
      $('work').classList.remove('hidden');
      resetResult();
      selectTab('photo');
      $('go').disabled = false;
      status('Проверьте рамку листа и нажмите «Собрать документ Word».');
    } catch (e) {
      status('Не удалось открыть файл: ' + (e.message || e) + '. Подойдёт JPG, PNG или WebP.', true);
    }
  }

  function setSource(cv) {
    S.src = cv;
    S.img = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height);
    S.quad = DocCore.detectQuad(S.img);
    drawPhoto();
  }

  // ---------- рамка листа ----------
  const photo = $('photo'), stage = $('stage');
  const handles = [0, 1, 2, 3].map((i) => {
    const h = document.createElement('div');
    h.className = 'handle';
    h.title = 'Угол листа: потяните';
    stage.appendChild(h);
    h.addEventListener('pointerdown', (e) => { e.preventDefault(); h.setPointerCapture(e.pointerId); h.dragging = true; });
    h.addEventListener('pointerup', () => { h.dragging = false; });
    h.addEventListener('pointermove', (e) => {
      if (!h.dragging) return;
      const r = photo.getBoundingClientRect(), sc = S.img.width / r.width;
      S.quad.pts[i] = { x: clampN((e.clientX - r.left) * sc, 0, S.img.width - 1), y: clampN((e.clientY - r.top) * sc, 0, S.img.height - 1) };
      S.quad.full = false;
      drawPhoto();
    });
    return h;
  });
  const clampN = (v, a, b) => Math.max(a, Math.min(b, v));

  function drawPhoto() {
    if (!S.src) return;
    const maxW = 1400, ds = Math.min(1, maxW / S.src.width);
    photo.width = Math.round(S.src.width * ds); photo.height = Math.round(S.src.height * ds);
    const ctx = photo.getContext('2d');
    ctx.drawImage(S.src, 0, 0, photo.width, photo.height);
    const p = S.quad.pts.map((q) => ({ x: q.x * ds, y: q.y * ds }));
    ctx.save();
    ctx.fillStyle = 'rgba(17,24,39,.45)';
    ctx.beginPath();
    ctx.rect(0, 0, photo.width, photo.height);
    ctx.moveTo(p[0].x, p[0].y); for (const q of p.slice(1)) ctx.lineTo(q.x, q.y); ctx.closePath();
    ctx.fill('evenodd');
    ctx.strokeStyle = '#2563eb'; ctx.lineWidth = Math.max(2, photo.width / 400);
    ctx.beginPath(); ctx.moveTo(p[0].x, p[0].y); for (const q of p.slice(1)) ctx.lineTo(q.x, q.y); ctx.closePath(); ctx.stroke();
    ctx.restore();
    placeHandles();
  }
  function placeHandles() {
    if (!S.img) return;
    const sc = photo.clientWidth / S.img.width;
    S.quad.pts.forEach((q, i) => { handles[i].style.left = q.x * sc + 'px'; handles[i].style.top = q.y * sc + 'px'; });
  }
  window.addEventListener('resize', () => { placeHandles(); fitPreview(); });

  $('rot').addEventListener('click', () => {
    const s = S.src, cv = document.createElement('canvas');
    cv.width = s.height; cv.height = s.width;
    const ctx = cv.getContext('2d');
    ctx.translate(cv.width, 0); ctx.rotate(Math.PI / 2); ctx.drawImage(s, 0, 0);
    setSource(cv);
    resetResult();
  });
  $('full').addEventListener('click', () => { S.quad = DocCore.fullQuad(S.img); drawPhoto(); });
  $('redetect').addEventListener('click', () => { S.quad = DocCore.detectQuad(S.img); drawPhoto(); });
  $('again').addEventListener('click', () => {
    S.src = S.img = S.cache = S.doc = S.bytes = null;
    $('work').classList.add('hidden'); drop.classList.remove('hidden');
    $('go').disabled = true; resetResult(); progress(0, 'Загрузите фото, чтобы начать.');
  });

  // ---------- вкладки ----------
  const tabBtns = [...document.querySelectorAll('.tabs button')];
  tabBtns.forEach((b) => b.addEventListener('click', () => !b.disabled && selectTab(b.dataset.tab)));
  function selectTab(name) {
    tabBtns.forEach((b) => b.classList.toggle('on', b.dataset.tab === name));
    document.querySelectorAll('[data-view]').forEach((v) => v.classList.toggle('hidden', v.dataset.view !== name));
    if (name === 'photo') placeHandles();
    if (name === 'preview') fitPreview();
  }

  // ---------- распознавание ----------
  function status(text, err) { const s = $('status'); s.textContent = text; s.classList.toggle('err', !!err); }
  function progress(pct, text, err) { $('barFill').style.width = Math.max(0, Math.min(100, pct)) + '%'; if (text) status(text, err); }
  function resetResult() {
    S.cache = S.doc = S.bytes = S.savedPath = null;
    $('result').classList.add('hidden');
    $('afterSave').classList.add('hidden');
    tabBtns.forEach((b) => { if (b.dataset.tab !== 'photo') b.disabled = true; });
  }

  function b64bytes(s) {
    const bin = atob(s), out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  // Движок распознавания (около 16 МБ) грузится в фоне: страница доступна сразу, фото можно снимать, пока он докачивается.
  const OCR_FILES = ['ocr/tesseract.min.js', 'ocr/engine.js', 'ocr/lang-rus.js', 'ocr/lang-eng.js'];
  let ocrLoading = null;
  function loadOcr() {
    if (window.OCR_ENGINE && window.OCR_LANGS && window.OCR_LANGS.eng && typeof Tesseract !== 'undefined') return Promise.resolve();
    if (!ocrLoading) {
      $('engine').textContent = 'Загружаю движок распознавания (около 16 МБ, один раз)…';
      ocrLoading = swControl().then(() => Promise.all(OCR_FILES.map((src) => new Promise((res, rej) => {
        const s = document.createElement('script');
        s.src = src; s.onload = res;
        s.onerror = () => rej(new Error('не загрузился файл ' + src + (location.protocol === 'file:' ? ' — рядом со страницей должна быть папка ocr' : ' — проверьте интернет')));
        document.body.appendChild(s);
      })))).then(() => {
        $('engine').textContent = navigator.serviceWorker && navigator.serviceWorker.controller
          ? 'Движок готов. Модуль сохранён на устройстве: дальше он открывается и без интернета.' : 'Движок распознавания готов.';
      }, (e) => { ocrLoading = null; $('engine').textContent = ''; throw e; });
    }
    return ocrLoading;
  }
  // сайт: файлы движка идут через сохранение на устройство (service worker), чтобы не качать их дважды
  function swControl() {
    if (!SW_ON || navigator.serviceWorker.controller) return Promise.resolve();
    return new Promise((res) => {
      navigator.serviceWorker.addEventListener('controllerchange', res, { once: true });
      setTimeout(res, 4000);
    });
  }
  const SW_ON = 'serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost');
  if (SW_ON) navigator.serviceWorker.register('sw.js').catch(() => {});
  setTimeout(() => loadOcr().catch(() => {}), 200);

  async function getWorker() {
    if (S.worker) return S.worker;
    await loadOcr();
    const E = window.OCR_ENGINE, L = window.OCR_LANGS;
    if (!E || !L || !L.rus || typeof Tesseract === 'undefined') throw new Error('рядом со страницей нет папки ocr с файлами распознавания');
    // ядро WebAssembly кладётся прямо в поток: страница открыта с диска, загружать файлы потоку нельзя
    const simd = WebAssembly.validate(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0, 1, 5, 1, 96, 0, 1, 123, 3, 2, 1, 0, 10, 10, 1, 8, 0, 65, 0, 253, 15, 253, 98, 11]));
    const url = URL.createObjectURL(new Blob([simd ? E.core : E.coreNoSimd, '\n;\n', E.worker], { type: 'application/javascript' }));
    const langs = Object.keys(L).map((code) => ({ code, data: b64bytes(L[code]) }));
    S.worker = await Tesseract.createWorker(langs, 1, {
      workerPath: url, workerBlobURL: false, corePath: url, cacheMethod: 'none', gzip: true,
      logger: (m) => { if (m.status === 'recognizing text' && S.ocrStep) S.ocrStep(25 + 70 * m.progress, 'Распознаю текст… ' + Math.round(m.progress * 100) + ' %'); },
    });
    await S.worker.setParameters({ tessedit_pageseg_mode: '3', preserve_interword_spaces: '1' });
    return S.worker;
  }

  function toCanvas(img) {
    const cv = document.createElement('canvas');
    cv.width = img.width; cv.height = img.height;
    cv.getContext('2d').putImageData(new ImageData(img.data, img.width, img.height), 0, 0);
    return cv;
  }

  function showClean(page) {
    const cv = $('clean');
    cv.width = page.w; cv.height = page.h;
    const id = new ImageData(page.w, page.h), d = id.data;
    for (let j = 0; j < page.w * page.h; j++) { d[j * 4] = d[j * 4 + 1] = d[j * 4 + 2] = page.gray[j]; d[j * 4 + 3] = 255; }
    cv.getContext('2d').putImageData(id, 0, 0);
  }

  function options() {
    return {
      font: $('optFont').value || undefined, keepBreaks: $('optBreaks').checked, highlight: $('optHl').checked, title: S.name,
    };
  }

  // Один снимок: выпрямление, очистка, линии, распознавание. Возвращает всё, что нужно для сборки docx.
  // на телефоне лист выпрямляется в меньшем разрешении: памяти у браузера там заметно меньше
  const MOBILE = /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent) || (navigator.maxTouchPoints > 1 && /Macintosh/.test(navigator.userAgent));
  const TARGET_W = MOBILE ? 1800 : 2100;

  async function processImage(img, quad, removeColor, onStep) {
    onStep(3, 'Выравниваю лист…'); await tick();
    const plan = DocCore.planPage(img, quad, TARGET_W);
    const warped = DocCore.warp(img, quad, plan.outW, plan.outH);
    onStep(10, 'Убираю тени, печати и подписи…'); await tick();
    const page = DocCore.cleanPage(warped, { removeColor });
    const lines = DocCore.detectLines(page);
    DocCore.removeBlobs(page, lines);
    const oimg = DocCore.ocrImage(page, lines);
    onStep(18, 'Запускаю распознавание…'); await tick();
    const worker = await getWorker();
    onStep(25, 'Распознаю текст…');
    S.ocrStep = onStep;
    const { data } = await worker.recognize(toCanvas(oimg), {}, { blocks: true, text: true });
    return { page, lines, data, plan, full: quad.full };
  }

  async function run() {
    if (S.busy || !S.img) return;
    S.busy = true; $('go').disabled = true; resetResult();
    try {
      S.cache = await processImage(S.img, S.quad, $('optColor').checked, progress);
      showClean(S.cache.page);
      progress(96, 'Собираю документ Word…'); await tick();
      build();
      progress(100, window.desktop ? 'Готово. Проверьте предпросмотр и сохраните документ.' : 'Готово. Проверьте предпросмотр и скачайте документ.');
      selectTab('preview');
      if (matchMedia('(max-width: 640px)').matches) $('actionCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (e) {
      console.error(e);
      progress(0, 'Ошибка: ' + (e.message || e), true);
    } finally {
      S.busy = false; $('go').disabled = false;
    }
  }
  $('go').addEventListener('click', run);

  // пересборка без повторного распознавания (смена шрифта, переносов, подсветки)
  function build() {
    const c = S.cache;
    const { doc, bytes } = DocCore.toDocx(c.page, c.lines, c.data, { ...options(), paper: c.plan.paper, full: c.full });
    if (!bytes) throw new Error('на фото не найден печатный текст');
    S.doc = doc; S.bytes = bytes;
    renderPreview(doc);
    showStats(doc, c.data);
    $('result').classList.remove('hidden');
    tabBtns.forEach((b) => { b.disabled = false; });
  }
  for (const id of ['optFont', 'optBreaks', 'optHl']) $(id).addEventListener('change', () => { if (S.cache && !S.busy) build(); });
  $('optColor').addEventListener('change', () => { if (S.cache) status('Чтобы применить, нажмите «Собрать документ Word» ещё раз.'); });

  function showStats(doc, data) {
    const words = doc.debug.chunks.flatMap((c) => c.words).filter((w) => !w.blank);
    const doubt = [...new Set(words.filter((w) => w.conf < 70).map((w) => w.text))];
    $('stats').innerHTML = `Шрифт: <b>${doc.font}, ${String(doc.bodySize).replace('.', ',')} пт</b><br>` +
      `Слов: <b>${words.length}</b>, уверенность распознавания: <b>${Math.round(data.confidence)} %</b>` +
      (doc.compat === 14 ? '<br>Файл в режиме совместимости Word 2010: так строки переносятся точно как в оригинале.' : '');
    $('doubts').innerHTML = doubt.length
      ? '<div class="note" style="margin-bottom:4px">Проверьте в Word эти слова:</div>' + doubt.slice(0, 40).map((t) => `<span class="doubt">${esc(t)}</span>`).join(' ')
      : '';
  }
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  async function save() {
    if (!S.bytes) return;
    if (window.desktop) {
      const p = await window.desktop.saveDocx(S.bytes, S.name + '.docx');
      if (!p) return;
      S.savedPath = p;
      status('Сохранено: ' + p);
      $('afterSave').classList.remove('hidden');
      return;
    }
    const blob = new Blob([S.bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = S.name + '.docx';
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }
  $('save').addEventListener('click', save);

  // телефон: отправить документ сразу в Word, Telegram, почту
  const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const canShareFiles = !window.desktop && navigator.canShare && (() => {
    try { return navigator.canShare({ files: [new File(['x'], 'x.docx', { type: DOCX_TYPE })] }); } catch (e) { return false; }
  })();
  if (canShareFiles) $('share').classList.remove('hidden');
  $('share').addEventListener('click', async () => {
    if (!S.bytes) return;
    try { await navigator.share({ files: [new File([S.bytes], S.name + '.docx', { type: DOCX_TYPE })], title: S.name }); } catch (e) { /* пользователь передумал */ }
  });
  // ---------- программа для Windows: пакетный режим, сохранение, открытие в Word ----------
  const baseName = (n) => (String(n).replace(/\.[^.]+$/, '') || 'документ').trim();

  async function batch() {
    if (S.busy || !window.desktop) return;
    const files = await window.desktop.pickImages();
    if (!files.length) return;
    const dir = await window.desktop.pickFolder();
    if (!dir) return;
    S.busy = true; $('go').disabled = true; $('batch').disabled = true;
    $('openBatchDir').classList.add('hidden');
    const done = [], failed = [];
    try {
      for (let i = 0; i < files.length; i++) {
        const f = files[i], head = `Фото ${i + 1} из ${files.length} (${f.name}): `;
        const step = (pct, text) => progress(((i + pct / 100) / files.length) * 100, head + text);
        try {
          const cv = await decodeToCanvas(new Blob([await window.desktop.readFile(f.path)]));
          const img = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height);
          const c = await processImage(img, DocCore.detectQuad(img), $('optColor').checked, step);
          step(97, 'Собираю документ Word…'); await tick();
          const { bytes } = DocCore.toDocx(c.page, c.lines, c.data, { ...options(), title: baseName(f.name), paper: c.plan.paper, full: c.full });
          if (!bytes) throw new Error('не найден печатный текст');
          done.push(await window.desktop.writeDocx(dir, baseName(f.name) + '.docx', bytes));
        } catch (e) {
          failed.push(f.name + ' — ' + (e.message || e));
        }
      }
      S.batchDir = dir;
      progress(100, `Готово: ${done.length} из ${files.length} в папке «${dir}».` + (failed.length ? ` Не получилось: ${failed.join('; ')}.` : ''), !done.length);
      if (done.length) $('openBatchDir').classList.remove('hidden');
    } finally {
      S.busy = false; $('go').disabled = !S.img; $('batch').disabled = false;
    }
  }

  if (window.desktop) {
    $('desktopOnly').classList.remove('hidden');
    $('save').textContent = 'Сохранить .docx…';
    $('batch').addEventListener('click', batch);
    $('openWord').addEventListener('click', () => S.savedPath && window.desktop.openPath(S.savedPath));
    $('showFolder').addEventListener('click', () => S.savedPath && window.desktop.showInFolder(S.savedPath));
    $('openBatchDir').addEventListener('click', () => S.batchDir && window.desktop.openPath(S.batchDir));
    window.desktop.onOpenFile((f) => load(new Blob([f.data]), f.name));
    window.desktop.onMenu((cmd) => { if (cmd === 'save') save(); if (cmd === 'batch') batch(); });
  }

  // ---------- предпросмотр (приблизительно, по модели документа) ----------
  function el(tag, css, text) { const e = document.createElement(tag); if (css) e.style.cssText = css; if (text !== undefined) e.textContent = text; return e; }

  function runEl(r) {
    const s = el('span', '', r.text);
    if (r.b) s.style.fontWeight = '700';
    if (r.i) s.style.fontStyle = 'italic';
    if (r.u) s.style.textDecoration = 'underline';
    if (r.hl) s.style.background = '#fde68a';
    if (r.size) s.style.fontSize = r.size + 'pt';
    return s;
  }

  function paraEl(p, widthPt) {
    const d = el('div', `margin-top:${p.before || 0}pt;line-height:${p.line || 12}pt;position:relative;`);
    if (p.rule) {
      d.style.height = p.line + 'pt';
      d.style.borderBottom = '0.75pt solid #000';
      d.style.marginLeft = (p.indL || 0) + 'pt'; d.style.marginRight = (p.indR || 0) + 'pt';
      return d;
    }
    d.style.fontSize = (p.size || 12) + 'pt';
    if (p.border) {
      d.style.marginLeft = (p.indL || 0) + 'pt'; d.style.marginRight = (p.indR || 0) + 'pt';
      d.style.borderBottom = '0.75pt solid #000'; d.style.paddingBottom = (p.border.space || 0) + 'pt';
    } else {
      d.style.paddingLeft = (p.indL || 0) + 'pt'; d.style.paddingRight = (p.indR || 0) + 'pt';
    }
    if (p.firstLine) d.style.textIndent = p.firstLine + 'pt';
    d.style.textAlign = p.align === 'both' ? 'justify' : p.align || 'left';
    const runs = p.runs || [];
    if (!runs.some((r) => r.tab)) {
      for (const r of runs) d.appendChild(runEl(r));
      return d;
    }
    // строка с табуляцией: куски ставятся на позиции табуляторов
    d.style.whiteSpace = 'nowrap';
    d.style.height = (p.line || 12) + 'pt';
    const offset = p.border ? p.indL || 0 : 0;
    let ti = 0, prev = p.indL || 0, seg = null;
    for (const r of runs) {
      if (r.tab) {
        const t = (p.tabs || [])[ti++] || { pos: prev + 36, val: 'left' };
        if (r.u) d.appendChild(el('span', `position:absolute;left:${prev - offset}pt;width:${Math.max(0, t.pos - prev)}pt;bottom:${(p.line || 12) * 0.2}pt;border-bottom:0.75pt solid #000`));
        seg = el('span', t.val === 'right'
          ? `position:absolute;right:${widthPt - t.pos - (p.border ? p.indR || 0 : 0)}pt;top:0`
          : `position:absolute;left:${t.pos - offset}pt;top:0`);
        d.appendChild(seg);
        prev = t.pos;
        continue;
      }
      (seg || d).appendChild(runEl(r));
    }
    return d;
  }

  function blocksEl(blocks, widthPt, doc) {
    const frag = document.createDocumentFragment();
    for (const b of blocks) {
      if (b.type === 'p') frag.appendChild(paraEl(b, widthPt));
      else if (b.type === 'grid') {
        const row = el('div', `display:flex;width:${widthPt}pt`);
        const ws = DocCore.gridCols(b, widthPt, doc.k);
        for (const [i, c] of b.cells.entries()) {
          const w = ws[i], cell = el('div', `width:${w}pt;flex:none`);
          cell.appendChild(blocksEl(c.blocks, w, doc));
          row.appendChild(cell);
        }
        frag.appendChild(row);
      } else if (b.type === 'table') {
        if (b.gapBefore > 1) frag.appendChild(el('div', `height:${b.gapBefore}pt`));
        const cols = [];
        for (let i = 0; i + 1 < b.xs.length; i++) cols.push((b.xs[i + 1] - b.xs[i]) * doc.k);
        const t = el('table', `border-collapse:collapse;table-layout:fixed;margin-left:${(b.xs[0] - doc.L) * doc.k}pt`);
        const cg = el('colgroup');
        cols.forEach((w) => cg.appendChild(el('col', `width:${w}pt`)));
        t.appendChild(cg);
        b.rows.forEach((row, ri) => {
          const tr = el('tr', `height:${row.h}pt`);
          for (const c of row.cells) {
            if (c.vmerge === 'continue') continue;
            const pad = b.pad * doc.k, w = cols.slice(c.c, c.c + c.span).reduce((a, v) => a + v, 0) - 2 * pad;
            const td = el('td', `border:0.5pt solid #000;vertical-align:top;padding:0 ${pad}pt`);
            if (c.span > 1) td.colSpan = c.span;
            if (c.vmerge === 'restart') {
              let n = 1;
              while (b.rows[ri + n] && b.rows[ri + n].cells.some((x) => x.c === c.c && x.vmerge === 'continue')) n++;
              td.rowSpan = n;
            }
            td.appendChild(blocksEl(c.blocks, w, doc));
            tr.appendChild(td);
          }
          t.appendChild(tr);
        });
        frag.appendChild(t);
      }
    }
    return frag;
  }

  function renderPreview(doc) {
    const P = doc.page, host = $('preview');
    host.innerHTML = '';
    host.style.cssText = `width:${P.w}pt;min-height:${P.h}pt;padding:${P.mT}pt ${P.mR}pt ${P.mB}pt ${P.mL}pt;` +
      `font-family:"${doc.font}", ${doc.font === 'Arial' ? 'sans-serif' : 'serif'};box-sizing:border-box;`;
    host.appendChild(blocksEl(doc.blocks, P.w - P.mL - P.mR, doc));
    fitPreview();
  }
  function fitPreview() {
    const host = $('preview'), wrap = $('previewWrap');
    if (!host.firstChild || !wrap.clientWidth) return;
    host.style.zoom = 1;
    host.style.zoom = Math.min(1.3, (wrap.clientWidth - 34) / host.offsetWidth);
  }

  // ---------- для автоматической проверки ----------
  window.DocApp = {
    async runDataURL(dataURL, name) {
      const blob = await (await fetch(dataURL)).blob();
      await load(blob, name || 'тест');
      await run();
      if (!S.bytes) throw new Error($('status').textContent);
      let bin = '';
      for (let i = 0; i < S.bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, S.bytes.subarray(i, i + 0x8000));
      const ocrLines = [];
      for (const b of S.cache.data.blocks || []) for (const pa of b.paragraphs) for (const l of pa.lines) ocrLines.push(l.words.map((w) => w.text + '(' + Math.round(w.confidence) + ')').join(' '));
      return { docx: btoa(bin), status: $('status').textContent, stats: $('stats').textContent, clean: $('clean').toDataURL('image/png'), ocr: ocrLines.join('\n'), removed: S.cache.page.removed };
    },
  };
})();
