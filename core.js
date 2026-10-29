/* Фото документа → Word. Ядро без DOM и зависимостей: работает в браузере (window.DocCore) и в Node.
 * Этапы: поиск листа → выпрямление → очистка от цветных печатей и подписей → поиск линий →
 * (распознавание делает вызывающая сторона) → разметка страницы → docx. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.DocCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const A4 = { w: 595.28, h: 841.89 };
  const FONTS = {
    'Times New Roman': { cap: 0.662, xh: 0.448, serif: true },
    'Arial': { cap: 0.716, xh: 0.519, serif: false },
    'Courier New': { cap: 0.571, xh: 0.423, serif: true },
  };
  const BETA = 0.8; // доля высоты строки «точно» от её верха до базовой линии в Word

  const clamp = (v, a, b) => (v < a ? a : v > b ? b : v);
  const median = (arr) => {
    if (!arr.length) return 0;
    const s = Array.from(arr).sort((a, b) => a - b), m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  };
  const quantile = (arr, q) => {
    if (!arr.length) return 0;
    const s = Array.from(arr).sort((a, b) => a - b);
    return s[clamp(Math.round(q * (s.length - 1)), 0, s.length - 1)];
  };

  function otsu(values, lo, hi) {
    const bins = 256, hist = new Float64Array(bins);
    let n = 0;
    for (let i = 0; i < values.length; i++) {
      const v = values[i];
      if (v < lo || v > hi) continue;
      hist[Math.min(bins - 1, Math.floor(((v - lo) / (hi - lo)) * bins))]++;
      n++;
    }
    if (!n) return (lo + hi) / 2;
    let sum = 0;
    for (let i = 0; i < bins; i++) sum += i * hist[i];
    let wB = 0, sumB = 0, best = 0, th = 0;
    for (let i = 0; i < bins; i++) {
      wB += hist[i];
      if (!wB) continue;
      const wF = n - wB;
      if (!wF) break;
      sumB += i * hist[i];
      const mB = sumB / wB, mF = (sum - sumB) / wF, v = wB * wF * (mB - mF) * (mB - mF);
      if (v > best) { best = v; th = i; }
    }
    return lo + ((th + 0.5) / bins) * (hi - lo);
  }

  // ---------- поиск листа ----------

  function boxDown(img, f) {
    const w = Math.max(1, Math.floor(img.width / f)), h = Math.max(1, Math.floor(img.height / f));
    const r = new Float32Array(w * h), g = new Float32Array(w * h), b = new Float32Array(w * h);
    const W = img.width, d = img.data, k = 1 / (f * f * 255);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      let sr = 0, sg = 0, sb = 0;
      for (let dy = 0; dy < f; dy++) {
        let i = ((y * f + dy) * W + x * f) * 4;
        for (let dx = 0; dx < f; dx++, i += 4) { sr += d[i]; sg += d[i + 1]; sb += d[i + 2]; }
      }
      const j = y * w + x;
      r[j] = sr * k; g[j] = sg * k; b[j] = sb * k;
    }
    return { w, h, r, g, b };
  }

  function fullQuad(img) {
    const w = img.width - 1, h = img.height - 1;
    return { pts: [{ x: 0, y: 0 }, { x: w, y: 0 }, { x: w, y: h }, { x: 0, y: h }], full: true };
  }

  function robustFit(pts) { // v = a*u + b, отбрасывая выбросы
    let use = pts;
    let a = 0, b = 0;
    for (let it = 0; it < 3; it++) {
      let su = 0, sv = 0, suu = 0, suv = 0;
      for (const p of use) { su += p.u; sv += p.v; suu += p.u * p.u; suv += p.u * p.v; }
      const n = use.length, den = n * suu - su * su;
      a = den ? (n * suv - su * sv) / den : 0;
      b = (sv - a * su) / n;
      const res = pts.map((p) => Math.abs(p.v - (a * p.u + b)));
      const lim = Math.max(1.5, quantile(res, 0.7) * 1.5);
      const next = pts.filter((p, i) => res[i] <= lim);
      if (next.length < 5) break;
      use = next;
    }
    return { a, b };
  }

  function detectQuad(img) {
    const f = Math.max(1, Math.round(Math.max(img.width, img.height) / 500));
    const s = boxDown(img, f), n = s.w * s.h;
    const L = new Float32Array(n), C = new Float32Array(n);
    for (let i = 0; i < n; i++) {
      const r = s.r[i], g = s.g[i], b = s.b[i];
      L[i] = 0.299 * r + 0.587 * g + 0.114 * b;
      C[i] = Math.max(r, g, b) - Math.min(r, g, b);
    }
    const t = otsu(L, 0, 1);
    const m = new Uint8Array(n);
    for (let i = 0; i < n; i++) m[i] = L[i] > t && C[i] < 0.3 ? 1 : 0;
    // скан или скриншот: по краям кадра почти везде светлая бумага — лист занимает весь кадр
    {
      let edge = 0, bright = 0;
      const bw = Math.max(1, Math.round(0.02 * Math.min(s.w, s.h)));
      for (let y = 0; y < s.h; y++) for (let x = 0; x < s.w; x++) {
        if (x >= bw && y >= bw && x < s.w - bw && y < s.h - bw) continue;
        edge++; if (L[y * s.w + x] > Math.max(t, 0.75)) bright++;
      }
      if (bright > 0.85 * edge) return fullQuad(img);
    }
    // крупнейшая светлая область
    const lab = new Int32Array(n).fill(-1);
    let best = -1, bestSize = 0, cur = 0;
    const stack = new Int32Array(n);
    for (let i = 0; i < n; i++) {
      if (!m[i] || lab[i] >= 0) continue;
      let sp = 0, size = 0;
      stack[sp++] = i; lab[i] = cur;
      while (sp) {
        const j = stack[--sp]; size++;
        const x = j % s.w, y = (j / s.w) | 0;
        if (x > 0 && m[j - 1] && lab[j - 1] < 0) { lab[j - 1] = cur; stack[sp++] = j - 1; }
        if (x < s.w - 1 && m[j + 1] && lab[j + 1] < 0) { lab[j + 1] = cur; stack[sp++] = j + 1; }
        if (y > 0 && m[j - s.w] && lab[j - s.w] < 0) { lab[j - s.w] = cur; stack[sp++] = j - s.w; }
        if (y < s.h - 1 && m[j + s.w] && lab[j + s.w] < 0) { lab[j + s.w] = cur; stack[sp++] = j + s.w; }
      }
      if (size > bestSize) { bestSize = size; best = cur; }
      cur++;
    }
    if (bestSize < 0.15 * n) return fullQuad(img);
    let x0 = s.w, x1 = 0, y0 = s.h, y1 = 0;
    for (let i = 0; i < n; i++) if (lab[i] === best) {
      const x = i % s.w, y = (i / s.w) | 0;
      if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y;
    }
    const inC = (x, y) => lab[y * s.w + x] === best;
    const top = [], bot = [], lef = [], rig = [];
    for (let x = Math.round(x0 + 0.1 * (x1 - x0)); x <= x1 - 0.1 * (x1 - x0); x++) {
      for (let y = y0; y <= y1; y++) if (inC(x, y)) { top.push({ u: x, v: y }); break; }
      for (let y = y1; y >= y0; y--) if (inC(x, y)) { bot.push({ u: x, v: y + 1 }); break; }
    }
    for (let y = Math.round(y0 + 0.1 * (y1 - y0)); y <= y1 - 0.1 * (y1 - y0); y++) {
      for (let x = x0; x <= x1; x++) if (inC(x, y)) { lef.push({ u: y, v: x }); break; }
      for (let x = x1; x >= x0; x--) if (inC(x, y)) { rig.push({ u: y, v: x + 1 }); break; }
    }
    if (top.length < 5 || lef.length < 5) return fullQuad(img);
    const T = robustFit(top), B = robustFit(bot), Lf = robustFit(lef), R = robustFit(rig);
    const cross = (H, V) => { // y = H.a x + H.b ; x = V.a y + V.b
      const x = (V.a * H.b + V.b) / (1 - V.a * H.a);
      return { x, y: H.a * x + H.b };
    };
    const pts = [cross(T, Lf), cross(T, R), cross(B, R), cross(B, Lf)].map((p) => ({
      x: clamp(p.x * f, 0, img.width - 1), y: clamp(p.y * f, 0, img.height - 1),
    }));
    const area = Math.abs(polyArea(pts));
    if (area < 0.15 * img.width * img.height) return fullQuad(img);
    // рамка листа — выпуклый четырёхугольник с углами 60–120°, иначе что-то пошло не так
    for (let i = 0; i < 4; i++) {
      const a = pts[(i + 3) % 4], b = pts[i], c = pts[(i + 1) % 4];
      const v1x = a.x - b.x, v1y = a.y - b.y, v2x = c.x - b.x, v2y = c.y - b.y;
      const cos = (v1x * v2x + v1y * v2y) / (Math.hypot(v1x, v1y) * Math.hypot(v2x, v2y) || 1);
      if (Math.abs(cos) > 0.5) return fullQuad(img);
    }
    const nearEdge = pts.filter((p) => p.x < 2 * f || p.y < 2 * f || p.x > img.width - 1 - 2 * f || p.y > img.height - 1 - 2 * f).length;
    return { pts, full: nearEdge === 4 && area > 0.93 * img.width * img.height };
  }

  function polyArea(p) {
    let a = 0;
    for (let i = 0; i < p.length; i++) { const q = p[(i + 1) % p.length]; a += p[i].x * q.y - q.x * p[i].y; }
    return a / 2;
  }

  // ---------- выпрямление ----------

  function solve(A, b) { // Гаусс с выбором ведущего
    const n = b.length;
    for (let c = 0; c < n; c++) {
      let p = c;
      for (let r = c + 1; r < n; r++) if (Math.abs(A[r][c]) > Math.abs(A[p][c])) p = r;
      [A[c], A[p]] = [A[p], A[c]]; [b[c], b[p]] = [b[p], b[c]];
      for (let r = c + 1; r < n; r++) {
        const k = A[r][c] / A[c][c];
        for (let j = c; j < n; j++) A[r][j] -= k * A[c][j];
        b[r] -= k * b[c];
      }
    }
    const x = new Array(n).fill(0);
    for (let r = n - 1; r >= 0; r--) {
      let s = b[r];
      for (let j = r + 1; j < n; j++) s -= A[r][j] * x[j];
      x[r] = s / A[r][r];
    }
    return x;
  }

  function homography(from, to) { // to = H(from)
    const A = [], b = [];
    for (let i = 0; i < 4; i++) {
      const { x, y } = from[i], u = to[i].x, v = to[i].y;
      A.push([x, y, 1, 0, 0, 0, -u * x, -u * y]); b.push(u);
      A.push([0, 0, 0, x, y, 1, -v * x, -v * y]); b.push(v);
    }
    const h = solve(A, b);
    return [h[0], h[1], h[2], h[3], h[4], h[5], h[6], h[7], 1];
  }

  // Размер выпрямленного листа: А4, если пропорции листа похожи, иначе фрагмент в исходных пропорциях.
  function planPage(img, quad, targetW) {
    targetW = targetW || 2100;
    const p = quad.pts, d = (a, b) => Math.hypot(a.x - b.x, a.y - b.y);
    const wq = (d(p[0], p[1]) + d(p[3], p[2])) / 2, hq = (d(p[0], p[3]) + d(p[1], p[2])) / 2;
    const asp = hq / wq;
    let paper = null;
    if (Math.abs(asp - Math.SQRT2) < (quad.full ? 0.05 : 0.12)) paper = 'A4';
    else if (!quad.full && Math.abs(asp - 1 / Math.SQRT2) < 0.06) paper = 'A4L';
    const outW = targetW;
    const outH = paper === 'A4' ? Math.round(outW * Math.SQRT2) : paper === 'A4L' ? Math.round(outW / Math.SQRT2) : Math.round(outW * asp);
    return { outW, outH, paper };
  }

  function warp(img, quad, outW, outH) {
    const H = homography([{ x: 0, y: 0 }, { x: outW, y: 0 }, { x: outW, y: outH }, { x: 0, y: outH }], quad.pts.map((q, i) => ({
      x: q.x + (i === 1 || i === 2 ? 1 : 0), y: q.y + (i >= 2 ? 1 : 0),
    })));
    const out = new Uint8ClampedArray(outW * outH * 4), src = img.data, W = img.width, Hh = img.height;
    const p = quad.pts, up = outW / Math.max(1, Math.hypot(p[1].x - p[0].x, p[1].y - p[0].y));
    // мелкий снимок растягивается в несколько раз: бикубика даёт буквам ровные края, по ним лучше распознаётся
    const cub = (t) => [((-t + 2) * t - 1) * t / 2, ((3 * t - 5) * t * t + 2) / 2, ((-3 * t + 4) * t + 1) * t / 2, (t - 1) * t * t / 2];
    for (let v = 0; v < outH; v++) {
      for (let u = 0; u < outW; u++) {
        const uu = u + 0.5, vv = v + 0.5;
        const z = H[6] * uu + H[7] * vv + 1;
        const x = (H[0] * uu + H[1] * vv + H[2]) / z - 0.5, y = (H[3] * uu + H[4] * vv + H[5]) / z - 0.5;
        if (up > 1.4) {
          const xf = Math.floor(x), yf = Math.floor(y), wx = cub(x - xf), wy = cub(y - yf), o = (v * outW + u) * 4;
          let r = 0, g = 0, b = 0;
          for (let j = 0; j < 4; j++) {
            const yy = clamp(yf - 1 + j, 0, Hh - 1);
            let rr = 0, gg = 0, bb = 0;
            for (let i = 0; i < 4; i++) {
              const k = (yy * W + clamp(xf - 1 + i, 0, W - 1)) * 4;
              rr += src[k] * wx[i]; gg += src[k + 1] * wx[i]; bb += src[k + 2] * wx[i];
            }
            r += rr * wy[j]; g += gg * wy[j]; b += bb * wy[j];
          }
          out[o] = r; out[o + 1] = g; out[o + 2] = b; out[o + 3] = 255;
          continue;
        }
        const x0 = clamp(Math.floor(x), 0, W - 2), y0 = clamp(Math.floor(y), 0, Hh - 2);
        const fx = clamp(x - x0, 0, 1), fy = clamp(y - y0, 0, 1);
        const i00 = (y0 * W + x0) * 4, i01 = i00 + 4, i10 = i00 + W * 4, i11 = i10 + 4;
        const o = (v * outW + u) * 4;
        for (let c = 0; c < 3; c++) {
          const a = src[i00 + c] + (src[i01 + c] - src[i00 + c]) * fx;
          const b = src[i10 + c] + (src[i11 + c] - src[i10 + c]) * fx;
          out[o + c] = a + (b - a) * fy;
        }
        out[o + 3] = 255;
      }
    }
    return { width: outW, height: outH, data: out };
  }

  // ---------- очистка ----------
  // Фон оценивается по блокам и делится на него по каждому каналу: тени уходят, бумага становится белой.
  // Цветные чернила (синие печати, подписи) по насыщенности после нормировки заменяются белым.
  function cleanPage(img, opts) {
    opts = opts || {};
    const removeColor = opts.removeColor !== false;
    const W = img.width, H = img.height, d = img.data;
    const B = Math.max(16, Math.round(W / 64));
    const bw = Math.ceil(W / B), bh = Math.ceil(H / B);
    const bg = [new Float32Array(bw * bh), new Float32Array(bw * bh), new Float32Array(bw * bh)];
    const vals = [[], [], []];
    for (let by = 0; by < bh; by++) for (let bx = 0; bx < bw; bx++) {
      vals[0].length = vals[1].length = vals[2].length = 0;
      for (let y = by * B; y < Math.min(H, by * B + B); y += 2) for (let x = bx * B; x < Math.min(W, bx * B + B); x += 2) {
        const i = (y * W + x) * 4;
        vals[0].push(d[i]); vals[1].push(d[i + 1]); vals[2].push(d[i + 2]);
      }
      for (let c = 0; c < 3; c++) bg[c][by * bw + bx] = quantile(vals[c], 0.85);
    }
    // блоки, почти целиком занятые чернилами, получают фон соседей
    for (let c = 0; c < 3; c++) {
      let a = bg[c];
      for (let pass = 0; pass < 2; pass++) {
        const o = new Float32Array(a.length);
        for (let y = 0; y < bh; y++) for (let x = 0; x < bw; x++) {
          let mx = 0;
          for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
            const yy = clamp(y + dy, 0, bh - 1), xx = clamp(x + dx, 0, bw - 1);
            mx = Math.max(mx, a[yy * bw + xx]);
          }
          o[y * bw + x] = Math.max(a[y * bw + x], mx * 0.9);
        }
        a = o;
      }
      const o = new Float32Array(a.length); // сглаживание 3x3
      for (let y = 0; y < bh; y++) for (let x = 0; x < bw; x++) {
        let s = 0, k = 0;
        for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
          const yy = y + dy, xx = x + dx;
          if (yy < 0 || xx < 0 || yy >= bh || xx >= bw) continue;
          s += a[yy * bw + xx]; k++;
        }
        o[y * bw + x] = Math.max(20, s / k);
      }
      bg[c] = o;
    }
    const n = W * H, Y = new Float32Array(n), BL = new Float32Array(n), CH = new Float32Array(n);
    for (let y = 0; y < H; y++) {
      const gy = clamp((y + 0.5) / B - 0.5, 0, bh - 1), y0 = Math.min(bh - 2, Math.floor(gy)), fy = bh > 1 ? gy - y0 : 0;
      for (let x = 0; x < W; x++) {
        const gx = clamp((x + 0.5) / B - 0.5, 0, bw - 1), x0 = Math.min(bw - 2, Math.floor(gx)), fx = bw > 1 ? gx - x0 : 0;
        const i = (y * W + x) * 4, j = y * W + x;
        const nn = [0, 0, 0];
        for (let c = 0; c < 3; c++) {
          const a = bg[c], i0 = Math.max(0, y0) * bw + Math.max(0, x0);
          const v00 = a[i0], v01 = bw > 1 ? a[i0 + 1] : v00, v10 = bh > 1 ? a[i0 + bw] : v00, v11 = bw > 1 && bh > 1 ? a[i0 + bw + 1] : v00;
          const bgv = (v00 + (v01 - v00) * fx) * (1 - fy) + (v10 + (v11 - v10) * fx) * fy;
          nn[c] = Math.min(1, d[i + c] / bgv);
        }
        const mx = Math.max(nn[0], nn[1], nn[2]), mn = Math.min(nn[0], nn[1], nn[2]);
        Y[j] = 0.299 * nn[0] + 0.587 * nn[1] + 0.114 * nn[2];
        BL[j] = nn[2] - (nn[0] + nn[1]) / 2;
        CH[j] = mx - mn;
      }
    }
    let removed = 0, fringe = false, spread = 0;
    const colorMask = new Uint8Array(n);
    if (removeColor) {
      // цветная окантовка букв (сглаживание шрифта ClearType на скриншоте): тёплых и холодных цветных
      // пикселей примерно поровну и все они у тёмных штрихов. У синей печати почти одни холодные, у красной — тёплые.
      let warm = 0, cool = 0, adj = 0;
      for (let y = 2; y < H - 2; y += 2) for (let x = 2; x < W - 2; x += 2) {
        const j = y * W + x, i = j * 4;
        if (Y[j] > 0.9 || CH[j] < 0.2) continue;
        const r = d[i], b = d[i + 2];
        if (r > b + 40) warm++; else if (b > r + 40) cool++; else continue;
        let near = false;
        for (let dy = -2; dy <= 2 && !near; dy++) for (let dx = -2; dx <= 2; dx++) if (Y[j + dy * W + dx] < 0.4) { near = true; break; }
        if (near) adj++;
      }
      const tot = warm + cool;
      if (tot > 300 && Math.min(warm, cool) > 0.4 * Math.max(warm, cool) && adj > 0.7 * tot) fringe = true;
    }
    if (removeColor && !fringe) {
      // «синева», усреднённая по окрестности: отдельные цветные пиксели JPEG не в счёт, синие штрихи — да.
      // Тёмные нейтральные пиксели (чёрный текст поверх печати) остаются.
      // оттенок основного текста: у некоторых сканов вся краска синеватая — это не печать
      const darkB = [], darkC = [];
      for (let j = 0; j < n; j += 5) if (Y[j] < 0.45) { darkB.push(BL[j]); darkC.push(CH[j]); }
      const b0 = darkB.length > 200 ? Math.max(0, quantile(darkB, 0.5)) : 0, c0 = darkC.length > 200 ? quantile(darkC, 0.5) : 0;
      const r = Math.max(2, Math.round(W / 500)), I = new Float64Array((W + 1) * (H + 1));
      for (let y = 0; y < H; y++) {
        let row = 0;
        for (let x = 0; x < W; x++) {
          const j = y * W + x;
          row += Y[j] < 0.85 ? BL[j] - b0 : 0;
          I[(y + 1) * (W + 1) + x + 1] = I[y * (W + 1) + x + 1] + row;
        }
      }
      for (let y = 0; y < H; y++) {
        const ya = Math.max(0, y - r), yb = Math.min(H, y + r + 1);
        for (let x = 0; x < W; x++) {
          const j = y * W + x;
          if (Y[j] > 0.92) continue;
          const xa = Math.max(0, x - r), xb = Math.min(W, x + r + 1);
          const sm = (I[yb * (W + 1) + xb] - I[ya * (W + 1) + xb] - I[yb * (W + 1) + xa] + I[ya * (W + 1) + xa]) / ((yb - ya) * (xb - xa));
          const blackNeutral = Y[j] < 0.38 && BL[j] - b0 < 0.035 && CH[j] < 0.12 + c0;
          if ((sm > 0.035 && !blackNeutral) || CH[j] > 0.25 + c0) { removed++; colorMask[j] = 1; }
        }
      }
      // печать и подпись — несколько мест на листе; если «цветное» нашлось почти везде, где есть текст, это оттенок снимка
      {
        const BS = 64, gw = Math.ceil(W / BS), gh = Math.ceil(H / BS), inkB = new Int32Array(gw * gh), colB = new Int32Array(gw * gh);
        for (let y = 0; y < H; y += 2) for (let x = 0; x < W; x += 2) {
          const j = y * W + x, g = ((y / BS) | 0) * gw + ((x / BS) | 0);
          if (Y[j] < 0.6) inkB[g]++;
          if (colorMask[j]) colB[g]++;
        }
        let withInk = 0, withCol = 0;
        for (let g = 0; g < gw * gh; g++) if (inkB[g] + colB[g] >= 20) { withInk++; if (colB[g] >= 0.02 * (BS * BS) / 4) withCol++; }
        spread = withInk ? withCol / withInk : 0;

      }
      for (let j = 0; j < n; j++) if (colorMask[j]) Y[j] = 1;
      // ореолы: рядом со стёртыми цветными штрихами остаются тёмные контуры — убираем всё, кроме насыщенно-чёрного.
      // Только если цветного на листе много (печать, подпись), а не цветная окантовка букв скриншота или оттенок скана.
      const r2 = Math.max(2, Math.round(W / 600)), C2 = removed > 0.0045 * n ? new Int32Array((W + 1) * (H + 1)) : null;
      if (C2) {
      for (let y = 0; y < H; y++) {
        let row = 0;
        for (let x = 0; x < W; x++) { row += colorMask[y * W + x]; C2[(y + 1) * (W + 1) + x + 1] = C2[y * (W + 1) + x + 1] + row; }
      }
      for (let y = 0; y < H; y++) {
        const ya = Math.max(0, y - r2), yb = Math.min(H, y + r2 + 1);
        for (let x = 0; x < W; x++) {
          const j = y * W + x;
          if (colorMask[j] || Y[j] > 0.92) continue;
          const xa = Math.max(0, x - r2), xb = Math.min(W, x + r2 + 1);
          const cnt = C2[yb * (W + 1) + xb] - C2[ya * (W + 1) + xb] - C2[yb * (W + 1) + xa] + C2[ya * (W + 1) + xa];
          if (cnt < 0.15 * (yb - ya) * (xb - xa)) continue; // одиночные цветные пиксели сжатия — не повод
          if (Y[j] < 0.3 && BL[j] - b0 < 0.03 && CH[j] < 0.12 + c0) continue;
          Y[j] = 1; removed++;
        }
      }
      }
    }
    const inkVals = [];
    for (let j = 0; j < n; j += 7) if (Y[j] < 0.8) inkVals.push(Y[j]);
    const black = inkVals.length ? quantile(inkVals, 0.03) : 0;
    const white = 0.93;
    const gray = new Uint8ClampedArray(n);
    for (let j = 0; j < n; j++) gray[j] = clamp(((Y[j] - black) / (white - black)) * 255, 0, 255);
    const th = clamp(otsu(gray, 0, 254.9), 90, 190);
    const ink = new Uint8Array(n), lineInk = new Uint8Array(n), thL = Math.min(225, th + 45);
    for (let j = 0; j < n; j++) { ink[j] = gray[j] < th ? 1 : 0; lineInk[j] = gray[j] < thL ? 1 : 0; }
    return { w: W, h: H, gray, ink, lineInk, colorMask, threshold: th, removed, fringe, spread };
  }

  // ---------- линии ----------

  function findSegments(page, horizontal, loose) {
    const { w, h } = page, ink = page.lineInk || page.ink;
    const A = horizontal ? h : w, Bn = horizontal ? w : h;
    const at = horizontal ? (a, b) => ink[a * w + b] : (a, b) => ink[b * w + a];
    const cmk = loose ? page.colorMask : null, cmAt = !cmk ? null : horizontal
      ? (a, b) => cmk[a * w + b] || (a > 0 && cmk[(a - 1) * w + b]) || (a + 1 < h && cmk[(a + 1) * w + b])
      : (a, b) => cmk[b * w + a] || (a > 0 && cmk[b * w + a - 1]) || (a + 1 < w && cmk[b * w + a + 1]);
    const minLen = horizontal ? Math.max(60, Math.round(0.035 * w)) : Math.max(50, Math.round(0.022 * h));
    const gapTol = loose ? Math.max(3, Math.round(w / 500)) : 3, fillMin = loose ? 0.8 : 0.85, maxThick = Math.max(6, Math.round(0.008 * w));
    const segs = [], open = [];
    for (let a = 0; a < A; a++) {
      const runs = [];
      let b = 0;
      while (b < Bn) {
        if (!at(a, b)) { b++; continue; }
        const s = b; let last = b, ink1 = 0, cm1 = 0;
        // разрыв линейки там, где стёрты синие чернила (вписанный от руки текст), — продолжение линии
        while (b < Bn && b - last <= gapTol + 1) {
          if (at(a, b)) { last = b; ink1++; } else if (cmAt && cmAt(a, b)) { last = b; cm1++; }
          b++;
        }
        const len = last - s + 1;
        if (len >= minLen && ink1 + cm1 >= fillMin * len && ink1 >= 0.5 * len) runs.push([s, last]);
        b = last + 1;
      }
      for (const [s, e] of runs) {
        let hit = null;
        for (const o of open) {
          if (o.a2 < a - 1) continue;
          const ov = Math.min(e, o.b2) - Math.max(s, o.b1);
          if (ov > 0.6 * Math.min(e - s, o.b2 - o.b1)) { hit = o; break; }
        }
        if (hit) { hit.b1 = Math.min(hit.b1, s); hit.b2 = Math.max(hit.b2, e); hit.a2 = a; hit.rows++; }
        else open.push({ a1: a, a2: a, b1: s, b2: e, rows: 1 });
      }
      for (let i = open.length - 1; i >= 0; i--) if (open[i].a2 < a - 1) { segs.push(open[i]); open.splice(i, 1); }
    }
    segs.push(...open);
    let out = segs.filter((s) => s.a2 - s.a1 + 1 <= maxThick);
    // сшивка ступенек слегка наклонной линии
    out.sort((p, q) => p.b1 - q.b1);
    const merged = [];
    for (const s of out) {
      const m = merged.find((t) => s.b1 - t.b2 < 12 && s.b1 >= t.b1 && Math.abs((s.a1 + s.a2) / 2 - (t.a1 + t.a2) / 2) <= maxThick);
      if (m) { m.b2 = Math.max(m.b2, s.b2); m.a1 = Math.min(m.a1, s.a1); m.a2 = Math.max(m.a2, s.a2); }
      else merged.push({ ...s });
    }
    // бледные следы (остатки снятых цветных штампов) — не линии
    const g = page.gray;
    const dark = (s) => {
      let best = 255;
      for (let a = s.a1; a <= s.a2; a++) {
        let sum = 0, n = 0;
        for (let b = s.b1; b <= s.b2; b += 2) { sum += horizontal ? g[a * w + b] : g[b * w + a]; n++; }
        best = Math.min(best, sum / n);
      }
      return best;
    };
    // линия вдоль стёртых цветных чернил — остаток штампа
    const cm = page.colorMask;
    const nearColor = (s) => {
      if (!cm) return false;
      let hit = 0, tot = 0;
      for (let b = s.b1; b <= s.b2; b += 3) {
        tot++;
        for (let a = Math.max(0, s.a1 - 4); a <= Math.min(A - 1, s.a2 + 4); a++) if (cm[horizontal ? a * w + b : b * w + a]) { hit++; break; }
      }
      return hit > 0.3 * tot;
    };
    const sideFree = (s) => {
      const strict = page.ink, d = Math.max(3, s.a2 - s.a1 + 1);
      const frac = (a0, a1) => {
        let hit = 0, tot = 0;
        for (let b = s.b1; b <= s.b2; b += 2) {
          tot++;
          for (let a = Math.max(0, a0); a <= Math.min(A - 1, a1); a++) if (strict[horizontal ? a * w + b : b * w + a]) { hit++; break; }
        }
        return tot ? hit / tot : 1;
      };
      return Math.min(frac(s.a1 - d - 3, s.a1 - d), frac(s.a2 + d, s.a2 + d + 3)) < 0.4;
    };
    return merged.filter((s) => { const dk = dark(s); return dk < 215 && !(dk > 140 && nearColor(s)) && sideFree(s); }).map((s) => horizontal
      ? { x1: s.b1, x2: s.b2, y1: s.a1, y2: s.a2, y: (s.a1 + s.a2) / 2, t: s.a2 - s.a1 + 1 }
      : { y1: s.b1, y2: s.b2, x1: s.a1, x2: s.a2, x: (s.a1 + s.a2) / 2, t: s.a2 - s.a1 + 1 });
  }

  function detectLines(page) {
    const mx = 0.012 * page.w, my = 0.012 * page.h;
    const h = findSegments(page, true).filter((s) => s.y > my && s.y < page.h - my);
    const v = findSegments(page, false).filter((s) => s.x > mx && s.x < page.w - mx);
    const longH = 0.12 * page.w, longV = 0.05 * page.h;
    for (const s of h) s.sure = s.x2 - s.x1 >= longH;
    for (const s of v) s.sure = s.y2 - s.y1 >= longV;
    // мягкий проход: прерывистые линии сканов — только кандидаты, их проверят по словам после распознавания
    const cover = (a, b, horiz) => horiz
      ? Math.abs(a.y - b.y) < 6 && Math.min(a.x2, b.x2) - Math.max(a.x1, b.x1) > 0.5 * (a.x2 - a.x1)
      : Math.abs(a.x - b.x) < 6 && Math.min(a.y2, b.y2) - Math.max(a.y1, b.y1) > 0.5 * (a.y2 - a.y1);
    for (const s of findSegments(page, true, true)) if (s.y > my && s.y < page.h - my && !h.some((t) => cover(s, t, true))) h.push(s);
    for (const s of findSegments(page, false, true)) if (s.x > mx && s.x < page.w - mx && !v.some((t) => cover(s, t, false))) v.push(s);
    return { h, v };
  }

  // Короткие отрезки, лежащие внутри распознанных слов, — это слившиеся засечки, а не линии.
  function dropTextArtifacts(lines, words) {
    const bad = (s, horiz) => !s.sure && words.some((w) => {
      if (horiz) {
        const ov = Math.min(s.x2, w.x2) - Math.max(s.x1, w.x1);
        return ov > 0.5 * (s.x2 - s.x1) && s.y >= w.y0 - 2 && s.y <= Math.max(w.y1, w.base + 4) + 2;
      }
      const ov = Math.min(s.y2, w.y1) - Math.max(s.y1, w.y0);
      return ov > 0.5 * (s.y2 - s.y1) && s.x >= w.x1 - 2 && s.x <= w.x2 + 2;
    });
    return { h: lines.h.filter((s) => !bad(s, true)), v: lines.v.filter((s) => !bad(s, false)) };
  }

  // Изображение для распознавания: найденные линии стёрты, чтобы не мешали буквам.
  function ocrImage(page, lines) {
    const { w, h } = page, g = new Uint8ClampedArray(page.gray);
    const wipe = (x1, x2, y1, y2) => {
      for (let y = Math.max(0, y1); y <= Math.min(h - 1, y2); y++)
        for (let x = Math.max(0, x1); x <= Math.min(w - 1, x2); x++) g[y * w + x] = 255;
    };
    for (const s of lines.h) if (s.sure) wipe(s.x1 - 1, s.x2 + 1, s.y1 - 1, s.y2 + 1);
    for (const s of lines.v) if (s.sure) wipe(s.x1 - 1, s.x2 + 1, s.y1 - 1, s.y2 + 1);
    page.grayNoLines = g;
    const rgba = new Uint8ClampedArray(w * h * 4);
    for (let j = 0; j < w * h; j++) { rgba[j * 4] = rgba[j * 4 + 1] = rgba[j * 4 + 2] = g[j]; rgba[j * 4 + 3] = 255; }
    return { width: w, height: h, data: rgba };
  }

  // ---------- разметка ----------

  const ASC = /[А-ЯЁA-Z0-9бйёbdfhklt]/;
  const DESC = /[друфцщДЦЩgjpqyQ,;()[\]{}\/|@_]/;
  const GARBAGE = /^[|_~=+*<>`^¦[\]{}°]+$/, SOFTJUNK = /^[-—–'".,:;!\\/()«»]+$/;

  function baseAt(line, x) {
    const b = line.baseline;
    if (!b || b.has_baseline === false || b.x1 === b.x0) return line.bbox.y1;
    return b.y0 + ((b.y1 - b.y0) * (x - b.x0)) / (b.x1 - b.x0);
  }

  // Толщина штрихов внутри слова: вертикальные (по горизонтальным пробегам) и горизонтальные (по вертикальным).
  // Отношение одной к другой — контраст: у шрифтов с засечками горизонтали волосяные.
  function strokeStats(page, x1, x2, base, xh) {
    const { w, h, ink } = page;
    const mid = [], bars = [];
    x1 = Math.max(0, x1); x2 = Math.min(w - 1, x2);
    for (let y = Math.round(base - 0.7 * xh); y <= Math.round(base - 0.35 * xh); y++) {
      if (y < 0 || y >= h) continue;
      let run = 0;
      for (let x = x1; x <= x2; x++) {
        if (ink[y * w + x]) run++;
        else if (run) { mid.push(run); run = 0; }
      }
      if (run) mid.push(run);
    }
    const ya = Math.max(0, Math.round(base - 1.08 * xh)), yb = Math.min(h - 1, Math.round(base + 0.04 * xh));
    for (let x = x1; x <= x2; x++) {
      let run = 0;
      for (let y = ya; y <= yb; y++) {
        if (ink[y * w + x]) run++;
        else if (run) { bars.push(run); run = 0; }
      }
      if (run) bars.push(run);
    }
    const m = mid.filter((r) => r < xh * 0.6), b = bars.filter((r) => r < xh * 0.35);
    return { stem: m.length >= 3 ? median(m) : 0, runs: m.length, bar: b.length >= 3 ? quantile(b, 0.3) : 0 };
  }

  const MET = (typeof module === 'object' && module.exports) ? require('./metrics.js') : globalThis.DocMetrics;
  const MIDX = new Map([...MET.chars].map((c, i) => [c, i]));
  const gi = (ch) => (MIDX.has(ch) ? MIDX.get(ch) : MIDX.get('o'));
  // ширина «по чернилам» слова в долях кегля: сумма ширин минус боковые отступы крайних знаков
  function inkWidth(text, font, bold) {
    const m = MET.fonts[font][bold ? 'b' : 'r'];
    let s = 0, first = -1, last = -1;
    for (const ch of text) { const i = gi(ch); s += m[4 * i]; if (first < 0) first = i; last = i; }
    if (first < 0) return 0;
    return (s - m[4 * first + 1] - m[4 * last + 2]) / 1000;
  }
  // ширина так, как её считает Word при переносе строк
  function advWidth(text, font, bold) {
    const m = MET.fonts[font][bold ? 'b' : 'r'];
    let s = 0;
    for (const ch of text) s += m[4 * gi(ch)];
    return s / 1000;
  }
  // площадь чернил слова обычным начертанием, в квадратах кегля
  function inkArea(text, font) {
    const m = MET.fonts[font].r;
    let s = 0;
    for (const ch of text) s += m[4 * gi(ch) + 3];
    return s / 10000;
  }
  const BOLD_GAIN = {};
  for (const f of Object.keys(MET.fonts)) {
    const r = MET.fonts[f].r, b = MET.fonts[f].b, rat = [];
    for (const ch of 'абвгдежзиклмнопрстуфхцчшщыэюяАБВГДЕЖЗИКЛМНОПРСТУФХ') { const i = gi(ch); rat.push(b[4 * i + 3] / r[4 * i + 3]); }
    BOLD_GAIN[f] = median(rat);
  }

  // Курсив: если сдвинуть верх слова влево на 0,22 высоты, столбцы чернил у наклонных букв становятся резче.
  function slantRatio(page, w) {
    const { w: W, h: Hh, ink } = page;
    const x1 = Math.max(0, Math.round(w.x1) - 2), x2 = Math.min(W - 1, Math.round(w.x2) + 2);
    const yTop = Math.max(0, Math.round(w.y0)), yBot = Math.min(Hh - 1, Math.round(w.base));
    const hgt = yBot - yTop;
    if (hgt < 8 || x2 - x1 < 2 * hgt || w.text.length < 3) return 0;
    const sharp = (sh) => {
      const col = new Float64Array(x2 - x1 + Math.ceil(sh * hgt) + 3);
      let tot = 0;
      for (let y = yTop; y <= yBot; y++) {
        const off = sh * hgt - sh * (yBot - y);
        for (let x = x1; x <= x2; x++) if (ink[y * W + x]) { col[Math.round(x - x1 + off)]++; tot++; }
      }
      let sq = 0;
      for (const c of col) sq += c * c;
      return tot ? sq / (tot * tot) : 0;
    };
    const s0 = sharp(0);
    return s0 ? sharp(0.22) / s0 : 0;
  }

  function twoMeans(vals) {
    if (vals.length < 6) return null;
    let a = quantile(vals, 0.2), b = quantile(vals, 0.85);
    for (let it = 0; it < 20; it++) {
      const A = [], Bv = [];
      for (const v of vals) (Math.abs(v - a) <= Math.abs(v - b) ? A : Bv).push(v);
      if (!A.length || !Bv.length) return null;
      a = median(A); b = median(Bv);
    }
    const nb = vals.filter((v) => Math.abs(v - b) < Math.abs(v - a)).length;
    return { a, b, nb, na: vals.length - nb };
  }

  const ITALIC_TH = 1.08;
  const STD_SIZES = [8, 9, 10, 10.5, 11, 11.5, 12, 13, 14, 15, 16, 18, 20, 22, 24, 26, 28];
  const snapSize = (s) => STD_SIZES.reduce((p, q) => (Math.abs(q - s) < Math.abs(p - s) ? q : p));

  // Крупные пятна чернил (остатки подписей и печатей, росчерки) убираются до распознавания.
  // Пятна, через которые идут длинные линии (рамки таблиц), не трогаются.
  function removeBlobs(page, lines) {
    const { w, h, ink, gray } = page, n = w * h;
    const lab = new Int32Array(n).fill(-1), stack = new Int32Array(n), comps = [];
    let id = 0;
    for (let i = 0; i < n; i++) {
      if (!ink[i] || lab[i] >= 0) continue;
      let sp = 0, x0 = w, x1 = 0, y0 = h, y1 = 0, cnt = 0;
      stack[sp++] = i; lab[i] = id;
      while (sp) {
        const j = stack[--sp], x = j % w, y = (j / w) | 0;
        cnt++;
        if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y;
        for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
          const xx = x + dx, yy = y + dy;
          if (xx < 0 || yy < 0 || xx >= w || yy >= h) continue;
          const k = yy * w + xx;
          if (ink[k] && lab[k] < 0) { lab[k] = id; stack[sp++] = k; }
        }
      }
      comps.push({ x0, x1, y0, y1, cnt });
      id++;
    }
    const medH = median(comps.filter((c) => c.cnt >= 8 && c.y1 - c.y0 < 0.05 * h).map((c) => c.y1 - c.y0 + 1)) || 20;
    const sure = lines.h.filter((s) => s.sure).concat(lines.v.filter((s) => s.sure));
    const kill = new Uint8Array(id);
    let killed = 0;
    comps.forEach((c, i) => {
      const ch = c.y1 - c.y0 + 1, cw = c.x1 - c.x0 + 1;
      if (!(ch > 3.8 * medH && cw > 1.2 * medH)) return;
      if (sure.some((s) => s.x1 <= c.x1 && s.x2 >= c.x0 && s.y1 <= c.y1 && s.y2 >= c.y0)) return;
      kill[i] = 1; killed++;
    });
    if (killed) for (let i = 0; i < n; i++) if (lab[i] >= 0 && kill[lab[i]]) { ink[i] = 0; gray[i] = 255; if (page.lineInk) page.lineInk[i] = 0; }
    return killed;
  }

  // Ширина полосы набора, при которой Word перенесёт строки там же, где в оригинале:
  // текст строки помещается, а первое слово следующей — уже нет.
  // Для каждой строки абзаца: [ширина строки, ширина строки с первым словом следующей) в пунктах.
  // Последняя строка даёт только нижнюю границу.
  function lineIntervals(b, fontName, shrink, withIndR) {
    const sp0 = advWidth(' ', fontName, false), out = [];
    const S = b.size, sp = shrink && b.align === 'both' ? 0.8 * sp0 : sp0;
    const n = b.lineWords.length;
    for (let i = 0; i < n; i++) {
      const ws = b.lineWords[i], nx = i + 1 < n ? b.lineWords[i + 1][0] : null;
      if (!ws.length || (nx && /[-‐]$/.test(ws[ws.length - 1].text))) continue;
      let wid = 0;
      ws.forEach((w, j) => { wid += advWidth(w.text, fontName, w.bold) + (j ? sp : 0); });
      const ind = (b.indL || 0) + (withIndR ? b.indR || 0 : 0) + (i === 0 ? (b.firstLine || 0) : 0);
      out.push([wid * S + ind, nx ? (wid + sp + advWidth(nx.text, fontName, nx.bold)) * S + ind : Infinity]);
    }
    return out;
  }

  function calibrateWidth(blocks, fontName, shrink) {
    const iv = [];
    for (const b of blocks) {
      if (b.type !== 'p' || !b.lineWords || b.lineWords.length < 2 || (b.align !== 'both' && b.align !== 'left')) continue;
      iv.push(...lineIntervals(b, fontName, shrink, true));
    }
    if (iv.filter((v) => v[1] < Infinity).length < 2) return null;
    const ev = [];
    for (const [lo, hi] of iv) { ev.push([lo, 1]); ev.push([hi, -1]); }
    ev.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    let cnt = 0, best = -1, lo = 0, hi = 0;
    for (let i = 0; i + 1 < ev.length; i++) {
      cnt += ev[i][1];
      if (cnt > best) { best = cnt; lo = ev[i][0]; hi = ev[i + 1][0]; }
    }
    if (!isFinite(hi)) hi = lo + 6;
    return { width: lo + 0.35 * (hi - lo), hits: best, total: iv.length };
  }

  // Ширины колонок шапки (пт): все, кроме последней, — как на фото; последняя забирает остаток полосы.
  function gridCols(b, widthPt, k) {
    const m = b.cells.map((c) => (c.x2 - c.x1) * k);
    const rest = widthPt - m.slice(0, -1).reduce((a, v) => a + v, 0);
    if (rest >= 0.6 * m[m.length - 1]) return m.slice(0, -1).concat([rest]);
    const tot = m.reduce((a, v) => a + v, 0);
    return m.map((v) => (v * widthPt) / tot);
  }

  // Абзацы в ячейках и с рваным краем: правый отступ такой, чтобы Word перенёс строки как в оригинале.
  function fitParagraphs(blocks, widthPt, fontName, shrink, k) {
    for (const b of blocks) {
      if (b.type === 'p' && b.lineWords && b.lineWords.length >= 2 && b.align !== 'center') {
        const iv = lineIntervals(b, fontName, shrink, false);
        if (iv.length < 2) continue;
        const lo = Math.max(...iv.map((v) => v[0])), hi = Math.min(...iv.map((v) => v[1]));
        const avail = widthPt - (b.indR || 0);
        if (lo >= hi || (avail >= lo && avail < hi)) continue;
        const target = lo + 0.5 * (Math.min(hi, lo + 12) - lo);
        if (target <= widthPt) b.indR = widthPt - target;
      } else if (b.type === 'grid') {
        const ws = gridCols(b, widthPt, k);
        b.cells.forEach((c, i) => fitParagraphs(c.blocks, ws[i], fontName, shrink, k));
      } else if (b.type === 'table') {
        for (const row of b.rows) for (const c of row.cells) fitParagraphs(c.blocks, (b.xs[c.c + c.span] - b.xs[c.c] - 2 * b.pad) * k, fontName, shrink, k);
      }
    }
  }

  function analyze(page, lines, ocr, opts) {
    opts = opts || {};
    const W = page.w;
    // 1. слова; строки, которые движок склеил через колонки, режутся по большим пробелам
    const segs = [];
    for (const bl of ocr.blocks || []) for (const pa of bl.paragraphs || []) for (const ln of pa.lines || []) {
      const ws = (ln.words || []).filter((w) => w.text && w.text.trim()).map((w) => {
        const t0 = w.text.trim(), t = t0.replace(/^[\\/|¦√]+(?=[А-ЯЁа-яёA-Za-z0-9«"(])/, '');
        const b = w.bbox, cut = t0.length > t.length ? Math.round(((b.x1 - b.x0) * (t0.length - t.length)) / t0.length) : 0;
        return { text: t, x1: b.x0 + cut, x2: b.x1, y0: b.y0, y1: b.y1, conf: w.confidence, base: baseAt(ln, (b.x0 + b.x1) / 2) };
      }).filter((w) => w.text).sort((a, b) => a.x1 - b.x1);
      if (!ws.length) continue;
      for (let i = 0; i < ws.length; i++) {
        const w = ws[i];
        w.text = w.text.replace(/^(?:N[eе°oо]|M[eе]|Мe|Ме)(?=$|[0-9lI|),.;:])/, '№').replace(/^№[lI|]/, '№1')
          .replace(/((?:^|[^А-Яа-яЁё])(?:кв|д|ст|ч|п|корп|стр|оф)\.)б(?=$|[,.;])/, '$16');
        if (/^[|lI]$/.test(w.text) && i > 0 && /№$/.test(ws[i - 1].text)) w.text = '1';
      }
      const hgt = median(ws.map((w) => w.y1 - w.y0)) || 20;
      for (let i = ws.length - 1; i > 0; i--) { // осколки одного слова: рамки вплотную
        const a = ws[i - 1], b = ws[i];
        if (b.x1 - a.x2 < 0.12 * hgt && Math.min(a.conf, b.conf) < 75 && /[А-ЯЁа-яё]$/.test(a.text) && /^[а-яё]/.test(b.text)) {
          a.text += b.text; a.x2 = Math.max(a.x2, b.x2); a.y0 = Math.min(a.y0, b.y0); a.y1 = Math.max(a.y1, b.y1); a.conf = Math.min(a.conf, b.conf);
          ws.splice(i, 1);
        }
      }
      const parts = [];
      const gaps = ws.slice(1).map((w, i) => w.x1 - ws[i].x2), mg = median(gaps) || 0;
      let cur = [];
      for (const w of ws) {
        const g = cur.length ? w.x1 - cur[cur.length - 1].x2 : 0;
        if (cur.length && g > Math.max(1.6 * hgt, 2.5 * Math.min(mg, 1.1 * hgt))) { parts.push(cur); cur = []; }
        cur.push(w);
      }
      if (cur.length) parts.push(cur);
      // базовая линия куска — по низу слов без выносных элементов (общая линия движка врёт, если он склеил колонки)
      for (const seg of parts) {
        const flat = seg.filter((w) => !DESC.test(w.text));
        if (parts.length === 1) { /* целая строка: базовая линия движка точнее */ }
        else if (flat.length) { const b = median(flat.map((w) => w.y1)); for (const w of seg) w.base = b; }
        else { const b = quantile(seg.map((w) => w.y1), 0.2) - 0.2 * hgt; for (const w of seg) w.base = b; }
        segs.push(seg);
      }
    }
    // мусор: остатки печатей и подписей, рукописные пометки
    const CYR = /[А-ЯЁа-яё]/, LAT = /[A-Za-z]/;
    const flatW = segs.flat();
    const docCyr = flatW.filter((w) => CYR.test(w.text)).length > 0.7 * flatW.length;
    if (docCyr) { // латинские двойники в русском тексте: AM. → АМ., OOO → ООО
      const LAT2CYR = { A: 'А', B: 'В', C: 'С', E: 'Е', H: 'Н', K: 'К', M: 'М', O: 'О', P: 'Р', T: 'Т', X: 'Х', Y: 'У', a: 'а', c: 'с', e: 'е', o: 'о', p: 'р', x: 'х', y: 'у' };
      for (const w of flatW) {
        if (/[A-Za-z]/.test(w.text) && !/[@\/]/.test(w.text) && [...w.text].every((ch) => !/[A-Za-z]/.test(ch) || LAT2CYR[ch]) && (CYR.test(w.text) || /^[A-Z.]+[,.;:]?$/.test(w.text)))
          w.text = [...w.text].map((ch) => LAT2CYR[ch] || ch).join('');
      }
    }
    const tLines = [];
    for (const s of segs) {
      const cyr = docCyr ? s.length + 1 : s.filter((w) => CYR.test(w.text)).length, lat = s.filter((w) => LAT.test(w.text) && !CYR.test(w.text)).length;
      const keep = s.filter((w, i) => {
        const t = w.text;
        if (GARBAGE.test(t) && (w.conf < 90 || s.length === 1)) return false;
        if (SOFTJUNK.test(t) && w.conf < 50) return false;
        if (w.conf >= 60) return !(LAT.test(t) && !CYR.test(t) && cyr > lat && w.conf < 88 && !/[@.:/]/.test(t));
        if (/@/.test(t) && w.conf >= 15) return true;
        // неуверенное слово остаётся, только если это русское слово вплотную к уверенному
        if (!CYR.test(t) || LAT.test(t)) return false;
        return [s[i - 1], s[i + 1]].some((o) => o && o.conf >= 80 && Math.min(Math.abs(o.x1 - w.x2), Math.abs(w.x1 - o.x2)) < (w.y1 - w.y0) * 1.2);
      });
      if (!keep.length) continue;
      const conf = keep.reduce((a, w) => a + w.conf, 0) / keep.length;
      if (conf < 50 || Math.max(...keep.map((w) => w.conf)) < 60) continue;
      tLines.push({ words: keep });
    }
    const allWords = tLines.flatMap((L) => L.words);
    if (!allWords.length) return { empty: true };
    lines = dropTextArtifacts(lines, segs.flat());

    // 2. высота прописных над базовой линией
    for (const w of allWords) {
      const hgt = w.base - w.y0;
      if (hgt > 2) { if (ASC.test(w.text)) w.capH = hgt; else w.xH = hgt; }
    }
    for (const L of tLines) {
      const caps = L.words.filter((w) => w.capH).map((w) => w.capH), xhs = L.words.filter((w) => w.xH).map((w) => w.xH);
      L.capH = caps.length ? median(caps) : xhs.length ? median(xhs) / 0.7 : 0;
    }
    const roughCap = median(tLines.filter((L) => L.capH).map((L) => L.capH)) || 20;

    // 3. шрифт: у шрифтов с засечками горизонтальные штрихи заметно тоньше вертикальных
    const contr = [];
    for (const L of tLines) for (const w of L.words) {
      const cap = L.capH || roughCap;
      if (w.x2 - w.x1 < cap) continue;
      const st = strokeStats(page, Math.round(w.x1), Math.round(w.x2), w.base, cap * 0.68);
      if (st.stem && st.bar) contr.push(st.stem / st.bar);
    }
    const contrast = median(contr);
    const xhr = median(allWords.filter((w) => w.xH).map((w) => w.xH)) / roughCap;
    // какой шрифт лучше объясняет соотношение ширин слов в строке (машинопись — моноширинная)
    const resid = {};
    for (const F of Object.keys(FONTS)) {
      const rs = [];
      for (const L of tLines) {
        const ws = L.words.filter((w) => w.text.length >= 3 && w.conf >= 70);
        if (ws.length < 4) continue;
        const e = ws.map((w) => (w.x2 - w.x1) / inkWidth(w.text, F, false)), m = median(e);
        rs.push(median(e.map((v) => Math.abs(Math.log(v / m)))));
      }
      resid[F] = rs.length >= 3 ? median(rs) : null;
    }
    let fontName = opts.font && FONTS[opts.font] ? opts.font : null;
    if (!fontName && resid['Courier New'] !== null && resid['Courier New'] < 0.75 * Math.min(resid['Times New Roman'], resid.Arial)) fontName = 'Courier New';
    if (!fontName) fontName = (contrast < 1.3 && xhr > 0.74) || contrast < 1.15 ? 'Arial' : 'Times New Roman';
    const FM = FONTS[fontName];

    // 4. кегль куска строки — по ширине слов в выбранном шрифте, с проверкой по высоте прописных
    let capQ = 1, emRef = roughCap / FM.cap; // поправка на шрифт документа (узкий/широкий) и ориентир кегля
    const lineEm = (L) => {
      const capEm = ((L.capH || roughCap) / FM.cap) * capQ, ems = [];
      for (const w of L.words) {
        if (w.text.length < 2) continue;
        const iw = inkWidth(w.text, fontName, w.bold);
        if (iw > 0) ems.push({ w, iw, e: (w.x2 - w.x1) / iw });
      }
      if (!ems.length) return capEm;
      const m = median(ems.map((o) => o.e));
      let sa = 0, sp = 0, agree = 0;
      for (const o of ems) if (Math.abs(o.e - m) < 0.12 * m) { sa += o.w.x2 - o.w.x1; sp += o.iw; agree++; }
      const wem = sp ? sa / sp : m;
      L.wem = agree >= 3 && agree >= 0.6 * ems.length ? wem : 0;
      if (Math.abs(wem - capEm) < 0.22 * capEm) return wem;
      if (agree >= 2 && agree >= 0.5 * ems.length) return wem;
      return Math.abs(Math.log(wem / emRef)) < Math.abs(Math.log(capEm / emRef)) ? wem : capEm;
    };
    for (const L of tLines) L.em = lineEm(L);
    {
      const qs = tLines.filter((L) => L.wem && L.capH).map((L) => L.wem / (L.capH / FM.cap));
      if (qs.length >= 3) {
        capQ = clamp(median(qs), 0.6, 1.5);
        emRef = median(tLines.filter((L) => L.wem).map((L) => L.wem));
        for (const L of tLines) L.em = lineEm(L);
      }
    }

    // 5. жирный: краски в слове больше, чем положено обычному начертанию (размытие её не меняет)
    const ratios = [];
    for (const L of tLines) for (const w of L.words) {
      const iw = inkWidth(w.text, fontName, false);
      const wem = w.text.length >= 4 && iw > 0 ? (w.x2 - w.x1) / iw : L.em;
      const em = Math.abs(wem / L.em - 1) < 0.25 ? wem : L.em;
      const area = inkArea(w.text, fontName) * em * em;
      w.ratio = 0;
      if (area <= 0 || w.text.length < 2) continue;
      let mass = 0;
      const xa = Math.max(0, Math.round(w.x1) - 2), xb = Math.min(W - 1, Math.round(w.x2) + 2);
      const ya = Math.max(0, Math.round(w.y0) - 2), yb = Math.min(page.h - 1, Math.round(w.y1) + 2);
      const gm = page.grayNoLines || page.gray;
      for (let y = ya; y <= yb; y++) for (let x = xa; x <= xb; x++) mass += 255 - gm[y * W + x];
      w.ratio = mass / 255 / area;
      if (w.conf >= 70 && w.text.length >= 3) ratios.push(w.ratio);
    }
    const regular = quantile(ratios, 0.35);
    // без явных двух групп жирным считается только то, где краски заметно больше обычного начертания
    let boldTh = regular * Math.sqrt(BOLD_GAIN[fontName]);
    const lr = ratios.map(Math.log), km = twoMeans(lr);
    let boldD = null;
    if (km && km.b - km.a > Math.log(1.17) && km.nb >= 2) {
      // группы должны быть разделены, а не быть половинками одного размытого «горба» (D Эшмана)
      const A = lr.filter((v) => Math.abs(v - km.a) <= Math.abs(v - km.b)), B = lr.filter((v) => Math.abs(v - km.b) < Math.abs(v - km.a));
      const sd = (arr, m) => Math.sqrt(arr.reduce((acc, v) => acc + (v - m) * (v - m), 0) / Math.max(1, arr.length - 1));
      const D = (Math.SQRT2 * (km.b - km.a)) / (Math.sqrt(sd(A, km.a) ** 2 + sd(B, km.b) ** 2) || 1e-6);
      // жирного в документах обычно меньше трети: если «жирной» вышла почти половина слов — нужен очень чёткий разрыв
      const share = B.length / lr.length;
      if (D >= (share > 0.4 ? 3.5 : 2.2)) boldTh = Math.exp((km.a + km.b) / 2);
      boldD = D;
    }
    if (fontName === 'Courier New') boldTh = Infinity; // машинопись: густота ленты гуляет, жирного там почти не бывает
    for (const L of tLines) {
      const z = L.words.map((w) => (w.ratio ? Math.log(w.ratio / boldTh) : null));
      L.words.forEach((w, i) => {
        if (z[i] === null) { w.bold = null; return; }
        const hasP = z[i - 1] !== null && z[i - 1] !== undefined, hasN = z[i + 1] !== null && z[i + 1] !== undefined;
        const nw = hasP && hasN ? 0.2 : 0.4;
        let s = 0.6 * z[i], k = 0.6;
        if (hasP) { s += nw * z[i - 1]; k += nw; }
        if (hasN) { s += nw * z[i + 1]; k += nw; }
        w.bold = s / k > 0;
      });
      // короткие слова и знаки берут жирность соседей
      for (let i = 0; i < L.words.length; i++) {
        const w = L.words[i];
        if (w.bold !== null && w.text.length > 2) continue;
        const p = L.words[i - 1], q = L.words[i + 1];
        if (w.bold === null) w.bold = p && p.bold !== null ? p.bold : q && q.bold !== null ? q.bold : false;
        else if (p && q && p.bold === q.bold && p.bold !== null) w.bold = p.bold;
      }
    }
    for (const L of tLines) { L.em = lineEm(L); for (const w of L.words) { w.lineEm = L.em; w.em = L.em; } }
    for (const L of tLines) {
      for (const w of L.words) w.slant = slantRatio(page, w);
      L.words.forEach((w, i) => {
        const near = [L.words[i - 1], w, L.words[i + 1]].filter((x) => x && x.slant);
        w.italic = near.length ? median(near.map((x) => x.slant)) > ITALIC_TH : false;
      });
    }

    // основной кегль — самый частый по числу знаков
    const roughEm = roughCap / FM.cap;
    const hist = new Map();
    for (const L of tLines) {
      const key = Math.round(L.em * 2) / 2;
      hist.set(key, (hist.get(key) || 0) + L.words.reduce((s, w) => s + w.text.length, 0));
    }
    let bodyEm = roughEm, bestN = -1;
    for (const [k] of hist) {
      const near = [...hist].filter(([k2]) => Math.abs(k2 - k) <= k * 0.04).reduce((s, [, v2]) => s + v2, 0);
      if (near > bestN) { bestN = near; bodyEm = k; }
    }
    {
      const near = tLines.filter((L) => Math.abs(L.em - bodyEm) <= bodyEm * 0.06).map((L) => L.em);
      if (near.length) bodyEm = median(near);
    }

    // масштаб: кегль основного текста — стандартный
    let left = Infinity, right = -Infinity;
    for (const w of allWords) { left = Math.min(left, w.x1); right = Math.max(right, w.x2); }
    for (const s of lines.h) { left = Math.min(left, s.x1); right = Math.max(right, s.x2); }
    let textW = right - left;
    let paper = opts.paper || null;
    if (paper && opts.full && left / W < 0.06) paper = null; // кадр без полей — это фрагмент, а не лист целиком
    let pageW = paper === 'A4L' ? A4.h : A4.w, pageH = paper === 'A4L' ? A4.w : A4.h;
    let bodySize, fragment = !paper;
    if (paper) bodySize = snapSize((bodyEm * pageW) / W);
    else {
      // фрагмент без краёв листа: кегль подбирается так, чтобы строки легли в полосу А4;
      // широкий кадр (бланки вроде ТОРГ-12) может оказаться альбомным листом
      const wide = W / page.h > 1.15;
      let best = null;
      for (const s of [12, 14, 13, 11, 10, 9, 8, 7]) {
        const cw = (textW * s) / bodyEm;
        for (const land of wide ? [false, true] : [false]) {
          const target = land ? 735 : 467.7, maxW = land ? 790 : 515;
          const score = Math.abs(cw - target) + (cw > maxW ? 1000 : 0) + (s < 10 ? 30 : 0) + (s === 12 || s === 14 ? 0 : 10) + (land ? 20 : 0);
          if (!best || score < best.score) best = { s, score, land };
        }
      }
      bodySize = best.s;
      if (best.land) { pageW = A4.h; pageH = A4.w; }
    }
    const k = bodySize / bodyEm; // пт на пиксель
    const sizeOf = (em) => {
      const s = em * k;
      if (Math.abs(s - bodySize) <= Math.max(0.6, bodySize * 0.05)) return bodySize;
      return snapSize(s);
    };
    const fr = contrast;

    // 5. куски строк, разделённые большими пробелами
    const chunks = [];
    for (const L of tLines) {
      let cur = null;
      const gaps = L.words.slice(1).map((w, i) => w.x1 - L.words[i].x2), mg = median(gaps) || 0;
      for (const w of L.words) {
        if (cur && w.x1 - cur.x2 > Math.max(1.8 * L.em, 2.5 * Math.min(mg, 0.8 * L.em))) { chunks.push(cur); cur = null; }
        if (!cur) cur = { kind: 'text', words: [], x1: w.x1, x2: w.x2, y0: w.y0, y1: w.y1, em: L.em };
        cur.words.push(w); cur.x2 = Math.max(cur.x2, w.x2); cur.y0 = Math.min(cur.y0, w.y0); cur.y1 = Math.max(cur.y1, w.y1);
      }
      if (cur) chunks.push(cur);
    }
    // рукописные числа (входящие номера, пометки на полях): крупные одиночные числа без букв
    const gridBoxes = findTables(lines, bodyEm).map((t) => ({ x1: t.xs[0], x2: t.xs[t.xs.length - 1], y0: t.ys[0], y1: t.ys[t.ys.length - 1] }));
    const inGrid = (c) => gridBoxes.some((g) => (c.x1 + c.x2) / 2 > g.x1 && (c.x1 + c.x2) / 2 < g.x2 && (c.y0 + c.y1) / 2 > g.y0 && (c.y0 + c.y1) / 2 < g.y1);
    for (let i = chunks.length - 1; i >= 0; i--) {
      const c = chunks[i];
      const noLetters = c.words.every((w) => !/[А-ЯЁа-яёA-Za-z]/.test(w.text));
      if (c.words.length <= 2 && c.em > 1.5 * bodyEm && noLetters) chunks.splice(i, 1);
      else if (c.words.length === 1 && /^\d{1,3}$/.test(c.words[0].text) && c.words[0].conf < 90 && !inGrid(c)) chunks.splice(i, 1); // росчерк, прочитанный как цифры
    }
    for (const c of chunks) {
      c.base = median(c.words.map((w) => w.base));
      const tol = c.words.length === 1 ? 0.2 : c.words.length <= 3 ? 0.12 : 0;
      if (tol && Math.abs(c.em / bodyEm - 1) < tol) { c.em = bodyEm; for (const w of c.words) { w.lineEm = bodyEm; w.em = bodyEm; } }
      c.size = sizeOf(c.em);
    }
    left = Math.min(...chunks.map((c) => c.x1), ...lines.h.map((l) => l.x1));
    right = Math.max(...chunks.map((c) => c.x2), ...lines.h.map((l) => l.x2));
    textW = right - left;

    // 6. таблицы по сетке линий
    const tables = findTables(lines, bodyEm);
    const used = new Set();
    for (const t of tables) for (const s of t.hs.concat(t.vs)) used.add(s);
    const rules = lines.h.filter((s) => !used.has(s) && s.x2 - s.x1 > 1.5 * bodyEm).map((s) => ({ kind: 'rule', x1: s.x1, x2: s.x2, y: s.y, y0: s.y1, y1: s.y2, base: s.y }));

    const ctx = { k, bodyEm, bodySize, sizeOf, left, right, log: [], font: fontName };
    let items = [];
    for (const t of tables) {
      t.kind = 'table';
      t.x1 = t.xs[0]; t.x2 = t.xs[t.xs.length - 1]; t.y0 = t.ys[0]; t.y1 = t.ys[t.ys.length - 1]; t.base = t.y0;
      t.words = [];
    }
    for (const c of chunks) {
      const cx = (c.x1 + c.x2) / 2, cy = (c.y0 + c.y1) / 2;
      const t = tables.find((t) => cx > t.x1 && cx < t.x2 && cy > t.y0 && cy < t.y1);
      if (t) t.words.push(...c.words.map((w) => ({ ...w, em: c.em, size: c.size })));
      else items.push(c);
    }
    const freeRules = [];
    for (const r of rules) {
      const cx = (r.x1 + r.x2) / 2;
      const t = tables.find((t) => cx > t.x1 && cx < t.x2 && r.y > t.y0 && r.y < t.y1);
      if (t) (t.rules = t.rules || []).push(r); else freeRules.push(r);
    }
    items = items.concat(freeRules, tables);

    // 7. страница: поля
    let contentW = textW * k;
    const marginL = paper ? Math.max(14, left * k) : clamp(pageW - contentW - 42.5, 20, 85.05);
    let marginR = pageW - marginL - contentW;
    let mL = marginL;
    if (marginR < 20) { mL = Math.max(14, pageW - contentW - 20); marginR = Math.max(14, pageW - mL - contentW); }
    const firstTop = Math.min(...items.map((i) => i.y0));
    const marginT = paper ? Math.max(14, firstTop * k - 2) : 56.7;
    ctx.L = left; ctx.R = right; ctx.pageTopPx = firstTop; ctx.marginT = marginT;
    ctx.yPt = (ypx) => marginT + (ypx - firstTop) * k;

    const blocks = layoutRegion(items, left, right, ctx);
    const doc = { page: { w: pageW, h: pageH, mL, mR: marginR, mT: marginT, mB: 42.5 }, font: fontName, bodySize, blocks, k, L: left };
    const curW = pageW - mL - marginR;
    const c14 = calibrateWidth(blocks, fontName, false), c15 = calibrateWidth(blocks, fontName, true);
    let cw = c15;
    doc.compat = 15;
    if (c14 && (!c15 || c14.hits > c15.hits)) { cw = c14; doc.compat = 14; }
    if (cw && cw.hits >= 0.6 * cw.total && Math.abs(cw.width - curW) < 0.08 * curW) doc.page.mR = Math.max(14, pageW - mL - cw.width);
    else doc.compat = 15;
    fitParagraphs(blocks, pageW - mL - doc.page.mR, fontName, doc.compat === 15, k);
    placeVertically(doc, ctx);
    doc.debug = { chunks, rules, tables, lines, bodyEm, k, fontName, contrast, xhr, resid, boldTh, boldD, stems: regular, calib: cw, compat: doc.compat, c14, c15, joinLog: ctx.log };
    return doc;
  }

  // ---------- таблицы ----------

  function cluster(vals, tol) {
    const s = vals.slice().sort((a, b) => a - b), out = [];
    for (const v of s) {
      const last = out[out.length - 1];
      if (last && v - last.max <= tol) { last.max = v; last.sum += v; last.n++; }
      else out.push({ min: v, max: v, sum: v, n: 1 });
    }
    return out.map((c) => c.sum / c.n);
  }

  function findTables(lines, em) {
    const tol = Math.max(8, em * 0.25);
    const H = lines.h, V = lines.v;
    if (!V.length) return [];
    const parent = new Map();
    const find = (a) => { while (parent.get(a) !== a) a = parent.get(a); return a; };
    const unite = (a, b) => parent.set(find(a), find(b));
    for (const s of H.concat(V)) parent.set(s, s);
    for (const h of H) for (const v of V) {
      if (v.x >= h.x1 - tol && v.x <= h.x2 + tol && h.y >= v.y1 - tol && h.y <= v.y2 + tol) unite(h, v);
    }
    const groups = new Map();
    for (const s of H.concat(V)) {
      const r = find(s);
      if (!groups.has(r)) groups.set(r, { hs: [], vs: [] });
      (H.includes(s) ? groups.get(r).hs : groups.get(r).vs).push(s);
    }
    const tables = [];
    for (const g of groups.values()) {
      if (g.hs.length < 2 || g.vs.length < 2) continue;
      const xs = cluster(g.vs.map((v) => v.x), tol), ys = cluster(g.hs.map((h) => h.y), tol);
      if (xs.length < 2 || ys.length < 2) continue;
      const nr = ys.length - 1, nc = xs.length - 1;
      const vEdge = (r, c) => g.vs.some((v) => Math.abs(v.x - xs[c]) < tol &&
        Math.min(v.y2, ys[r + 1]) - Math.max(v.y1, ys[r]) > 0.6 * (ys[r + 1] - ys[r]));
      const hEdge = (r, c) => g.hs.some((h) => Math.abs(h.y - ys[r]) < tol &&
        Math.min(h.x2, xs[c + 1]) - Math.max(h.x1, xs[c]) > 0.6 * (xs[c + 1] - xs[c]));
      const rows = [];
      for (let r = 0; r < nr; r++) {
        const cells = [];
        let c = 0;
        while (c < nc) {
          let span = 1;
          while (c + span < nc && !vEdge(r, c + span)) span++;
          const cont = r > 0 && !hEdge(r, c);
          cells.push({ c, span, vmerge: cont ? 'continue' : null, x1: xs[c], x2: xs[c + span] });
          c += span;
        }
        rows.push({ y0: ys[r], y1: ys[r + 1], cells });
      }
      for (let r = 0; r < nr; r++) for (const cell of rows[r].cells) {
        if (cell.vmerge) continue;
        let rr = r + 1;
        while (rr < nr && rows[rr].cells.some((c2) => c2.c === cell.c && c2.span === cell.span && c2.vmerge)) rr++;
        if (rr > r + 1) cell.vmerge = 'restart';
        cell.y0 = ys[r]; cell.y1 = ys[rr];
      }
      tables.push({ xs, ys, rows, hs: g.hs, vs: g.vs, ruled: true });
    }
    return tables;
  }

  // ---------- области, колонки, строки ----------

  function bandsOf(items, gap) {
    const s = items.slice().sort((a, b) => a.y0 - b.y0), out = [];
    let cur = null, bottom = -Infinity;
    for (const it of s) {
      if (cur && it.y0 - bottom > gap) { out.push(cur); cur = null; }
      if (!cur) { cur = []; bottom = -Infinity; }
      cur.push(it); bottom = Math.max(bottom, it.y1);
    }
    if (cur) out.push(cur);
    return out;
  }

  function columnsOf(items, minGap) {
    const s = items.slice().sort((a, b) => a.x1 - b.x1), cols = [];
    let cur = null, right = -Infinity;
    for (const it of s) {
      if (cur && it.x1 - right > minGap) { cols.push(cur); cur = null; }
      if (!cur) { cur = []; right = -Infinity; }
      cur.push(it); right = Math.max(right, it.x2);
    }
    if (cur) cols.push(cur);
    return cols;
  }

  function layoutRegion(items, L, R, ctx) {
    const em = ctx.bodyEm, out = [];
    const tabs = items.filter((i) => i.kind === 'table');
    const bands = bandsOf(items.filter((i) => i.kind !== 'table'), 0.6 * em).concat(tabs.map((t) => [t]));
    bands.sort((a, b) => Math.min(...a.map((i) => i.y0)) - Math.min(...b.map((i) => i.y0)));
    for (const band of bands) {
      const tbl = band.find((i) => i.kind === 'table');
      if (tbl && band.length === 1) { out.push(tableBlock(tbl, ctx)); continue; }
      const cols = band.some((i) => i.kind === 'table') ? [band] : columnsOf(band, 2 * em);
      if (cols.length > 1 && !rowsAligned(cols, em)) {
        if (ctx.log) ctx.log.push('колонки: ' + cols.map((c) => c.map((i) => i.kind === 'text' ? i.words.map((w) => w.text).join(' ') : i.kind + '@' + Math.round(i.base)).join(' | ')).join('  ||  '));
        // независимые колонки → таблица без рамок
        const bounds = [];
        for (let c = 0; c < cols.length; c++) {
          const x1 = c === 0 ? L : Math.min(...cols[c].map((i) => i.x1));
          bounds.push(x1);
        }
        bounds.push(R);
        const cells = cols.map((col, c) => ({ x1: bounds[c], x2: bounds[c + 1], blocks: layoutRegion(col, bounds[c], bounds[c + 1], ctx) }));
        out.push({ type: 'grid', x1: L, x2: R, y0: Math.min(...band.map((i) => i.y0)), cells });
      } else {
        out.push(...flow(band, L, R, ctx));
      }
    }
    return out;
  }

  function rowsAligned(cols, em) {
    // если почти все куски правых колонок стоят на одной базовой линии с левыми — это строки с табуляцией
    let tot = 0, ok = 0;
    for (let c = 1; c < cols.length; c++) for (const it of cols[c]) {
      tot++;
      const others = cols.filter((_, j) => j !== c).flat();
      // линейка лежит чуть ниже базовой линии текста — для неё допуск шире
      if (others.some((o) => Math.abs(o.base - it.base) < (o.kind === 'rule' || it.kind === 'rule' ? 0.7 : 0.2) * em)) ok++;
    }
    return tot > 0 && ok / tot >= 0.7;
  }

  function flow(items, L, R, ctx) {
    const em = ctx.bodyEm;
    const texts = items.filter((i) => i.kind === 'text'), rules = items.filter((i) => i.kind === 'rule');
    // подчёркивания слов и строк
    const rows = [];
    for (const c of texts.sort((a, b) => a.base - b.base)) {
      const r = rows.find((r) => Math.abs(r.base - c.base) < 0.45 * c.em && r.els.every((e) => e.x2 < c.x1 || e.x1 > c.x2));
      if (r) { r.els.push(c); r.em = Math.max(r.em, c.em); }
      else rows.push({ base: c.base, em: c.em, els: [c], y0: c.y0, y1: c.y1 });
    }
    for (const r of rows) { r.y0 = Math.min(...r.els.map((e) => e.y0)); r.y1 = Math.max(...r.els.map((e) => e.y1)); }
    const freeRules = [];
    for (const s of rules) {
      // строка для линии: по высоте рядом, а по горизонтали — та, чей текст над линией или ближе всего к ней
      const near = rows.filter((r) => s.y > r.base - 0.2 * r.em && s.y < r.base + 0.6 * r.em);
      const dist = (r) => Math.min(...r.els.map((e) => Math.max(0, e.x1 - s.x2, s.x1 - e.x2)));
      const r = near.sort((a, b) => dist(a) - dist(b))[0];
      if (r && dist(r) > 3 * r.em) { freeRules.push(s); continue; }
      if (!r) { freeRules.push(s); continue; }
      const over = r.els.filter((e) => e.kind === 'text' && Math.min(e.x2, s.x2) - Math.max(e.x1, s.x1) > 0);
      if (!over.length) { r.els.push({ kind: 'blank', x1: s.x1, x2: s.x2, base: r.base, y0: s.y0, y1: s.y1 }); continue; }
      const ovLen = over.reduce((a, e) => a + Math.min(e.x2, s.x2) - Math.max(e.x1, s.x1), 0);
      if (ovLen < 0.25 * (s.x2 - s.x1)) {
        // линейка лишь задевает соседний текст (скобку, кавычку) — это пустое поле, обрезаем его по краю текста
        let x1 = s.x1, x2 = s.x2;
        for (const e of over) { if ((e.x1 + e.x2) / 2 > (s.x1 + s.x2) / 2) x2 = Math.min(x2, e.x1 - 2); else x1 = Math.max(x1, e.x2 + 2); }
        if (x2 - x1 > 0.8 * em) r.els.push({ kind: 'blank', x1, x2, base: r.base, y0: s.y0, y1: s.y1 });
        continue;
      }
      const tx1 = Math.min(...over.map((e) => e.x1)), tx2 = Math.max(...over.map((e) => e.x2));
      const words = over.flatMap((e) => e.words).filter((w) => w.x1 >= s.x1 - 0.3 * em && w.x2 <= s.x2 + 0.3 * em);
      if (s.x1 >= tx1 - 0.6 * em && s.x2 <= tx2 + 0.6 * em && words.length) { for (const w of words) w.ul = true; }
      else if (s.x2 - s.x1 > 1.3 * (tx2 - tx1)) r.border = s;
      else r.els.push({ kind: 'blank', x1: Math.max(s.x1, tx2 + 0.2 * em), x2: s.x2, base: r.base, y0: s.y0, y1: s.y1 });
    }
    for (const r of rows) {
      r.els.sort((a, b) => a.x1 - b.x1);
      // линейка вплотную к тексту — это подчёркивания в строке, а не отдельное поле с табуляцией
      if (r.els.some((e) => e.kind === 'blank') && r.els.some((e) => e.kind === 'text')) {
        const em0 = r.em, uw = advWidth('_', ctx.font, false) * em0, merged = [];
        for (const e of r.els) {
          const asWords = e.kind === 'blank'
            ? [{ text: '_'.repeat(Math.max(2, Math.round((e.x2 - e.x1) / uw))), x1: e.x1, x2: e.x2, y0: e.y0, y1: e.y1, base: r.base, bold: false, blank: true, lineEm: em0, em: em0 }]
            : e.words;
          const last = merged[merged.length - 1];
          if (last && e.x1 - last.x2 < 1.8 * em0) {
            last.words.push(...asWords); last.x2 = Math.max(last.x2, e.x2); last.y0 = Math.min(last.y0, e.y0); last.y1 = Math.max(last.y1, e.y1);
            if (e.kind === 'text') { last.kind = 'text'; last.em = e.em; }
          } else merged.push(e.kind === 'blank' ? { ...e, words: asWords, em: em0 } : { ...e, words: e.words.slice() });
        }
        r.els = merged.map((m) => (m.kind === 'blank' && m.words.length > 1 ? { ...m, kind: 'text' } : m));
      }
    }
    const seq = rows.map((r) => ({ type: 'row', ...r })).concat(freeRules.map((s) => ({ type: 'rule', base: s.y, y0: s.y0, y1: s.y1, s })));
    seq.sort((a, b) => a.base - b.base);

    // центр колонки: если строки дружно центрованы вокруг одной оси
    const singles = rows.filter((r) => r.els.length === 1 && !r.border);
    let colCenter = null;
    if (singles.length >= 3) {
      const cs = singles.map((r) => (r.els[0].x1 + r.els[0].x2) / 2), ls = singles.map((r) => r.els[0].x1);
      const mc = median(cs);
      const spreadC = median(cs.map((c) => Math.abs(c - mc))), spreadL = median(ls.map((l) => Math.abs(l - median(ls))));
      if (spreadC < 0.5 * em && spreadL > 1.0 * em) colCenter = mc;
    }

    const paras = [];
    for (const it of seq) {
      if (it.type === 'rule') { paras.push({ kind: 'rule', s: it.s, base: it.s.y }); continue; }
      const single = it.els.length === 1 && it.els[0].kind === 'text' && !it.border;
      const prev = paras[paras.length - 1];
      if (single && prev && prev.kind === 'text' && canJoin(prev, it, L, R, colCenter, ctx)) { prev.lines.push(it); continue; }
      paras.push(single ? { kind: 'text', lines: [it] } : { kind: 'tabs', row: it });
    }
    return paras.map((p) => buildPara(p, L, R, colCenter, ctx));
  }

  function lineCenter(r) { return (r.els[0].x1 + r.els[r.els.length - 1].x2) / 2; }

  function canJoin(p, r, L, R, colCenter, ctx) {
    const last = p.lines[p.lines.length - 1], e = r.els[0], le = last.els[0];
    const em = Math.max(e.em, le.em);
    const no = (why) => {
      if (ctx.log) ctx.log.push(why + ': «' + le.words.slice(-2).map((w) => w.text).join(' ') + '» | «' + e.words.slice(0, 2).map((w) => w.text).join(' ') + '»');
      return false;
    };
    if (Math.abs(e.em - le.em) > 0.06 * em) return no('кегль ' + le.em.toFixed(1) + '/' + e.em.toFixed(1));
    const pitch = r.base - last.base;
    if (pitch < 0.95 * em || pitch > 2.1 * em) return no('шаг ' + (pitch / em).toFixed(2));
    if (p.lines.length >= 2) {
      const pp = last.base - p.lines[p.lines.length - 2].base;
      if (Math.abs(pitch - pp) > 0.35 * em) return no('неровный шаг ' + ((pitch - pp) / em).toFixed(2));
    }
    // центрованные строки не склеиваются: ширину, по которой их переносили, не узнать
    const centered = colCenter !== null && Math.abs((le.x1 + le.x2) / 2 - colCenter) < 0.6 * em && Math.abs((e.x1 + e.x2) / 2 - colCenter) < 0.6 * em;
    if (centered) return no('по центру');
    if (p.lines.length === 1 && Math.abs((le.x1 + le.x2) / 2 - (e.x1 + e.x2) / 2) < 0.6 * em && Math.abs(le.x1 - e.x1) > 1 * em) return no('пара по центру');
    const firstWord = e.words[0], wordW = firstWord.x2 - firstWord.x1;
    if (R - le.x2 > wordW + 0.6 * em) return no('слово влезло бы'); // значит, строку оборвали намеренно
    if (p.lines.length >= 2) return Math.abs(e.x1 - p.lines[1].els[0].x1) < 0.7 * em || no('левый край');
    const d = p.lines[0].els[0].x1 - e.x1;
    return (d > -0.5 * em && d < 6 * em) || no('отступ');
  }

  function runsOf(words, sizeOf, em) {
    const runs = [];
    for (let i = 0; i < words.length; i++) {
      const w = words[i];
      const st = { b: !!w.bold, i: !!w.italic, u: !!w.ul, size: sizeOf(w.lineEm || em), hl: !w.blank && w.conf < 70 };
      const last = runs[runs.length - 1];
      let space = i > 0 ? ' ' : '';
      if (last && last.b === st.b && last.i === st.i && last.u === st.u && last.size === st.size && last.hl === st.hl) { last.text += space + w.text; continue; }
      // пробел между ранами разного вида — в неподчёркнутый ран
      if (space && last && !last.u) { last.text += space; space = ''; }
      runs.push({ ...st, text: space + w.text });
    }
    return runs;
  }

  function buildPara(p, L, R, colCenter, ctx) {
    const { k, bodyEm, sizeOf } = ctx, em = bodyEm, pt = (px) => px * k;
    if (p.kind === 'rule') {
      return { type: 'p', rule: true, indL: pt(Math.max(0, p.s.x1 - L)), indR: pt(Math.max(0, R - p.s.x2)), lineY: p.s.y, runs: [], size: ctx.bodySize };
    }
    if (p.kind === 'tabs') {
      const r = p.row, tabs = [], runs = [];
      let indL = 0, indR = 0, cursor = L;
      if (r.border) { indL = pt(Math.max(0, r.border.x1 - L)); indR = pt(Math.max(0, R - r.border.x2)); cursor = r.border.x1; }
      for (let i = 0; i < r.els.length; i++) {
        const e = r.els[i];
        const isLast = i === r.els.length - 1;
        if (e.kind === 'blank') {
          if (e.x1 - cursor > 0.4 * em) { tabs.push({ pos: pt(e.x1 - L), val: 'left' }); runs.push({ tab: true }); }
          tabs.push({ pos: pt(e.x2 - L), val: 'left' }); runs.push({ tab: true, u: true });
          cursor = e.x2; continue;
        }
        const gap = e.x1 - cursor;
        let start = null;
        if (isLast && i > 0 && R - e.x2 < 0.8 * em && gap > 2 * em) { tabs.push({ pos: pt(R - L), val: 'right' }); runs.push({ tab: true }); }
        else if (i === 0 && !r.border) { if (gap > 0.3 * em) indL = pt(gap); start = indL; }
        else if (gap > 0.3 * em || i > 0) { tabs.push({ pos: pt(e.x1 - L), val: 'left' }); runs.push({ tab: true }); start = pt(e.x1 - L); }
        else start = pt(cursor - L);
        const er = runsOf(e.words, sizeOf, e.em);
        runs.push(...er);
        // кусок не должен упираться в правый край: иначе Word перенесёт его на вторую строку
        if (start !== null) {
          const wPt = er.reduce((a, rr) => a + advWidth(rr.text, ctx.font, rr.b) * rr.size, 0);
          const room = pt(R - L) - indR - start;
          if (wPt + 1.5 > room) indR = Math.max(0, indR - (wPt + 1.5 - room));
        }
        cursor = e.x2;
      }
      return {
        type: 'p', align: 'left', indL, indR, firstLine: 0, tabs, runs,
        base: r.base, nLines: 1, size: sizeOf(r.em), border: r.border ? { y: r.border.y } : null,
      };
    }
    const lines = p.lines, n = lines.length, els = lines.map((l) => l.els[0]);
    const size = sizeOf(median(els.map((e) => e.em)));
    let align = 'left', indL = 0, indR = 0, first = 0;
    const x1s = els.map((e) => e.x1), x2s = els.map((e) => e.x2);
    const centerOf = (e) => (e.x1 + e.x2) / 2;
    const setCenter = (c) => {
      align = 'center';
      const mid = (L + R) / 2;
      if (c < mid) indR = pt(R - (2 * c - L)); else indL = pt(2 * c - R - L);
    };
    if (n >= 2) {
      const cont = median(x1s.slice(1)), nonLast = x2s.slice(0, -1);
      const cs = els.map(centerOf), c0 = median(cs);
      const toRight = x2s.slice(0, -1).every((x) => R - x < 0.6 * em);
      const centeredAll = !toRight && cs.every((c) => Math.abs(c - c0) < 0.6 * em) && Math.max(...x1s) - Math.min(...x1s) > 1.2 * em;
      if (centeredAll) setCenter(c0);
      else {
        indL = pt(Math.max(0, cont - L));
        first = pt(x1s[0] - cont);
        const rr = Math.max(...nonLast);
        indR = pt(Math.max(0, R - rr));
        const justified = nonLast.every((x) => rr - x < 0.5 * em);
        align = justified ? 'both' : 'left';
        if (!justified) indR = 0;
        if (indL < pt(0.3 * em)) indL = 0;
        if (Math.abs(first) < pt(0.3 * em)) first = 0;
        if (indR < pt(0.4 * em)) indR = 0;
      }
    } else {
      const e = els[0], dl = e.x1 - L, dr = R - e.x2, c = centerOf(e), mid = (L + R) / 2;
      if (colCenter !== null && Math.abs(c - colCenter) < 0.6 * em && dl > 0.5 * em) setCenter(colCenter);
      else if (dl < 0.5 * em) { align = 'left'; }
      else if (dr < 0.6 * em && dl > 3 * em) { align = 'right'; indR = pt(dr); }
      else if (Math.abs(c - mid) < 1.0 * em) setCenter(c);
      else { align = 'left'; indL = pt(dl); }
    }
    // текст: строки склеиваются; перенос по слогам убирается
    const words = [];
    lines.forEach((l, i) => {
      const ws = l.els[0].words.map((w) => ({ ...w }));
      const prevW = words[words.length - 1];
      if (i > 0 && prevW && /[-­‐]$/.test(prevW.text) && prevW.text.length > 1 && /^[а-яёa-z]/.test(ws[0].text)) {
        prevW.text = prevW.text.replace(/[-­‐]$/, '') + ws[0].text;
        prevW.x2 = ws[0].x2;
        ws.shift();
      }
      words.push(...ws);
    });
    const runs = runsOf(words, sizeOf, els[0].em);
    const pitches = [];
    for (let i = 1; i < n; i++) pitches.push(lines[i].base - lines[i - 1].base);
    return {
      type: 'p', align, indL, indR, firstLine: first, tabs: [], runs, size,
      base: lines[0].base, nLines: n, pitch: n >= 2 ? median(pitches) : null,
      breaks: lines.map((l) => l.els[0].words.length), lineWords: lines.map((l) => l.els[0].words),
    };
  }

  function tableBlock(t, ctx) {
    const em = ctx.bodyEm, pad = Math.max(3, 0.15 * em);
    const tem = median(t.words.map((w) => w.em));
    // в ячейках слова короткие и оценка кегля шумная: текст таблицы приводим к общему кеглю
    for (const w of t.words) if (w.em / tem < 2.2 && w.em / tem > 0.5) { w.em = tem; w.lineEm = tem; }
    const rows = t.rows.map((row) => ({
      y0: row.y0, y1: row.y1,
      cells: row.cells.map((cell) => {
        let blocks = [];
        if (cell.vmerge !== 'continue') {
          const ws = t.words.filter((w) => {
            const cx = (w.x1 + w.x2) / 2, cy = (w.y0 + w.y1) / 2;
            return cx > cell.x1 && cx < cell.x2 && cy > cell.y0 && cy < cell.y1;
          });
          const lineMap = [];
          for (const w of ws.sort((a, b) => a.base - b.base)) {
            let L = lineMap.find((l) => Math.abs(l.base - w.base) < 0.45 * w.em);
            if (!L) { L = { base: w.base, ws: [], em: w.em }; lineMap.push(L); }
            L.ws.push(w);
          }
          const items = [];
          for (const l of lineMap) {
            l.ws.sort((a, b) => a.x1 - b.x1);
            let cur = null;
            for (const w of l.ws) {
              if (cur && w.x1 - cur.x2 > 1.8 * l.em) { items.push(cur); cur = null; }
              if (!cur) cur = { kind: 'text', words: [], x1: w.x1, x2: w.x2, y0: w.y0, y1: w.y1, em: l.em, base: l.base };
              cur.words.push(w); cur.x2 = Math.max(cur.x2, w.x2); cur.y0 = Math.min(cur.y0, w.y0); cur.y1 = Math.max(cur.y1, w.y1);
            }
            if (cur) items.push(cur);
          }
          for (const r of t.rules || []) {
            const cx = (r.x1 + r.x2) / 2;
            if (cx > cell.x1 && cx < cell.x2 && r.y > cell.y0 && r.y < cell.y1) items.push(r);
          }
          blocks = items.length ? flow(items, cell.x1 + pad, cell.x2 - pad, ctx) : [];
        }
        return { ...cell, blocks };
      }),
    }));
    return { type: 'table', ruled: true, xs: t.xs, ys: t.ys, y0: t.y0, rows, pad };
  }

  // ---------- вертикальная раскладка (пункты) ----------
  // Каждый абзац ставится на свою базовую линию: интервал перед абзацем = цель − текущее положение.
  function placeVertically(doc, ctx) {
    const { k, yPt } = ctx;
    const placeList = (blocks, startY) => {
      let cur = startY;
      for (const b of blocks) {
        if (b.type === 'p') {
          if (b.rule) {
            b.line = 2; // пт, точно
            const top = yPt(b.lineY) - b.line - 1;
            b.before = Math.max(0, top - cur);
            b.border = { space: 0 };
            cur = cur + b.before + b.line + 1;
            continue;
          }
          b.line = b.pitch ? b.pitch * k : Math.round(b.size * 1.15 * 2) / 2;
          const top = yPt(b.base) - BETA * b.line;
          b.before = Math.max(0, top - cur);
          cur = cur + b.before + b.nLines * b.line;
          if (b.border) {
            const want = yPt(b.border.y) - (cur);
            b.border.space = clamp(want, 0, 31);
            cur += b.border.space + 1;
          }
        } else if (b.type === 'grid') {
          let bottom = cur;
          for (const c of b.cells) bottom = Math.max(bottom, placeList(c.blocks, cur));
          b.top = cur;
          cur = bottom;
        } else if (b.type === 'table') {
          const top = yPt(b.y0);
          b.gapBefore = Math.max(0, top - cur);
          cur = cur + b.gapBefore;
          for (const row of b.rows) {
            row.h = (row.y1 - row.y0) * k;
            for (const c of row.cells) placeList(c.blocks, yPt(row.y0) + 1);
          }
          cur = yPt(b.ys[b.ys.length - 1]);
        }
      }
      return cur;
    };
    const end = placeList(doc.blocks, doc.page.mT);
    const limit = doc.page.h - 34;
    if (end > limit) { // не вылезти на вторую страницу: ужимаем интервалы
      const all = [];
      const walk = (bs) => { for (const b of bs) { if (b.type === 'p') all.push(b); if (b.cells) b.cells.forEach((c) => walk(c.blocks)); if (b.rows) b.rows.forEach((r) => r.cells.forEach((c) => walk(c.blocks))); } };
      walk(doc.blocks);
      const sum = all.reduce((s, b) => s + (b.before || 0), 0);
      const f = sum > 0 ? Math.max(0, (sum - (end - limit)) / sum) : 1;
      for (const b of all) b.before = (b.before || 0) * f;
    }
    doc.page.mB = clamp(doc.page.h - Math.min(end, limit) - 16, 14, 56.7);
  }

  // ---------- docx ----------

  const esc = (s) => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const tw = (pt) => Math.round(pt * 20);

  function runXml(r, opts) {
    const rp = [];
    if (r.b) rp.push('<w:b/><w:bCs/>');
    if (r.i) rp.push('<w:i/><w:iCs/>');
    if (r.size) rp.push(`<w:sz w:val="${Math.round(r.size * 2)}"/><w:szCs w:val="${Math.round(r.size * 2)}"/>`);
    if (r.hl && opts && opts.highlight) rp.push('<w:highlight w:val="yellow"/>');
    if (r.u) rp.push('<w:u w:val="single"/>');
    const rpr = rp.length ? `<w:rPr>${rp.join('')}</w:rPr>` : '';
    if (r.tab) return `<w:r>${rpr}<w:tab/></w:r>`;
    if (r.br) return `<w:r>${rpr}<w:br/></w:r>`;
    return `<w:r>${rpr}<w:t xml:space="preserve">${esc(r.text)}</w:t></w:r>`;
  }

  function paraXml(p, opts) {
    const pp = [];
    if (p.border) pp.push(`<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="${Math.round(clamp(p.border.space || 0, 0, 31))}" w:color="000000"/></w:pBdr>`);
    if (p.tabs && p.tabs.length) {
      const seen = new Set();
      const ts = p.tabs.filter((t) => { const key = tw(t.pos); if (seen.has(key) || key <= 0) return false; seen.add(key); return true; })
        .sort((a, b) => a.pos - b.pos).map((t) => `<w:tab w:val="${t.val}" w:pos="${tw(t.pos)}"/>`);
      if (ts.length) pp.push(`<w:tabs>${ts.join('')}</w:tabs>`);
    }
    pp.push(`<w:spacing w:before="${tw(p.before || 0)}" w:after="0" w:line="${Math.max(12, tw(p.line || 12))}" w:lineRule="exact"/>`);
    const ind = [];
    if (p.indL) ind.push(`w:left="${tw(p.indL)}"`);
    if (p.indR) ind.push(`w:right="${tw(p.indR)}"`);
    if (p.firstLine > 0) ind.push(`w:firstLine="${tw(p.firstLine)}"`);
    if (p.firstLine < 0) ind.push(`w:hanging="${tw(-p.firstLine)}"`);
    if (ind.length) pp.push(`<w:ind ${ind.join(' ')}/>`);
    if (p.align && p.align !== 'left') pp.push(`<w:jc w:val="${p.align}"/>`);
    pp.push(`<w:rPr><w:sz w:val="${Math.round((p.size || 12) * 2)}"/></w:rPr>`);
    let runs = p.runs || [];
    if (opts.keepBreaks && p.breaks && p.breaks.length > 1) runs = withBreaks(runs, p.breaks);
    return `<w:p><w:pPr>${pp.join('')}</w:pPr>${runs.map((r) => runXml(r, opts)).join('')}</w:p>`;
  }

  // Переносы как в оригинале: разрезаем раны по числу слов в строках и ставим разрыв строки.
  function withBreaks(runs, breaks) {
    const out = [];
    let lineIdx = 0, left = breaks[0];
    for (const r of runs) {
      if (r.tab) { out.push(r); continue; }
      const parts = r.text.split(' ');
      let buf = [];
      const lead = parts[0] === '' ? (parts.shift(), ' ') : '';
      let leadUsed = false;
      for (const w of parts) {
        if (w === '') continue;
        buf.push(w); left--;
        if (left === 0 && lineIdx < breaks.length - 1) {
          out.push({ ...r, text: (leadUsed ? '' : lead) + buf.join(' ') }); leadUsed = true;
          out.push({ br: true }); buf = []; lineIdx++; left = breaks[lineIdx];
        }
      }
      if (buf.length) { out.push({ ...r, text: (leadUsed ? ' ' : lead) + buf.join(' ') }); }
      else if (!leadUsed && lead && out.length && !out[out.length - 1].br) out.push({ ...r, text: lead });
    }
    // пробел в начале строки после разрыва не нужен
    for (let i = 1; i < out.length; i++) if (out[i - 1].br && out[i].text) out[i].text = out[i].text.replace(/^ /, '');
    return out;
  }

  function blocksXml(blocks, opts, widthPt) {
    const parts = [];
    let prevTable = false;
    for (const b of blocks) {
      if (b.type === 'p') { parts.push(paraXml(b, opts)); prevTable = false; continue; }
      if (prevTable) parts.push(paraXml({ line: 1, runs: [], size: 1 }, opts));
      if (b.type === 'grid') {
        const cols = gridCols(b, widthPt, opts.k).map(tw);
        parts.push(`<w:tbl><w:tblPr><w:tblW w:w="${cols.reduce((s, v) => s + v, 0)}" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/>` +
          '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>' +
          '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>' +
          '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>' +
          `<w:tblGrid>${cols.map((c) => `<w:gridCol w:w="${c}"/>`).join('')}</w:tblGrid><w:tr>` +
          b.cells.map((c, i) => `<w:tc><w:tcPr><w:tcW w:w="${cols[i]}" w:type="dxa"/></w:tcPr>${c.blocks.length ? blocksXml(c.blocks, opts, cols[i] / 20) : ''}${!c.blocks.length || c.blocks[c.blocks.length - 1].type !== 'p' ? paraXml({ line: 1, runs: [], size: 1 }, opts) : ''}</w:tc>`).join('') +
          '</w:tr></w:tbl>');
      } else if (b.type === 'table') {
        if (b.gapBefore > 1) parts.push(paraXml({ line: Math.min(1584, b.gapBefore), runs: [], size: 1 }, opts));
        const cols = [];
        for (let i = 0; i + 1 < b.xs.length; i++) cols.push(tw((b.xs[i + 1] - b.xs[i]) * opts.k));
        const indent = tw((b.xs[0] - opts.L) * opts.k);
        const padTw = tw(b.pad * opts.k);
        const bd = (n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`;
        let x = `<w:tbl><w:tblPr><w:tblW w:w="${cols.reduce((s, v) => s + v, 0)}" w:type="dxa"/><w:tblInd w:w="${indent}" w:type="dxa"/>` +
          `<w:tblBorders>${bd('top')}${bd('left')}${bd('bottom')}${bd('right')}${bd('insideH')}${bd('insideV')}</w:tblBorders>` +
          `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="${padTw}" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="${padTw}" w:type="dxa"/></w:tblCellMar>` +
          '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>' +
          `<w:tblGrid>${cols.map((c) => `<w:gridCol w:w="${c}"/>`).join('')}</w:tblGrid>`;
        for (const row of b.rows) {
          x += `<w:tr><w:trPr><w:trHeight w:val="${tw(row.h)}" w:hRule="atLeast"/></w:trPr>`;
          for (const c of row.cells) {
            const wdt = cols.slice(c.c, c.c + c.span).reduce((s, v) => s + v, 0);
            x += `<w:tc><w:tcPr><w:tcW w:w="${wdt}" w:type="dxa"/>${c.span > 1 ? `<w:gridSpan w:val="${c.span}"/>` : ''}` +
              `${c.vmerge === 'restart' ? '<w:vMerge w:val="restart"/>' : c.vmerge === 'continue' ? '<w:vMerge/>' : ''}</w:tcPr>`;
            const inner = c.blocks.length ? blocksXml(c.blocks, opts, wdt / 20) : '';
            x += inner && c.blocks[c.blocks.length - 1].type === 'p' ? inner : inner + paraXml({ line: 12, runs: [], size: 10 }, opts);
            x += '</w:tc>';
          }
          x += '</w:tr>';
        }
        parts.push(x + '</w:tbl>');
      }
      prevTable = true;
    }
    return parts.join('');
  }

  function documentXml(doc, opts) {
    const p = doc.page;
    const body = blocksXml(doc.blocks, { ...opts, k: doc.k, L: doc.L }, p.w - p.mL - p.mR);
    const orient = p.w > p.h ? ' w:orient="landscape"' : '';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<w:body>${body}<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>` +
      `<w:sectPr><w:pgSz w:w="${tw(p.w)}" w:h="${tw(p.h)}"${orient}/>` +
      `<w:pgMar w:top="${tw(p.mT)}" w:right="${tw(p.mR)}" w:bottom="${tw(p.mB)}" w:left="${tw(p.mL)}" w:header="567" w:footer="567" w:gutter="0"/>` +
      '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>';
  }

  function stylesXml(doc) {
    const f = esc(doc.font), sz = Math.round(doc.bodySize * 2);
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>` +
      `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/><w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>` +
      '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:rPr><w:lang w:val="ru-RU"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:default="1" w:styleId="a0"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>' +
      '<w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>' +
      '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
      '</w:styles>';
  }

  const SETTINGS = (compat) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/>' +
    '<w:defaultTabStop w:val="709"/><w:characterSpacingControl w:val="doNotCompress"/>' +
    '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="' + (compat || 15) + '"/></w:compat>' +
    '<w:themeFontLang w:val="ru-RU"/></w:settings>';

  function buildDocx(doc, opts) {
    opts = opts || {};
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const files = [
      ['[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${opts.template ? 'template' : 'document'}.main+xml"/>` +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'],
      ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'],
      ['word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'],
      ['word/document.xml', documentXml(doc, opts)],
      ['word/styles.xml', stylesXml(doc)],
      ['word/settings.xml', SETTINGS(opts.compat || doc.compat)],
      ['docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        `<dc:title>${esc(opts.title || 'Документ')}</dc:title><dc:language>ru-RU</dc:language>` +
        `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`],
      ['docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application></Properties>'],
    ];
    return zip(files);
  }

  // ---------- zip без сжатия ----------
  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(b) { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
  function utf8(s) {
    if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(s);
    return Uint8Array.from(Buffer.from(s, 'utf8'));
  }
  function zip(files) {
    const chunks = [], central = [];
    let offset = 0;
    const d = new Date(), dt = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    const tm = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    for (const [name, content] of files) {
      const nb = utf8(name), data = typeof content === 'string' ? utf8(content) : content, crc = crc32(data);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
      h.setUint16(10, tm, true); h.setUint16(12, dt, true); h.setUint32(14, crc, true);
      h.setUint32(18, data.length, true); h.setUint32(22, data.length, true); h.setUint16(26, nb.length, true); h.setUint16(28, 0, true);
      chunks.push(new Uint8Array(h.buffer), nb, data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 0, true);
      c.setUint16(12, tm, true); c.setUint16(14, dt, true); c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true);
      c.setUint16(28, nb.length, true); c.setUint32(42, offset, true);
      central.push(new Uint8Array(c.buffer), nb);
      offset += 30 + nb.length + data.length;
    }
    const cdSize = central.reduce((s, a) => s + a.length, 0);
    const e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, cdSize, true); e.setUint32(16, offset, true);
    const all = chunks.concat(central, [new Uint8Array(e.buffer)]);
    const out = new Uint8Array(all.reduce((s, a) => s + a.length, 0));
    let p = 0;
    for (const a of all) { out.set(a, p); p += a.length; }
    return out;
  }

  // Полный разбор после распознавания: модель документа + docx.
  function toDocx(page, lines, ocr, opts) {
    const doc = analyze(page, lines, ocr, opts);
    if (doc.empty) return { doc, bytes: null };
    return { doc, bytes: buildDocx(doc, opts) };
  }

  return { gridCols, detectQuad, planPage, warp, cleanPage, detectLines, removeBlobs, ocrImage, analyze, buildDocx, toDocx, fullQuad, A4 };
});
