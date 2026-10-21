/* Заставка «Магазин Чорсу» — плоская 2D-векторная анимация.
   Детерминированный таймлайн: render(t) рисует кадр для времени t (секунды).
   Используется и для экспорта видео, и как заставка в приложении. */
(function (root) {
  const C = { bg: '#1D5C45', dark: '#12382B', light: '#F2F5F0', amber: '#E8B05A', coral: '#D9573B' };
  const DUR = 2.7;
  const NS = 'http://www.w3.org/2000/svg';
  // --- easing ---
  const clamp = (x, a = 0, b = 1) => Math.min(b, Math.max(a, x));
  const seg = (t, a, b) => clamp((t - a) / (b - a));
  const io = x => x < .5 ? 4 * x * x * x : 1 - Math.pow(-2 * x + 2, 3) / 2;       // ease-in-out cubic
  const io5 = x => x < .5 ? 16 * Math.pow(x, 5) : 1 - Math.pow(-2 * x + 2, 5) / 2; // stronger ease-in-out
  const out = x => 1 - Math.pow(1 - x, 3);
  const back = (x, s = 1.4) => 1 + (s + 1) * Math.pow(x - 1, 3) + s * Math.pow(x - 1, 2);
  const lerp = (a, b, k) => a + (b - a) * k;
  const mix = (c1, c2, k) => { const p = c => [1, 3, 5].map(i => parseInt(c.slice(i, i + 2), 16)); const a = p(c1), b = p(c2); return '#' + a.map((v, i) => Math.round(lerp(v, b[i], k)).toString(16).padStart(2, '0')).join(''); };
  const el = (tag, attrs, parent) => { const e = document.createElementNS(NS, tag); for (const k in attrs) e.setAttribute(k, attrs[k]); if (parent) parent.appendChild(e); return e; };
  const set = (e, attrs) => { for (const k in attrs) e.setAttribute(k, attrs[k]); };

  // --- layout in a 1000×1000 logical square ---
  const CX = 500;
  const COUNTER_Y = 660, LEFT = 262, RIGHT = 738;          // прилавок
  const TOP = 312;                                         // верх ставни / навеса
  const SHELF1 = 486, SHELF2 = 632;                        // полки
  const PW = 74, PH = 104;                                 // пачка
  const packX = i => 300 + (i % 4) * 112;                  // 4 в ряд
  const packs = [0, 1, 2, 3, 4, 5, 6, 7].map(i => ({
    i, x: packX(i), floor: i < 4 ? SHELF1 : SHELF2,
    body: [C.amber, C.coral, C.light, C.amber, C.coral, C.light, C.amber, C.coral][i],
    band: [C.light, C.light, C.coral, C.light, C.light, C.amber, C.light, C.light][i],
    drop: 0.98 + [0, 1, 2, 3, 4, 5, 6, 7][i] * 0.07
  }));
  // буква «Ч» из трёх планок
  const CH = { L: [418, 438, 46, 118], M: [418, 520, 164, 46], R: [536, 438, 46, 226] };
  const target = [CH.L, CH.M, CH.M, CH.R, CH.L, CH.M, CH.R, CH.R];
  const FINAL_AWN = { x: 384, y: 352, w: 232 };            // навес над буквой
  const WORD = 'Магазин Чорсу', WORD_Y = 772;

  function mount(svg, opts = {}) {
    const font = opts.font || "'Golos Text', 'Segoe UI', Roboto, Arial, sans-serif";
    svg.innerHTML = '';
    const defs = el('defs', {}, svg);
    const clip = el('clipPath', { id: 'csWordClip' }, defs); el('rect', { x: -500, y: WORD_Y - 80, width: 2000, height: 96 }, clip);
    const kclip = el('clipPath', { id: 'csKiosk' }, defs); el('rect', { x: LEFT, y: TOP + 26, width: RIGHT - LEFT, height: 420 }, kclip);
    el('rect', { x: -2000, y: -2000, width: 5000, height: 5000, fill: C.bg }, svg);
    const cam = el('g', {}, svg);
    // прилавок
    const counter = el('line', { x1: CX, y1: COUNTER_Y, x2: CX, y2: COUNTER_Y, stroke: C.light, 'stroke-width': 10, 'stroke-linecap': 'round' }, cam);
    // полки
    const shelves = [SHELF1, SHELF2].map(y => el('rect', { x: LEFT + 14, y, width: RIGHT - LEFT - 28, height: 9, rx: 4.5, fill: C.dark }, cam));
    const posts = [LEFT + 14, RIGHT - 23].map(x => el('rect', { x, y: TOP + 40, width: 9, height: COUNTER_Y - TOP - 40, rx: 4.5, fill: C.dark }, cam));
    // пачки
    const packG = el('g', { 'clip-path': 'url(#csKiosk)' }, cam);
    const packEls = packs.map(p => {
      const g = el('g', {}, packG);
      const body = el('rect', { rx: 7, fill: p.body }, g);
      const band = el('rect', { fill: p.band }, g);
      const shade = el('rect', { fill: C.dark, opacity: .14 }, g);
      return { g, body, band, shade };
    });
    // ставня (вырастает из линии прилавка) и рулон/планка навеса
    const shutter = el('rect', { x: LEFT, width: RIGHT - LEFT, fill: C.light }, cam);
    const slats = [0, 1, 2, 3, 4, 5].map(() => el('rect', { x: LEFT + 18, width: RIGHT - LEFT - 36, height: 6, rx: 3, fill: C.bg, opacity: .16 }, cam));
    const bar = el('rect', { rx: 11, fill: C.light }, cam);
    // фестоны навеса: 6 полукругов чередующихся цветов
    const scal = [0, 1, 2, 3, 4, 5].map(k => el('path', { fill: k % 2 ? C.amber : C.light }, cam));
    // буква «Ч»: финальные планки и скругление угла
    function measure(){}               // надписи больше нет — оставлено для совместимости

    function scallops(x, y, w, n, drop) { // drop[k] 0..1 — насколько опущен каждый фестон
      const sw = w / n;
      scal.forEach((s, k) => {
        const d = drop[k], r = sw / 2, h = r * 1.0 * d;
        const x0 = x + k * sw;
        s.setAttribute('d', `M${x0} ${y} h${sw} v${h * 0.15} a${r} ${Math.max(h, .01)} 0 0 1 ${-sw} 0 Z`);
        s.setAttribute('opacity', d > 0.01 ? 1 : 0);
      });
    }

    function render(t) {
      t = clamp(t, 0, DUR);
      // камера: наезд 1.12→1.00, откат на сборке, посадка
      const push = io(seg(t, 0, 2.1));
      const s = lerp(1.14, 1.0, push);
      const cy = lerp(540, 486, io(seg(t, 0.2, 2.1)));
      set(cam, { transform: `translate(${CX} 500) scale(${s}) translate(${-CX} ${-cy})` });

      // 1) прилавок рисуется из центра
      const draw = io(seg(t, 0.0, 0.38));
      const collapse = 0;                                         // магазин остаётся до конца
      const halfW = lerp(0, (RIGHT - LEFT) / 2, draw) * (1 - collapse);
      set(counter, { x1: CX - halfW, x2: CX + halfW, opacity: 1 });

      // ставня вырастает из линии (морф: линия → прямоугольник), затем сворачивается вверх
      const grow = io(seg(t, 0.28, 0.72)), roll = io(seg(t, 0.78, 1.12));
      const shTop = lerp(COUNTER_Y, TOP, grow);
      const shBottom = lerp(COUNTER_Y, TOP + 24, roll);
      const shH = Math.max(shBottom - shTop, 0);
      set(shutter, { y: shTop, height: shH, opacity: (grow > 0 && roll < 1) ? 1 : 0 });
      slats.forEach((sl, k) => { const y = shTop + 36 + k * 52; set(sl, { y, opacity: (y < shBottom - 12 && grow > .6) ? .16 : 0 }); });

      // рулон → планка навеса → навес над буквой
      const toFinal = 0;
      const barShow = roll > 0.92 || t > 1.1;
      const bx = lerp(LEFT - 10, FINAL_AWN.x, toFinal), bw = lerp(RIGHT - LEFT + 20, FINAL_AWN.w, toFinal);
      const by = lerp(TOP - 6, FINAL_AWN.y, toFinal), bh = lerp(26, 22, toFinal);
      set(bar, { x: bx, y: by, width: bw, height: bh, opacity: barShow ? 1 : 0 });
      const drops = scal.map((_, k) => back(seg(t, 1.12 + k * 0.07, 1.42 + k * 0.07), 1.2));
      scallops(bx, by + bh - 1, bw, 6, drops);

      // полки и стойки появляются, когда ставня уходит, и тают при сборке
      const shelfIn = io(seg(t, 0.82, 1.1)), shelfOut = 0;
      shelves.forEach((sh, k) => { const w = (RIGHT - LEFT - 28) * shelfIn * (1 - shelfOut * .6); set(sh, { x: CX - w / 2, width: w, opacity: 1 - shelfOut }); });
      posts.forEach(p => set(p, { opacity: shelfIn * (1 - shelfOut) }));

      // 2) пачки: падают по очереди с приседанием, 3) сливаются в «Ч»
      const merge = 0;
      packs.forEach((p, k) => {
        const e = packEls[k];
        const fall = seg(t, p.drop, p.drop + 0.26);
        const land = seg(t, p.drop + 0.26, p.drop + 0.48);                  // касание и отскок
        let sx = 1, sy = 1, y;
        if (fall < 1){ sy = 1 + 0.12 * fall; sx = 1 - 0.06 * fall; y = lerp(p.floor - PH - 140, p.floor - PH * sy, Math.pow(fall, 2.2)); }
        else { const sq = Math.sin(Math.PI * land) * (1 - land); sy = 1 - 0.22 * sq; sx = 1 + 0.16 * sq; y = p.floor - PH * sy; }
        const w = PW * sx, h = PH * sy, x = p.x + (PW - w) / 2;
        // слияние в планку буквы
        const tg = target[k], m = merge;
        const fx = lerp(x, tg[0], m), fy = lerp(y, tg[1], m), fw = lerp(w, tg[2], m), fh = lerp(h, tg[3], m);
        const vis = t >= p.drop ? 1 : 0;
        set(e.body, { x: fx, y: fy, width: fw, height: fh, rx: lerp(7, 4, m), fill: p.body });
        const bandH = fh * 0.24 * (1 - m);
        set(e.band, { x: fx, y: fy + fh * 0.16 * (1 - m), width: fw, height: bandH, opacity: 1 - m });
        set(e.shade, { x: fx + fw * 0.68, y: fy, width: fw * 0.32 * (1 - m), height: fh, opacity: .14 * (1 - m) });
        set(e.g, { opacity: vis });
      });
      // скругление внутреннего угла «Ч» проявляется после слияния
      set(packG, { 'clip-path': 'url(#csKiosk)' });
    }
    render(0);
    return { render, measure, DUR };
  }
  root.ChorsuSplash = { mount, DUR, colors: C };
})(typeof window !== 'undefined' ? window : this);
