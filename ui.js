/* Utilidades de interfaz compartidas · SUGERIDO DE VENTAS · Venejugos 2015 */
(function () {
  'use strict';
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();

  /* Avisos tipo Sonner */
  function aviso(titulo, texto = '', tipo = 'ok', opciones = {}) {
    let cont = $('.avisos');
    if (!cont) { cont = document.createElement('div'); cont.className = 'avisos'; cont.setAttribute('role', 'status'); cont.setAttribute('aria-live', 'polite'); document.body.appendChild(cont); }
    const el = document.createElement('div');
    el.className = 'aviso ' + tipo;
    const ic = tipo === 'error' ? '!' : tipo === 'info' ? 'i' : '✓';
    el.innerHTML = `<span class="ic" aria-hidden="true">${ic}</span><div><b>${NUC.esc(titulo)}</b>${texto ? `<span>${NUC.esc(texto)}</span>` : ''}${opciones.accion ? '<div class="acc"></div>' : ''}</div>`;
    if (opciones.accion) {
      const b = document.createElement('button');
      b.className = 'btn chico';
      b.textContent = opciones.accion.texto;
      b.addEventListener('click', () => { opciones.accion.fn(); quitar(); });
      el.querySelector('.acc').appendChild(b);
    }
    cont.appendChild(el);
    while (cont.children.length > 4) cont.firstElementChild.remove();
    let t = setTimeout(quitar, opciones.duracion || (tipo === 'error' ? 8000 : 5200));
    el.addEventListener('mouseenter', () => clearTimeout(t));
    el.addEventListener('mouseleave', () => { t = setTimeout(quitar, 2400); });
    function quitar() { el.classList.add('sale'); setTimeout(() => el.remove(), 220); }
    return el;
  }

  /* Controles segmentados (radiogroup) */
  function segmento(el, valor, alCambiar) {
    const ops = $$('.seg-op', el);
    const marcar = v => ops.forEach(o => o.setAttribute('aria-checked', String(o.dataset.v === v)));
    marcar(valor);
    ops.forEach(o => o.addEventListener('click', () => { marcar(o.dataset.v); alCambiar(o.dataset.v); }));
    el.addEventListener('keydown', e => {
      if (!['ArrowLeft', 'ArrowRight'].includes(e.key)) return;
      const i = ops.findIndex(o => o.getAttribute('aria-checked') === 'true');
      const j = (i + (e.key === 'ArrowRight' ? 1 : -1) + ops.length) % ops.length;
      ops[j].focus(); ops[j].click();
    });
    return { marcar };
  }

  /* Gráficos (Chart.js) con los colores del tema */
  const graficos = new Map();
  function colores() {
    return {
      azul: css('--azul-2') || '#2350C8', azul3: css('--azul-3') || '#8FA8F0', rojo: css('--rojo') || '#E61C24',
      tinta: css('--tinta') || '#0F1B3D', tinta2: css('--tinta-2'), tinta3: css('--tinta-3') || '#8A93AA',
      linea: css('--linea') || 'rgba(15,27,61,.08)', verde: css('--verde'), ambar: css('--ambar'), sup: css('--sup')
    };
  }
  function grafico(id, config) {
    const lienzo = document.getElementById(id);
    if (!lienzo || !window.Chart) return null;
    const c = colores();
    Chart.defaults.font.family = '"Archivo", ui-sans-serif, system-ui, sans-serif';
    Chart.defaults.font.size = 12;
    Chart.defaults.color = c.tinta3;
    Chart.defaults.borderColor = c.linea;
    const prev = graficos.get(id);
    if (prev) prev.destroy();
    const ch = new Chart(lienzo, config);
    graficos.set(id, ch);
    return ch;
  }
  const reducido = () => window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  /* Etiqueta de datos en el encabezado */
  function etiquetaDatos(M, local) {
    const p = [];
    if (M.capas && M.capas.historico) p.push(`histórico ${NUC.fmtFecha(M.capas.historico.desde)} – ${NUC.fmtFecha(M.capas.historico.hasta)}`);
    p.push(`ventas al ${NUC.fmtFecha(M.diaMax)}`);
    if (M.invFecha) p.push(`inventario ${NUC.fmtFecha(NUC.diaDeIso(M.invFecha))}`);
    return (local ? 'Vista previa local · ' : '') + p.join(' · ');
  }

  function descargar(nombre, contenido, tipo = 'application/json') {
    const blob = contenido instanceof Blob ? contenido : new Blob([contenido], { type: tipo });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = nombre;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  const leerLS = (k, def = null) => { try { const v = localStorage.getItem(k); return v == null ? def : JSON.parse(v); } catch (e) { return def; } };
  const guardarLS = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } };

  /* GitHub: lectura y escritura de archivos del repositorio */
  const GH_LS = 'vjs_github';
  const b64 = texto => {
    const bytes = new TextEncoder().encode(texto);
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  };
  function ghConfig() { return leerLS(GH_LS, null); }
  async function ghArchivo(g, ruta) {
    const url = `https://api.github.com/repos/${g.repo}/contents/${ruta.split('/').map(encodeURIComponent).join('/')}`;
    const h = { Authorization: `Bearer ${g.token}`, 'X-GitHub-Api-Version': '2022-11-28' };
    const meta = await fetch(`${url}?ref=${encodeURIComponent(g.rama)}`, { headers: { ...h, Accept: 'application/vnd.github+json' }, cache: 'no-store' });
    if (meta.status === 404) return { url, h, sha: undefined, json: null };
    if (meta.status === 401) throw new Error('GitHub no acepta el token: está mal copiado o venció.');
    if (!meta.ok) throw new Error(`GitHub respondió ${meta.status} al leer ${ruta}.`);
    const sha = (await meta.json()).sha;
    let json = null;
    const raw = await fetch(`${url}?ref=${encodeURIComponent(g.rama)}`, { headers: { ...h, Accept: 'application/vnd.github.raw+json' }, cache: 'no-store' });
    if (raw.ok) { try { json = await raw.json(); } catch (e) { json = null; } }
    return { url, h, sha, json };
  }
  async function ghEscribir(g, ruta, contenido, mensaje) {
    const a = await ghArchivo(g, ruta);
    const r = await fetch(a.url, {
      method: 'PUT', headers: { ...a.h, Accept: 'application/vnd.github+json' },
      body: JSON.stringify({ message: mensaje, content: b64(typeof contenido === 'string' ? contenido : JSON.stringify(contenido)), branch: g.rama, sha: a.sha })
    });
    if (!r.ok) {
      let m = '';
      try { m = (await r.json()).message || ''; } catch (e) { /* nada */ }
      throw new Error(`GitHub respondió ${r.status} al guardar ${ruta}${m ? ': ' + m : ''}.`);
    }
    return true;
  }
  const rutaDatos = (g, nombre) => (g.carpeta ? g.carpeta.replace(/^\/|\/$/g, '') + '/' : 'datos/') + nombre;

  window.UI = { $, $$, css, aviso, segmento, grafico, colores, reducido, etiquetaDatos, descargar, leerLS, guardarLS, ghConfig, ghArchivo, ghEscribir, rutaDatos, GH_LS };
})();
