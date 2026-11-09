/* =====================================================================
   SUGERIDO DE VENTAS · Venejugos 2015 — núcleo compartido
   Lectura de Excel (por partes para el histórico grande), procesamiento,
   fusión de capas, agrupador de descripciones, analítico y sugerido.
   Funciona en el navegador (window.NUC) y en Node (module.exports).
   ===================================================================== */
(function (raiz, fabrica) {
  const api = fabrica(raiz);
  if (typeof module === 'object' && module.exports) module.exports = api;
  else raiz.NUC = api;
})(typeof self !== 'undefined' ? self : globalThis, function (raiz) {
  'use strict';

  const libFflate = () => raiz.fflate || (typeof require === 'function' ? require('fflate') : null);
  const libXLSX = () => raiz.XLSX || (typeof require === 'function' ? require('xlsx') : null);
  const pausa = () => new Promise(r => setTimeout(r, 0));

  /* ---------------------------------------------------------------
     Utilidades
     --------------------------------------------------------------- */
  const norm = s => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
  const r2 = n => Math.round(n * 100) / 100;
  const r3 = n => Math.round(n * 1000) / 1000;

  function num(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    if (v == null || v === '' || typeof v === 'boolean') return 0;
    let s = String(v).trim().replace(/[^\d.,\-]/g, '');
    if (!s) return 0;
    const c = s.lastIndexOf(','), p = s.lastIndexOf('.');
    if (c >= 0 && p >= 0) s = c > p ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    else if (c >= 0) s = (s.match(/,/g).length > 1) ? s.replace(/,/g, '') : s.replace(',', '.');
    const n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }

  /* Códigos de producto y RIF: sin espacios, mayúsculas, sin ceros a la izquierda. */
  function codigo(v) {
    if (v == null) return '';
    let s = typeof v === 'number' ? String(Math.round(v * 1e6) / 1e6) : String(v);
    s = s.trim().toUpperCase().replace(/\s+/g, '');
    if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
    const z = s.replace(/^0+/, '');
    return z || (s ? '0' : '');
  }
  const rifNorm = v => String(v ?? '').toUpperCase().replace(/[\s.\-]/g, '').trim();

  /* Fechas: se trabaja con el número de día desde 1970-01-01 (UTC). */
  const DIA_MS = 864e5;
  function diaDe(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : Math.floor(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()) / DIA_MS);
    if (typeof v === 'number') return v > 20000 && v < 80000 ? Math.floor(v) - 25569 : null;
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return Math.floor(Date.UTC(+m[1], +m[2] - 1, +m[3]) / DIA_MS);
    m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
    if (m) { const a = +m[3] < 100 ? 2000 + +m[3] : +m[3]; return Math.floor(Date.UTC(a, +m[2] - 1, +m[1]) / DIA_MS); }
    if (/^\d+(\.\d+)?$/.test(s)) return diaDe(+s);
    return null;
  }
  const isoDe = d => new Date(d * DIA_MS).toISOString().slice(0, 10);
  const diaDeIso = iso => Math.floor(Date.parse(iso + 'T00:00:00Z') / DIA_MS);
  const fechaDe = d => new Date(d * DIA_MS);
  const dowDe = d => (fechaDe(d).getUTCDay() + 6) % 7;            // 0 = lunes
  const MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  const MESES_L = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const DIAS_S = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'];

  function encabezado(filas, requeridas, limite = 30) {
    const req = requeridas.map(norm);
    for (let i = 0; i < Math.min(filas.length, limite); i++) {
      const fila = (filas[i] || []).map(norm);
      if (req.every(r => fila.includes(r))) {
        const col = {};
        fila.forEach((n, j) => { if (n && !(n in col)) col[n] = j; });
        return { fila: i, col };
      }
    }
    return null;
  }
  const pick = (col, ...nombres) => { for (const n of nombres) { const k = norm(n); if (k in col) return col[k]; } return -1; };
  const val = (f, i) => (i >= 0 ? f[i] : undefined);
  const txt = (f, i) => { const v = val(f, i); return v == null ? '' : String(v).trim(); };

  /* ---------------------------------------------------------------
     Lectura de .xlsx por partes (sirve para archivos de cientos de MB)
     --------------------------------------------------------------- */
  function zipEntradas(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('El archivo no tiene formato .xlsx válido.');
    const n = dv.getUint16(eocd + 10, true);
    let off = dv.getUint32(eocd + 16, true);
    if (off === 0xffffffff) throw new Error('El archivo supera el tamaño que el panel puede leer (formato ZIP64).');
    const td = new TextDecoder();
    const ent = {};
    for (let k = 0; k < n; k++) {
      if (dv.getUint32(off, true) !== 0x02014b50) break;
      const metodo = dv.getUint16(off + 10, true);
      const comp = dv.getUint32(off + 20, true), tam = dv.getUint32(off + 24, true);
      const ln = dv.getUint16(off + 28, true), le = dv.getUint16(off + 30, true), lc = dv.getUint16(off + 32, true);
      const loc = dv.getUint32(off + 42, true);
      const nombre = td.decode(u8.subarray(off + 46, off + 46 + ln)).replace(/^\//, '');
      ent[nombre.toLowerCase()] = { nombre, metodo, comp, tam, loc };
      off += 46 + ln + le + lc;
    }
    return ent;
  }
  function zipDatos(u8, e) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const ln = dv.getUint16(e.loc + 26, true), le = dv.getUint16(e.loc + 28, true);
    const ini = e.loc + 30 + ln + le;
    return u8.subarray(ini, ini + e.comp);
  }
  function zipTexto(u8, ent, nombre) {
    const e = ent[nombre.toLowerCase()];
    if (!e) return null;
    const d = zipDatos(u8, e);
    const bytes = e.metodo === 0 ? d : libFflate().inflateSync(d);
    return new TextDecoder().decode(bytes);
  }
  async function zipFlujo(u8, e, alTexto, progreso) {
    const datos = zipDatos(u8, e);
    const dec = new TextDecoder('utf-8');
    let parar = false;
    const PASO = 1 << 20;
    if (e.metodo === 0) {
      for (let i = 0; i < datos.length && !parar; i += PASO * 4) {
        const fin = Math.min(datos.length, i + PASO * 4);
        if (alTexto(dec.decode(datos.subarray(i, fin), { stream: fin < datos.length })) === false) parar = true;
        if (progreso) progreso(fin / datos.length);
        await pausa();
      }
      return;
    }
    const inf = new (libFflate().Inflate)((chunk, final) => {
      if (parar) return;
      if (alTexto(dec.decode(chunk, { stream: !final })) === false) parar = true;
    });
    for (let i = 0; i < datos.length && !parar; i += PASO) {
      const fin = Math.min(datos.length, i + PASO);
      inf.push(datos.subarray(i, fin), fin === datos.length);
      if (progreso) progreso(fin / datos.length);
      await pausa();
    }
  }

  const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  function desescapar(s) {
    if (s.indexOf('&') >= 0) s = s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, e) => {
      if (e[0] === '#') return String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
      return ENT[e.toLowerCase()] || m;
    });
    if (s.indexOf('_x') >= 0) s = s.replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
    return s;
  }
  function leerSST(xml) {
    const out = [];
    if (!xml) return out;
    let pos = 0;
    for (;;) {
      const a = xml.indexOf('<si', pos);
      if (a < 0) break;
      const g = xml.indexOf('>', a);
      if (xml.charCodeAt(g - 1) === 47) { out.push(''); pos = g + 1; continue; }
      const b = xml.indexOf('</si>', g);
      let s = xml.slice(g + 1, b);
      if (s.indexOf('<rPh') >= 0) s = s.replace(/<rPh[\s\S]*?<\/rPh>/g, '');
      let t = '', p = 0;
      for (;;) {
        const i = s.indexOf('<t', p);
        if (i < 0) break;
        const c = s.charCodeAt(i + 2);
        if (c !== 62 && c !== 32) { p = i + 2; continue; }
        const j = s.indexOf('>', i);
        if (s.charCodeAt(j - 1) === 47) { p = j + 1; continue; }
        const k = s.indexOf('</t>', j);
        t += s.slice(j + 1, k);
        p = k + 4;
      }
      out.push(desescapar(t));
      pos = b + 5;
    }
    return out;
  }
  function colIdx(ref) {
    let n = 0;
    for (let i = 0; i < ref.length; i++) {
      const c = ref.charCodeAt(i);
      if (c < 65 || c > 90) break;
      n = n * 26 + (c - 64);
    }
    return n - 1;
  }
  function atrib(tag, nombre) {
    const k = tag.indexOf(' ' + nombre + '="');
    if (k < 0) return null;
    const i = k + nombre.length + 3;
    return tag.slice(i, tag.indexOf('"', i));
  }
  function filaXml(xml, ini, fin, sst) {
    const fila = [];
    let pos = ini, col = -1;
    for (;;) {
      const a = xml.indexOf('<c', pos);
      if (a < 0 || a >= fin) break;
      const ch = xml.charCodeAt(a + 2);
      if (ch !== 32 && ch !== 62 && ch !== 47) { pos = a + 2; continue; }
      const g = xml.indexOf('>', a);
      const tag = xml.slice(a + 2, g);
      const r = atrib(tag, 'r');
      col = r ? colIdx(r) : col + 1;
      if (xml.charCodeAt(g - 1) === 47) { pos = g + 1; continue; }
      const z = xml.indexOf('</c>', g);
      pos = z + 4;
      const t = atrib(tag, 't');
      let v = null;
      if (t === 'inlineStr') {
        const i = xml.indexOf('<t', g);
        if (i >= 0 && i < z) { const j = xml.indexOf('>', i); const k = xml.indexOf('</t>', j); if (k >= 0 && k < z) v = desescapar(xml.slice(j + 1, k)); }
      } else {
        const i = xml.indexOf('<v>', g);
        if (i >= 0 && i < z) v = xml.slice(i + 3, xml.indexOf('</v>', i));
      }
      if (v === null) continue;
      if (t === 's') v = sst[+v];
      else if (t === 'str') v = desescapar(v);
      else if (t === 'b') v = v === '1';
      else if (t === 'e') v = null;
      else if (t !== 'inlineStr') { const n = +v; v = isNaN(n) ? v : n; }
      if (v !== null && v !== undefined) fila[col] = v;
    }
    return fila;
  }
  function lectorHoja(sst, alFila) {
    let buf = '';
    let parado = false;
    return function (texto) {
      if (parado) return false;
      buf = buf ? buf + texto : texto;
      let pos = 0, corte = -1;
      for (;;) {
        const a = buf.indexOf('<row', pos);
        if (a < 0) break;
        const g = buf.indexOf('>', a);
        if (g < 0) { corte = a; break; }
        if (buf.charCodeAt(g - 1) === 47) { pos = g + 1; continue; }
        const b = buf.indexOf('</row>', g);
        if (b < 0) { corte = a; break; }
        const fila = filaXml(buf, g + 1, b, sst);
        pos = b + 6;
        if (alFila(fila) === false) { parado = true; return false; }
      }
      buf = corte >= 0 ? buf.slice(corte) : buf.slice(Math.max(pos, buf.length - 8));
      return true;
    };
  }

  /* Abre un libro y devuelve la lista de hojas con un recorrido fila por fila.
     .xlsx se lee por partes; .xls / .csv con SheetJS. */
  async function abrirLibro(buffer, nombreArchivo = '') {
    const u8 = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    const esZip = u8[0] === 0x50 && u8[1] === 0x4b;
    if (esZip) {
      const ent = zipEntradas(u8);
      if (!ent['xl/workbook.xml']) throw new Error('El archivo .xlsx no tiene libro de trabajo.');
      const wb = zipTexto(u8, ent, 'xl/workbook.xml');
      const rels = zipTexto(u8, ent, 'xl/_rels/workbook.xml.rels') || '';
      const destino = {};
      rels.replace(/<Relationship\b[^>]*>/g, tag => {
        const id = (tag.match(/\bId="([^"]+)"/) || [])[1], t = (tag.match(/\bTarget="([^"]+)"/) || [])[1];
        if (id && t) destino[id] = t.replace(/^\/?(xl\/)?/, 'xl/');
        return tag;
      });
      const hojas = [];
      wb.replace(/<sheet\b[^>]*>/g, tag => {
        const nombre = desescapar((tag.match(/\bname="([^"]*)"/) || [])[1] || '');
        const rid = (tag.match(/\br:id="([^"]+)"/) || tag.match(/\bid="([^"]+)"/) || [])[1];
        const ruta = destino[rid];
        if (ruta && ent[ruta.toLowerCase()]) hojas.push({ nombre, ruta });
        return tag;
      });
      let sst = null;
      const cargarSST = () => { if (!sst) sst = leerSST(zipTexto(u8, ent, 'xl/sharedStrings.xml')); return sst; };
      return {
        tipo: 'xlsx',
        hojas: hojas.map(h => h.nombre),
        tamano: u8.length,
        async recorrer(nombreHoja, alFila, progreso) {
          const h = hojas.find(x => x.nombre === nombreHoja);
          if (!h) throw new Error(`No se encontró la hoja ${nombreHoja}.`);
          const lector = lectorHoja(cargarSST(), alFila);
          await zipFlujo(u8, ent[h.ruta.toLowerCase()], lector, progreso);
        }
      };
    }
    const X = libXLSX();
    if (!X) throw new Error('Falta la librería para leer archivos .xls.');
    // SheetJS avisa en consola de registros que no necesita (por ejemplo 0x27d): se silencian durante la lectura.
    const errorOriginal = console.error;
    console.error = (...a) => { if (!/Missing Info for XLS Record/.test(String(a[0]))) errorOriginal.apply(console, a); };
    let libro;
    try { libro = X.read(u8, { type: 'array', dense: true, cellDates: false, cellNF: false, cellText: false }); }
    finally { console.error = errorOriginal; }
    return {
      tipo: 'xls',
      hojas: libro.SheetNames.slice(),
      tamano: u8.length,
      async recorrer(nombreHoja, alFila, progreso) {
        const filas = X.utils.sheet_to_json(libro.Sheets[nombreHoja], { header: 1, raw: true, defval: null, blankrows: false });
        for (let i = 0; i < filas.length; i++) {
          if (alFila(filas[i]) === false) break;
          if (progreso && i % 5000 === 0) { progreso(i / filas.length); await pausa(); }
        }
        if (progreso) progreso(1);
      }
    };
  }

  /* ---------------------------------------------------------------
     Detección del tipo de archivo por encabezados
     --------------------------------------------------------------- */
  const TIPOS = {
    historico: { nombre: 'Histórico de ventas', req: ['total_venta_dolar', 'codprod', 'MODULOS'] },
    b: { nombre: 'Ventas Módulo B', req: ['total_venta_dolar', 'codprod', 'producto'] },
    mv: { nombre: 'Ventas MobilVendor', req: ['RucCliente', 'CodProducto', 'PrecioTotal'] },
    inventario: { nombre: 'Inventario', req: ['Codigo', 'Productos', 'Stock'] }
  };
  function detectarTipo(filas) {
    for (const t of ['historico', 'mv', 'inventario', 'b']) {
      const h = encabezado(filas, TIPOS[t].req);
      if (h) return { tipo: t, ...h };
    }
    return null;
  }

  /* ---------------------------------------------------------------
     Procesadores de filas (acumuladores)
     --------------------------------------------------------------- */
  const UNIDADES_BASE = {
    codigos: { 'UNI': 'UNI', 'UND.': 'UNI', 'UND': 'UNI', 'UN': 'UNI', 'UNIDAD': 'UNI', 'CAJ': 'CAJ', 'CAJAS': 'CAJ', 'CAJA': 'CAJ', 'CJ': 'CAJ', 'CJAX36': 'CAJ', 'DIS': 'DIS', 'DISPLAY': 'DIS', 'BUL': 'BUL', 'BULTO': 'BUL', 'KG': 'KG', 'KG.': 'KG', 'KGS': 'KG', 'SIX': 'SIX', 'MAN': 'MAN', 'MBU': 'MBU', 'MULTIPACK': 'MUL', '#EMP': 'CAJ' },
    nombres: { UNI: 'Unidad', CAJ: 'Caja', DIS: 'Display', BUL: 'Bulto', KG: 'Kilo', SIX: 'Six pack', MAN: 'Mano', MBU: 'Multibulto', MUL: 'Multipack' }
  };
  const UNIDADES_SUELTAS = ['UNI', 'KG'];

  function nuevoAcum(fuente, cfg) {
    const excl = ((cfg && cfg.excluirClientes) || ['INLUVENCA']).map(norm).filter(Boolean);
    return {
      fuente, hechos: new Map(), clientes: new Map(), items: new Map(), pres: new Map(), nombresCli: new Map(),
      meta: { filas: 0, lineas: 0, desde: null, hasta: null, venta: 0, devol: 0, dias: new Set(), excluidas: 0, montoExcluido: 0, omitidas: 0, avisos: [] },
      excluido(nombre, rif) { return excl.length && excl.some(x => norm(nombre).includes(x) || rifNorm(rif) === x); }
    };
  }
  function sumarLinea(A, L) {
    // L: {dia, item, desc, uni, fac, monto, q, qp, rif, cli, vend, unH, unR, ma, li, sl, unCli}
    const M = A.meta;
    M.lineas++;
    if (L.dia < (M.desde ?? Infinity)) M.desde = L.dia;
    if (L.dia > (M.hasta ?? -Infinity)) M.hasta = L.dia;
    M.dias.add(L.dia);
    if (L.monto >= 0) M.venta += L.monto; else M.devol -= L.monto;
    const kh = L.dia + '|' + L.item + '|' + L.uni;
    let h = A.hechos.get(kh);
    if (!h) { h = [0, 0, 0, 0]; A.hechos.set(kh, h); }
    if (L.monto >= 0) h[0] += L.monto; else h[1] -= L.monto;
    h[2] += L.q; h[3] += L.qp;
    // Catálogo
    let it = A.items.get(L.item);
    if (!it) { it = { desc: L.desc, dia: L.dia, unH: new Map(), unR: new Map(), mm: new Map(), ma: new Map(), li: new Map(), sl: new Map() }; A.items.set(L.item, it); }
    if (L.dia >= it.dia && L.desc) { it.desc = L.desc; it.dia = L.dia; }
    const peso = Math.abs(L.monto) + 0.01;
    const suma = (m, k) => { if (k) m.set(k, (m.get(k) || 0) + peso); };
    suma(it.unH, L.unH); suma(it.unR, L.unR); suma(it.mm, L.mvMarca); suma(it.ma, L.ma); suma(it.li, L.li); suma(it.sl, L.sl);
    const regPres = (u, fac) => {
      const kp = L.item + '|' + u;
      let p = A.pres.get(kp);
      if (!p) { p = new Map(); A.pres.set(kp, p); }
      const f = r3(fac);
      p.set(f, (p.get(f) || 0) + 1);
    };
    if (L.uni && L.fac > 0) regPres(L.uni, L.fac);
    if (L.facEmp > 1) regPres('#EMP', L.facEmp);
    // Clientes
    if (L.rif) {
      const mod = L.item[0];
      const kc = L.dia + '|' + L.rif + '|' + (L.unCli || '') + '|' + mod;
      let c = A.clientes.get(kc);
      if (!c) { c = [0, '']; A.clientes.set(kc, c); }
      c[0] += L.monto;
      if (L.vend) c[1] = L.vend;
      if (L.cli) A.nombresCli.set(L.rif, L.cli);
    }
  }

  const malo = s => !s || s === '0' || s === '#N/A' || s === 'N/A' || s === 'POR ASIGNAR' || s === 'NO APLICA';
  const limpio = s => { const n = norm(s); return malo(n) ? '' : n; };

  /* Histórico consolidado (DINAMICA) y archivo del Módulo B: misma estructura base. */
  function procesadorSistema(enc, cfg, fuente, opciones = {}) {
    const A = nuevoAcum(fuente, cfg);
    const c = enc.col;
    const I = {
      mod: pick(c, 'MODULOS', 'MODULO'), fecha: pick(c, 'fecha', 'fechafact'), rif: pick(c, 'RIF REAL', 'rif'), cli: pick(c, 'nomclie'),
      vend: pick(c, 'nomvend', 'ASESOR ACTUAL'), cod: pick(c, 'codprod'), alt: pick(c, 'alterno'), prod: pick(c, 'producto'),
      uni: pick(c, 'unidad'), emp: pick(c, 'emp_aux'), qb: pick(c, 'total_emp_aux'), qn: pick(c, 'venta_neta'), monto: pick(c, 'total_venta_dolar'),
      un: pick(c, 'UNIDAD NEGOCIO'), unS: pick(c, 'unidad_negocio'), ma: pick(c, 'MARCA'), maS: pick(c, 'marca3'),
      li: pick(c, 'LINEA'), liS: pick(c, 'linea4'), sl: pick(c, 'SUBLINEA'), slS: pick(c, 'sublinea5')
    };
    const esHist = fuente === 'historico';
    return {
      A,
      fila(f) {
        A.meta.filas++;
        const dia = diaDe(val(f, I.fecha));
        if (dia == null) return;
        const modTxt = norm(val(f, I.mod));
        let mod = esHist ? (/MOBIL/.test(modTxt) ? 'M' : 'B') : 'B';
        if (!esHist && I.mod >= 0 && /MOBIL/.test(modTxt)) { if (opciones.omitirMV) { A.meta.omitidas++; return; } mod = 'M'; }
        const cp = codigo(val(f, I.cod)), alt = codigo(val(f, I.alt));
        let item;
        if (mod === 'M') item = cp ? 'M:' + cp : null;
        else item = alt ? 'B:' + alt : (cp ? 'B:#' + cp : null);
        if (!item) return;
        const rif = rifNorm(val(f, I.rif)), cli = txt(f, I.cli);
        const monto = num(val(f, I.monto));
        if (A.excluido(cli, rif)) { A.meta.excluidas++; A.meta.montoExcluido += monto; return; }
        const emp = num(val(f, I.emp)) || 1;
        const uniRaw = norm(val(f, I.uni));
        const suelta = UNIDADES_SUELTAS.includes(UNIDADES_BASE.codigos[uniRaw] || uniRaw);
        const tea = I.qb >= 0 ? num(val(f, I.qb)) : null;
        const vn = I.qn >= 0 ? num(val(f, I.qn)) : (tea != null ? tea / emp : 0);
        // Unidad suelta (UNI, UND., KG): venta_neta ya está en unidades; emp_aux es el empaque de compra.
        // Empaque (CAJ, DIS, BUL…): total_emp_aux trae las unidades sueltas.
        const q = suelta ? vn : (tea != null ? tea : vn * emp);
        const qp = vn;
        if (!monto && !q) return;
        const unH = esHist ? limpio(val(f, I.un)) : '';
        const unR = limpio(val(f, I.unS));
        const ma = limpio(val(f, I.ma)) || limpio(val(f, I.maS));
        const li = limpio(val(f, I.li)) || limpio(val(f, I.liS));
        const sl = limpio(val(f, I.sl)) || limpio(val(f, I.slS));
        sumarLinea(A, {
          dia, item, desc: txt(f, I.prod), uni: uniRaw, fac: suelta ? 1 : emp, facEmp: suelta && emp > 1 ? emp : 0, monto, q, qp,
          rif, cli, vend: norm(val(f, I.vend)), unH, unR, ma, li, sl, unCli: unH || unR
        });
      },
      cerrar: () => cerrarAcum(A)
    };
  }

  function procesadorMV(enc, cfg) {
    const A = nuevoAcum('mv', cfg);
    const c = enc.col;
    const I = {
      fecha: pick(c, 'FechaFactura', 'Fecha'), rif: pick(c, 'RucCliente'), cli: pick(c, 'RazónSocial', 'RazonSocial'), vend: pick(c, 'NombreVend'),
      cod: pick(c, 'CodProducto'), prod: pick(c, 'DescrProducto'), uni: pick(c, 'UnidadVenta'), fac: pick(c, 'UnidadFactor'),
      cant: pick(c, 'Cantidad'), cantB: pick(c, 'CantidadEnUnidadInventario'), monto: pick(c, 'PrecioTotal'),
      trans: pick(c, 'CodTrans'), dtrans: pick(c, 'DescTrans'), marca: pick(c, 'DescrMarca'), fam: pick(c, 'DescrFamilia'),
      cat: pick(c, 'DescrCategoría', 'DescrCategoria'), sub: pick(c, 'DescrSubCategoría', 'DescrSubCategoria')
    };
    return {
      A,
      fila(f) {
        A.meta.filas++;
        const dia = diaDe(val(f, I.fecha));
        if (dia == null) return;
        const tr = norm(val(f, I.trans)), dtr = norm(val(f, I.dtrans));
        if (tr === 'PD' || /PEDIDO/.test(dtr)) { A.meta.omitidas++; return; }
        const cp = codigo(val(f, I.cod));
        if (!cp) return;
        const rif = rifNorm(val(f, I.rif)), cli = txt(f, I.cli);
        const monto = num(val(f, I.monto));
        if (A.excluido(cli, rif)) { A.meta.excluidas++; A.meta.montoExcluido += monto; return; }
        const uniRaw = norm(val(f, I.uni));
        const suelta = UNIDADES_SUELTAS.includes(UNIDADES_BASE.codigos[uniRaw] || uniRaw);
        const fac = suelta ? 1 : (num(val(f, I.fac)) || 1);
        const cant = num(val(f, I.cant));
        const q = suelta ? cant : (I.cantB >= 0 && val(f, I.cantB) != null ? num(val(f, I.cantB)) : cant * fac);
        if (!monto && !q) return;
        const unR = limpio(val(f, I.marca));
        sumarLinea(A, {
          dia, item: 'M:' + cp, desc: txt(f, I.prod), uni: uniRaw, fac, monto, q, qp: cant,
          rif, cli, vend: norm(val(f, I.vend)), unH: '', unR, ma: limpio(val(f, I.fam)), li: limpio(val(f, I.cat)), sl: limpio(val(f, I.sub)), unCli: unR, mvMarca: unR
        });
      },
      cerrar: () => cerrarAcum(A)
    };
  }

  function procesadorInventario(enc) {
    const c = enc.col;
    const I = {
      mod: pick(c, 'Modulo', 'Módulo'), cod: pick(c, 'Codigo', 'Código'), prod: pick(c, 'Productos', 'Producto'), uni: pick(c, 'Unidad'),
      stock: pick(c, 'Stock', 'Existencia'), precio: pick(c, 'Precio'), costo: pick(c, 'Costo'), un: pick(c, 'Unidad de Negocio'),
      sl: pick(c, 'Sublinea'), li: pick(c, 'Lineas', 'Linea'), es: pick(c, 'estatus', 'Estatus'), cat: pick(c, 'CATEGORIAS')
    };
    const filas = [];
    const meta = { filas: 0, validas: 0, avisos: [] };
    return {
      fila(f) {
        meta.filas++;
        const cod = codigo(val(f, I.cod)), desc = txt(f, I.prod);
        if (!cod || !desc || norm(desc) === 'PRODUCTOS') return;
        const m = norm(val(f, I.mod));
        filas.push({
          mod: /MOBIL/.test(m) ? 'M' : /AMBOS/.test(m) ? 'A' : 'B', cod, desc, uni: norm(val(f, I.uni)),
          s: num(val(f, I.stock)), p: num(val(f, I.precio)), c: num(val(f, I.costo)),
          un: limpio(val(f, I.un)), li: limpio(val(f, I.li)), sl: limpio(val(f, I.sl)), es: norm(val(f, I.es))
        });
        meta.validas++;
      },
      cerrar: () => ({ filas, meta })
    };
  }

  function mayor(m) { let k = '', v = -1; for (const [a, b] of m) if (b > v) { v = b; k = a; } return k; }
  function cerrarAcum(A) {
    const M = A.meta;
    return {
      fuente: A.fuente,
      meta: {
        filas: M.filas, lineas: M.lineas, desde: M.desde, hasta: M.hasta, dias: [...M.dias].sort((a, b) => a - b),
        venta: r2(M.venta), devol: r2(M.devol), excluidas: M.excluidas, montoExcluido: r2(M.montoExcluido), omitidas: M.omitidas
      },
      hechos: [...A.hechos].map(([k, h]) => { const [d, i, u] = k.split('|'); return { d: +d, i, u, vb: h[0], dv: h[1], q: h[2], qp: h[3] }; }),
      clientes: [...A.clientes].map(([k, c]) => { const [d, r, n, m] = k.split('|'); return { d: +d, r, n, m, v: c[0], a: c[1] }; }),
      items: [...A.items].map(([k, it]) => ({ k, d: it.desc, dia: it.dia, unH: mayor(it.unH), unR: mayor(it.unR), mm: mayor(it.mm), ma: mayor(it.ma), li: mayor(it.li), sl: mayor(it.sl) })),
      pres: [...A.pres].map(([k, m]) => { const i = k.lastIndexOf('|'); return { i: k.slice(0, i), u: k.slice(i + 1), f: +mayor(m), n: [...m.values()].reduce((a, b) => a + b, 0) }; }),
      nombresCli: [...A.nombresCli]
    };
  }

  /* Lee un archivo completo: detecta hoja y tipo, y lo procesa en una sola pasada. */
  async function procesarArchivo(buffer, nombre, cfg, opciones = {}) {
    const libro = await abrirLibro(buffer, nombre);
    const esperado = opciones.tipo || null;
    for (const hoja of libro.hojas) {
      const primeras = [];
      let det = null, proc = null, ignorar = false;
      await libro.recorrer(hoja, fila => {
        if (proc) { proc.fila(fila); return; }
        primeras.push(fila);
        det = detectarTipo(primeras);
        if (det && esperado && det.tipo !== esperado) {
          if (!(esperado === 'b' && det.tipo === 'historico' && opciones.permitirSistema)) { ignorar = true; return false; }
        }
        if (det) {
          const t = esperado || det.tipo;
          proc = t === 'mv' ? procesadorMV(det, cfg)
            : t === 'inventario' ? procesadorInventario(det)
            : procesadorSistema(det, cfg, t === 'historico' ? 'historico' : 'b', opciones);
          for (let i = det.fila + 1; i < primeras.length; i++) proc.fila(primeras[i]);
          return;
        }
        if (primeras.length >= 30) return false;
      }, opciones.progreso);
      if (proc) return { tipo: esperado || det.tipo, hoja, datos: proc.cerrar() };
      if (ignorar && esperado) continue;
    }
    throw new Error(esperado
      ? `No se encontraron las columnas de ${TIPOS[esperado].nombre} (${TIPOS[esperado].req.join(', ')}).`
      : 'No se reconoció el archivo: revisa que sea histórico, ventas MobilVendor, ventas Módulo B o inventario.');
  }

  /* ---------------------------------------------------------------
     Conjunto de ventas (ventas.json): codificar, decodificar y fusionar capas
     --------------------------------------------------------------- */
  function vacioVentas() {
    return { capas: { historico: null, M: null, B: null }, hechos: [], clientes: [], items: new Map(), pres: new Map(), nombresCli: new Map() };
  }
  function decodificarVentas(J) {
    const V = vacioVentas();
    if (!J || !J.h) return V;
    V.capas = J.capas || V.capas;
    const D = J.dims;
    const keys = J.items.k;
    J.items.k.forEach((k, j) => V.items.set(k, {
      k, d: J.items.d[j], dia: J.items.dia[j], unH: D.un[J.items.unH[j]] || '', unR: D.un[J.items.unR[j]] || '', mm: D.un[(J.items.mm || [])[j]] || '',
      ma: D.ma[J.items.ma[j]] || '', li: D.li[J.items.li[j]] || '', sl: D.sl[J.items.sl[j]] || ''
    }));
    const h = J.h;
    for (let j = 0; j < h.d.length; j++) V.hechos.push({ d: h.d[j], i: keys[h.i[j]], u: D.u[h.u[j]], vb: h.vb[j], dv: h.dv[j], q: h.q[j], qp: h.qp[j], o: h.o[j] });
    const c = J.c;
    for (let j = 0; j < c.d.length; j++) V.clientes.push({ d: c.d[j], r: J.cli.r[c.k[j]], n: D.un[c.n[j]] || '', m: c.m[j] ? 'B' : 'M', v: c.v[j], a: D.as[c.a[j]] || '', o: c.o[j] });
    J.cli.r.forEach((r, j) => V.nombresCli.set(r, J.cli.n[j]));
    (J.pf || []).forEach(([i, u, f, n]) => V.pres.set(keys[i] + '|' + D.u[u], { f, n }));
    return V;
  }
  function codificarVentas(V) {
    const dims = { un: [''], ma: [''], li: [''], sl: [''], u: [], as: [''] };
    const idx = {};
    const di = (t, v) => { v = v || ''; const m = idx[t] || (idx[t] = new Map(dims[t].map((x, j) => [x, j]))); let j = m.get(v); if (j == null) { j = dims[t].length; dims[t].push(v); m.set(v, j); } return j; };
    const keys = [...V.items.keys()].sort();
    const ki = new Map(keys.map((k, j) => [k, j]));
    const items = { k: keys, d: [], dia: [], unH: [], unR: [], mm: [], ma: [], li: [], sl: [] };
    for (const k of keys) {
      const it = V.items.get(k);
      items.d.push(it.d); items.dia.push(it.dia ?? 0); items.unH.push(di('un', it.unH)); items.unR.push(di('un', it.unR)); items.mm.push(di('un', it.mm));
      items.ma.push(di('ma', it.ma)); items.li.push(di('li', it.li)); items.sl.push(di('sl', it.sl));
    }
    const hs = V.hechos.slice().sort((a, b) => a.d - b.d || (ki.get(a.i) - ki.get(b.i)));
    const h = { d: [], i: [], u: [], vb: [], dv: [], q: [], qp: [], o: [] };
    for (const x of hs) {
      if (!ki.has(x.i)) continue;
      h.d.push(x.d); h.i.push(ki.get(x.i)); h.u.push(di('u', x.u)); h.vb.push(r2(x.vb)); h.dv.push(r2(x.dv)); h.q.push(r3(x.q)); h.qp.push(r3(x.qp)); h.o.push(x.o);
    }
    const rifs = [...new Set(V.clientes.map(c => c.r))].sort();
    const ri = new Map(rifs.map((r, j) => [r, j]));
    const cs = V.clientes.slice().sort((a, b) => a.d - b.d);
    const c = { d: [], k: [], n: [], m: [], v: [], a: [], o: [] };
    for (const x of cs) { c.d.push(x.d); c.k.push(ri.get(x.r)); c.n.push(di('un', x.n)); c.m.push(x.m === 'B' ? 1 : 0); c.v.push(r2(x.v)); c.a.push(di('as', x.a)); c.o.push(x.o); }
    const pf = [];
    for (const [k, p] of V.pres) { const j = k.lastIndexOf('|'); const i = ki.get(k.slice(0, j)); if (i != null) pf.push([i, di('u', k.slice(j + 1)), p.f, p.n]); }
    return {
      version: 1, app: 'sugerido-ventas', generado: new Date().toISOString(), capas: V.capas, dims, items, h, c,
      cli: { r: rifs, n: rifs.map(r => V.nombresCli.get(r) || '') }, pf, equivAuto: equivalenciasAuto(V)
    };
  }

  /* Aprende las equivalencias de marca hacia la unidad de negocio consolidada del histórico:
     DescrMarca de MobilVendor (mm) y unidad_negocio del sistema (unR). */
  function equivalenciasAuto(V) {
    const cuenta = new Map();
    const sumar = (r, h) => { if (!r || !h) return; const m = cuenta.get(r) || new Map(); m.set(h, (m.get(h) || 0) + 1); cuenta.set(r, m); };
    for (const it of V.items.values()) { sumar(it.unR, it.unH); }
    for (const it of V.items.values()) { sumar(it.mm, it.unH); sumar(it.mm, it.unH); }
    const out = {};
    for (const [r, m] of cuenta) out[r] = mayor(m);
    return out;
  }

  /* Fusiona una capa nueva (histórico o diario de un módulo) en el conjunto. */
  function fusionarCapa(V, capa, archivo) {
    const esHist = capa.fuente === 'historico';
    const resumen = { reemplazados: 0, descartados: 0, agregados: 0 };
    if (esHist) {
      const hasta = capa.meta.hasta;
      const antes = V.hechos.length;
      V.hechos = V.hechos.filter(x => x.o === 1 && x.d > hasta);
      V.clientes = V.clientes.filter(x => x.o === 1 && x.d > hasta);
      resumen.reemplazados = antes - V.hechos.length;
      V.capas.historico = { desde: capa.meta.desde, hasta, archivo, filas: capa.meta.lineas, venta: capa.meta.venta, devol: capa.meta.devol, procesado: new Date().toISOString() };
      for (const m of ['M', 'B']) if (V.capas[m]) V.capas[m].dias = V.capas[m].dias.filter(d => d > hasta);
    }
    const hastaHist = V.capas.historico ? V.capas.historico.hasta : -Infinity;
    const mods = esHist ? null : new Set(capa.hechos.map(x => x.i[0]));
    if (!esHist) {
      const dias = new Set(capa.meta.dias);
      const antes = V.hechos.length;
      V.hechos = V.hechos.filter(x => !(x.o === 1 && mods.has(x.i[0]) && dias.has(x.d)));
      V.clientes = V.clientes.filter(x => !(x.o === 1 && mods.has(x.m) && dias.has(x.d)));
      resumen.reemplazados = antes - V.hechos.length;
    }
    for (const x of capa.hechos) {
      if (!esHist && x.d <= hastaHist) { resumen.descartados++; continue; }
      V.hechos.push({ ...x, o: esHist ? 0 : 1 });
      resumen.agregados++;
    }
    for (const x of capa.clientes) {
      if (!esHist && x.d <= hastaHist) continue;
      V.clientes.push({ ...x, o: esHist ? 0 : 1 });
    }
    if (!esHist) {
      for (const m of mods) {
        const dias = capa.meta.dias.filter(d => d > hastaHist);
        const prev = V.capas[m] ? V.capas[m].dias : [];
        V.capas[m] = {
          dias: [...new Set([...prev, ...dias])].sort((a, b) => a - b), archivo, ultimaCarga: new Date().toISOString(),
          filas: capa.meta.lineas, venta: capa.meta.venta, devol: capa.meta.devol, rango: [capa.meta.desde, capa.meta.hasta]
        };
      }
    }
    // Catálogo de productos
    for (const it of capa.items) {
      const prev = V.items.get(it.k);
      if (!prev) { V.items.set(it.k, { ...it }); continue; }
      if ((it.dia ?? 0) >= (prev.dia ?? 0) && it.d) { prev.d = it.d; prev.dia = it.dia; }
      if (esHist) { prev.unH = it.unH || prev.unH; }
      for (const f of ['unR', 'ma', 'li', 'sl']) if (it[f] && (esHist || !prev[f])) prev[f] = it[f];
      if (it.mm) prev.mm = it.mm;
    }
    for (const p of capa.pres) {
      const k = p.i + '|' + p.u;
      const prev = V.pres.get(k);
      if (!prev || p.n >= prev.n || esHist) V.pres.set(k, { f: p.f, n: Math.max(p.n, prev ? prev.n : 0) });
    }
    for (const [r, n] of capa.nombresCli) V.nombresCli.set(r, n);
    // Unidad de negocio de los clientes del diario: se resuelve con el catálogo
    if (!esHist) {
      const eq = equivalenciasAuto(V);
      const unDe = new Map([...V.items.values()].map(it => [it.k, it.unH || eq[it.unR] || it.unR]));
      // Las filas de clientes del diario vienen con la marca cruda: se traducen
      for (const x of V.clientes) if (x.o === 1 && x.n && eq[x.n]) x.n = eq[x.n];
      void unDe;
    }
    return resumen;
  }

  function codificarInventario(res, archivo, fechaIso, V, C) {
    const keysV = V ? new Set(V.items.keys()) : new Set();
    const vistos = new Map();
    const out = { version: 1, generado: new Date().toISOString(), fecha: fechaIso, archivo, filas: { k: [], d: [], u: [], s: [], c: [], p: [], un: [], li: [], sl: [], es: [] }, avisos: [] };
    const F = out.filas;
    const dup = [];
    for (const r of res.filas) {
      let k;
      if (r.mod === 'M') k = 'M:' + r.cod;
      else if (r.mod === 'B') k = 'B:' + r.cod;
      else k = keysV.has('M:' + r.cod) ? 'M:' + r.cod : 'B:' + r.cod;
      if (vistos.has(k)) {
        const prev = vistos.get(k);
        if (claveDesc(prev) !== claveDesc(r.desc)) { dup.push(`${k} (${prev} / ${r.desc})`); k = k + '~' + (dup.length); }
      }
      vistos.set(k, r.desc);
      F.k.push(k); F.d.push(r.desc); F.u.push(r.uni); F.s.push(r3(r.s)); F.c.push(r.c); F.p.push(r.p); F.un.push(r.un); F.li.push(r.li); F.sl.push(r.sl); F.es.push(r.es);
    }
    if (dup.length) out.avisos.push({ tipo: 'duplicado', texto: `Códigos repetidos con descripción distinta en el inventario: ${dup.slice(0, 8).join('; ')}${dup.length > 8 ? '…' : ''}` });
    void C;
    return out;
  }

  /* ---------------------------------------------------------------
     Agrupador de descripciones
     --------------------------------------------------------------- */
  const PARADAS = new Set(['DE', 'DEL', 'LA', 'EL', 'LOS', 'LAS', 'CON', 'EN', 'Y', 'X', 'VE', 'SABOR', 'PAQ', 'PAQUETE', 'TIPO', 'P', 'C', 'A', 'AL', 'S']);
  function tokens(desc) {
    let s = norm(desc).replace(/[®™©]/g, '').replace(/(\d),(\d)/g, '$1.$2');
    s = s.replace(/([A-Z])(\d)/g, '$1 $2').replace(/(\d)([A-Z])/g, '$1 $2').replace(/[^A-Z0-9.]+/g, ' ');
    const crudo = s.split(' ').filter(Boolean);
    const UNI = { G: 'G', GR: 'G', GRS: 'G', GRAMOS: 'G', GRM: 'G', KG: 'KG', KGS: 'KG', KILO: 'KG', L: 'L', LT: 'L', LTS: 'L', LTR: 'L', LITRO: 'L', LITROS: 'L', ML: 'ML', CC: 'ML', UND: 'UN', UNID: 'UN', UNIDADES: 'UN', UN: 'UN', U: 'UN', UNDS: 'UN', OZ: 'OZ', MTS: 'M', M: 'M' };
    const pal = [], med = [];
    for (let i = 0; i < crudo.length; i++) {
      const t = crudo[i];
      if (/^\d+(\.\d+)?$/.test(t)) {
        const u = UNI[crudo[i + 1]];
        if (u) { med.push(+t + u); i++; } else med.push(String(+t));
        continue;
      }
      const t2 = t.replace(/\./g, '');
      if (!t2 || PARADAS.has(t2) || UNI[t2]) continue;
      pal.push(t2);
    }
    return { pal, med };
  }
  const claveDesc = desc => { const t = tokens(desc); return t.pal.join(' ') + '|' + t.med.join(' '); };
  function similitud(ta, tb) {
    const A = new Set(ta.pal), B = new Set(tb.pal);
    if (!A.size || !B.size) return 0;
    let inter = 0; for (const x of A) if (B.has(x)) inter++;
    const jac = inter / (A.size + B.size - inter);
    const MA = new Set(ta.med), MB = new Set(tb.med);
    if (MA.size && MB.size) {
      let im = 0; for (const x of MA) if (MB.has(x)) im++;
      if (im !== MA.size || im !== MB.size) return jac * (im ? 0.55 : 0.3);
    } else if (MA.size || MB.size) return jac * 0.85;
    return jac;
  }
  function unionFind(n) {
    const p = new Int32Array(n).map((_, i) => i);
    const f = x => { while (p[x] !== x) { p[x] = p[p[x]]; x = p[x]; } return x; };
    return { f, u(a, b) { a = f(a); b = f(b); if (a !== b) p[Math.max(a, b)] = Math.min(a, b); } };
  }

  /* ---------------------------------------------------------------
     Modelo para las vistas: une ventas + inventario + configuración
     --------------------------------------------------------------- */
  const CONFIG_BASE = {
    version: 1,
    agrupador: { unir: [], separar: [], distintos: [], grupos: {} },
    equiv: { 'NIM': 'NESTLE', 'NP': 'NESTLE', 'BIGOTT': 'DISTRIBUIDORA BIGOTT', 'LA NORMANDA': 'NORMANDA' },
    unidades: UNIDADES_BASE,
    sugerido: { ventana: 90, reciente: 30, pesoReciente: 0.5, cobertura: 15, seguridad: 5, exceso: 60, redondeo: 0.25 },
    marcas: {},
    cigarros: { un: 'BIGOTT|CIGARR', linea: 'TABACO|CIGARR' },
    noProducto: ['AJUSTES'],
    excluirClientes: ['INLUVENCA']
  };
  function completarConfig(C) {
    const B = JSON.parse(JSON.stringify(CONFIG_BASE));
    if (!C) return B;
    const out = { ...B, ...C };
    out.agrupador = { ...B.agrupador, ...(C.agrupador || {}) };
    out.equiv = { ...B.equiv, ...(C.equiv || {}) };
    out.unidades = { codigos: { ...B.unidades.codigos, ...((C.unidades || {}).codigos || {}) }, nombres: { ...B.unidades.nombres, ...((C.unidades || {}).nombres || {}) } };
    out.sugerido = { ...B.sugerido, ...(C.sugerido || {}) };
    out.cigarros = { ...B.cigarros, ...(C.cigarros || {}) };
    return out;
  }

  function factorDescripcion(desc) {
    const s = norm(desc).replace(/(\d),(\d)/g, '$1.$2');
    let m = s.match(/\b(\d{1,3})\s*\(\s*(\d{1,3})\s*X/);              // 16(20X32G) -> 20 por display
    if (m) return { f: +m[2], fuente: 'descripción' };
    m = s.match(/\bX\s*(\d{1,3})\s*(UND|UNID|UN|U)\b/) || s.match(/\b(\d{1,3})\s*(UND|UNID|UNIDADES)\b/) || s.match(/\bCAJA\s*(\d{1,3})\b/) || s.match(/\b(\d{1,3})\s*X\s*\d/) || s.match(/\b(\d{1,3})\s*\/\s*\d/);
    if (m && +m[1] > 1) return { f: +m[1], fuente: 'descripción' };
    return null;
  }

  function construirModelo(J, INV, Cin) {
    const C = completarConfig(Cin);
    const D = J.dims;
    const canonU = raw => C.unidades.codigos[raw] || C.unidades.codigos[norm(raw)] || norm(raw) || 'UNI';
    const nombreU = cu => C.unidades.nombres[cu] || cu;
    const reCigUN = new RegExp(C.cigarros.un || '^$'), reCigLi = new RegExp(C.cigarros.linea || '^$');
    const noProd = new Set((C.noProducto || []).map(norm));
    const eq = { ...(J.equivAuto || {}), ...(C.equiv || {}) };

    /* Productos (ítems) de ventas + inventario */
    const items = [];
    const porClave = new Map();
    const nuevoItem = (k, d) => { const it = { k, mod: k[0] === 'M' ? 'M' : 'B', cod: k.slice(2), d, unH: '', unR: '', mm: '', unInv: '', ma: '', li: '', sl: '', inv: null, pres: new Map(), venta: 0, venta90: 0 }; porClave.set(k, items.length); items.push(it); return it; };
    J.items.k.forEach((k, j) => {
      const it = nuevoItem(k, J.items.d[j]);
      it.unH = D.un[J.items.unH[j]] || ''; it.unR = D.un[J.items.unR[j]] || ''; it.mm = D.un[(J.items.mm || [])[j]] || '';
      it.ma = D.ma[J.items.ma[j]] || ''; it.li = D.li[J.items.li[j]] || ''; it.sl = D.sl[J.items.sl[j]] || '';
      it.ult = J.items.dia[j];
    });
    for (const [i, u, f, n] of (J.pf || [])) {
      const cu = canonU(D.u[u]);
      const p = items[i].pres.get(cu);
      if (!p || n > p.n) items[i].pres.set(cu, { f, n: (p ? p.n : 0) + n });
      else p.n += n;
    }
    // Inventario
    const invF = INV && INV.filas ? INV.filas : null;
    if (invF) {
      for (let j = 0; j < invF.k.length; j++) {
        const k = invF.k[j];
        let ix = porClave.get(k);
        const it = ix == null ? nuevoItem(k, invF.d[j]) : items[ix];
        const cu = canonU(invF.u[j]);
        let fac = 1, fuente = 'unidad';
        if (!UNIDADES_SUELTAS.includes(cu)) {
          const p = it.pres.get(cu);
          if (p && p.f > 0) { fac = p.f; fuente = 'ventas'; }
          else {
            const fd = factorDescripcion(invF.d[j]);
            if (fd) { fac = fd.f; fuente = fd.fuente; } else fuente = 'por confirmar';
          }
        }
        const prev = it.inv;
        const reg = { s: invF.s[j], u: invF.u[j], cu, fac, fuenteFac: fuente, c: invF.c[j], p: invF.p[j], es: invF.es[j], d: invF.d[j] };
        if (prev) { prev.extra = (prev.extra || []).concat(reg); } else it.inv = reg;
        it.unInv = invF.un[j] || it.unInv;
        if (!it.li) it.li = invF.li[j] || '';
        if (!it.sl) it.sl = invF.sl[j] || '';
      }
    }
    // Razón típica costo/precio del inventario (para estimar costos inconsistentes o faltantes)
    let ratioCP = 0.78;
    if (invF) {
      const rs = [];
      for (let j = 0; j < invF.k.length; j++) { const c = invF.c[j], p = invF.p[j]; if (c > 0 && p > 0 && c / p > 0.3 && c / p < 1.5) rs.push(c / p); }
      if (rs.length > 10) { rs.sort((x, y) => x - y); ratioCP = rs[Math.floor(rs.length / 2)]; }
    }
    for (const it of items) {
      it.un = it.unH || (it.mm ? (eq[it.mm] || '') : '') || (it.unR ? (eq[it.unR] || it.unR) : '') || it.unInv || (it.mm ? it.mm : '') || 'SIN CLASIFICAR';
    }

    /* Hechos */
    const H = J.h, nH = H.d.length;
    const hItem = Int32Array.from(H.i), hDia = Int32Array.from(H.d), hU = new Int16Array(nH);
    const uCanon = D.u.map(canonU);
    const unidadesUsadas = [...new Set(uCanon)];
    const uIdx = new Map(unidadesUsadas.map((u, j) => [u, j]));
    for (let j = 0; j < nH; j++) hU[j] = uIdx.get(uCanon[H.u[j]]);
    const hVb = Float64Array.from(H.vb), hDv = Float64Array.from(H.dv), hQ = Float64Array.from(H.q), hQp = Float64Array.from(H.qp);
    const diaMax = nH ? Math.max(...hDia.slice(-500)) : diaDe(new Date());
    let diaMin = Infinity; for (let j = 0; j < nH; j++) if (hDia[j] < diaMin) diaMin = hDia[j];
    for (let j = 0; j < nH; j++) {
      const it = items[hItem[j]];
      const v = hVb[j] - hDv[j];
      it.venta += v;
      if (hDia[j] > diaMax - 90) { it.venta90 += v; if (hQ[j] > 0 && hVb[j] > 0) { it.pv = (it.pv || 0) + hVb[j]; it.pq = (it.pq || 0) + hQ[j]; } }
    }

    /* Agrupación */
    const AG = C.agrupador;
    const separar = new Set(AG.separar || []);
    const distintos = new Set((AG.distintos || []).map(p => p.slice().sort().join('&')));
    const uf = unionFind(items.length);
    const tok = items.map(it => tokens(it.d));
    const porCod = new Map();
    items.forEach((it, j) => {
      if (separar.has(it.k) || it.cod.startsWith('#')) return;
      const c = it.cod.replace(/~\d+$/, '');
      if (!porCod.has(c)) porCod.set(c, []);
      porCod.get(c).push(j);
    });
    for (const js of porCod.values()) {
      if (js.length < 2) continue;
      for (let a = 1; a < js.length; a++) {
        const x = items[js[0]], y = items[js[a]];
        if (x.mod === y.mod && x.k.replace(/~\d+$/, '') === y.k.replace(/~\d+$/, '')) continue;
        if (distintos.has([x.k, y.k].sort().join('&'))) continue;
        if (similitud(tok[js[0]], tok[js[a]]) >= 0.2 || x.un === y.un) uf.u(js[0], js[a]);
      }
    }
    const porDesc = new Map();
    items.forEach((it, j) => {
      if (separar.has(it.k)) return;
      const c = claveDesc(it.d);
      if (c.length < 4) return;
      if (porDesc.has(c)) { const o = porDesc.get(c); if (!distintos.has([items[o].k, it.k].sort().join('&'))) uf.u(o, j); }
      else porDesc.set(c, j);
    });
    for (const grupo of (AG.unir || [])) {
      const js = grupo.map(k => porClave.get(k)).filter(x => x != null);
      for (let a = 1; a < js.length; a++) uf.u(js[0], js[a]);
    }
    const raizG = new Map();
    const grupos = [];
    items.forEach((it, j) => {
      const r = uf.f(j);
      let g = raizG.get(r);
      if (g == null) { g = grupos.length; raizG.set(r, g); grupos.push({ id: '', miembros: [] }); }
      grupos[g].miembros.push(j);
      it.g = g;
    });
    const voto = (miembros, campo) => { const m = new Map(); for (const j of miembros) { const it = items[j]; const v = it[campo]; if (v && v !== 'SIN CLASIFICAR') m.set(v, (m.get(v) || 0) + Math.abs(it.venta) + 1); } return mayor(m); };
    for (const G of grupos) {
      const ms = G.miembros;
      ms.sort((a, b) => items[b].venta90 - items[a].venta90 || items[b].venta - items[a].venta);
      G.id = ms.map(j => items[j].k).sort()[0];
      let ov = null;
      for (const j of ms) { const o = AG.grupos && AG.grupos[items[j].k]; if (o) ov = { ...(ov || {}), ...o }; }
      G.ov = ov || {};
      const lider = items[ms[0]];
      G.desc = G.ov.d || lider.d;
      G.un = G.ov.un || voto(ms, 'un') || 'SIN CLASIFICAR';
      G.ma = G.ov.ma || voto(ms, 'ma') || '';
      G.li = G.ov.li || voto(ms, 'li') || '';
      G.sl = voto(ms, 'sl') || '';
      G.cig = reCigUN.test(G.un) || reCigLi.test(G.li);
      G.noProd = noProd.has(norm(G.un));
      G.mods = [...new Set(ms.map(j => items[j].mod))].sort().join('');
      // Presentaciones del grupo
      const pres = new Map();
      for (const j of ms) for (const [cu, p] of items[j].pres) { const q = pres.get(cu); if (!q || p.n > q.n) pres.set(cu, { f: p.f, n: (q ? q.n : 0) + p.n }); else q.n += p.n; }
      G.pres = pres;
      // Inventario del grupo (en unidades base)
      let stock = 0, valor = 0, valorP = 0, conInv = 0, factorInvMax = 1, porConfirmar = false;
      const regs = [];
      for (const j of ms) {
        const inv = items[j].inv;
        if (!inv) continue;
        for (const r of [inv, ...(inv.extra || [])]) {
          conInv++;
          const base = r.s * r.fac;
          stock += base;
          valor += r.s * r.c;
          valorP += r.s * r.p;
          if (r.fac > factorInvMax) factorInvMax = r.fac;
          if (r.fuenteFac === 'por confirmar') porConfirmar = true;
          regs.push({ k: items[j].k, ...r, base });
        }
      }
      G.inv = regs;
      G.conInv = conInv > 0;
      G.stock = stock;
      G.valorStock = valor;
      // Costo y precio por unidad base
      let cu = 0, pu = 0, wsum = 0;
      G.costoEstimado = false;
      for (const r of regs) {
        if (!(r.fac > 0)) continue;
        let c = r.c;
        if (r.p > 0 && (c <= 0 || c > r.p * 1.8)) { c = r.p * ratioCP; G.costoEstimado = true; }
        if (c > 0) { const w = Math.max(r.s, 0.0001); cu += (c / r.fac) * w; pu += ((r.p || c / ratioCP) / r.fac) * w; wsum += w; }
      }
      G.costoBase = wsum ? cu / wsum : 0;
      G.precioBase = wsum ? pu / wsum : 0;
      if (!G.costoBase) {
        let pv = 0, pq = 0;
        for (const j of ms) { pv += items[j].pv || 0; pq += items[j].pq || 0; }
        if (pq > 0) { G.precioBase = pv / pq; G.costoBase = G.precioBase * ratioCP; G.costoEstimado = true; }
      }
      G.facPorConfirmar = porConfirmar;
      // Unidad de compra
      let fc = 1, uc = 'UNI';
      if (G.ov.f > 0) { fc = +G.ov.f; uc = G.ov.u || 'CAJ'; }
      else if (factorInvMax > 1) { fc = factorInvMax; uc = (regs.find(r => r.fac === factorInvMax) || {}).cu || 'CAJ'; }
      else {
        let mejor = null;
        for (const [u, p] of pres) if (!UNIDADES_SUELTAS.includes(u) && p.f > 1 && (!mejor || p.n > mejor.n)) mejor = { u, ...p };
        if (mejor) { fc = mejor.f; uc = mejor.u; }
      }
      G.fc = fc; G.uc = uc;
      G.ucNombre = fc > 1 ? `${nombreU(uc)} ×${fc % 1 ? fc.toFixed(2) : fc}` : nombreU(uc === 'KG' ? 'KG' : 'UNI');
    }

    /* Marcas (unidades de negocio) */
    const marcasSet = new Set(grupos.map(g => g.un));
    const C_ = J.c;
    const unCli = C_.n.map(n => D.un[n] || '');
    for (const n of new Set(unCli)) if (n) marcasSet.add(eq[n] || n);
    const marcas = [...marcasSet].sort((a, b) => a.localeCompare(b, 'es'));
    const marcaIdx = new Map(marcas.map((m, j) => [m, j]));
    const gUn = Int32Array.from(grupos, g => marcaIdx.get(g.un));
    const marcaCig = marcas.map(m => reCigUN.test(m));
    for (const G of grupos) if (G.cig) marcaCig[marcaIdx.get(G.un)] = marcaCig[marcaIdx.get(G.un)] || reCigUN.test(G.un);

    const hG = new Int32Array(nH), hMod = new Uint8Array(nH), hCig = new Uint8Array(nH), hUn = new Int32Array(nH), hNoProd = new Uint8Array(nH);
    for (let j = 0; j < nH; j++) {
      const it = items[hItem[j]];
      const G = grupos[it.g];
      hG[j] = it.g; hMod[j] = it.mod === 'B' ? 1 : 0; hCig[j] = G.cig ? 1 : 0; hUn[j] = gUn[it.g]; hNoProd[j] = G.noProd ? 1 : 0;
    }
    // Clientes
    const nC = C_.d.length;
    const cDia = Int32Array.from(C_.d), cK = Int32Array.from(C_.k), cMod = Uint8Array.from(C_.m), cV = Float64Array.from(C_.v), cA = Int32Array.from(C_.a);
    const cUn = new Int32Array(nC), cCig = new Uint8Array(nC);
    for (let j = 0; j < nC; j++) { const n = unCli[j]; const m = eq[n] || n; const ix = marcaIdx.get(m); cUn[j] = ix == null ? -1 : ix; cCig[j] = reCigUN.test(m) ? 1 : 0; }

    // Días con facturación
    const diasSet = new Set();
    for (let j = 0; j < nH; j++) if (hVb[j] > 0) diasSet.add(hDia[j]);
    const diasVenta = [...diasSet].sort((a, b) => a - b);

    return {
      C, J, INV, items, grupos, marcas, ratioCP, marcaIdx, marcaCig, unidades: unidadesUsadas, nombreU, canonU,
      h: { n: nH, dia: hDia, item: hItem, g: hG, u: hU, vb: hVb, dv: hDv, q: hQ, qp: hQp, mod: hMod, cig: hCig, un: hUn, noProd: hNoProd },
      c: { n: nC, dia: cDia, k: cK, mod: cMod, v: cV, a: cA, un: cUn, cig: cCig, rif: J.cli.r, nombre: J.cli.n, asesores: D.as },
      diaMin: isFinite(diaMin) ? diaMin : diaMax, diaMax, diasVenta,
      capas: J.capas, invFecha: INV ? INV.fecha : null
    };
  }

  /* ---------------------------------------------------------------
     Períodos
     --------------------------------------------------------------- */
  const GRAN = { d: 'Día', s: 'Semana', m: 'Mes', t: 'Trimestre', a: 'Año' };
  function periodoDe(dia, gran) {
    const f = fechaDe(dia), y = f.getUTCFullYear(), m = f.getUTCMonth();
    if (gran === 'd') return dia;
    if (gran === 's') return dia - dowDe(dia);
    if (gran === 'm') return y * 12 + m;
    if (gran === 't') return y * 4 + Math.floor(m / 3);
    return y;
  }
  function etiquetaPeriodo(p, gran, corta) {
    if (gran === 'd') { const f = fechaDe(p); return `${DIAS_S[dowDe(p)]} ${f.getUTCDate()} ${MESES[f.getUTCMonth()]}`; }
    if (gran === 's') { const a = fechaDe(p), b = fechaDe(p + 6); return corta ? `${a.getUTCDate()} ${MESES[a.getUTCMonth()]}` : `${a.getUTCDate()} ${MESES[a.getUTCMonth()]} – ${b.getUTCDate()} ${MESES[b.getUTCMonth()]}`; }
    if (gran === 'm') { const y = Math.floor(p / 12), m = p % 12; return corta ? `${MESES[m]} ${String(y).slice(2)}` : `${MESES_L[m]} ${y}`; }
    if (gran === 't') return `T${(p % 4) + 1} ${Math.floor(p / 4)}`;
    return String(p);
  }
  function primerDiaPeriodo(p, gran) {
    if (gran === 'd' || gran === 's') return p;
    if (gran === 'm') return Math.floor(Date.UTC(Math.floor(p / 12), p % 12, 1) / DIA_MS);
    if (gran === 't') return Math.floor(Date.UTC(Math.floor(p / 4), (p % 4) * 3, 1) / DIA_MS);
    return Math.floor(Date.UTC(p, 0, 1) / DIA_MS);
  }
  function siguientePeriodo(p, gran) { return gran === 's' ? p + 7 : p + 1; }
  function periodosEntre(desde, hasta, gran) {
    const out = [];
    let p = periodoDe(desde, gran);
    const fin = periodoDe(hasta, gran);
    while (p <= fin && out.length < 2000) { out.push(p); p = siguientePeriodo(p, gran); }
    return out;
  }

  /* ---------------------------------------------------------------
     Analítico
     --------------------------------------------------------------- */
  function filtroHechos(M, f) {
    const h = M.h;
    const marcas = f.marcas && f.marcas.length ? new Set(f.marcas.map(m => M.marcaIdx.get(m))) : null;
    const grupo = f.grupo != null ? f.grupo : null;
    return j => {
      const d = h.dia[j];
      if (d < f.desde || d > f.hasta) return false;
      if (f.modulo === 'M' && h.mod[j] !== 0) return false;
      if (f.modulo === 'B' && h.mod[j] !== 1) return false;
      if (f.linea === 'base' && h.cig[j]) return false;
      if (f.linea === 'cig' && !h.cig[j]) return false;
      if (marcas && !marcas.has(h.un[j])) return false;
      if (grupo != null && h.g[j] !== grupo) return false;
      return true;
    };
  }
  function filtroClientes(M, f) {
    const c = M.c;
    const marcas = f.marcas && f.marcas.length ? new Set(f.marcas.map(m => M.marcaIdx.get(m))) : null;
    return j => {
      const d = c.dia[j];
      if (d < f.desde || d > f.hasta) return false;
      if (f.modulo === 'M' && c.mod[j] !== 0) return false;
      if (f.modulo === 'B' && c.mod[j] !== 1) return false;
      if (f.linea === 'base' && c.cig[j]) return false;
      if (f.linea === 'cig' && !c.cig[j]) return false;
      if (marcas && !marcas.has(c.un[j])) return false;
      return true;
    };
  }
  function diasVentaEntre(M, a, b) { let n = 0; for (const d of M.diasVenta) if (d >= a && d <= b) n++; return n; }

  function resumen(M, f) {
    const h = M.h, ok = filtroHechos(M, f);
    let vb = 0, dv = 0, q = 0;
    const gs = new Set();
    for (let j = 0; j < h.n; j++) {
      if (!ok(j)) continue;
      vb += h.vb[j]; dv += h.dv[j];
      if (!h.noProd[j]) { q += h.q[j]; if (h.vb[j] > 0) gs.add(h.g[j]); }
    }
    const c = M.c, okc = filtroClientes(M, f);
    const clientes = new Map();
    for (let j = 0; j < c.n; j++) { if (!okc(j)) continue; clientes.set(c.k[j], (clientes.get(c.k[j]) || 0) + c.v[j]); }
    let atendidos = 0; for (const v of clientes.values()) if (v > 0) atendidos++;
    const dias = diasVentaEntre(M, f.desde, f.hasta);
    const neta = vb - dv;
    return { vb, dv, neta, q, skus: gs.size, clientes: atendidos, ticket: atendidos ? neta / atendidos : 0, dias, promDia: dias ? neta / dias : 0, pDev: vb ? dv / vb : 0 };
  }
  function rangoAnterior(f) { const n = f.hasta - f.desde + 1; return { ...f, desde: f.desde - n, hasta: f.desde - 1 }; }

  function serie(M, f, gran, partir = 'mod') {
    const h = M.h, ok = filtroHechos(M, f);
    const per = periodosEntre(f.desde, f.hasta, gran);
    const ix = new Map(per.map((p, j) => [p, j]));
    const L = per.length;
    const mv = new Float64Array(L), b = new Float64Array(L), qm = new Float64Array(L), qb = new Float64Array(L), dv = new Float64Array(L);
    for (let j = 0; j < h.n; j++) {
      if (!ok(j)) continue;
      const k = ix.get(periodoDe(h.dia[j], gran));
      if (k == null) continue;
      const v = h.vb[j] - h.dv[j];
      dv[k] += h.dv[j];
      if (h.mod[j]) { b[k] += v; if (!h.noProd[j]) qb[k] += h.q[j]; } else { mv[k] += v; if (!h.noProd[j]) qm[k] += h.q[j]; }
    }
    void partir;
    const dias = per.map(p => { const a = Math.max(primerDiaPeriodo(p, gran), f.desde); const z = Math.min(primerDiaPeriodo(siguientePeriodo(p, gran), gran) - 1, f.hasta); return diasVentaEntre(M, a, z); });
    return { periodos: per, etiquetas: per.map(p => etiquetaPeriodo(p, gran, true)), largas: per.map(p => etiquetaPeriodo(p, gran, false)), mv: [...mv], b: [...b], total: [...mv].map((x, j) => x + b[j]), qm: [...qm], qb: [...qb], q: [...qm].map((x, j) => x + qb[j]), dv: [...dv], dias };
  }

  function porDimension(M, f, dim, medida = 'v') {
    // dim: 'marca' | 'grupo' | 'pres' | 'dow'
    const h = M.h, ok = filtroHechos(M, f);
    const acc = new Map();
    for (let j = 0; j < h.n; j++) {
      if (!ok(j)) continue;
      if ((dim === 'grupo' || dim === 'pres') && h.noProd[j]) continue;
      const k = dim === 'marca' ? h.un[j] : dim === 'grupo' ? h.g[j] : dim === 'pres' ? h.u[j] : dowDe(h.dia[j]);
      let a = acc.get(k);
      if (!a) { a = { k, v: 0, vb: 0, dv: 0, q: 0, qp: 0 }; acc.set(k, a); }
      a.v += h.vb[j] - h.dv[j]; a.vb += h.vb[j]; a.dv += h.dv[j]; a.q += h.q[j]; a.qp += h.qp[j];
    }
    const out = [...acc.values()];
    out.sort((x, y) => (medida === 'q' ? y.q - x.q : y.v - x.v));
    return out;
  }

  function matriz(M, f, gran, dim, limite = 40) {
    const h = M.h, ok = filtroHechos(M, f);
    const per = periodosEntre(f.desde, f.hasta, gran);
    const ix = new Map(per.map((p, j) => [p, j]));
    const filas = new Map();
    for (let j = 0; j < h.n; j++) {
      if (!ok(j)) continue;
      if (dim === 'grupo' && h.noProd[j]) continue;
      const k = dim === 'marca' ? h.un[j] : h.g[j];
      const c = ix.get(periodoDe(h.dia[j], gran));
      if (c == null) continue;
      let fila = filas.get(k);
      if (!fila) { fila = { k, v: new Float64Array(per.length), q: new Float64Array(per.length), tot: 0, totQ: 0 }; filas.set(k, fila); }
      const v = h.vb[j] - h.dv[j];
      fila.v[c] += v; fila.tot += v;
      if (!h.noProd[j]) { fila.q[c] += h.q[j]; fila.totQ += h.q[j]; }
    }
    const out = [...filas.values()].sort((a, b) => b.tot - a.tot).slice(0, limite);
    const dias = per.map(p => { const a = Math.max(primerDiaPeriodo(p, gran), f.desde); const z = Math.min(primerDiaPeriodo(siguientePeriodo(p, gran), gran) - 1, f.hasta); return diasVentaEntre(M, a, z); });
    return { periodos: per, etiquetas: per.map(p => etiquetaPeriodo(p, gran, true)), filas: out, dias };
  }

  function clientesTop(M, f, limite = 15) {
    const c = M.c, ok = filtroClientes(M, f);
    const acc = new Map();
    for (let j = 0; j < c.n; j++) {
      if (!ok(j)) continue;
      let a = acc.get(c.k[j]);
      if (!a) { a = { k: c.k[j], v: 0, dias: new Set(), a: c.a[j] }; acc.set(c.k[j], a); }
      a.v += c.v[j]; a.dias.add(c.dia[j]); a.a = c.a[j];
    }
    return [...acc.values()].sort((x, y) => y.v - x.v).slice(0, limite).map(x => ({ rif: M.c.rif[x.k], nombre: M.c.nombre[x.k], v: x.v, visitas: x.dias.size, asesor: M.c.asesores[x.a] || '' }));
  }
  function asesores(M, f) {
    const c = M.c, ok = filtroClientes(M, f);
    const acc = new Map();
    for (let j = 0; j < c.n; j++) {
      if (!ok(j)) continue;
      let a = acc.get(c.a[j]);
      if (!a) { a = { a: c.a[j], v: 0, cli: new Set() }; acc.set(c.a[j], a); }
      a.v += c.v[j]; if (c.v[j] > 0) a.cli.add(c.k[j]);
    }
    return [...acc.values()].map(x => ({ nombre: M.c.asesores[x.a] || 'SIN ASESOR', v: x.v, clientes: x.cli.size })).sort((x, y) => y.v - x.v);
  }

  /* ---------------------------------------------------------------
     Sugerido de compra por marca (unidad de negocio)
     --------------------------------------------------------------- */
  function parametrosMarca(M, un) {
    const S = M.C.sugerido, P = (M.C.marcas || {})[un] || {};
    return {
      ventana: +P.ventana || S.ventana, reciente: S.reciente, pesoReciente: S.pesoReciente,
      cobertura: P.cobertura != null && P.cobertura !== '' ? +P.cobertura : S.cobertura,
      seguridad: P.seguridad != null && P.seguridad !== '' ? +P.seguridad : S.seguridad,
      exceso: S.exceso, redondeo: S.redondeo, presupuesto: +P.presupuesto || 0
    };
  }

  function calcularSugerido(M, opciones = {}) {
    const h = M.h;
    const ref = opciones.diaRef != null ? opciones.diaRef : M.diaMax;
    const S = M.C.sugerido;
    const vMax = Math.max(S.ventana, 180);
    const q90 = new Map(), q30 = new Map(), v30 = new Map(), qMeses = new Map(), ultVenta = new Map();
    const mesRef = periodoDe(ref, 'm');
    for (let j = 0; j < h.n; j++) {
      const d = h.dia[j];
      if (d > ref || h.noProd[j]) continue;
      const g = h.g[j];
      if (h.vb[j] > 0 && (ultVenta.get(g) ?? -1) < d) ultVenta.set(g, d);
      if (d <= ref - vMax) continue;
      const un = M.grupos[g].un;
      const P = parametrosMarca(M, un);
      if (d > ref - P.ventana) q90.set(g, (q90.get(g) || 0) + h.q[j]);
      if (d > ref - P.reciente) { q30.set(g, (q30.get(g) || 0) + h.q[j]); v30.set(g, (v30.get(g) || 0) + h.vb[j] - h.dv[j]); }
      const mm = mesRef - periodoDe(d, 'm');
      if (mm >= 0 && mm < 6) { const a = qMeses.get(g) || new Float64Array(6); a[5 - mm] += h.q[j]; qMeses.set(g, a); }
    }
    const diasCache = new Map();
    const diasV = n => { if (!diasCache.has(n)) diasCache.set(n, Math.max(1, diasVentaEntre(M, ref - n + 1, ref))); return diasCache.get(n); };
    const diasMes = (() => { const meses = new Map(); for (const d of M.diasVenta) { const k = periodoDe(d, 'm'); meses.set(k, (meses.get(k) || 0) + 1); } const v = [...meses.values()]; v.pop(); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : 21; })();

    const porMarca = new Map();
    M.grupos.forEach((G, g) => {
      if (G.noProd) return;
      const P = parametrosMarca(M, G.un);
      const qW = Math.max(0, q90.get(g) || 0), qR = Math.max(0, q30.get(g) || 0);
      const vpdW = qW / diasV(P.ventana), vpdR = qR / diasV(P.reciente);
      const vpd = P.pesoReciente * vpdR + (1 - P.pesoReciente) * vpdW;
      if (!G.conInv && vpd <= 0) return;
      const stock = Math.max(0, G.stock);
      const fc = G.fc || 1;
      const ultV = ultVenta.get(g);
      const inactivo = !G.conInv && (ultV == null || ultV <= ref - P.reciente);
      const objetivo = vpd * (P.cobertura + P.seguridad);
      const necesidad = Math.max(0, objetivo - stock);
      const red = P.redondeo;
      const sug = inactivo ? 0 : (necesidad > 0 ? Math.max(0, Math.ceil(necesidad / fc - red)) : 0);
      const minimoBase = Math.max(0, vpd * P.seguridad - stock);
      const minimo = inactivo ? 0 : (minimoBase > 0 ? Math.max(1, Math.ceil(minimoBase / fc - red)) : 0);
      const cobertura = vpd > 0 ? stock / vpd : (stock > 0 ? Infinity : 0);
      const costoPack = G.costoBase * fc;
      const ult = ultVenta.get(g);
      let estado = 'ok';
      if (inactivo) estado = 'inactivo';
      else if (vpd > 0 && stock <= 0) estado = 'quiebre';
      else if (vpd > 0 && cobertura < P.seguridad) estado = 'riesgo';
      else if (vpd <= 0 && stock > 0) estado = 'sinmov';
      else if (cobertura > P.exceso) estado = 'exceso';
      const fila = {
        g, id: G.id, desc: G.desc, un: G.un, ma: G.ma, li: G.li, mods: G.mods, codigos: G.miembros.map(j => M.items[j].k),
        fc, uc: G.uc, ucNombre: G.ucNombre, vpd, vpdW, vpdR, tendencia: vpdW > 0 ? vpdR / vpdW - 1 : (vpdR > 0 ? 1 : 0),
        promMes: vpd * diasMes / fc, stock: stock / fc, stockBase: stock, cobertura, objetivo: objetivo / fc,
        algoritmo: sug, minimo, costo: costoPack, precio: G.precioBase * fc, venta30: v30.get(g) || 0,
        meses: qMeses.get(g) ? [...qMeses.get(g)].map(x => x / fc) : [0, 0, 0, 0, 0, 0],
        estado, conInv: G.conInv, facPorConfirmar: G.facPorConfirmar, costoEstimado: G.costoEstimado, ultimaVenta: ult != null ? ult : null, P
      };
      if (!porMarca.has(G.un)) porMarca.set(G.un, []);
      porMarca.get(G.un).push(fila);
    });
    const marcas = [];
    for (const [un, filas] of porMarca) {
      filas.sort((a, b) => (b.algoritmo * b.costo) - (a.algoritmo * a.costo) || b.venta30 - a.venta30);
      const P = parametrosMarca(M, un);
      const tot = { skus: filas.filter(r => r.estado !== 'inactivo').length, inactivos: 0, conMov: 0, venta30: 0, stockValor: 0, sugeridoValor: 0, quiebres: 0, riesgo: 0, exceso: 0, sinmov: 0, consumoValor: 0, sinInv: 0, porConfirmar: 0 };
      for (const r of filas) {
        if (r.estado === 'inactivo') { tot.inactivos++; continue; }
        if (r.vpd > 0) tot.conMov++;
        tot.venta30 += r.venta30;
        tot.stockValor += r.stock * r.costo;
        tot.sugeridoValor += r.algoritmo * r.costo;
        tot.consumoValor += r.vpd / r.fc * r.costo;
        if (r.estado === 'quiebre') tot.quiebres++;
        if (r.estado === 'riesgo') tot.riesgo++;
        if (r.estado === 'exceso') tot.exceso++;
        if (r.estado === 'sinmov') tot.sinmov++;
        if (!r.conInv) tot.sinInv++;
        if (r.facPorConfirmar) tot.porConfirmar++;
      }
      tot.cobertura = tot.consumoValor > 0 ? tot.stockValor / tot.consumoValor : null;
      marcas.push({ un, filas, tot, P, cig: M.marcaCig[M.marcaIdx.get(un)] });
    }
    marcas.sort((a, b) => b.tot.sugeridoValor - a.tot.sugeridoValor || b.tot.venta30 - a.tot.venta30);
    return { ref, diasMes, marcas };
  }

  function tendenciaMarca(M, un, meses = 6, ref = M.diaMax) {
    const h = M.h, ix = M.marcaIdx.get(un);
    const mRef = periodoDe(ref, 'm');
    const v = new Float64Array(meses), q = new Float64Array(meses);
    for (let j = 0; j < h.n; j++) {
      if (h.un[j] !== ix) continue;
      const k = mRef - periodoDe(h.dia[j], 'm');
      if (k < 0 || k >= meses) continue;
      v[meses - 1 - k] += h.vb[j] - h.dv[j];
      if (!h.noProd[j]) q[meses - 1 - k] += h.q[j];
    }
    const etiquetas = [];
    for (let k = meses - 1; k >= 0; k--) etiquetas.push(etiquetaPeriodo(mRef - k, 'm', true));
    return { etiquetas, v: [...v], q: [...q] };
  }

  /* Sugerencias de agrupación (para el panel administrativo). */
  function sugerenciasAgrupador(M, limite = 120) {
    const out = [];
    const distintos = new Set((M.C.agrupador.distintos || []).map(p => p.slice().sort().join('&')));
    const porMarca = new Map();
    M.grupos.forEach((G, g) => { if (G.noProd) return; const k = G.un; if (!porMarca.has(k)) porMarca.set(k, []); porMarca.get(k).push(g); });
    const tok = M.grupos.map(G => tokens(G.desc));
    for (const [un, gs] of porMarca) {
      for (let a = 0; a < gs.length; a++) for (let b = a + 1; b < gs.length; b++) {
        const A = M.grupos[gs[a]], B = M.grupos[gs[b]];
        const s = similitud(tok[gs[a]], tok[gs[b]]);
        if (s < 0.6) continue;
        const ka = A.miembros.map(j => M.items[j].k), kb = B.miembros.map(j => M.items[j].k);
        if (ka.some(x => kb.some(y => distintos.has([x, y].sort().join('&'))))) continue;
        const venta = M.items[A.miembros[0]].venta + M.items[B.miembros[0]].venta;
        out.push({ a: gs[a], b: gs[b], s, un, venta, cruzaModulos: A.mods !== B.mods });
      }
      void un;
    }
    out.sort((x, y) => (y.cruzaModulos - x.cruzaModulos) || y.s - x.s || y.venta - x.venta);
    return out.slice(0, limite);
  }

  /* ---------------------------------------------------------------
     Formato
     --------------------------------------------------------------- */
  const fmtN = (n, dec = 0) => (n == null || !isFinite(n)) ? '—' : new Intl.NumberFormat('es-VE', { minimumFractionDigits: dec, maximumFractionDigits: dec }).format(n);
  const fmtUSD = (n, dec) => {
    if (n == null || !isFinite(n)) return '—';
    const d = dec != null ? dec : (Math.abs(n) >= 1000 ? 0 : 2);
    return (n < 0 ? '−$' : '$') + fmtN(Math.abs(n), d);
  };
  const fmtCorto = n => {
    if (n == null || !isFinite(n)) return '—';
    const a = Math.abs(n), s = n < 0 ? '−' : '';
    if (a >= 1e6) return s + '$' + fmtN(a / 1e6, a >= 1e7 ? 1 : 2) + ' M';
    if (a >= 1e4) return s + '$' + fmtN(a / 1e3, 0) + ' mil';
    return s + '$' + fmtN(a, a >= 100 ? 0 : 2);
  };
  const fmtPct = (n, dec = 1) => (n == null || !isFinite(n)) ? '—' : (n > 0 ? '+' : n < 0 ? '−' : '') + fmtN(Math.abs(n) * 100, dec) + ' %';
  const fmtFecha = d => { if (d == null) return '—'; const f = fechaDe(d); return `${f.getUTCDate()} ${MESES[f.getUTCMonth()]} ${f.getUTCFullYear()}`; };
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  /* ---------------------------------------------------------------
     Carga de datos publicados (y vista previa local del panel)
     --------------------------------------------------------------- */
  const IDB = { nombre: 'vjs_datos', almacen: 'archivos' };
  function idb() {
    return new Promise((ok, mal) => {
      if (!raiz.indexedDB) return mal(new Error('Sin IndexedDB'));
      const r = raiz.indexedDB.open(IDB.nombre, 1);
      r.onupgradeneeded = () => r.result.createObjectStore(IDB.almacen);
      r.onsuccess = () => ok(r.result);
      r.onerror = () => mal(r.error);
    });
  }
  async function guardarLocal(clave, valor) {
    const db = await idb();
    return new Promise((ok, mal) => { const t = db.transaction(IDB.almacen, 'readwrite'); t.objectStore(IDB.almacen).put(valor, clave); t.oncomplete = () => ok(true); t.onerror = () => mal(t.error); });
  }
  async function leerLocal(clave) {
    try {
      const db = await idb();
      return await new Promise((ok) => { const t = db.transaction(IDB.almacen, 'readonly'); const r = t.objectStore(IDB.almacen).get(clave); r.onsuccess = () => ok(r.result || null); r.onerror = () => ok(null); });
    } catch (e) { return null; }
  }
  async function borrarLocal(clave) {
    try { const db = await idb(); await new Promise(ok => { const t = db.transaction(IDB.almacen, 'readwrite'); t.objectStore(IDB.almacen).delete(clave); t.oncomplete = () => ok(); t.onerror = () => ok(); }); } catch (e) { /* nada */ }
  }
  async function leerPublicado(nombre) {
    try {
      const r = await fetch(`datos/${nombre}.json?v=${Date.now()}`, { cache: 'no-store' });
      if (!r.ok) return null;
      return await r.json();
    } catch (e) { return null; }
  }
  /* Usa la versión local (vista previa del panel) si es más reciente que la publicada. */
  async function cargarDatos() {
    const res = {};
    let local = false;
    for (const n of ['ventas', 'inventario', 'config']) {
      const [pub, loc] = await Promise.all([leerPublicado(n), leerLocal(n)]);
      const tPub = pub && (pub.generado || pub.actualizado) || '';
      const tLoc = loc && (loc.generado || loc.actualizado) || '';
      if (loc && (!pub || tLoc > tPub)) { res[n] = loc; local = true; } else res[n] = pub;
    }
    res.local = local;
    return res;
  }

  return {
    // utilidades
    norm, num, codigo, rifNorm, diaDe, isoDe, diaDeIso, fechaDe, dowDe, MESES, MESES_L, DIAS_S, r2, pausa,
    // lectura y procesamiento
    abrirLibro, detectarTipo, procesarArchivo, TIPOS,
    vacioVentas, decodificarVentas, codificarVentas, fusionarCapa, codificarInventario, equivalenciasAuto,
    // modelo y cálculos
    CONFIG_BASE, completarConfig, construirModelo, tokens, claveDesc, similitud, factorDescripcion,
    GRAN, periodoDe, etiquetaPeriodo, primerDiaPeriodo, periodosEntre, diasVentaEntre,
    resumen, rangoAnterior, serie, porDimension, matriz, clientesTop, asesores,
    parametrosMarca, calcularSugerido, tendenciaMarca, sugerenciasAgrupador,
    // formato y datos
    fmtN, fmtUSD, fmtCorto, fmtPct, fmtFecha, esc,
    guardarLocal, leerLocal, borrarLocal, leerPublicado, cargarDatos
  };
});
