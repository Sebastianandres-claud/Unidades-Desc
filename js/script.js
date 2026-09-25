/* ==========================================================================
   SITRANS DASHBOARD - OPERATIONAL LOGIC & REALTIME ENGINE
   ========================================================================== */

const RAW_URL = "https://github-live-proxy.samiranda.workers.dev";
const REFRESH_MS = 10000;
const CODIGOS_EXCLUIDOS = ["TIP", "TQ", "OUT"];
const DIAS = { LU:0, MA:1, MI:2, JU:3, VI:4, SA:5, DO:6 };

let ultimasFilas = [];

let sortDescGeneral = true;
let sortDescSinConexion = true;
let sortDescEmbarque = true;

// ==========================================
// EXPORTACIÓN A EXCEL NATIVA
// ==========================================

function exportarExcel() {
  if (!ultimasFilas || ultimasFilas.length === 0) {
    alert('Todavía no hay datos cargados para exportar.');
    return;
  }

  const panels = document.querySelectorAll('.panel');
  if (!panels || panels.length === 0) return;

  let excelHTML = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" 
          xmlns:x="urn:schemas-microsoft-com:office:excel" 
          xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="UTF-8">
      <style>
        table { border-collapse: collapse; margin-bottom: 20px; font-family: sans-serif; }
        th { background-color: #FAFAFA; color: #5B6670; font-weight: bold; border: 1px solid #DFE3E0; padding: 6px; }
        td { border: 1px solid #EEF0EE; padding: 6px; }
        h3 { font-family: sans-serif; color: #10161C; margin-top: 15px; }
      </style>
    </head>
    <body>
  `;

  panels.forEach((panel) => {
    const title = panel.querySelector('.panel-head h2')?.innerText.trim() || 'Tabla';
    const table = panel.querySelector('table');
    if (table) {
      excelHTML += `<h3>${title}</h3>`;
      excelHTML += table.outerHTML;
      excelHTML += `<br>`;
    }
  });

  excelHTML += `</body></html>`;

  const blob = new Blob(['\ufeff' + excelHTML], {
    type: 'application/vnd.ms-excel;charset=utf-8'
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const fecha = new Date().toISOString().slice(0, 10);
  
  a.href = url;
  a.download = `Reporte_SITRANS_${fecha}.xls`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ==========================================
// LÓGICA DE BADGES Y REMARKS (AC / COT / INSTANCIAS)
// ==========================================

// Generador de cápsula según la Instancia de la unidad
function obtenerBadgeInstancia(instancia) {
  if (!instancia) return `<span class="badge badge-calle">-</span>`;
  const inst = String(instancia).toUpperCase();
  let clase = 'badge-calle';

  if (inst.includes('DESPACHO')) {
    clase = 'badge-despacho';
  } else if (inst.includes('MOVIMIENTO')) {
    clase = 'badge-movimiento';
  } else if (inst.includes('BORDO') || inst.includes('EMBARCADO')) {
    clase = 'badge-bordo';
  } else if (inst.includes('EMBARQUE')) {
    clase = 'badge-embarque';
  } else if (inst.includes('DESCARGA')) {
    clase = 'badge-descarga';
  } else if (inst.includes('CALLE')) {
    clase = 'badge-calle';
  }

  return `<span class="badge ${clase}">${esc(instancia)}</span>`;
}

function obtenerBadgeMov(esMov) {
  if (esMov) return `<span class="badge badge-mov-si">SÍ</span>`;
  return `<span class="badge badge-mov-no">NO</span>`;
}

// EVALUACIÓN DE AC Y USDA/COT
function obtenerBadgesHTML(remarks) {
  if (!remarks) return '';
  const rem = String(remarks).trim();
  let html = '';

// EVALUACIÓN DE AC Y USDA/COT
function obtenerBadgesHTML(remarks) {
  if (!remarks) return '';
  const rem = String(remarks).trim();
  let html = '';

  // 1. REGLA USDA -> Genera cuadrito "COT"
  if (/\bUSDA\b/i.test(rem)) {
    html += `<span class="badge-cot">COT</span>`;
  }

  // 2. REGLA AC -> Se activa OBLIGATORIAMENTE si existe una fracción numérica (ej. 6/4, 5/2, 10-2)
  // Ignora si solo dice "AC" sin números.
  const tieneFraccionNum = /\b\d+[\/\.-]\d+\b/.test(rem);

  if (tieneFraccionNum) {
    html += `<span class="badge-ac">AC</span>`;
  }

  return html;
}

function obtenerClaseFila(instancia) {
  const inst = String(instancia || '').toUpperCase();
  if (inst.includes('CALLE') || inst.includes('DESCARGA') || inst.includes('SOBRE CAMION')) return 'row-calle';
  if (inst.includes('MOVIMIENTO') || inst.includes('DESPACHO')) return 'row-movimiento';
  if (inst.includes('EMBARCADO') || inst.includes('ABORDO')) return 'row-a-bordo';
  if (inst.includes('EMBARQUE')) return 'row-embarque';
  return '';
}

function formatPosicion(pos) {
  if (pos === null || pos === undefined) return '-';
  const txt = String(pos).trim();
  if (!txt) return '-';
  const partes = txt.split('*').map(p => p.trim()).filter(p => p.length > 0);
  if (partes.length === 0) return txt;
  return partes[partes.length - 1];
}

function claseTiempo(minutos) {
  if (minutos === null || minutos === undefined) return '';
  if (minutos >= 26) return 'style="color:#e11d48; font-weight:700;"';
  if (minutos < 15) return 'style="color:#059669; font-weight:600;"';
  return 'style="color:#d97706; font-weight:600;"';
}

function promedio(arr) {
  if (arr.length === 0) return '0.0';
  const suma = arr.reduce((acc, r) => acc + (r.TiempoMin ?? 0), 0);
  return (suma / arr.length).toFixed(1);
}

function esc(v) {
  return String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function esVacio(v) { return v === null || v === undefined || String(v).trim() === ''; }

// ==========================================
// PARSERS Y CÁLCULOS
// ==========================================

function esPosicionVigente(pos, anioActual) {
  const limpio = String(pos ?? '').trim();
  const basura = CODIGOS_EXCLUIDOS.includes(limpio.toUpperCase());
  const larga = limpio.length > 10;
  let anioRot = null;
  const prefijo = parseInt(limpio.substring(0, 2), 10);
  if (!isNaN(prefijo)) anioRot = 2000 + prefijo;

  if (basura) return false;
  if (!larga) return true;
  return anioRot === anioActual || anioRot === anioActual - 1;
}

function mondayIndex(date) { return (date.getDay() + 6) % 7; }

function parseDiscon(valor) {
  if (!valor) return null;
  const txt = String(valor).trim();
  if (!txt) return null;

  if (/^\d{10}$/.test(txt)) {
    const yy = parseInt(txt.substring(0, 2), 10);
    const mm = parseInt(txt.substring(2, 4), 10);
    const dd = parseInt(txt.substring(4, 6), 10);
    const hh = parseInt(txt.substring(6, 8), 10);
    const mi = parseInt(txt.substring(8, 10), 10);
    if (![yy, mm, dd, hh, mi].some(isNaN)) {
      return new Date(2000 + yy, mm - 1, dd, hh, mi, 0);
    }
  }

  const parsed = Date.parse(txt);
  if (!isNaN(parsed)) return new Date(parsed);
  return null;
}

function parseHoraSimple(valor, ahora) {
  const txt = String(valor ?? '').trim();
  if (txt.length < 6) return null;
  const codigoDia = txt.substring(0, 2).toUpperCase();
  const sufijo = txt.substring(txt.length - 4);
  const hh = parseInt(sufijo.substring(0, 2), 10);
  const mi = parseInt(sufijo.substring(2, 4), 10);
  const diaObjetivo = DIAS.hasOwnProperty(codigoDia) ? DIAS[codigoDia] : null;

  const valido = diaObjetivo !== null && !isNaN(hh) && !isNaN(mi) && hh >= 0 && hh <= 23 && mi >= 0 && mi <= 59;
  if (!valido) return null;

  const hoy = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate());
  let offsetMatch = 0;
  for (let o = 0; o <= 6; o++) {
    const candidatoDia = new Date(hoy);
    candidatoDia.setDate(hoy.getDate() - o);
    if (mondayIndex(candidatoDia) === diaObjetivo) { offsetMatch = o; break; }
  }
  const fechaBase = new Date(hoy);
  fechaBase.setDate(hoy.getDate() - offsetMatch);
  let candidato = new Date(fechaBase.getFullYear(), fechaBase.getMonth(), fechaBase.getDate(), hh, mi, 0);
  if (candidato > ahora) {
    candidato = new Date(candidato.getTime() - 7 * 24 * 60 * 60 * 1000);
  }
  return candidato;
}

function formatearHHMM(diffMs) {
  if (diffMs === null || diffMs === undefined) return null;
  const totalMin = Math.round(diffMs / 60000);
  const horas = Math.floor(totalMin / 60);
  const minutos = totalMin % 60;
  return String(horas).padStart(2, '0') + ':' + String(minutos).padStart(2, '0');
}

function extraerFechaReporte(text) {
  if (!text) return null;
  const lines = text.replace(/\r/g, '').split('\n').map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length === 0) return null;

  for (let i = lines.length - 1; i >= 0; i--) {
    const line = lines[i];
    const matchTSV = line.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})[\t\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (matchTSV) {
      let yr = Number(matchTSV[3]);
      if (yr < 100) yr += 2000;
      const sec = matchTSV[6] ? Number(matchTSV[6]) : 0;
      return new Date(yr, Number(matchTSV[2]) - 1, Number(matchTSV[1]), Number(matchTSV[4]), Number(matchTSV[5]), sec);
    }
    if (/DATOS\s+AL|OPERACI[OÓ]N/i.test(line)) {
      const matchTexto = line.match(/\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\s+\d{1,2}:\d{2}(:\d{2})?(\s*[ap]\.?\s*m\.?)?/i);
      if (matchTexto) return "Datos al " + matchTexto[0].toUpperCase();
    }
  }
  return null;
}

function parseTSV(text) {
  const cleanText = text.replace(/^\ufeff/, '');
  const lines = cleanText.replace(/\r/g, '').split('\n');
  if (lines.length === 0) return [];

  let headerIndex = lines.findIndex(l => l.includes('Container No.') || l.includes('Container'));
  if (headerIndex === -1) headerIndex = 0;

  const headers = lines[headerIndex].split('\t').map(h => h.trim());
  const rows = [];
  for (let i = headerIndex + 1; i < lines.length; i++) {
    const line = lines[i];
    if (!line || !line.includes('\t')) continue;
    const cells = line.split('\t');
    const row = {};
    headers.forEach((h, idx) => row[h] = (cells[idx] ?? '').trim());
    rows.push(row);
  }
  return rows;
}

function calcularFila(fila, ahora, anioActual) {
  if (!fila) return null;

  const powerTxt = String(fila['Power'] ?? '');
  const seMovio = powerTxt.includes('(') ? 'Sí' : 'No';
  const powerNum = parseInt(powerTxt.replace(/[()]/g, '').trim(), 10) || 0;

  const posActual = String(fila['Current Position'] ?? '').trim();
  const discon = fila['Discon'];
  const remarks = fila['Remarks'];
  const yardPlan = fila['Yard Plan'];
  const kind = String(fila['Kind'] ?? '').trim();
  const moveStage = String(fila['Move Stage'] ?? '').trim();

  let instancia;

  if (posActual.length > 10) {
    if (kind === 'DSCH') instancia = 'DESCARGA';
    else if (kind === 'LOAD') instancia = 'EMBARCADO SIN CONEX';
    else instancia = 'EMBARCADO SIN CONEXION (revisar)';
  } else {
    if (kind === 'DSCH') {
      instancia = 'DESCARGA';
    } else if (esVacio(discon)) {
      instancia = esVacio(remarks) ? 'CALLE' : 'CALLE REVISAR';
    } else {
      if (!esVacio(yardPlan)) instancia = 'MOVIMIENTO';
      else if (moveStage === 'Completed') instancia = 'MOVIMIENTO';
      else {
        if (powerNum === 0 && kind === 'LOAD') instancia = 'EMBARQUE';
        else if (powerNum === 0) instancia = 'DESPACHO';
        else instancia = 'SIN CLASIFICAR';
      }
    }
  }

  if (posActual.startsWith('*TR')) {
    if (instancia === 'EMBARQUE') instancia = 'EMBARQUE SOBRE CAMION';
    else if (instancia === 'MOVIMIENTO') instancia = 'MOVIMIENTO SOBRE CAMION';
    else instancia = 'SOBRE CAMION';
  }

  const usaDiscon = ['MOVIMIENTO', 'EMBARQUE', 'DESPACHO', 'SOBRE CAMION', 'SIN CLASIFICAR', 'DESCARGA', 'MOVIMIENTO SOBRE CAMION', 'EMBARQUE SOBRE CAMION'].includes(instancia);
  const usaYardIn = ['CALLE', 'CALLE REVISAR'].includes(instancia);
  const usaComplete = ['EMBARCADO SIN CONEX', 'EMBARCADO SIN CONEXION (revisar)'].includes(instancia);

  let fechaEvento = null;
  if (usaDiscon) fechaEvento = parseDiscon(discon);
  else if (usaYardIn) fechaEvento = parseHoraSimple(fila['Yard In'], ahora);
  else if (usaComplete) fechaEvento = parseHoraSimple(fila['Complete'], ahora);

  const tiempoTranscurrido = fechaEvento ? formatearHHMM(ahora - fechaEvento) : null;
  const tiempoMin = fechaEvento ? Math.round((ahora - fechaEvento) / 60000) : null;

  const esMov = (moveStage === 'Completed') && !esVacio(discon) && (kind === 'YARD') && (powerTxt.trim() === '(0)');

  return {
    Contenedor: fila['Container No.'],
    Nave: fila['Outbound Carrier Name'],
    Posicion: posActual,
    Instancia: instancia,
    Tiempo: tiempoTranscurrido,
    TiempoMin: tiempoMin,
    ReqsPower: seMovio,
    Mov: esMov,
    Remarks: remarks
  };
}

function procesar(text) {
  const ahora = new Date();
  const anioActual = ahora.getFullYear();
  const rows = parseTSV(text);

  const filasValidas = rows.filter(r => r && !esVacio(r['Container No.']));
  const sinBasura = filasValidas.filter(r => esPosicionVigente(r['Current Position'], anioActual));
  const calculadas = sinBasura.map(r => calcularFila(r, ahora, anioActual)).filter(Boolean);
  return calculadas.filter(r => !String(r.Remarks ?? '').toUpperCase().includes('NO CONECTAR'));
}

// ==========================================
// RENDERIZADO DE TABLAS
// ==========================================

function renderGeneral(rows) {
  const tbody = document.getElementById('tbody-patio');
  if (!tbody) return;

  const data = rows.filter(r => !['EMBARCADO SIN CONEX', 'EMBARCADO SIN CONEXION (revisar)', 'EMBARQUE', 'EMBARQUE SOBRE CAMION'].includes(r.Instancia));

  const elCount = document.getElementById('count-patio');
  if (elCount) elCount.textContent = data.length;

  const calleRows = data.filter(r => r.Instancia === 'CALLE' || r.Instancia === 'CALLE REVISAR');
  const movimientoRows = data.filter(r => r.Instancia === 'MOVIMIENTO' || r.Instancia === 'MOVIMIENTO SOBRE CAMION');

  const elCalleCount = document.getElementById('kpi-calle-cnt');
  if (elCalleCount) elCalleCount.textContent = calleRows.length;
  const elCalleProm = document.getElementById('kpi-calle-avg');
  if (elCalleProm) elCalleProm.textContent = promedio(calleRows);

  const elMovCount = document.getElementById('kpi-mov-cnt');
  if (elMovCount) elMovCount.textContent = movimientoRows.length;
  const elMovProm = document.getElementById('kpi-mov-avg');
  if (elMovProm) elMovProm.textContent = promedio(movimientoRows);

  if (data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6" class="empty-state">sin registros</td></tr>';
    return;
  }

  const sorted = data.sort((a, b) => sortDescGeneral ? (b.TiempoMin ?? -1) - (a.TiempoMin ?? -1) : (a.TiempoMin ?? -1) - (b.TiempoMin ?? -1));

  let html = '';
  sorted.forEach(r => {
    const claseFila = obtenerClaseFila(r.Instancia);
    const badges = obtenerBadgesHTML(r.Remarks);
    const badgeInstancia = obtenerBadgeInstancia(r.Instancia);
    const badgeMov = obtenerBadgeMov(r.Mov);

    html += `
      <tr class="${claseFila}">
        <td><span class="container-id">${esc(r.Contenedor)}</span>${badges}</td>
        <td>${badgeInstancia}</td>
        <td><span class="pos-id">${esc(formatPosicion(r.Posicion))}</span></td>
        <td>${esc(r.Nave)}</td>
        <td ${claseTiempo(r.TiempoMin)}>${r.Tiempo ?? 'N/A'}</td>
        <td>${badgeMov}</td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
}

function renderSinConexion(rows) {
  const tbody = document.getElementById('tbody-bordo');
  if (!tbody) return;

  const data = rows.filter(r => ['EMBARCADO SIN CONEX', 'EMBARCADO SIN CONEXION (revisar)'].includes(r.Instancia));

  const elCount = document.getElementById('count-bordo');
  if (elCount) elCount.textContent = data.length;

  const elKpiCount = document.getElementById('kpi-bordo-cnt');
  if (elKpiCount) elKpiCount.textContent = data.length;

  const elKpiProm = document.getElementById('kpi-bordo-avg');
  if (elKpiProm) elKpiProm.textContent = promedio(data);

  if (data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-state">sin unidades a bordo sin conexión</td></tr>';
    return;
  }

  const sorted = data.sort((a, b) => sortDescSinConexion ? (b.TiempoMin ?? -1) - (a.TiempoMin ?? -1) : (a.TiempoMin ?? -1) - (b.TiempoMin ?? -1));

  let html = '';
  sorted.forEach(r => {
    const claseFila = obtenerClaseFila(r.Instancia);
    const badges = obtenerBadgesHTML(r.Remarks);

    html += `
      <tr class="${claseFila}">
        <td><span class="container-id">${esc(r.Contenedor)}</span>${badges}</td>
        <td>${esc(r.Nave)}</td>
        <td ${claseTiempo(r.TiempoMin)}>${r.Tiempo ?? 'N/A'}</td>
        <td><span class="pos-id">${esc(formatPosicion(r.Posicion))}</span></td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
}

function renderEmbarque(rows) {
  const tbody = document.getElementById('tbody-embarque');
  if (!tbody) return;

  const data = rows.filter(r => r.Instancia === 'EMBARQUE' || r.Instancia === 'EMBARQUE SOBRE CAMION');

  const elCount = document.getElementById('count-embarque');
  if (elCount) elCount.textContent = data.length;

  const elKpiCount = document.getElementById('kpi-emb-cnt');
  if (elKpiCount) elKpiCount.textContent = data.length;

  const elKpiProm = document.getElementById('kpi-emb-avg');
  if (elKpiProm) elKpiProm.textContent = promedio(data);

  if (data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="3" class="empty-state">sin unidades desconectadas para embarque</td></tr>';
    return;
  }

  const sorted = data.sort((a, b) => sortDescEmbarque ? (b.TiempoMin ?? -1) - (a.TiempoMin ?? -1) : (a.TiempoMin ?? -1) - (b.TiempoMin ?? -1));

  let html = '';
  sorted.forEach(r => {
    const claseFila = obtenerClaseFila(r.Instancia);
    const badges = obtenerBadgesHTML(r.Remarks);

    html += `
      <tr class="${claseFila}">
        <td><span class="container-id">${esc(r.Contenedor)}</span>${badges}</td>
        <td>${esc(r.Nave)}</td>
        <td ${claseTiempo(r.TiempoMin)}>${r.Tiempo ?? 'N/A'}</td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
}

// ==========================================
// CONSULTA PERIÓDICA DE DATOS Y EVENTOS
// ==========================================

async function fetchData() {
  try {
    const url = `${RAW_URL}?t=${Date.now()}`;
    const res = await fetch(url, { cache: 'no-store' });

    if (!res.ok) throw new Error('HTTP ' + res.status);
    let text = await res.text();

    if (!text || text.includes("Esperando primera carga")) return;

    if (text.charCodeAt(0) === 0xFEFF) {
      text = text.slice(1);
    }

    const rows = procesar(text);
    ultimasFilas = rows;

    const elTotal = document.getElementById('kpi-total-cnt');
    if (elTotal) elTotal.textContent = rows.length;

    renderGeneral(rows);
    renderSinConexion(rows);
    renderEmbarque(rows);

    const fechaReporte = extraerFechaReporte(text);
    const opClock = document.getElementById('last-update');

    if (opClock) {
      if (fechaReporte instanceof Date) {
        const dStr = fechaReporte.toLocaleDateString('es-CL', { day: '2-digit', month: '2-digit', year: '2-digit' });
        const tStr = fechaReporte.toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        opClock.textContent = 'Datos al ' + dStr + ' ' + tStr;
      } else if (typeof fechaReporte === 'string' && fechaReporte.trim() !== '') {
        opClock.textContent = fechaReporte;
      } else {
        const ahora = new Date();
        const dStr = ahora.toLocaleDateString('es-CL', { day: '2-digit', month: '2-digit', year: '2-digit' });
        const tStr = ahora.toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        opClock.textContent = 'ACTUALIZADO ' + dStr + ' ' + tStr;
      }
    }

  } catch (err) {
    console.error("Error al actualizar datos:", err);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  // Cambio de pestañas
  const tabButtons = document.querySelectorAll('.tab-link');
  const pageContainers = document.querySelectorAll('.page-container');

  tabButtons.forEach(button => {
    button.addEventListener('click', (e) => {
      const targetId = e.currentTarget.getAttribute('data-target');

      tabButtons.forEach(btn => btn.classList.remove('active'));
      e.currentTarget.classList.add('active');

      pageContainers.forEach(container => {
        if (container.id === targetId) {
          container.classList.add('active');
        } else {
          container.classList.remove('active');
        }
      });
    });
  });

  // Listener para exportar a Excel
  const btnExport = document.getElementById('btn-export-excel') || document.getElementById('btn-export');
  if (btnExport) {
    btnExport.addEventListener('click', exportarExcel);
  }

  fetchData();
  setInterval(fetchData, REFRESH_MS);
});
