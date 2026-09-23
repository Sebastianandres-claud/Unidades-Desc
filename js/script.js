const RAW_URL = "https://github-live-proxy.samiranda.workers.dev";
const REFRESH_MS = 10000;
const CODIGOS_EXCLUIDOS = ["TIP", "TQ", "OUT"];
const DIAS = { LU:0, MA:1, MI:2, JU:3, VI:4, SA:5, DO:6 };

let expandedGeneral = false;
let ultimasFilas = [];
const LIMIT_GENERAL = 8;

let prevGeneral = {};
let prevSinConexion = {};
let prevEmbarque = {};

let sortDescGeneral = true;
let sortDescSinConexion = true;
let sortDescEmbarque = true;

let ultimaFechaReporte = null;

function exportarExcel(){
  if(!ultimasFilas || ultimasFilas.length === 0){
    alert('Todavía no hay datos cargados para exportar.');
    return;
  }

  const filasOrdenadas = [...ultimasFilas].sort((a,b) => (b.TiempoMin ?? -1) - (a.TiempoMin ?? -1));

  const datos = filasOrdenadas.map(r => ({
    'Contenedor': r.Contenedor,
    'Instancia': r.Instancia,
    'Posición': r.Posicion,
    'Nave': r.Nave,
    'Tiempo': r.Tiempo,
    'Minutos': r.TiempoMin,
    'Requiere Power': r.ReqsPower,
    'Mov': r.Mov ? 'Sí' : 'No',
    'Etiqueta': etiquetaDestacado(r.Remarks) ?? '',
    'Remarks': r.Remarks
  }));

  const ws = XLSX.utils.json_to_sheet(datos);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Unidades Desconectadas');

  const ahora = new Date();
  const sello = ahora.toISOString().slice(0,19).replace(/[:T]/g,'-');
  XLSX.writeFile(wb, `unidades_desconectadas_${sello}.xlsx`);
}

function mostrarPagina(pageId, btn){
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById(pageId).classList.add('active');
  btn.classList.add('active');
}

function toggleSortGeneral(){
  sortDescGeneral = !sortDescGeneral;
  renderGeneral(ultimasFilas);
}

function toggleSortSinConexion(){
  sortDescSinConexion = !sortDescSinConexion;
  renderSinConexion(ultimasFilas);
}

function toggleSortEmbarque(){
  sortDescEmbarque = !sortDescEmbarque;
  renderEmbarque(ultimasFilas);
}

function formatPosicion(pos){
  if(pos === null || pos === undefined) return '-';
  const txt = String(pos).trim();
  if(!txt) return '-';

  const partes = txt.split('*').map(p => p.trim()).filter(p => p.length > 0);
  if(partes.length === 0) return txt;
  return partes[partes.length - 1];
}

function claseTiempo(minutos){
  if(minutos === null || minutos === undefined) return 'time-neutral';
  if(minutos >= 60) return 'time-red';
  if(minutos < 10) return 'time-green';
  return 'time-yellow';
}

function etiquetaDestacado(remarks){
  const txt = String(remarks ?? '').toUpperCase();
  const esUSDA = /\bUSDA\b/.test(txt);
  const esAC = /\bAC\b/.test(txt) || /\d+\s*\/\s*\d+/.test(txt);

  const partes = [];
  if(esUSDA) partes.push('COT');
  if(esAC) partes.push('AC');
  return partes.length > 0 ? partes.join('/') : null;
}

function promedio(arr){
  if(arr.length === 0) return '0.0';
  const suma = arr.reduce((acc,r) => acc + (r.TiempoMin ?? 0), 0);
  return (suma/arr.length).toFixed(1);
}

function toggleGeneral(){
  expandedGeneral = !expandedGeneral;
  renderGeneral(ultimasFilas);
}

function esVacio(v){ return v === null || v === undefined || String(v).trim() === ''; }

function esPosicionVigente(pos, anioActual){
  const limpio = String(pos ?? '').trim();
  const basura = CODIGOS_EXCLUIDOS.includes(limpio.toUpperCase());
  const larga = limpio.length > 10;
  let anioRot = null;
  const prefijo = parseInt(limpio.substring(0,2), 10);
  if(!isNaN(prefijo)) anioRot = 2000 + prefijo;

  if(basura) return false;
  if(!larga) return true;
  return anioRot === anioActual || anioRot === anioActual - 1;
}

function mondayIndex(date){
  return (date.getDay() + 6) % 7;
}

function parseDiscon(valor){
  const txt = String(valor ?? '').trim();
  if(txt.length !== 10) return null;
  const yy = parseInt(txt.substring(0,2),10);
  const mm = parseInt(txt.substring(2,4),10);
  const dd = parseInt(txt.substring(4,6),10);
  const hh = parseInt(txt.substring(6,8),10);
  const mi = parseInt(txt.substring(8,10),10);
  if([yy,mm,dd,hh,mi].some(isNaN)) return null;
  return new Date(2000+yy, mm-1, dd, hh, mi, 0);
}

function parseHoraSimple(valor, ahora){
  const txt = String(valor ?? '').trim();
  if(txt.length < 6) return null;
  const codigoDia = txt.substring(0,2).toUpperCase();
  const sufijo = txt.substring(txt.length-4);
  const hh = parseInt(sufijo.substring(0,2),10);
  const mi = parseInt(sufijo.substring(2,4),10);
  const diaObjetivo = DIAS.hasOwnProperty(codigoDia) ? DIAS[codigoDia] : null;

  const valido = diaObjetivo !== null && !isNaN(hh) && !isNaN(mi) && hh>=0 && hh<=23 && mi>=0 && mi<=59;
  if(!valido) return null;

  const hoy = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate());
  let offsetMatch = 0;
  for(let o=0; o<=6; o++){
    const candidatoDia = new Date(hoy);
    candidatoDia.setDate(hoy.getDate()-o);
    if(mondayIndex(candidatoDia) === diaObjetivo){ offsetMatch = o; break; }
  }
  const fechaBase = new Date(hoy);
  fechaBase.setDate(hoy.getDate()-offsetMatch);
  let candidato = new Date(fechaBase.getFullYear(), fechaBase.getMonth(), fechaBase.getDate(), hh, mi, 0);
  if(candidato > ahora){
    candidato = new Date(candidato.getTime() - 7*24*60*60*1000);
  }
  return candidato;
}

function formatearHHMM(diffMs){
  if(diffMs === null || diffMs === undefined) return null;
  const totalMin = Math.round(diffMs/60000);
  const horas = Math.floor(totalMin/60);
  const minutos = totalMin % 60;
  return String(horas).padStart(2,'0') + ':' + String(minutos).padStart(2,'0');
}

function extraerFechaReporte(text){
  const lines = text.replace(/\r/g,'').split('\n').map(l => l.trim()).filter(l => l.length > 0);
  if(lines.length === 0) return null;
  const ultima = lines[lines.length-1];
  const match = ultima.match(/^(\d{2})-(\d{2})-(\d{4})\t(\d{2}):(\d{2}):(\d{2})$/);
  if(!match) return null;
  return new Date(Number(match[3]), Number(match[2])-1, Number(match[1]), Number(match[4]), Number(match[5]), Number(match[6]));
}

function parseTSV(text){
  const cleanText = text.replace(/^\ufeff/, '');
  const lines = cleanText.replace(/\r/g,'').split('\n');
  if(lines.length === 0) return [];
  const headers = lines[0].split('\t').map(h => h.trim());
  const rows = [];
  for(let i=1;i<lines.length;i++){
    const line = lines[i];
    if(!line || !line.includes('\t')) continue;
    const cells = line.split('\t');
    const row = {};
    headers.forEach((h,idx) => row[h] = (cells[idx] ?? '').trim());
    rows.push(row);
  }
  return rows;
}

function calcularFila(fila, ahora, anioActual){
  const powerTxt = String(fila['Power'] ?? '');
  const seMovio = powerTxt.includes('(') ? 'Sí' : 'No';
  const powerNum = parseInt(powerTxt.replace(/[()]/g,'').trim(), 10) || 0;

  const posActual = String(fila['Current Position'] ?? '').trim();
  const discon = fila['Discon'];
  const remarks = fila['Remarks'];
  const yardPlan = fila['Yard Plan'];
  const kind = String(fila['Kind'] ?? '').trim();
  const moveStage = String(fila['Move Stage'] ?? '').trim();

  let instancia;

  if(posActual.length > 10){
    if(kind === 'DSCH') instancia = 'DESCARGA';
    else if(kind === 'LOAD') instancia = 'EMBARCADO SIN CONEX';
    else instancia = 'EMBARCADO SIN CONEXION (revisar)';
  } else {
    if(kind === 'DSCH'){
      instancia = 'DESCARGA';
    } else if(esVacio(discon)){
      instancia = esVacio(remarks) ? 'CALLE' : 'CALLE REVISAR';
    } else {
      if(!esVacio(yardPlan)) instancia = 'MOVIMIENTO';
      else if(moveStage === 'Completed') instancia = 'MOVIMIENTO';
      else {
        if(powerNum === 0 && kind === 'LOAD') instancia = 'EMBARQUE';
        else if(powerNum === 0) instancia = 'DESPACHO';
        else instancia = 'SIN CLASIFICAR';
      }
    }
  }

  if(posActual.startsWith('*TR')){
    if(instancia === 'EMBARQUE'){
      instancia = 'EMBARQUE SOBRE CAMION';
    } else if(instancia === 'MOVIMIENTO'){
      instancia = 'MOVIMIENTO SOBRE CAMION';
    } else {
      instancia = 'SOBRE CAMION';
    }
  }

  const usaDiscon = ['MOVIMIENTO','EMBARQUE','DESPACHO','SOBRE CAMION','SIN CLASIFICAR','DESCARGA','MOVIMIENTO SOBRE CAMION','EMBARQUE SOBRE CAMION'].includes(instancia);
  const usaYardIn = ['CALLE','CALLE REVISAR'].includes(instancia);
  const usaComplete = ['EMBARCADO SIN CONEX','EMBARCADO SIN CONEXION (revisar)'].includes(instancia);

  let fechaEvento = null;
  if(usaDiscon) fechaEvento = parseDiscon(discon);
  else if(usaYardIn) fechaEvento = parseHoraSimple(fila['Yard In'], ahora);
  else if(usaComplete) fechaEvento = parseHoraSimple(fila['Complete'], ahora);

  const tiempoTranscurrido = fechaEvento ? formatearHHMM(ahora - fechaEvento) : null;
  const tiempoMin = fechaEvento ? Math.round((ahora - fechaEvento)/60000) : null;
  const alerta = instancia === 'SIN CLASIFICAR' ? 'Revisar' : null;

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
    Alerta: alerta,
    Remarks: remarks
  };
}

function procesar(text){
  const ahora = new Date();
  const anioActual = ahora.getFullYear();
  const rows = parseTSV(text);

  const filasValidas = rows.filter(r => !esVacio(r['Container No.']));
  const sinBasura = filasValidas.filter(r => esPosicionVigente(r['Current Position'], anioActual));
  const calculadas = sinBasura.map(r => calcularFila(r, ahora, anioActual));
  const filtradas = calculadas.filter(r => !String(r.Remarks ?? '').toUpperCase().includes('NO CONECTAR'));

  return filtradas;
}

function badgeInstancia(inst){
  const map = {
    'DESCARGA':'badge-descarga', 'CALLE':'badge-calle', 'CALLE REVISAR':'badge-revisar',
    'EMBARQUE':'badge-embarque', 'MOVIMIENTO':'badge-movimiento', 'DESPACHO':'badge-movimiento',
    'SOBRE CAMION':'badge-calle', 'SIN CLASIFICAR':'badge-revisar',
    'EMBARQUE SOBRE CAMION': 'badge-embarque',
    'MOVIMIENTO SOBRE CAMION': 'badge-movimiento'
  };
  const cls = map[inst] || '';
  return `<span class="badge ${cls}">${inst}</span>`;
}

function esc(v){
  return String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function renderGeneral(rows){
  const container = document.getElementById('tableGeneral');
  if(!container) return;

  const data = rows.filter(r => !['EMBARCADO SIN CONEX','EMBARCADO SIN CONEXION (revisar)','EMBARQUE', 'EMBARQUE SOBRE CAMION'].includes(r.Instancia));
  
  const elCount = document.getElementById('countGeneral');
  if(elCount) elCount.textContent = data.length;

  const calleRows = data.filter(r => r.Instancia === 'CALLE' || r.Instancia === 'CALLE REVISAR');
  const movimientoRows = data.filter(r => r.Instancia === 'MOVIMIENTO' || r.Instancia === 'MOVIMIENTO SOBRE CAMION');

  const elCalleCount = document.getElementById('kpiCalleCount');
  if(elCalleCount) elCalleCount.textContent = calleRows.length;

  const elCalleProm = document.getElementById('kpiCalleProm');
  if(elCalleProm) elCalleProm.textContent = promedio(calleRows);

  const elMovCount = document.getElementById('kpiMovimientoCount');
  if(elMovCount) elMovCount.textContent = movimientoRows.length;

  const elMovProm = document.getElementById('kpiMovimientoProm');
  if(elMovProm) elMovProm.textContent = promedio(movimientoRows);

  if(data.length === 0){ container.innerHTML = '<div class="empty-state">sin registros</div>'; return; }

  const sorted = data.sort((a,b)=> sortDescGeneral ? (b.TiempoMin??-1) - (a.TiempoMin??-1) : (a.TiempoMin??-1) - (b.TiempoMin??-1));
  const toShow = expandedGeneral ? sorted : sorted.slice(0, LIMIT_GENERAL);

  const nuevoPrev = {};
  let html = `<table><thead><tr><th style="width:24%">Contenedor</th><th style="width:26%">Instancia</th><th style="width:14%">Posición</th><th style="width:18%">Nave</th><th style="width:11%" class="sortable" onclick="toggleSortGeneral()">Tiempo <span class="sort-arrow">${sortDescGeneral ? '▼' : '▲'}</span></th><th style="width:7%">Mov</th></tr></thead><tbody>`;
  toShow.forEach(r => {
    const etiqueta = etiquetaDestacado(r.Remarks);
    const cambio = prevGeneral[r.Contenedor] !== undefined && prevGeneral[r.Contenedor] !== r.Instancia;
    nuevoPrev[r.Contenedor] = r.Instancia;
    html += `<tr class="${etiqueta ? 'row-highlight' : ''} ${cambio ? 'flash-update' : ''}"><td>${esc(r.Contenedor)}</td><td>${badgeInstancia(r.Instancia)}${etiqueta ? `<span class="flag">${etiqueta}</span>` : ''}</td><td>${esc(formatPosicion(r.Posicion))}</td><td>${esc(r.Nave)}</td><td class="time-cell ${claseTiempo(r.TiempoMin)}">${r.Tiempo ?? 'N/A'}</td><td class="${r.Mov ? 'mov-si' : 'mov-no'}">${r.Mov ? 'Sí' : 'No'}</td></tr>`;
  });
  html += '</tbody></table>';
  prevGeneral = nuevoPrev;

  if(sorted.length > LIMIT_GENERAL){
    html += `<button class="toggle-btn" onclick="toggleGeneral()">${expandedGeneral ? 'Mostrar menos' : `Ver todas (${sorted.length})`}</button>`;
  }

  container.innerHTML = html;
}

function renderSinConexion(rows){
  const container = document.getElementById('tableSinConexion');
  if(!container) return;

  const data = rows.filter(r => ['EMBARCADO SIN CONEX','EMBARCADO SIN CONEXION (revisar)'].includes(r.Instancia));
  
  const elCount = document.getElementById('countSinConexion');
  if(elCount) elCount.textContent = data.length;

  const elKpiCount = document.getElementById('kpiSinConexionCount');
  if(elKpiCount) elKpiCount.textContent = data.length;

  const elKpiProm = document.getElementById('kpiSinConexionProm');
  if(elKpiProm) elKpiProm.textContent = promedio(data);

  if(data.length === 0){ container.innerHTML = '<div class="empty-state">sin unidades a bordo sin conexión</div>'; return; }

  const sorted = data.sort((a,b)=> sortDescSinConexion ? (b.TiempoMin??-1) - (a.TiempoMin??-1) : (a.TiempoMin??-1) - (b.TiempoMin??-1));

  const nuevoPrev = {};
  let html = `<table><thead><tr><th style="width:36%">Contenedor</th><th style="width:27%">Nave</th><th style="width:18%" class="sortable" onclick="toggleSortSinConexion()">Tiempo <span class="sort-arrow">${sortDescSinConexion ? '▼' : '▲'}</span></th><th style="width:19%">Posición</th></tr></thead><tbody>`;
  sorted.forEach(r => {
    const etiqueta = etiquetaDestacado(r.Remarks);
    const cambio = prevSinConexion[r.Contenedor] === undefined;
    nuevoPrev[r.Contenedor] = true;
    html += `<tr class="${etiqueta ? 'row-highlight' : ''} ${cambio ? 'flash-update' : ''}"><td>${esc(r.Contenedor)}${etiqueta ? `<span class="flag">${etiqueta}</span>` : ''}</td><td>${esc(r.Nave)}</td><td class="time-cell ${claseTiempo(r.TiempoMin)}">${r.Tiempo ?? 'N/A'}</td><td>${esc(formatPosicion(r.Posicion))}</td></tr>`;
  });
  html += '</tbody></table>';
  prevSinConexion = nuevoPrev;
  container.innerHTML = html;
}

function renderEmbarque(rows){
  const container = document.getElementById('tableEmbarque');
  if(!container) return;

  const data = rows.filter(r => r.Instancia === 'EMBARQUE' || r.Instancia === 'EMBARQUE SOBRE CAMION');

  const elCount = document.getElementById('countEmbarque');
  if(elCount) elCount.textContent = data.length;

  const elKpiCount = document.getElementById('kpiEmbarqueCount');
  if(elKpiCount) elKpiCount.textContent = data.length;

  const elKpiProm = document.getElementById('kpiEmbarqueProm');
  if(elKpiProm) elKpiProm.textContent = promedio(data);

  if(data.length === 0){ container.innerHTML = '<div class="empty-state">sin unidades desconectadas para embarque</div>'; return; }

  const sorted = data.sort((a,b)=> sortDescEmbarque ? (b.TiempoMin??-1) - (a.TiempoMin??-1) : (a.TiempoMin??-1) - (b.TiempoMin??-1));

  const nuevoPrev = {};
  let html = `<table><thead><tr><th style="width:36%">Contenedor</th><th style="width:27%">Nave</th><th style="width:18%" class="sortable" onclick="toggleSortEmbarque()">Tiempo <span class="sort-arrow">${sortDescEmbarque ? '▼' : '▲'}</span></th><th style="width:19%">Posición</th></tr></thead><tbody>`;
  sorted.forEach(r => {
    const etiqueta = etiquetaDestacado(r.Remarks);
    const cambio = prevEmbarque[r.Contenedor] === undefined;
    nuevoPrev[r.Contenedor] = true;
    html += `<tr class="${etiqueta ? 'row-highlight' : ''} ${cambio ? 'flash-update' : ''}"><td>${esc(r.Contenedor)}${etiqueta ? `<span class="flag">${etiqueta}</span>` : ''}</td><td>${esc(r.Nave)}</td><td class="time-cell ${claseTiempo(r.TiempoMin)}">${r.Tiempo ?? 'N/A'}</td><td>${esc(formatPosicion(r.Posicion))}</td></tr>`;
  });
  html += '</tbody></table>';
  prevEmbarque = nuevoPrev;
  container.innerHTML = html;
}

async function fetchData(){
  const statusDot = document.getElementById('statusDot');
  const statusText = document.getElementById('statusText');
  const opClock = document.getElementById('opClock');

  if(statusText) statusText.textContent = 'actualizando...';

  try{
    const url = `${RAW_URL}?t=${Date.now()}`;
    const res = await fetch(url, { cache: 'no-store' });
    if(!res.ok) throw new Error('HTTP ' + res.status);
    let text = await res.text();

    if (text.charCodeAt(0) === 0xFEFF) {
      text = text.slice(1);
    }

    const rows = procesar(text);
    ultimasFilas = rows;

    const elTotal = document.getElementById('kpiTotalCount');
    if(elTotal) elTotal.textContent = rows.length;

    renderGeneral(rows);
    renderSinConexion(rows);
    renderEmbarque(rows);

    if(statusDot) statusDot.classList.remove('offline');
    if(statusText) statusText.textContent = 'en vivo';

    const fechaReporte = extraerFechaReporte(text);

    if(opClock){
      if(fechaReporte){
        opClock.textContent = 'DATOS AL ' + fechaReporte.toLocaleDateString('es-CL',{day:'2-digit',month:'2-digit',year:'2-digit'}) + ' ' + fechaReporte.toLocaleTimeString('es-CL',{hour:'2-digit',minute:'2-digit',second:'2-digit'});

        if(ultimaFechaReporte !== null && fechaReporte.getTime() !== ultimaFechaReporte.getTime()){
          const strip = document.querySelector('.kpi-strip');
          if(strip){
            strip.classList.remove('data-updated');
            void strip.offsetWidth;
            strip.classList.add('data-updated');
          }
        }
        ultimaFechaReporte = fechaReporte;
      } else {
        const lastModifiedHeader = res.headers.get('last-modified');
        const fileDate = lastModifiedHeader ? new Date(lastModifiedHeader) : null;
        if(fileDate){
          opClock.textContent = 'ARCHIVO ACTUALIZADO ' + fileDate.toLocaleDateString('es-CL',{day:'2-digit',month:'2-digit',year:'2-digit'}) + ' ' + fileDate.toLocaleTimeString('es-CL',{hour:'2-digit',minute:'2-digit',second:'2-digit'});
        } else {
          const now = new Date();
          opClock.textContent = 'VERIFICADO ' + now.toLocaleTimeString('es-CL',{hour:'2-digit',minute:'2-digit',second:'2-digit'}) + ' (sin fecha de archivo disponible)';
        }
      }
    }

  }catch(err){
    if(statusDot) statusDot.classList.add('offline');
    if(statusText) statusText.textContent = 'sin conexión';
    console.error("Error al actualizar datos:", err);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  fetchData();
  setInterval(fetchData, REFRESH_MS);
});
