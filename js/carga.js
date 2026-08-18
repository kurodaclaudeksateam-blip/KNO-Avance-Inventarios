const params = new URLSearchParams(window.location.search);
const cargaId = params.get('id');

let carga = null;
let areas = [];
let responsables = [];
let cortes = [];
let zoneChart = null;
let responsableChart = null;
let currentTab = 'detalle';
let modalResponsableContextAreaId = null;

const els = {
  nombreCarga: document.getElementById('nombre-carga'),
  metaCarga: document.getElementById('meta-carga'),
  uploadSection: document.getElementById('upload-section'),
  workspace: document.getElementById('workspace'),
  tablaBody: document.getElementById('tabla-body'),
  btnReemplazar: document.getElementById('btn-reemplazar'),
  btnAgregarArea: document.getElementById('btn-agregar-area'),
  btnExportar: document.getElementById('btn-exportar'),
  excelFile: document.getElementById('excel-file'),
  btnEmpezarVacio: document.getElementById('btn-empezar-vacio'),
  btnResponsables: document.getElementById('btn-responsables'),
  modalResponsables: document.getElementById('modal-responsables'),
  tituloModalResponsables: document.getElementById('titulo-modal-responsables'),
  descripcionModalResponsables: document.getElementById('descripcion-modal-responsables'),
  listaResponsables: document.getElementById('lista-responsables'),
  inputNuevoResponsable: document.getElementById('input-nuevo-responsable'),
  btnAgregarResponsable: document.getElementById('btn-agregar-responsable'),
  btnCerrarResponsables: document.getElementById('btn-cerrar-responsables'),
  inputValorTotal: document.getElementById('input-valor-total'),
  inputValorCotizado: document.getElementById('input-valor-cotizado'),
  inputSkusTotal: document.getElementById('input-skus-total'),
  inputSkusCotizado: document.getElementById('input-skus-cotizado'),
  cortesBody: document.getElementById('cortes-body'),
  btnNuevoCorte: document.getElementById('btn-nuevo-corte'),
  proyeccionCard: document.getElementById('proyeccion-card'),
};

const guardarCampo = debounce(async (key, id, field, value) => {
  const { error } = await supabaseClient.from('areas_almacen').update({ [field]: value }).eq('id', id);
  if (error) {
    console.error(error);
    showToast('No se pudo guardar el cambio.', 'error');
    return;
  }
  touchCarga();
}, 500);

async function init() {
  if (!cargaId) {
    window.location.href = 'index.html';
    return;
  }

  const { data: cargaData, error: cargaError } = await supabaseClient
    .from('cargas')
    .select('*')
    .eq('id', cargaId)
    .maybeSingle();

  if (cargaError || !cargaData) {
    showToast('No se encontró esa carga.', 'error');
    setTimeout(() => (window.location.href = 'index.html'), 1200);
    return;
  }

  carga = cargaData;
  els.nombreCarga.textContent = carga.nombre;
  els.metaCarga.textContent = `Actualizado ${formatFecha(carga.updated_at)}`;

  await Promise.all([cargarAreas(), cargarResponsables(), cargarCortes()]);

  if (areas.length === 0) {
    mostrarUpload();
  } else {
    mostrarWorkspace();
  }
}

async function cargarAreas() {
  const { data, error } = await supabaseClient
    .from('areas_almacen')
    .select('*')
    .eq('carga_id', cargaId)
    .order('orden', { ascending: true });

  if (error) {
    console.error(error);
    showToast('No se pudieron cargar las áreas.', 'error');
    areas = [];
    return;
  }
  areas = data || [];
}

async function cargarResponsables() {
  const { data, error } = await supabaseClient
    .from('responsables')
    .select('*')
    .eq('carga_id', cargaId)
    .order('nombre', { ascending: true });

  if (error) {
    console.error(error);
    responsables = [];
    return;
  }
  responsables = data || [];
}

async function cargarCortes() {
  const { data, error } = await supabaseClient
    .from('cortes_avance')
    .select('*')
    .eq('carga_id', cargaId)
    .order('fecha', { ascending: true });

  if (error) {
    console.error(error);
    cortes = [];
    return;
  }
  cortes = data || [];
}

function mostrarUpload() {
  els.uploadSection.classList.remove('hidden');
  els.workspace.classList.add('hidden');
  els.btnReemplazar.classList.add('hidden');
  els.btnAgregarArea.classList.add('hidden');
  els.btnExportar.classList.add('hidden');
  els.btnResponsables.classList.add('hidden');
}

function mostrarWorkspace() {
  els.uploadSection.classList.add('hidden');
  els.workspace.classList.remove('hidden');
  els.btnReemplazar.classList.remove('hidden');
  els.btnAgregarArea.classList.remove('hidden');
  els.btnExportar.classList.remove('hidden');
  els.btnResponsables.classList.remove('hidden');
  renderTabla();
  renderInforme();
  renderAvance();
}

// ---------- Excel (importar) ----------

function normalizarTexto(str) {
  return str.trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function buscarColumna(keys, nombreBuscado) {
  return keys.find((k) => normalizarTexto(k) === nombreBuscado);
}

els.excelFile.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async (ev) => {
    const data = new Uint8Array(ev.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    if (json.length === 0) {
      showToast('El archivo está vacío.', 'error');
      return;
    }

    const keys = Object.keys(json[0]);
    const columnaArea = buscarColumna(keys, 'areas del almacen');
    const columnaCotizacion = buscarColumna(keys, 'cotizacion');

    if (!columnaArea) {
      showToast('No se encontró la columna "areas del almacen".', 'error');
      return;
    }

    const nuevasAreas = json.map((row, index) => ({
      carga_id: cargaId,
      area: String(row[columnaArea] ?? ''),
      cotizacion: columnaCotizacion ? String(row[columnaCotizacion] ?? '') : '',
      limpieza: false,
      acomodo: false,
      etiquetado: false,
      avance_cotizacion: 0,
      duplicados: 0,
      orden: index,
    })).filter((a) => a.area.trim() !== '');

    const { error } = await supabaseClient.from('areas_almacen').insert(nuevasAreas);
    if (error) {
      console.error(error);
      showToast('No se pudieron guardar las áreas.', 'error');
      return;
    }

    await supabaseClient.from('cargas').update({ archivo_nombre: file.name, updated_at: new Date().toISOString() }).eq('id', cargaId);

    await cargarAreas();
    mostrarWorkspace();
    showToast(`Se cargaron ${nuevasAreas.length} áreas.`, 'success');
  };
  reader.readAsArrayBuffer(file);
});

els.btnEmpezarVacio.addEventListener('click', () => {
  mostrarWorkspace();
});

els.btnReemplazar.addEventListener('click', async () => {
  if (areas.length > 0 && !window.confirm('Esto borrará las áreas actuales de esta carga antes de subir el nuevo archivo. ¿Continuar?')) return;
  if (areas.length > 0) {
    const { error } = await supabaseClient.from('areas_almacen').delete().eq('carga_id', cargaId);
    if (error) {
      console.error(error);
      showToast('No se pudo reemplazar el archivo.', 'error');
      return;
    }
    areas = [];
  }
  mostrarUpload();
});

// ---------- Tabla / Detalle ----------

function nombreResponsable(id) {
  return id ? (responsables.find((r) => r.id === id)?.nombre || '') : '';
}

function renderResponsableTag(item) {
  const nombre = nombreResponsable(item.responsable_id);
  const asignado = !!nombre;
  return `<button type="button" class="responsable-tag ${asignado ? 'responsable-tag--assigned' : ''}" data-action="asignar-responsable" title="Clic para asignar responsable">${asignado ? escapeHtml(nombre) : '+ Asignar'}</button>`;
}

function renderTabla() {
  els.tablaBody.innerHTML = areas.map((item) => `
    <tr data-id="${item.id}">
      <td class="font-medium">${escapeHtml(item.area)}</td>
      <td><input type="text" value="${escapeHtml(item.cotizacion)}" data-field="cotizacion" placeholder="Nº Cotización"></td>
      <td style="text-align:center"><input type="checkbox" ${item.limpieza ? 'checked' : ''} data-field="limpieza"></td>
      <td style="text-align:center"><input type="checkbox" ${item.acomodo ? 'checked' : ''} data-field="acomodo"></td>
      <td style="text-align:center"><input type="checkbox" ${item.etiquetado ? 'checked' : ''} data-field="etiquetado"></td>
      <td style="text-align:center"><input type="number" min="0" max="100" value="${item.avance_cotizacion}" data-field="avance_cotizacion" style="text-align:center"></td>
      <td style="text-align:center"><input type="number" min="0" value="${item.duplicados}" data-field="duplicados" style="text-align:center; color:var(--color-danger)"></td>
      <td style="text-align:right; font-weight:600" id="subtotal-${item.id}">0.0%</td>
      <td>${renderResponsableTag(item)}</td>
      <td style="text-align:center; white-space:nowrap">
        <button class="btn-icon ${item.comentario ? 'has-comment' : ''}" data-action="comentario" title="${item.comentario ? 'Editar comentario' : 'Agregar comentario'}">💬</button>
        <button class="btn-icon" data-action="borrar" title="Borrar área">
          <svg width="15" height="15" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2">
            <path stroke-linecap="round" stroke-linejoin="round" d="M6 7h12M9 7V5a1 1 0 011-1h4a1 1 0 011 1v2m2 0-.7 12.1a2 2 0 01-2 1.9H8.7a2 2 0 01-2-1.9L6 7h12z" />
          </svg>
        </button>
      </td>
    </tr>
  `).join('');

  els.tablaBody.querySelectorAll('tr').forEach((tr) => {
    const id = tr.dataset.id;
    tr.querySelectorAll('input').forEach((input) => {
      const field = input.dataset.field;
      const eventName = input.type === 'checkbox' ? 'change' : 'input';
      input.addEventListener(eventName, () => onCampoCambiado(id, field, input));
    });
    tr.querySelector('[data-action="borrar"]').addEventListener('click', () => borrarArea(id));
    tr.querySelector('[data-action="comentario"]').addEventListener('click', () => editarComentario(id));
    tr.querySelector('[data-action="asignar-responsable"]').addEventListener('click', () => abrirModalResponsables(id));
  });

  calcularTotales();
}

function onCampoCambiado(id, field, input) {
  const item = areas.find((a) => a.id === id);
  if (!item) return;

  let value;
  if (input.type === 'checkbox') {
    value = input.checked;
  } else if (field === 'avance_cotizacion') {
    value = Math.min(Math.max(Number(input.value) || 0, 0), 100);
    input.value = value;
  } else if (field === 'duplicados') {
    value = Math.max(Number(input.value) || 0, 0);
    input.value = value;
  } else {
    value = input.value;
  }

  item[field] = value;
  calcularTotales();
  if (currentTab === 'informe') renderInforme();
  guardarCampo(`${id}-${field}`, id, field, value);
}

function calcularTotales() {
  const n = areas.length;
  if (n === 0) {
    ['lbl-avance-total', 'lbl-limpieza', 'lbl-acomodo', 'lbl-etiquetado', 'lbl-duplicados'].forEach((id) => {
      document.getElementById(id).textContent = '0.0%';
    });
    return;
  }

  let totalLimpieza = 0, totalAcomodo = 0, totalEtiquetado = 0, totalDuplicados = 0, totalPonderado = 0;

  areas.forEach((item) => {
    const limVal = item.limpieza ? 100 : 0;
    const acoVal = item.acomodo ? 100 : 0;
    const etiVal = item.etiquetado ? 100 : 0;
    const avCotVal = Math.min(Math.max(item.avance_cotizacion, 0), 100);
    const dupVal = Number(item.duplicados) || 0;

    totalLimpieza += limVal;
    totalAcomodo += acoVal;
    totalEtiquetado += etiVal;
    totalDuplicados += dupVal;

    const subtotal = item.subtotal = Math.max((limVal + acoVal + etiVal + avCotVal) / 4 - dupVal * 5, -100);
    totalPonderado += subtotal;

    const subEl = document.getElementById(`subtotal-${item.id}`);
    if (subEl) {
      subEl.textContent = subtotal.toFixed(1) + '%';
      subEl.style.color = subtotal < 0 ? 'var(--color-danger)' : 'inherit';
    }
  });

  const impactoDuplicados = (totalDuplicados * 5) / n;

  document.getElementById('lbl-limpieza').textContent = (totalLimpieza / n).toFixed(1) + '%';
  document.getElementById('lbl-acomodo').textContent = (totalAcomodo / n).toFixed(1) + '%';
  document.getElementById('lbl-etiquetado').textContent = (totalEtiquetado / n).toFixed(1) + '%';
  document.getElementById('lbl-duplicados').textContent = (impactoDuplicados === 0 ? '0.0' : '-' + impactoDuplicados.toFixed(1)) + '%';
  document.getElementById('lbl-avance-total').textContent = (totalPonderado / n).toFixed(1) + '%';
}

async function borrarArea(id) {
  if (!window.confirm('¿Borrar esta área?')) return;
  const { error } = await supabaseClient.from('areas_almacen').delete().eq('id', id);
  if (error) {
    console.error(error);
    showToast('No se pudo borrar el área.', 'error');
    return;
  }
  areas = areas.filter((a) => a.id !== id);
  touchCarga();
  if (areas.length === 0) {
    mostrarUpload();
  } else {
    renderTabla();
    renderInforme();
    renderAvance();
  }
}

function editarComentario(id) {
  const item = areas.find((a) => a.id === id);
  if (!item) return;
  const nuevo = window.prompt(`Comentario para "${item.area}":`, item.comentario || '');
  if (nuevo === null) return;
  item.comentario = nuevo;
  renderTabla();
  guardarCampo(`${id}-comentario`, id, 'comentario', nuevo);
}

els.btnAgregarArea.addEventListener('click', async () => {
  const nombre = window.prompt('Nombre de la nueva área:');
  if (!nombre || !nombre.trim()) return;

  const { data, error } = await supabaseClient
    .from('areas_almacen')
    .insert({ carga_id: cargaId, area: nombre.trim(), orden: areas.length })
    .select('*')
    .single();

  if (error) {
    console.error(error);
    showToast('No se pudo agregar el área.', 'error');
    return;
  }

  areas.push(data);
  touchCarga();

  if (els.workspace.classList.contains('hidden')) mostrarWorkspace();
  else { renderTabla(); renderInforme(); renderAvance(); }
});

// ---------- Responsables (gestión y asignación) ----------

function abrirModalResponsables(areaId = null) {
  modalResponsableContextAreaId = areaId;

  if (areaId) {
    const item = areas.find((a) => a.id === areaId);
    els.tituloModalResponsables.textContent = `Responsable de "${item ? item.area : ''}"`;
    els.descripcionModalResponsables.textContent = 'Elige quién es responsable de esta área, o agrega uno nuevo.';
  } else {
    els.tituloModalResponsables.textContent = 'Responsables';
    els.descripcionModalResponsables.textContent = 'Agrega personas y asígnalas a cada área haciendo clic en la etiqueta "Responsable" de la tabla.';
  }

  renderListaResponsables();
  els.modalResponsables.classList.remove('hidden');
  els.inputNuevoResponsable.focus();
}

function cerrarModalResponsables() {
  els.modalResponsables.classList.add('hidden');
  modalResponsableContextAreaId = null;
}

function renderListaResponsables() {
  const enContexto = !!modalResponsableContextAreaId;
  const itemActual = enContexto ? areas.find((a) => a.id === modalResponsableContextAreaId) : null;

  const limpiar = enContexto
    ? `<button type="button" class="responsable-row responsable-row--clear" data-action="quitar-asignacion">Sin asignar</button>`
    : '';

  if (responsables.length === 0) {
    els.listaResponsables.innerHTML = limpiar + '<p class="responsables-empty">Todavía no hay responsables. Agrega el primero abajo.</p>';
  } else {
    els.listaResponsables.innerHTML = limpiar + responsables.map((r) => `
      <div class="responsable-row" data-id="${r.id}">
        <button type="button" class="responsable-row__name" data-action="elegir-responsable">${escapeHtml(r.nombre)}${itemActual && itemActual.responsable_id === r.id ? ' ✓' : ''}</button>
        <button class="btn-icon" data-action="borrar-responsable" title="Quitar responsable">
          <svg width="14" height="14" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2">
            <path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12" />
          </svg>
        </button>
      </div>
    `).join('');
  }

  const btnLimpiar = els.listaResponsables.querySelector('[data-action="quitar-asignacion"]');
  if (btnLimpiar) btnLimpiar.addEventListener('click', () => asignarResponsable(modalResponsableContextAreaId, null));

  els.listaResponsables.querySelectorAll('[data-action="borrar-responsable"]').forEach((btn) => {
    const id = btn.closest('.responsable-row').dataset.id;
    btn.addEventListener('click', () => eliminarResponsable(id));
  });

  if (enContexto) {
    els.listaResponsables.querySelectorAll('[data-action="elegir-responsable"]').forEach((btn) => {
      const id = btn.closest('.responsable-row').dataset.id;
      btn.addEventListener('click', () => asignarResponsable(modalResponsableContextAreaId, id));
    });
  }
}

function asignarResponsable(areaId, responsableId) {
  if (!areaId) return;
  const item = areas.find((a) => a.id === areaId);
  if (!item) return;

  item.responsable_id = responsableId;
  renderTabla();
  if (currentTab === 'informe') renderInforme();
  guardarCampo(`${areaId}-responsable_id`, areaId, 'responsable_id', responsableId);
  cerrarModalResponsables();
  showToast(responsableId ? `Asignado a ${nombreResponsable(responsableId)}.` : 'Área sin asignar.', 'success');
}

async function agregarResponsable() {
  const nombre = els.inputNuevoResponsable.value.trim();
  if (!nombre) return;

  const { data, error } = await supabaseClient
    .from('responsables')
    .insert({ carga_id: cargaId, nombre })
    .select('*')
    .single();

  if (error) {
    console.error(error);
    showToast('No se pudo agregar el responsable.', 'error');
    return;
  }

  responsables.push(data);
  responsables.sort((a, b) => a.nombre.localeCompare(b.nombre));
  els.inputNuevoResponsable.value = '';

  if (modalResponsableContextAreaId) {
    asignarResponsable(modalResponsableContextAreaId, data.id);
    return;
  }

  renderListaResponsables();
  renderTabla();
  if (currentTab === 'informe') renderInforme();
}

async function eliminarResponsable(id) {
  if (!window.confirm('¿Quitar este responsable? Las áreas asignadas quedarán sin responsable.')) return;

  const { error } = await supabaseClient.from('responsables').delete().eq('id', id);
  if (error) {
    console.error(error);
    showToast('No se pudo quitar el responsable.', 'error');
    return;
  }

  responsables = responsables.filter((r) => r.id !== id);
  areas.forEach((a) => { if (a.responsable_id === id) a.responsable_id = null; });
  renderListaResponsables();
  renderTabla();
  if (currentTab === 'informe') renderInforme();
}

els.btnResponsables.addEventListener('click', () => abrirModalResponsables(null));
els.btnCerrarResponsables.addEventListener('click', cerrarModalResponsables);
els.modalResponsables.addEventListener('click', (e) => { if (e.target === els.modalResponsables) cerrarModalResponsables(); });
els.btnAgregarResponsable.addEventListener('click', agregarResponsable);
els.inputNuevoResponsable.addEventListener('keydown', (e) => { if (e.key === 'Enter') agregarResponsable(); });

// ---------- Informe por zona ----------

function renderInforme() {
  const zoneCardsEl = document.getElementById('zone-cards');
  if (areas.length === 0) {
    zoneCardsEl.innerHTML = '<p style="color:var(--color-text-muted)">Todavía no hay áreas para este informe.</p>';
    document.getElementById('responsable-cards').innerHTML = '';
    if (zoneChart) { zoneChart.destroy(); zoneChart = null; }
    if (responsableChart) { responsableChart.destroy(); responsableChart = null; }
    return;
  }

  const ordenadas = [...areas].sort((a, b) => (a.subtotal ?? 0) - (b.subtotal ?? 0));

  zoneCardsEl.innerHTML = ordenadas.map((item) => {
    const valor = item.subtotal ?? 0;
    const nivel = nivelAvance(valor);
    return `
      <div class="zone-card">
        <div class="zone-card__top">
          <span class="zone-card__area">${escapeHtml(item.area)}</span>
          <span class="zone-badge ${nivel}">${valor.toFixed(1)}%</span>
        </div>
        <div class="progress-track"><div class="progress-fill ${nivel === 'danger' ? 'danger' : nivel === 'warning' ? 'warning' : ''}" style="width:${Math.max(Math.min(valor, 100), 0)}%"></div></div>
      </div>
    `;
  }).join('');

  zoneChart = renderBarChart('zone-chart', zoneChart, ordenadas.map((a) => a.area), ordenadas.map((a) => a.subtotal ?? 0));

  // Avance por responsable
  const grupos = new Map();
  areas.forEach((item) => {
    const key = item.responsable_id || 'sin-asignar';
    if (!grupos.has(key)) {
      const nombre = item.responsable_id ? (responsables.find((r) => r.id === item.responsable_id)?.nombre || 'Responsable') : 'Sin asignar';
      grupos.set(key, { nombre, valores: [] });
    }
    grupos.get(key).valores.push(item.subtotal ?? 0);
  });

  const gruposOrdenados = [...grupos.values()]
    .map((g) => ({ nombre: g.nombre, valor: g.valores.reduce((a, b) => a + b, 0) / g.valores.length, count: g.valores.length }))
    .sort((a, b) => a.valor - b.valor);

  document.getElementById('responsable-cards').innerHTML = gruposOrdenados.map((g) => {
    const nivel = nivelAvance(g.valor);
    return `
      <div class="zone-card">
        <div class="zone-card__top">
          <span class="zone-card__area">${escapeHtml(g.nombre)}</span>
          <span class="zone-badge ${nivel}">${g.valor.toFixed(1)}%</span>
        </div>
        <p class="zone-card__meta">${g.count} área${g.count === 1 ? '' : 's'}</p>
        <div class="progress-track"><div class="progress-fill ${nivel === 'danger' ? 'danger' : nivel === 'warning' ? 'warning' : ''}" style="width:${Math.max(Math.min(g.valor, 100), 0)}%"></div></div>
      </div>
    `;
  }).join('');

  responsableChart = renderBarChart('responsable-chart', responsableChart, gruposOrdenados.map((g) => g.nombre), gruposOrdenados.map((g) => g.valor));
}

function renderBarChart(canvasId, existingChart, labels, valores) {
  const ctx = document.getElementById(canvasId);
  const chartData = {
    labels,
    datasets: [{
      label: 'Avance (%)',
      data: valores.map((v) => v.toFixed(1)),
      backgroundColor: valores.map((v) => {
        const nivel = nivelAvance(v);
        return nivel === 'success' ? '#16a34a' : nivel === 'warning' ? '#d97706' : '#dc2626';
      }),
      borderRadius: 6,
    }],
  };

  if (existingChart) {
    existingChart.data = chartData;
    existingChart.update();
    return existingChart;
  }

  return new Chart(ctx, {
    type: 'bar',
    data: chartData,
    options: {
      indexAxis: 'y',
      responsive: true,
      plugins: { legend: { display: false } },
      scales: { x: { min: -100, max: 100 } },
    },
  });
}

// ---------- Avance y Proyección (importes / cortes) ----------

function formatoMoneda(valor) {
  return '$' + (Number(valor) || 0).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

const guardarCampoCarga = debounce(async (key, field, value) => {
  const { error } = await supabaseClient.from('cargas').update({ [field]: value }).eq('id', cargaId);
  if (error) {
    console.error(error);
    showToast('No se pudo guardar el cambio.', 'error');
  }
}, 500);

function calcularAgregadosImportes() {
  const duplicados = areas.reduce((acc, item) => acc + (Number(item.duplicados) || 0), 0);
  return {
    valorTotal: Number(carga.valor_total) || 0,
    valorCotizado: Number(carga.valor_cotizado) || 0,
    skusTotal: Number(carga.skus_total) || 0,
    skusCotizado: Number(carga.skus_cotizado) || 0,
    duplicados,
  };
}

function renderImportesForm() {
  els.inputValorTotal.value = carga.valor_total;
  els.inputValorCotizado.value = carga.valor_cotizado;
  els.inputSkusTotal.value = carga.skus_total;
  els.inputSkusCotizado.value = carga.skus_cotizado;

  [
    [els.inputValorTotal, 'valor_total'],
    [els.inputValorCotizado, 'valor_cotizado'],
    [els.inputSkusTotal, 'skus_total'],
    [els.inputSkusCotizado, 'skus_cotizado'],
  ].forEach(([input, field]) => {
    input.oninput = () => {
      const value = Math.max(Number(input.value) || 0, 0);
      carga[field] = value;
      renderProyeccion(renderResumenImportes());
      guardarCampoCarga(field, field, value);
    };
  });
}

function renderResumenImportes() {
  const agg = calcularAgregadosImportes();
  const valorPendiente = Math.max(agg.valorTotal - agg.valorCotizado, 0);
  const skusPendiente = Math.max(agg.skusTotal - agg.skusCotizado, 0);

  document.getElementById('lbl-valor-total').textContent = formatoMoneda(agg.valorTotal);
  document.getElementById('lbl-valor-cotizado').textContent = formatoMoneda(agg.valorCotizado);
  document.getElementById('lbl-valor-pendiente').textContent = formatoMoneda(valorPendiente);
  document.getElementById('lbl-skus-total').textContent = agg.skusTotal;
  document.getElementById('lbl-skus-cotizado').textContent = agg.skusCotizado;
  document.getElementById('lbl-skus-pendiente').textContent = skusPendiente;
  document.getElementById('lbl-avance-duplicados').textContent = agg.duplicados;

  return { ...agg, valorPendiente, skusPendiente };
}

function renderCortesTabla() {
  if (cortes.length === 0) {
    els.cortesBody.innerHTML = '<tr><td colspan="5" style="color:var(--color-text-muted); text-align:center; padding:1.2rem">Todavía no hay cortes guardados. Usa "Iniciar nuevo corte" para tomar la primera foto del avance.</td></tr>';
    return;
  }

  els.cortesBody.innerHTML = [...cortes].reverse().map((c) => {
    const pctValor = c.valor_total > 0 ? (c.valor_cotizado / c.valor_total) * 100 : 0;
    return `
      <tr>
        <td>${formatFecha(c.fecha)}</td>
        <td style="text-align:right">${formatoMoneda(c.valor_cotizado)}</td>
        <td style="text-align:right">${pctValor.toFixed(1)}%</td>
        <td style="text-align:right">${c.skus_cotizado}</td>
        <td>${escapeHtml(c.nota || '')}</td>
      </tr>
    `;
  }).join('');
}

function renderProyeccion(agg) {
  if (cortes.length === 0) {
    els.proyeccionCard.innerHTML = `
      <p class="proyeccion-card__title">Proyección de término</p>
      <p class="proyeccion-card__main">Sin datos suficientes</p>
      <p class="proyeccion-card__detail">Guarda tu primer corte con "Iniciar nuevo corte" para empezar a calcular cuándo se llegará al 100% cotizado.</p>
    `;
    return;
  }

  const primerCorte = cortes[0];
  const ahora = new Date();
  const diasTranscurridos = Math.max((ahora - new Date(primerCorte.fecha)) / 86400000, 0.01);
  const avanceValor = agg.valorCotizado - Number(primerCorte.valor_cotizado);
  const ratePorDia = avanceValor / diasTranscurridos;

  if (agg.valorPendiente <= 0) {
    els.proyeccionCard.innerHTML = `
      <p class="proyeccion-card__title">Proyección de término</p>
      <p class="proyeccion-card__main" style="color:var(--color-success)">¡Completado! 100% cotizado</p>
      <p class="proyeccion-card__detail">El valor pendiente llegó a $0.00.</p>
    `;
    return;
  }

  if (ratePorDia <= 0) {
    els.proyeccionCard.innerHTML = `
      <p class="proyeccion-card__title">Proyección de término</p>
      <p class="proyeccion-card__main">Sin avance registrado todavía</p>
      <p class="proyeccion-card__detail">Desde el primer corte (${formatFecha(primerCorte.fecha)}) no se ha reducido el valor pendiente. Sigue capturando avance y guarda un nuevo corte.</p>
    `;
    return;
  }

  const diasRestantes = agg.valorPendiente / ratePorDia;
  const fechaEstimada = new Date(ahora.getTime() + diasRestantes * 86400000);

  els.proyeccionCard.innerHTML = `
    <p class="proyeccion-card__title">Proyección de término</p>
    <p class="proyeccion-card__main">${new Intl.DateTimeFormat('es-MX', { dateStyle: 'long' }).format(fechaEstimada)}</p>
    <p class="proyeccion-card__detail">A un ritmo de ${formatoMoneda(ratePorDia)}/día desde el ${formatFecha(primerCorte.fecha)}, faltan ~${Math.ceil(diasRestantes)} día${Math.ceil(diasRestantes) === 1 ? '' : 's'} para llegar al 100% cotizado y $0.00 pendiente.</p>
  `;
}

function renderAvance() {
  renderImportesForm();
  const agg = renderResumenImportes();
  renderCortesTabla();
  renderProyeccion(agg);
}

els.btnNuevoCorte.addEventListener('click', async () => {
  const agg = calcularAgregadosImportes();
  const nota = window.prompt('Nota para este corte (opcional):', '') || '';

  const { data, error } = await supabaseClient
    .from('cortes_avance')
    .insert({
      carga_id: cargaId,
      valor_total: agg.valorTotal,
      valor_cotizado: agg.valorCotizado,
      skus_total: agg.skusTotal,
      skus_cotizado: agg.skusCotizado,
      duplicados: agg.duplicados,
      nota,
    })
    .select('*')
    .single();

  if (error) {
    console.error(error);
    showToast('No se pudo guardar el corte.', 'error');
    return;
  }

  cortes.push(data);
  renderCortesTabla();
  renderProyeccion(renderResumenImportes());
  showToast('Corte guardado.', 'success');
});

// ---------- Exportar a Excel ----------

function slugify(str) {
  return normalizarTexto(str).replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '') || 'carga';
}

els.btnExportar.addEventListener('click', async () => {
  if (typeof ExcelJS === 'undefined') {
    showToast('No se pudo cargar el generador de Excel.', 'error');
    return;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'KNO Avance de Inventarios';

  const sheet = workbook.addWorksheet('Detalle');

  sheet.mergeCells('A1:I1');
  sheet.getCell('A1').value = `${carga.nombre} — Avance de Almacén`;
  sheet.getCell('A1').font = { bold: true, size: 14 };

  sheet.mergeCells('A2:I2');
  sheet.getCell('A2').value = `Actualizado: ${formatFecha(carga.updated_at)}`;
  sheet.getCell('A2').font = { italic: true, size: 10, color: { argb: 'FF6B7280' } };

  const metricHeaders = ['Avance Total Ponderado', 'Limpieza Promedio', 'Acomodo Promedio', 'Etiquetado Promedio', 'Impacto Duplicados'];
  const metricValues = ['lbl-avance-total', 'lbl-limpieza', 'lbl-acomodo', 'lbl-etiquetado', 'lbl-duplicados'].map((id) => document.getElementById(id).textContent);
  metricHeaders.forEach((h, i) => {
    const cell = sheet.getCell(4, i + 1);
    cell.value = h;
    cell.font = { bold: true, size: 9, color: { argb: 'FF6B7280' } };
  });
  metricValues.forEach((v, i) => {
    const cell = sheet.getCell(5, i + 1);
    cell.value = v;
    cell.font = { bold: true, size: 13 };
  });

  const headerRowIndex = 7;
  const headers = ['Área del Almacén', 'Nº Cotización', 'Limpieza', 'Acomodo', 'Etiquetado', 'Validación de Cotización (%)', 'Duplicados', 'Subtotal (%)', 'Responsable'];
  headers.forEach((h, i) => {
    const cell = sheet.getCell(headerRowIndex, i + 1);
    cell.value = h;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });

  areas.forEach((item, i) => {
    const rowIndex = headerRowIndex + 1 + i;
    const nivel = nivelAvance(item.subtotal ?? 0);
    const fillColor = nivel === 'success' ? 'FFDCFCE7' : nivel === 'warning' ? 'FFFEF3C7' : 'FFFEE2E2';
    const responsableNombre = item.responsable_id ? (responsables.find((r) => r.id === item.responsable_id)?.nombre || '') : '';
    const values = [
      item.area,
      item.cotizacion || '',
      item.limpieza ? 'Sí' : 'No',
      item.acomodo ? 'Sí' : 'No',
      item.etiquetado ? 'Sí' : 'No',
      Number(item.avance_cotizacion) || 0,
      Number(item.duplicados) || 0,
      Number((item.subtotal ?? 0).toFixed(1)),
      responsableNombre,
    ];
    values.forEach((v, colIdx) => {
      const cell = sheet.getCell(rowIndex, colIdx + 1);
      cell.value = v;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
    });
  });

  sheet.columns = [
    { width: 26 }, { width: 16 }, { width: 10 }, { width: 10 }, { width: 11 }, { width: 20 }, { width: 12 }, { width: 12 }, { width: 20 },
  ];

  const sheetComentarios = workbook.addWorksheet('Comentarios');
  ['Área', 'Responsable', 'Comentario'].forEach((h, i) => {
    const cell = sheetComentarios.getCell(1, i + 1);
    cell.value = h;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } };
  });

  const conComentario = areas.filter((a) => a.comentario && a.comentario.trim() !== '');
  if (conComentario.length === 0) {
    sheetComentarios.getCell(2, 1).value = 'Sin comentarios registrados.';
  } else {
    conComentario.forEach((item, i) => {
      const rowIndex = i + 2;
      const responsableNombre = item.responsable_id ? (responsables.find((r) => r.id === item.responsable_id)?.nombre || '') : '';
      sheetComentarios.getCell(rowIndex, 1).value = item.area;
      sheetComentarios.getCell(rowIndex, 2).value = responsableNombre;
      sheetComentarios.getCell(rowIndex, 3).value = item.comentario;
    });
  }
  sheetComentarios.columns = [{ width: 26 }, { width: 20 }, { width: 60 }];

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${slugify(carga.nombre)}-avance.xlsx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  showToast('Excel generado.', 'success');
});

// ---------- Tabs ----------

document.querySelectorAll('.tab-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    currentTab = btn.dataset.tab;
    document.querySelectorAll('.tab-btn').forEach((b) => b.classList.toggle('active', b === btn));
    document.getElementById('tab-detalle').classList.toggle('hidden', currentTab !== 'detalle');
    document.getElementById('tab-informe').classList.toggle('hidden', currentTab !== 'informe');
    document.getElementById('tab-avance').classList.toggle('hidden', currentTab !== 'avance');
    if (currentTab === 'informe') renderInforme();
    if (currentTab === 'avance') renderAvance();
  });
});

// ---------- Utils ----------

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str ?? '';
  return div.innerHTML;
}

async function touchCarga() {
  await supabaseClient.from('cargas').update({ updated_at: new Date().toISOString() }).eq('id', cargaId);
}

init();
