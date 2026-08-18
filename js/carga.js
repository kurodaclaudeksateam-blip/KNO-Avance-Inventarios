const params = new URLSearchParams(window.location.search);
const cargaId = params.get('id');

let carga = null;
let areas = [];
let zoneChart = null;
let currentTab = 'detalle';

const els = {
  nombreCarga: document.getElementById('nombre-carga'),
  metaCarga: document.getElementById('meta-carga'),
  uploadSection: document.getElementById('upload-section'),
  workspace: document.getElementById('workspace'),
  tablaBody: document.getElementById('tabla-body'),
  btnReemplazar: document.getElementById('btn-reemplazar'),
  btnAgregarArea: document.getElementById('btn-agregar-area'),
  excelFile: document.getElementById('excel-file'),
  btnEmpezarVacio: document.getElementById('btn-empezar-vacio'),
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

  await cargarAreas();

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

function mostrarUpload() {
  els.uploadSection.classList.remove('hidden');
  els.workspace.classList.add('hidden');
  els.btnReemplazar.classList.add('hidden');
  els.btnAgregarArea.classList.add('hidden');
}

function mostrarWorkspace() {
  els.uploadSection.classList.add('hidden');
  els.workspace.classList.remove('hidden');
  els.btnReemplazar.classList.remove('hidden');
  els.btnAgregarArea.classList.remove('hidden');
  renderTabla();
  renderInforme();
}

// ---------- Excel ----------

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

function renderTabla() {
  els.tablaBody.innerHTML = areas.map((item, index) => `
    <tr data-id="${item.id}">
      <td class="font-medium">${escapeHtml(item.area)}</td>
      <td><input type="text" value="${escapeHtml(item.cotizacion)}" data-field="cotizacion" placeholder="Nº Cotización"></td>
      <td style="text-align:center"><input type="checkbox" ${item.limpieza ? 'checked' : ''} data-field="limpieza"></td>
      <td style="text-align:center"><input type="checkbox" ${item.acomodo ? 'checked' : ''} data-field="acomodo"></td>
      <td style="text-align:center"><input type="number" min="0" max="100" value="${item.avance_cotizacion}" data-field="avance_cotizacion" style="text-align:center"></td>
      <td style="text-align:center"><input type="number" min="0" value="${item.duplicados}" data-field="duplicados" style="text-align:center; color:var(--color-danger)"></td>
      <td style="text-align:right; font-weight:600" id="subtotal-${item.id}">0.0%</td>
      <td style="text-align:center">
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
    ['lbl-avance-total', 'lbl-limpieza', 'lbl-acomodo', 'lbl-duplicados'].forEach((id) => {
      document.getElementById(id).textContent = '0.0%';
    });
    return;
  }

  let totalLimpieza = 0, totalAcomodo = 0, totalDuplicados = 0, totalPonderado = 0;

  areas.forEach((item) => {
    const limVal = item.limpieza ? 100 : 0;
    const acoVal = item.acomodo ? 100 : 0;
    const avCotVal = Math.min(Math.max(item.avance_cotizacion, 0), 100);
    const dupVal = Number(item.duplicados) || 0;

    totalLimpieza += limVal;
    totalAcomodo += acoVal;
    totalDuplicados += dupVal;

    const subtotal = item.subtotal = Math.max((limVal + acoVal + avCotVal) / 3 - dupVal * 5, -100);
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
  }
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
  else { renderTabla(); renderInforme(); }
});

// ---------- Informe por zona ----------

function renderInforme() {
  const zoneCardsEl = document.getElementById('zone-cards');
  if (areas.length === 0) {
    zoneCardsEl.innerHTML = '<p style="color:var(--color-text-muted)">Todavía no hay áreas para este informe.</p>';
    if (zoneChart) { zoneChart.destroy(); zoneChart = null; }
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

  const ctx = document.getElementById('zone-chart');
  const chartData = {
    labels: ordenadas.map((a) => a.area),
    datasets: [{
      label: 'Avance por área (%)',
      data: ordenadas.map((a) => (a.subtotal ?? 0).toFixed(1)),
      backgroundColor: ordenadas.map((a) => {
        const nivel = nivelAvance(a.subtotal ?? 0);
        return nivel === 'success' ? '#16a34a' : nivel === 'warning' ? '#d97706' : '#dc2626';
      }),
      borderRadius: 6,
    }],
  };

  if (zoneChart) {
    zoneChart.data = chartData;
    zoneChart.update();
  } else {
    zoneChart = new Chart(ctx, {
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
}

// ---------- Tabs ----------

document.querySelectorAll('.tab-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    currentTab = btn.dataset.tab;
    document.querySelectorAll('.tab-btn').forEach((b) => b.classList.toggle('active', b === btn));
    document.getElementById('tab-detalle').classList.toggle('hidden', currentTab !== 'detalle');
    document.getElementById('tab-informe').classList.toggle('hidden', currentTab !== 'informe');
    if (currentTab === 'informe') renderInforme();
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
