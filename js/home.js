const deckEl = document.getElementById('deck');
const modalEl = document.getElementById('modal-nueva');
const inputNombreEl = document.getElementById('input-nombre-carga');

const NEW_CARD_HTML = `
  <div class="carga-card carga-card--new" id="card-nueva" role="button" tabindex="0" aria-label="Crear nueva carga">
    <svg fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2">
      <path stroke-linecap="round" stroke-linejoin="round" d="M12 4.5v15m7.5-7.5h-15" />
    </svg>
    Nueva carga
  </div>
`;

function onClickOrEnter(el, handler) {
  el.addEventListener('click', handler);
  el.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      handler(e);
    }
  });
}

async function cargarDeck() {
  const { data, error } = await supabaseClient
    .from('vista_resumen_cargas')
    .select('*')
    .order('updated_at', { ascending: false });

  if (error) {
    console.error(error);
    deckEl.innerHTML = NEW_CARD_HTML;
    showToast('No se pudieron cargar las cargas guardadas.', 'error');
    return;
  }

  renderDeck(data || []);
}

function renderDeck(cargas) {
  const cardsHtml = cargas.map(renderCargaCard).join('');
  deckEl.innerHTML = cardsHtml + NEW_CARD_HTML;

  cargas.forEach((carga) => {
    const card = document.getElementById(`card-${carga.id}`);
    if (!card) return;
    onClickOrEnter(card, () => {
      window.location.href = `carga.html?id=${carga.id}`;
    });
    const btnBorrar = card.querySelector('.btn-icon');
    if (btnBorrar) {
      btnBorrar.addEventListener('click', (e) => {
        e.stopPropagation();
        borrarCarga(carga.id, carga.nombre);
      });
    }
  });

  onClickOrEnter(document.getElementById('card-nueva'), abrirModal);
}

function renderCargaCard(carga) {
  const avance = Number(carga.avance_promedio) || 0;
  const nivel = nivelAvance(avance);
  const anchoBarra = Math.max(Math.min(avance, 100), 0);
  const totalAreas = carga.total_areas || 0;

  return `
    <div class="carga-card" id="card-${carga.id}" role="button" tabindex="0" aria-label="Abrir carga ${escapeHtml(carga.nombre)}">
      <div class="carga-card__top">
        <div>
          <p class="carga-card__title">${escapeHtml(carga.nombre)}</p>
          <p class="carga-card__meta">${totalAreas} área${totalAreas === 1 ? '' : 's'} · ${formatFecha(carga.updated_at)}</p>
        </div>
        <button class="btn-icon" title="Borrar carga">
          <svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2">
            <path stroke-linecap="round" stroke-linejoin="round" d="M6 7h12M9 7V5a1 1 0 011-1h4a1 1 0 011 1v2m2 0-.7 12.1a2 2 0 01-2 1.9H8.7a2 2 0 01-2-1.9L6 7h12z" />
          </svg>
        </button>
      </div>
      <div class="carga-card__score">
        <span class="carga-card__score-value" style="color: var(--color-${nivel === 'success' ? 'success' : nivel === 'warning' ? 'warning' : 'danger'})">${avance.toFixed(1)}%</span>
        <span class="carga-card__meta">avance</span>
      </div>
      <div class="progress-track"><div class="progress-fill ${nivel === 'danger' ? 'danger' : nivel === 'warning' ? 'warning' : ''}" style="width:${anchoBarra}%"></div></div>
    </div>
  `;
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str ?? '';
  return div.innerHTML;
}

function abrirModal() {
  modalEl.classList.remove('hidden');
  inputNombreEl.value = '';
  inputNombreEl.focus();
}

function cerrarModal() {
  modalEl.classList.add('hidden');
}

async function crearCarga() {
  const nombre = inputNombreEl.value.trim();
  if (!nombre) {
    showToast('Escribe un nombre para la carga.', 'error');
    return;
  }

  const { data, error } = await supabaseClient
    .from('cargas')
    .insert({ nombre })
    .select('id')
    .single();

  if (error) {
    console.error(error);
    showToast('No se pudo crear la carga.', 'error');
    return;
  }

  window.location.href = `carga.html?id=${data.id}`;
}

async function borrarCarga(id, nombre) {
  if (!window.confirm(`¿Borrar la carga "${nombre}"? Esta acción no se puede deshacer.`)) return;

  const { error } = await supabaseClient.from('cargas').delete().eq('id', id);
  if (error) {
    console.error(error);
    showToast('No se pudo borrar la carga.', 'error');
    return;
  }

  showToast('Carga borrada.', 'success');
  cargarDeck();
}

document.getElementById('btn-cancelar-nueva').addEventListener('click', cerrarModal);
document.getElementById('btn-crear-nueva').addEventListener('click', crearCarga);
modalEl.addEventListener('click', (e) => { if (e.target === modalEl) cerrarModal(); });
inputNombreEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') crearCarga(); });

cargarDeck();
