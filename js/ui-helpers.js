// Utilidades compartidas entre index.html y carga.html: toasts no bloqueantes
// (reemplazan los alert() de la versión anterior) y helpers de formato.

function showToast(message, type = 'info') {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast toast--${type}`;
  toast.textContent = message;
  container.appendChild(toast);

  setTimeout(() => {
    toast.classList.add('toast--out');
    setTimeout(() => toast.remove(), 200);
  }, 3200);
}

function formatFecha(iso) {
  if (!iso) return '';
  return new Intl.DateTimeFormat('es-MX', { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(iso));
}

function nivelAvance(valor) {
  if (valor >= 80) return 'success';
  if (valor >= 50) return 'warning';
  return 'danger';
}

function debounce(fn, wait) {
  const timers = new Map();
  return (key, ...args) => {
    clearTimeout(timers.get(key));
    timers.set(key, setTimeout(() => fn(key, ...args), wait));
  };
}
