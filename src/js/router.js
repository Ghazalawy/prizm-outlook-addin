/**
 * Hash router. Each manifest menu entry sets a hash like #/create-task
 * so adding a new view = add a route here and register a view module.
 */
const routes = new Map();
let currentTeardown = null;

export function registerRoute(path, loader, { title } = {}) {
  routes.set(path, { loader, title });
}

export function go(path) {
  if (window.location.hash !== `#${path}`) {
    window.location.hash = path;
  } else {
    render();
  }
}

export function back() {
  if (window.history.length > 1) {
    window.history.back();
  } else {
    go('/home');
  }
}

async function render() {
  const path = (window.location.hash || '#/home').replace(/^#/, '') || '/home';
  let route = routes.get(path);
  if (!route) {
    console.warn(`No route for ${path}, falling back to /home`);
    route = routes.get('/home');
    if (!route) {
      document.getElementById('view').textContent = `Unknown view: ${path}`;
      return;
    }
  }

  if (typeof currentTeardown === 'function') {
    try { currentTeardown(); } catch (_) { /* ignore */ }
    currentTeardown = null;
  }

  const viewModule = await route.loader();
  if (route.title) {
    const crumb = document.getElementById('crumb');
    if (crumb) crumb.textContent = route.title;
  }
  const teardown = await viewModule.render?.();
  if (typeof teardown === 'function') currentTeardown = teardown;
}

export function startRouter() {
  window.addEventListener('hashchange', render);
  render();
}
