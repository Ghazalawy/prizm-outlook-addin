/**
 * Add-in entry point.
 * 1. Wait for Office.js
 * 2. Register routes (one per ERP action)
 * 3. Wire topbar + footer
 * 4. Start the router (renders the view that matches the URL hash, which the
 *    manifest sets when a menu item is clicked).
 */
import { Office } from './office.js';
import { Config } from './config.js';
import { registerRoute, startRouter, back, go } from './router.js';

function setupChrome() {
  document.getElementById('splash').hidden = true;
  document.getElementById('topbar').hidden = false;
  document.getElementById('view').hidden   = false;
  document.getElementById('footer').hidden = false;

  document.getElementById('backBtn').addEventListener('click', back);
  document.getElementById('reloadBtn').addEventListener('click', () => {
    // Force re-fetch of all assets — useful after a redeploy.
    window.location.reload();
  });

  const { erpBase, version } = Config.all();
  document.getElementById('footerVersion').textContent = `v${version}`;
  const link = document.getElementById('footerErpLink');
  link.href = erpBase;
  link.textContent = new URL(erpBase).hostname;
}

function registerRoutes() {
  registerRoute('/home',                () => import('./views/home.js'),               { title: 'Prizm ERP' });
  registerRoute('/create-task',         () => import('./views/create-task.js'),        { title: 'Create Task' });
  registerRoute('/create-opportunity',  () => import('./views/create-opportunity.js'), { title: 'Create Opportunity' });
  registerRoute('/create-lead',         () => import('./views/create-lead.js'),        { title: 'Create Lead' });
  registerRoute('/create-ticket',       () => import('./views/create-ticket.js'),      { title: 'Create Ticket' });
  registerRoute('/link-record',         () => import('./views/link-record.js'),        { title: 'Link to Record' });
  registerRoute('/lookup-sender',       () => import('./views/lookup-sender.js'),      { title: 'Look up sender' });
  registerRoute('/settings',            () => import('./views/settings.js'),           { title: 'Settings' });
}

async function main() {
  await Office.ready();
  setupChrome();
  registerRoutes();
  startRouter();
}

main().catch((err) => {
  console.error('Add-in boot failed', err);
  const splash = document.getElementById('splash');
  if (splash) {
    splash.innerHTML = `<div style="color:#b3261e;text-align:center;padding:16px">
      <strong>Failed to load.</strong><br/>${err?.message || err}
    </div>`;
  }
});

// Expose for views to navigate
window.__prizmGo = go;
