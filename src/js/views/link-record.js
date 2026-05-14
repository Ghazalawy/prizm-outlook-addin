/**
 * Link this email (and optionally attachments) to an existing ERP record.
 * Search-as-you-type against GET {apiBase}/outlook/search?type=&q=
 */
import { Office } from '../office.js';
import { Api, ApiError } from '../api.js';
import { el, mount, field, row, banner, contextBanner } from '../ui.js';

const TYPES = [
  { v: 'project',     l: 'Project' },
  { v: 'customer',    l: 'Customer' },
  { v: 'lead',        l: 'Lead' },
  { v: 'opportunity', l: 'Opportunity' },
  { v: 'invoice',     l: 'Invoice' },
  { v: 'estimate',    l: 'Estimate' },
  { v: 'contract',    l: 'Contract' },
  { v: 'ticket',      l: 'Ticket' },
];

export async function render() {
  mount(el('div', {}, banner('info', 'Loading email context...')));
  const snap = await Office.snapshot();

  let debounceTimer = null;
  let chosen = null;

  const typeSel = el('select', {}, ...TYPES.map((t) => el('option', { value: t.v, text: t.l })));
  const search  = el('input', { type: 'search', placeholder: 'Type to search...' });
  const results = el('div', { class: 'results' });
  const attachEmail = el('input', { type: 'checkbox', checked: true });
  const attachFiles = el('input', { type: 'checkbox', checked: !!snap.attachments?.length });
  const status = el('div', {});
  const submit = el('button', { class: 'btn btn--block', type: 'submit' }, 'Link');
  const cancel = el('button', { class: 'btn btn--ghost', type: 'button', onclick: () => window.__prizmGo('/home') }, 'Cancel');

  function renderResults(list) {
    results.replaceChildren();
    if (!list?.length) {
      results.appendChild(el('div', { class: 'field__hint', text: 'No matches.' }));
      return;
    }
    list.forEach((r) => {
      const tile = el('button', {
        class: 'action-tile',
        type: 'button',
        style: { gridColumn: '1 / -1' },
        onclick: () => {
          chosen = { type: typeSel.value, id: r.id, label: r.label };
          renderResults([r]);
          tile.style.borderColor = 'var(--prizm-primary)';
        },
      },
        el('span', { class: 'action-tile__title', text: r.label }),
        r.sub ? el('span', { class: 'action-tile__sub', text: r.sub }) : null,
      );
      results.appendChild(tile);
    });
  }

  search.addEventListener('input', () => {
    chosen = null;
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(async () => {
      const q = search.value.trim();
      if (!q) { results.replaceChildren(); return; }
      try {
        const list = await Api.search(typeSel.value, q);
        renderResults(list || []);
      } catch (err) {
        results.replaceChildren(banner('err', `Search failed: ${err.message}`));
      }
    }, 300);
  });

  const form = el('form', {
    onsubmit: async (e) => {
      e.preventDefault();
      if (!chosen) {
        status.replaceChildren(banner('err', 'Pick a record first.'));
        return;
      }
      submit.disabled = true; submit.textContent = 'Linking...'; status.replaceChildren();
      try {
        await Api.linkEmail({
          target: { type: chosen.type, id: chosen.id },
          email: {
            itemId: snap.itemId,
            internetMessageId: snap.internetMessageId,
            subject: snap.subject,
            from: snap.from,
            to: snap.to,
            receivedAt: snap.receivedAt,
            attachEmailAsEml: attachEmail.checked,
            attachments: attachFiles.checked ? snap.attachments : [],
          },
        });
        status.replaceChildren(banner('ok', `Email linked to ${chosen.type} "${chosen.label}".`));
        submit.textContent = 'Done'; submit.disabled = true;
      } catch (err) {
        const msg = err instanceof ApiError ? `Failed (${err.status||'net'}): ${err.message}` : `Failed: ${err.message}`;
        status.replaceChildren(banner('err', msg));
        submit.disabled = false; submit.textContent = 'Retry';
      }
    },
  },
    contextBanner(snap),
    row(field('Record type', typeSel), field('Search', search)),
    results,
    el('div', { class: 'field__check' }, attachEmail, el('label', { text: 'Attach the email as .eml' })),
    el('div', { class: 'field__check' }, attachFiles, el('label', { text: `Attach ${snap.attachments?.length || 0} attachment(s)` })),
    el('div', { class: 'btn-row' }, cancel, submit),
    status,
  );

  mount(form);
}
