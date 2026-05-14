/**
 * Look up the email sender in ERP — show related records and quick deep links.
 * GET {apiBase}/outlook/lookup?email=
 */
import { Office } from '../office.js';
import { Api, ApiError } from '../api.js';
import { Config } from '../config.js';
import { el, mount, banner, contextBanner, fmtBytes } from '../ui.js';

export async function render() {
  mount(el('div', {}, banner('info', 'Looking up sender...')));
  const snap = await Office.snapshot();
  const email = snap.from?.email;

  if (!email) {
    mount(el('div', {}, contextBanner(snap), banner('err', 'No sender email available on this item.')));
    return;
  }

  try {
    const data = await Api.lookupContact(email);
    const container = el('div', {});

    container.appendChild(contextBanner(snap));

    if (!data || (!data.customer && !data.lead && !(data.recent?.length))) {
      container.appendChild(banner('info', `No records found for ${email}.`));
      container.appendChild(el('div', { class: 'btn-row' },
        el('button', { class: 'btn btn--ghost btn--block', type: 'button', onclick: () => window.__prizmGo('/create-lead') }, 'Create lead from sender'),
      ));
      mount(container);
      return;
    }

    if (data.customer) {
      const url = data.customer.url || `${Config.get('erpBase')}/admin/clients/client/${data.customer.id}`;
      container.appendChild(el('div', { class: 'card' },
        el('h3', { text: `Customer: ${data.customer.name}` }),
        el('p', { text: data.customer.industry || '' }),
        el('a', { href: url, target: '_blank', rel: 'noopener', text: 'Open customer in ERP' }),
      ));
    }
    if (data.lead) {
      const url = data.lead.url || `${Config.get('erpBase')}/admin/leads/index/${data.lead.id}`;
      container.appendChild(el('div', { class: 'card' },
        el('h3', { text: `Lead: ${data.lead.name}` }),
        el('p', { text: data.lead.status || '' }),
        el('a', { href: url, target: '_blank', rel: 'noopener', text: 'Open lead in ERP' }),
      ));
    }
    if (data.recent?.length) {
      const list = el('div', {});
      data.recent.forEach((r) => {
        const url = r.url || '#';
        list.appendChild(el('div', { class: 'card' },
          el('h3', { text: `${r.type}: ${r.label}` }),
          r.sub ? el('p', { text: r.sub }) : null,
          el('a', { href: url, target: '_blank', rel: 'noopener', text: 'Open in ERP' }),
        ));
      });
      container.appendChild(el('h3', { text: 'Recent activity', style: { margin: '12px 0 6px' } }));
      container.appendChild(list);
    }

    container.appendChild(el('div', { class: 'btn-row' },
      el('button', { class: 'btn btn--ghost', type: 'button', onclick: () => window.__prizmGo('/home') }, 'Home'),
      el('button', { class: 'btn', type: 'button', onclick: () => window.__prizmGo('/link-record') }, 'Link this email'),
    ));

    mount(container);
  } catch (err) {
    const msg = err instanceof ApiError ? `Failed (${err.status||'net'}): ${err.message}` : `Failed: ${err.message}`;
    mount(el('div', {}, contextBanner(snap), banner('err', msg)));
  }
}
