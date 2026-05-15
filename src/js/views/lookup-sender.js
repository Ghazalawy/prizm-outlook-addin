/**
 * Lookup Sender — comprehensive view of the email sender's status in Prizm:
 *  - The domain (canonical identifier for the company)
 *  - Whether the domain is registered as a customer / business partner / vendor
 *  - The matching contact in the system (if any)
 *  - Quick relationship summary for the company (counts of invoices, opps, ...)
 *  - Recent activity rows
 *  - Buttons to add a contact and/or a company, with the strict rule:
 *    "contact cannot be added unless the company is already registered"
 *
 * Add forms pre-fill from the email signature (name/phone/mobile) using a
 * regex heuristic — no AI calls.
 */
import { Office } from '../office.js';
import { Api, ApiError } from '../api.js';
import { Config } from '../config.js';
import { el, mount, field, row, banner, contextBanner } from '../ui.js';

function chipPill(label, kind) {
  return el('span', { class: `pill pill--${kind}`, text: label });
}

function renderRelationshipSummary(activity) {
  if (!activity || !Object.keys(activity).length) return null;
  const rows = [];
  Object.entries(activity).forEach(([k, v]) => {
    if (!v) return;
    rows.push(el('div', { class: 'kv' },
      el('span', { class: 'kv__key', text: k }),
      el('span', { class: 'kv__value', text: String(v) }),
    ));
  });
  if (!rows.length) return null;
  return el('div', { class: 'card' },
    el('h3', { text: 'Relationship summary' }),
    el('div', { class: 'kv-grid' }, ...rows),
  );
}

function renderBusiness(business, domain) {
  if (!business) return null;
  const items = [];
  if (business.customer) {
    items.push(el('div', { class: 'card' },
      el('h3', {}, 'Customer ', chipPill('registered', 'ok')),
      el('p', { text: business.customer.name }),
      business.customer.website ? el('p', { class: 'mono', text: business.customer.website }) : null,
      el('a', { href: business.customer.url, target: '_blank', rel: 'noopener', text: 'Open customer in ERP →' }),
    ));
  }
  if (business.partner) {
    const flags = [];
    if (business.partner.customer) flags.push('customer');
    if (business.partner.supplier) flags.push('supplier');
    items.push(el('div', { class: 'card' },
      el('h3', {}, 'Business partner ', chipPill(flags.join(' / ') || '—', 'info')),
      el('p', { text: business.partner.name }),
      el('a', { href: business.partner.url, target: '_blank', rel: 'noopener', text: 'Open partner in ERP →' }),
    ));
  }
  if (business.vendor) {
    items.push(el('div', { class: 'card' },
      el('h3', {}, 'Vendor / Supplier ', chipPill('registered', 'ok')),
      el('p', { text: business.vendor.name }),
      el('a', { href: business.vendor.url, target: '_blank', rel: 'noopener', text: 'Open vendor in ERP →' }),
    ));
  }
  if (!items.length) {
    items.push(el('div', { class: 'card' },
      el('h3', {}, `Domain ${domain || ''} `, chipPill('not registered', 'warn')),
      el('p', { class: 'muted', text: 'No customer, partner or vendor in the system matches this domain.' }),
    ));
  }
  return el('div', {}, ...items);
}

function showAddCompanyForm({ snap, domain, onCreated, status }) {
  const inputs = {
    company: el('input', { type: 'text', value: '' }),
    website: el('input', { type: 'url',  value: domain ? 'https://' + domain : '' }),
    phone:   el('input', { type: 'tel',  value: '' }),
    email:   el('input', { type: 'email', value: snap.from?.email || '' }),
    customerFlag: el('input', { type: 'checkbox', checked: true }),
    supplierFlag: el('input', { type: 'checkbox', checked: false }),
  };
  const submit = el('button', { class: 'btn', type: 'submit' }, 'Save company');
  const cancel = el('button', { class: 'btn btn--ghost', type: 'button', onclick: () => window.__prizmGo('/lookup-sender') }, 'Cancel');

  const form = el('form', {
    onsubmit: async (e) => {
      e.preventDefault();
      if (!inputs.company.value.trim()) {
        status.replaceChildren(banner('err', 'Company name is required.'));
        return;
      }
      submit.disabled = true; submit.textContent = 'Saving…';
      try {
        const res = await Api.createCustomer({
          company:        inputs.company.value.trim(),
          website:        inputs.website.value.trim(),
          phone:          inputs.phone.value.trim(),
          email:          inputs.email.value.trim(),
          mark_as_customer: inputs.customerFlag.checked,
          mark_as_supplier: inputs.supplierFlag.checked,
        });
        status.replaceChildren(banner('ok', `Company saved (#${res.id}). Refreshing lookup…`));
        if (typeof onCreated === 'function') onCreated(res);
      } catch (err) {
        const msg = err instanceof ApiError ? `Failed (${err.status||'net'}): ${err.message}` : `Failed: ${err.message}`;
        status.replaceChildren(banner('err', msg));
        submit.disabled = false; submit.textContent = 'Save company';
      }
    },
  },
    el('h3', { text: 'Add company' }),
    field('Company name', inputs.company, { required: true }),
    field('Website', inputs.website, { hint: 'Domain is the unique identifier. e.g. https://example.com' }),
    row(field('Phone', inputs.phone), field('Email', inputs.email)),
    el('div', { class: 'field__check' }, inputs.customerFlag, el('label', { text: 'Mark as customer' })),
    el('div', { class: 'field__check' }, inputs.supplierFlag, el('label', { text: 'Mark as supplier' })),
    el('div', { class: 'btn-row btn-row--split' }, submit, cancel),
  );
  return form;
}

function showAddContactForm({ snap, parsedSig, onCreated, status }) {
  const fromEmail = snap.from?.email || '';
  const fromName  = snap.from?.name || parsedSig.name || '';
  const [first = '', ...rest] = (fromName || '').split(/\s+/);
  const last = rest.join(' ');

  const inputs = {
    firstname: el('input', { type: 'text', value: first }),
    lastname:  el('input', { type: 'text', value: last }),
    email:     el('input', { type: 'email', value: fromEmail }),
    title:     el('input', { type: 'text', value: parsedSig.title || '' }),
    phone:     el('input', { type: 'tel',  value: parsedSig.phone || '' }),
    mobile:    el('input', { type: 'tel',  value: parsedSig.mobile || '' }),
  };
  const submit = el('button', { class: 'btn', type: 'submit' }, 'Save contact');
  const cancel = el('button', { class: 'btn btn--ghost', type: 'button', onclick: () => window.__prizmGo('/lookup-sender') }, 'Cancel');

  const form = el('form', {
    onsubmit: async (e) => {
      e.preventDefault();
      const required = ['firstname','lastname','email'];
      for (const k of required) {
        if (!inputs[k].value.trim()) {
          status.replaceChildren(banner('err', `${k} is required.`));
          return;
        }
      }
      submit.disabled = true; submit.textContent = 'Saving…';
      try {
        const res = await Api.createContact({
          firstname: inputs.firstname.value.trim(),
          lastname:  inputs.lastname.value.trim(),
          email:     inputs.email.value.trim(),
          title:     inputs.title.value.trim(),
          phone:     inputs.phone.value.trim() || inputs.mobile.value.trim(),
        });
        status.replaceChildren(banner('ok', `Contact saved under customer #${res.company_id}. Refreshing…`));
        if (typeof onCreated === 'function') onCreated(res);
      } catch (err) {
        if (err instanceof ApiError && err.status === 409) {
          status.replaceChildren(banner('err', err.message));
        } else {
          const msg = err instanceof ApiError ? `Failed (${err.status||'net'}): ${err.message}` : `Failed: ${err.message}`;
          status.replaceChildren(banner('err', msg));
        }
        submit.disabled = false; submit.textContent = 'Save contact';
      }
    },
  },
    el('h3', { text: 'Add contact under existing company' }),
    parsedSig.name ? el('p', { class: 'muted', text: 'Fields below were pre-filled from the email signature where possible — please review.' }) : null,
    row(field('First name', inputs.firstname, { required: true }), field('Last name', inputs.lastname, { required: true })),
    field('Email', inputs.email, { required: true, hint: 'The domain of this email must match the registered company.' }),
    field('Title', inputs.title),
    row(field('Phone', inputs.phone), field('Mobile', inputs.mobile)),
    el('div', { class: 'btn-row btn-row--split' }, submit, cancel),
  );
  return form;
}

export async function render() {
  mount(el('div', {}, banner('info', 'Looking up sender…')));
  const snap = await Office.snapshot();
  const email = snap.from?.email;

  if (!email) {
    mount(el('div', {}, contextBanner(snap), banner('err', 'No sender email available on this item.')));
    return;
  }

  // Parse signature once for the add-contact form pre-fill.
  const parsedSig = Office.parseSignature(snap.bodyText || '');

  async function load() {
    const status = el('div', {});
    let data;
    try {
      data = await Api.lookupContact(email);
    } catch (err) {
      mount(el('div', {}, contextBanner(snap),
        banner('err', err instanceof ApiError ? `Failed (${err.status||'net'}): ${err.message}` : `Failed: ${err.message}`)));
      return;
    }

    const container = el('div', {});
    container.appendChild(contextBanner(snap));

    // Domain status header
    const domain = data.domain || '';
    const business = data.business || {};
    container.appendChild(renderBusiness(business, domain));

    // Contact / lead
    const contactRows = [];
    if (data.contact) {
      contactRows.push(el('div', { class: 'card' },
        el('h3', {}, 'Contact ', chipPill('on file', 'ok')),
        el('p', { text: data.contact.name + (data.contact.industry ? ' · ' + data.contact.industry : '') }),
        el('a', { href: data.contact.url, target: '_blank', rel: 'noopener', text: 'Open in ERP →' }),
      ));
    }
    if (data.lead) {
      contactRows.push(el('div', { class: 'card' },
        el('h3', {}, 'Lead ', chipPill(data.lead.status || '—', 'info')),
        el('p', { text: data.lead.name }),
        el('a', { href: data.lead.url, target: '_blank', rel: 'noopener', text: 'Open lead in ERP →' }),
      ));
    }
    contactRows.forEach((c) => container.appendChild(c));

    // Activity summary
    const summary = renderRelationshipSummary(data.activity);
    if (summary) container.appendChild(summary);

    // Recent rows (existing block)
    if (data.recent?.length) {
      const list = el('div', {});
      data.recent.forEach((r) => {
        list.appendChild(el('div', { class: 'card' },
          el('h3', { text: `${r.type}: ${r.label}` }),
          r.sub ? el('p', { text: r.sub }) : null,
          el('a', { href: r.url || '#', target: '_blank', rel: 'noopener', text: 'Open in ERP →' }),
        ));
      });
      container.appendChild(el('h3', { text: 'Recent', style: { margin: '12px 0 6px' } }));
      container.appendChild(list);
    }

    // Action buttons — Add Contact is disabled when domain isn't registered.
    const domainRegistered = !!business.domain_registered;
    const hasCustomer       = !!business.customer;
    const addContactBtn = el('button', {
      class: 'btn',
      type: 'button',
      disabled: !hasCustomer || undefined,
      title: hasCustomer
        ? 'Add this person as a contact under the registered company'
        : 'Disabled — register the company first before adding a contact',
      onclick: () => mount(el('div', {},
        contextBanner(snap),
        showAddContactForm({ snap, parsedSig, status,
          onCreated: () => setTimeout(load, 400),
        }),
        status,
      )),
    }, 'Add contact');

    const addCompanyBtn = el('button', {
      class: 'btn btn--ghost',
      type: 'button',
      onclick: () => mount(el('div', {},
        contextBanner(snap),
        showAddCompanyForm({ snap, domain, status,
          onCreated: () => setTimeout(load, 400),
        }),
        status,
      )),
    }, hasCustomer ? 'Add another company for this domain' : 'Add company');

    container.appendChild(el('div', { class: 'btn-row btn-row--split' }, addContactBtn, addCompanyBtn));

    if (!hasCustomer) {
      container.appendChild(el('div', { class: 'field__hint', style: { marginTop: '6px' } },
        '⚠ Contacts can only be added under a registered company. Add the company first — the domain becomes its unique identifier.'));
    }

    container.appendChild(status);
    mount(container);
  }

  await load();
}
