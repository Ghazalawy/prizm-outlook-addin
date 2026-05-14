/**
 * Tiny DOM helpers — no framework. Keeps views terse.
 */

export function el(tag, props = {}, ...children) {
  const node = document.createElement(tag);
  for (const [key, value] of Object.entries(props || {})) {
    if (value === false || value === null || value === undefined) continue;
    if (key === 'class')   { node.className = value; continue; }
    if (key === 'html')    { node.innerHTML = value; continue; }
    if (key === 'text')    { node.textContent = value; continue; }
    if (key === 'style' && typeof value === 'object') {
      Object.assign(node.style, value); continue;
    }
    if (key.startsWith('on') && typeof value === 'function') {
      node.addEventListener(key.slice(2).toLowerCase(), value);
      continue;
    }
    if (key === 'attrs' && typeof value === 'object') {
      Object.entries(value).forEach(([k, v]) => node.setAttribute(k, v));
      continue;
    }
    if (value === true) { node.setAttribute(key, ''); continue; }
    node.setAttribute(key, value);
  }
  for (const child of children.flat(Infinity)) {
    if (child === null || child === undefined || child === false) continue;
    node.appendChild(child instanceof Node ? child : document.createTextNode(String(child)));
  }
  return node;
}

export function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); }

export function mount(viewNode) {
  const host = document.getElementById('view');
  clear(host);
  host.appendChild(viewNode);
}

export function setCrumb(text) {
  const c = document.getElementById('crumb');
  if (c) c.textContent = text;
}

export function banner(kind, message) {
  return el('div', { class: `banner banner--${kind}`, text: message });
}

export function field(label, input, { required = false, hint } = {}) {
  return el('div', { class: 'field' },
    el('label', {}, label, required ? el('span', { class: 'required', text: '*' }) : null),
    input,
    hint ? el('div', { class: 'field__hint', text: hint }) : null,
  );
}

export function row(...fields) {
  return el('div', { class: 'field__row' }, ...fields);
}

export function fmtBytes(n) {
  if (!n) return '';
  const u = ['B','KB','MB','GB'];
  let i = 0;
  while (n >= 1024 && i < u.length - 1) { n /= 1024; i++; }
  return `${n.toFixed(n >= 10 || i === 0 ? 0 : 1)} ${u[i]}`;
}

export function contextBanner(snapshot) {
  const rows = [];
  if (snapshot.from) {
    rows.push(el('div', { class: 'ctx__row' },
      el('span', { class: 'ctx__label', text: 'From' }),
      el('span', { class: 'ctx__value', text: snapshot.from.name ? `${snapshot.from.name} <${snapshot.from.email}>` : snapshot.from.email }),
    ));
  }
  if (snapshot.subject) {
    rows.push(el('div', { class: 'ctx__row' },
      el('span', { class: 'ctx__label', text: 'Subject' }),
      el('span', { class: 'ctx__value', text: snapshot.subject }),
    ));
  }
  if (snapshot.attachments?.length) {
    rows.push(el('div', { class: 'ctx__row' },
      el('span', { class: 'ctx__label', text: 'Files' }),
      el('span', { class: 'ctx__value', text: snapshot.attachments.map((a) => `${a.name} (${fmtBytes(a.size)})`).join(', ') }),
    ));
  }
  return el('div', { class: 'ctx' }, ...rows);
}
