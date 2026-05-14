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

/**
 * Searchable multi-select with chip-style removable selections.
 *
 * @param {object} opts
 * @param {Array<{id:any,name:string}>} opts.options - choices
 * @param {Array<any>} [opts.selected=[]] - ids initially selected
 * @param {string} [opts.placeholder='Search...']
 * @returns {{node: HTMLElement, getSelected: () => any[]}}
 */
export function chipsPicker({ options = [], selected = [], placeholder = 'Search...', allowCreate = false } = {}) {
  const picked = new Map();
  options.forEach((o) => { if (selected.includes(o.id)) picked.set(o.id, o.name); });

  const chipsRow = el('div', { class: 'chips__row' });
  const input    = el('input', { type: 'search', placeholder, class: 'chips__input' });
  const dropdown = el('div', { class: 'chips__dropdown', hidden: true });
  const wrap     = el('div', { class: 'chips' }, chipsRow, input, dropdown);

  function renderChips() {
    chipsRow.replaceChildren();
    [...picked.entries()].forEach(([id, name]) => {
      chipsRow.appendChild(el('span', { class: 'chip' },
        el('span', { class: 'chip__name', text: name }),
        el('button', {
          type: 'button', class: 'chip__x', 'aria-label': `Remove ${name}`,
          onclick: () => { picked.delete(id); renderChips(); },
        }, '×'),
      ));
    });
  }

  function pick(id, name) {
    picked.set(id, name);
    input.value = '';
    renderChips();
    renderDropdown('');
  }

  function renderDropdown(q) {
    dropdown.replaceChildren();
    const needle = (q || '').trim().toLowerCase();
    const matches = options
      .filter((o) => !picked.has(o.id))
      .filter((o) => !needle || (o.name || '').toLowerCase().includes(needle))
      .slice(0, 50);

    matches.forEach((o) => {
      dropdown.appendChild(el('button', {
        type: 'button', class: 'chips__opt',
        onmousedown: (e) => { e.preventDefault(); pick(o.id, o.name); },
      }, o.name));
    });

    const exact = options.some((o) => (o.name || '').toLowerCase() === needle);
    if (allowCreate && needle && !exact && !picked.has(needle)) {
      dropdown.appendChild(el('button', {
        type: 'button', class: 'chips__opt chips__opt--create',
        onmousedown: (e) => { e.preventDefault(); pick(q.trim(), q.trim()); },
      }, `+ Create "${q.trim()}"`));
    }

    if (!matches.length && !(allowCreate && needle)) {
      dropdown.appendChild(el('div', { class: 'chips__empty', text: needle ? 'No matches.' : 'Start typing to filter.' }));
    }
  }

  input.addEventListener('focus', () => { renderDropdown(input.value); dropdown.hidden = false; });
  input.addEventListener('input', () => { renderDropdown(input.value); dropdown.hidden = false; });
  input.addEventListener('blur',  () => { setTimeout(() => { dropdown.hidden = true; }, 150); });
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      const needle = input.value.trim();
      if (!needle) return;
      const exact = options.find((o) => (o.name || '').toLowerCase() === needle.toLowerCase());
      if (exact) { pick(exact.id, exact.name); return; }
      if (allowCreate) pick(needle, needle);
    } else if (e.key === 'Backspace' && !input.value && picked.size) {
      // Quick remove: empty input + backspace pops the last chip
      const last = [...picked.keys()].pop();
      picked.delete(last);
      renderChips();
    }
  });

  renderChips();
  return { node: wrap, getSelected: () => [...picked.keys()] };
}

/**
 * Single-select async search picker.
 * Type → debounced async call → dropdown of results → click sets the chosen
 * record (and fills the input with its label). Re-typing clears the choice.
 *
 * @param {object} opts
 * @param {(q: string) => Promise<Array<{id:any,label:string,sub?:string}>>} opts.search
 * @param {string} [opts.placeholder='Search...']
 * @param {(r: object) => void} [opts.onChange]
 * @returns {{node: HTMLElement, get: () => object|null, clear: () => void, setPlaceholder: (s:string)=>void}}
 */
export function asyncSearchPicker({ search, placeholder = 'Search...', onChange } = {}) {
  const input    = el('input', { type: 'search', placeholder, class: 'chips__input' });
  const dropdown = el('div',   { class: 'chips__dropdown', hidden: true });
  const wrap     = el('div',   { class: 'chips' }, input, dropdown);

  let chosen = null;
  let debounceTimer = null;
  let reqId = 0;

  function publish(value) {
    chosen = value;
    if (typeof onChange === 'function') onChange(chosen);
  }

  function showResults(items) {
    dropdown.replaceChildren();
    if (!items?.length) {
      dropdown.appendChild(el('div', { class: 'chips__empty', text: 'No matches.' }));
      return;
    }
    items.forEach((r) => {
      const row = el('button', {
        type: 'button', class: 'chips__opt',
        onmousedown: (e) => {
          e.preventDefault();
          publish(r);
          input.value = r.label;
          dropdown.hidden = true;
        },
      },
        el('div', { class: 'chips__opt-title', text: r.label }),
        r.sub ? el('div', { class: 'chips__opt-sub', text: r.sub }) : null,
      );
      dropdown.appendChild(row);
    });
  }

  input.addEventListener('input', () => {
    publish(null);
    clearTimeout(debounceTimer);
    const q = input.value.trim();
    if (!q) { dropdown.hidden = true; return; }
    const myReq = ++reqId;
    debounceTimer = setTimeout(async () => {
      dropdown.hidden = false;
      dropdown.replaceChildren(el('div', { class: 'chips__empty', text: 'Searching…' }));
      try {
        const results = await search(q);
        if (myReq !== reqId) return; // stale response from earlier keystroke
        showResults(results);
      } catch (e) {
        if (myReq !== reqId) return;
        dropdown.replaceChildren(el('div', { class: 'chips__empty', text: 'Error: ' + e.message }));
      }
    }, 250);
  });
  input.addEventListener('focus', () => { if (input.value.trim()) dropdown.hidden = false; });
  input.addEventListener('blur',  () => { setTimeout(() => { dropdown.hidden = true; }, 150); });

  return {
    node: wrap,
    get: () => chosen,
    clear: () => { publish(null); input.value = ''; dropdown.hidden = true; },
    setPlaceholder: (s) => { input.placeholder = s; },
  };
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
