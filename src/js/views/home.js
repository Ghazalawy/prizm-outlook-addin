/**
 * Home view — grid of all available ERP actions.
 *
 * Each tile checks the v2 email-link log to see if its action has already
 * been performed on the current email. If so, the tile turns green and
 * shows a "Created on <date>" caption (or "Linked to N records" etc.).
 *
 * Same set of actions surfaced by the Outlook menu, so users get the same
 * choices whether they pick from Outlook or open the add-in directly.
 */
import { Office } from '../office.js';
import { Api } from '../api.js';
import { el, mount } from '../ui.js';

const TILES = [
  { path: '/create-task',        icon: '✓', title: 'Create Task',        sub: 'New task from this email',     doneKey: 'task' },
  { path: '/create-opportunity', icon: '✯', title: 'Create Opportunity', sub: 'Convert email to opportunity', doneKey: 'opportunity' },
  { path: '/create-lead',        icon: '★', title: 'Create Lead',        sub: 'From sender details',          doneKey: 'lead' },
  { path: '/create-ticket',      icon: '⚠', title: 'Create Ticket',      sub: 'Open support ticket',          doneKey: 'ticket' },
  // 'Link to Record' rows are stored with the target's rel_type, so the
  // marker is "linked to N records" rather than a single done flag.
  { path: '/link-record',        icon: '⛓', title: 'Link to Record',    sub: 'Attach to project/customer',   doneKey: '__link' },
  { path: '/lookup-sender',      icon: '⌕', title: 'Lookup Sender',      sub: 'Find related records' },
  { path: '/settings',           icon: '⚙', title: 'Settings',           sub: 'API key & ERP URL' },
];

function fmtWhen(iso) {
  if (!iso) return '';
  const d = new Date(iso.replace(' ', 'T') + 'Z');
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
}

function captionFor(doneKey, actions) {
  if (!actions) return null;
  if (doneKey === '__link') {
    // Anything that's NOT a create-action counts as a "link" target.
    const linkTypes = ['project', 'customer', 'lead', 'opportunity', 'invoice', 'estimate', 'contract'];
    const links = linkTypes.flatMap((t) => actions[t] || []);
    if (!links.length) return null;
    const newest = links.reduce((a, b) => (a.created_at > b.created_at ? a : b));
    return `Linked to ${links.length} record${links.length > 1 ? 's' : ''} · last ${fmtWhen(newest.created_at)}`;
  }
  const list = actions[doneKey];
  if (!list?.length) return null;
  const newest = list[list.length - 1]; // backend ordered by id asc
  const prefix = list.length === 1 ? '#' + newest.id : `${list.length}× last #${newest.id}`;
  return `${prefix} · created ${fmtWhen(newest.created_at)}`;
}

export async function render() {
  // Render the grid synchronously first so the user sees something even if
  // the email-status round-trip is slow or fails.
  let mounted;
  function paint(actions) {
    const grid = el('div', { class: 'action-grid' },
      ...TILES.map((t) => {
        const caption = t.doneKey ? captionFor(t.doneKey, actions) : null;
        const done = !!caption;
        return el('button', {
          class: done ? 'action-tile action-tile--done' : 'action-tile',
          type: 'button',
          onclick: () => window.__prizmGo(t.path),
        },
          el('span', { class: 'action-tile__icon', text: t.icon }),
          el('span', { class: 'action-tile__title', text: t.title }),
          el('span', { class: 'action-tile__sub', text: t.sub }),
          caption ? el('span', { class: 'action-tile__done', text: caption }) : null,
        );
      }),
    );

    const node = el('div', {},
      el('div', { class: 'card' },
        el('h3', { text: 'Pick an action' }),
        el('p', { text: 'These are also reachable directly from the Prizm ERP menu on any email.' }),
      ),
      grid,
    );
    mount(node);
    mounted = node;
  }

  paint({});

  // Fetch in the background; re-paint with done-state when it lands.
  let messageId = null;
  try {
    const snap = await Office.snapshot();
    messageId = snap?.internetMessageId || null;
  } catch (_) { /* not in Outlook, no email context */ }
  if (!messageId) return;

  try {
    const status = await Api.emailStatus(messageId);
    if (status && status.actions) paint(status.actions);
  } catch (_) {
    // Silent — if the user has no key yet they were already gated; if the
    // endpoint is temporarily down, the plain grid is still useful.
  }
}
