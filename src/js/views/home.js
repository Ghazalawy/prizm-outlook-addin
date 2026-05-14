/**
 * Home view — grid of all available ERP actions.
 * Same set of actions surfaced by the Outlook menu, so users get the same
 * choices whether they pick from Outlook or open the add-in directly.
 */
import { el, mount } from '../ui.js';

const TILES = [
  { path: '/create-task',        icon: '✓', title: 'Create Task',        sub: 'New task from this email' },
  { path: '/create-opportunity', icon: '✯', title: 'Create Opportunity', sub: 'Convert email to opportunity' },
  { path: '/create-lead',        icon: '★', title: 'Create Lead',        sub: 'From sender details' },
  { path: '/create-ticket',      icon: '⚠', title: 'Create Ticket',      sub: 'Open support ticket' },
  { path: '/link-record',        icon: '⛓', title: 'Link to Record', sub: 'Attach to project/customer' },
  { path: '/lookup-sender',      icon: '⌕', title: 'Lookup Sender',      sub: 'Find related records' },
  { path: '/settings',           icon: '⚙', title: 'Settings',           sub: 'API key & ERP URL' },
];

export async function render() {
  const grid = el('div', { class: 'action-grid' },
    ...TILES.map((t) =>
      el('button', {
        class: 'action-tile',
        type: 'button',
        onclick: () => window.__prizmGo(t.path),
      },
        el('span', { class: 'action-tile__icon', text: t.icon }),
        el('span', { class: 'action-tile__title', text: t.title }),
        el('span', { class: 'action-tile__sub', text: t.sub }),
      )
    ),
  );

  mount(el('div', {},
    el('div', { class: 'card' },
      el('h3', { text: 'Pick an action' }),
      el('p', { text: 'These are also reachable directly from the Prizm ERP menu on any email.' }),
    ),
    grid,
  ));
}
