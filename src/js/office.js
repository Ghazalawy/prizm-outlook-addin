/**
 * Thin wrappers over Office.js. Returns Promises so views can await everything.
 * If Office.js isn't initialised yet, we return a "stub" item so the views still
 * render in a normal browser for local testing.
 */
export const Office = {
  ready() {
    return new Promise((resolve) => {
      if (typeof window.Office === 'undefined') {
        resolve({ host: 'browser', platform: 'browser' });
        return;
      }
      window.Office.onReady((info) => resolve(info));
    });
  },

  item() {
    if (typeof window.Office === 'undefined' || !window.Office.context?.mailbox?.item) {
      return null;
    }
    return window.Office.context.mailbox.item;
  },

  userProfile() {
    return window.Office?.context?.mailbox?.userProfile || {
      displayName: 'Local Test User',
      emailAddress: 'local.test@example.com',
    };
  },

  /** Get the email subject (works in read + compose modes). */
  getSubject() {
    return new Promise((resolve) => {
      const item = this.item();
      if (!item) { resolve(''); return; }
      if (typeof item.subject === 'string') { resolve(item.subject); return; }
      item.subject.getAsync((r) => resolve(r.status === 'succeeded' ? (r.value || '') : ''));
    });
  },

  /** Get the email body as plain text. */
  getBodyText() {
    return new Promise((resolve) => {
      const item = this.item();
      if (!item || !item.body) { resolve(''); return; }
      item.body.getAsync('text', (r) => resolve(r.status === 'succeeded' ? (r.value || '') : ''));
    });
  },

  /** From (sender) — read mode only; compose mode has no concept of "from". */
  getFrom() {
    const item = this.item();
    if (!item) return null;
    if (item.from) {
      return { name: item.from.displayName || '', email: item.from.emailAddress || '' };
    }
    return null;
  },

  /** Recipients: To array (read or compose). */
  getTo() {
    return new Promise((resolve) => {
      const item = this.item();
      if (!item || !item.to) { resolve([]); return; }
      if (Array.isArray(item.to)) {
        resolve(item.to.map((r) => ({ name: r.displayName || '', email: r.emailAddress || '' })));
        return;
      }
      item.to.getAsync((r) => {
        if (r.status === 'succeeded') {
          resolve((r.value || []).map((x) => ({ name: x.displayName || '', email: x.emailAddress || '' })));
        } else { resolve([]); }
      });
    });
  },

  /** Attachments list (read mode). */
  getAttachments() {
    const item = this.item();
    if (!item || !item.attachments) return [];
    return (item.attachments || []).map((a) => ({
      id: a.id,
      name: a.name,
      size: a.size,
      contentType: a.contentType,
      isInline: a.isInline,
    }));
  },

  /** Internet message id (used to link email to ERP record). */
  getInternetMessageId() {
    const item = this.item();
    return item?.internetMessageId || null;
  },

  /**
   * Get the item ID converted to REST (Graph) format.
   * Outlook gives us an EWS ID by default; Graph endpoints expect the v2 REST ID.
   * `convertToRestId` is a no-op when the ID is already in REST format.
   * Returns the raw itemId if Office.js isn't available (browser test mode).
   */
  getRestItemId() {
    const item = this.item();
    if (!item?.itemId) return null;
    const mbx = window.Office?.context?.mailbox;
    const enums = window.Office?.MailboxEnums;
    if (mbx?.convertToRestId && enums?.RestVersion?.v2_0) {
      try {
        return mbx.convertToRestId(item.itemId, enums.RestVersion.v2_0);
      } catch (_) { /* fall through */ }
    }
    return item.itemId;
  },

  /**
   * Get a token for the ERP backend to call Graph on behalf of the user
   * (only relevant if the ERP is set up as an Azure AD app). Falls back to null.
   */
  getCallbackToken() {
    return new Promise((resolve) => {
      const mbx = window.Office?.context?.mailbox;
      if (!mbx?.getCallbackTokenAsync) { resolve(null); return; }
      mbx.getCallbackTokenAsync({ isRest: true }, (r) => {
        resolve(r.status === 'succeeded' ? r.value : null);
      });
    });
  },

  /**
   * Standard email envelope every create-X view sends to the ERP.
   * Centralised so we don't drift between views.
   *
   * @param {object} snap  - result of snapshot()
   * @param {object} opts  - { attachEmailAsEml: bool, attachFiles: bool }
   */
  envelope(snap, opts = {}) {
    return {
      itemId: snap.itemId,                  // REST/Graph-compatible
      rawItemId: snap.rawItemId,            // EWS form for fallback
      internetMessageId: snap.internetMessageId,
      conversationId: snap.conversationId,
      subject: snap.subject,
      from: snap.from,
      to: snap.to,
      receivedAt: snap.receivedAt,
      mailbox: snap.mailbox,                // tells the ERP which user's mailbox to fetch from via Graph
      attachEmailAsEml: !!opts.attachEmailAsEml,
      attachments: opts.attachFiles ? (snap.attachments || []) : [],
    };
  },

  /** Build a snapshot of the current email for any "create from email" view. */
  async snapshot() {
    const item = this.item();
    const [subject, body, to] = await Promise.all([
      this.getSubject(),
      this.getBodyText(),
      this.getTo(),
    ]);
    const restItemId = this.getRestItemId();
    return {
      subject,
      bodyText: body,
      bodyExcerpt: (body || '').slice(0, 500),
      from: this.getFrom(),
      to,
      attachments: this.getAttachments(),
      internetMessageId: this.getInternetMessageId(),
      itemId: restItemId,
      rawItemId: item?.itemId || null,
      conversationId: item?.conversationId || null,
      receivedAt: item?.dateTimeCreated || null,
      mailbox: this.userProfile().emailAddress || null,
    };
  },
};
