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

  /**
   * Detect whether we're in compose mode (user is drafting/replying/forwarding)
   * vs read mode (user opened an existing message). In compose mode subject is
   * an object with getAsync; in read mode it's a plain string. Same for to/cc.
   */
  isComposeMode() {
    const item = this.item();
    if (!item) return false;
    // Compose: item.subject has getAsync (function-like object)
    return !!(item.subject && typeof item.subject === 'object' && typeof item.subject.getAsync === 'function');
  },

  userProfile() {
    return window.Office?.context?.mailbox?.userProfile || {
      displayName: 'Local Test User',
      emailAddress: 'local.test@example.com',
    };
  },

  /**
   * Best-effort signature parser. No AI — pure regex / heuristics.
   * Given the email body text, tries to extract:
   *   - name        (capitalized words near a "regards"/"sincerely" closing)
   *   - title       (the line immediately after the name, before the company)
   *   - phones      (international or local digit clusters)
   *   - mobiles     (phone lines labeled mobile/cell)
   *
   * Returns {} when nothing usable is found — caller should treat fields as
   * suggestions only, never autosubmit.
   */
  parseSignature(bodyText) {
    const out = { name: '', title: '', phone: '', mobile: '' };
    if (!bodyText) return out;
    const lines = bodyText.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);

    // Find a sign-off line and look at the next non-empty lines as candidates
    const signoffRe = /^(best regards?|kind regards?|regards|sincerely|cheers|thanks(?:\s+(?:and\s+)?(?:regards|bests?))?|yours\s+truly)[,.]?\s*$/i;
    let signoffIdx = -1;
    for (let i = 0; i < lines.length; i++) {
      if (signoffRe.test(lines[i])) { signoffIdx = i; break; }
    }
    const tail = signoffIdx >= 0 ? lines.slice(signoffIdx + 1, signoffIdx + 12) : lines.slice(-12);

    // Name = first line in the tail that looks like a person's name (2-4
    // capitalized words, no @, no digits)
    const nameRe = /^([A-Z][a-z'’\-]+(?:\s+[A-Z][a-z'’\-]+){1,3})\s*$/;
    for (const ln of tail) {
      const m = ln.match(nameRe);
      if (m) { out.name = m[1]; break; }
    }

    // Title — first non-name line after the name candidate
    if (out.name) {
      const nameIdx = tail.findIndex((l) => l.startsWith(out.name));
      const afterName = tail.slice(nameIdx + 1, nameIdx + 4);
      for (const ln of afterName) {
        if (!/[@+\d]/.test(ln) && ln.length < 80) { out.title = ln; break; }
      }
    }

    // Phones: international "+" prefix or 7+ digit clusters with separators
    const phoneRe = /(\+?\d[\d\s\-().]{6,}\d)/g;
    const allPhones = [];
    for (const ln of tail) {
      const isMobile = /mob(?:ile)?|cell/i.test(ln);
      let m;
      while ((m = phoneRe.exec(ln)) !== null) {
        const num = m[1].replace(/[^\d+]/g, '');
        if (num.length < 7) continue;
        allPhones.push({ num, isMobile });
      }
    }
    const mobile = allPhones.find((p) => p.isMobile);
    if (mobile)            out.mobile = mobile.num;
    const landline = allPhones.find((p) => !p.isMobile && p.num !== out.mobile);
    if (landline)          out.phone = landline.num;
    if (!out.phone && !out.mobile && allPhones[0]) out.phone = allPhones[0].num;

    return out;
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
  getTo() { return this._recipients('to'); },
  /** Recipients: CC array (read or compose). */
  getCc() { return this._recipients('cc'); },

  _recipients(field) {
    return new Promise((resolve) => {
      const item = this.item();
      const ref = item ? item[field] : null;
      if (!ref) { resolve([]); return; }
      if (Array.isArray(ref)) {
        resolve(ref.map((r) => ({ name: r.displayName || '', email: r.emailAddress || '' })));
        return;
      }
      ref.getAsync((r) => {
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
    const composeMode = this.isComposeMode();
    const [subject, body, to, cc] = await Promise.all([
      this.getSubject(),
      this.getBodyText(),
      this.getTo(),
      this.getCc(),
    ]);
    const restItemId = this.getRestItemId();
    return {
      mode: composeMode ? 'compose' : 'read',
      subject,
      bodyText: body,
      bodyExcerpt: (body || '').slice(0, 500),
      from: this.getFrom(),
      to,
      cc,
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
