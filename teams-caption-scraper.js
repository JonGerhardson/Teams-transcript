/* Teams live-caption scraper — paste into DevTools console on teams.microsoft.com
 *
 * Prereq: turn live captions ON in the meeting first
 *   (... More  ->  Language and speech  ->  Turn on live captions)
 *
 * Then:  TeamsCaps.text()   -> transcript as a string
 *        TeamsCaps.save()   -> download .txt + .json
 *        TeamsCaps.copy()   -> copy to clipboard
 *        TeamsCaps.stop()   -> stop polling
 *        TeamsCaps.debug()  -> dump caption DOM (if capture looks broken)
 */
(() => {
  // Re-pasting replaces the running instance. stop() flushes pending lines to
  // localStorage, which the new instance restores below, so nothing is lost.
  if (window.TeamsCaps) {
    try { window.TeamsCaps.stop(); } catch {}
    delete window.TeamsCaps;
    console.log('[caps] replaced previous instance');
  }

  const KEY = 'teamsCaps.v1';
  const POLL_MS = 400;
  const AUTOSAVE_MIN = 5; // minutes between autosaves; 0 disables

  const CONTAINERS = [
    '[data-tid="closed-caption-v2-window"]',
    '[data-tid="closed-caption-v2-window-wrapper"]',
    '[data-tid="closed-caption-renderer-wrapper"]',
    '[data-tid="closed-captions-renderer"]',
    '[class*="closedCaption"]',
    '[class*="closed-caption"]',
  ];
  const LINES = [
    '[data-tid="closed-caption-message"]',
    '[data-tid="closed-caption-message-content"]',
    '[data-tid="closed-caption-text"]',
  ];

  const AUTHOR = [
    '[data-tid="author"]',
    '[data-tid="closed-caption-author"]',
    '[class*="author"]',
    '[class*="Author"]',
  ].join(',');

  const state = {
    entries: [], live: new Map(), timer: null, lastSpeaker: '',
    autosaveTimer: null, handle: null, jsonHandle: null,
  };

  // restore anything from a previous run / tab reload
  try {
    const prev = JSON.parse(localStorage.getItem(KEY) || '[]');
    if (Array.isArray(prev) && prev.length) {
      state.entries = prev;
      console.log(`[caps] restored ${prev.length} prior lines (TeamsCaps.reset() to clear)`);
    }
  } catch {}

  const findContainer = () => {
    for (const s of CONTAINERS) {
      const el = document.querySelector(s);
      if (el) return el;
    }
    return null;
  };

  const findLines = (root) => {
    for (const s of LINES) {
      const n = root.querySelectorAll(s);
      if (n.length) return [...n];
    }
    // fallback: any element holding an author chip, deepest match wins
    const cands = [...root.querySelectorAll('*')].filter(
      (e) => e.querySelector('[data-tid="author"]') && e.textContent.trim()
    );
    return cands.filter((e) => !cands.some((o) => o !== e && e.contains(o)));
  };

  // The author chip is frequently a sibling or grandparent of the text node,
  // not a descendant — so walk up a few levels before giving up.
  const findSpeaker = (el) => {
    let node = el;
    for (let i = 0; i < 4 && node; i++, node = node.parentElement) {
      const a = node.querySelector?.(AUTHOR);
      const name = a?.textContent.trim();
      if (name) return name;
    }
    node = el;
    for (let i = 0; i < 4 && node; i++, node = node.parentElement) {
      const av = node.querySelector?.('[data-tid="avatar"],[class*="avatar"],[class*="Avatar"]');
      const label = av?.getAttribute('aria-label') || av?.getAttribute('title');
      if (label) return label.replace(/\s*\((?:guest|external)\)\s*$/i, '').trim();
    }
    return '';
  };

  const parse = (el) => {
    const speaker = findSpeaker(el);
    const textEl = el.querySelector('[data-tid="closed-caption-text"]') || el;
    const clone = textEl.cloneNode(true);
    clone.querySelectorAll(AUTHOR).forEach((n) => n.remove());
    let text = clone.textContent.trim();
    if (speaker && text.startsWith(speaker)) text = text.slice(speaker.length).trim();
    return { speaker, text: text.replace(/\s+/g, ' ') };
  };

  const persist = () => {
    try { localStorage.setItem(KEY, JSON.stringify(state.entries)); } catch {}
  };

  const commit = (rec) => {
    if (!rec.text) return;
    const last = state.entries[state.entries.length - 1];
    // Teams refines a line in place, and long utterances get re-emitted —
    // collapse when the new text is a superset of the previous one.
    if (last && last.speaker === rec.speaker) {
      if (rec.text.startsWith(last.text)) { state.entries[state.entries.length - 1] = rec; persist(); return; }
      if (last.text.startsWith(rec.text)) return;
    }
    state.entries.push(rec);
    persist();
  };

  const tick = () => {
    const root = findContainer();
    // Meeting ended / captions toggled off: the container is gone, so anything
    // still live will never scroll out. Commit it now rather than losing it.
    if (!root) { flushLive(); return; }
    const els = new Set(findLines(root));

    for (const el of els) {
      const { speaker, text } = parse(el);
      if (!text) continue;
      const rec = state.live.get(el) || { t: new Date().toISOString(), speaker: '', text: '' };
      // Teams drops the name on continuation lines — inherit the last one seen.
      rec.speaker = speaker || rec.speaker || state.lastSpeaker;
      if (speaker) state.lastSpeaker = speaker;
      rec.text = text;
      state.live.set(el, rec);
    }
    // lines that scrolled out of the DOM are final
    for (const [el, rec] of state.live) {
      if (!els.has(el)) { commit(rec); state.live.delete(el); }
    }
  };

  const flushLive = () => {
    for (const [el, rec] of state.live) { commit(rec); state.live.delete(el); }
  };

  const download = (name, body, type = 'text/plain') => {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([body], { type }));
    a.download = name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  };

  const renderText = () => {
    flushLive();
    return state.entries
      .map((e) => `[${e.t.slice(11, 19)}] ${e.speaker ? e.speaker + ': ' : ''}${e.text}`)
      .join('\n');
  };

  const renderJson = () => { flushLive(); return JSON.stringify(state.entries, null, 2); };

  const stamp = () => new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);

  const writeHandle = async (handle, body) => {
    const w = await handle.createWritable();
    await w.write(body);
    await w.close();
  };

  // Chromium can hold writable handles and rewrite the same files forever.
  // Firefox has no equivalent, so it gets fresh timestamped downloads instead.
  const autosaveTick = async () => {
    const body = renderText();
    if (!body) return;
    const json = renderJson();

    if (state.handle) {
      try {
        await writeHandle(state.handle, body);
        if (state.jsonHandle) await writeHandle(state.jsonHandle, json);
        console.log(`[caps] autosaved ${state.entries.length} lines`);
        return;
      } catch (e) {
        console.warn('[caps] handle write failed, reverting to downloads', e);
        state.handle = state.jsonHandle = null;
      }
    }
    const s = stamp();
    download(`teams-captions-${s}.txt`, body);
    download(`teams-captions-${s}.json`, json, 'application/json');
    console.log(`[caps] autosaved ${state.entries.length} lines (download)`);
  };

  window.TeamsCaps = {
    get entries() { flushLive(); return state.entries; },
    text: renderText,
    // Chromium only: pick the files once, then every autosave rewrites them.
    // Pass false to skip the JSON and get a single dialog.
    async pickFile(withJson = true) {
      if (!window.showSaveFilePicker) { console.warn('[caps] not supported in this browser'); return; }
      const s = stamp();
      state.handle = await showSaveFilePicker({
        suggestedName: `teams-captions-${s}.txt`,
        types: [{ description: 'Text', accept: { 'text/plain': ['.txt'] } }],
      });
      if (withJson) {
        // A second picker can be refused for lack of user activation — the txt
        // handle still stands, and JSON just falls back to periodic downloads.
        try {
          state.jsonHandle = await showSaveFilePicker({
            suggestedName: `teams-captions-${s}.json`,
            types: [{ description: 'JSON', accept: { 'application/json': ['.json'] } }],
          });
        } catch (e) { console.warn('[caps] JSON handle not granted, txt only', e); }
      }
      await autosaveTick();
    },
    autosave(min = AUTOSAVE_MIN) {
      clearInterval(state.autosaveTimer);
      state.autosaveTimer = setInterval(autosaveTick, min * 60000);
      console.log(`[caps] autosaving every ${min} min`);
    },
    autosaveOff() { clearInterval(state.autosaveTimer); console.log('[caps] autosave off'); },
    save() {
      const body = renderText();
      const s = stamp();
      download(`teams-captions-${s}.txt`, body);
      download(`teams-captions-${s}.json`, JSON.stringify(state.entries, null, 2), 'application/json');
      console.log(`[caps] saved ${state.entries.length} lines`);
    },
    async copy() { await navigator.clipboard.writeText(this.text()); console.log('[caps] copied'); },
    stop() {
      clearInterval(state.timer);
      clearInterval(state.autosaveTimer);
      flushLive();
      console.log(`[caps] stopped — ${state.entries.length} lines held`);
    },
    reset() { state.entries = []; state.live.clear(); persist(); console.log('[caps] cleared'); },
    debug() {
      const root = findContainer();
      if (!root) { console.warn('[caps] no caption container found — are captions turned on?'); return; }
      const lines = findLines(root);
      console.log('[caps] container:', root, '| lines:', lines.length);
      console.table(lines.map((el) => ({ speaker: parse(el).speaker || '(none)', text: parse(el).text.slice(0, 60) })));
      lines.forEach((el) => console.log('[caps] line el:', el, '| author match:', el.closest('*')?.querySelector(AUTHOR)));
      console.log(root.outerHTML.slice(0, 4000));
    },
  };

  // Last-chance persist if the tab is closed or navigated away mid-meeting.
  addEventListener('pagehide', flushLive);
  addEventListener('beforeunload', flushLive);
  addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') flushLive(); });

  state.timer = setInterval(tick, POLL_MS);
  if (AUTOSAVE_MIN > 0) window.TeamsCaps.autosave(AUTOSAVE_MIN);
  console.log('[caps] capturing… TeamsCaps.save() when done');
})();
