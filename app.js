'use strict';

(() => {
  /* =====================================================================
   *  Constantes et utilitaires
   * ===================================================================== */
  const POLL_MS = 30000;       // vérifie le Google Sheet toutes les 30 s (onglet visible)
  const SAVE_DELAY = 1200;     // enregistre 1,2 s après la dernière modification
  const LS_CFG = 'arbitrage.connexion';
  const LS_CACHE = 'arbitrage.cache';

  const $ = (sel, el = document) => el.querySelector(sel);
  const $$ = (sel, el = document) => [...el.querySelectorAll(sel)];
  const money = new Intl.NumberFormat('fr-CA', { style: 'currency', currency: 'CAD' });
  const fmt = (n) => money.format(n || 0);
  const round2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
  const uid = () => Math.random().toString(36).slice(2, 10);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const pad = (n) => String(n).padStart(2, '0');
  const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };

  const TRASH = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14M10 11v6M14 11v6"/></svg>';

  /* =====================================================================
   *  État
   * ===================================================================== */
  const state = {
    rows: [], tariffs: [], arenas: [],
    version: null,          // dernière version du Google Sheet connue
    dirty: false,           // modifications locales pas encore enregistrées
    saving: false, pulling: false, resave: false,
    editGen: 0,             // compteur de modifications (détecte les éditions pendant un enregistrement)
    lastEditAt: 0,
    conflict: null,
    loaded: false,
    lastSync: null,
    filter: { mode: 'all', arena: '', month: '', q: '' }
  };
  let cfg = readCfg();
  let saveTimer = null;
  let toastTimer = null;
  let priceMap = new Map();
  let arenaMap = new Map();

  /* =====================================================================
   *  Calculs (mêmes règles que les formules du fichier Excel)
   * ===================================================================== */
  const descOf = (t) => `${t.cat} - ${t.double ? 'Double lettre' : 'Simple lettre'} - ${t.type}`;

  function buildLookups() {
    priceMap = new Map();
    state.tariffs.forEach((t) => { const d = descOf(t); if (!priceMap.has(d)) priceMap.set(d, t.cost); });
    arenaMap = new Map();
    state.arenas.forEach((a) => { if (!arenaMap.has(a.name)) arenaMap.set(a.name, a.travelCost); });
  }

  function calc(r) {
    const price = r.category ? (priceMap.get(r.category) ?? 0) : 0;
    const travel = r.travel ? (arenaMap.get(r.arena) ?? 0) : 0;
    const extra = Number(r.extra) || 0;
    return {
      price, travel, extra,
      total: round2(price + travel + extra),
      missingCategory: !!r.category && !priceMap.has(r.category),
      missingArena: !!r.arena && !arenaMap.has(r.arena)
    };
  }

  /* =====================================================================
   *  Données : conversion et cache local
   * ===================================================================== */
  function hydrate(d) {
    state.rows = (d.rows || []).map((r) => ({
      id: uid(), date: r.date || '', arena: r.arena || '', category: r.category || '',
      travel: !!r.travel, extra: Number(r.extra) || 0, counted: !!r.counted,
      comment: r.comment || '', tournament: !!r.tournament, paid: !!r.paid
    }));
    state.tariffs = (d.tariffs || []).map((t) => ({ id: uid(), cat: t.cat || '', double: !!t.double, type: t.type || '', cost: Number(t.cost) || 0 }));
    state.arenas = (d.arenas || []).map((a) => ({ id: uid(), name: a.name || '', travelCost: Number(a.travelCost) || 0 }));
  }

  function serialize() {
    return {
      rows: state.rows.map(({ id, ...r }) => ({ ...r, extra: Number(r.extra) || 0 })),
      tariffs: state.tariffs.map(({ id, ...t }) => ({ ...t, cost: Number(t.cost) || 0 })),
      arenas: state.arenas.map(({ id, ...a }) => ({ ...a, travelCost: Number(a.travelCost) || 0 }))
    };
  }

  function cache() {
    try { localStorage.setItem(LS_CACHE, JSON.stringify({ data: serialize(), version: state.version, dirty: state.dirty })); } catch (_) { /* stockage plein ou bloqué */ }
  }
  function readCache() {
    try { return JSON.parse(localStorage.getItem(LS_CACHE) || 'null'); } catch (_) { return null; }
  }

  /* =====================================================================
   *  Connexion (adresse Apps Script + mot de passe, gardés sur l'appareil)
   * ===================================================================== */
  function readCfg() {
    try { const c = JSON.parse(localStorage.getItem(LS_CFG) || 'null'); return c && c.url && c.token ? c : null; } catch (_) { return null; }
  }
  function writeCfg(c) { localStorage.setItem(LS_CFG, JSON.stringify(c)); }

  const b64enc = (s) => btoa(String.fromCharCode(...new TextEncoder().encode(s)));
  const b64dec = (s) => new TextDecoder().decode(Uint8Array.from(atob(s), (c) => c.charCodeAt(0)));

  function consumePairingHash() {
    const m = location.hash.match(/^#cfg=(.+)$/);
    if (!m) return;
    try {
      const o = JSON.parse(b64dec(m[1]));
      if (o.url && o.token) writeCfg({ url: o.url, token: o.token });
    } catch (_) { /* lien invalide */ }
    history.replaceState(null, '', location.pathname + location.search);
  }

  async function api(action, payload = {}, conf = cfg) {
    if (!conf) throw new Error('not-configured');
    const ctl = new AbortController();
    const timer = setTimeout(() => ctl.abort(), 30000);
    try {
      const res = await fetch(conf.url, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },  // évite la requête préliminaire CORS
        body: JSON.stringify({ token: conf.token, action, ...payload }),
        signal: ctl.signal
      });
      const text = await res.text();
      try { return JSON.parse(text); } catch (_) { throw new Error('bad-response'); }
    } finally { clearTimeout(timer); }
  }
  const isNetworkError = (e) => e && (e.name === 'AbortError' || e instanceof TypeError);

  /* =====================================================================
   *  Statut, message éphémère, bannière
   * ===================================================================== */
  function setStatus(kind, text) {
    const el = $('#status');
    el.dataset.kind = kind;
    const t = {
      ok: state.lastSync ? `Synchronisé à ${pad(state.lastSync.getHours())}:${pad(state.lastSync.getMinutes())}` : 'Synchronisé',
      saving: 'Enregistrement…',
      pending: 'Modifications à enregistrer',
      loading: 'Chargement…',
      offline: 'Hors ligne, en attente du réseau',
      error: 'Erreur de synchronisation',
      conflict: 'Conflit à résoudre',
      disconnected: 'Non connecté à Google Sheet'
    }[kind];
    $('#status-text').textContent = text || t;
  }

  function toast(msg, action) {
    const el = $('#toast');
    el.innerHTML = `<span>${esc(msg)}</span>`;
    if (action) {
      const b = document.createElement('button');
      b.type = 'button'; b.textContent = action.label;
      b.onclick = () => { hideToast(); action.fn(); };
      el.appendChild(b);
    }
    el.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(hideToast, action ? 7000 : 4000);
  }
  function hideToast() { $('#toast').hidden = true; }

  function showBanner(html, kind = '') {
    const el = $('#banner');
    el.className = 'banner' + (kind ? ' ' + kind : '');
    el.innerHTML = html;
    el.hidden = false;
  }
  function hideBanner() { $('#banner').hidden = true; }

  /* =====================================================================
   *  Enregistrement automatique et rechargement automatique
   * ===================================================================== */
  function markDirty() {
    state.dirty = true;
    state.editGen++;
    state.lastEditAt = Date.now();
    cache();
    if (!cfg) { setStatus('disconnected'); return; }
    if (state.conflict) return;
    setStatus('pending');
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => flush(), SAVE_DELAY);
  }

  async function flush(force = false) {
    clearTimeout(saveTimer);
    if (!cfg || !state.dirty || state.conflict) return;
    if (state.saving) { state.resave = true; return; }
    state.saving = true;
    setStatus('saving');
    const gen = state.editGen;
    let again = 0;
    try {
      const r = await api('save', { data: serialize(), baseVersion: state.version, force });
      if (r.ok) {
        state.version = r.version;
        state.lastSync = new Date();
        if (!state.conflict) hideBanner();
        if (gen === state.editGen) { state.dirty = false; cache(); setStatus('ok'); }
        else again = 300;
      } else if (r.conflict) {
        state.conflict = r;
        setStatus('conflict');
        showBanner(`<p><strong>Le Google Sheet a changé ailleurs</strong> pendant que tu modifiais ici. Choisis quelle version garder.</p>
          <button class="btn" type="button" data-act="take-remote">Recharger la version du Sheet</button>
          <button class="btn primary" type="button" data-act="keep-mine">Garder mes modifications</button>`, '');
      } else {
        handleApiError(r);
      }
    } catch (e) {
      if (isNetworkError(e)) { setStatus('offline'); again = 15000; }
      else { setStatus('error'); showBanner('<p><strong>Enregistrement impossible.</strong> Vérifie l’adresse du service dans Réglages, onglet Connexion.</p>', 'err'); }
    } finally {
      state.saving = false;
      if (state.resave) { state.resave = false; again = again || 300; }
      if (again && state.dirty && !state.conflict) saveTimer = setTimeout(() => flush(), again);
    }
  }

  async function pull(manual = false) {
    if (!cfg || state.pulling) return;
    state.pulling = true;
    try {
      const r = await api('load');
      if (!r.ok) { handleApiError(r); return; }
      state.lastSync = new Date();
      state.loaded = true;
      if (r.version === state.version) {
        if (!state.dirty && !state.saving) setStatus('ok');
        if (manual) toast('Déjà à jour.');
        if (!$('#rows').children.length) renderAll();
        return;
      }
      if (state.dirty || state.saving || isEditing() || dlg.open) return;   // on réessaiera au prochain tour
      hideBanner();
      hydrate(r.data);
      state.version = r.version;
      cache();
      renderAll();
      setStatus('ok');
      toast(manual ? 'Données rechargées depuis le Google Sheet.' : 'Mis à jour depuis le Google Sheet.');
    } catch (e) {
      if (isNetworkError(e)) setStatus('offline');
      else { setStatus('error'); if (!state.loaded) showBanner('<p><strong>Le service Google ne répond pas comme prévu.</strong> Vérifie l’adresse dans Réglages, onglet Connexion, et que le déploiement est ouvert à « Tout le monde ».</p>', 'err'); }
    } finally { state.pulling = false; }
  }

  function handleApiError(r) {
    if (r.error === 'auth') {
      state.authFailed = true;
      setStatus('error', 'Mot de passe refusé');
      openSettings('connexion');
      setCfgMsg('Mot de passe refusé par le service Google. Vérifie qu’il est identique à celui de Code.gs.', false);
    } else {
      setStatus('error');
      showBanner(`<p><strong>Erreur du service Google :</strong> ${esc(r.error || 'inconnue')}</p>`, 'err');
    }
  }

  // Vrai seulement si la personne est en train de saisir (champ actif ET modifié il y a moins de 8 s),
  // pour ne pas redessiner le tableau sous ses doigts.
  function isEditing() {
    const a = document.activeElement;
    return !!a && ['INPUT', 'SELECT', 'TEXTAREA'].includes(a.tagName) && !!a.closest('#rows') && Date.now() - state.lastEditAt < 8000;
  }

  function tick() {
    if (document.hidden || !cfg) return;
    if (state.saving || state.pulling || state.conflict || state.authFailed) return;
    if (state.dirty) flush(); else if (dlg.open) return; else pull();
  }

  $('#banner').addEventListener('click', (e) => {
    const act = e.target.dataset.act;
    if (!act || !state.conflict) return;
    const c = state.conflict;
    state.conflict = null;
    hideBanner();
    if (act === 'take-remote') {
      hydrate(c.data); state.version = c.version; state.dirty = false; cache(); renderAll(); setStatus('ok');
    } else {
      state.version = c.version; flush(true);
    }
  });

  /* =====================================================================
   *  Filtres et affichage du tableau
   * ===================================================================== */
  function visibleRows() {
    const f = state.filter;
    const q = f.q.trim().toLowerCase();
    return state.rows.filter((r) => {
      if (f.mode === 'unpaid' && r.paid) return false;
      if (f.mode === 'uncounted' && r.counted) return false;
      if (f.arena && r.arena !== f.arena) return false;
      if (f.month && !(r.date || '').startsWith(f.month)) return false;
      if (q && !`${r.comment} ${r.category} ${r.arena}`.toLowerCase().includes(q)) return false;
      return true;
    });
  }

  function optionsHTML(list, current, placeholder) {
    const seen = new Set();
    let html = `<option value=""${current ? '' : ' selected'}>${esc(placeholder)}</option>`;
    list.forEach((v) => {
      if (!v || seen.has(v)) return;
      seen.add(v);
      html += `<option value="${esc(v)}"${v === current ? ' selected' : ''}>${esc(v)}</option>`;
    });
    if (current && !seen.has(current)) html += `<option value="${esc(current)}" selected>⚠ ${esc(current)} (introuvable)</option>`;
    return html;
  }

  function rowHTML(r, arenaNames, catNames) {
    const c = calc(r);
    const cls = [r.paid ? 'paid' : '', c.missingCategory || c.missingArena ? 'warn' : ''].filter(Boolean).join(' ');
    const chk = (f, label) => `<input type="checkbox" data-f="${f}" aria-label="${label}"${r[f] ? ' checked' : ''}>`;
    return `<tr data-id="${r.id}" class="${cls}">
      <td class="c-date" data-label="Date"><input type="date" data-f="date" value="${esc(r.date)}" aria-label="Date"></td>
      <td class="c-arena" data-label="Aréna"><select data-f="arena" aria-label="Aréna">${optionsHTML(arenaNames, r.arena, 'Choisir une aréna')}</select></td>
      <td class="c-cat" data-label="Catégorie"><select data-f="category" aria-label="Catégorie">${optionsHTML(catNames, r.category, 'Choisir une catégorie')}</select></td>
      <td class="c-price n calc" data-label="Prix" data-c="price">${fmt(c.price)}</td>
      <td class="c-tpay c" data-label="Déplacement à payer">${chk('travel', 'Déplacement à payer')}</td>
      <td class="c-travel n calc" data-label="Déplacement" data-c="travel">${fmt(c.travel)}</td>
      <td class="c-extra n" data-label="Supplément"><input type="number" step="0.01" data-f="extra" value="${r.extra || ''}" placeholder="0" aria-label="Supplément" inputmode="decimal"></td>
      <td class="c-total n calc total" data-label="Total" data-c="total">${fmt(c.total)}</td>
      <td class="c-count c" data-label="Comptabilisé">${chk('counted', 'Comptabilisé')}</td>
      <td class="c-comment" data-label="Commentaire"><input type="text" data-f="comment" value="${esc(r.comment)}" aria-label="Commentaire"></td>
      <td class="c-tour c" data-label="Tournoi">${chk('tournament', 'Tournoi')}</td>
      <td class="c-paid c" data-label="Payé">${chk('paid', 'Payé')}</td>
      <td class="c-del c"><button class="icon-btn" type="button" data-act="del" aria-label="Supprimer cette partie">${TRASH}</button></td>
    </tr>`;
  }

  function renderTable() {
    const arenaNames = state.arenas.map((a) => a.name);
    const catNames = state.tariffs.map(descOf);
    $('#rows').innerHTML = visibleRows().map((r) => rowHTML(r, arenaNames, catNames)).join('');
    renderEmpty();
  }

  function renderEmpty() {
    const el = $('#empty');
    const anyData = state.rows.length || state.tariffs.length || state.arenas.length;
    if (!state.loaded && cfg && !state.rows.length) {
      el.innerHTML = '<strong>Chargement…</strong>Lecture du Google Sheet en cours.';
      el.hidden = false;
    } else if (!cfg && !anyData) {
      el.innerHTML = '<strong>Pas encore connecté</strong>Ouvre les Réglages pour relier ce site à ton Google Sheet.';
      el.hidden = false;
    } else if (!anyData) {
      el.innerHTML = `<strong>Le Google Sheet est vide</strong>Importe tes données de départ, ou commence avec une liste de tarifs et d’arénas vide.
        <div class="row-actions"><button class="btn primary" type="button" data-act="import">Importer un fichier .json</button><button class="btn" type="button" data-act="settings">Ouvrir les réglages</button></div>`;
      el.hidden = false;
    } else if (!visibleRows().length) {
      el.innerHTML = state.rows.length
        ? '<strong>Aucune partie ne correspond</strong>Change ou retire les filtres pour les voir.'
        : '<strong>Aucune partie pour l’instant</strong>Ajoute la première avec le bouton orange.';
      el.hidden = false;
    } else {
      el.hidden = true;
    }
  }

  function renderFilters() {
    const f = state.filter;
    const arenas = [...new Set([...state.arenas.map((a) => a.name), ...state.rows.map((r) => r.arena)].filter(Boolean))];
    $('#f-arena').innerHTML = optionsHTML(arenas, f.arena, 'Toutes les arénas');
    const months = [...new Set(state.rows.map((r) => (r.date || '').slice(0, 7)).filter((m) => /^\d{4}-\d{2}$/.test(m)))].sort();
    let html = `<option value=""${f.month ? '' : ' selected'}>Tous les mois</option>`;
    months.forEach((m) => {
      const [y, mo] = m.split('-').map(Number);
      const label = new Date(y, mo - 1, 1).toLocaleDateString('fr-CA', { month: 'long', year: 'numeric' });
      html += `<option value="${m}"${m === f.month ? ' selected' : ''}>${esc(label)}</option>`;
    });
    $('#f-month').innerHTML = html;
    $$('.seg button').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.mode === f.mode)));
  }

  function refreshBoard() {
    let price = 0, travel = 0, extra = 0, due = 0;
    const vis = visibleRows();
    vis.forEach((r) => {
      const c = calc(r);
      price += c.price; travel += c.travel; extra += c.extra;
      if (!r.paid) due += c.total;
    });
    $('#b-price').textContent = fmt(round2(price));
    $('#b-travel').textContent = fmt(round2(travel));
    $('#b-extra').textContent = fmt(round2(extra));
    $('#b-total').textContent = fmt(round2(price + travel + extra));
    $('#b-due').textContent = fmt(round2(due));
    $('#b-count').textContent = String(vis.length);
  }

  function refreshRow(tr, r) {
    const c = calc(r);
    $('[data-c="price"]', tr).textContent = fmt(c.price);
    $('[data-c="travel"]', tr).textContent = fmt(c.travel);
    $('[data-c="total"]', tr).textContent = fmt(c.total);
    tr.classList.toggle('paid', !!r.paid);
    tr.classList.toggle('warn', c.missingCategory || c.missingArena);
  }

  function renderAll() {
    buildLookups();
    renderFilters();
    renderTable();
    refreshBoard();
  }

  /* =====================================================================
   *  Édition du tableau
   * ===================================================================== */
  // Les champs texte/nombre réagissent à « input », les autres à « change ».
  function onRowEdit(e) {
    const el = e.target;
    const f = el.dataset.f;
    if (!f) return;
    if ((e.type === 'input') !== (el.type === 'text' || el.type === 'number')) return;
    const tr = el.closest('tr');
    const r = state.rows.find((x) => x.id === tr.dataset.id);
    if (!r) return;
    r[f] = el.type === 'checkbox' ? el.checked : el.type === 'number' ? (parseFloat(el.value) || 0) : el.value;
    refreshRow(tr, r);
    refreshBoard();
    markDirty();
  }

  $('#rows').addEventListener('input', onRowEdit);
  $('#rows').addEventListener('change', onRowEdit);

  $('#rows').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-act="del"]');
    if (!btn) return;
    const id = btn.closest('tr').dataset.id;
    const idx = state.rows.findIndex((x) => x.id === id);
    if (idx < 0) return;
    const [removed] = state.rows.splice(idx, 1);
    renderAll();
    markDirty();
    toast('Partie supprimée.', {
      label: 'Annuler',
      fn: () => { state.rows.splice(Math.min(idx, state.rows.length), 0, removed); renderAll(); markDirty(); }
    });
  });

  function defaultTravel(date, arena) {
    if (!(arenaMap.get(arena) > 0)) return false;
    return !state.rows.some((r) => r.date === date && r.arena === arena && r.travel);
  }

  function addRow() {
    const last = state.rows[state.rows.length - 1];
    const date = todayISO();
    const arena = (last && last.arena) || (state.arenas[0] && state.arenas[0].name) || '';
    const category = (last && last.category) || '';
    const r = { id: uid(), date, arena, category, travel: defaultTravel(date, arena), extra: 0, counted: false, comment: '', tournament: false, paid: false };
    state.rows.push(r);
    state.filter = { mode: 'all', arena: '', month: '', q: '' };
    $('#f-q').value = '';
    renderAll();
    markDirty();
    const tr = $(`#rows tr[data-id="${r.id}"]`);
    if (tr) {
      tr.classList.add('fresh');
      tr.scrollIntoView({ block: 'center', behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' });
    }
  }

  function sortByDate() {
    state.rows.sort((a, b) => (a.date || '9999').localeCompare(b.date || '9999'));
    renderAll();
    markDirty();
    toast('Parties triées par date.');
  }

  $('#add').addEventListener('click', addRow);
  $('#sort').addEventListener('click', sortByDate);

  $('.seg').addEventListener('click', (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    state.filter.mode = b.dataset.mode;
    renderFilters(); renderTable(); refreshBoard();
  });
  $('#f-arena').addEventListener('change', (e) => { state.filter.arena = e.target.value; renderTable(); refreshBoard(); });
  $('#f-month').addEventListener('change', (e) => { state.filter.month = e.target.value; renderTable(); refreshBoard(); });
  $('#f-q').addEventListener('input', (e) => { state.filter.q = e.target.value; renderTable(); refreshBoard(); });

  $('#empty').addEventListener('click', (e) => {
    const act = e.target.dataset.act;
    if (act === 'import') { openSettings('donnees'); $('#d-file').click(); }
    if (act === 'settings') openSettings('tarifs');
  });

  /* =====================================================================
   *  Réglages : tarifs, arénas, connexion, données
   * ===================================================================== */
  const dlg = $('#settings');

  function openSettings(tab) {
    renderTariffs();
    renderArenas();
    $('#cfg-url').value = cfg ? cfg.url : '';
    $('#cfg-token').value = cfg ? cfg.token : '';
    selectTab(tab || 'tarifs');
    if (!dlg.open) dlg.showModal();
  }

  function selectTab(name) {
    $$('.tabs button').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === name)));
    $$('.panel').forEach((p) => p.classList.toggle('on', p.id === `tab-${name}`));
  }

  $('#open-settings').addEventListener('click', () => openSettings('tarifs'));
  $('#close-settings').addEventListener('click', () => dlg.close());
  dlg.addEventListener('close', () => { renderAll(); });
  $('.tabs').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) selectTab(b.dataset.tab); });

  /* ---- Tarifs ---- */
  function renderTariffs() {
    $('#t-body').innerHTML = state.tariffs.map((t) => `<tr data-id="${t.id}">
      <td class="t-cat" data-label="Catégorie"><input type="text" list="dl-cats" data-f="cat" value="${esc(t.cat)}" aria-label="Catégorie"></td>
      <td class="t-dbl c" data-label="Double lettre"><input type="checkbox" data-f="double" ${t.double ? 'checked' : ''} aria-label="Double lettre"></td>
      <td class="t-type" data-label="Type"><input type="text" list="dl-types" data-f="type" value="${esc(t.type)}" aria-label="Type"></td>
      <td class="t-cost" data-label="Coût ($)"><input type="number" step="0.01" min="0" data-f="cost" value="${t.cost}" aria-label="Coût" inputmode="decimal"></td>
      <td class="t-desc desc" data-c="desc">${esc(descOf(t))}</td>
      <td class="t-del"><button class="icon-btn" type="button" data-act="del" aria-label="Supprimer ce tarif">${TRASH}</button></td>
    </tr>`).join('');
    $('#dl-cats').innerHTML = [...new Set(state.tariffs.map((t) => t.cat).filter(Boolean))].map((v) => `<option value="${esc(v)}">`).join('');
    $('#dl-types').innerHTML = [...new Set(state.tariffs.map((t) => t.type).filter(Boolean))].map((v) => `<option value="${esc(v)}">`).join('');
  }

  function onTariffEdit(e) {
    const el = e.target;
    const f = el.dataset.f;
    if (!f) return;
    if ((e.type === 'input') !== (el.type === 'text' || el.type === 'number')) return;
    const tr = el.closest('tr');
    const t = state.tariffs.find((x) => x.id === tr.dataset.id);
    if (!t) return;
    const before = descOf(t);
    const unique = state.tariffs.filter((x) => descOf(x) === before).length === 1;
    t[f] = el.type === 'checkbox' ? el.checked : el.type === 'number' ? (parseFloat(el.value) || 0) : el.value;
    const after = descOf(t);
    if (before !== after && unique) state.rows.forEach((r) => { if (r.category === before) r.category = after; });
    $('[data-c="desc"]', tr).textContent = after;
    buildLookups();
    markDirty();
  }
  $('#t-body').addEventListener('input', onTariffEdit);
  $('#t-body').addEventListener('change', onTariffEdit);

  $('#t-body').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-act="del"]');
    if (!btn) return;
    const id = btn.closest('tr').dataset.id;
    const t = state.tariffs.find((x) => x.id === id);
    const used = state.rows.filter((r) => r.category === descOf(t)).length;
    if (used && !confirm(`${used} partie(s) utilisent ce tarif et se retrouveront sans prix. Supprimer quand même ?`)) return;
    state.tariffs = state.tariffs.filter((x) => x.id !== id);
    buildLookups(); renderTariffs(); markDirty();
  });

  $('#t-add').addEventListener('click', () => {
    state.tariffs.push({ id: uid(), cat: '', double: false, type: '', cost: 0 });
    renderTariffs(); markDirty();
    const inputs = $$('#t-body tr:last-child input');
    if (inputs[0]) inputs[0].focus();
  });

  /* ---- Arénas ---- */
  function renderArenas() {
    $('#a-body').innerHTML = state.arenas.map((a) => `<tr data-id="${a.id}">
      <td class="a-name" data-label="Aréna"><input type="text" data-f="name" value="${esc(a.name)}" aria-label="Nom de l’aréna"></td>
      <td class="a-cost" data-label="Coût de déplacement ($)"><input type="number" step="0.01" min="0" data-f="travelCost" value="${a.travelCost}" aria-label="Coût de déplacement" inputmode="decimal"></td>
      <td class="a-del"><button class="icon-btn" type="button" data-act="del" aria-label="Supprimer cette aréna">${TRASH}</button></td>
    </tr>`).join('');
  }

  function onArenaEdit(e) {
    const el = e.target;
    const f = el.dataset.f;
    if (!f || e.type !== 'input') return;
    const a = state.arenas.find((x) => x.id === el.closest('tr').dataset.id);
    if (!a) return;
    if (f === 'name') {
      const before = a.name;
      const unique = state.arenas.filter((x) => x.name === before).length === 1;
      a.name = el.value;
      if (unique && before !== a.name) state.rows.forEach((r) => { if (r.arena === before) r.arena = a.name; });
    } else {
      a.travelCost = parseFloat(el.value) || 0;
    }
    buildLookups();
    markDirty();
  }
  $('#a-body').addEventListener('input', onArenaEdit);

  $('#a-body').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-act="del"]');
    if (!btn) return;
    const id = btn.closest('tr').dataset.id;
    const a = state.arenas.find((x) => x.id === id);
    const used = state.rows.filter((r) => r.arena === a.name).length;
    if (used && !confirm(`${used} partie(s) sont dans cette aréna. Supprimer quand même ?`)) return;
    state.arenas = state.arenas.filter((x) => x.id !== id);
    buildLookups(); renderArenas(); markDirty();
  });

  $('#a-add').addEventListener('click', () => {
    state.arenas.push({ id: uid(), name: '', travelCost: 0 });
    renderArenas(); markDirty();
    const inputs = $$('#a-body tr:last-child input');
    if (inputs[0]) inputs[0].focus();
  });

  /* ---- Connexion ---- */
  function setCfgMsg(text, ok) {
    const el = $('#cfg-msg');
    el.textContent = text;
    el.className = 'msg' + (text ? (ok ? ' ok' : ' bad') : '');
  }

  $('#cfg-save').addEventListener('click', async () => {
    const url = $('#cfg-url').value.trim();
    const token = $('#cfg-token').value;
    if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) {
      setCfgMsg('L’adresse doit commencer par https://script.google.com/ et se terminer par /exec.', false);
      return;
    }
    if (!token) { setCfgMsg('Entre le mot de passe.', false); return; }
    setCfgMsg('Test de la connexion…', true);
    try {
      const conf = { url, token };
      const r = await api('load', {}, conf);
      if (!r.ok) {
        setCfgMsg(r.error === 'auth' ? 'Mot de passe refusé. Il doit être identique à celui de Code.gs.' : `Erreur : ${r.error}`, false);
        return;
      }
      const changed = !cfg || cfg.url !== url;
      const remoteEmpty = !r.data.rows.length && !r.data.tariffs.length && !r.data.arenas.length;
      writeCfg(conf);
      cfg = conf;
      state.authFailed = false;
      if (state.dirty && remoteEmpty) {
        state.version = r.version;   // on garde les données locales et on les envoie
      } else if (changed || !state.dirty) {
        hydrate(r.data); state.version = r.version; state.dirty = false; state.conflict = null; hideBanner();
        cache(); renderAll(); renderTariffs(); renderArenas();
      }
      state.loaded = true;
      state.lastSync = new Date();
      setStatus('ok');
      const n = r.data.rows.length;
      setCfgMsg(n ? `Connecté. ${n} partie(s) chargée(s) depuis le Google Sheet.` : 'Connecté. Le Google Sheet est vide : importe tes données dans l’onglet Données.', true);
      if (state.dirty) flush();
    } catch (e) {
      setCfgMsg(isNetworkError(e)
        ? 'Impossible de joindre Google. Vérifie ta connexion Internet.'
        : 'Réponse inattendue. Vérifie l’adresse et que le déploiement est ouvert à « Tout le monde ».', false);
    }
  });

  $('#cfg-pair').addEventListener('click', async () => {
    if (!cfg) { setCfgMsg('Enregistre d’abord la connexion sur cet appareil.', false); return; }
    const link = `${location.origin}${location.pathname}#cfg=${b64enc(JSON.stringify(cfg))}`;
    try { await navigator.clipboard.writeText(link); setCfgMsg('Lien copié. Ouvre-le sur l’autre appareil.', true); }
    catch (_) { prompt('Copie ce lien et ouvre-le sur l’autre appareil :', link); }
  });

  $('#cfg-forget').addEventListener('click', () => {
    if (!confirm('Déconnecter cet appareil ? Les données restent dans le Google Sheet.')) return;
    localStorage.removeItem(LS_CFG);
    localStorage.removeItem(LS_CACHE);
    state.dirty = false;
    location.reload();
  });

  /* ---- Données ---- */
  $('#d-reload').addEventListener('click', async () => {
    if (state.dirty && !confirm('Des modifications ne sont pas encore enregistrées. Les remplacer par la version du Google Sheet ?')) return;
    state.dirty = false; state.version = null;
    await pull(true);
    renderTariffs(); renderArenas();
  });

  $('#d-export').addEventListener('click', () => {
    const blob = new Blob([JSON.stringify(serialize(), null, 1)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `arbitrage-${todayISO()}.json`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  });

  $('#d-import').addEventListener('click', () => $('#d-file').click());
  $('#d-file').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    try {
      const d = JSON.parse(await file.text());
      if (!Array.isArray(d.rows) || !Array.isArray(d.tariffs) || !Array.isArray(d.arenas)) throw new Error('format');
      const has = state.rows.length || state.tariffs.length || state.arenas.length;
      if (has && !confirm(`Remplacer les données actuelles par ${d.rows.length} partie(s), ${d.tariffs.length} tarif(s) et ${d.arenas.length} aréna(s) ?`)) return;
      hydrate(d);
      renderAll(); renderTariffs(); renderArenas();
      markDirty();
      toast(`${d.rows.length} partie(s) importée(s).`);
    } catch (_) {
      toast('Fichier invalide : utilise un .json exporté depuis ce site.');
    }
  });

  /* =====================================================================
   *  Démarrage
   * ===================================================================== */
  async function start() {
    consumePairingHash();
    cfg = readCfg();

    const cached = readCache();
    if (cached && cached.data) {
      hydrate(cached.data);
      state.version = cached.version || null;
      state.dirty = !!cached.dirty;
      state.loaded = true;
    }
    renderAll();

    setInterval(tick, POLL_MS);
    document.addEventListener('visibilitychange', () => { if (!document.hidden) tick(); });
    window.addEventListener('online', tick);
    window.addEventListener('offline', () => setStatus('offline'));

    if (!cfg) {
      setStatus('disconnected');
      openSettings('connexion');
      return;
    }
    if (state.dirty) { setStatus('pending'); await flush(); }
    else { setStatus('loading'); await pull(); }
  }

  window.addEventListener('beforeunload', (e) => {
    if (state.dirty) { e.preventDefault(); e.returnValue = ''; }
  });

  start();
})();
