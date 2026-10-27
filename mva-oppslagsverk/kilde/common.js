/* Felles kjerne: data, gjengivelse, utvalg og PDF-eksport. Ingen eksterne avhengigheter. */
const D = window.DATA;
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const seg = p => p.map(x => `<span class="f-${x.farge}">${x.fet ? `<strong>${esc(x.tekst)}</strong>` : esc(x.tekst)}</span>`).join('');
const paras = l => l.map(p => `<p>${seg(p)}</p>`).join('');
const plain = p => p.map(x => x.tekst).join('');

const ITEMS = [
  ...D.scenarioer.map(s => ({ id: 's' + s.nr, type: 'scenario', s })),
  { id: 'yrk', type: 'yrk' }
];
const ITEM = Object.fromEntries(ITEMS.map(i => [i.id, i]));
const itemTittel = i => i.type === 'scenario' ? `Scenario ${i.s.nr}` : 'Yrkesmessig utleie';

/* Tekst som kan søkes i */
function sokTekst(i) {
  if (i.type === 'scenario') {
    const s = i.s;
    return [`scenario ${s.nr}`, s.gruppe, plain(s.problemstilling), s.problemstilling_alt ? plain(s.problemstilling_alt) : '',
      s.problemstilling_tillegg ? plain(s.problemstilling_tillegg) : '', ...s.konklusjon.map(plain),
      s.fotnote ? plain(s.fotnote) : ''].join(' ').toLowerCase();
  }
  if (i.type === 'yrk') return (plain(D.yrkesmessig_utleie.tittel) + ' ' + D.yrkesmessig_utleie.avsnitt.map(plain).join(' ')).toLowerCase();
}
/* Merkelapper til filtrering, utledet av ordlyden i problemstillingen */
function tagsFor(i) {
  if (i.type !== 'scenario') return [];
  const s = i.s, t = (plain(s.problemstilling) + ' ' + (s.problemstilling_alt ? plain(s.problemstilling_alt) : '') + ' ' + (s.problemstilling_tillegg ? plain(s.problemstilling_tillegg) : '')).toLowerCase();
  const r = [s.gruppe === 'Utleie' ? 'Utleie' : 'Andre forhold'];
  if (/fossilt/.test(t)) r.push('Fossilt');
  if (/elektrisk|elbil/.test(t)) r.push('Elektrisk');
  if (/varebil/.test(t)) r.push('Varebil klasse 2');
  if (/demobil/.test(t)) r.push('Demobil');
  if (/firmabil/.test(t)) r.push('Firmabil');
  if (/under 30/.test(t)) r.push('Under 30 dager');
  if (/over 30/.test(t)) r.push('Over 30 dager');
  return r;
}
const ALLE_TAGS = ['Utleie', 'Andre forhold', 'Fossilt', 'Elektrisk', 'Varebil klasse 2', 'Demobil', 'Firmabil', 'Under 30 dager', 'Over 30 dager'];

/* Gjengivelse av ett scenario (konklusjonstekst) */
function titler(s) {
  let h = `<div class="titler"><p>${seg(s.problemstilling)}</p>`;
  if (s.problemstilling_alt) h += `<p>${seg(s.problemstilling_alt)}</p>`;
  if (s.problemstilling_tillegg) h += `<p>${seg(s.problemstilling_tillegg)}</p>`;
  return h + '</div>';
}
function konklusjon(s) {
  let h = `<div class="konk">${paras(s.konklusjon)}</div>`;
  if (s.fotnote) h += `<p class="fotnote">${seg(s.fotnote)}</p>`;
  return h;
}
function yrkHtml() {
  return `<div class="yrk-avsnitt">${D.yrkesmessig_utleie.avsnitt.map(a => `<p>${seg(a)}</p>`).join('')}</div>`;
}
const forutsetningHtml = () => `<p class="forutsetning">${seg(D.tekst_over_tabell.forutsetning)}</p>`;
const opprinnelig = s => s.opprinnelig_nr.length > 1 ? `PP-scenario ${s.opprinnelig_nr.join(' og ')}` : `PP-scenario ${s.opprinnelig_nr[0]}`;

/* Utvalg */
const valgt = new Set(ITEMS.map(i => i.id));
const velgBoks = (id, tekst = 'Ta med i PDF') =>
  `<label class="velg"><input type="checkbox" data-velg="${id}" ${valgt.has(id) ? 'checked' : ''}><span>${tekst}</span></label>`;
const kunKnapp = id => `<button type="button" class="kun" data-kun="${id}">Kun denne til PDF</button>`;
function synk() {
  $$('[data-velg]').forEach(c => c.checked = valgt.has(c.dataset.velg));
  $$('.antall').forEach(a => a.textContent = `${valgt.size} av ${ITEMS.length} valgt`);
  window.etterSynk && window.etterSynk();
}
function toast(t) {
  const e = $('#toast'); e.textContent = t; e.className = 'vis';
  clearTimeout(toast.t); toast.t = setTimeout(() => e.className = '', 3200);
}

/* Utskriftsdokument */
const dato = () => new Date().toLocaleDateString('nb-NO', { day: 'numeric', month: 'long', year: 'numeric' });
const lockHtml = () => `<div class="lock"><div class="bauda">Bauda<i></i></div>${$(".bdo") ? `<span class="strek-u"></span><div class="bdo-u">${$(".bdo").innerHTML}</div>` : ""}</div>`;
function artikkel(i) {
  if (i.type === 'scenario') {
    const s = i.s;
    return `<article class="u-side"><p class="gruppe">${esc(s.gruppe)}</p><h2>Scenario ${s.nr}</h2>
        <h3>Problemstilling</h3>${titler(s)}<h3>Konklusjon/mva-behandling</h3><div class="kon">${konklusjon(s)}</div></article>`;
  }
  return `<article class="u-side"><p class="gruppe">Bakgrunn</p><h2>${seg(D.yrkesmessig_utleie.tittel)}</h2>${yrkHtml()}</article>`;
}
const innhold = i => `<li><b>${itemTittel(i)}</b><span>${i.type === 'scenario' ? esc(plain(i.s.problemstilling)) : 'Regler for yrkesmessig utleie'}</span></li>`;
function byggUtskrift(ids, modus) {
  const valg = ITEMS.filter(i => ids.includes(i.id));
  let h = '';
  if (modus === 'alt') {
    /* Komplett utgave: forside og innholdsfortegnelse */
    h += `<section class="u-forside">${lockHtml()}
    <h1>${esc(D.tittel)}</h1><p class="ut">Oppslagsverk for ${esc(D.undertittel.toLowerCase().replace('bauda-', 'Bauda-'))}</p>
    ${$("#bil") ? `<img class="bil-u" src="${$("#bil").src}" alt="">` : ""}<div class="ma">Komplett utgave · ${dato()}</div></section>`;
    h += `<section class="u-toc"><h2>Innhold</h2><ol>${valg.map(innhold).join('')}</ol></section>`;
  } else {
    /* Utvalg: kompakt topp på første side, uten forside og uten egen innholdsside */
    h += `<section class="u-topp">${lockHtml()}<h1>${esc(D.tittel)}</h1><p class="ut">Utvalg · ${dato()}</p>${valg.length > 1 ? `<ol class="u-liste">${valg.map(innhold).join('')}</ol>` : ''}</section>`;
  }
  h += valg.map(artikkel).join('');
  $('#utskrift').innerHTML = h;
}
function eksporter(ids, modus) {
  if (!ids.length) { toast('Ingenting er valgt. Kryss av minst ett element, eller bruk «Eksporter alt».'); return; }
  byggUtskrift(ids, modus);
  setTimeout(() => window.print(), 60);
}
window.__bygg = (ids, modus) => byggUtskrift(ids, modus);   // brukes til testing uten utskriftsdialog

document.addEventListener('change', e => {
  const t = e.target;
  if (t.dataset.velg) { t.checked ? valgt.add(t.dataset.velg) : valgt.delete(t.dataset.velg); synk(); }
});
document.addEventListener('click', e => {
  const k = e.target.closest('[data-kun]');
  if (k) { eksporter([k.dataset.kun], 'valgt'); return; }
  const b = e.target.closest('[data-h]');
  if (!b) return;
  const h = b.dataset.h;
  if (h === 'alle') { ITEMS.forEach(i => valgt.add(i.id)); synk(); }
  if (h === 'ingen') { valgt.clear(); synk(); }
  if (h === 'synlige') {
    const s = (window.synlige ? window.synlige() : ITEMS.map(i => i.id));
    valgt.clear(); s.forEach(id => valgt.add(id)); synk();
    toast(`${s.length} synlige valgt`);
  }
  if (h === 'valgte') eksporter(ITEMS.filter(i => valgt.has(i.id)).map(i => i.id), 'valgt');
  if (h === 'eksalt') eksporter(ITEMS.map(i => i.id), 'alt');
});

/* Filter: søk + merkelapper. Returnerer objekt med visible(item). */
function lagFilter(mount, onChange, tags = ALLE_TAGS) {
  const st = { q: '', aktive: new Set() };
  mount.innerHTML = `<label class="sok"><span class="sr">Søk</span><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
    <input type="search" placeholder="Søk i teksten" aria-label="Søk i teksten"></label>
    <div class="chips" role="group" aria-label="Filter">${tags.map(t => `<button class="chip" aria-pressed="false" data-tag="${t}">${t}</button>`).join('')}
    <button class="chip" data-nullstill hidden>Nullstill</button></div>`;
  const nul = $('[data-nullstill]', mount);
  const opp = () => { nul.hidden = !st.q && !st.aktive.size; onChange(); };
  $('input', mount).addEventListener('input', e => { st.q = e.target.value.trim().toLowerCase(); opp(); });
  mount.addEventListener('click', e => {
    const c = e.target.closest('[data-tag]');
    if (c) { const t = c.dataset.tag; st.aktive.has(t) ? st.aktive.delete(t) : st.aktive.add(t); c.setAttribute('aria-pressed', st.aktive.has(t)); opp(); }
    if (e.target.closest('[data-nullstill]')) {
      st.q = ''; st.aktive.clear(); $('input', mount).value = '';
      $$('[data-tag]', mount).forEach(x => x.setAttribute('aria-pressed', 'false')); opp();
    }
  });
  st.visible = i => {
    if (st.q && !sokTekst(i).includes(st.q)) return false;
    if (st.aktive.size) { if (i.type !== 'scenario') return false; const tg = tagsFor(i); for (const a of st.aktive) if (!tg.includes(a)) return false; }
    return true;
  };
  return st;
}

/* Felles oppstart */
function start() {
  const stil = document.createElement('style');
  stil.textContent = `@media print{@page{@top-right{content:"${dato()}";font:9pt "Segoe UI",Arial,sans-serif;color:#6B6560}}}`;
  document.head.appendChild(stil);
  document.body.insertAdjacentHTML('beforeend', '<div id="utskrift"></div><div id="toast" role="status"></div>');
  synk();
  const p = new URLSearchParams(location.search).get('pdf');   // ?pdf=alt eller ?pdf=s1,s3,yrk (kun for testing)
  if (p) byggUtskrift(p === 'alt' ? ITEMS.map(i => i.id) : p.split(','), p === 'alt' ? 'alt' : 'valgt');
}
window.addEventListener('DOMContentLoaded', () => { window.init && window.init(); start(); });
