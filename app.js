/* Mušičarska liga Crne Gore. Rang liste iz Excel fajlova saveza.
   Logika je prevedena iz popuni_tabelu.py: isti aliasi klubova, isti ključ za
   uparivanje imena, isti službeni izvor plasmana (sheet POJEDINACNO) i ista
   pravila kazne za preskočeno kolo. */

/* ---------- podešavanja ---------- */

// Fajlovi kola se traže kao data/kolo-1.xlsx, data/kolo-2.xlsx ... do MAX_KOLA.
// Novo kolo = novi fajl u data/ folderu, ništa drugo se ne dira.
const MAX_KOLA = 12;
const KAZNA_PRESKOK = 30;          // plasman-poena po preskočenom kolu
const REPO_DATA_URL = 'https://github.com/alt-DooM/mlcg/tree/main/data';

const KLUB_ALIASI = {
  'SRK Gorštak - Kolašin':         ['SRK GORŠTAK - KOLAŠIN', 'SRK GORŠTAK KL'],
  'SRFFK Ravnjak - Mojkovac':      ['SRFFK RAVNJAK', 'SRFFK RAVNJAK - MOJKOVAC'],
  'SRK Kolašin':                   ['SRK KOLAŠIN'],
  'SRK Vodene lisice - Podgorica': ['SRK VODENE LISICE', 'SRK VODENE LISICE - PG'],
  'SRK Lipljen - Pljevlja':        ['SRK LIPLJEN', 'SRK LIPLJEN - PLJEVLJA'],
  'SRK Lim - Berane':              ['SRK LIM - BERANE', 'SRK LIM BA'],
  'SRK Lim - Andrijevica':         ['SRK LIM AN', 'SRK LIM - ANDRIJEVICA'],
  'SRFFK Maniro - Kolašin':        ['SRFFK MANIRO', 'SRFFK MANIRO - KOLAŠIN'],
  'SRK Plavsko jezero':            ['SRK PLAVSKO JEZERO'],
  'SRK Tara - Mojkovac':           ['SRK TARA - MOJKOVAC', 'SRK TARA MK'],
  'SRK EPCG':                      ['SRK EPCG'],
};

const RIMSKI = ['', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];
const BOJE = ['#0f4c53', '#c8712a', '#2f6b45', '#7a5aa8', '#a83e4c', '#3b6b8c'];
const MAX_SERIJA = 5;              // koliko takmičara može istovremeno na grafikon

/* ---------- tekst ---------- */

const normalizuj = s => s == null ? '' : String(s).normalize('NFC').toUpperCase().split(/\s+/).filter(Boolean).join(' ');

// Ključ za uparivanje imena. Toleriše kvačice i viška razmake, da se
// 'MILAN FUŠTIĆ' i 'Milan Fuštić' prepoznaju kao ista osoba.
const kljuc = s => normalizuj(s)
  .replace(/Đ/g, 'D')
  .normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/DJ/g, 'D');

const lijepoIme = s => String(s || '').trim().split(/\s+/)
  .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');

// 'SRK VODENE LISICE' -> 'SRK Vodene lisice'; kratke skraćenice ostaju velike.
function lijepKlub(s) {
  return String(s || '').trim().split(/\s+/).map((w, i) => {
    if (i === 0 || w.length <= 2) return w.toUpperCase();
    if (i === 1) return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
    return w.toLowerCase();
  }).join(' ');
}

const ALIAS_INDEX = (() => {
  const m = new Map();
  for (const [kanon, varijante] of Object.entries(KLUB_ALIASI)) {
    m.set(kljuc(kanon), kanon);
    for (const v of varijante) m.set(kljuc(v), kanon);
  }
  return m;
})();

const kanonKlub = raw => ALIAS_INDEX.get(kljuc(raw)) || lijepKlub(raw);

const kratkiKlub = naziv => String(naziv).split(' - ')[0];
const broj = n => Number(n || 0).toLocaleString('de-DE');            // 7540 -> 7.540
const dec = (n, d = 1) => Number(n || 0).toFixed(d).replace('.', ',');
const zbir = a => a.reduce((x, y) => x + y, 0);

/* ---------- čitanje Excel fajla ---------- */

// Uvijek čitamo od ćelije A1, i kad opseg sheeta počinje od druge kolone.
// Inače bi se sve kolone pomjerile ulijevo.
function asGrid(ws) {
  if (!ws || !ws['!ref']) return [];
  const r = XLSX.utils.decode_range(ws['!ref']);
  r.s.c = 0; r.s.r = 0;
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true, range: r });
}
const cell = (g, r, c) => (g[r - 1] || [])[c - 1] ?? null;   // 1-indeksirano, kao openpyxl
const maxCol = g => g.reduce((m, row) => Math.max(m, row ? row.length : 0), 0);

function procitajMeta(g) {
  let kolo = null, mjesto = '', datum = '', godina = '';
  const mc = maxCol(g);
  for (let r = 1; r <= Math.min(g.length, 12); r++) {
    for (let c = 1; c <= mc; c++) {
      const v = cell(g, r, c);
      if (!v) continue;
      const s = String(v);
      let m;
      if ((m = s.match(/Kolo\s*:\s*(\d+)/i))) kolo = parseInt(m[1], 10);
      if ((m = s.match(/Mjesto\s+održavanja\s*:\s*(.+)/i))) mjesto = m[1].trim().replace(/\.$/, '');
      if ((m = s.match(/Datum\s*:\s*(.+)/i))) {
        datum = m[1].trim().replace(/\s*godine\.?\s*$/i, '').replace(/\.$/, '');
        const y = s.match(/(\d{4})/);
        if (y) godina = y[1];
      }
    }
  }
  return { kolo, mjesto, datum, godina };
}

function procitajTabelu(g) {
  let hr = null;
  for (let r = 1; r <= g.length; r++) if (normalizuj(cell(g, r, 1)) === 'ŠIFRA EKIPE') { hr = r; break; }
  if (hr === null) throw new Error("U sheetu TABELA nije nađeno zaglavlje 'ŠIFRA EKIPE'.");

  const grupaKol = {};
  const mapa = { 'GRUPA I': 1, 'GRUPA II': 2, 'GRUPA III': 3 };
  const mc = maxCol(g);
  for (let c = 1; c <= mc; c++) {
    const v = normalizuj(cell(g, hr, c));
    if (mapa[v]) grupaKol[mapa[v]] = c;
  }
  if (Object.keys(grupaKol).length !== 3) throw new Error('U sheetu TABELA nisu nađene sve tri grupe.');

  const ekipe = [];
  for (let r = hr + 1; r <= g.length; r++) {
    const sifra = cell(g, r, 1);
    if (typeof sifra !== 'number' || sifra < 1) continue;
    const grupe = {};
    for (const [gr, c] of Object.entries(grupaKol)) {
      const v = cell(g, r, c);
      grupe[gr] = v ? String(v).trim() : null;
    }
    ekipe.push({ sifra: Math.trunc(sifra), klub: String(cell(g, r, 2) || '').trim(), grupe });
  }
  if (!ekipe.length) throw new Error('U sheetu TABELA nije nađena nijedna ekipa.');
  return ekipe;
}

/* Sheet POJEDINACNO je službena istina. Uključuje i sudijske kazne koje se ne
   vide u poenima. Vraća null ako sheet ne postoji ili nema očekivano zaglavlje. */
function procitajPojedinacno(g) {
  if (!g.length) return null;
  let hr = null;
  for (let r = 1; r <= Math.min(g.length, 10); r++) if (normalizuj(cell(g, r, 2)) === 'IME I PREZIME') { hr = r; break; }
  if (hr === null) return null;

  const kolPoena = [], kolPlasman = [];
  let kolUkP = null, kolZbir = null, kolKlub = null;
  const mc = maxCol(g);
  for (let c = 1; c <= mc; c++) {
    const v = normalizuj(cell(g, hr + 1, c));
    if (v.startsWith('POENA')) kolPoena.push(c);
    else if (v === 'SEKTORSKI PLASMAN') kolPlasman.push(c);
    else if (v === 'UKUPNO POENA') kolUkP = c;
    else if (v.startsWith('ZBIR SEKTORSKIH')) kolZbir = c;
    if (normalizuj(cell(g, hr, c)) === 'NAZIV KLUBA') kolKlub = c;
  }
  if (kolPoena.length < 3 || kolPlasman.length < 3 || !kolUkP || !kolZbir) return null;

  const out = new Map();
  for (let r = hr + 2; r <= g.length; r++) {
    const ime = cell(g, r, 2);
    if (!ime || cell(g, r, kolZbir) == null) continue;
    const sesije = [0, 1, 2].map(i => ({
      poena: Number(cell(g, r, kolPoena[i]) || 0),
      plasman: Number(cell(g, r, kolPlasman[i]) || 0),
    }));
    out.set(kljuc(ime), {
      ime: lijepoIme(ime),
      klub: kolKlub ? kanonKlub(cell(g, r, kolKlub)) : '',
      sesije,
      poena: Number(cell(g, r, kolUkP) || 0),
      plasman: Number(cell(g, r, kolZbir) || 0),
    });
  }
  return out.size ? out : null;
}

/* Iz GRUPA sheeta po sesiji: broj riba, poeni i najduža riba.
   {imeKljuc: {1: {riba, poena, najduza}, 2: {...}, 3: {...}}} */
function procitajGrupu(g) {
  const imenaR = [], poeniR = [], ribaR = [], duzR = [];
  for (let r = 1; r <= g.length; r++) {
    const v = normalizuj(cell(g, r, 2));
    if (v === 'IME I PREZIME') imenaR.push(r + 1);
    else if (v === 'UKUPNO POENA') poeniR.push(r);
    else if (v === 'UKUPNO RIBA') ribaR.push(r);
    else if (v === 'NAJDUŽA RIBA') duzR.push(r);
  }
  const n = Math.min(imenaR.length, poeniR.length, ribaR.length);
  const out = new Map();
  const mc = maxCol(g);
  for (let s = 0; s < n; s++) {
    for (let c = 2; c <= mc; c += 2) {
      const ime = cell(g, imenaR[s], c);
      if (!ime) continue;
      const norm = normalizuj(ime);
      if (norm === 'DUŽINA-CM' || norm === 'IME I PREZIME') continue;
      const k = kljuc(ime);
      if (!out.has(k)) out.set(k, {});
      out.get(k)[s + 1] = {
        riba: Number(cell(g, ribaR[s], c + 1) || 0),
        poena: Number(cell(g, poeniR[s], c + 1) || 0),
        najduza: duzR[s] ? Number(cell(g, duzR[s], c + 1) || 0) : 0,
      };
    }
  }
  return out;
}

function parsirajKolo(buf, izvor) {
  const wb = XLSX.read(buf, { type: 'array' });
  const sheet = ime => wb.Sheets[ime] ? asGrid(wb.Sheets[ime]) : [];
  const tabela = sheet('TABELA');
  if (!tabela.length) throw new Error('Fajl nema sheet TABELA, ne izgleda kao fajl kola.');

  const ekipe = procitajTabelu(tabela);
  const meta = procitajMeta(tabela);
  const grupe = { 1: procitajGrupu(sheet('GRUPA I')), 2: procitajGrupu(sheet('GRUPA II')), 3: procitajGrupu(sheet('GRUPA III')) };
  const pojedinacno = procitajPojedinacno(sheet('POJEDINACNO'));

  const kolo = { izvor, ekipe, grupe, ...meta };
  kolo.rezultati = rezultatiKola(kolo, pojedinacno);
  kolo.sluzbeno = !!pojedinacno;
  return kolo;
}

/* Rezultat kola po takmičaru. Ako postoji POJEDINACNO, to je istina.
   Ako ga nema, plasmani se računaju iz poena: 0 poena -> broj takmičara u grupi,
   inače rang po poenima unutar grupe (isti fallback kao u popuni_tabelu.py). */
function rezultatiKola(kolo, pojedinacno) {
  const out = new Map();
  const info = new Map();     // imeKljuc -> {grupa, klub, sifra}
  for (const e of kolo.ekipe) {
    for (const gr of [1, 2, 3]) {
      const ime = e.grupe[gr];
      if (ime) info.set(kljuc(ime), { grupa: Number(gr), klub: kanonKlub(e.klub), sifra: e.sifra });
    }
  }

  const izGrupe = (gr, k, polje) => {
    const p = gr && kolo.grupe[gr] && kolo.grupe[gr].get(k);
    return [1, 2, 3].map(s => (p && p[s] ? p[s][polje] : 0));
  };

  if (pojedinacno) {
    for (const [k, v] of pojedinacno) {
      const i = info.get(k) || {};
      out.set(k, {
        kljuc: k, ime: v.ime, klub: v.klub || i.klub || '', grupa: i.grupa || null, sifra: i.sifra || null,
        sesije: v.sesije,
        riba: izGrupe(i.grupa, k, 'riba'),
        najduza: izGrupe(i.grupa, k, 'najduza'),
        poena: v.poena, plasman: v.plasman,
      });
    }
    return out;
  }

  const maxNeg = kolo.ekipe.length;
  for (const gr of [1, 2, 3]) {
    const imena = kolo.ekipe.map(e => e.grupe[gr]).filter(Boolean);
    for (const s of [1, 2, 3]) {
      const poeniSesije = new Map(imena.map(i => {
        const p = kolo.grupe[gr].get(kljuc(i));
        return [kljuc(i), p && p[s] ? p[s].poena : 0];
      }));
      for (const ime of imena) {
        const k = kljuc(ime);
        const p = poeniSesije.get(k);
        const neg = p === 0 ? maxNeg : 1 + [...poeniSesije.values()].filter(v => v > p).length;
        if (!out.has(k)) {
          const i = info.get(k) || {};
          out.set(k, {
            kljuc: k, ime: lijepoIme(ime), klub: i.klub || '', grupa: gr, sifra: i.sifra || null,
            sesije: [], riba: izGrupe(gr, k, 'riba'), najduza: izGrupe(gr, k, 'najduza'),
            poena: 0, plasman: 0,
          });
        }
        const rec = out.get(k);
        rec.sesije[s - 1] = { poena: p, plasman: neg };
        rec.poena += p;
        rec.plasman += neg;
      }
    }
  }
  return out;
}

/* ---------- sezona ---------- */

// Manji zbir sektorskih plasmana je bolji; kod istog zbira odlučuje više poena.
const poredak = (a, b) => (a.plasman - b.plasman) || (b.poena - a.poena);

function sezona(kola) {
  kola = kola.slice().sort((a, b) => (a.kolo || 0) - (b.kolo || 0));

  const uk = new Map();          // pojedinačno, kumulativno
  const ukEk = new Map();        // ekipno, kumulativno
  const snapshots = [], snapshotsEk = [];

  for (const kolo of kola) {
    const rk = kolo.rezultati;

    // odsutan iz ovog kola -> 0 poena, +kazna plasmana
    for (const [k, v] of uk) if (!rk.has(k)) v.plasman += KAZNA_PRESKOK;

    for (const [k, v] of rk) {
      if (uk.has(k)) {
        const z = uk.get(k);
        z.ime = v.ime; z.klub = v.klub || z.klub;
        z.poena += v.poena; z.plasman += v.plasman;
      } else {
        // nov takmičar -> kazna za sva prethodna kola
        uk.set(k, {
          kljuc: k, ime: v.ime, klub: v.klub,
          poena: v.poena,
          plasman: v.plasman + KAZNA_PRESKOK * Math.max(0, (kolo.kolo || 1) - 1),
        });
      }
    }

    // ekipno: prazan slot u ekipi = 0 poena, 3 x broj ekipa plasmana
    const maxNeg = kolo.ekipe.length;
    for (const e of kolo.ekipe) {
      const klub = kanonKlub(e.klub);
      let poena = 0, plasman = 0;
      for (const gr of [1, 2, 3]) {
        const ime = e.grupe[gr];
        const r = ime ? rk.get(kljuc(ime)) : null;
        if (r) { poena += r.poena; plasman += r.plasman; }
        else plasman += 3 * maxNeg;
      }
      const kk = kljuc(klub);
      if (!ukEk.has(kk)) ukEk.set(kk, { kljuc: kk, ime: klub, poena: 0, plasman: 0 });
      const z = ukEk.get(kk);
      z.ime = klub; z.poena += poena; z.plasman += plasman;
    }

    snapshots.push([...uk.values()].map(v => ({ ...v })).sort(poredak));
    snapshotsEk.push([...ukEk.values()].map(v => ({ ...v })).sort(poredak));
  }

  return { kola, snapshots, snapshotsEk };
}

const mjesta = snap => new Map(snap.map((v, i) => [v.kljuc, i + 1]));

function saPromjenom(snapshots) {
  if (!snapshots.length) return [];
  const sad = snapshots[snapshots.length - 1];
  const prije = snapshots.length > 1 ? mjesta(snapshots[snapshots.length - 2]) : null;
  return sad.map((v, i) => {
    const p = prije ? prije.get(v.kljuc) : null;
    return { ...v, mjesto: i + 1, promjena: p == null ? null : p - (i + 1) };
  });
}

/* ---------- statistika sezone ---------- */

function statistika(sez) {
  const kola = sez.kola;
  const brojSesija = kola.length * 3;

  const t = new Map();   // kljuc -> agregat po takmičaru
  const rekordSesija = [];   // sve sesije sezone, za rekorde

  kola.forEach((kolo, ki) => {
    const poredani = [...kolo.rezultati.values()].sort(poredak);
    const mjestoUKolu = new Map(poredani.map((r, i) => [r.kljuc, i + 1]));

    for (const r of poredani) {
      if (!t.has(r.kljuc)) {
        t.set(r.kljuc, {
          kljuc: r.kljuc, ime: r.ime, klub: r.klub,
          kolaOdigrao: 0, poena: 0, riba: 0, plasman: 0,
          sesija: 0, nule: 0, najduza: 0,
          mjesta: [], sesijskiPlasmani: [],
        });
      }
      const a = t.get(r.kljuc);
      a.klub = r.klub || a.klub;
      a.kolaOdigrao++;
      a.poena += r.poena;
      a.plasman += r.plasman;
      a.riba += zbir(r.riba);
      a.mjesta.push({ kolo: kolo.kolo, mjesto: mjestoUKolu.get(r.kljuc) });
      a.najduza = Math.max(a.najduza, ...r.najduza);

      for (let s = 0; s < 3; s++) {
        const ses = r.sesije[s] || { poena: 0, plasman: 0 };
        const rib = r.riba[s] || 0;
        a.sesija++;
        a.sesijskiPlasmani.push(ses.plasman);
        if (rib === 0) a.nule++;
        rekordSesija.push({
          ime: r.ime, klub: r.klub, kolo: kolo.kolo, kolIdx: ki, sesija: s + 1,
          riba: rib, poena: ses.poena, plasman: ses.plasman, najduza: r.najduza[s] || 0,
        });
      }
    }
  });

  const krajnjaMjesta = mjesta(sez.snapshots[sez.snapshots.length - 1] || []);
  const perTakmicar = [...t.values()].map(a => ({
    mjesto: krajnjaMjesta.get(a.kljuc) ?? null,
    ...a,
    prosjekPoena: a.sesija ? a.poena / a.sesija : 0,
    prosjekRiba: a.sesija ? a.riba / a.sesija : 0,
    prosjekPlasmana: a.sesijskiPlasmani.length ? zbir(a.sesijskiPlasmani) / a.sesijskiPlasmani.length : 0,
    najbolje: a.mjesta.length ? Math.min(...a.mjesta.map(m => m.mjesto)) : null,
    najgore: a.mjesta.length ? Math.max(...a.mjesta.map(m => m.mjesto)) : null,
    raspon: a.mjesta.length ? Math.max(...a.mjesta.map(m => m.mjesto)) - Math.min(...a.mjesta.map(m => m.mjesto)) : null,
  }));

  // klubovi
  const kl = new Map();
  for (const a of perTakmicar) {
    if (!a.klub) continue;
    const k = kljuc(a.klub);
    if (!kl.has(k)) kl.set(k, { ime: a.klub, poena: 0, riba: 0, takmicara: 0, sesija: 0 });
    const c = kl.get(k);
    c.poena += a.poena; c.riba += a.riba; c.takmicara++; c.sesija += a.sesija;
  }
  const perKlub = [...kl.values()].map(c => ({
    ...c,
    prosjekPoSesiji: c.sesija ? c.poena / c.sesija : 0,
    prosjekRibaPoSesiji: c.sesija ? c.riba / c.sesija : 0,
  })).sort((a, b) => b.prosjekPoSesiji - a.prosjekPoSesiji);

  // rekordi
  const najSesija = rekordSesija.slice().sort((a, b) => (b.riba - a.riba) || (b.poena - a.poena))[0] || null;
  const najPoenaSesija = rekordSesija.slice().sort((a, b) => b.poena - a.poena)[0] || null;
  const najDuza = rekordSesija.slice().sort((a, b) => b.najduza - a.najduza)[0] || null;
  const najRiba = perTakmicar.slice().sort((a, b) => b.riba - a.riba)[0] || null;
  const konstantan = perTakmicar
    .filter(a => a.kolaOdigrao === kola.length && kola.length >= 2)
    .sort((a, b) => (a.raspon - b.raspon) || (a.prosjekPlasmana - b.prosjekPlasmana))[0] || null;

  // najveći skok u ukupnom plasmanu između dva uzastopna kola
  let skok = null;
  for (let i = 1; i < sez.snapshots.length; i++) {
    const prije = mjesta(sez.snapshots[i - 1]);
    sez.snapshots[i].forEach((v, idx) => {
      const p = prije.get(v.kljuc);
      if (p == null) return;
      const d = p - (idx + 1);
      if (d > 0 && (!skok || d > skok.d)) skok = { ime: v.ime, d, sa: p, na: idx + 1, kolo: kola[i].kolo };
    });
  }

  const ukupno = {
    takmicara: perTakmicar.length,
    klubova: perKlub.length,
    kola: kola.length,
    sesija: brojSesija,
    riba: zbir(perTakmicar.map(a => a.riba)),
    poena: zbir(perTakmicar.map(a => a.poena)),
    nule: zbir(perTakmicar.map(a => a.nule)),
    najduza: najDuza ? najDuza.najduza : 0,
  };
  ukupno.prosjekRiba = ukupno.sesija && perTakmicar.length
    ? ukupno.riba / zbir(perTakmicar.map(a => a.sesija)) : 0;

  return { ukupno, perTakmicar, perKlub, rekordi: { najSesija, najPoenaSesija, najDuza, najRiba, konstantan, skok } };
}

/* ---------- render ---------- */

const $ = sel => document.querySelector(sel);
const el = (tag, cls, html) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (html != null) n.innerHTML = html;
  return n;
};

// 'Vlatko Peković' -> 'V. Peković'. Puni inicijali ('VP') se ne razlikuju
// dovoljno kad u ligi ima 29 takmičara, a prezime staje i na uski ekran.
function kratkoIme(ime) {
  const d = String(ime || '').trim().split(/\s+/);
  return d.length < 2 ? String(ime || '') : d[0].charAt(0) + '. ' + d.slice(1).join(' ');
}

function promjenaHtml(p) {
  if (p == null) return '<span class="chg flat">novo</span>';
  if (p > 0) return `<span class="chg up">▲ ${p}</span>`;
  if (p < 0) return `<span class="chg down">▼ ${-p}</span>`;
  return '<span class="chg flat">0</span>';
}

const medalja = m => m === 1 ? 'm1' : m === 2 ? 'm2' : m === 3 ? 'm3' : '';

const S = {
  kola: [],
  sez: null,
  stat: null,
  tab: 'rang',
  round: 0,
  izabrani: [],
  sort: {},          // {idTabele: {key, dir}}
};

/* ---------- jedna tabela za cijeli sajt ----------
   Sortiranje klikom na zaglavlje, zakovane kolone "Mj." i ime pri horizontalnom
   skrolu, i skraćeno ime na uskom ekranu.

   kolona: { key, lbl, tip, uloga, asc, render, kratko }
     tip    'txt' | 'int' | 'dec' | 'dec0'
     uloga  'mjesto' ili 'ime' -> kolona ostaje zakovana pri skrolu
     asc    prvi klik sortira rastuće (za plasmane, gdje je manje bolje)
     render funkcija(red) -> HTML ćelije
     kratko funkcija(red) -> kraća verzija za telefon (podrazumijevano kratkoIme)  */

const FORMAT = {
  int: v => broj(v),
  dec: v => dec(v, 1),
  dec0: v => broj(Math.round(v)),
  txt: v => v,
};

function sadrzaj(red, k) {
  if (k.render) return k.render(red);
  const v = red[k.key];
  if (v == null || v === '') return '-';
  return (FORMAT[k.tip] || FORMAT.txt)(v);
}

function sortiraj(redovi, kolone, st) {
  const k = kolone.find(c => c.key === st.key);
  if (!k) return redovi;
  return redovi.slice().sort((a, b) => {
    const x = a[st.key], y = b[st.key];
    if (k.tip === 'txt') return st.dir * String(x ?? '').localeCompare(String(y ?? ''), 'sr');
    // prazne vrijednosti idu na kraj bez obzira na smjer sortiranja
    if (x == null && y == null) return 0;
    if (x == null) return 1;
    if (y == null) return -1;
    return st.dir * (x - y);
  });
}

function klaseKolone(k) {
  const c = [];
  if (k.uloga === 'mjesto') c.push('c-pos');
  if (k.uloga === 'ime') c.push('c-name');
  if (k.tip !== 'txt' && k.uloga !== 'mjesto') c.push('r');
  if (k.klasa) c.push(k.klasa);
  return c.join(' ');
}

function crtajTabelu(id, prikaz, kolone, redovi, opcije = {}) {
  if (!S.sort[id]) S.sort[id] = { ...(opcije.pocetni || { key: kolone[0].key, dir: 1 }) };
  const st = S.sort[id];
  const poredani = sortiraj(redovi, kolone, st);

  const th = kolone.map(k => {
    const strelica = st.key === k.key ? (st.dir === 1 ? ' ▲' : ' ▼') : '';
    return `<th class="${klaseKolone(k)} sortable" data-tabela="${id}" data-prikaz="${prikaz}"` +
      ` data-key="${k.key}" data-tip="${k.tip}" data-asc="${k.asc ? 1 : 0}"` +
      ` title="Sortiraj po: ${k.lbl}">${k.lbl}${strelica}</th>`;
  }).join('');

  const tr = poredani.map(r => {
    const celije = kolone.map(k => {
      const puno = sadrzaj(r, k);
      const unutra = k.uloga === 'ime'
        ? `<span class="ime-puno">${puno}</span><span class="ime-kratko">${k.kratko ? k.kratko(r) : kratkoIme(puno)}</span>`
        : puno;
      return `<td class="${klaseKolone(k)}">${unutra}</td>`;
    }).join('');
    return `<tr class="${opcije.klasaReda ? opcije.klasaReda(r) : ''}">${celije}</tr>`;
  }).join('');

  return `<table class="tbl"><thead><tr>${th}</tr></thead><tbody>${tr}</tbody></table>`;
}

/* ---------- pojedinačno i ekipno ---------- */

const KOL_RANG = [
  { key: 'mjesto', lbl: 'Mj.', tip: 'int', uloga: 'mjesto', asc: true },
  { key: 'ime', lbl: 'Ime i prezime', tip: 'txt', uloga: 'ime' },
  { key: 'klub', lbl: 'Klub', tip: 'txt', render: r => kratkiKlub(r.klub || '-') },
  { key: 'poena', lbl: 'Poeni', tip: 'int' },
  { key: 'plasman', lbl: 'Zbir sekt. plasmana', tip: 'int', asc: true },
  { key: 'promjena', lbl: 'Promjena', tip: 'int', render: r => promjenaHtml(r.promjena) },
];

const KOL_EKIPNO = [
  { key: 'mjesto', lbl: 'Mj.', tip: 'int', uloga: 'mjesto', asc: true },
  { key: 'ime', lbl: 'Klub', tip: 'txt', uloga: 'ime', kratko: r => kratkiKlub(r.ime) },
  { key: 'poena', lbl: 'Poeni', tip: 'int' },
  { key: 'plasman', lbl: 'Zbir sekt. plasmana', tip: 'int', asc: true },
  { key: 'promjena', lbl: 'Promjena', tip: 'int', render: r => promjenaHtml(r.promjena) },
];

function renderRang() {
  const redovi = saPromjenom(S.sez.snapshots);
  $('#rang-tabela').innerHTML = crtajTabelu('rang', 'rang', KOL_RANG, redovi, {
    pocetni: { key: 'mjesto', dir: 1 },
    klasaReda: r => medalja(r.mjesto),
  });
  $('#rang-sub').textContent = `${redovi.length} takmičara, ${S.sez.kola.length} odigranih kola`;
  $('#legend-kolo').textContent = S.sez.kola.length > 1
    ? `napredovao u odnosu na ${RIMSKI[S.sez.kola[S.sez.kola.length - 2].kolo] || ''} kolo`
    : 'napredovao u odnosu na prošlo kolo';
  return S.sez.kola[S.sez.kola.length - 1];
}

function renderEkipno() {
  const redovi = saPromjenom(S.sez.snapshotsEk);
  $('#ekipno-tabela').innerHTML = crtajTabelu('ekipno', 'ekipno', KOL_EKIPNO, redovi, {
    pocetni: { key: 'mjesto', dir: 1 },
    klasaReda: r => medalja(r.mjesto),
  });
  $('#ekipno-sub').textContent = `${redovi.length} klubova`;
}

/* ---------- kola ---------- */

const KOL_GRUPA = [
  { key: 'ime', lbl: 'Takmičar', tip: 'txt', uloga: 'ime' },
  { key: 'klub', lbl: 'Klub', tip: 'txt' },
  { key: 's1', lbl: 'S1', tip: 'int', asc: true },
  { key: 's2', lbl: 'S2', tip: 'int', asc: true },
  { key: 's3', lbl: 'S3', tip: 'int', asc: true },
  { key: 'riba', lbl: 'Riba', tip: 'int' },
  { key: 'plasman', lbl: 'Zbir', tip: 'int', asc: true },
  { key: 'poena', lbl: 'Poeni', tip: 'int' },
];

function renderKola() {
  const pick = $('#round-picker');
  pick.innerHTML = '';
  S.sez.kola.forEach((k, i) => {
    const b = el('button', 'round-btn' + (i === S.round ? ' is-active' : ''), `${RIMSKI[k.kolo] || k.kolo} kolo`);
    b.type = 'button';
    b.onclick = () => { S.round = i; renderKola(); };
    pick.appendChild(b);
  });

  const kolo = S.sez.kola[S.round];
  if (!kolo) return;

  $('#round-head').innerHTML =
    `<div><div class="lbl">${RIMSKI[kolo.kolo] || kolo.kolo} kolo</div>` +
    `<div class="mjesto">${kolo.mjesto || 'Mjesto nije upisano'}</div></div>` +
    `<div class="datum">${kolo.datum || ''}</div>`;

  const svi = [...kolo.rezultati.values()];
  const gwrap = $('#round-groups');
  gwrap.innerHTML = '';
  for (const gr of [1, 2, 3]) {
    const redovi = svi.filter(r => r.grupa === gr).map(r => ({
      ime: r.ime,
      klub: kratkiKlub(r.klub || '-'),
      s1: r.sesije[0] ? r.sesije[0].plasman : null,
      s2: r.sesije[1] ? r.sesije[1].plasman : null,
      s3: r.sesije[2] ? r.sesije[2].plasman : null,
      riba: zbir(r.riba),
      plasman: r.plasman,
      poena: r.poena,
    }));
    if (!redovi.length) continue;
    const card = el('div', 'card scroll-x');
    card.appendChild(el('div', 'group-title', `Grupa ${RIMSKI[gr]}`));
    card.insertAdjacentHTML('beforeend',
      crtajTabelu(`grupa-${gr}`, 'kola', KOL_GRUPA, redovi, { pocetni: { key: 'plasman', dir: 1 } }));
    gwrap.appendChild(card);
  }

  const best = svi.slice().sort(poredak)[0];
  $('#round-best').innerHTML = best ? (
    `<div><div class="kicker">Najbolji u kolu</div>` +
    `<div class="name">${best.ime}</div><div class="club">${best.klub || ''}</div></div>` +
    `<div class="nums">` +
    `<div><div class="k">Poeni</div><div class="v">${broj(best.poena)}</div></div>` +
    `<div><div class="k">Zbir plasmana</div><div class="v">${best.plasman}</div></div>` +
    `<div><div class="k">Riba</div><div class="v">${zbir(best.riba)}</div></div>` +
    `</div>`) : '';
}

/* ---------- statistika: prikaz ---------- */

function serije() {
  const mape = S.sez.snapshots.map(mjesta);
  const zadnji = S.sez.snapshots[S.sez.snapshots.length - 1];
  return zadnji.map(v => ({
    kljuc: v.kljuc, ime: v.ime,
    v: mape.map(m => m.get(v.kljuc) ?? null),   // mjesto poslije svakog kola
  }));
}

function crtajLinije(izabrane, maxRank) {
  const n = S.sez.kola.length;
  const W = 1000, H = 430, L = 96, R = W - 104, T = 54, B = H - 70;

  // Y osa se skuplja na izabrane takmicare. Bez toga tri linije iz vrha plutaju
  // u praznoj tabeli od 29 mjesta i ne vidi se nikakva razlika medju njima.
  const sve = izabrane.flatMap(x => x.v.filter(p => p != null));
  let lo = sve.length ? Math.max(1, Math.min(...sve) - 1) : 1;
  let hi = sve.length ? Math.min(maxRank, Math.max(...sve) + 1) : maxRank;
  if (hi - lo < 3) hi = Math.min(maxRank, lo + 3);
  if (hi - lo < 3) lo = Math.max(1, hi - 3);

  const xOf = i => n === 1 ? (L + R) / 2 : L + i * ((R - L) / (n - 1));
  const yOf = p => T + (p - lo) * ((B - T) / Math.max(1, hi - lo));
  const korak = Math.max(1, Math.ceil((hi - lo + 1) / 9));

  let g = '';
  g += `<text x="${L - 18}" y="${T - 24}" text-anchor="end" style="font:600 15px Archivo,sans-serif;fill:#1f7a4d">bolje</text>`;
  g += `<text x="${L - 18}" y="${B + 30}" text-anchor="end" style="font:600 15px Archivo,sans-serif;fill:#b23b2e">slabije</text>`;

  for (let p = lo; p <= hi; p += korak) {
    g += `<line x1="${L}" x2="${R}" y1="${yOf(p)}" y2="${yOf(p)}" stroke="#ece7dd" stroke-width="2"/>` +
      `<text x="${L - 18}" y="${yOf(p) + 6}" text-anchor="end" style="font:600 17px Archivo,sans-serif;fill:#8b9295">${p}.</text>`;
  }
  // uspravne linije kola, da se vidi gdje se cita koja tacka
  S.sez.kola.forEach((k, i) => {
    g += `<line x1="${xOf(i)}" x2="${xOf(i)}" y1="${T - 10}" y2="${B + 10}" stroke="#f3efe7" stroke-width="2"/>` +
      `<text x="${xOf(i)}" y="${H - 22}" text-anchor="middle" style="font:600 18px Archivo,sans-serif;fill:#3d5157">${RIMSKI[k.kolo] || k.kolo} kolo</text>`;
  });

  izabrane.forEach((serija, si) => {
    const boja = BOJE[si % BOJE.length];
    const tacke = serija.v.map((p, i) => p == null ? null : { x: xOf(i), y: yOf(p), p }).filter(Boolean);
    if (tacke.length > 1) {
      g += `<polyline points="${tacke.map(t => t.x + ',' + t.y).join(' ')}" fill="none" stroke="${boja}" stroke-width="5" stroke-linejoin="round" stroke-linecap="round"/>`;
    }
    for (const t of tacke) {
      g += `<circle cx="${t.x}" cy="${t.y}" r="17" fill="#fff" stroke="${boja}" stroke-width="4"/>` +
        `<text x="${t.x}" y="${t.y + 6}" text-anchor="middle" style="font:700 16px Archivo,sans-serif;fill:${boja}">${t.p}</text>`;
    }
  });

  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Mjesto na rang listi poslije svakog kola">${g}</svg>`;
}

function crtajTrake(stavke, vrijednost, oznaka) {
  const maxV = Math.max(1, ...stavke.map(vrijednost));
  return stavke.map(s => {
    const w = Math.round(100 * vrijednost(s) / maxV);
    return `<div class="bar-row"><div class="bar-lbl">${kratkiKlub(s.ime)}</div>` +
      `<div class="bar-track"><div class="bar-fill" style="width:${w}%"></div></div>` +
      `<div class="bar-val">${oznaka(s)}</div></div>`;
  }).join('');
}

const KOL_STAT = [
  { key: 'mjesto', lbl: 'Mj.', tip: 'int', uloga: 'mjesto', asc: true },
  { key: 'ime', lbl: 'Takmičar', tip: 'txt', uloga: 'ime' },
  { key: 'klub', lbl: 'Klub', tip: 'txt', render: r => kratkiKlub(r.klub || '-') },
  { key: 'kolaOdigrao', lbl: 'Kola', tip: 'int' },
  { key: 'poena', lbl: 'Poeni', tip: 'int' },
  { key: 'prosjekPoena', lbl: 'Poeni / sesija', tip: 'dec0' },
  { key: 'riba', lbl: 'Riba', tip: 'int' },
  { key: 'prosjekRiba', lbl: 'Riba / sesija', tip: 'dec' },
  { key: 'najduza', lbl: 'Najduža (cm)', tip: 'int' },
  { key: 'najbolje', lbl: 'Najbolje u kolu', tip: 'int', asc: true },
  { key: 'najgore', lbl: 'Najgore u kolu', tip: 'int', asc: true },
  { key: 'prosjekPlasmana', lbl: 'Prosj. sekt. plasman', tip: 'dec', asc: true },
  { key: 'nule', lbl: 'Sesija bez ribe', tip: 'int', asc: true },
];

function renderTabelaStat() {
  $('#stat-tabela').innerHTML =
    crtajTabelu('stat', 'stat', KOL_STAT, S.stat.perTakmicar, { pocetni: { key: 'mjesto', dir: 1 } });
}

function renderStat() {
  const u = S.stat.ukupno;
  $('#stat-kpi').innerHTML = [
    ['Takmičara', broj(u.takmicara)],
    ['Klubova', broj(u.klubova)],
    ['Odigranih kola', broj(u.kola)],
    ['Ulovljenih riba', broj(u.riba)],
    ['Riba po sesiji', dec(u.prosjekRiba, 2)],
    ['Najduža riba', u.najduza ? u.najduza + ' cm' : '-'],
    ['Ukupno poena', broj(u.poena)],
    ['Sesija bez ribe', broj(u.nule)],
  ].map(([k, v]) => `<div class="kpi"><div class="kpi-k">${k}</div><div class="kpi-v">${v}</div></div>`).join('');

  // grafikon kretanja
  const sve = serije();
  const maxRank = sve.length;
  if (!S.izabrani.length) S.izabrani = sve.slice(0, 3).map(s => s.kljuc);

  const chips = $('#stat-chips');
  chips.innerHTML = '';
  sve.forEach(s => {
    const idx = S.izabrani.indexOf(s.kljuc);
    const boja = idx >= 0 ? BOJE[idx % BOJE.length] : null;
    const b = el('button', 'chip', s.ime);
    b.type = 'button';
    if (boja) { b.style.background = boja; b.style.borderColor = boja; b.style.color = '#fff'; }
    b.onclick = () => {
      if (idx >= 0) S.izabrani = S.izabrani.filter(k => k !== s.kljuc);
      else if (S.izabrani.length >= MAX_SERIJA) S.izabrani = S.izabrani.slice(1).concat(s.kljuc);
      else S.izabrani = S.izabrani.concat(s.kljuc);
      renderStat();
    };
    chips.appendChild(b);
  });

  const izabrane = S.izabrani.map(k => sve.find(s => s.kljuc === k)).filter(Boolean);
  $('#stat-chart').innerHTML = crtajLinije(izabrane, maxRank);
  $('#stat-legend').innerHTML = izabrane.map((s, i) =>
    `<span><i style="background:${BOJE[i % BOJE.length]}"></i>${s.ime}</span>`).join('');

  // rekordi
  const r = S.stat.rekordi;
  const karta = (k, v, d) => `<div class="stat-card"><div class="k">${k}</div><div class="v">${v}</div><div class="d">${d}</div></div>`;
  let rek = '';
  if (r.najSesija) rek += karta('Najbolja sesija sezone', r.najSesija.ime,
    `${RIMSKI[r.najSesija.kolo] || r.najSesija.kolo} kolo, sesija ${r.najSesija.sesija}, ${r.najSesija.riba} riba (${broj(r.najSesija.poena)} poena)`);
  if (r.najPoenaSesija) rek += karta('Najviše poena u jednoj sesiji', r.najPoenaSesija.ime,
    `${RIMSKI[r.najPoenaSesija.kolo] || r.najPoenaSesija.kolo} kolo, sesija ${r.najPoenaSesija.sesija}, ${broj(r.najPoenaSesija.poena)} poena`);
  if (r.najRiba) rek += karta('Najviše riba u sezoni', r.najRiba.ime,
    `${r.najRiba.riba} riba u ${r.najRiba.sesija} sesija (${dec(r.najRiba.prosjekRiba, 2)} po sesiji)`);
  if (r.najDuza && r.najDuza.najduza) rek += karta('Najduža riba sezone', r.najDuza.ime,
    `${r.najDuza.najduza} cm, ${RIMSKI[r.najDuza.kolo] || r.najDuza.kolo} kolo, sesija ${r.najDuza.sesija}`);
  if (r.konstantan) rek += karta('Najkonstantniji takmičar', r.konstantan.ime,
    `Plasman od ${r.konstantan.najbolje} do ${r.konstantan.najgore} u svim kolima`);
  if (r.skok) rek += karta('Najveći skok na tabeli', r.skok.ime,
    `Sa ${r.skok.sa}. na ${r.skok.na}. mjesto poslije ${RIMSKI[r.skok.kolo] || r.skok.kolo} kola`);
  $('#stat-cards').innerHTML = rek;

  // klubovi
  $('#klub-poeni').innerHTML = crtajTrake(
    S.stat.perKlub, k => k.prosjekPoSesiji, k => broj(Math.round(k.prosjekPoSesiji)));
  $('#klub-riba').innerHTML = crtajTrake(
    S.stat.perKlub.slice().sort((a, b) => b.riba - a.riba), k => k.riba, k => broj(k.riba));

  renderTabelaStat();
}

// koji prikaz se ponovo crta poslije klika na zaglavlje tabele
const PRIKAZI = { rang: renderRang, ekipno: renderEkipno, kola: renderKola, stat: renderTabelaStat };

/* ---------- tabovi ---------- */

function prikaziTab(id) {
  if (!document.querySelector(`.panel[data-panel="${id}"]`)) id = 'rang';
  S.tab = id;
  document.querySelectorAll('.tab').forEach(b => b.classList.toggle('is-active', b.dataset.tab === id));
  document.querySelectorAll('.panel').forEach(p => { p.hidden = p.dataset.panel !== id; });
  if (id === 'kola') renderKola();
  if (id === 'stat') renderStat();
  if (location.hash.slice(1) !== id) history.replaceState(null, '', '#' + id);
}

function renderSve() {
  S.sez = sezona(S.kola);
  S.stat = statistika(S.sez);
  const zadnje = renderRang();
  renderEkipno();
  S.round = S.sez.kola.length - 1;

  $('#hero-godina').textContent = zadnje.godina || '';
  $('#hero-badge').textContent =
    `Nakon ${RIMSKI[zadnje.kolo] || zadnje.kolo} kola: ${zadnje.mjesto || 'nepoznato mjesto'}`;
  document.title = `Mušičarska liga Crne Gore ${zadnje.godina || ''}: rang lista`;
  $('#next-name').textContent = `kolo-${(zadnje.kolo || S.kola.length) + 1}.xlsx`;

  $('#loader').hidden = true;
  prikaziTab(location.hash.slice(1) || S.tab);
}

/* ---------- učitavanje ---------- */

async function ucitaj() {
  const pokusaji = [];
  for (let i = 1; i <= MAX_KOLA; i++) {
    pokusaji.push(fetch(`data/kolo-${i}.xlsx`, { cache: 'no-cache' })
      .then(r => r.ok ? r.arrayBuffer().then(b => ({ i, b })) : null)
      .catch(() => null));
  }
  const nadjeni = (await Promise.all(pokusaji)).filter(Boolean);

  const greske = [];
  for (const { i, b } of nadjeni) {
    try {
      const k = parsirajKolo(b, `kolo-${i}.xlsx`);
      if (!k.kolo) k.kolo = i;
      S.kola.push(k);
    } catch (e) {
      greske.push(`kolo-${i}.xlsx: ${e.message}`);
    }
  }

  if (!S.kola.length) {
    $('#loader').hidden = true;
    const f = $('#fatal');
    f.hidden = false;
    f.innerHTML = 'Nema učitanih kola. Ubaci fajl <code>data/kolo-1.xlsx</code> u repozitorijum.' +
      (greske.length ? '<br><br>' + greske.join('<br>') : '');
    return;
  }
  renderSve();
  if (greske.length) console.warn('Preskočeni fajlovi:', greske);
}

/* ---------- "Dodaj kolo" (pregled u browseru) ---------- */

function status(tekst, lose) {
  const s = $('#status');
  s.hidden = false;
  s.className = 'status' + (lose ? ' bad' : '');
  s.textContent = tekst;
}

async function primiFajl(f) {
  if (!f) return;
  if (!/\.xlsx?$/i.test(f.name)) {
    status(`Greška: "${f.name}" nije Excel fajl. Sačuvaj tabelu kao .xlsx i probaj ponovo.`, true);
    return;
  }
  try {
    const kolo = parsirajKolo(await f.arrayBuffer(), f.name);
    if (!kolo.kolo) throw new Error('U sheetu TABELA nije nađen broj kola (ćelija "Kolo: N").');
    const i = S.kola.findIndex(k => k.kolo === kolo.kolo);
    if (i >= 0) S.kola[i] = kolo; else S.kola.push(kolo);
    S.izabrani = [];
    renderSve();
    status(`Fajl "${f.name}" je učitan. Prepoznato ${RIMSKI[kolo.kolo] || kolo.kolo} kolo, ` +
      `${kolo.rezultati.size} takmičara, ${kolo.ekipe.length} ekipa. Rang lista je osvježena.` +
      (kolo.sluzbeno ? '' : ' Napomena: fajl nema sheet POJEDINACNO, pa su plasmani izračunati iz poena.'));
  } catch (e) {
    status(`Greška: ${e.message}`, true);
  }
}

/* ---------- start ---------- */

if (typeof document !== 'undefined') {
  document.addEventListener('click', e => {
    const th = e.target.closest('th.sortable');
    if (!th) return;
    const { tabela, prikaz, key, tip, asc } = th.dataset;
    const st = S.sort[tabela];
    S.sort[tabela] = { key, dir: st && st.key === key ? -st.dir : (tip === 'txt' || asc === '1' ? 1 : -1) };
    (PRIKAZI[prikaz] || (() => {}))();
  });

  document.getElementById('tabs').addEventListener('click', e => {
    const b = e.target.closest('.tab');
    if (b) prikaziTab(b.dataset.tab);
  });
  window.addEventListener('hashchange', () => {
    const id = location.hash.slice(1);
    if (id && id !== S.tab) prikaziTab(id);
  });

  const modal = $('#modal'), drop = $('#drop'), file = $('#file');
  $('#open-modal').onclick = () => { modal.hidden = false; $('#status').hidden = true; };
  $('#close-modal').onclick = () => { modal.hidden = true; };
  modal.addEventListener('click', e => { if (e.target === modal) modal.hidden = true; });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') modal.hidden = true; });
  $('#pick').onclick = () => file.click();
  file.onchange = () => primiFajl(file.files[0]);
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('is-over'); });
  drop.addEventListener('dragleave', e => { e.preventDefault(); drop.classList.remove('is-over'); });
  drop.addEventListener('drop', e => {
    e.preventDefault();
    drop.classList.remove('is-over');
    primiFajl(e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]);
  });
  $('#repo-link').href = REPO_DATA_URL;

  ucitaj();
}

// za test.js (node); u browseru ovo ne postoji i ne radi ništa
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { kljuc, kanonKlub, parsirajKolo, sezona, saPromjenom, statistika, poredak, asGrid, crtajLinije, S, crtajTabelu, kratkoIme, KOL_RANG };
}
