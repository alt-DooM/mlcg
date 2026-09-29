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
  'SRK Plavsko jezero - Plav':     ['SRK PLAVSKO JEZERO', 'SRK PLAVSKO JEZERO - PLAV'],
  'SRK Tara - Mojkovac':           ['SRK TARA - MOJKOVAC', 'SRK TARA MK'],
  'SRK EPCG':                      ['SRK EPCG'],
};

const RIMSKI = ['', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];
const BOJE = ['#0f4c53', '#c8712a', '#2f6b45', '#7a5aa8', '#a83e4c', '#3b6b8c'];
const MAX_SERIJA = 5;              // koliko takmičara može istovremeno na grafikon
const UKUPNO_KOLA = 6;             // koliko kola ima sezona; mijenja se ovdje
const SIMULACIJA = { broj: 10000, ublazavanje: 6, sjeme: 20260914 };

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
      // svaka ulovljena riba je svoj red: duzina u cm, pa poeni
      const duzine = [];
      for (let rr = imenaR[s] + 2; rr < poeniR[s]; rr++) {
        const d = cell(g, rr, c);
        if (typeof d === 'number' && d > 0) duzine.push(d);
      }
      out.get(k)[s + 1] = {
        riba: Number(cell(g, ribaR[s], c + 1) || 0),
        poena: Number(cell(g, poeniR[s], c + 1) || 0),
        najduza: duzR[s] ? Number(cell(g, duzR[s], c + 1) || 0) : 0,
        duzine,
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
  const duzineIz = (gr, k) => {
    const p = gr && kolo.grupe[gr] && kolo.grupe[gr].get(k);
    return [1, 2, 3].map(s => (p && p[s] && p[s].duzine ? p[s].duzine : []));
  };

  if (pojedinacno) {
    for (const [k, v] of pojedinacno) {
      const i = info.get(k) || {};
      out.set(k, {
        kljuc: k, ime: v.ime, klub: v.klub || i.klub || '', grupa: i.grupa || null, sifra: i.sifra || null,
        sesije: v.sesije,
        riba: izGrupe(i.grupa, k, 'riba'),
        najduza: izGrupe(i.grupa, k, 'najduza'),
        duzine: duzineIz(i.grupa, k),
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
            duzine: duzineIz(gr, k),
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

/* Sudijske kazne iz data/kazne.json. Fajl saveza ih ne nosi (vidi IV kolo i
   žuti karton Nikole Trebješanina), a zvanični dokument ih ima. Drže se odvojeno
   da prežive ponovni upload originalnog xlsx-a. Zove se tačno jednom po kolu,
   odmah poslije parsiranja, jer sabira na postojeći plasman. */
function primijeniKazne(kolo, kazne) {
  const primijenjene = [];
  for (const k of (kazne || [])) {
    if (Number(k.kolo) !== Number(kolo.kolo)) continue;
    const r = kolo.rezultati.get(kljuc(k.takmicar));
    if (!r) {
      console.warn(`Kazna za "${k.takmicar}" u ${k.kolo}. kolu: takmičar nije nađen, kazna nije primijenjena.`);
      continue;
    }
    const d = Number(k.plasman || 0);
    r.plasman += d;
    r.kazna = (r.kazna || 0) + d;
    r.razlogKazne = k.razlog || '';
    primijenjene.push({ ime: r.ime, plasman: d, razlog: k.razlog || '' });
  }
  kolo.kazne = primijenjene;
  return primijenjene;
}

/* ---------- sezona ---------- */

// Manji zbir sektorskih plasmana je bolji; kod istog zbira odlučuje više poena.
const poredak = (a, b) => (a.plasman - b.plasman) || (b.poena - a.poena);

/* Ekipni rezultat jednog kola: {klubKljuc: {ime, poena, plasman, takmicari}}.
   Prazan slot u ekipi (nema takmičara u grupi) nosi 0 poena i 3 x broj ekipa
   plasman-poena, isto kao u zvaničnom EKIPNI PLASMAN sheetu saveza. */
function rezultatiEkipno(kolo) {
  const maxNeg = kolo.ekipe.length;
  const out = new Map();
  for (const e of kolo.ekipe) {
    const klub = kanonKlub(e.klub);
    let poena = 0, plasman = 0;
    const takmicari = [];
    for (const gr of [1, 2, 3]) {
      const ime = e.grupe[gr];
      const r = ime ? kolo.rezultati.get(kljuc(ime)) : null;
      if (r) { poena += r.poena; plasman += r.plasman; takmicari.push(r); }
      else plasman += 3 * maxNeg;
    }
    out.set(kljuc(klub), { ime: klub, poena, plasman, takmicari });
  }
  return out;
}

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

    for (const [kk, v] of rezultatiEkipno(kolo)) {
      if (!ukEk.has(kk)) ukEk.set(kk, { kljuc: kk, ime: v.ime, poena: 0, plasman: 0 });
      const z = ukEk.get(kk);
      z.ime = v.ime; z.poena += v.poena; z.plasman += v.plasman;
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

/* ---------- karton rijeke ---------- */

/* Bodovanje je linearno i provjereno na svih 864 ribe: svaki centimetar nosi
   20 poena, a svaka riba jos ravnih 100 povrh toga. Zato dvije ribe od 20cm
   (1000) nose vise od jedne od 40cm (900): broj riba uvijek tuce velicinu. */
const POENI_PO_CM = 20, POENI_PO_RIBI = 100;
const poeniZaDuzinu = cm => POENI_PO_CM * cm + POENI_PO_RIBI;

const KANTE = [[20, 24], [25, 29], [30, 34], [35, 60]];

function medijana(a) {
  if (!a.length) return 0;
  const s = a.slice().sort((x, y) => x - y);
  const m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}

/* Profil svakog terena iz svih odigranih kola na njemu. */
function profilTerena(sez) {
  const po = new Map();
  for (const kolo of sez.kola) {
    const mjesto = kolo.mjesto || 'Nepoznato mjesto';
    if (!po.has(mjesto)) {
      po.set(mjesto, {
        mjesto, kola: [], datumi: [], duzine: [], poSesiji: [0, 0, 0],
        bezRibe: [0, 0, 0], ucesnika: 0, dobitne: [], dobitniPoeni: [],
      });
    }
    const t = po.get(mjesto);
    t.kola.push(kolo.kolo);
    t.datumi.push(kolo.datum || '');
    t.ucesnika = Math.max(t.ucesnika, kolo.rezultati.size);

    for (const r of kolo.rezultati.values()) {
      for (let s = 0; s < 3; s++) {
        const d = (r.duzine && r.duzine[s]) || [];
        t.duzine.push(...d);
        t.poSesiji[s] += d.length;
        if (!d.length) t.bezRibe[s]++;
      }
    }
    // koliko riba je trebalo da se dobije sesija
    for (let s = 0; s < 3; s++) {
      let naj = null;
      for (const r of kolo.rezultati.values()) {
        const d = (r.duzine && r.duzine[s]) || [];
        const bod = d.reduce((a, cm) => a + poeniZaDuzinu(cm), 0);
        if (bod > 0 && (!naj || bod > naj.bod)) naj = { bod, riba: d.length };
      }
      if (naj) { t.dobitne.push(naj.riba); t.dobitniPoeni.push(naj.bod); }
    }
  }

  return [...po.values()].map(t => {
    const n = t.duzine.length || 1;
    return {
      ...t,
      brojKola: t.kola.length,
      ukupnoRiba: t.duzine.length,
      prosjek: t.duzine.reduce((a, b) => a + b, 0) / n,
      medijana: medijana(t.duzine),
      najveca: t.duzine.length ? Math.max(...t.duzine) : 0,
      raspodjela: KANTE.map(([a, b]) => ({
        a, b, udio: t.duzine.filter(x => x >= a && x <= b).length / n,
        broj: t.duzine.filter(x => x >= a && x <= b).length,
      })),
      dobitnaSesija: Math.round(medijana(t.dobitne)),
      dobitniPoeni: Math.round(medijana(t.dobitniPoeni)),
      // koliko sesija od ukupno je zavrsilo bez ijedne ribe
      praznihUdio: (t.bezRibe[0] + t.bezRibe[1] + t.bezRibe[2]) / (t.ucesnika * 3 * t.kola.length || 1),
    };
  }).sort((a, b) => b.ukupnoRiba - a.ukupnoRiba);
}

/* Kratki savjeti, svaki izveden iz brojki ovog terena. Namjerno se ne pise
   nista sto podaci ne nose: nema savjeta o musicama, dubini ni tehnici. */
function savjetiZaTeren(t, tereni) {
  const s = [];
  const [s1, s2, s3] = t.poSesiji;

  // kako rijeka radi kroz dan
  if (s1 && s3) {
    const odnos = s1 / s3;
    if (odnos >= 1.4) s.push({ gdje: 'sesije', ik: 'trend', broj: dec(odnos, 1) + '×', tekst: 'više ribe u prvoj nego u trećoj sesiji' });
    else if (odnos <= 1 / 1.4) s.push({ gdje: 'sesije', ik: 'trend', broj: dec(1 / odnos, 1) + '×', tekst: 'više ribe u trećoj nego u prvoj sesiji' });
    else s.push({ gdje: 'sesije', ik: 'talasi', broj: '≈', tekst: 'ulov ujednačen kroz cijeli dan' });
  }

  // veličina ribe u odnosu na ostale terene
  const sveDuzine = tereni.flatMap(x => x.duzine);
  const ligaProsjek = sveDuzine.reduce((a, b) => a + b, 0) / (sveDuzine.length || 1);
  const razlika = t.prosjek - ligaProsjek;
  if (Math.abs(razlika) >= 1) {
    s.push({
      gdje: 'duzine', ik: 'ruler', broj: (razlika > 0 ? '+' : '') + dec(razlika, 1) + ' cm',
      tekst: razlika > 0 ? 'krupnija riba nego na ostalim terenima' : 'sitnija riba nego na ostalim terenima',
    });
  }

  // koliko riba je dobijalo sesiju
  if (t.dobitnaSesija) {
    s.push({ gdje: 'bodovi', ik: 'trophy', broj: t.dobitnaSesija, tekst: 'riba je u prosjeku dobijalo sesiju' });
  }

  // gdje je najveći rizik da se ostane prazan
  const najgora = t.bezRibe.indexOf(Math.max(...t.bezRibe));
  const odUkupno = t.ucesnika * t.brojKola;
  if (odUkupno && t.bezRibe[najgora] / odUkupno >= 0.2) {
    s.push({
      gdje: 'sesije', ik: 'prazno', broj: t.bezRibe[najgora] + '/' + odUkupno,
      tekst: `praznih u ${najgora + 1}. sesiji, tu se najviše gubi`,
    });
  }

  // ima li smisla čekati kapitalca
  const krupne = t.raspodjela[t.raspodjela.length - 1].udio;
  s.push({
    gdje: 'duzine', ik: 'fish', broj: dec(100 * krupne, 0) + '%',
    tekst: krupne < 0.03
      ? 'ulova preko 35 cm, kapitalac se praktično ne javlja'
      : 'ulova preko 35 cm',
  });

  return s;
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

  // klubovi: sve se skuplja po kolima, sa pripadnoscu kluba iz tog kola,
  // a ne iz zadnjeg, da promjena kluba usred sezone ne pomjeri istoriju
  const kl = new Map();
  for (const kolo of kola) {
    const ekipno = [...rezultatiEkipno(kolo).entries()]
      .map(([kk, v]) => ({ kk, ...v }))
      .sort(poredak);
    ekipno.forEach((e, i) => {
      if (!kl.has(e.kk)) {
        kl.set(e.kk, {
          kljuc: e.kk, ime: e.ime, kolaOdigrao: 0, poena: 0, plasman: 0,
          riba: 0, sesija: 0, nule: 0, najduza: 0, mjesta: [], imena: new Set(),
        });
      }
      const c = kl.get(e.kk);
      c.ime = e.ime;
      c.kolaOdigrao++;
      c.poena += e.poena;
      c.plasman += e.plasman;
      c.mjesta.push(i + 1);
      for (const r of e.takmicari) {
        c.imena.add(r.kljuc);
        c.riba += zbir(r.riba);
        c.sesija += 3;
        c.nule += r.riba.filter(x => x === 0).length;
        c.najduza = Math.max(c.najduza, ...r.najduza);
      }
    });
  }
  const krajnjaEkipna = mjesta(sez.snapshotsEk[sez.snapshotsEk.length - 1] || []);
  const perKlub = [...kl.values()].map(c => ({
    ...c,
    mjesto: krajnjaEkipna.get(c.kljuc) ?? null,
    takmicara: c.imena.size,
    prosjekPoSesiji: c.sesija ? c.poena / c.sesija : 0,
    prosjekRibaPoSesiji: c.sesija ? c.riba / c.sesija : 0,
    najbolje: c.mjesta.length ? Math.min(...c.mjesta) : null,
    najgore: c.mjesta.length ? Math.max(...c.mjesta) : null,
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

/* ---------- rekordi sezone ----------

   Knjiga rekorda. Sve se racuna iz odigranih kola, nista se ne unosi rucno.
   Sest kartica je ranije stajalo u statistici; ovdje su zajedno sa ostalima.  */

const MIN_RIBA_ZA_PROSJEK = 20;    // ispod ovoga prosjek duzine nije rekord nego slucajnost

function rekordiSezone(sez, stat) {
  const kola = sez.kola;
  const mjestoKola = new Map(kola.map(k => [k.kolo, k.mjesto || '']));
  const rim = n => (RIMSKI[n] || n) + ' kolo';
  const rimSa = (n, nastavak) => (RIMSKI[n] || n) + ' kol' + nastavak;
  // 2, 3 i 4 traze drugi oblik: 23 ribe, ali 25 riba i 13 riba
  const malo = n => n % 10 >= 2 && n % 10 <= 4 && (n % 100 < 11 || n % 100 > 14);
  const ribe = n => n + (malo(n) ? ' ribe' : ' riba');
  const plasmana = n => n + (malo(n) ? ' sektorska' : ' sektorskih') + ' plasmana';
  const gdje = n => rim(n) + (mjestoKola.get(n) ? ', ' + mjestoKola.get(n) : '');

  const poKolu = [];            // ucinak jednog takmicara u jednom kolu
  const niz = new Map();        // kljuc -> tekuci i najduzi niz sesija sa ribom
  const duzine = new Map();     // kljuc -> sve duzine sezone
  const poTerenu = new Map();   // mjesto -> najduza riba tog terena

  for (const kolo of kola) {
    // propusteno kolo prekida niz, isto kao i sesija bez ribe
    for (const [k, n] of niz) if (!kolo.rezultati.has(k)) n.tekuci = 0;

    for (const r of kolo.rezultati.values()) {
      poKolu.push({ ime: r.ime, kolo: kolo.kolo, riba: zbir(r.riba), poena: r.poena, plasman: r.plasman });

      if (!niz.has(r.kljuc)) niz.set(r.kljuc, { ime: r.ime, tekuci: 0, najduzi: 0, kraj: null });
      const n = niz.get(r.kljuc);
      n.ime = r.ime;
      for (let s = 0; s < 3; s++) {
        if ((r.riba[s] || 0) > 0) {
          n.tekuci++;
          if (n.tekuci > n.najduzi) { n.najduzi = n.tekuci; n.kraj = { kolo: kolo.kolo, sesija: s + 1 }; }
        } else {
          n.tekuci = 0;
        }
      }

      const d = (r.duzine || []).reduce((a, x) => a.concat(x || []), []);
      if (d.length) {
        if (!duzine.has(r.kljuc)) duzine.set(r.kljuc, { ime: r.ime, sve: [] });
        duzine.get(r.kljuc).sve.push(...d);
      }

      const naj = Math.max(0, ...(r.najduza || []));
      const m = kolo.mjesto || 'Nepoznato mjesto';
      if (naj && (!poTerenu.has(m) || naj > poTerenu.get(m).cm)) {
        poTerenu.set(m, { mjesto: m, cm: naj, ime: r.ime, kolo: kolo.kolo });
      }
    }
  }

  const najRibaKolo = poKolu.slice().sort((a, b) => (b.riba - a.riba) || (b.poena - a.poena))[0] || null;
  const najKolo = poKolu.slice().sort(poredak)[0] || null;
  const najNiz = [...niz.values()].sort((a, b) => b.najduzi - a.najduzi)[0] || null;
  const najKrupniji = [...duzine.values()]
    .filter(x => x.sve.length >= MIN_RIBA_ZA_PROSJEK)
    .map(x => ({ ime: x.ime, riba: x.sve.length, prosjek: zbir(x.sve) / x.sve.length }))
    .sort((a, b) => b.prosjek - a.prosjek)[0] || null;

  const r = (stat && stat.rekordi) || {};
  const karte = [];
  const dodaj = (ik, k, v, d) => { if (v) karte.push({ ik, k, v, d }); };

  if (r.najDuza && r.najDuza.najduza) dodaj('ruler', 'Najduža riba sezone', r.najDuza.ime,
    `${r.najDuza.najduza} cm, ${gdje(r.najDuza.kolo)}, ${r.najDuza.sesija}. sesija`);
  if (najKolo) dodaj('trophy', 'Najbolje odigrano kolo', najKolo.ime,
    `${plasmana(najKolo.plasman)}, ${gdje(najKolo.kolo)}`);
  if (r.najSesija) dodaj('fish', 'Najviše riba u jednoj sesiji', r.najSesija.ime,
    `${ribe(r.najSesija.riba)}, ${gdje(r.najSesija.kolo)}, ${r.najSesija.sesija}. sesija`);
  if (r.najPoenaSesija) dodaj('trophy', 'Najviše poena u jednoj sesiji', r.najPoenaSesija.ime,
    `${broj(r.najPoenaSesija.poena)} poena, ${gdje(r.najPoenaSesija.kolo)}, ${r.najPoenaSesija.sesija}. sesija`);
  if (najRibaKolo) dodaj('fish', 'Najviše riba u jednom kolu', najRibaKolo.ime,
    `${ribe(najRibaKolo.riba)} u tri sesije, ${gdje(najRibaKolo.kolo)}`);
  if (r.najRiba) dodaj('fish', 'Najviše riba u sezoni', r.najRiba.ime,
    `${ribe(r.najRiba.riba)} u ${r.najRiba.sesija} sesija, ${dec(r.najRiba.prosjekRiba, 2)} po sesiji`);
  if (najNiz && najNiz.najduzi > 1) dodaj('trend', 'Najduži niz sesija sa ribom', najNiz.ime,
    `${najNiz.najduzi} sesija zaredom, zaključno sa ${najNiz.kraj.sesija}. sesijom u ${rimSa(najNiz.kraj.kolo, 'u')}`);
  if (najKrupniji) dodaj('ruler', 'Najkrupnija riba u prosjeku', najKrupniji.ime,
    `${dec(najKrupniji.prosjek, 1)} cm u prosjeku, na ${najKrupniji.riba} riba`);
  if (r.konstantan) dodaj('talasi', 'Najkonstantniji takmičar', r.konstantan.ime,
    `Plasman od ${r.konstantan.najbolje} do ${r.konstantan.najgore} u svim kolima`);
  if (r.skok) dodaj('trend', 'Najveći skok na tabeli', r.skok.ime,
    `Sa ${r.skok.sa}. na ${r.skok.na}. mjesto poslije ${rimSa(r.skok.kolo, 'a')}`);

  return { karte, tereni: [...poTerenu.values()].sort((a, b) => b.cm - a.cm) };
}

/* ---------- prognoza: kalkulator titule i vjerovatnoće ----------

   Kalkulator je običan račun, ne pogađanje: uneseš pretpostavku i dobiješ
   tačan ishod. Vjerovatnoće su simulacija i mogu da promaše, pa stranica
   pokazuje i koliko su promašivale na odigranim kolima.                     */

// Isti niz slučajnih brojeva pri svakom otvaranju, da se procenti ne mijenjaju
// sami od sebe kad neko osvježi stranicu.
function mulberry32(a) {
  return function () {
    a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

// Svi sesijski plasmani koje je takmičar do sada ostvario: {kljuc: [1,9,2,...]}
function sesijskaIstorija(kola) {
  const h = new Map();
  for (const kolo of kola) {
    for (const [k, r] of kolo.rezultati) {
      if (!h.has(k)) h.set(k, []);
      for (const s of r.sesije) h.get(k).push(s.plasman);
    }
  }
  return h;
}

/* Jedna simulirana sesija. Sa malo odigranih kola je istorija preuska, pa bi
   model bio pretjerano siguran u sebe: poslije I kola je davao favoritu 72%
   a stvarnom pobjedniku 0,0%. Zato se uz stvarnu istoriju miješa i UBLAZAVANJE
   nasumičnih sesija. Provjereno na odigranim kolima: greška padne sa 0,100 na
   0,088 (nasumično pogađanje je 0,180). */
function izvuciSesiju(hist, r, ublazavanje, brojUGrupi) {
  const w = hist.length / (hist.length + ublazavanje);
  return (hist.length && r() < w)
    ? hist[Math.floor(r() * hist.length)]
    : 1 + Math.floor(r() * brojUGrupi);
}

/* Vjerovatnoće za sljedeće kolo i za titulu poslije svih preostalih kola.
   Vraća {sljedece, titula, ekipnaTitula} kao mape kljuc -> udio 0..1. */
function prognoza(sez, preostalo, opcije = {}) {
  const broj = opcije.broj || SIMULACIJA.broj;
  const ubl = opcije.ublazavanje ?? SIMULACIJA.ublazavanje;
  const r = mulberry32(opcije.sjeme ?? SIMULACIJA.sjeme);

  const kola = sez.kola;
  const zadnje = kola[kola.length - 1];
  if (!zadnje) return { sljedece: new Map(), titula: new Map(), ekipnaTitula: new Map(), broj };

  const hist = sesijskaIstorija(kola);
  const uGrupi = Math.max(3, Math.round(zadnje.rezultati.size / 3));

  // ko nastupa: svi iz posljednjeg kola
  const ucesnici = [...zadnje.rezultati.keys()];
  const trenutno = new Map(sez.snapshots[sez.snapshots.length - 1].map(v => [v.kljuc, v]));
  const klubOd = new Map([...zadnje.rezultati].map(([k, v]) => [k, kljuc(v.klub || '')]));
  const klubovi = [...new Set([...klubOd.values()].filter(Boolean))];

  const pobjeda = new Map(ucesnici.map(k => [k, 0]));
  const top3 = new Map(ucesnici.map(k => [k, 0]));
  const titula = new Map(ucesnici.map(k => [k, 0]));
  const ekipna = new Map(klubovi.map(k => [k, 0]));
  const ekPobjeda = new Map(klubovi.map(k => [k, 0]));
  const ekTop3 = new Map(klubovi.map(k => [k, 0]));

  const ukupno = new Map();
  for (let n = 0; n < broj; n++) {
    // startna pozicija: trenutni zbir; ko nije u tabeli krece od svog kola
    for (const k of ucesnici) ukupno.set(k, trenutno.get(k) ? trenutno.get(k).plasman : 0);

    for (let kolo = 0; kolo < Math.max(1, preostalo); kolo++) {
      const rez = ucesnici.map(k => {
        const h = hist.get(k) || [];
        let z = 0;
        for (let s = 0; s < 3; s++) z += izvuciSesiju(h, r, ubl, uGrupi);
        return { k, z: z + r() * 0.001 };   // sitan šum razbija izjednačenja
      });
      if (kolo === 0) {
        const poredani = rez.slice().sort((a, b) => a.z - b.z);
        pobjeda.set(poredani[0].k, pobjeda.get(poredani[0].k) + 1);
        for (const x of poredani.slice(0, 3)) top3.set(x.k, top3.get(x.k) + 1);

        const poKlubu = new Map(klubovi.map(k => [k, 0]));
        for (const x of rez) {
          const kk = klubOd.get(x.k);
          if (kk && poKlubu.has(kk)) poKlubu.set(kk, poKlubu.get(kk) + x.z);
        }
        const klubRang = [...poKlubu.entries()].sort((a, b) => a[1] - b[1]);
        if (klubRang.length) {
          ekPobjeda.set(klubRang[0][0], ekPobjeda.get(klubRang[0][0]) + 1);
          for (const [kk] of klubRang.slice(0, 3)) ekTop3.set(kk, ekTop3.get(kk) + 1);
        }
      }
      if (kolo < preostalo) for (const x of rez) ukupno.set(x.k, ukupno.get(x.k) + x.z);
    }

    if (preostalo > 0) {
      let naj = null;
      for (const k of ucesnici) if (naj === null || ukupno.get(k) < ukupno.get(naj)) naj = k;
      if (naj) titula.set(naj, titula.get(naj) + 1);

      // ekipno: zbir njegova tri takmičara u istom scenariju
      const poKlubu = new Map(klubovi.map(k => [k, 0]));
      for (const k of ucesnici) {
        const kk = klubOd.get(k);
        if (kk && poKlubu.has(kk)) poKlubu.set(kk, poKlubu.get(kk) + ukupno.get(k));
      }
      let najK = null;
      for (const kk of klubovi) if (najK === null || poKlubu.get(kk) < poKlubu.get(najK)) najK = kk;
      if (najK) ekipna.set(najK, ekipna.get(najK) + 1);
    }
  }

  const udio = m => new Map([...m].map(([k, v]) => [k, v / broj]));
  return {
    broj,
    sljedece: new Map(ucesnici.map(k => [k, { pobjeda: pobjeda.get(k) / broj, top3: top3.get(k) / broj }])),
    ekipnoSljedece: new Map(klubovi.map(k => [k, { pobjeda: ekPobjeda.get(k) / broj, top3: ekTop3.get(k) / broj }])),
    titula: udio(titula),
    ekipnaTitula: udio(ekipna),
  };
}

/* ---------- kalkulator ---------- */

// Raspon zbira sektorskih plasmana u jednom kolu, iz odigranih kola.
function rasponKola(sez, ekipno) {
  const svi = [];
  for (const kolo of sez.kola) {
    const v = ekipno
      ? [...rezultatiEkipno(kolo).values()].map(x => x.plasman)
      : [...kolo.rezultati.values()].map(x => x.plasman);
    svi.push(...v);
  }
  if (!svi.length) return { min: 3, max: 27 };
  return { min: Math.min(...svi), max: Math.max(...svi) };
}

/* Klizač radi sa zbirom sektorskih plasmana, a ribolovac razmišlja u mjestima.
   Ovo je prevod, iz stvarno odigranih kola: koliko je rezultata bilo bolje od
   datog zbira, podijeljeno brojem kola, plus jedan. */
function prevodUMjesto(sez, ekipno) {
  const svi = [];
  for (const kolo of sez.kola) {
    const v = ekipno
      ? [...rezultatiEkipno(kolo).values()].map(x => x.plasman)
      : [...kolo.rezultati.values()].map(x => x.plasman);
    svi.push(...v);
  }
  svi.sort((a, b) => a - b);
  const brojKola = Math.max(1, sez.kola.length);
  const ucesnika = Math.max(1, Math.round(svi.length / brojKola));
  return zbir => {
    let bolji = 0;
    while (bolji < svi.length && svi[bolji] < zbir) bolji++;
    return Math.min(ucesnika, Math.max(1, Math.round(1 + bolji / brojKola)));
  };
}

/* Konačna tabela pod pretpostavkom da svako u SVAKOM preostalom kolu ostvari
   zbir `pretpostavke.get(kljuc)`. Ko nema pretpostavku, zadržava svoj prosjek. */
function scenarij(osnova, pretpostavke, preostalo, prosjeci) {
  return osnova.map(v => {
    const po = pretpostavke.has(v.kljuc) ? pretpostavke.get(v.kljuc) : (prosjeci.get(v.kljuc) ?? 0);
    return { ...v, poKolu: po, konacni: v.plasman + po * preostalo };
  }).sort((a, b) => (a.konacni - b.konacni) || (b.poena - a.poena))
    .map((v, i) => ({ ...v, konacnoMjesto: i + 1 }));
}

/* Može li još do titule, u najboljem slučaju za njega i najgorem za ostale.
   Granice su labave (ne uzimaju u obzir da unutar grupe samo jedan može biti
   prvi), pa je ovo "matematički moguće", ne i "vjerovatno". */
function mozeDoTitule(osnova, preostalo, min, max) {
  const najbolji = Math.min(...osnova.map(v => v.plasman));
  return new Map(osnova.map(v => {
    const mojeNajbolje = v.plasman + min * preostalo;
    const tudjeNajgore = (v.plasman === najbolji
      ? Math.min(...osnova.filter(x => x.kljuc !== v.kljuc).map(x => x.plasman))
      : najbolji) + max * preostalo;
    return [v.kljuc, mojeNajbolje <= tudjeNajgore];
  }));
}

/* Najslabiji prosjek koji jos donosi titulu, uz pretpostavku da ostali odigraju
   svoj dosadasnji prosjek. Ovo je obrnuto pitanje od klizaca: ne "sta ako", nego
   "sta mu treba". Vraca {kljuc: {granica, moguce}} gdje je granica zbir po kolu. */
function staTreba(osnova, prosjeci, preostalo, min, max) {
  const out = new Map();
  for (const v of osnova) {
    let granica = null;
    for (let z = min; z <= max; z++) {
      const sc = scenarij(osnova, new Map([[v.kljuc, z]]), preostalo, prosjeci);
      if (sc[0].kljuc !== v.kljuc) break;
      granica = z;
    }
    out.set(v.kljuc, { granica, moguce: granica !== null });
  }
  return out;
}

/* Najslabiji prosjek sa kojim `ja` jos zavrsavam ispred `protivnika`, uz to da
   protivnik odigra svoju pretpostavku. Vraca null ako ni najbolji ishod ne stize.
   Isto pitanje kao staTreba, samo sto meta nije prvo mjesto nego jedan covjek. */
function granicaProtiv(osnova, pretpostavke, preostalo, min, max, ja, protivnik) {
  const nadji = (sc, k) => sc.find(v => v.kljuc === k);
  let granica = null;
  for (let z = min; z <= max; z++) {
    const sc = scenarij(osnova, new Map([[ja, z]]), preostalo, pretpostavke);
    const A = nadji(sc, ja), B = nadji(sc, protivnik);
    if (!A || !B) break;
    if (A.konacnoMjesto > B.konacnoMjesto) break;
    granica = z;
  }
  return granica;
}

/* Sve sto jednog takmicara zanima za njegovu licnu trku, bez obzira na titulu. */
function licnaTrka(osnova, pretpostavke, preostalo, min, max, ja, koliko = 4) {
  const nadji = (sc, k) => sc.find(v => v.kljuc === k);
  const ocekivano = scenarij(osnova, new Map(), preostalo, pretpostavke);
  const sad = osnova.findIndex(v => v.kljuc === ja);
  const mojeOcekivano = nadji(ocekivano, ja);
  const najbolje = nadji(scenarij(osnova, new Map([[ja, min]]), preostalo, pretpostavke), ja);
  const najgore = nadji(scenarij(osnova, new Map([[ja, max]]), preostalo, pretpostavke), ja);

  const komsije = osnova
    .filter((v, i) => v.kljuc !== ja && Math.abs(i - sad) <= koliko)
    .map(v => {
      const granica = granicaProtiv(osnova, pretpostavke, preostalo, min, max, ja, v.kljuc);
      const obrnuto = granicaProtiv(osnova, pretpostavke, preostalo, min, max, v.kljuc, ja);
      return { ...v, granica, mozeMeStici: obrnuto !== null };
    });

  return {
    sada: osnova[sad],
    ocekivano: mojeOcekivano,
    najbolje: najbolje ? najbolje.konacnoMjesto : null,
    najgore: najgore ? najgore.konacnoMjesto : null,
    komsije,
  };
}

/* Po jedan predstavnik zbira za svako mjesto, da ose matrice ne ponavljaju
   isto mjesto dvaput (zbir 3 i 4 su oba prvo mjesto). */
function zbiroviPoMjestu(uMjesto, min, max, koliko) {
  const vidjeno = new Map();
  for (let z = min; z <= max; z++) {
    const m = uMjesto(z);
    if (!vidjeno.has(m)) vidjeno.set(m, z);
  }
  return [...vidjeno.entries()].sort((a, b) => a[0] - b[0]).slice(0, koliko);
}

/* Ko je prvak za svaku kombinaciju plasmana dvojice kandidata. */
function matricaDvoboja(osnova, prosjeci, preostalo, a, b, parovi) {
  return parovi.map(([, za]) => parovi.map(([, zb]) => {
    const sc = scenarij(osnova, new Map([[a, za], [b, zb]]), preostalo, prosjeci);
    return sc[0].kljuc;
  }));
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
  kazne: [],
  preostalo: 1,
  kalkEkipno: false,
  prognoza: null,
  dvoboj: null,
  matricaCtx: null,
  licni: null,
  tereni: null,
  rekordi: null,
  teren: null,
  sljedece: null,
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
  pct: v => dec(v * 100, 1) + '%',
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
      // red sa sudijskom kaznom dobija zvjezdicu uz ime, kao i u dokumentu saveza
      const zvjezdica = k.uloga === 'ime' && r.kazna
        ? ` <span class="kazna" title="${r.razlogKazne || 'Sudijska kazna'}: +${r.kazna} plasman-poena">*</span>` : '';
      const unutra = k.uloga === 'ime'
        ? `<span class="ime-puno">${puno}</span><span class="ime-kratko">${k.kratko ? k.kratko(r) : kratkoIme(puno)}</span>${zvjezdica}`
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
      kazna: r.kazna || 0,
      razlogKazne: r.razlogKazne || '',
    }));
    if (!redovi.length) continue;
    const card = el('div', 'card scroll-x');
    card.appendChild(el('div', 'group-title', `Grupa ${RIMSKI[gr]}`));
    card.insertAdjacentHTML('beforeend',
      crtajTabelu(`grupa-${gr}`, 'kola', KOL_GRUPA, redovi, { pocetni: { key: 'plasman', dir: 1 } }));
    gwrap.appendChild(card);
  }

  $('#round-kazne').innerHTML = (kolo.kazne || []).length
    ? '* ' + kolo.kazne.map(k => `${k.ime}: ${k.razlog || 'sudijska kazna'}, +${k.plasman} plasman-poena`).join(' · ')
    : '';

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

// Da li crtamo za uski ekran. Bitno je zato sto se SVG skalira: isti viewBox
// od 1000px na telefonu se stisne na ~350px i tekst od 17px postane 6px.
// Zato uski ekran dobija svoj, uzi i visi viewBox sa krupnijim odnosima.
const uskiEkran = () => typeof window !== 'undefined' && window.innerWidth < 620;

function crtajLinije(izabrane, maxRank) {
  const n = S.sez.kola.length;
  const usko = uskiEkran();
  const W = usko ? 380 : 1000;
  const H = usko ? 440 : 430;
  const L = usko ? 46 : 96;
  const R = W - (usko ? 34 : 104);
  const T = usko ? 46 : 54;
  const B = H - (usko ? 54 : 70);
  const fOsa = usko ? 15 : 17;      // brojevi mjesta lijevo
  const fKolo = usko ? 14 : 18;     // nazivi kola ispod
  const fRub = usko ? 12 : 15;      // "bolje" / "slabije"
  const rTacka = usko ? 13 : 17;    // poluprecnik kruzica
  const fTacka = usko ? 13 : 16;    // broj u kruzicu

  // Y osa se skuplja na izabrane takmicare. Bez toga tri linije iz vrha plutaju
  // u praznoj tabeli od 29 mjesta i ne vidi se nikakva razlika medju njima.
  const sve = izabrane.flatMap(x => x.v.filter(p => p != null));
  let lo = sve.length ? Math.max(1, Math.min(...sve) - 1) : 1;
  let hi = sve.length ? Math.min(maxRank, Math.max(...sve) + 1) : maxRank;
  if (hi - lo < 3) hi = Math.min(maxRank, lo + 3);
  if (hi - lo < 3) lo = Math.max(1, hi - 3);

  const xOf = i => n === 1 ? (L + R) / 2 : L + i * ((R - L) / (n - 1));
  const yOf = p => T + (p - lo) * ((B - T) / Math.max(1, hi - lo));
  const korak = Math.max(1, Math.ceil((hi - lo + 1) / (usko ? 6 : 9)));

  let g = '';
  // na uskom ekranu nema mjesta lijevo od ose, pa oznake idu iznad i ispod
  const xRub = usko ? L - 6 : L - 18;
  const sidro = usko ? 'start' : 'end';
  g += `<text x="${usko ? 4 : xRub}" y="${T - 20}" text-anchor="${sidro}" style="font:600 ${fRub}px Archivo,sans-serif;fill:#1f7a4d">bolje</text>`;
  g += `<text x="${usko ? 4 : xRub}" y="${B + 26}" text-anchor="${sidro}" style="font:600 ${fRub}px Archivo,sans-serif;fill:#b23b2e">slabije</text>`;

  for (let p = lo; p <= hi; p += korak) {
    g += `<line x1="${L}" x2="${R}" y1="${yOf(p)}" y2="${yOf(p)}" stroke="#ece7dd" stroke-width="2"/>` +
      `<text x="${L - 10}" y="${yOf(p) + 6}" text-anchor="end" style="font:600 ${fOsa}px Archivo,sans-serif;fill:#8b9295">${p}.</text>`;
  }
  // uspravne linije kola, da se vidi gdje se cita koja tacka
  S.sez.kola.forEach((k, i) => {
    g += `<line x1="${xOf(i)}" x2="${xOf(i)}" y1="${T - 10}" y2="${B + 10}" stroke="#f3efe7" stroke-width="2"/>` +
      `<text x="${xOf(i)}" y="${H - 16}" text-anchor="middle" style="font:600 ${fKolo}px Archivo,sans-serif;fill:#3d5157">${RIMSKI[k.kolo] || k.kolo} kolo</text>`;
  });

  izabrane.forEach((serija, si) => {
    const boja = BOJE[si % BOJE.length];
    const tacke = serija.v.map((p, i) => p == null ? null : { x: xOf(i), y: yOf(p), p }).filter(Boolean);
    if (tacke.length > 1) {
      g += `<polyline points="${tacke.map(t => t.x + ',' + t.y).join(' ')}" fill="none" stroke="${boja}" stroke-width="${usko ? 4 : 5}" stroke-linejoin="round" stroke-linecap="round"/>`;
    }
    for (const t of tacke) {
      g += `<circle cx="${t.x}" cy="${t.y}" r="${rTacka}" fill="#fff" stroke="${boja}" stroke-width="${usko ? 3 : 4}"/>` +
        `<text x="${t.x}" y="${t.y + fTacka / 2.6}" text-anchor="middle" style="font:700 ${fTacka}px Archivo,sans-serif;fill:${boja}">${t.p}</text>`;
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

const KOL_KLUB = [
  { key: 'mjesto', lbl: 'Mj.', tip: 'int', uloga: 'mjesto', asc: true },
  { key: 'ime', lbl: 'Klub', tip: 'txt', uloga: 'ime', kratko: r => kratkiKlub(r.ime) },
  { key: 'kolaOdigrao', lbl: 'Kola', tip: 'int' },
  { key: 'takmicara', lbl: 'Takmičara', tip: 'int' },
  { key: 'poena', lbl: 'Poeni', tip: 'int' },
  { key: 'prosjekPoSesiji', lbl: 'Poeni / sesija', tip: 'dec0' },
  { key: 'plasman', lbl: 'Zbir sekt. plasmana', tip: 'int', asc: true },
  { key: 'riba', lbl: 'Riba', tip: 'int' },
  { key: 'prosjekRibaPoSesiji', lbl: 'Riba / sesija', tip: 'dec' },
  { key: 'najduza', lbl: 'Najduža (cm)', tip: 'int' },
  { key: 'najbolje', lbl: 'Najbolje u kolu', tip: 'int', asc: true },
  { key: 'najgore', lbl: 'Najgore u kolu', tip: 'int', asc: true },
  { key: 'nule', lbl: 'Sesija bez ribe', tip: 'int', asc: true },
];

function renderTabelaStat() {
  $('#stat-tabela').innerHTML =
    crtajTabelu('stat', 'stat', KOL_STAT, S.stat.perTakmicar, { pocetni: { key: 'mjesto', dir: 1 } });
}

function renderTabelaKlub() {
  $('#klub-tabela').innerHTML =
    crtajTabelu('statKlub', 'statKlub', KOL_KLUB, S.stat.perKlub, {
      pocetni: { key: 'mjesto', dir: 1 },
      klasaReda: r => medalja(r.mjesto),
    });
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

  // klubovi
  $('#klub-poeni').innerHTML = crtajTrake(
    S.stat.perKlub, k => k.prosjekPoSesiji, k => broj(Math.round(k.prosjekPoSesiji)));
  $('#klub-riba').innerHTML = crtajTrake(
    S.stat.perKlub.slice().sort((a, b) => b.riba - a.riba), k => k.riba, k => broj(k.riba));

  renderTabelaKlub();
  renderTabelaStat();
}

/* ---------- prognoza: prikaz ---------- */

const KOL_PROGNOZA = [
  { key: 'mjesto', lbl: 'Mj.', tip: 'int', uloga: 'mjesto', asc: true },
  { key: 'ime', lbl: 'Takmičar', tip: 'txt', uloga: 'ime', kratko: r => r.jeKlub ? kratkiKlub(r.ime) : kratkoIme(r.ime) },
  { key: 'plasman', lbl: 'Sada', tip: 'int', asc: true },
  { key: 'pobjeda', lbl: 'Pobjeda u kolu', tip: 'pct' },
  { key: 'top3', lbl: 'Top 3 u kolu', tip: 'pct' },
  { key: 'titula', lbl: 'Šansa za titulu', tip: 'pct' },
];

function preostaloKola() {
  return Math.max(0, UKUPNO_KOLA - S.sez.kola.length);
}

function osnovaZa(ekipno) {
  return ekipno ? saPromjenom(S.sez.snapshotsEk) : saPromjenom(S.sez.snapshots);
}

/* Prosječan zbir sektorskih plasmana po ODIGRANOM kolu. Namjerno se ne dijeli
   kumulativni zbir iz rang liste, jer on nosi i kazne od 30 za propuštena kola:
   Edinu Sokoloviću bi tako "prosjek" ispao 113 po kolu, a najslabije moguće
   odigrano kolo je 27. Ovdje se dijeli čist zbir iz kola u kojima je nastupio. */
function prosjeciPoKolu(ekipno) {
  const izvor = ekipno ? S.stat.perKlub : S.stat.perTakmicar;
  return new Map(izvor.map(v => [v.kljuc, v.kolaOdigrao ? Math.round(v.plasman / v.kolaOdigrao) : 0]));
}

function renderKontrole() {
  const pre = S.preostalo;
  const kola = [];
  for (let i = 1; i <= Math.max(3, preostaloKola()); i++) {
    kola.push(`<button type="button" class="round-btn${i === pre ? ' is-active' : ''}" data-preostalo="${i}">${i}</button>`);
  }
  $('#prog-kontrole').innerHTML =
    `<div class="prog-grupa"><span class="prog-lbl">Preostalo kola</span><div class="prog-dugmad">${kola.join('')}</div></div>` +
    `<div class="prog-grupa"><span class="prog-lbl">Pogled</span><div class="prog-dugmad">` +
    `<button type="button" class="round-btn${S.kalkEkipno ? '' : ' is-active'}" data-ekipno="0">Pojedinačno</button>` +
    `<button type="button" class="round-btn${S.kalkEkipno ? ' is-active' : ''}" data-ekipno="1">Ekipno</button>` +
    `</div></div>`;
}

function renderKalkulator() {
  const ekipno = S.kalkEkipno;
  const preostalo = S.preostalo;
  const osnova = osnovaZa(ekipno);
  const prosjeci = prosjeciPoKolu(ekipno);
  const { min, max } = rasponKola(S.sez, ekipno);
  const kandidati = osnova.slice(0, ekipno ? osnova.length : 8);
  const moze = mozeDoTitule(osnova, preostalo, min, max);

  const redovi = kandidati.map(v => {
    const pocetna = Math.round(prosjeci.get(v.kljuc) ?? min);
    const ime = ekipno ? kratkiKlub(v.ime) : v.ime;
    const naOsi = (S.dvoboj || []).includes(v.kljuc);
    return `<div class="kalk-red${naOsi ? ' na-osi' : ''}" data-kljuc="${v.kljuc}">
      <div class="kalk-ko"><span class="kalk-ime">${ime}</span>
        <span class="kalk-sad">sada ${v.plasman}</span>
        ${naOsi ? '<span class="kalk-osa">na osi matrice</span>' : ''}
        ${moze.get(v.kljuc) ? '' : '<span class="kalk-ne">bez šanse za titulu</span>'}</div>
      <div class="kalk-klizac">
        <input type="range" min="${min}" max="${max}" value="${pocetna}" data-kljuc="${v.kljuc}" aria-label="Prosječno mjesto kroz preostala kola za ${ime}">
        <output class="kalk-vrijednost"><b class="kalk-mj">-</b><span class="kalk-zbir">-</span></output>
      </div>
      <div class="kalk-krajnje"><span class="kalk-mjesto">-</span><span class="kalk-ukupno">-</span></div>
    </div>`;
  }).join('');

  const uMjesto = prevodUMjesto(S.sez, ekipno);
  $('#kalk-tabela').innerHTML =
    `<div class="kalk-zaglavlje"><span>${ekipno ? 'Klub' : 'Takmičar'}</span>` +
    `<span>Prosječno mjesto kroz preostala ${preostalo} kola` +
    ` (lijevo ${uMjesto(min)}, desno ${uMjesto(max)}.)</span>` +
    `<span>Kraj sezone</span></div>${redovi}`;

  osvjeziKalkulator();
}

// Mijenja samo brojeve, ne i klizače, da se ne gubi prst sa klizača dok se vuče.
function osvjeziKalkulator() {
  const ekipno = S.kalkEkipno;
  const preostalo = S.preostalo;
  const osnova = osnovaZa(ekipno);
  const prosjeci = prosjeciPoKolu(ekipno);
  const uMjesto = prevodUMjesto(S.sez, ekipno);

  const pretpostavke = new Map();
  $('#kalk-tabela').querySelectorAll('input[type=range]').forEach(el => {
    pretpostavke.set(el.dataset.kljuc, Number(el.value));
  });

  const konacno = scenarij(osnova, pretpostavke, preostalo, prosjeci);
  const poKljucu = new Map(konacno.map(v => [v.kljuc, v]));
  const prvak = konacno[0];

  $('#kalk-tabela').querySelectorAll('.kalk-red').forEach(red => {
    const v = poKljucu.get(red.dataset.kljuc);
    if (!v) return;
    red.querySelector('.kalk-mj').textContent = uMjesto(v.poKolu) + '.';
    red.querySelector('.kalk-zbir').textContent = 'zbir ' + v.poKolu;
    red.querySelector('.kalk-mjesto').textContent = v.konacnoMjesto + '.';
    red.querySelector('.kalk-ukupno').textContent = v.konacni;
    red.classList.toggle('je-prvak', v.kljuc === prvak.kljuc);
  });

  $('#kalk-ishod').innerHTML =
    `<span class="kalk-kruna">Prvak</span><b>${ekipno ? kratkiKlub(prvak.ime) : prvak.ime}</b>` +
    `<span class="kalk-detalj">sa ${prvak.konacni} sektorskih plasmana</span>`;
}

function renderStaTreba() {
  const ekipno = S.kalkEkipno;
  const preostalo = S.preostalo;
  const osnova = osnovaZa(ekipno);
  const prosjeci = prosjeciPoKolu(ekipno);
  const { min, max } = rasponKola(S.sez, ekipno);
  const uMjesto = prevodUMjesto(S.sez, ekipno);
  const treba = staTreba(osnova, prosjeci, preostalo, min, max);

  // Dvije razlicite stvari, ranije pomijesane:
  //  - `treba`: ko stize ako ostali odigraju svoj prosjek (realno)
  //  - `matematicki`: ko jos uopste moze, uz tudji potpuni podbacaj
  const matematicki = mozeDoTitule(osnova, preostalo, min, max);
  const zivi = osnova.filter(v => treba.get(v.kljuc).moguce);
  const josMogu = osnova.filter(v => matematicki.get(v.kljuc)).length;
  const rijec = ekipno ? 'klub' : 'takmičar';
  const oblik = n => n === 1 ? rijec : (rijec === 'klub' ? 'kluba' : 'takmičara');

  const kartice = zivi.map((v, i) => {
    const g = treba.get(v.kljuc).granica;
    const ime = ekipno ? kratkiKlub(v.ime) : v.ime;
    const koliko = preostalo === 1 ? 'u posljednjem kolu' : `u prosjeku kroz preostala ${preostalo} kola`;
    return `<div class="treba-karta${i === 0 ? ' vodi' : ''}">
      <div class="treba-mj">${v.mjesto}.</div>
      <div class="treba-ko"><b>${ime}</b><span>sada ${v.plasman} sektorskih plasmana</span></div>
      <div class="treba-sta"><b>${uMjesto(g)}. mjesto</b><span>${koliko}</span></div>
    </div>`;
  }).join('');

  $('#treba-lista').innerHTML = kartice || `<p class="fineprint">Nema preostalih kola.</p>`;
  $('#treba-rezime').innerHTML = zivi.length === 0
    ? 'Nema preostalih kola.'
    : `Ako ostali odigraju svoj prosjek, titulu ${zivi.length === 1 ? 'može' : 'mogu'} ` +
      `<b>${zivi.length} ${oblik(zivi.length)}</b>.` +
      (josMogu > zivi.length
        ? ` Matematički ${josMogu === 1 ? 'je može' : 'ih je mogu'} još ${josMogu}, ali samo uz tuđi potpuni podbačaj.`
        : '');

  return { osnova, prosjeci, min, max, uMjesto, zivi };
}

/* Sta je trenutno postavljeno: vrijednost klizaca ako ga ima, inace prosjek. */
function pretpostavkeSvih(osnova, prosjeci) {
  const m = new Map(prosjeci);
  const kalk = document.querySelector('#kalk-tabela');
  if (kalk) kalk.querySelectorAll('input[type=range]').forEach(el => {
    m.set(el.dataset.kljuc, Number(el.value));
  });
  return m;
}

function renderMatrica(ctx) {
  const { osnova, prosjeci, min, max, uMjesto } = ctx;
  const ekipno = S.kalkEkipno;
  const preostalo = S.preostalo;
  const imeOd = v => ekipno ? kratkiKlub(v.ime) : v.ime;

  if (!S.dvoboj || !osnova.find(v => v.kljuc === S.dvoboj[0]) || !osnova.find(v => v.kljuc === S.dvoboj[1])) {
    S.dvoboj = [osnova[0].kljuc, (osnova[1] || osnova[0]).kljuc];
  }
  const [a, b] = S.dvoboj;
  const izbor = (izabran, koji) => `<select data-dvoboj="${koji}">` +
    osnova.slice(0, 12).map(v => `<option value="${v.kljuc}"${v.kljuc === izabran ? ' selected' : ''}>${imeOd(v)}</option>`).join('') +
    '</select>';

  const parovi = zbiroviPoMjestu(uMjesto, min, max, 10);
  // ostali igraju onako kako su postavljeni klizacima, a ne nuzno svoj prosjek
  const ostali = pretpostavkeSvih(osnova, prosjeci);
  const m = matricaDvoboja(osnova, ostali, preostalo, a, b, parovi);
  const imena = new Map(osnova.map(v => [v.kljuc, imeOd(v)]));

  const glava = parovi.map(([mj]) => `<th>${mj}.</th>`).join('');
  const redovi = parovi.map(([mj], i) => {
    const celije = m[i].map(k => {
      const tip = k === a ? 'a' : k === b ? 'b' : 'c';
      return `<td class="mx mx-${tip}" title="${imena.get(k) || ''}"></td>`;
    }).join('');
    return `<tr><th>${mj}.</th>${celije}</tr>`;
  }).join('');

  $('#dvoboj-izbor').innerHTML =
    `<label>Redovi: ${izbor(a, 0)}</label><label>Kolone: ${izbor(b, 1)}</label>`;
  $('#dvoboj-matrica').innerHTML =
    `<table class="mx-tabela"><thead><tr><th class="mx-ugao">mj.</th>${glava}</tr></thead><tbody>${redovi}</tbody></table>`;
  $('#dvoboj-legenda').innerHTML =
    `<span><i class="mx-a"></i>prvak je ${imena.get(a)}</span>` +
    `<span><i class="mx-b"></i>prvak je ${imena.get(b)}</span>` +
    `<span><i class="mx-c"></i>prvak je neko treći</span>`;
}

function renderLicnaTrka(ctx) {
  const { osnova, prosjeci, min, max, uMjesto } = ctx;
  const ekipno = S.kalkEkipno;
  const preostalo = S.preostalo;
  const imeOd = v => ekipno ? kratkiKlub(v.ime) : v.ime;

  if (!S.licni || !osnova.find(v => v.kljuc === S.licni)) S.licni = osnova[0].kljuc;
  const pretpostavke = pretpostavkeSvih(osnova, prosjeci);
  const t = licnaTrka(osnova, pretpostavke, preostalo, min, max, S.licni);

  $('#licna-izbor').innerHTML = `<label>${ekipno ? 'Klub' : 'Takmičar'}: <select id="licni-select">` +
    osnova.map(v => `<option value="${v.kljuc}"${v.kljuc === S.licni ? ' selected' : ''}>${v.mjesto}. ${imeOd(v)}</option>`).join('') +
    '</select></label>';

  const kut = (k, v, d) => `<div class="kpi"><div class="kpi-k">${k}</div><div class="kpi-v">${v}</div>` +
    (d ? `<div class="licna-d">${d}</div>` : '') + '</div>';
  $('#licna-sazetak').innerHTML =
    kut('Sada', t.sada.mjesto + '.', t.sada.plasman + ' plasmana') +
    kut('Ako nastavi ovako', t.ocekivano.konacnoMjesto + '.', 'na kraju sezone') +
    kut('Najbolje moguće', t.najbolje + '.', 'sve prva mjesta') +
    kut('Najgore moguće', t.najgore + '.', 'sve posljednja mjesta');

  const ispred = t.komsije.filter(k => k.mjesto < t.sada.mjesto);
  const iza = t.komsije.filter(k => k.mjesto > t.sada.mjesto);

  const red = (k, gore) => {
    const ime = imeOd(k);
    const tekst = gore
      ? (k.granica === null
        ? '<span class="licna-ne">van domašaja</span>'
        : `prestižeš ga sa prosječno <b>${uMjesto(k.granica)}. mjesta</b>`)
      : (k.mozeMeStici
        ? `može te stići<span class="licna-tiho">ako budeš slabiji</span>`
        : '<span class="licna-ne">ne može te stići</span>');
    const odigrao = (S.stat && !ekipno) ? (S.stat.perTakmicar.find(x => x.kljuc === k.kljuc) || {}).kolaOdigrao : null;
    const propustio = odigrao != null && odigrao < S.sez.kola.length
      ? `<span class="licna-tiho">propustio ${S.sez.kola.length - odigrao} kola</span>` : '';
    return `<li><span class="licna-mj">${k.mjesto}.</span>` +
      `<span class="licna-ime">${ime}${propustio}</span>` +
      `<span class="licna-sta">${tekst}</span></li>`;
  };

  $('#licna-komsije').innerHTML =
    `<div><h4>Ispred</h4><ul class="licna-lista">${ispred.map(k => red(k, true)).join('') || '<li class="licna-prazno">Niko, na vrhu je.</li>'}</ul></div>` +
    `<div><h4>Iza</h4><ul class="licna-lista">${iza.map(k => red(k, false)).join('') || '<li class="licna-prazno">Niko, na dnu je.</li>'}</ul></div>`;
}

function renderPrognoza() {
  renderKontrole();
  const ctx = renderStaTreba();
  renderKalkulator();
  renderMatrica(ctx);
  renderLicnaTrka(ctx);
  S.matricaCtx = ctx;

  const ekipno = S.kalkEkipno;
  const preostalo = S.preostalo;
  const kljucProg = `${preostalo}`;
  if (!S.prognoza || S.prognoza.kljuc !== kljucProg) {
    S.prognoza = { kljuc: kljucProg, rez: prognoza(S.sez, preostalo) };
  }
  const p = S.prognoza.rez;

  const osnova = osnovaZa(ekipno);
  const sljedece = ekipno ? p.ekipnoSljedece : p.sljedece;
  const titule = ekipno ? p.ekipnaTitula : p.titula;
  const redovi = osnova.map(v => {
    const s = sljedece.get(v.kljuc) || { pobjeda: 0, top3: 0 };
    return {
      kljuc: v.kljuc, mjesto: v.mjesto, ime: v.ime, jeKlub: ekipno, poena: v.poena,
      plasman: v.plasman, pobjeda: s.pobjeda, top3: s.top3, titula: titule.get(v.kljuc) || 0,
    };
  });

  $('#prog-tabela').innerHTML = crtajTabelu('prognoza', 'prognoza', KOL_PROGNOZA, redovi, {
    pocetni: { key: 'titula', dir: -1 },
  });

  const sljedeceKolo = RIMSKI[S.sez.kola.length + 1] || (S.sez.kola.length + 1);
  $('#prog-uvod').innerHTML =
    `Sljedeće je <b>${sljedeceKolo} kolo</b>. Računar je odigrao ` +
    `${broj(p.broj)} zamišljenih sezona, svaki put izvlačeći svakom takmičaru sesije ` +
    `nalik onima koje je stvarno lovio, pa prebrojao koliko puta je ko bio prvi.`;
}

/* ---------- karton rijeke: prikaz ---------- */

// Lucide ikone, ucrtane direktno da sajt ne zavisi ni od kakvog CDN-a.
const IKONE = {
  fish: '<path d="M6.5 12c.94-3.46 4.94-6 8.5-6 3.56 0 6.06 2.54 7 6-.94 3.47-3.44 6-7 6s-7.56-2.53-8.5-6Z"/><path d="M18 12v.5"/><path d="M16 17.93a9.77 9.77 0 0 1 0-11.86"/><path d="M7 10.67C7 8 5.58 5.97 2.73 5.5c-1 1.5-1 5 .23 6.5-1.24 1.5-1.24 5-.23 6.5C5.58 18.03 7 16 7 13.33"/><path d="M10.46 7.26C10.2 5.88 9.17 4.24 8 3h5.8a2 2 0 0 1 1.98 1.67l.23 1.4"/><path d="m16.01 17.93-.23 1.4A2 2 0 0 1 13.8 21H9.5a5.96 5.96 0 0 0 1.49-3.98"/>',
  ruler: '<path d="M21.3 15.3a2.4 2.4 0 0 1 0 3.4l-2.6 2.6a2.4 2.4 0 0 1-3.4 0L2.7 8.7a2.41 2.41 0 0 1 0-3.4l2.6-2.6a2.41 2.41 0 0 1 3.4 0Z"/><path d="m14.5 12.5 2-2"/><path d="m11.5 9.5 2-2"/><path d="m8.5 6.5 2-2"/><path d="m17.5 15.5 2-2"/>',
  trophy: '<path d="M6 9H4.5a2.5 2.5 0 0 1 0-5H6"/><path d="M18 9h1.5a2.5 2.5 0 0 0 0-5H18"/><path d="M4 22h16"/><path d="M10 14.66V17c0 .55-.47.98-.97 1.21C7.85 18.75 7 20.24 7 22"/><path d="M14 14.66V17c0 .55.47.98.97 1.21C16.15 18.75 17 20.24 17 22"/><path d="M18 2H6v7a6 6 0 0 0 12 0V2Z"/>',
  trend: '<path d="M16 17h6v-6"/><path d="m22 17-8.5-8.5-5 5L2 7"/>',
  pin: '<path d="M20 10c0 4.993-5.539 10.193-7.399 11.799a1 1 0 0 1-1.202 0C9.539 20.193 4 14.993 4 10a8 8 0 0 1 16 0"/><circle cx="12" cy="10" r="3"/>',
  kalendar: '<path d="M8 2v4"/><path d="M16 2v4"/><rect width="18" height="18" x="3" y="4" rx="2"/><path d="M3 10h18"/>',
  termometar: '<path d="M14 4v10.54a4 4 0 1 1-4 0V4a2 2 0 0 1 4 0Z"/>',
  vjetar: '<path d="M12.8 19.6A2 2 0 1 0 14 16H2"/><path d="M17.5 8a2.5 2.5 0 1 1 2 4H2"/><path d="M9.8 4.4A2 2 0 1 1 11 8H2"/>',
  kisa: '<path d="M4 14.899A7 7 0 1 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242"/><path d="M16 14v6"/><path d="M8 14v6"/><path d="M12 16v6"/>',
  prazno: '<circle cx="12" cy="12" r="10"/><path d="m4.9 4.9 14.2 14.2"/>',
  talasi: '<path d="M2 6c.6.5 1.2 1 2.5 1C7 7 7 5 9.5 5c2.6 0 2.4 2 5 2 2.5 0 2.5-2 5-2 1.3 0 1.9.5 2.5 1"/><path d="M2 12c.6.5 1.2 1 2.5 1 2.5 0 2.5-2 5-2 2.6 0 2.4 2 5 2 2.5 0 2.5-2 5-2 1.3 0 1.9.5 2.5 1"/><path d="M2 18c.6.5 1.2 1 2.5 1 2.5 0 2.5-2 5-2 2.6 0 2.4 2 5 2 2.5 0 2.5-2 5-2 1.3 0 1.9.5 2.5 1"/>',
  sunce: '<circle cx="12" cy="12" r="4"/><path d="M12 2v2"/><path d="M12 20v2"/><path d="m4.93 4.93 1.41 1.41"/><path d="m17.66 17.66 1.41 1.41"/><path d="M2 12h2"/><path d="M20 12h2"/><path d="m6.34 17.66-1.41 1.41"/><path d="m19.07 4.93-1.41 1.41"/>',
};

const ikona = (ime, cls = '') =>
  `<svg class="ik ${cls}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"` +
  ` stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${IKONE[ime] || ''}</svg>`;

/* Histogram duzina: glavna slika kartona. Jedna traka po centimetru. */
function crtajHistogram(duzine) {
  if (!duzine.length) return '';
  const od = Math.min(...duzine), do_ = Math.max(...duzine);
  const broj = new Map();
  for (const d of duzine) broj.set(d, (broj.get(d) || 0) + 1);
  const maxN = Math.max(...broj.values());

  const W = 1000, H = 300, L = 8, R = W - 8, T = 16, B = H - 42;
  const n = do_ - od + 1;
  const sirina = (R - L) / n;
  const boja = cm => cm >= 35 ? '#c8712a' : cm >= 30 ? '#2f6b45' : cm >= 25 ? '#0f4c53' : '#4d8f96';

  let g = '';
  for (let cm = od; cm <= do_; cm++) {
    const c = broj.get(cm) || 0;
    const h = maxN ? (c / maxN) * (B - T) : 0;
    const x = L + (cm - od) * sirina;
    g += `<rect x="${x + 1}" y="${B - h}" width="${Math.max(1, sirina - 2)}" height="${h}" rx="2" fill="${boja(cm)}"><title>${cm} cm: ${c} riba</title></rect>`;
    if (cm % 5 === 0 || cm === od || cm === do_) {
      g += `<text x="${x + sirina / 2}" y="${B + 26}" text-anchor="middle" style="font:600 18px Archivo,sans-serif;fill:#5d6b6e">${cm}</text>`;
    }
  }
  g += `<line x1="${L}" x2="${R}" y1="${B}" y2="${B}" stroke="#ddd6c9" stroke-width="2"/>`;
  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Raspodjela dužina ulovljenih riba">${g}</svg>`;
}

/* Tri sesije kao tri trake: koliko je riba ulovljeno u svakoj. */
function crtajSesije(poSesiji, bezRibe, ucesnika, brojKola) {
  const max = Math.max(1, ...poSesiji);
  return poSesiji.map((n, i) => {
    const h = Math.round(100 * n / max);
    const praznih = bezRibe[i];
    const odUkupno = ucesnika * brojKola;
    return `<div class="ses">
      <div class="ses-traka"><div class="ses-punjenje" style="height:${h}%"></div></div>
      <div class="ses-broj">${broj(n)}</div>
      <div class="ses-ime">${i + 1}. sesija</div>
      <div class="ses-prazno">${ikona('prazno')}${praznih}/${odUkupno}</div>
    </div>`;
  }).join('');
}

function renderRijeke() {
  renderSljedece();
  const tereni = S.tereni || (S.tereni = profilTerena(S.sez));
  if (!tereni.length) return;
  if (!S.teren || !tereni.find(t => t.mjesto === S.teren)) {
    const ref = S.sljedece && S.sljedece.referenca;
    S.teren = (ref && tereni.find(t => t.mjesto === ref) ? ref : tereni[0].mjesto);
  }
  const t = tereni.find(x => x.mjesto === S.teren);

  $('#rijeka-izbor').innerHTML = tereni.map(x =>
    `<button type="button" class="round-btn${x.mjesto === S.teren ? ' is-active' : ''}" data-teren="${x.mjesto}">` +
    `${ikona('talasi')}${x.mjesto.replace(/^Rijeka\s+/i, '')}</button>`).join('');

  const kpi = (ime, v, k) => `<div class="kpi"><div class="kpi-k">${ikona(ime)}${k}</div><div class="kpi-v">${v}</div></div>`;
  $('#rijeka-kpi').innerHTML =
    kpi('kalendar', t.brojKola, t.brojKola === 1 ? 'odigrano kolo' : 'odigranih kola') +
    kpi('fish', broj(t.ukupnoRiba), 'ulovljenih riba') +
    kpi('ruler', dec(t.prosjek, 1) + ' cm', 'prosječna dužina') +
    kpi('trophy', t.najveca + ' cm', 'najveća ulovljena');

  $('#rijeka-histogram').innerHTML = crtajHistogram(t.duzine);
  $('#rijeka-raspodjela').innerHTML = t.raspodjela.map(r =>
    `<div class="rasp"><span class="rasp-boja" style="background:${r.a >= 35 ? '#c8712a' : r.a >= 30 ? '#2f6b45' : r.a >= 25 ? '#0f4c53' : '#4d8f96'}"></span>` +
    `<b>${dec(100 * r.udio, 0)}%</b><span>${r.a}-${r.b === 60 ? '48' : r.b} cm</span></div>`).join('');

  $('#rijeka-sesije').innerHTML = crtajSesije(t.poSesiji, t.bezRibe, t.ucesnika, t.brojKola);

  // pravilo bodovanja, pokazano brojkama a ne tekstom
  $('#rijeka-bodovi').innerHTML =
    `<div class="bod-red"><span class="bod-slika">${ikona('fish')}${ikona('fish')}</span>` +
    `<span class="bod-tekst">2 × 20 cm</span><b class="bod-v">${broj(2 * poeniZaDuzinu(20))}</b></div>` +
    `<div class="bod-vs">veće od</div>` +
    `<div class="bod-red"><span class="bod-slika">${ikona('fish', 'velika')}</span>` +
    `<span class="bod-tekst">1 × 40 cm</span><b class="bod-v">${broj(poeniZaDuzinu(40))}</b></div>`;

  const sv = savjetiZaTeren(t, tereni);
  for (const gdje of ['duzine', 'sesije', 'bodovi']) {
    const cilj = $('#savjeti-' + gdje);
    if (!cilj) continue;
    cilj.innerHTML = sv.filter(x => x.gdje === gdje).map(x =>
      `<div class="savjet">${ikona(x.ik)}<b>${x.broj}</b><span>${x.tekst}</span></div>`).join('');
  }

  renderVrijeme(t);
}

function renderRekordi() {
  const rk = S.rekordi || (S.rekordi = rekordiSezone(S.sez, S.stat));
  $('#rekordi-cards').innerHTML = rk.karte.map(c =>
    `<div class="stat-card"><div class="k">${ikona(c.ik)}${c.k}</div>` +
    `<div class="v">${c.v}</div><div class="d">${c.d}</div></div>`).join('');
  $('#rekordi-tereni').innerHTML = rk.tereni.map(t =>
    `<div class="kpi"><div class="kpi-k">${ikona('talasi')}${t.mjesto.replace(/^Rijeka\s+/i, '')}</div>` +
    `<div class="kpi-v">${t.cm} cm</div>` +
    `<div class="kpi-d">${t.ime}, ${RIMSKI[t.kolo] || t.kolo} kolo</div></div>`).join('');
}

/* ---------- vrijeme ----------
   Open-Meteo, bez registracije i bez kljuca. Temperatura vode i vodostaj se
   nigdje ne mogu besplatno povuci za crnogorske rijeke, pa ih ovdje nema.
   Sve sto ne stigne se jednostavno ne iscrta; vrijeme je dodatak, ne uslov. */

const KOORDINATE = [
  [/ćehotina|cehotina|pljevlja/i, 43.356, 19.358],
  [/tara.*kolašin|tara.*kolasin/i, 42.822, 19.517],
  [/tara.*mojkovac/i, 42.960, 19.583],
  [/lim.*berane/i, 42.844, 19.871],
  [/lim.*plav|plavsko/i, 42.598, 19.944],
];

const koordinateZa = mjesto => {
  const n = (KOORDINATE.find(([re]) => re.test(mjesto || '')) || null);
  return n ? { lat: n[1], lon: n[2] } : null;
};

// "27. 09. 2026" -> "2026-09-27"
function uISO(datum) {
  const d = String(datum || '').match(/(\d{1,2})\D+(\d{1,2})\D+(\d{4})/);
  if (!d) return null;
  return `${d[3]}-${String(d[2]).padStart(2, '0')}-${String(d[1]).padStart(2, '0')}`;
}

async function vrijemeZa(mjesto, datum) {
  const k = koordinateZa(mjesto), dan = uISO(datum);
  if (!k || !dan) return null;
  const kljucKesa = `vrijeme:${dan}:${k.lat},${k.lon}`;
  try {
    const kes = localStorage.getItem(kljucKesa);
    if (kes) return JSON.parse(kes);
  } catch (e) { /* privatni prozor, nema veze */ }

  const url = 'https://archive-api.open-meteo.com/v1/archive'
    + `?latitude=${k.lat}&longitude=${k.lon}&start_date=${dan}&end_date=${dan}`
    + '&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,wind_speed_10m_max'
    + '&timezone=Europe%2FBelgrade';
  try {
    const r = await fetch(url);
    if (!r.ok) return null;
    const j = await r.json();
    const d = j && j.daily;
    if (!d || d.temperature_2m_max == null || d.temperature_2m_max[0] == null) return null;
    const v = {
      dan,
      tmax: d.temperature_2m_max[0], tmin: d.temperature_2m_min[0],
      kisa: d.precipitation_sum[0], vjetar: d.wind_speed_10m_max[0],
    };
    try { localStorage.setItem(kljucKesa, JSON.stringify(v)); } catch (e) { /* nema veze */ }
    return v;
  } catch (e) {
    return null;
  }
}

/* Prognoza za buduci dan. Drugi endpoint od arhive, i kes traje samo tri sata
   jer se prognoza mijenja. Dalje od 16 dana Open-Meteo ne daje nista. */
async function prognozaVremena(mjesto, datum) {
  const k = koordinateZa(mjesto), dan = uISO(datum);
  if (!k || !dan) return null;
  const kljucKesa = `prognoza:${dan}:${k.lat},${k.lon}`;
  try {
    const kes = JSON.parse(localStorage.getItem(kljucKesa) || 'null');
    if (kes && Date.now() - kes.kad < 3 * 3600e3) return kes.v;
  } catch (e) { /* nema veze */ }

  const url = 'https://api.open-meteo.com/v1/forecast'
    + `?latitude=${k.lat}&longitude=${k.lon}&start_date=${dan}&end_date=${dan}`
    + '&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,wind_speed_10m_max,precipitation_probability_max'
    + '&timezone=Europe%2FBelgrade';
  try {
    const r = await fetch(url);
    if (!r.ok) return null;
    const j = await r.json();
    const d = j && j.daily;
    if (!d || !d.temperature_2m_max || d.temperature_2m_max[0] == null) return null;
    const v = {
      dan, tmax: d.temperature_2m_max[0], tmin: d.temperature_2m_min[0],
      kisa: d.precipitation_sum[0], vjetar: d.wind_speed_10m_max[0],
      sansaKise: d.precipitation_probability_max ? d.precipitation_probability_max[0] : null,
    };
    try { localStorage.setItem(kljucKesa, JSON.stringify({ kad: Date.now(), v })); } catch (e) { /* nema veze */ }
    return v;
  } catch (e) {
    return null;
  }
}

const DANI = ['nedjelja', 'ponedjeljak', 'utorak', 'srijeda', 'četvrtak', 'petak', 'subota'];

async function renderSljedece() {
  const cilj = $('#sljedece-kolo');
  const sk = S.sljedece;
  if (!cilj) return;
  if (!sk || !sk.datum) { cilj.hidden = true; return; }

  const iso = uISO(sk.datum);
  const dat = iso ? new Date(iso + 'T00:00:00') : null;
  const danas = new Date(); danas.setHours(0, 0, 0, 0);
  const zaDana = dat ? Math.round((dat - danas) / 86400e3) : null;
  const odbrojavanje = zaDana == null ? ''
    : zaDana > 1 ? `za ${zaDana} dana` : zaDana === 1 ? 'sjutra' : zaDana === 0 ? 'danas' : '';

  cilj.hidden = false;
  cilj.innerHTML =
    `<div class="sk-glava">
       <span class="sk-oznaka">${ikona('kalendar')}${RIMSKI[sk.kolo] || sk.kolo} kolo</span>
       <span class="sk-mjesto">${ikona('pin')}${(sk.mjesto || '').replace(/^Rijeka\s+/i, '')}</span>
       <span class="sk-datum">${dat ? DANI[dat.getDay()] + ', ' : ''}${sk.datum}${odbrojavanje ? ' · ' + odbrojavanje : ''}</span>
     </div>
     <div class="sk-vrijeme" id="sk-vrijeme"></div>`;

  const v = await prognozaVremena(sk.mjesto, sk.datum);
  const p = $('#sk-vrijeme');
  if (!p) return;
  p.innerHTML = v ? (
    `<span class="vr-p">${ikona('termometar')}${dec(v.tmin, 0)}–${dec(v.tmax, 0)}°</span>` +
    `<span class="vr-p">${ikona('kisa')}${dec(v.kisa, 1)} mm${v.sansaKise != null ? ' · ' + dec(v.sansaKise, 0) + '%' : ''}</span>` +
    `<span class="vr-p">${ikona('vjetar')}${dec(v.vjetar, 0)} km/h</span>`
  ) : `<span class="sk-nema">${ikona('prazno')}prognoza još nije dostupna</span>`;
}

async function renderVrijeme(t) {
  const cilj = $('#rijeka-vrijeme');
  if (!cilj) return;
  const kola = S.sez.kola.filter(k => (k.mjesto || '') === t.mjesto);
  const podaci = await Promise.all(kola.map(k => vrijemeZa(k.mjesto, k.datum)));
  if (S.teren !== t.mjesto) return;          // u međuvremenu je promijenjena rijeka

  const redovi = kola.map((k, i) => {
    const v = podaci[i];
    if (!v) return '';
    return `<div class="vr-red">
      <span class="vr-kolo">${RIMSKI[k.kolo] || k.kolo}</span>
      <span class="vr-dan">${k.datum || ''}</span>
      <span class="vr-p">${ikona('termometar')}${dec(v.tmin, 0)}–${dec(v.tmax, 0)}°</span>
      <span class="vr-p">${ikona('kisa')}${dec(v.kisa, 1)} mm</span>
      <span class="vr-p">${ikona('vjetar')}${dec(v.vjetar, 0)} km/h</span>
    </div>`;
  }).filter(Boolean).join('');

  cilj.innerHTML = redovi;
  const sekcija = $('#rijeka-vrijeme-sekcija');
  if (sekcija) sekcija.hidden = !redovi;
}

// koji prikaz se ponovo crta poslije klika na zaglavlje tabele
const PRIKAZI = { rang: renderRang, ekipno: renderEkipno, kola: renderKola,
  stat: renderTabelaStat, statKlub: renderTabelaKlub, prognoza: renderPrognoza };

/* ---------- tabovi ---------- */

function prikaziTab(id) {
  if (!document.querySelector(`.panel[data-panel="${id}"]`)) id = 'rang';
  S.tab = id;
  document.querySelectorAll('.tab').forEach(b => b.classList.toggle('is-active', b.dataset.tab === id));
  document.querySelectorAll('.panel').forEach(p => { p.hidden = p.dataset.panel !== id; });
  if (id === 'kola') renderKola();
  if (id === 'stat') renderStat();
  if (id === 'prognoza') renderPrognoza();
  if (id === 'rijeke') renderRijeke();
  if (id === 'rekordi') renderRekordi();
  if (location.hash.slice(1) !== id) history.replaceState(null, '', '#' + id);
}

function renderSve() {
  S.sez = sezona(S.kola);
  S.stat = statistika(S.sez);
  const zadnje = renderRang();
  renderEkipno();
  S.round = S.sez.kola.length - 1;
  S.preostalo = Math.max(1, UKUPNO_KOLA - S.sez.kola.length);
  S.prognoza = null;
  S.tereni = null;
  S.rekordi = null;

  $('#hero-godina').textContent = zadnje.godina || '';
  $('#hero-badge').textContent =
    `Nakon ${RIMSKI[zadnje.kolo] || zadnje.kolo} kola: ${zadnje.mjesto || 'nepoznato mjesto'}`;
  const sk = S.sljedece;
  const hs = $('#hero-sljedece');
  if (hs) {
    const jos = sk && sk.datum && !S.kola.some(k => k.kolo === sk.kolo);
    hs.hidden = !jos;
    if (jos) hs.textContent = `Sljedeće: ${RIMSKI[sk.kolo] || sk.kolo} kolo, ` +
      `${(sk.mjesto || '').replace(/^Rijeka\s+/i, '')}, ${sk.datum}`;
  }
  document.title = `Mušičarska liga Crne Gore ${zadnje.godina || ''}: rang lista`;
  $('#next-name').textContent = `kolo-${(zadnje.kolo || S.kola.length) + 1}.xlsx`;

  $('#loader').hidden = true;
  prikaziTab(location.hash.slice(1) || S.tab);
}

/* ---------- učitavanje ---------- */

async function ucitaj() {
  const kazne = await fetch('data/kazne.json', { cache: 'no-cache' })
    .then(r => r.ok ? r.json() : null)
    .then(j => (j && j.kazne) || [])
    .catch(() => []);
  S.kazne = kazne;
  S.sljedece = await fetch('data/sljedece-kolo.json', { cache: 'no-cache' })
    .then(r => r.ok ? r.json() : null)
    .catch(() => null);

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
      primijeniKazne(k, S.kazne);
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
    primijeniKazne(kolo, S.kazne);
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
  let bioUzak = uskiEkran();
  window.addEventListener('resize', () => {
    const sad = uskiEkran();
    if (sad === bioUzak) return;
    bioUzak = sad;
    if (S.tab === 'stat' && S.stat) renderStat();
  });

  document.addEventListener('click', e => {
    const rijeka = e.target.closest('#rijeka-izbor button');
    if (rijeka) { S.teren = rijeka.dataset.teren; renderRijeke(); return; }

    const dugme = e.target.closest('#prog-kontrole button');
    if (dugme) {
      if (dugme.dataset.preostalo) S.preostalo = Number(dugme.dataset.preostalo);
      if (dugme.dataset.ekipno) S.kalkEkipno = dugme.dataset.ekipno === '1';
      renderPrognoza();
      return;
    }

    const th = e.target.closest('th.sortable');
    if (!th) return;
    const { tabela, prikaz, key, tip, asc } = th.dataset;
    const st = S.sort[tabela];
    S.sort[tabela] = { key, dir: st && st.key === key ? -st.dir : (tip === 'txt' || asc === '1' ? 1 : -1) };
    (PRIKAZI[prikaz] || (() => {}))();
  });

  document.addEventListener('input', e => {
    if (e.target.matches('#kalk-tabela input[type=range]')) {
      osvjeziKalkulator();
      if (S.matricaCtx) { renderMatrica(S.matricaCtx); renderLicnaTrka(S.matricaCtx); }
    }
  });
  document.addEventListener('change', e => {
    const sel = e.target.closest('#dvoboj-izbor select');
    if (sel) {
      S.dvoboj[Number(sel.dataset.dvoboj)] = sel.value;
      renderPrognoza();
      return;
    }
    if (e.target.id === 'licni-select') {
      S.licni = e.target.value;
      if (S.matricaCtx) renderLicnaTrka(S.matricaCtx);
    }
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
  module.exports = { kljuc, kanonKlub, parsirajKolo, sezona, saPromjenom, statistika, poredak, asGrid, crtajLinije, S, crtajTabelu, kratkoIme, KOL_RANG, KOL_KLUB, rezultatiEkipno, primijeniKazne, profilTerena, poeniZaDuzinu, koordinateZa, uISO, DANI, savjetiZaTeren, rekordiSezone,
    prognoza, scenarij, mozeDoTitule, rasponKola, sesijskaIstorija, prevodUMjesto, staTreba, matricaDvoboja, zbiroviPoMjestu, licnaTrka, granicaProtiv, UKUPNO_KOLA,
    KOL_PROGNOZA };
}
