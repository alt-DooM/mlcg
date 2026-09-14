/* Provjere obračuna. Pokretanje:  node test.js

   Dva nezavisna izvora istine:
   1) POPUNJENO fajl koji je napravila Python skripta popuni_tabelu.py
      (kumulativno kroz III kolo),
   2) zvanični dokument saveza za IV kolo (EKIPNO i POJEDINACNO IV KOLO),
      prepisan u REF_IV_* ispod.                                            */

const fs = require('fs');
const path = require('path');
const assert = require('assert');

global.XLSX = require('./vendor/xlsx.full.min.js');
const { parsirajKolo, sezona, kljuc, kanonKlub, asGrid, statistika, crtajLinije, S,
        crtajTabelu, kratkoIme, KOL_RANG, KOL_KLUB, saPromjenom, rezultatiEkipno } = require('./app.js');

const KOLA = [1, 2, 3, 4].map(i => `data/kolo-${i}.xlsx`);
const REFERENCA = path.join('test-data', 'referenca-III-kolo.xlsx');

const ucitaj = p => XLSX.read(new Uint8Array(fs.readFileSync(p)), { type: 'array' });
const grid = (wb, ime) => asGrid(wb.Sheets[ime]);
const cell = (g, r, c) => (g[r - 1] || [])[c - 1] ?? null;
const poredak = (a, b) => (a.plasman - b.plasman) || (b.poena - a.poena);

// {kljuc: {ime, poena, plasman}} iz POPUNJENO sheeta (C ime, H poena, I plasman).
// Klubovi se porede u kanonskom obliku, jer se kanonski naziv mijenja kad savez
// uvede novu varijantu pisanja (npr. 'SRK PLAVSKO JEZERO - PLAV' u IV kolu).
function referenca(g, jeKlub) {
  const out = new Map();
  for (let r = 3; r <= g.length; r++) {
    const ime = cell(g, r, 3);
    if (!ime) continue;
    const k = jeKlub ? kljuc(kanonKlub(ime)) : kljuc(ime);
    out.set(k, { ime: String(ime), poena: Number(cell(g, r, 8) || 0), plasman: Number(cell(g, r, 9) || 0) });
  }
  return out;
}

const kola = KOLA.map(p => parsirajKolo(new Uint8Array(fs.readFileSync(p)), p));
const sez = sezona(kola);
const zadnji = sez.snapshots.length - 1;

assert.strictEqual(sez.kola.length, 4, 'očekivana su 4 kola');
assert.deepStrictEqual(sez.kola.map(k => k.kolo), [1, 2, 3, 4], 'brojevi kola iz TABELA sheeta');
assert.ok(sez.kola.every(k => k.sluzbeno), 'sva kola moraju koristiti službeni POJEDINACNO sheet');

/* --- 1. kumulativno kroz III kolo protiv Python izlaza --- */

const sez3 = sezona(kola.slice(0, 3));
const ref = ucitaj(REFERENCA);
let ok = 0;
for (const [sheet, nas] of [['Rang lista', sez3.snapshots[2]], ['Ekipni plasman', sez3.snapshotsEk[2]]]) {
  const oc = referenca(grid(ref, sheet), sheet === 'Ekipni plasman');
  assert.strictEqual(nas.length, oc.size, `${sheet}: broj redova (naš ${nas.length}, Python ${oc.size})`);

  const nasImena = nas.map(v => v.kljuc);
  const ocSort = [...oc.entries()].sort((a, b) => (a[1].plasman - b[1].plasman) || (b[1].poena - a[1].poena)).map(e => e[0]);
  assert.deepStrictEqual(nasImena, ocSort, `${sheet}: redosljed se ne poklapa`);

  for (const v of nas) {
    const o = oc.get(v.kljuc);
    assert.ok(o, `${sheet}: '${v.ime}' nema u Python izlazu`);
    assert.strictEqual(v.poena, o.poena, `${sheet}: '${v.ime}' poeni (naš ${v.poena}, Python ${o.poena})`);
    assert.strictEqual(v.plasman, o.plasman, `${sheet}: '${v.ime}' plasman (naš ${v.plasman}, Python ${o.plasman})`);
    ok++;
  }
  console.log(`  OK  ${sheet} kroz III kolo: ${oc.size} redova identično Python izlazu`);
}

/* --- 2. IV kolo protiv zvaničnog dokumenta saveza ---

   Prepisano iz 'EKIPNO i POJEDINACNO IV KOLO - MLCG.docx'. To je plasman
   SAMO za IV kolo, ne kumulativno.

   ZNANO NESLAGANJE: dokument vodi Nikolu Trebješanina sa 19 sektorskih bodova
   i zvjezdicom pored imena, a xlsx sa 14 (9+2+3). Razlika od 5 je sudijska
   kazna koju xlsx ne nosi. Ista razlika ide i na njegov klub, SRK Gorštak
   (45 u xlsx-u, 50 u dokumentu), i mijenja mu mjesto u kolu sa 5. na 6.
   Dok savez ne razjasni, sajt računa po xlsx-u, kao i sva ranija kola.
   Kad se razjasni: ispravi xlsx ili unesi kaznu, pa ovdje obriši KAZNA_IV. */

const KAZNA_IV = { takmicar: 'NIKOLA TREBJESANIN', klub: 'SRK GORSTAK - KOLASIN', kazna: 5 };

const REF_IV_POJEDINACNO = [
  ['IVICA RAJKOVIĆ', 10460, 4], ['PETAR OBRADOVIĆ', 10220, 5], ['BOŠKO VULEVIĆ', 9280, 6],
  ['ARSO JEREMIĆ', 10840, 7], ['VLATKO PEKOVIĆ', 10780, 7], ['MILUN ĐUROVIĆ', 6360, 9],
  ['NIKOLA ĐALOVIĆ', 6340, 10], ['RADOVAN KRKOVIĆ', 5360, 10], ['RADULE MIĆOVIĆ', 5920, 11],
  ['OMAR BAŠIĆ', 4560, 11], ['GORAN OBRENOVIĆ', 4680, 12], ['MARKO BAKIĆ', 6200, 16],
  ['DAMIR CANOVIĆ', 4740, 17], ['KEMAL AJANOVIĆ', 4580, 18], ['NIKOLA TREBJEŠANIN', 2300, 19],
  ['BOGDAN GOVEDARICA', 3220, 20], ['MILAN FUŠTIĆ', 2880, 20], ['BOGDAN MARKOVIĆ', 2820, 20],
  ['BOJAN GOVEDARICA', 1840, 20], ['ORHAN BAŠIĆ', 1720, 20], ['SAVO RAIČEVIĆ', 4280, 21],
  ['MEHMED DŽIDIĆ', 1720, 21], ['MUMIN HEKALO', 3000, 23], ['ALEKSANDAR GAŠEVIĆ', 1300, 23],
];

const REF_IV_EKIPNO = [
  ['SRK LIM - BERANE', 21560, 26], ['SRK VODENE LISICE - PODGORICA', 22760, 31],
  ['SRK KOLAŠIN', 22360, 35], ['SRFFK MANIRO - KOLAŠIN', 12860, 42],
  ['SRK PLAVSKO JEZERO - PLAV', 11020, 48], ['SRK GORŠTAK - KOLAŠIN', 13700, 50],
  ['SRK TARA - MOJKOVAC', 9340, 61], ['SRK LIPLJEN - PLJEVLJA', 9300, 62],
  ['SRFFK RAVNJAK - MOJKOVAC', 4780, 69],
];

const kolo4 = sez.kola[3];
let poredjeno = 0;
for (const [ime, poena, plasman] of REF_IV_POJEDINACNO) {
  const r = kolo4.rezultati.get(kljuc(ime));
  assert.ok(r, `IV kolo: '${ime}' nije nađen u xlsx-u`);
  assert.strictEqual(r.poena, poena, `IV kolo: '${ime}' poeni (naš ${r.poena}, savez ${poena})`);
  const ocekivano = kljuc(ime) === KAZNA_IV.takmicar ? plasman - KAZNA_IV.kazna : plasman;
  assert.strictEqual(r.plasman, ocekivano, `IV kolo: '${ime}' plasman (naš ${r.plasman}, ocekivano ${ocekivano})`);
  poredjeno++;
}

const ekipno4 = rezultatiEkipno(kolo4);
for (const [ime, poena, plasman] of REF_IV_EKIPNO) {
  const e = ekipno4.get(kljuc(kanonKlub(ime)));
  assert.ok(e, `IV kolo ekipno: '${ime}' nije nađen`);
  assert.strictEqual(e.poena, poena, `IV kolo ekipno: '${ime}' poeni (naš ${e.poena}, savez ${poena})`);
  const ocekivano = kljuc(ime) === KAZNA_IV.klub ? plasman - KAZNA_IV.kazna : plasman;
  assert.strictEqual(e.plasman, ocekivano, `IV kolo ekipno: '${ime}' plasman (naš ${e.plasman}, ocekivano ${ocekivano})`);
  poredjeno++;
}
console.log(`  OK  IV kolo protiv dokumenta saveza: ${poredjeno} redova (uz znanu kaznu od ${KAZNA_IV.kazna})`);

/* --- 3. aliasi i ključevi --- */

assert.strictEqual(kanonKlub('SRK LIPLJEN'), kanonKlub('SRK LIPLJEN - PLJEVLJA'));
assert.strictEqual(kanonKlub('SRK VODENE LISICE'), kanonKlub('SRK VODENE LISICE - PG'));
assert.strictEqual(kanonKlub('SRK VODENE LISICE'), kanonKlub('SRK VODENE LISICE - PODGORICA'));
assert.strictEqual(kanonKlub('SRK PLAVSKO JEZERO'), kanonKlub('SRK PLAVSKO JEZERO - PLAV'));
assert.notStrictEqual(kanonKlub('SRK LIM - BERANE'), kanonKlub('SRK LIM AN'));
assert.strictEqual(kljuc('MILAN FUŠTIĆ'), kljuc('Milan Fuštić'));
// broj klubova ne smije da naraste kad savez promijeni način pisanja
assert.strictEqual(sez.snapshotsEk[zadnji].length, 9, 'devet klubova poslije spajanja aliasa');
console.log('  OK  aliasi klubova i ključ imena');

/* --- 4. statistika po takmičaru i po klubu --- */

const st = statistika(sez);
assert.strictEqual(st.ukupno.kola, 4);
assert.strictEqual(st.ukupno.sesija, 12);
assert.strictEqual(st.perTakmicar.length, sez.snapshots[zadnji].length, 'statistika pokriva sve takmičare');
assert.ok(st.ukupno.riba > 0 && st.ukupno.najduza > 0);
assert.strictEqual(st.perKlub.length, 9);

// zbirovi moraju da se slažu iz oba ugla
assert.strictEqual(
  st.perKlub.reduce((a, k) => a + k.riba, 0),
  st.perTakmicar.reduce((a, t) => a + t.riba, 0),
  'riba po klubovima mora biti isto kao riba po takmičarima');
assert.strictEqual(
  st.perKlub.reduce((a, k) => a + k.nule, 0),
  st.perTakmicar.reduce((a, t) => a + t.nule, 0),
  'sesije bez ribe se moraju slagati iz oba ugla');
assert.strictEqual(
  Math.max(...st.perKlub.map(k => k.najduza)),
  Math.max(...st.perTakmicar.map(t => t.najduza)),
  'najduža riba se mora slagati iz oba ugla');

// ekipni zbirovi moraju biti isti kao u ekipnoj rang listi
const ekTabela = new Map(sez.snapshotsEk[zadnji].map(v => [v.kljuc, v]));
for (const k of st.perKlub) {
  const e = ekTabela.get(k.kljuc);
  assert.ok(e, `klub '${k.ime}' nije u ekipnoj rang listi`);
  assert.strictEqual(k.poena, e.poena, `${k.ime}: poeni se ne slažu sa ekipnom tabelom`);
  assert.strictEqual(k.plasman, e.plasman, `${k.ime}: plasman se ne slaže sa ekipnom tabelom`);
  assert.strictEqual(k.kolaOdigrao, 4, `${k.ime}: svi klubovi su nastupili u sva 4 kola`);
  assert.ok(k.najbolje >= 1 && k.najgore >= k.najbolje && k.najgore <= 9, `${k.ime}: raspon plasmana u kolu`);
  assert.ok(k.takmicara >= 3, `${k.ime}: klub ima najmanje tri takmičara`);
  assert.ok(k.nule <= k.sesija, `${k.ime}: sesija bez ribe ne može biti više od odigranih`);
}
for (const t of st.perTakmicar) {
  assert.ok(t.nule <= t.sesija, `${t.ime}: sesija bez ribe ne može biti više od odigranih sesija`);
  assert.ok(t.najbolje >= 1 && t.najgore >= t.najbolje, `${t.ime}: raspon plasmana`);
}
assert.ok(st.rekordi.najSesija && st.rekordi.najRiba, 'rekordi sezone su izračunati');
console.log(`  OK  statistika: ${st.perTakmicar.length} takmičara, ${st.perKlub.length} klubova, ${st.ukupno.riba} riba`);

/* --- 5. grafikon --- */

S.sez = sez;
const mape = sez.snapshots.map(sn => new Map(sn.map((v, i) => [v.kljuc, i + 1])));
const svi = sez.snapshots[zadnji].map(v => ({ kljuc: v.kljuc, ime: v.ime, v: mape.map(m => m.get(v.kljuc) ?? null) }));
for (const w of [null, 380]) {
  if (w === null) delete global.window; else global.window = { innerWidth: w };
  for (const slucaj of [svi.slice(0, 3), svi.slice(0, 1), [{ ime: 'x', v: [1, 1, 1, 1] }], []]) {
    const svg = crtajLinije(slucaj, svi.length);
    assert.ok(!/NaN|Infinity|undefined/.test(svg), `grafikon (${w || 'desktop'}) ima nevalidnu koordinatu`);
    const vb = svg.match(/viewBox="0 0 (\d+) (\d+)"/);
    assert.ok(vb, 'grafikon nema viewBox');
    const W = +vb[1], H = +vb[2];
    for (const m of svg.matchAll(/<(?:text|circle) [^>]*?(?:x|cx)="([\d.]+)"[^>]*?(?:y|cy)="([\d.]+)"/g)) {
      assert.ok(+m[1] >= 0 && +m[1] <= W, `grafikon (${w || 'desktop'}): x=${m[1]} van okvira ${W}`);
      assert.ok(+m[2] >= 0 && +m[2] <= H, `grafikon (${w || 'desktop'}): y=${m[2]} van okvira ${H}`);
    }
  }
}
delete global.window;
console.log('  OK  grafikon: validne koordinate, desktop i telefon');

/* --- 6. zajednička tabela --- */

assert.strictEqual(kratkoIme('Vlatko Peković'), 'V. Peković');
assert.strictEqual(kratkoIme('Bogdan Marković'), 'B. Marković');
assert.strictEqual(kratkoIme('Pele'), 'Pele', 'jednorječno ime ostaje kakvo jeste');

const redovi = saPromjenom(sez.snapshots);
const imenaIz = html => [...html.matchAll(/<span class="ime-puno">([^<]+)</g)].map(m => m[1]);

S.sort = {};
const podrazumijevana = crtajTabelu('t', 'rang', KOL_RANG, redovi, { pocetni: { key: 'mjesto', dir: 1 } });
assert.deepStrictEqual(imenaIz(podrazumijevana), redovi.map(r => r.ime), 'podrazumijevano ide po mjestu');
assert.ok(podrazumijevana.includes('class="c-pos sortable"'), 'kolona Mj. je zakovana');
assert.ok(podrazumijevana.includes('class="c-name sortable"'), 'kolona imena je zakovana');
assert.ok(podrazumijevana.includes('<span class="ime-kratko">'), 'svaki red nosi i kratku verziju imena');

S.sort.t = { key: 'poena', dir: -1 };
assert.deepStrictEqual(
  imenaIz(crtajTabelu('t', 'rang', KOL_RANG, redovi)),
  redovi.slice().sort((a, b) => b.poena - a.poena).map(r => r.ime),
  'sortiranje po poenima opadajuce');

S.sort.t = { key: 'ime', dir: 1 };
assert.deepStrictEqual(
  imenaIz(crtajTabelu('t', 'rang', KOL_RANG, redovi)),
  redovi.map(r => r.ime).sort((a, b) => a.localeCompare(b, 'sr')),
  'sortiranje po imenu');

// prazne vrijednosti (npr. 'novo' u koloni Promjena) uvijek idu na kraj
for (const dir of [1, -1]) {
  S.sort.t = { key: 'promjena', dir };
  const html = crtajTabelu('t', 'rang', KOL_RANG, redovi);
  const nizNovih = [...html.matchAll(/chg (flat|up|down)">(novo)?/g)].map(m => !!m[2]);
  const prvoNovo = nizNovih.indexOf(true);
  if (prvoNovo >= 0) assert.ok(nizNovih.slice(prvoNovo).every(Boolean), 'prazne vrijednosti idu na kraj');
}

// ista tabela radi i za klubove
S.sort = {};
const klubTabela = crtajTabelu('k', 'statKlub', KOL_KLUB, st.perKlub, { pocetni: { key: 'mjesto', dir: 1 } });
assert.deepStrictEqual(
  imenaIz(klubTabela),
  st.perKlub.slice().sort((a, b) => a.mjesto - b.mjesto).map(k => k.ime),
  'tabela klubova ide po ekipnom mjestu');
assert.ok(klubTabela.includes('class="c-name sortable"'), 'i tabela klubova ima zakovanu kolonu naziva');
S.sort.k = { key: 'riba', dir: -1 };
const poRibi = imenaIz(crtajTabelu('k', 'statKlub', KOL_KLUB, st.perKlub));
assert.deepStrictEqual(poRibi, st.perKlub.slice().sort((a, b) => b.riba - a.riba).map(k => k.ime), 'klubovi po ribi');
S.sort = {};
console.log('  OK  zajednicka tabela: sortiranje, zakovane kolone, kratko ime, klubovi');

console.log(`\nSve provjere prošle (${ok} redova protiv Python izlaza, ${poredjeno} protiv dokumenta saveza).`);
