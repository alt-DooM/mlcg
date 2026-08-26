/* Provjera: rang lista i ekipni plasman koje računa app.js moraju se poklopiti
   sa POPUNJENO fajlovima koje je napravila Python skripta popuni_tabelu.py.
   Pokretanje:  node test.js                                                   */

const fs = require('fs');
const path = require('path');
const assert = require('assert');

global.XLSX = require('./vendor/xlsx.full.min.js');
const { parsirajKolo, sezona, kljuc, kanonKlub, asGrid, statistika, crtajLinije, S,
        crtajTabelu, kratkoIme, KOL_RANG, saPromjenom } = require('./app.js');

const KOLA = ['data/kolo-1.xlsx', 'data/kolo-2.xlsx', 'data/kolo-3.xlsx'];
const REFERENCA = path.join('test-data', 'referenca-III-kolo.xlsx');

const ucitaj = p => XLSX.read(new Uint8Array(fs.readFileSync(p)), { type: 'array' });
const grid = (wb, ime) => asGrid(wb.Sheets[ime]);
const cell = (g, r, c) => (g[r - 1] || [])[c - 1] ?? null;

// {kljuc: {ime, poena, plasman}} iz POPUNJENO sheeta (C ime, H poena, I plasman)
function referenca(g) {
  const out = new Map();
  for (let r = 3; r <= g.length; r++) {
    const ime = cell(g, r, 3);
    if (!ime) continue;
    out.set(kljuc(ime), { ime: String(ime), poena: Number(cell(g, r, 8) || 0), plasman: Number(cell(g, r, 9) || 0) });
  }
  return out;
}

const kola = KOLA.map(p => parsirajKolo(new Uint8Array(fs.readFileSync(p)), p));
const sez = sezona(kola);

assert.strictEqual(sez.kola.length, 3, 'očekivana su 3 kola');
assert.deepStrictEqual(sez.kola.map(k => k.kolo), [1, 2, 3], 'brojevi kola iz TABELA sheeta');
assert.ok(sez.kola.every(k => k.sluzbeno), 'sva kola moraju koristiti službeni POJEDINACNO sheet');

const ref = ucitaj(REFERENCA);
const provjere = [
  ['Rang lista', sez.snapshots[2]],
  ['Ekipni plasman', sez.snapshotsEk[2]],
];

let ok = 0;
for (const [sheet, nas] of provjere) {
  const oc = referenca(grid(ref, sheet));
  assert.strictEqual(nas.length, oc.size, `${sheet}: broj redova (naš ${nas.length}, Python ${oc.size})`);

  // isti redosljed
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
  console.log(`  OK  ${sheet}: ${oc.size} redova identično`);
}

// aliasi klubova: iste varijante moraju dati isti kanonski naziv
assert.strictEqual(kanonKlub('SRK LIPLJEN'), kanonKlub('SRK LIPLJEN - PLJEVLJA'));
assert.strictEqual(kanonKlub('SRK VODENE LISICE'), kanonKlub('SRK VODENE LISICE - PG'));
assert.notStrictEqual(kanonKlub('SRK LIM - BERANE'), kanonKlub('SRK LIM AN'));
assert.strictEqual(kljuc('MILAN FUŠTIĆ'), kljuc('Milan Fuštić'));
console.log('  OK  aliasi klubova i ključ imena');

// statistika sezone
const st = statistika(sez);
assert.strictEqual(st.ukupno.kola, 3);
assert.strictEqual(st.ukupno.sesija, 9);
assert.strictEqual(st.perTakmicar.length, sez.snapshots[2].length, 'statistika pokriva sve takmičare');
assert.ok(st.ukupno.riba > 0, 'ukupan broj riba mora biti veći od nule');
assert.ok(st.ukupno.najduza > 0, 'najduža riba se čita iz GRUPA sheeta');
assert.strictEqual(st.perKlub.length, 9, 'devet klubova poslije spajanja aliasa');
assert.strictEqual(
  st.perKlub.reduce((a, k) => a + k.riba, 0),
  st.perTakmicar.reduce((a, t) => a + t.riba, 0),
  'riba po klubovima mora biti isto kao riba po takmičarima');
for (const t of st.perTakmicar) {
  assert.ok(t.nule <= t.sesija, `${t.ime}: sesija bez ribe ne može biti više od odigranih sesija`);
  assert.ok(t.najbolje >= 1 && t.najgore >= t.najbolje, `${t.ime}: raspon plasmana`);
}
assert.ok(st.rekordi.najSesija && st.rekordi.najRiba, 'rekordi sezone su izračunati');
console.log(`  OK  statistika: ${st.perTakmicar.length} takmičara, ${st.perKlub.length} klubova, ${st.ukupno.riba} riba`);

// grafikon: nikad NaN u koordinatama, ni kad su svi izabrani na istom mjestu
S.sez = sez;
const mape = sez.snapshots.map(sn => new Map(sn.map((v, i) => [v.kljuc, i + 1])));
const svi = sez.snapshots[2].map(v => ({ kljuc: v.kljuc, ime: v.ime, v: mape.map(m => m.get(v.kljuc) ?? null) }));
for (const slucaj of [svi.slice(0, 3), svi.slice(0, 1), [{ ime: 'x', v: [1, 1, 1] }], []]) {
  const svg = crtajLinije(slucaj, svi.length);
  assert.ok(!/NaN|Infinity|undefined/.test(svg), 'grafikon ima nevalidnu koordinatu');
  assert.ok(svg.startsWith('<svg') && svg.endsWith('</svg>'), 'grafikon nije validan SVG');
}
console.log('  OK  grafikon: validne koordinate u svim slucajevima');

// zajednicka tabela: sortiranje, zakovane kolone, kratko ime
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
const poPoenima = imenaIz(crtajTabelu('t', 'rang', KOL_RANG, redovi));
const ocekivano = redovi.slice().sort((a, b) => b.poena - a.poena).map(r => r.ime);
assert.deepStrictEqual(poPoenima, ocekivano, 'sortiranje po poenima opadajuce');

S.sort.t = { key: 'ime', dir: 1 };
const poImenu = imenaIz(crtajTabelu('t', 'rang', KOL_RANG, redovi));
assert.deepStrictEqual(poImenu, redovi.map(r => r.ime).sort((a, b) => a.localeCompare(b, 'sr')), 'sortiranje po imenu');

// prazne vrijednosti (npr. 'novo' u koloni Promjena) uvijek idu na kraj
for (const dir of [1, -1]) {
  S.sort.t = { key: 'promjena', dir };
  const html = crtajTabelu('t', 'rang', KOL_RANG, redovi);
  const nizNovih = [...html.matchAll(/chg (flat|up|down)">(novo)?/g)].map(m => !!m[2]);
  const prvoNovo = nizNovih.indexOf(true);
  if (prvoNovo >= 0) assert.ok(nizNovih.slice(prvoNovo).every(Boolean), 'prazne vrijednosti idu na kraj');
}
S.sort = {};
console.log('  OK  zajednicka tabela: sortiranje, zakovane kolone, kratko ime');

console.log(`\nSve provjere prošle (${ok} poređenih redova).`);
