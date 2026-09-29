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
        crtajTabelu, kratkoIme, KOL_RANG, KOL_KLUB, saPromjenom, rezultatiEkipno, primijeniKazne,
        prognoza, scenarij, mozeDoTitule, rasponKola, UKUPNO_KOLA,
        staTreba, matricaDvoboja, zbiroviPoMjestu, prevodUMjesto, licnaTrka,
        profilTerena, poeniZaDuzinu, koordinateZa, uISO, DANI, savjetiZaTeren } = require('./app.js');

const KOLA = [1, 2, 3, 4, 5].map(i => `data/kolo-${i}.xlsx`);
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

const KAZNE = JSON.parse(fs.readFileSync('data/kazne.json', 'utf8')).kazne;
const kola = KOLA.map(p => parsirajKolo(new Uint8Array(fs.readFileSync(p)), p));
kola.forEach(k => primijeniKazne(k, KAZNE));
const sez = sezona(kola);
const zadnji = sez.snapshots.length - 1;

assert.strictEqual(sez.kola.length, 5, 'očekivano je 5 kola');
assert.deepStrictEqual(sez.kola.map(k => k.kolo), [1, 2, 3, 4, 5], 'brojevi kola iz TABELA sheeta');
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

   Nikola Trebješanin je u IV kolu dobio žuti karton i 5 kaznenih plasman-poena.
   Fajl saveza to ne nosi (xlsx ima 14, dokument 19), pa kazna stoji u
   data/kazne.json. Poslije njene primjene sve se poklapa red po red. */

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
  assert.strictEqual(r.plasman, plasman, `IV kolo: '${ime}' plasman (naš ${r.plasman}, savez ${plasman})`);
  poredjeno++;
}

const ekipno4 = rezultatiEkipno(kolo4);
for (const [ime, poena, plasman] of REF_IV_EKIPNO) {
  const e = ekipno4.get(kljuc(kanonKlub(ime)));
  assert.ok(e, `IV kolo ekipno: '${ime}' nije nađen`);
  assert.strictEqual(e.poena, poena, `IV kolo ekipno: '${ime}' poeni (naš ${e.poena}, savez ${poena})`);
  assert.strictEqual(e.plasman, plasman, `IV kolo ekipno: '${ime}' plasman (naš ${e.plasman}, savez ${plasman})`);
  poredjeno++;
}
// kazna je stvarno primijenjena i vidi se na takmičaru
const kaznjen = kolo4.rezultati.get(kljuc('NIKOLA TREBJEŠANIN'));
assert.strictEqual(kaznjen.kazna, 5, 'kazna iz kazne.json mora biti primijenjena');
assert.strictEqual(kolo4.kazne.length, 1, 'IV kolo ima tačno jednu kaznu');
// kazna se ne smije primijeniti dvaput
const ponovo = parsirajKolo(new Uint8Array(fs.readFileSync(KOLA[3])), 'x');
primijeniKazne(ponovo, KAZNE);
assert.strictEqual(ponovo.rezultati.get(kljuc('NIKOLA TREBJEŠANIN')).plasman, kaznjen.plasman,
  'ponovno parsiranje daje isti rezultat');
// kazna na nepoznato ime ne smije nista da promijeni
const nepostojeci = parsirajKolo(new Uint8Array(fs.readFileSync(KOLA[3])), 'x');
primijeniKazne(nepostojeci, [{ kolo: 4, takmicar: 'NEKO KOGA NEMA', plasman: 99 }]);
assert.strictEqual(nepostojeci.kazne.length, 0, 'kazna na nepoznato ime se preskace');
console.log(`  OK  IV kolo protiv dokumenta saveza: ${poredjeno} redova, kazna primijenjena`);

/* --- 2b. V kolo protiv zvanicnog dokumenta saveza ---

   Marko Bakic je dobio zuti karton i 5 kaznenih plasmana, isto kao Trebjesanin
   u IV kolu; kazna stoji u data/kazne.json.

   ZNANO NESLAGANJE: dokument vodi Gorana Obrenovica sa 14 sektorskih bodova,
   a xlsx sa 15. Xlsx je sam sa sobom saglasan na dva mjesta: sesije mu daju
   7+6+2=15, i njegova sopstvena celija ZBIR SEKTORSKIH PLASMANA kaze 15.
   Razlika ide u korist takmicara, pa nije sudijska kazna. Na ukupne tabele ne
   utice nista (provjereno u oba smjera), pa sajt racuna po xlsx-u kao i uvijek.
   Kad savez razjasni: obrisi RAZLIKA_V ispod. */

const RAZLIKA_V = { takmicar: 'GORAN OBRENOVIC', klub: 'SRFFK MANIRO - KOLASIN', razlika: 1 };

const REF_V_POJEDINACNO = [
  ['RADULE MIĆOVIĆ', 7220, 6], ['MILUN ĐUROVIĆ', 8640, 7], ['ALEKSANDAR GAŠEVIĆ', 8060, 7],
  ['RADOVAN KRKOVIĆ', 8000, 8], ['MEHMED DŽIDIĆ', 3920, 10], ['VLATKO PEKOVIĆ', 7000, 11],
  ['BALŠA FUŠTIĆ', 3540, 11], ['BOJAN GOVEDARICA', 5700, 12], ['BOŠKO VULEVIĆ', 4840, 12],
  ['ARSO JEREMIĆ', 3940, 13], ['BOGDAN MARKOVIĆ', 3100, 13], ['MILAN FUŠTIĆ', 4520, 14],
  ['GORAN OBRENOVIĆ', 3160, 14], ['IVICA RAJKOVIĆ', 4940, 15], ['PETAR OBRADOVIĆ', 4060, 16],
  ['NIKOLA POPOVIĆ', 3100, 16], ['KEMAL AJANOVIĆ', 3560, 17], ['NIKOLA TREBJEŠANIN', 1880, 17],
  ['MARKO BAKIĆ', 3940, 18], ['NIKOLA ĐALOVIĆ', 3100, 19], ['OMAR BAŠIĆ', 2060, 19],
  ['SAVO RAIČEVIĆ', 1460, 20], ['DAMIR CANOVIĆ', 1580, 21], ['BOGDAN GOVEDARICA', 1780, 23],
  ['MUMIN HEKALO', 780, 24], ['MILJAN ANĐIĆ', 520, 27], ['ORHAN BAŠIĆ', 0, 27],
];

const REF_V_EKIPNO = [
  ['SRK LIM - BERANE', 20700, 25], ['SRFFK RAVNJAK - MOJKOVAC', 16120, 32],
  ['SRFFK MANIRO - KOLAŠIN', 14260, 35], ['SRK KOLAŠIN', 11980, 44],
  ['SRK LIPLJEN - PLJEVLJA', 8260, 51], ['SRK VODENE LISICE - PODGORICA', 11100, 53],
  ['SRK GORŠTAK - KOLAŠIN', 9400, 55], ['SRK TARA - MOJKOVAC', 8940, 55],
  ['SRK PLAVSKO JEZERO', 3640, 67],
];

const kolo5 = sez.kola[4];
for (const [ime, poena, plasman] of REF_V_POJEDINACNO) {
  const r = kolo5.rezultati.get(kljuc(ime));
  assert.ok(r, `V kolo: '${ime}' nije nađen u xlsx-u`);
  assert.strictEqual(r.poena, poena, `V kolo: '${ime}' poeni (naš ${r.poena}, savez ${poena})`);
  const ocekivano = kljuc(ime) === RAZLIKA_V.takmicar ? plasman + RAZLIKA_V.razlika : plasman;
  assert.strictEqual(r.plasman, ocekivano, `V kolo: '${ime}' plasman (naš ${r.plasman}, ocekivano ${ocekivano})`);
  poredjeno++;
}
const ekipno5 = rezultatiEkipno(kolo5);
for (const [ime, poena, plasman] of REF_V_EKIPNO) {
  const e = ekipno5.get(kljuc(kanonKlub(ime)));
  assert.ok(e, `V kolo ekipno: '${ime}' nije nađen`);
  assert.strictEqual(e.poena, poena, `V kolo ekipno: '${ime}' poeni (naš ${e.poena}, savez ${poena})`);
  const ocekivano = kljuc(ime) === RAZLIKA_V.klub ? plasman + RAZLIKA_V.razlika : plasman;
  assert.strictEqual(e.plasman, ocekivano, `V kolo ekipno: '${ime}' plasman (naš ${e.plasman}, ocekivano ${ocekivano})`);
  poredjeno++;
}
assert.strictEqual(kolo5.rezultati.get(kljuc('MARKO BAKIĆ')).kazna, 5, 'zuti karton Marka Bakica');
assert.strictEqual(kolo5.kazne.length, 1, 'V kolo ima tacno jednu kaznu');
// sporni bod ne smije da pomjeri nikoga na ukupnoj tabeli
const bezSpora = sezona(KOLA.map(p => {
  const k = parsirajKolo(new Uint8Array(fs.readFileSync(p)), p);
  primijeniKazne(k, KAZNE);
  return k;
}).map((k, i) => {
  if (i === 4) k.rezultati.get(kljuc('GORAN OBRENOVIĆ')).plasman -= RAZLIKA_V.razlika;
  return k;
}));
assert.deepStrictEqual(
  bezSpora.snapshots[4].map(v => v.kljuc), sez.snapshots[4].map(v => v.kljuc),
  'sporni bod ne smije da promijeni redosljed pojedinacne tabele');
assert.deepStrictEqual(
  bezSpora.snapshotsEk[4].map(v => v.kljuc), sez.snapshotsEk[4].map(v => v.kljuc),
  'sporni bod ne smije da promijeni redosljed ekipne tabele');
console.log(`  OK  V kolo protiv dokumenta saveza: ${REF_V_POJEDINACNO.length + REF_V_EKIPNO.length} redova, sporni bod bez uticaja`);

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
assert.strictEqual(st.ukupno.kola, 5);
assert.strictEqual(st.ukupno.sesija, 15);
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
  assert.strictEqual(k.kolaOdigrao, 5, `${k.ime}: svi klubovi su nastupili u svih 5 kola`);
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
  for (const slucaj of [svi.slice(0, 3), svi.slice(0, 1), [{ ime: 'x', v: [1, 1, 1, 1, 1] }], []]) {
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

/* --- 7. prognoza: kalkulator je racun, simulacija je simulacija --- */

const preostalo = UKUPNO_KOLA - sez.kola.length;
assert.strictEqual(preostalo, 1, 'sezona ima 6 kola, odigrano je 5');

const osnova = saPromjenom(sez.snapshots);
const { min, max } = rasponKola(sez, false);
assert.ok(min >= 3 && max <= 27 && min < max, `raspon zbira po kolu (${min}-${max})`);

// kalkulator mora da daje tacno ono sto se rucno izracuna
const prosjeci = new Map(st.perTakmicar.map(v => [v.kljuc, v.plasman / v.kolaOdigrao]));
const vodeci = osnova[0], drugi = osnova[1];
let sc = scenarij(osnova, new Map([[vodeci.kljuc, max], [drugi.kljuc, min]]), preostalo, prosjeci);
let poK = new Map(sc.map(v => [v.kljuc, v]));
assert.strictEqual(poK.get(vodeci.kljuc).konacni, vodeci.plasman + max * preostalo);
assert.strictEqual(poK.get(drugi.kljuc).konacni, drugi.plasman + min * preostalo);
assert.ok(poK.get(drugi.kljuc).konacnoMjesto < poK.get(vodeci.kljuc).konacnoMjesto,
  'ko odigra najbolje mora prestici onoga ko odigra najgore');

// bez ijedne pretpostavke svako zadrzava svoj prosjek
sc = scenarij(osnova, new Map(), preostalo, prosjeci);
assert.strictEqual(sc.length, osnova.length, 'scenarij pokriva cijelu tabelu');
assert.strictEqual(sc[0].konacni, Math.min(...sc.map(v => v.konacni)), 'tabela je sortirana po konacnom zbiru');

// matematicka mogucnost: vodeci uvijek moze, posljednji ne moze
const moze = mozeDoTitule(osnova, preostalo, min, max);
assert.strictEqual(moze.get(osnova[0].kljuc), true, 'vodeci uvijek moze do titule');
assert.strictEqual(moze.get(osnova[osnova.length - 1].kljuc), false, 'posljednji ne moze');

// simulacija: iste brojke pri svakom pokretanju, i sve sanse se sabiraju u 1
const p1 = prognoza(sez, preostalo, { broj: 3000 });
const p2 = prognoza(sez, preostalo, { broj: 3000 });
assert.deepStrictEqual([...p1.titula], [...p2.titula], 'simulacija mora biti ponovljiva');
const suma = m => [...m.values()].reduce((a, v) => a + (typeof v === 'number' ? v : v.pobjeda), 0);
for (const [naziv, m] of [['titula', p1.titula], ['ekipna titula', p1.ekipnaTitula],
                          ['pobjeda u kolu', p1.sljedece], ['ekipna pobjeda', p1.ekipnoSljedece]]) {
  assert.ok(Math.abs(suma(m) - 1) < 0.02, `${naziv}: sanse se moraju sabrati u 100% (dobijeno ${suma(m).toFixed(3)})`);
  for (const v of m.values()) {
    const x = typeof v === 'number' ? v : v.pobjeda;
    assert.ok(x >= 0 && x <= 1, `${naziv}: vjerovatnoca van 0..1`);
  }
}
// vodeci na tabeli mora imati najvecu sansu za titulu
assert.strictEqual([...p1.titula.entries()].sort((a, b) => b[1] - a[1])[0][0], osnova[0].kljuc,
  'vodeci mora imati najvecu sansu za titulu');
// ko matematicki ne moze do titule, ne smije imati nijednu simuliranu titulu
for (const [k, v] of p1.titula) if (moze.get(k) === false) assert.strictEqual(v, 0, 'nemoguce a simulirano');
// bez preostalih kola nista ne puca
const p0 = prognoza(sez, 0, { broj: 200 });
assert.ok(p0.sljedece.size > 0 && [...p0.titula.values()].every(v => v === 0));
// "sta kome treba": granica mora stvarno da bude granica
const uMjesto = prevodUMjesto(sez, false);
const treba = staTreba(osnova, prosjeci, preostalo, min, max);
const zivi = osnova.filter(v => treba.get(v.kljuc).moguce);
assert.ok(zivi.length >= 1 && zivi.length < osnova.length, 'neko moze do titule, ali ne svi');
assert.strictEqual(zivi[0].kljuc, osnova[0].kljuc, 'vodeci je uvijek medju zivima');
for (const v of zivi) {
  const g = treba.get(v.kljuc).granica;
  const taman = scenarij(osnova, new Map([[v.kljuc, g]]), preostalo, prosjeci);
  assert.strictEqual(taman[0].kljuc, v.kljuc, `${v.ime}: sa granicom ${g} mora biti prvak`);
  if (g < max) {
    const malo_gore = scenarij(osnova, new Map([[v.kljuc, g + 1]]), preostalo, prosjeci);
    assert.notStrictEqual(malo_gore[0].kljuc, v.kljuc, `${v.ime}: jedan slabiji od granice vise nije prvak`);
  }
}
for (const v of osnova) if (!treba.get(v.kljuc).moguce) {
  const najbolje = scenarij(osnova, new Map([[v.kljuc, min]]), preostalo, prosjeci);
  assert.notStrictEqual(najbolje[0].kljuc, v.kljuc, `${v.ime}: oznacen kao otpisan, a najbolji ishod ga vodi do titule`);
}

// ose matrice ne smiju dvaput da ponove isto mjesto
const parovi = zbiroviPoMjestu(uMjesto, min, max, 10);
assert.strictEqual(new Set(parovi.map(p => p[0])).size, parovi.length, 'svako mjesto se pojavljuje jednom');
assert.deepStrictEqual(parovi.map(p => p[0]), parovi.map(p => p[0]).slice().sort((a, b) => a - b), 'ose su rastuce');

// matrica: najbolji ishod za jednog i najgori za drugog mora dati prvog
const mx = matricaDvoboja(osnova, prosjeci, preostalo, osnova[0].kljuc, osnova[1].kljuc, parovi);
assert.strictEqual(mx.length, parovi.length);
assert.strictEqual(mx[0][mx[0].length - 1], osnova[0].kljuc, 'gornji desni ugao pripada prvom');
assert.strictEqual(mx[mx.length - 1][0], osnova[1].kljuc, 'donji lijevi ugao pripada drugom');
// prosjek po kolu ne smije da nosi kazne za propustena kola: sa kaznama bi
// Sokolovicu ispalo 113 po kolu, a najslabije moguce odigrano kolo je 27
for (const t of st.perTakmicar) {
  const pk = t.plasman / t.kolaOdigrao;
  assert.ok(pk >= min && pk <= max,
    `${t.ime}: prosjek po kolu ${pk.toFixed(1)} je van moguceg opsega ${min}-${max}`);
}
const saKaznama = osnova.find(v => v.ime === 'Edin Sokolović');
const cist = st.perTakmicar.find(v => v.kljuc === saKaznama.kljuc);
assert.ok(saKaznama.plasman > cist.plasman, 'kumulativni zbir nosi kazne, cisti ne');
assert.strictEqual(saKaznama.plasman - cist.plasman, 30 * (sez.kola.length - cist.kolaOdigrao),
  'razlika mora biti tacno 30 po propustenom kolu');

// licna trka: vrijedi za svakoga, ne samo za vrh tabele
for (const v of [osnova[0], osnova[Math.floor(osnova.length / 2)], osnova[osnova.length - 1]]) {
  const t = licnaTrka(osnova, prosjeci, preostalo, min, max, v.kljuc);
  assert.strictEqual(t.sada.kljuc, v.kljuc);
  assert.ok(t.najbolje <= t.ocekivano.konacnoMjesto, `${v.ime}: najbolji ishod ne smije biti losiji od ocekivanog`);
  assert.ok(t.najgore >= t.ocekivano.konacnoMjesto, `${v.ime}: najgori ishod ne smije biti bolji od ocekivanog`);
  assert.ok(t.komsije.length > 0 && t.komsije.every(k => k.kljuc !== v.kljuc), `${v.ime}: komsije ne ukljucuju njega samog`);
  // granica protiv komsije mora stvarno da bude granica
  for (const k of t.komsije.filter(x => x.granica !== null)) {
    const sc = scenarij(osnova, new Map([[v.kljuc, k.granica]]), preostalo, prosjeci);
    const ja = sc.find(x => x.kljuc === v.kljuc), on = sc.find(x => x.kljuc === k.kljuc);
    assert.ok(ja.konacnoMjesto <= on.konacnoMjesto, `${v.ime} protiv ${k.ime}: granica ${k.granica} ga ne drzi ispred`);
  }
}
// vodeceg niko ne moze prestici odozdo ako odigra najbolje
const vrh = licnaTrka(osnova, prosjeci, preostalo, min, max, osnova[0].kljuc);
assert.strictEqual(vrh.najbolje, 1, 'vodeci sa najboljim ishodom zavrsava prvi');
assert.strictEqual(vrh.komsije.filter(k => k.mjesto < vrh.sada.mjesto).length, 0, 'ispred vodeceg nema nikoga');

/* --- 8. karton rijeke --- */

// bodovanje: provjereno na svakoj pojedinacnoj ribi, bez izuzetka
let ribaUkupno = 0;
for (const kolo of sez.kola) {
  for (const r of kolo.rezultati.values()) {
    for (let s = 0; s < 3; s++) {
      const d = r.duzine[s] || [];
      assert.strictEqual(d.length, r.riba[s], `${r.ime}, kolo ${kolo.kolo}, sesija ${s + 1}: broj procitanih riba`);
      const bod = d.reduce((a, cm) => a + poeniZaDuzinu(cm), 0);
      assert.strictEqual(bod, r.sesije[s] ? r.sesije[s].poena : 0,
        `${r.ime}, kolo ${kolo.kolo}, sesija ${s + 1}: poeni iz duzina se ne slazu sa zbirom saveza`);
      if (d.length) assert.strictEqual(Math.max(...d), r.najduza[s], `${r.ime}: najduza riba`);
      ribaUkupno += d.length;
    }
  }
}
assert.strictEqual(ribaUkupno, 864, 'ocekivano je 864 ribe u pet kola');
assert.strictEqual(poeniZaDuzinu(20), 500);
assert.strictEqual(poeniZaDuzinu(48), 1060);
assert.ok(2 * poeniZaDuzinu(20) > poeniZaDuzinu(40), 'dvije male ribe nose vise od jedne velike');

const tereni = profilTerena(sez);
assert.strictEqual(tereni.length, 4, 'pet kola na cetiri terena');
assert.strictEqual(tereni.reduce((a, t) => a + t.ukupnoRiba, 0), ribaUkupno, 'sve ribe su rasporedjene po terenima');
assert.strictEqual(tereni.reduce((a, t) => a + t.brojKola, 0), sez.kola.length);
for (const t of tereni) {
  assert.strictEqual(t.poSesiji.reduce((a, b) => a + b, 0), t.ukupnoRiba, `${t.mjesto}: zbir po sesijama`);
  assert.ok(Math.abs(t.raspodjela.reduce((a, r) => a + r.udio, 0) - 1) < 1e-9, `${t.mjesto}: raspodjela se sabira u 1`);
  assert.ok(t.najveca >= t.prosjek && t.prosjek >= 20, `${t.mjesto}: prosjek i maksimum`);
  assert.ok(t.praznihUdio >= 0 && t.praznihUdio <= 1, `${t.mjesto}: udio praznih sesija`);
  assert.ok(koordinateZa(t.mjesto), `${t.mjesto}: nema koordinata za vrijeme`);
}
// Lim je rijeka sitne ribe, Tara krupnije; ako se to okrene, profil je pukao
const lim = tereni.find(t => /Lim/i.test(t.mjesto));
const tara = tereni.find(t => /Tara, Mojkovac/i.test(t.mjesto));
assert.ok(lim.prosjek < tara.prosjek, 'Lim ima sitniju ribu od Tare');
assert.strictEqual(uISO('27. 09. 2026'), '2026-09-27');
assert.strictEqual(uISO('nema datuma'), null);
// sljedece kolo: konfiguracija mora da se slaze sa onim sto sajt ocekuje
const SK = JSON.parse(fs.readFileSync('data/sljedece-kolo.json', 'utf8'));
assert.strictEqual(SK.kolo, sez.kola.length + 1, 'sljedece kolo je naredno po redu');
assert.ok(!sez.kola.some(k => k.kolo === SK.kolo), 'sljedece kolo jos nije odigrano');
const isoSK = uISO(SK.datum);
assert.ok(isoSK, 'datum sljedeceg kola se moze pretvoriti u ISO');
assert.ok(koordinateZa(SK.mjesto), `${SK.mjesto}: nema koordinata, prognoza se ne bi mogla povuci`);
assert.ok(tereni.find(t => t.mjesto === SK.referenca), 'referentni teren postoji u podacima');
assert.strictEqual(DANI[new Date(isoSK + 'T00:00:00').getDay()], 'nedjelja', '11. oktobar 2026. je nedjelja');
assert.ok(new Date(isoSK) > new Date(uISO(sez.kola[sez.kola.length - 1].datum)), 'sljedece kolo je poslije posljednjeg odigranog');
console.log(`  OK  sljedece kolo: ${SK.mjesto}, ${SK.datum}, referenca ${SK.referenca}`);

// savjeti: izvedeni iz brojki, pa moraju da prate brojke
for (const t of tereni) {
  const sv = savjetiZaTeren(t, tereni);
  assert.ok(sv.length >= 3 && sv.length <= 5, `${t.mjesto}: ocekivano 3 do 5 savjeta, dobijeno ${sv.length}`);
  for (const x of sv) {
    assert.ok(x.ik && x.broj !== undefined && x.tekst, `${t.mjesto}: savjet bez ikone, broja ili teksta`);
    assert.ok(!/undefined|NaN/.test(String(x.broj) + x.tekst), `${t.mjesto}: savjet sadrzi NaN ili undefined`);
    // svaki savjet mora da zna u koju sekciju ide, inace se nigdje ne iscrta
    assert.ok(['duzine', 'sesije', 'bodovi'].includes(x.gdje), `${t.mjesto}: savjet '${x.tekst}' nema sekciju`);
  }
  // savjet o padu kroz dan mora da se slaze sa stvarnim brojem riba
  const pad = sv.find(x => /prvoj|trećoj|ujednačen/.test(x.tekst));
  assert.ok(pad, `${t.mjesto}: nema savjeta o kretanju ulova kroz dan`);
  const [a, , c] = t.poSesiji;
  if (/prvoj nego u trećoj/.test(pad.tekst)) assert.ok(a > c, `${t.mjesto}: savjet kaze pad, a brojevi ne`);
  if (/trećoj nego u prvoj/.test(pad.tekst)) assert.ok(c > a, `${t.mjesto}: savjet kaze rast, a brojevi ne`);
}
// Lim je najizrazitiji pad kroz dan i najsitnija riba
const savjetiLim = savjetiZaTeren(lim, tereni);
assert.ok(savjetiLim.some(x => /sitnija/.test(x.tekst)), 'Lim mora nositi savjet o sitnijoj ribi');
assert.ok(savjetiLim.some(x => /kapitalac/.test(x.tekst)), 'Lim mora nositi savjet o izostanku kapitalca');
console.log(`  OK  savjeti: ${tereni.length} terena, svaki sa svojim brojkama`);

console.log(`  OK  karton rijeke: ${ribaUkupno} riba, ${tereni.length} terena, bodovanje 20/cm + 100`);

console.log(`  OK  prognoza: kalkulator tacan, granice tacne, simulacija ponovljiva (${preostalo} preostala kola, ${zivi.length} jos u igri)`);

console.log(`\nSve provjere prošle (${ok} redova protiv Python izlaza, ${poredjeno} protiv dokumenta saveza).`);
