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
        staTreba, matricaDvoboja, zbiroviPoMjestu, prevodUMjesto, licnaTrka } = require('./app.js');

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

const KAZNE = JSON.parse(fs.readFileSync('data/kazne.json', 'utf8')).kazne;
const kola = KOLA.map(p => parsirajKolo(new Uint8Array(fs.readFileSync(p)), p));
kola.forEach(k => primijeniKazne(k, KAZNE));
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

/* --- 7. prognoza: kalkulator je racun, simulacija je simulacija --- */

const preostalo = UKUPNO_KOLA - sez.kola.length;
assert.strictEqual(preostalo, 2, 'sezona ima 6 kola, odigrana su 4');

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

console.log(`  OK  prognoza: kalkulator tacan, granice tacne, simulacija ponovljiva (${preostalo} preostala kola, ${zivi.length} jos u igri)`);

console.log(`\nSve provjere prošle (${ok} redova protiv Python izlaza, ${poredjeno} protiv dokumenta saveza).`);
