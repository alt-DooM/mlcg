# Mušičarska liga Crne Gore, sajt sa rezultatima

Statički sajt koji čita zvanične Excel fajlove saveza direktno u browseru i iz njih
računa rang listu, ekipni plasman, pregled po kolima i statistiku sezone.
Nema servera, nema baze, nema troškova. Sve radi na GitHub Pages.

## Kako se dodaje novo kolo

1. Preimenuj Excel fajl kola u `kolo-4.xlsx` (broj = redni broj kola).
2. Otvori folder `data/` u ovom repozitorijumu na GitHub-u.
3. Klikni **Add file > Upload files**, prevuci fajl, pa **Commit changes**.
4. Za minut-dva sajt sam prikazuje novo kolo. Ništa drugo se ne mijenja.

Prije objave se rezultat može provjeriti. Na sajtu, dugme **Dodaj kolo** u dnu strane
otvara polje u koje se fajl prevuče. Rang lista se odmah preračuna, ali samo u tom
browseru. Ništa se ne šalje i ništa se ne objavljuje dok se fajl ne ubaci u `data/`.

## Objavljivanje (jednom)

1. Napravi repozitorijum na GitHub-u i ubaci sadržaj ovog foldera u njegov korijen.
2. **Settings > Pages > Source: Deploy from a branch**, grana `main`, folder `/ (root)`.
3. Sajt je na `https://alt-doom.github.io/<naziv-repozitorijuma>/`.

Ako repozitorijum ne nazoveš `mlcg`, izmijeni `REPO_DATA_URL`
na vrhu `app.js`. To je link koji vodi na `data/` folder iz modala "Dodaj kolo".

## Šta se čita iz Excel fajla

Novi način pisanja naziva kluba (savez ga mijenja iz kola u kolo) mora se dodati
u `KLUB_ALIASI` u `app.js`, inače se isti klub pojavi dvaput u ekipnoj tabeli.
U IV kolu se tako pojavilo `SRK PLAVSKO JEZERO - PLAV` umjesto `SRK PLAVSKO JEZERO`.

| Sheet | Šta se uzima |
|---|---|
| `TABELA` | broj kola, mjesto, datum, žrijeb (koji takmičar je u kojoj grupi) i klubovi |
| `POJEDINACNO` | službeni poeni i sektorski plasmani po sesiji, uključuju i sudijske kazne |
| `GRUPA I/II/III` | broj riba i najduža riba po sesiji |

Ako fajl nema `POJEDINACNO`, plasmani se računaju iz poena (0 poena = zadnje mjesto
u grupi), isto kao rezervni put u `popuni_tabelu.py`. Sajt u tom slučaju to i napiše.

## Sudijske kazne

Fajl saveza ne nosi sudijske kazne (žuti karton i slično), a zvanični dokument
koji savez objavi poslije kola ih ima. Zato kazne stoje u `data/kazne.json`:

```json
{ "kolo": 4, "takmicar": "NIKOLA TREBJEŠANIN", "plasman": 5, "razlog": "Žuti karton" }
```

`plasman` je broj sektorskih plasman-poena koji se **dodaje** takmičaru u tom kolu.
Ekipni plasman kluba se sam popravi, jer se računa iz rezultata njegovih takmičara.

Kazna namjerno nije upisana u sam xlsx: ako neko kasnije ponovo uploaduje
originalni fajl saveza, kazna bi nestala bez traga. Ovako preživi.

Do sada viđene kazne: IV kolo Nikola Trebješanin +5 (žuti karton), V kolo
Marko Bakić +5 (žuti karton). Žuti karton je dosad uvijek nosio 5 plasman-poena,
ali se vrijednost čita iz dokumenta saveza, ne pretpostavlja.

Ime se uparuje bez kvačica i viška razmaka. Ako se ne pronađe u tom kolu, sajt
ispiše upozorenje u konzoli browsera i ne mijenja ništa. Takmičar sa kaznom
dobija zvjezdicu uz ime, a ispod tabela tog kola piše razlog.

## Prognoza

Tab **Prognoza** ima dvije odvojene stvari, namjerno razdvojene jer nisu iste vrste:

**Kalkulator titule** je običan račun. Povuku se klizači (pretpostavljeni zbir
sektorskih plasmana po kolu) i vidi se ko je prvak pod tom pretpostavkom. Nikad
ne griješi jer ništa ne pogađa.

**Šanse** su Monte Carlo simulacija: svakom takmičaru se 10.000 puta izvlače
sesije nalik onima koje je stvarno lovio, pa se prebroji koliko puta je bio prvi.
Sjeme slučajnih brojeva je fiksno (`SIMULACIJA.sjeme`), pa se procenti ne mijenjaju
sami od sebe pri osvježavanju stranice.

Sa malo odigranih kola je istorija preuska i model postaje pretjerano siguran:
bez ublažavanja je poslije I kola davao favoritu 72%, a stvarnom pobjedniku 0,0%.
Zato se uz stvarnu istoriju miješa i `SIMULACIJA.ublazavanje` nasumičnih sesija.
Provjereno unazad na odigranim kolima, greška za pogađanje prva tri padne sa
0,100 na 0,088 (nasumično pogađanje je 0,180). Favorit modela je dobio kolo
1 od 3 puta. Te brojke stoje i na samoj stranici, da niko ne misli da je ovo
proročanstvo.

Broj kola u sezoni je `UKUPNO_KOLA` u `app.js` (trenutno 6). Iz njega se računa
koliko je kola preostalo, a na stranici se može i ručno promijeniti.

## Karton rijeke

Tab **Rijeke** gradi profil svakog terena iz pojedinačnih riba. U `GRUPA`
sheetovima je zapisana svaka ulovljena riba sa dužinom u cm, njih 864 kroz pet
kola. Iz toga ide raspodjela dužina, ulov po sesijama, koliko riba dobija sesiju
i koliko sesija završi bez ribe.

Bodovanje je linearno i provjereno na svih 864 ribe bez ijednog izuzetka:

```
poeni = 20 x dužina(cm) + 100
```

Onih ravnih 100 po ribi znači da broj riba uvijek tuče veličinu: dvije od 20 cm
nose 1000, jedna od 40 cm nosi 900. Test provjerava i da se poeni izračunati iz
dužina poklapaju sa zbirovima saveza za svaku sesiju svakog takmičara.

Vrijeme na dan kola se povlači sa Open-Meteo, bez registracije i bez ključa, i
kešira se u browseru. Koordinate terena su u `KOORDINATE` u `app.js`; novi teren
treba dodati tamo. Temperatura vode i vodostaj se ne mogu besplatno povući ni
sa jednog javnog servisa za crnogorske rijeke, pa ih nema. Ako vrijeme ne stigne,
sekcija se jednostavno ne prikaže.

## Pravila obračuna

Ista kao u `popuni_tabelu.py`:

- Bolji je **manji zbir sektorskih plasmana**. Kod istog zbira odlučuje više poena.
- Propušteno kolo nosi **30 kaznenih plasman-poena** (konstanta `KAZNA_PRESKOK`).
- Klubovi ne dobijaju kaznu za propuštena kola, ali prazan slot u ekipi nosi
  `3 x broj ekipa` plasman-poena.
- Isti klub savez piše različito iz kola u kolo (`SRK LIPLJEN` ili `SRK LIPLJEN - PLJEVLJA`).
  Varijante se svode na jedan naziv preko `KLUB_ALIASI` u `app.js`. Nova varijanta se
  samo doda u tu listu.
- Imena se uparuju bez kvačica i viška razmaka, pa je `MILAN FUŠTIĆ` isti čovjek kao
  `Milan Fuštić`.

## Provjera

```
node test.js
```

Poredi rang listu i ekipni plasman koje računa `app.js` sa fajlom
`test-data/referenca-III-kolo.xlsx`, koji je napravila Python skripta. Svi redovi
moraju biti identični: poeni, plasmani i redosljed.

## Fajlovi

```
index.html                    struktura strane
style.css                     stil
app.js                        parsiranje Excela, obračun, prikaz
vendor/xlsx.full.min.js       SheetJS (lokalno, bez CDN-a)
data/kolo-N.xlsx              fajlovi kola, ovdje se dodaje novo kolo
test-data/                    referentni fajl za test
test.js                       provjera obračuna
```
