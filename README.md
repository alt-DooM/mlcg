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
