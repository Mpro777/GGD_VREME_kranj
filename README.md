# GGD_VREME

Samodejni zajem vremenskih podatkov ARSO za postajo **Kranj**.

Projekt uporablja GitHub Actions in ob vsakem zagonu opravi dve nalogi:

1. naredi posnetek zaslona ARSO strani,
2. prebere podatke iz vremenske tabele in posodobi Excelovo datoteko.

## Vir podatkov

Podatki se zajemajo s strani ARSO:

`https://meteo.arso.gov.si/uploads/probase/www/observ/surface/text/sl/observationAms_KRANJ_history.html`

Stran prikazuje meritve za zadnjih 48 ur.

## Kaj projekt shranjuje

### Posnetki zaslona

Posnetki se shranjujejo v mapo:

```text
screenshots/
```

### Excelova tabela

Excelova datoteka se shranjuje v mapo:

```text
tabela_vremena/ARSO_Kranj.xlsx
```

Tabela vsebuje naslednje stolpce:

- Datum
- Ura
- Kraj
- Temperatura [°C]
- Vlažnost [%]
- Padavine [mm]
- Vsota padavin [mm]

### Lokalna baza in dnevnik

Projekt uporablja tudi:

```text
arso_kranj.sqlite
arso_kranj.log
```

Datoteka `arso_kranj.sqlite` preprečuje podvajanje meritev in hrani zgodovino zajetih podatkov.

Datoteka `arso_kranj.log` vsebuje zapis uspešnih zagonov in morebitnih napak.

## Glavne datoteke

```text
arso_screenshot.py
```

Naredi posnetek zaslona ARSO strani.

```text
arso_kranj.py
```

Prebere podatke iz ARSO tabele, shrani nove meritve in izdela Excelovo datoteko.

```text
requirements.txt
```

Vsebuje potrebne Python knjižnice.

```text
.github/workflows/arso.yml
```

Določa samodejni zagon programa v GitHub Actions.

## Samodejni zagon

Workflow se zažene vsak dan ob:

```text
22:55 UTC
```

To pomeni približno:

- 23:55 po slovenskem času pozimi,
- 00:55 po slovenskem času poleti.

GitHub Actions uporablja UTC in ne prilagaja samodejno poletnega ter zimskega časa.

Workflow je mogoče zagnati tudi ročno:

1. odprite zavihek **Actions**,
2. izberite **ARSO Kranj screenshot in Excel**,
3. kliknite **Run workflow**.

## Potek delovanja

Ob vsakem zagonu GitHub Actions:

1. prenese repozitorij,
2. namesti Python,
3. namesti Playwright in Chromium,
4. naredi screenshot,
5. prebere ARSO tabelo,
6. posodobi SQLite bazo,
7. ustvari ali posodobi Excelovo datoteko,
8. shrani spremembe nazaj v repozitorij.

## Preprečevanje dvojnikov

ARSO stran prikazuje zadnjih 48 ur podatkov, zato se večina vrstic pri naslednjem zajemu ponovi.

Program uporablja SQLite bazo in vsako vrstico preveri pred shranjevanjem. Že obstoječe meritve se ne dodajo ponovno.

## Struktura repozitorija

```text
GGD_VREME/
├── .github/
│   └── workflows/
│       └── arso.yml
├── screenshots/
├── tabela_vremena/
│   └── ARSO_Kranj.xlsx
├── arso_kranj.py
├── arso_screenshot.py
├── arso_kranj.sqlite
├── arso_kranj.log
├── requirements.txt
└── README.md
```

## Ročni lokalni zagon

Za lokalni zagon potrebujete Python 3.

Namestitev potrebnih knjižnic:

```bash
pip install -r requirements.txt
playwright install chromium
```

Izdelava screenshota:

```bash
python arso_screenshot.py
```

Posodobitev Excelove tabele:

```bash
python arso_kranj.py
```

## Opombe

- Podatki o vetru za postajo Kranj na uporabljeni ARSO strani niso na voljo.
- Če ARSO stran začasno ni dosegljiva, se napaka zapiše v `arso_kranj.log`.
- GitHubov načrtovani zagon se lahko zaradi obremenitve izvede nekaj minut pozneje.

## Spremljanje druge vremenske postaje

Za spremljanje druge ARSO postaje je treba spremeniti spletni naslov v obeh Python datotekah:

```text
arso_screenshot.py
arso_kranj.py
```

V obeh datotekah poiščite vrstico, ki vsebuje:

```python
URL = "https://meteo.arso.gov.si/uploads/probase/www/observ/surface/text/sl/observationAms_KRANJ_history.html"
```

Nato `KRANJ` zamenjajte z oznako druge postaje.

Primer za Ljubljano Bežigrad:

```python
URL = "https://meteo.arso.gov.si/uploads/probase/www/observ/surface/text/sl/observationAms_LJUBL-ANA_BEZIGRAD_history.html"
```

Pomembno je, da uporabite natančen naslov ARSO strani za želeno postajo. Najlažje ga dobite tako:

1. v spletnem brskalniku odprite ARSO stran želene postaje,
2. preverite, da prikazuje podatke za zadnjih 48 ur,
3. kopirajte celoten spletni naslov,
4. isti naslov prilepite v `arso_screenshot.py` in `arso_kranj.py`,
5. spremembe shranite z **Commit changes**,
6. workflow ročno preizkusite v zavihku **Actions**.

V datoteki `arso_kranj.py` je trenutno ime kraja določeno tudi neposredno v kodi. Poiščite:

```python
place = "Kranj"
```

in ga zamenjajte z novim imenom, na primer:

```python
place = "Ljubljana Bežigrad"
```

Prav tako poiščite preverjanje postaje:

```python
value.strip().casefold() == "kranj"
```

in ga spremenite v ime, kot je zapisano v ARSO tabeli, na primer:

```python
value.strip().casefold() == "ljubljana bežigrad"
```

Po želji lahko preimenujete tudi izhodne datoteke in naslove:

```python
DB_PATH = BASE_DIR / "arso_kranj.sqlite"
XLSX_PATH = XLSX_DIR / "ARSO_Kranj.xlsx"
LOG_PATH = BASE_DIR / "arso_kranj.log"
```

Primer za Ljubljano:

```python
DB_PATH = BASE_DIR / "arso_ljubljana.sqlite"
XLSX_PATH = XLSX_DIR / "ARSO_Ljubljana.xlsx"
LOG_PATH = BASE_DIR / "arso_ljubljana.log"
```

Če spremenite imena izhodnih datotek, jih ustrezno spremenite tudi v `.github/workflows/arso.yml` pri ukazih `git add`.

Primer:

```yaml
git add tabela_vremena/ARSO_Ljubljana.xlsx
git add arso_ljubljana.sqlite
git add arso_ljubljana.log
```

Pred prehodom na drugo postajo je priporočljivo obstoječo SQLite bazo preimenovati ali odstraniti, da se podatki različnih postaj ne pomešajo.
