# IzjavaBirača 2026

Čitanje lične karte preko Čelik API-ja i štampa izjave birača (obrazac **NPRS-3**).
Nastavak programa `Izbori2024izjava`, uz tri tražene izmene: novi `.doc` šablon,
CelikApi 1.4.2 i podesiv pomak (offset) za štampu.

---

## Šta je novo u odnosu na verziju 2024

| | 2024 | 2026 |
|---|---|---|
| Šablon | `IzjavaBiracaTemplate.dot` (bookmarkovi) | `NPRS-3 ....doc` (ćelije tabele) |
| CelikApi.dll | **1.4.0.0** | **1.4.2.0** (x86 i x64) |
| Pomak štampe | fiksna pozicija u šablonu | **podesiv X/Y u mm + probna štampa** |
| Naziv liste, datum, opština | hardkodirano (build po opštini) | **Podešavanja** (`Podesavanja.xml`) |
| Evidencija potpisnika | Excel `.xlsx` (traži instaliran Excel) | **CSV** (ne traži Excel) |
| Ako nema Word-a | nerazumljiva COM greška | jasna poruka na srpskom |
| Provera šablona | SHA512 hash | provera strukture |
| Start programa | gasio sve WINWORD i EXCEL procese | uredno otpuštanje COM objekata |

### O verziji Čelik API-ja

Stara verzija je bila **1.4.0.0**, nova je **1.4.2.0**. Poređenjem hedera utvrđeno je
da je jedina razlika **dve nove konstante**:

```c
const int EID_Cert_SIG_FIXED    = 4;
const int EID_Cert_SIG_VARIABLE = 5;
```

Svi prototipovi funkcija i sve strukture su **identični**, pa `EidStartup(4)` i dalje
važi i postojeći P/Invoke sloj radi bez izmena.

---

## Koju verziju instalirati: x86 ili x64?

> **Bira se prema bitnosti Microsoft Office-a, a NE prema bitnosti Windows-a.**

Program koristi Word preko COM interopa, koji se učitava **u isti proces**. Zato
32-bitni Word ne može da radi sa 64-bitnim programom i obrnuto.

| Office | Verzija programa |
|---|---|
| 32-bitni (najčešći slučaj) | `IzjavaBiraca2026_x86` |
| 64-bitni | `IzjavaBiraca2026_x64` |

Bitnost Office-a: **Word → File → Account → About Word** (piše na kraju naslova),
ili u registru `HKLM\SOFTWARE\Microsoft\Office\<verzija>\Outlook\Bitness`.

Uz svaku verziju ide odgovarajući `CelikApi.dll` (32-bitni uz x86, 64-bitni uz x64) —
build ga kopira automatski iz `lib\x86\` odnosno `lib\x64\`.

---

## Struktura projekta

```
Izbori2026izjava\
├─ Form1.vb / .Designer.vb      glavni ekran (izgled kao 2024 + dugme Podešavanja)
├─ FrmPodesavanja.vb / .Designer.vb   ekran sa podešavanjima i offsetom
├─ WordStampac.vb               popunjavanje i štampa NPRS-3 šablona
├─ PodaciBiraca.vb              podaci sa kartice spremni za štampu
├─ Podesavanja.vb               čitanje/pisanje Podesavanja.xml
├─ CsvEvidencija.vb             spisak potpisnika (CSV) + provera duplikata
├─ OfficeProvera.vb             provera da li je Word instaliran
├─ CelikApi.vb                  P/Invoke sloj (iz 2024, + konstante 1.4.2)
├─ SmartcardManager.vb          rad sa čitačem (nepromenjeno)
├─ Transliteration.vb           latinica ↔ ćirilica (nepromenjeno)
├─ lib\x86\CelikApi.dll         Čelik API 1.4.2, 32-bit
├─ lib\x64\CelikApi.dll         Čelik API 1.4.2, 64-bit
├─ doc\                         PDF dokumentacija API-ja i originalni ZIP-ovi
├─ setup\                     Inno Setup skripta za instalacione fajlove
└─ build.cmd                    build za obe platforme + portable ZIP
```

---

## Kako se popunjava šablon

Novi NPRS-3 `.doc` **nema nijedan bookmark, form field ni content control** (provereno),
za razliku od starog `.dot`-a koji je koristio bookmarkove `ime_prezime`, `jmbg_1..13` itd.
Zato se piše direktno u **Tabelu 3**, koja je blok „Б И Р А Ч":

```
Red 1  Б И Р А Ч
Red 2  [ИМЕ И ПРЕЗИМЕ]        ← upisuje se u 1. paragraf ćelije
Red 3  [0][1][0][1]...[3]     ← 13 kutijica, po jedna cifra
Red 4  (ЈМБГ)
Red 5  [МЕСТО, Улица 15/3]    ← upisuje se u 1. paragraf ćelije
Red 6  (место и адреса пребивалишта)
Red 7  ______________________ (za potpis)
Red 8  (потпис)
```

Šablon se otvara **read-only** i zatvara sa `SaveChanges:=False`, pa originalni fajl
na disku ostaje netaknut.

---

## Offset za štampu

U **Podešavanja → Štampa i pomak** unose se pomak levo/desno i gore/dole u milimetrima
(pozitivno = udesno i nadole). Dugme **Probna štampa** štampa test primerak sa izmišljenim
podacima da se pomak iskalibriše bez trošenja pravog obrasca.

### Zašto vertikalni pomak ima ograničenje

Testiranjem je utvrđeno da **NPRS-3 obrazac staje na tačno jednu stranu bez ikakve
rezerve** — svaki pomeraj nadole, čak i 1 mm, prelama dokument na dve strane.

Zato program radi ovako:

1. Horizontalni pomak se primeni uvek (ne utiče na prelom).
2. Vertikalni pomak se primeni, pa se **proveri broj strana**.
3. Ako je dokument prešao na 2 strane, pokušava se sa donjom marginom na 0.
4. Ako i dalje prelama, vertikalni pomak se **poništava** i program prikaže
   obaveštenje — bolje je odštampati bez pomaka nego pokvariti obrazac.

U praksi pomaci do ~3 mm prolaze; veći se prijave kao upozorenje.

---

## Podešavanja (`Podesavanja.xml`)

Fajl se pravi pored `.exe`-a pri prvom snimanju podešavanja.

| Polje | Značenje |
|---|---|
| `NazivIzborneListe` | naziv izborne liste (prva tabela) |
| `NazivPodnosioca` | stranka / koalicija / grupa građana |
| `DatumIzbora` | zamenjuje „17. децембар 2023" iz šablona |
| `Godina` | zamenjuje „2023." u „У ____, ____ 20__." |
| `MestoOvere` | mesto overe |
| `Opstina` | organ overe — upisuje se u kolonu C spiska potpisnika |
| `OvlasceniOveritelj`, `AdresaOveritelja` | podaci overitelja |
| `OffsetXmm`, `OffsetYmm` | pomak štampe u mm |
| `BrojKopija`, `Stampac` | broj kopija i štampač (prazno = podrazumevani) |
| `PrikaziPreStampe` | umesto štampe otvori dokument u Word-u |
| `Transliteracija` | podaci sa kartice na ćirilicu |
| `VodiEvidenciju` | vođenje spiska potpisnika (CSV) |

Pošto ništa nije hardkodirano, **jedan build radi za sve opštine** — u verziji 2024
je za svaku opštinu pravljen poseban program (Inđija, Niš, Palilula, Pantelej,
Medijana, Crveni krst, Niška Banja).

---

## Evidencija potpisnika

`ListaPotpisnika.csv` — program ga sam napravi pri prvom upisu.

| Kolona | Sadržaj |
|---|---|
| A | Prezime i ime birača |
| B | JMBG |
| C | Organ overe (iz Podešavanja) |
| D | Datum i vreme overe |

Pre upisa se proverava da li JMBG već postoji; ako postoji, program javlja duplikat
sa imenom već upisanog potpisnika i **ne upisuje** red — isto ponašanje kao 2024.

**Zašto CSV a ne Excel:** verzija iz 2024 je zahtevala instaliran Microsoft Excel
i pokretala ga u pozadini pri svakom upisu. CSV se otvara u Excelu, LibreOffice-u
ili Notepad-u, upis je trenutan, a Excel više uopšte nije potreban.
Fajl je UTF-8 sa BOM i tačkom-zarezom kao razdvajačem, pa se ćirilica ispravno
prikazuje kada se dvoklikom otvori u Excelu.

## Ako Word nije instaliran

Word je **obavezan** — bez njega se obrazac ne može odštampati. Program to
proverava pri pokretanju i pri štampi, pa umesto nerazumljive COM greške
prikazuje jasnu poruku na srpskom i isključi dugme „Štampaj".

Instaler takođe proverava Word pre instalacije i upozorava ako ga nema.

## Build

```cmd
build.cmd                 :: program (x86 + x64) + portable ZIP
setup\napravi_setup.cmd   :: instalacioni Setup.exe (x86 + x64)
```

Ručno, za jednu platformu:

```cmd
MSBuild Izbori2026izjava.vbproj /p:Configuration=Release /p:Platform=x86
```

## Instalacija (za krajnjeg korisnika)

U `dist\` se nalaze gotovi instaleri — korisnik ih samo dvoklikne:

| Fajl | Za koga |
|---|---|
| `IzjavaBiraca2026_Setup_x86.exe` | 32-bitni Microsoft Office (najčešće) |
| `IzjavaBiraca2026_Setup_x64.exe` | 64-bitni Microsoft Office |

Instaler:
- radi **bez administratorskih prava** (instalira u `%LocalAppData%\Programs`)
- pravi prečice u Start meniju i na radnoj površini
- proverava .NET Framework 4.7.2 i Microsoft Word
- **upozorava ako se bitnost programa i Office-a ne poklapaju** i odbija
  instalaciju pogrešne verzije u tihom režimu
- dodaje stavku u „Add/Remove Programs" sa urednom deinstalacijom

Instaleri se prave pomoću **Inno Setup 6** (besplatan):

```cmd
winget install JRSoftware.InnoSetup
```

> Napomena: prvobitno je planiran Visual Studio Installer projekat (`.vdproj`),
> kao u verziji 2024. Ispostavilo se da ta ekstenzija nije registrovana u
> VS 2022 na ovom računaru — i originalni 2024 setup projekat pada na isti
> način. Nekorišćeni `.vdproj` fajlovi su sačuvani u `setup\_vdproj_nekoriscen\`.

## Zahtevi

- Windows 7 ili noviji
- .NET Framework 4.7.2
- Microsoft Word (testirano sa Office 2010 / 14.0, 32-bit) — Excel NIJE potreban
- Čitač smart kartica i lična karta sa čipom
- `CelikApi.dll` 1.4.2 odgovarajuće bitnosti (ide uz program)
