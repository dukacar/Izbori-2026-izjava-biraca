# IzjavaBirača 2026 — uputstvo

Program čita ličnu kartu preko čitača i štampa **izjavu birača (obrazac NPRS-3)**.

---

## 1. Instalacija

### Šta vam treba pre instalacije

| | |
|---|---|
| Windows | 7 ili noviji |
| Microsoft Word | **obavezan** (bez njega nema štampe) |
| .NET Framework | 4.7.2 ili noviji (obično već postoji na Windows-u) |
| Čitač kartica | priključen na USB |
| Microsoft Excel | **nije potreban** |

### Koju verziju instalirati

Dva instalera:

| Fajl | Za koga |
|---|---|
| `IzjavaBiraca2026_Setup_x86.exe` | 32-bitni Microsoft Office **(najčešći slučaj)** |
| `IzjavaBiraca2026_Setup_x64.exe` | 64-bitni Microsoft Office |

> **Bira se prema Microsoft Office-u, a NE prema Windows-u.**
> Vrlo često je Windows 64-bitni, a Office 32-bitni — tada ide **x86** verzija.

**Kako proveriti bitnost Office-a:**
otvorite Word → **File → Account → About Word** → na kraju naslova piše
„32-bit" ili „64-bit".

Ako niste sigurni — pokrenite `x86`. Ako je pogrešna, instaler će vas
zaustaviti i reći koja vam treba.

### Postupak

1. Dvoklik na odgovarajući `IzjavaBiraca2026_Setup_*.exe`
2. Kliknite **Next** kroz čarobnjak (podrazumevane vrednosti su u redu)
3. **Nisu potrebna administratorska prava** — instalira se u vaš korisnički profil
4. Na kraju možete čekirati „Pokreni IzjavaBirača 2026"

Program se instalira u:

```
C:\Users\<korisnik>\AppData\Local\Programs\IzjavaBiraca2026_x86
```

Prečice se prave u **Start meniju** i na **radnoj površini**.

### Deinstalacija

**Settings → Apps → IzjavaBirača 2026 → Uninstall**, ili prečica
„Deinstaliraj IzjavaBirača 2026" u Start meniju.

---

## 2. Gde se nalaze koji fajlovi

Svi radni fajlovi su **u istom folderu gde je program** (`IzjavaBiraca2026.exe`).
Do njega najbrže dođete tako što desni klik na prečicu → **Open file location**.

| Fajl | Šta je | Odakle dolazi |
|---|---|---|
| `IzjavaBiraca2026.exe` | sam program | instaler |
| `NPRS-3 - izjava biraca da podrzava izbornu listu.doc` | **Word šablon obrasca** | instaler |
| `CelikApi.dll` | biblioteka za čitanje lične karte | instaler |
| `Podesavanja.xml` | vaša podešavanja | program ga sam napravi |
| `ListaPotpisnika.csv` | **spisak potpisnika** | program ga sam napravi |

### Ako treba zameniti šablon obrasca

Kada RIK objavi novi obrazac, ili ako želite izmenjenu verziju:

1. Zatvorite program
2. Novi fajl kopirajte u folder programa
3. Nazovite ga **tačno** `NPRS-3 - izjava biraca da podrzava izbornu listu.doc`
   (ili u `Podesavanja.xml` promenite `NazivSablona` na novi naziv)
4. Pokrenite program — pri startu proverava šablon i javi ako nešto ne valja

> Šablon mora biti **Word `.doc`** fajl sa istom strukturom kao originalni
> NPRS-3 obrazac: tabela sa poljem za ime, **13 kutijica za JMBG** i linijom
> za adresu. Program se ne oslanja na bookmarke — piše direktno u ćelije tabele.
>
> Program **nikada ne menja** šablon: otvara ga samo za čitanje.

### Spisak potpisnika (CSV)

`ListaPotpisnika.csv` nastaje sam pri prvom štampanju, u folderu programa.

| Kolona | Sadržaj |
|---|---|
| A | Prezime i ime birača |
| B | JMBG |
| C | Organ overe (iz Podešavanja) |
| D | Datum i vreme |

Otvara se dvoklikom u **Excelu**, ali i u **LibreOffice-u** ili **Notepad-u** —
Excel nije potreban za rad programa.

**Ako isti JMBG već postoji**, program to javi, kaže za koga je već upisan,
i **neće** ga upisati drugi put.

> Ako želite da počnete nov spisak (npr. za druge izbore), zatvorite program
> pa preimenujte ili premestite `ListaPotpisnika.csv`. Program će napraviti novi.

> Ne držite CSV otvoren u Excelu dok štampate — Excel zaključa fajl i upis
> neće uspeti. Program će vas o tome obavestiti.

---

## 3. Prvo pokretanje — podešavanja

Pri prvom pokretanju otvorite **Podešavanja** i popunite podatke.
Bez toga obrazac se štampa nepopunjen u delu izborne liste.

### Izborna lista

| Polje | Šta upisati |
|---|---|
| **Naziv izborne liste** | naziv kako glasi na listi |
| **Podnosilac** | politička stranka / koalicija / grupa građana |
| **Datum izbora** | npr. `15. март 2026` — zamenjuje datum iz obrasca |
| **Godina** | npr. `2026` — ide u „У ____, ____ 20__." |
| **Mesto overe** | npr. `ИНЂИЈА` |

### Ovlašćeni overitelj

| Polje | Šta upisati |
|---|---|
| **Organ overe (opština)** | npr. `Општинска управа Инђија` — ide i u spisak potpisnika |
| **Ime i prezime** | ime overitelja |
| **Mesto i adresa** | adresa organa overe |

### Štampa i pomak (offset)

| Polje | Značenje |
|---|---|
| **Pomak levo/desno (mm)** | pozitivno = udesno |
| **Pomak gore/dole (mm)** | pozitivno = nadole |
| **Probna štampa** | odštampa test primerak sa izmišljenim podacima |
| **Poništi pomak** | vraća oba pomaka na 0 |
| **Štampač** | prazno = podrazumevani štampač iz Windows-a |
| **Kopija** | koliko primeraka odjednom |
| **Prikaži dokument u Word-u umesto štampe** | za proveru pre štampe |

### Ostalo

- **Transliteruj podatke na ćirilicu** — podaci sa kartice se prebacuju u ćirilicu
- **Vodi spisak potpisnika (CSV fajl)** — isključite ako ne želite evidenciju

Na kraju **Sačuvaj**. Podešavanja se pamte u `Podesavanja.xml`.

---

## 4. Kalibracija štampe (offset)

Obrazac NPRS-3 se štampa na unapred odštampan formular, pa se tekst mora
poklopiti sa linijama. Postupak:

1. Otvorite **Podešavanja**
2. Kliknite **Probna štampa** sa pomacima 0 / 0
3. Uporedite odštampano sa formularom:
   - tekst previše **levo** → povećajte „Pomak levo/desno" (npr. `2`)
   - tekst previše **desno** → negativna vrednost (npr. `-2`)
   - tekst previše **visoko** → povećajte „Pomak gore/dole"
   - tekst previše **nisko** → negativna vrednost
4. Ponavljajte probnu štampu dok se ne poklopi
5. **Sačuvaj**

> **Savet:** prve probe radite na običnom papiru, ili izaberite štampač
> „Microsoft Print to PDF" pa pogledajte rezultat na ekranu — tako ne trošite
> formulare.

### Zašto pomak nadole ima granicu

Obrazac staje na **tačno jednu stranu bez rezerve**. Preveliki pomak nadole
gurnuo bi sadržaj na drugu stranu, pa ga program u tom slučaju **ne primeni**
i o tome vas obavesti. Pomak levo/desno se uvek primenjuje.

U praksi pomaci do oko **3 mm** prolaze bez problema.

---

## 5. Svakodnevni rad

1. Ubacite ličnu kartu u čitač
2. **Pročitaj podatke** — podaci i slika se pojave na ekranu
3. Proverite da su podaci ispravni
4. **Štampaj** — obrazac ide na štampač, a potpisnik se upiše u spisak
5. **Obriši podatke** pre sledećeg birača

**Ćirilica / Latinica** (desno dole) bira pismo. Ako promenite pismo posle
čitanja kartice, ponovo pročitajte karticu da bi se primenilo.

---

## 6. Rešavanje problema

| Poruka / pojava | Šta uraditi |
|---|---|
| „Proverite da li je čitač kartica priključen" | Priključite čitač; proverite da Windows vidi uređaj |
| „kartica nije ubačena u čitač" | Gurnite karticu do kraja, čipom nagore |
| „Microsoft Word nije pronađen" | Instalirajte Microsoft Office (Word) |
| „CelikApi.dll nije pronađen ili nije odgovarajuće arhitekture" | Instalirali ste pogrešnu bitnost — deinstalirajte pa instalirajte drugu verziju (x86 ↔ x64) |
| „Word šablon nije pronađen" | Vratite `.doc` šablon u folder programa (vidi tačku 2) |
| „Šablon nema očekivanu strukturu" | Šablon nije NPRS-3 obrazac ili je izmenjen — vratite ispravan |
| „JMBG ... već postoji u spisku" | Taj birač je već potpisao — provera duplikata radi ispravno |
| „Fajl ListaPotpisnika.csv je zauzet" | Zatvorite ga u Excelu pa ponovite |
| Tekst se ne poklapa sa formularom | Podesite offset (tačka 4) |
| Štampa prelazi na dve strane | Smanjite pomak nadole |

### Ćirilica se u Excelu prikazuje kao znakovi bez smisla

Fajl je ispravan (UTF-8). Ako se dvoklikom otvori pogrešno, u Excelu idite na
**Data → From Text/CSV**, izaberite fajl, i podesite:
- **File Origin:** `65001: Unicode (UTF-8)`
- **Delimiter:** `Semicolon` (tačka-zarez)

---

## 7. Rad za više opština

Jedan isti program radi za sve opštine — sve se menja u **Podešavanjima**,
ništa nije ugrađeno u sam program.

Ako radite za više opština na istom računaru, posle rada za jednu opštinu
premestite `ListaPotpisnika.csv` na sigurno, promenite podatke u Podešavanjima,
i nastavite — program će napraviti nov spisak.

---

## 8. Za programera

<details>
<summary>Tehnički detalji (kliknite da otvorite)</summary>

### Build

```cmd
build.cmd                 :: program (x86 + x64) + portable ZIP u dist\
setup\napravi_setup.cmd   :: instalacioni Setup.exe (x86 + x64) u dist\
```

Za instalere je potreban Inno Setup 6:

```cmd
winget install JRSoftware.InnoSetup
```

### Zašto x86 i x64

Word se koristi preko COM interopa i učitava se **u isti proces**, pa bitnost
programa mora da odgovara bitnosti Office-a. Uz svaku verziju ide odgovarajući
`CelikApi.dll` (iz `lib\x86\` odnosno `lib\x64\`) — build ga kopira sam.

### Čelik API

Koristi se **CelikApi.dll 1.4.2**, inicijalizacija `EidStartup(4)`.

### Kako se popunjava obrazac

NPRS-3 `.doc` nema bookmarke ni form fields, pa se piše direktno u ćelije
**Tabele 3** (blok „Б И Р А Ч"):

```
Red 2  ime i prezime      → 1. paragraf ćelije (oznaka ostaje u 2. paragrafu)
Red 3  JMBG               → 13 ćelija, po jedna cifra
Red 5  mesto i adresa     → 1. paragraf ćelije
```

Deo o overitelju se popunjava zamenom teksta, uz očuvanje dužine polja
(donje crte), da se ne promeni prelom strane.

Šablon se otvara `ReadOnly:=True` i zatvara sa `SaveChanges:=False`.

> **Napomena:** `Documents.Open` mora imati `Visible:=True`, inače dokument
> nema aktivan prozor i `PrintOut` puca sa porukom
> „This method or property is not available because a document window is not active".
> Sama Word aplikacija ostaje nevidljiva (`app.Visible = False`), pa korisnik
> ništa ne vidi.

### Struktura projekta

```
Izbori2026izjava\
├─ Form1.vb / .Designer.vb            glavni ekran
├─ FrmPodesavanja.vb / .Designer.vb   ekran sa podešavanjima
├─ WordStampac.vb                     popunjavanje i štampa obrasca
├─ PodaciBiraca.vb                    podaci sa kartice
├─ Podesavanja.vb                     Podesavanja.xml
├─ CsvEvidencija.vb                   spisak potpisnika + duplikati
├─ OfficeProvera.vb                   provera da li je Word instaliran
├─ CelikApi.vb                        P/Invoke sloj za CelikApi.dll
├─ SmartcardManager.vb                rad sa čitačem
├─ Transliteration.vb                 latinica ↔ ćirilica
├─ lib\x86\ , lib\x64\                CelikApi.dll 1.4.2 po bitnosti
├─ setup\                             Inno Setup skripta
└─ doc\                               dokumentacija Čelik API-ja
```

### Podesavanja.xml

| Polje | Značenje |
|---|---|
| `NazivIzborneListe`, `NazivPodnosioca` | podaci izborne liste |
| `DatumIzbora`, `Godina`, `MestoOvere` | datumi i mesto overe |
| `Opstina`, `OvlasceniOveritelj`, `AdresaOveritelja` | podaci overitelja |
| `OffsetXmm`, `OffsetYmm` | pomak štampe u mm |
| `BrojKopija`, `Stampac` | broj kopija, štampač (prazno = podrazumevani) |
| `PrikaziPreStampe` | otvori u Word-u umesto štampe |
| `Transliteracija` | ćirilica |
| `VodiEvidenciju` | vođenje spiska potpisnika |
| `NazivSablona`, `NazivCsvFajla` | nazivi fajlova u folderu programa |

</details>
