# Handoff — SuperOffice Health Check · Company Overview (kleur-, contrast- & tabel-refactor)

## Overzicht
Dit pakket beschrijft een **afgeronde design-iteratie** op het *Company · Overview*-dashboard van de SuperOffice Health Check (het "Data Analysis"-rapport, Hugo-gegenereerd). De iteratie is **geen nieuw scherm** — het is een systematische **kleur-, contrast- en tabel-refactor** van het bestaande dashboard. Doel: één samenhangend token-systeem, WCAG AA-tekst, twee strikt gescheiden kleur-rampen, en dichte datatabellen die crisp lezen — zonder de warme, eigen "SuperOffice-groen"-identiteit te verliezen.

De implementerende sessie moet deze wijzigingen doorvoeren in de **echte CSS/JS/HTML** van het dashboard, met behoud van de hieronder beschreven principes.

> **Scope-update (na review van de live app):** de refactor wordt **app-breed** uitgerold over álle Health Check-pagina's (Company, Contact, Sale, Requests, Project, Activities/Momentum, Adoption, Integrity, Analyze All, Settings, Intake). De bevestigde doel-shell is **Zwevend mono** — de live app migreert dus van witte kaarten + donkergroene sidebar naar getinte kaarten op een wit zwevend paneel met blended shell. Alle nieuwe componenttypes staan gerenderd in **`Component reference - Zwevend mono.html`** (de visuele spec); zie ook de sectie *App-brede uitrol* onderaan.

---

## Over de designbestanden
De bestanden in dit pakket zijn **design-referenties, gemaakt in HTML** — een prototype dat de bedoelde look en gedrag toont, **niet** productiecode om letterlijk te kopiëren. De opdracht is om dit ontwerp te **reproduceren in de bestaande codebase** (Hugo-templates + de bijbehorende CSS/JS) met de daar gebruikelijke patronen. Waar het prototype losse "review-hulpmiddelen" bevat (zie hieronder), horen die **niet** in productie.

### Fidelity: **hi-fi**
Definitieve kleuren (exacte hex), typografie, spacing en interacties. Reproduceer pixel-nauwkeurig; het prototype `Company Overview - Neutral shell.html` is de leidende bron van waarheid voor alles wat hieronder niet exact vermeld staat.

### Review-scaffolding — NIET naar productie
Drie elementen in het prototype zijn puur voor design-review en moeten worden **weggelaten** (of bewust als echte feature heringericht):
- **Annotatie-overlay** — de zwarte knop rechtsonder *"Toon verbeteringen"* + de oranje genummerde pins (`.annot-toggle`, `.pin`, `body.show-annot`, `.ann-target`). Puur toelichting; verwijderen.
- **Font-switcher** linksonder (`.fontpick`, *"Lato / Hanken / DM Sans"*). Keuze-instrument; de **beslissing is DM Sans** — zet dat vast als het enige font en verwijder de switcher.
- **Stijl-schakelaar** linksonder (`.surfacepick`, *"Plat & warm / Zwevend mono"*). Dit is een **keuze-instrument** om twee thema's te vergelijken. Kies voor productie het **standaardthema (Zwevend mono)**; exposeer thema-keuze alleen als je er bewust een echte instelling van maakt (zie *Surface-varianten*).

---

## De kernbeslissingen (MOET behouden blijven)
Dit is de essentie van de iteratie. Wie de code aanpast moet deze zeven principes respecteren, anders valt het systeem uit elkaar.

**P1 — Twee STRIKT gescheiden kleur-rampen. Nooit mengen.**
- **SEMANTISCH** (`--good` / `--mod` / `--bad`): *alleen* voor gezondheid/score/kwaliteitsstatus. Elke kleur heeft vier tokens: solid (`--good`) · zachte vulling (`--good-bg`) · AA-tekst (`--good-ink`) · hairline (`--good-line`).
- **CATEGORISCH** (`--c1`…`--c6` + `--cn`): brand-teal, *alleen* voor verdelingen/segmenten (niet-semantisch). `--cn` = neutraal "overig/rest".
- Gebruik nooit een semantische kleur voor een categorische slice of andersom. Concreet: **"Dormant tail" is overal neutraal (`--cn`)** — in de KPI-tegel én in de composition-balk — nooit amber.

**P2 — "Recessed" tokens stappen OMLAAG vanaf de kaart, niet omhoog naar wit.**
Lege rails/tracks (`--track` voor lege donut-boog, progress- en minibars), `--line`, `--line-soft` en `--group-band` moeten leesbaar zijn *tegen het kaartoppervlak*. In het getinte Zwevend-mono-thema worden ze dieper getint zodat ze niet wegvallen. Richtcijfer: een lege rail leest ~**1.35:1** t.o.v. zijn kaart (zoals de neutrale "rest"-stip).

**P3 — Twee chip-families.**
- **Semantische chip** = zachte gevulde pill (`-bg`) + gekleurde stip + AA-tekst (`-ink`) + hairline (`-line`). (Good/Moderate-badges, Required, passchip, "Worth checking"-tag.)
- **Neutrale/categorische chip** = hairline (rand) + minimale/transparante vulling; de tekst of stip draagt de betekenis. (typechips, importance Normal/Excluded, count-pills, groeps-tellingen, "Recommended"-tag.)
- Doel: geen "zwevende gekleurde blokjes". Gevuld = semantisch, omlijnd = categorisch/neutraal.

**P4 — Tekst haalt WCAG AA op de getinte kaart.**
`--muted #5a6760` (~5.2:1) en `--faint #6f7972` (~3.9:1) op de mono-kaart. De semantische `-ink`-tokens zijn AA op hun eigen `-bg`-vulling.

**P5 — De "Filtered"-balk zit in de kaart-familie** (niet in een losse zandkleur). De periode is een echte control-knop; "Last scan" is een zachte status-chip (good-familie).

**P6 — Amber is gedempte ocher (`--mod #b9832f`), geen felle oranje.** Houdt het salie/zand-palet rustig. Geldt voor badge, bar, tag én tabel-stippen.

**P7 — Datatabellen krijgen de "sheeted"-behandeling** (zie eigen sectie): getinte kaart + getinte kop, datarijen full-bleed wit tot de getinte rand, geen zebra. Insight-, KPI- en hero-kaarten blijven getint.

---

## Design tokens

### Thema-onafhankelijke tokens (basis — gelijk in beide thema's)

**Tekst / inkt**
| token | hex | gebruik |
|---|---|---|
| `--ink` | `#173a32` | koppen, primaire tekst, cijfers |
| `--ink-soft` | `#2c4a42` | tabel-cellen, body |
| `--muted` | `#5a6760` | secundaire tekst (AA) |
| `--faint` | `#6f7972` | labels/eyebrows, meta (AA-groot) |

**Semantische ramp** (alleen status/score)
| rol | solid | `-bg` (vulling) | `-ink` (tekst) | `-line` (rand) |
|---|---|---|---|---|
| good | `#2f8f5b` | `#e7f3ec` | `#226a43` | `#cbe4d3` |
| mod (ocher) | `#b9832f` | `#f3e4ca` | `#785012` | `#e4caa0` |
| bad | `#c0473a` | `#f8e9e6` | `#9e3529` | `#e6c1b6` |

> De `-bg`-waarden hierboven zijn de **witte-context** (basis) waarden; het Zwevend-mono-thema verdiept ze — zie themablok.

**Categorische ramp** (alleen verdelingen, brand-teal)
| token | hex | | token | hex |
|---|---|---|---|---|
| `--c1` | `#0d5f59` | | `--c4` | `#82c2b8` |
| `--c2` | `#2f817a` | | `--c5` | `#b3ddd5` |
| `--c3` | `#51a399` | | `--c6` | `#dcebe7` |
| `--cn` (neutraal "overig") | `#c8c2b4` | | | |

**Merk / sidebar**
| token | hex | gebruik |
|---|---|---|
| `--green` | `#15564b` | merk-groen (logo/woordmerk) |
| `--green-deep` | `#0c4a3f` | actieve tab-onderrand, primaire knoppen, actieve nav-tekst |
| `--side-text` | `#303a35` | nav-item tekst (rust **én** hover **én** actief — verandert niet per state) |
| `--side-faint` | `#837e72` | nav-sectiekoppen, credit |
| `--side-active` | `#e7e1d4` | achtergrond actief nav-item (subtiele warme tint op het beige menu) |
| `--side-2` | `#e7e1d4` | achtergrond hover nav-item — **gelijk aan `--side-active`** |

**Neutrale chip-taal (basis)**
`--chip-fill #f1efe8` · `--chip-line #ddd8cc` · `--chip-ink #625d51` · `--chip-line-teal #cfe0da`

### Surface-tokens per thema (effectieve, reeds samengevoegde waarden)

**Thema A — "Zwevend mono" (STANDAARD)** — `html[data-surface="float-mono"]`
Getinte kaarten op een wit zwevend paneel; eigenzinnig, "designed".
| token | waarde |
|---|---|
| `--shell` (canvas + menu) | `#f2efe9` |
| `--page` (zwevend paneel) | `#ffffff` |
| `--card` | `#f2efe9` |
| `--card-border` | `#e8e4dc` |
| `--card-shadow` | `none` |
| `--line` / `--line-soft` | `#e0dacd` / `#e6e1d5` |
| `--track` | `#d5cec0` |
| `--group-band` / `--zebra` | `#e6ece8` / `#ede9e0` |
| `--chip-fill` / `--chip-line` / `--chip-line-teal` | `#ece6db` / `#d7d0c2` / `#bcd2cc` |
| `--good-bg` / `--mod-bg` / `--bad-bg` | `#d3e8d8` / `#e8d5b0` / `#eccfc6` |
| `--side-active` / `--side-2` | `#e7e1d4` / `#e7e1d4` (gelijk) |

Alle overige tokens (tekst, semantische solids + `-ink` + `-line`, categorische ramp) = basiswaarden.

**Thema B — "Plat & warm"** — `html[data-surface="warm"]`
Licht menu + lichte kaarten op een warm-donker zwevend paneel; maximale data-leesbaarheid.
| token | waarde |
|---|---|
| `--shell` (canvas + menu, licht) | `#fcfbf7` |
| `--page` (zwevend paneel, warm-donker) | `#ece9e2` |
| `--card` | `#fcfbf7` |
| `--card-border` | `#e5e1d6` |
| `--side-active` (groene pill) | `#dfeae4` |
| `--side-2` (hover) | `rgba(21,86,75,.07)` |

Alle overige tokens = basiswaarden (o.a. `--track #e7e2d7`, chips en semantische `-bg` in witte-context).

> **Waarom de menu-hover/actief in Thema B anders is:** het menu is daar licht, dus de witte actief-pill uit Thema A zou onzichtbaar zijn. Actief wordt een zachte groene pill (`--side-active #dfeae4`), hover een lichte groene wash (`--side-2`).

**Beide thema's** laten het hoofd­paneel *zweven*:
`.main { margin: 12px 12px 12px 6px; border-radius: 18px; box-shadow: 0 2px 12px rgba(40,45,40,.07); }`
en `.topsticky { border-radius: 18px 18px 0 0; }`.

> Het kale `:root` (zonder `data-surface`) is **geen** los te kiezen thema meer — het levert alleen de basiswaarden waar A en B op voortbouwen. De schakelaar biedt uitsluitend A en B; oude waarden `flat`/`float-warm` in `localStorage` worden op `warm` gemapt.

---

## De "sheeted" tabelbehandeling (P7 — belangrijk)
Alle **echte datatabellen** krijgen de klasse `sheeted` op hun `.card`. Effect: de kaart en de kolomkop houden de kaart-tint, maar de **datarijen zijn puur wit en lopen full-bleed door tot de getinte kaartrand**, met afgeronde onderhoek en **zonder zebra**.

Kernregels (zie prototype voor het geheel):
```css
.card.sheeted { padding: 18px 0 0;
  /* reset recessed/​fill-tokens naar witte-context zodat de witte data crisp leest */
  --track:#e7e2d7; --chip-fill:#f1efe8; --chip-line:#ddd8cc;
  --line:#e6e2d9; --line-soft:#eeebe3; --group-band:#eef2f0;
  --good-bg:#e7f3ec; --mod-bg:#f3e4ca; --bad-bg:#f8e9e6; }
body:not(.show-annot) .card.sheeted { overflow:hidden; }      /* laatste rij loopt naar de ronde rand */
.card.sheeted .cardhead2,
.card.sheeted p.sub { padding: 0 24px; }                       /* kop houdt kaart-tint + inspringing */
.card.sheeted .tbl:not(.assoc) thead th { border-bottom:none; }
.card.sheeted .tbl:not(.assoc) thead th:first-child,
.card.sheeted .tbl:not(.assoc) tbody td:first-child { padding-left:24px; }
.card.sheeted .tbl:not(.assoc) thead th:last-child,
.card.sheeted .tbl:not(.assoc) tbody td:last-child  { padding-right:24px; }
.card.sheeted .tbl:not(.assoc) tbody td { background:#ffffff; }
.card.sheeted .tbl:not(.assoc) tbody tr:first-child td { border-top:1px solid var(--card-border); }
/* Associate breakdown: member-rijen wit (geen zebra), kolomkoppen + groepsbanden blijven getint */
.card.sheeted .assoc-scroll,
.card.sheeted .memberrow td { background:#ffffff; }
```

**Welke kaarten zijn `sheeted`:** Quality flags · Standard field completeness · Custom fields (Company) · Adoption score · Company segments · Associate breakdown — kortom elke datatabel.
**Blijven getint (NIET sheeted):** de health-hero (donut), KPI-tegels, Database composition, Category mix, de bar-chart en alle insight-kaarten.

> In Thema B (lichte kaarten) is `sheeted` in de praktijk een no-op qua kleur — kaart en data zijn daar toch al licht; de regels blijven veilig.

---

## Componenten (specificatie)

**Layout-shell** — `.app { display:grid; grid-template-columns:250px 1fr; height:100% }`. Sidebar (`.side`) is transparant en versmelt met `--shell`. Hoofdpaneel `.main` scrollt zelf; `.topsticky` (kop + tabs) plakt bovenaan en krijgt een schaduw bij scroll (`.stuck`).

**Sidebar** — merkblok met gradient-mark (`linear-gradient(160deg,#1f7a6b,#0c4a3f)`), nav-secties ("Entities/Other/Tools") in `--side-faint` uppercase.

**Nav-items — kleurgedrag (Zwevend mono):**
- **Rust:** achtergrond transparant (menu blend't met `--shell #f2efe9`); tekst + icoon `--side-text #303a35` (charcoal), gewicht 500.
- **Hover:** achtergrond `--side-2 #e7e1d4`; tekst/icoon ongewijzigd (`--side-text`).
- **Actief:** achtergrond `--side-active #e7e1d4` (**zelfde kleur als hover**); tekst/icoon `--side-text` (charcoal, **niet** groen); gewicht **700**. Het onderscheid rust↔actief komt dus van de tint + bold, niet van een kleurwissel; hover en actief delen de achtergrond en verschillen alleen in gewicht.
- Icoon volgt `currentColor` (dus automatisch charcoal). Radius 10px, padding 9px 12px.
- **Thema B (Plat & warm):** menu is licht, dus daar `--side-active #dfeae4` (zachte groene pill) en `--side-2 rgba(21,86,75,.07)`; verder identiek gedrag.

`.nav-item { color:var(--side-text); font-weight:500 } .nav-item:hover { background:var(--side-2) } .nav-item.active { background:var(--side-active); color:var(--side-text); font-weight:700 }`

**Tabs** — `.tab` 14.5px/600, `--muted`; actief `--ink` met 2.5px onderrand `--green-deep`.

**Filtered-balk** (P5) — `.filter` = kaart-familie (`--card` + `--card-border`, radius 12px). `.period-btn` echte knop (`--page`-vlak, `--chip-line`-rand). `.scan` zachte good-chip (pill).

**Health-hero** — `.card.hero`, grid `220px 1fr`. Donut 132px: `conic-gradient(from -90deg, var(--good) 0 81%, var(--track) 81% 100%)`, binnenring `inset:15px` met kaartkleur. Sub-rijen met `.badge` (good/mod), gewicht-chip `.w` (neutrale hairline-chip) en progress `.bar` (rail = `--track`).

**KPI-tegels** — `.kpi` radius 14px, zachte schaduw. Waarde 33px/800, `tabular-nums`. `.kdot` semantisch of neutraal (let op: **Dormant tail-stip = `--cn`**, niet mod).

**Database composition / Category mix / Engagement** — gestapelde balken met de **categorische** ramp; `--cn` voor de dormant/"overig"-rest. Legenda-swatches 10–11px, radius 3px.

**Bar-chart** — kolommen `--c1`, gedeeltelijke/prognose-kolom `--c4`; `.chartfoot` = neutrale hairline-notitie (`--chip-fill` + `--chip-line`).

**Insights** — `.insight` kaart, **geen** linker-accentbalk; prioriteit via tag + volgorde. `.insight.primary` = zachte teal band (`--group-band` + `--chip-line-teal`). Icoon-tegels gelijkgetrokken als **paar**: beide een zachte tint van hun eigen accent — Recommended teal (`background:#d3e6de; border:var(--chip-line-teal); color:var(--c1)`), Worth-checking ocher (`background:var(--mod-bg); border:var(--mod-line); color:var(--mod-ink)`). `.tag` = teal hairline-chip; `.tag.warn` = mod-familie. Primaire knop `--green-deep`.

**Tabellen** — `.tbl` 13.5px; `th` 11px uppercase `.07em`, `--faint`. Cellen `--ink-soft`, rij-rand `--line-soft`. `.minibar` rail = `--track`, vulling semantisch. Importance-chips: `.imp.req` = good-chip (gevuld), `.imp.norm` = neutrale hairline-chip, `.imp.exc` = transparant + hairline. `.typechip` = neutrale hairline-chip. `.count` = neutrale count-pill.

**Associate breakdown** — sticky kolomkoppen (`top:0`) + sticky, inklapbare groepskoppen (`.grouprow`, `top:37px`, `--group-band`). Groeps-telling `.gn` = teal hairline-pill. Member-rijen wit (sheeted), geen zebra.

---

## Interacties & gedrag (productie-relevant)
- **Tabs**: klik wisselt `.tabpanel[hidden]`; `main.scrollTop = 0` bij wissel.
- **Sticky-schaduw**: `.topsticky` krijgt `.stuck` zodra `main.scrollTop > 4`.
- **Inklapbare groepen** (Associate breakdown): klik op `.grouprow` toggelt `.collapsed` en verbergt volgende `.memberrow`-zussen; chevron `.gchev` roteert 90°.
- **Thema-schakelaar** (alleen als je thema-keuze wilt shippen): `applySurface(s)` zet/verwijdert `data-surface` op `<html>` en bewaart in `localStorage['so_surface']`. Standaard = `float-mono`. Legacy-waarden `flat`/`float-warm` → `warm`.

**Niet naar productie:** annotatie-overlay (`.annot-toggle`, `.pin`, `body.show-annot`, `.ann-target`).

---

## Typografie & iconen
- **Font**: DM Sans (Google Fonts), gewichten 400/500/600/700/800. Cijfers vaak `font-variant-numeric: tabular-nums`. (Gekozen boven Hanken Grotesk en Lato om leesbaarheid in dichte tabellen/KPI-cijfers; het prototype bevat een preview-switcher tussen de drie — review-only.)
- **Type scale**: zie **`Type scale - Zwevend mono.html`** (+ `.png`) — elk element gerenderd op zijn exacte px/gewicht/spacing/kleur. Trap (px): 33 KPI · 30 donut · 20 titel · 19 merk · 16 kaarttitel · 15.5 hero-label · 15 nav-top/% · 14.5 nav+tab · 14 body/filter · 13.5 tabel/legenda/subkop · 13 KPI-sub · 12.5 count/legenda · 12 chip/staaflabel · 11.5 sectie+KPI-label · 11 kolomkop/importance · 10.5 nav-sectie/tag. Gewichten: 800 display+KPI · 700 titels/labels/nadruk · 600 tabs/subkop · 500 nav-rust · 400 body/tabel.
- **Schaal (px)**: h1 20/800 · kaarttitel 16–17/700 · KPI-waarde 33/800 · donut-waarde 30/800 · body/tabel 13.5–14 · labels/eyebrows 10.5–11.5 uppercase, letter-spacing .07–.13em.
- **Iconen**: inline SVG in Feather/Lucide-stijl, `stroke-width` 1.8–1.9, ronde caps/joins, `currentColor`.

---

## Contrast-intentie (ter controle bij implementatie)
- `--muted` op mono-kaart ≈ **5.2:1**, `--faint` ≈ **3.9:1** (beide AA voor hun tekstgrootte).
- Lege `--track`-rails ≈ **1.35:1** t.o.v. de kaart (bewust zichtbaar, niet wegvallend).
- Semantische `-ink` op eigen `-bg` ≥ **4.5:1**.
Behoud deze verhoudingen als je tinten in de echte codebase minimaal moet bijstellen.

---

## App-brede uitrol (overige pagina's)
Na review van de live Health Check breidt dit pakket uit van *Company · Overview* naar de **hele app**. De onderstaande component-types komen niet voor op Company Overview en zijn daarom apart gespecificeerd. **`Component reference - Zwevend mono.html`** toont ze allemaal gerenderd met de exacte tokens — dat bestand is de visuele bron van waarheid; onderstaande tekst geeft de regels en hex.

### Twee extra rampen (naast semantisch + categorisch)
Om P1 overeind te houden krijgen niet-semantische, geordende attributen een **eigen** ramp — nooit groen/rood lenen.

**Neutrale nadruk-ramp — Weight (High/Medium/Low) & Importance (Required/Normal/Excluded).**
Rangorde, geen status. Zwaarder = donkerder neutraal.
| stap | vulling | tekst | rand |
|---|---|---|---|
| strong (High / Required) | `#3d4a44` | `#ffffff` | — |
| mid (Medium / Normal) | `#ece6db` (`--chip-fill`) | `#625d51` (`--chip-ink`) | `#d7d0c2` (`--chip-line`) |
| faint (Low / Excluded) | transparant | `#6f7972` (`--faint`) | `#d7d0c2` |
> **Vervangt** het groene "Required/High" uit de live app. Groen blijft strikt voor gezondheid.

**Level-ramp — engagement (Power/Regular/Low/Dormant).**
Eigen geordende schaal; **geen blauw**, geen alarm-rood. Dormant = neutraal (kans, geen fout — net als de dormant tail).
| level | soft-vulling | tekst | stip |
|---|---|---|---|
| Power | `#d3e8d8` | `#226a43` | `--good #2f8f5b` |
| Regular | `#d6e7e3` | `#1f6a63` | `#2f817a` (`--c2`) |
| Low | `#e8d5b0` | `#785012` | `--mod #b9832f` |
| Dormant | `#e4ded3` | `#625d51` | `--cn #c8c2b4` |

### Nieuwe componenten (met kleur-regel)
- **Score-pill** (grote "87%") — gevuld, kleurt naar band: ≥70 `#1f7a4d` · 40–69 `#b9832f` · <40 `#c0473a`, witte tekst, radius 9px. (Health-score → semantisch, P1-conform.)
- **Donut / pie** (Sale status split) — segmenten = **gedempte semantische** solids (win/verlies/open); "stalled/none" = `--cn`. Restring = `--track`.
- **Ranked bar list** — rail `--track`; vulling volgt de betekenis: **categorie-verdeling** (status/priority/type) = teal-ramp; **completeness-%** = semantisch (`--good`, `--mod` onder drempel).
- **Stacked distribution bar** — categorie = teal; "no value" = `--cn`; geordende waarschijnlijkheid mag ocher→teal→groen lopen.
- **Trend (lijn + staaf)** — staven teal (`--c4`), trendlijn één merk-accent (`--green-deep`), raster `--line`.
- **Funnel / pipeline-tabel** — completeness-minibars semantisch; conversie-cijfers neutraal `--ink`; weight = nadruk-ramp.
- **User-leaderboard** — level-badges (zie boven); sheeted-tabel.

### Controls (Intake · Analyze All · Settings)
Interactieve controls gebruiken **merk-groen** voor selectie/actief — dat is **UI-state, geen data**. P1 (twee rampen) geldt uitsluitend voor datavisualisatie.
- **Segmented control** — actief `--green-deep` + wit; track `--chip-fill` + `--chip-line`.
- **Radio-keuzekaarten** — geselecteerd `--good-bg` + `--good-line` + check; "Recommended"-tag = `--green-deep`-pill.
- **Slider** — gevuld `--green-deep`, rest `--track`, aparte huidig- + doel-marker.
- **Select / dropdown** — `--page`-vlak + `--chip-line`-rand. **Vervangt de lavendel/paarse dropdowns** uit Settings.
- **Modal** (Settings) — kaart-surface + linker-subnav; scrim `rgba(20,25,20,.4)`.
- **Sticky action bar** (Intake) — kaart-familie + samenvattings-chips + primaire CTA.

### Insight-kaarten — één verzoende vorm (app-breed)
De live app gebruikt overal een gekleurde **linker-accentbalk**; het prototype had die verwijderd. Voor consistentie over álle pagina's geldt app-breed: **één dunne (3px) gedempte linker-accent** als prioriteitssignaal + uppercase lead-tag; **geen** zware gevulde icoon-tegel. Accent-kleur = gedempt `--good`/`--mod`/`--bad`, of `--cn` voor pure context. Dit **supersedeert** de accentloze insight uit `Company Overview - Neutral shell.html`.

### Kruispunt-fixes (checklist bij de uitrol)
- **Weight/Importance** groen → neutrale nadruk-ramp.
- **User-levels** groen/blauw/amber/rood → level-ramp (blauw vervalt; Dormant neutraal).
- **Settings-dropdowns** lavendel/paars → neutrale select in het palet.
- **Insights** → dunne gedempte linker-accent overal.
- **KPI-"slechte" cijfers** fel rood → gedempt `--bad-ink #9e3529`.
- **Shell** → migreer naar Zwevend mono (getinte kaarten, blended shell, wit zwevend paneel).

---

## Implementatie-checklist
1. Neem het volledige **token-blok** over (basis + de twee `[data-surface]`-blokken). Houd de driedeling *basis → thema-override → `.sheeted`-reset* intact.
2. Zet **Zwevend mono** als standaardthema.
3. Voer de **twee chip-families** door (P3) — controleer elke chip: semantisch=gevuld, categorisch/neutraal=hairline.
4. Pas de **`sheeted`**-klasse toe op alle datatabel-kaarten (P7); laat insight/KPI/hero getint.
5. Verifieer **P1**: geen semantische kleur op een verdeling; "Dormant tail" overal `--cn`.
6. Strip de **review-scaffolding** (annotaties; en de stijl-schakelaar tenzij bewust als instelling gewenst).
7. Controleer de **contrast-intentie** hierboven na eventueel bijstellen.
8. **App-breed** (zie sectie hierboven + component reference): voer de nadruk-ramp (Weight/Importance), level-ramp, score-pill, de nieuwe datavis-/control-componenten en de verzoende insight-vorm consistent door; ruim de kruispunt-fixes op (groene weight, blauwe levels, paarse dropdowns, felle KPI-cijfers).

---

## Bestanden in dit pakket
- **`Company Overview - Neutral shell.html`** — het definitieve, zelfstandige prototype van *Company · Overview*. Leidende bron van waarheid voor die pagina. Open in een browser; wissel thema linksonder, bekijk de tabbladen (Overview / Data Quality / Adoption) voor de sheeted-tabellen.
- **`Component reference - Zwevend mono.html`** — de **visuele spec voor de app-brede uitrol**: elk nieuw component (rampen, chips, score-pill, donut, ranked/stacked bars, trend, sheeted-tabellen, level-badges, controls, insights) gerenderd in Zwevend mono met de exacte tokens in `:root`. Open naast deze README.
- **`Table treatments - Zwevend mono.html`** — *alleen ter onderbouwing*: de drie onderzochte tabel-richtingen (A wit vel / B lichte kaart / C zebra). De **gekozen** richting is A, toegepast op álle datatabellen (zie P7). Niet de target, puur rationale.
