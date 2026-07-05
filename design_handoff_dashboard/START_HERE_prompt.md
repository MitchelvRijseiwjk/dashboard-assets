# START HIER — kickoff-prompt voor de implementerende Claude

Plak de tekst hieronder als je **eerste bericht** in de sessie die de Health Check-code aanpast. Zorg dat de map `design_handoff_dashboard/` (deze README + de twee/drie HTML-bestanden) mee is als context of in de repo staat.

---

Je gaat een afgeronde design-refactor doorvoeren in de **echte codebase** van de SuperOffice Health Check (Hugo-gegenereerd dashboard: templates + CSS/JS). Dit is **geen nieuw scherm** en **geen kopieer-klus** — het is het app-breed doorvoeren van een kleur-, contrast-, tabel- en componentsysteem.

**Lees eerst, in deze volgorde, volledig:**
1. `design_handoff_dashboard/README.md` — de 7 kernbeslissingen (P1–P7), het complete token-systeem (basis + de twee `[data-surface]`-thema's), de `sheeted`-tabelbehandeling, en de sectie *App-brede uitrol* met de extra rampen en kruispunt-fixes.
2. `design_handoff_dashboard/Component reference - Zwevend mono.html` — de **visuele bron van waarheid** voor elk nieuw component. Open in een browser.
3. `design_handoff_dashboard/Company Overview - Neutral shell.html` — het volledige pagina-prototype (Company · Overview). Leidend voor die pagina; wissel thema linksonder naar **Zwevend mono**.

**De HTML-bestanden zijn design-referenties, geen productiecode.** Reproduceer de look in de bestaande Hugo/CSS-patronen — kopieer geen markup 1-op-1.

**Doel-shell = Zwevend mono** (getinte kaarten `#f2efe9` op een wit zwevend paneel, blended shell). De app migreert dus van witte kaarten + donkergroene sidebar naar mono.

**Hardste guardrails (niet schenden):**
- **P1 — twee gescheiden data-rampen.** Semantisch (good/mod/bad) alleen voor gezondheid/score; categorisch (teal `c1–c6` + neutraal `cn`) alleen voor verdelingen. Weight/Importance en User-levels krijgen hun **eigen** ramp (neutrale nadruk resp. level-ramp) — **nooit** groen/blauw lenen. "Dormant" overal neutraal.
- **Controls** (segmented/radio/slider/dropdown) mogen merk-groen gebruiken voor UI-state; P1 geldt alleen voor datavisualisatie.
- **Review-scaffolding niet shippen:** de annotatie-overlay ("Toon verbeteringen" + pins) en de stijl-schakelaar linksonder uit het prototype zijn review-tools — laat ze weg.
- Behoud de **WCAG AA**-contrastverhoudingen uit de README.

**Werkwijze:**
1. Bevestig kort je begrip van P1–P7 en de doel-shell vóór je begint.
2. Zet eerst het **token-fundament** neer (basis + `[data-surface="float-mono"]` + `.sheeted`-reset) op één centrale plek.
3. Doe daarna **één pagina volledig** (voorstel: Company · Overview, want die is 1-op-1 gespecificeerd) en laat het resultaat zien vóór je verder rolt.
4. Rol dan pagina voor pagina uit; per pagina de kruispunt-fixes meenemen (groene weight → nadruk-ramp, blauwe levels → level-ramp, paarse Settings-dropdowns → neutrale select, felle KPI-cijfers → gedempt `--bad-ink`, insights → dunne gedempte linker-accent).
5. Vraag bij twijfel of bij een grote refactor eerst even af; wijzig niet meer dan nodig.

Je kunt de live app niet zien — werk puur op de referenties. Begin met stap 1.

---

## Aandachtspunten die je zelf even moet checken in de repo
Deze kon ik vanuit de referenties niet vaststellen; laat je Claude ze vroeg verifiëren:
- **Waar leven de tokens nu** (CSS custom properties? SCSS-variabelen? een thema-partial?) — bepaalt waar het token-blok landt.
- **Sidebar**: mono blend't de sidebar in de shell (beige). De live app heeft een donkergroene sidebar. Bevestig of je die echt wilt opgeven of als merk-anker wilt houden (kleine afwijking van pure mono).
- **Grafiek-library**: donut/trend/bars in de live app — zelf-CSS/SVG of een chartlib? Kleuren moeten hoe dan ook uit de rampen komen.
