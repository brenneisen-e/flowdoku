/**
 * Die Start-Use-Cases — seit v1.2 die **echten Demos vom KI-Arbeitsplatz**.
 *
 * Vorher standen hier fuenf erfundene Agenten-Ideen aus einem Konzept-Deck
 * („Portfolio-Advisory-Agent", „Fraud-Investigation-Agent"). Sie waren alle
 * `Geplant` und zeigten auf nichts — eine Kachelwand, auf der jeder Klick ins
 * Leere geht, erklaert die Plattform nicht, sie entwertet sie. Nutzer-Ansage
 * 11.09.2026: „du solltest doch funktionierende Use Cases auf kiarbeitsplatz
 * nehmen und in der App hinterlegen".
 *
 * ## Woher die Liste kommt
 *
 * Aus **`kiarbeitsplatz/index.html`, Konstante `MODULE_CLUSTERS`** (Z. 374 ff.)
 * — 18 Kacheln in sechs Clustern, mit Label, Beschreibung, Ziel und Tags. Das
 * ist die maßgebliche Liste des Arbeitsplatzes selbst; eine zweite, hier neu
 * erfundene waere genau die Sorte Doppelung, die auseinanderlaeuft. Die langen
 * Beschreibungen stammen aus den READMEs und dem Quellcode der jeweiligen
 * Demo, nicht aus dem Kacheltext.
 *
 * ## Drei Korrekturen gegenueber der Vorlage
 *
 *  1. **Versicherungs-1×1** zeigt dort auf `Versicherungs1x1/index.html` —
 *     das ist der **Vite-Entwicklungs-Einstieg**, der `./src/main.jsx` laedt.
 *     Ein Browser fuehrt kein JSX aus, die Seite bleibt weiss. Richtig ist
 *     `Versicherungs1x1/dist/index.html`.
 *  2. **Bank der Zukunft** zeigt dort auf `BankDerZukunft/dist/index.html` —
 *     eine Huelle, die per `location.replace('/bank-der-zukunft-2035/dist/')`
 *     weiterleitet, und zwar mit ABSOLUTEM Pfad. Wir zeigen direkt auf das
 *     Ziel; die Huelle bricht, sobald das Repo nicht auf der Site-Wurzel liegt.
 *  3. **Abrechnungspruefung** und **Provisionsvalidierung** liegen auf eigenen
 *     Hosts ausserhalb des Arbeitsplatzes — ihre vollstaendige URL steht
 *     deshalb direkt im Eintrag, nicht hinter `PLATTFORM_BASIS`.
 *
 * ## Warum alle `aufrufArt: 'fenster'`
 *
 * Der ganze KI-Arbeitsplatz steht hinter einem Passwort
 * (`functions/_middleware.js` gated jede Datei; nachgemessen: die Startseite
 * antwortet mit **401**). Der erste Aufruf landet also auf einer Anmeldeseite,
 * danach traegt ein Cookie die Deeplinks. Eine Anmeldeseite in einem iframe
 * innerhalb von SharePoint ist der Fall, bei dem niemand mehr sagen kann, wo
 * die Eingabe hingeht — und Cookies in Drittanbieter-Kontexten sind die
 * naechste Fehlerquelle. Einbetten waere technisch erlaubt (`_headers` setzt
 * `X-Frame-Options: ALLOWALL`), ist hier aber die schlechtere Wahl.
 *
 * Ein Kurator kann das je Demo umstellen — die Voreinstellung ist nur die
 * Voreinstellung.
 */

import { UseCase } from '../types';

/**
 * Die Wurzel des KI-Arbeitsplatzes (Cloudflare Pages). Steht hier EINMAL:
 * Ein Umzug der Plattform ist sonst achtzehn Suchen-und-Ersetzen, von denen
 * das letzte vergessen wird.
 */
export const PLATTFORM_BASIS = 'https://kiarbeitsplatz.pages.dev/';

const url = (pfad: string): string => PLATTFORM_BASIS + pfad;

export const START_USE_CASES: Array<Partial<UseCase>> = [
  // ------------------------------------------------------------------ //
  // Vertrieb & Beratung                                                  //
  // ------------------------------------------------------------------ //
  {
    titel: 'KI-Vertriebsarbeitsplatz',
    bereich: 'Vertrieb & Beratung',
    kurzbeschreibung: 'Der Arbeitsplatz einer Beraterin: Tagesübersicht, Kundenliste, Kundenakte und Vertriebspotenziale in einer Oberfläche — mit KI-Assistent daneben.',
    beschreibung: '<p>Vier Ansichten greifen ineinander: <strong>Landing</strong> (Chat-Einstieg mit Schnellbefehlen), <strong>Übersicht</strong> (KPI-Kacheln plus sechs bis neun Reiter je Modus — Stammdaten, Neugeschäft, Bestand, Storno, Vergütung, Qualität, Kunden, Potenziale), <strong>Kundenakte</strong> (sieben bis acht Reiter von der Finanzübersicht über Open Finance bis Produktabschluss) und <strong>Cockpit</strong> (Vertriebssteuerung mit anklickbarer Deutschlandkarte, Vermittler-Ranking und Forecast).</p><p>Derselbe Arbeitsplatz lässt sich per Reiter zwischen <strong>Versicherung</strong>, <strong>Banking</strong> und <strong>Vermögensverwaltung</strong> umschalten — mitsamt Institut, Farbwelt, Rollenbezeichnung und KPI-Satz; je Modus stehen vier bis sechs weitere Institutsvarianten für White-Label-Vorführungen bereit. Die Daten sind seed-basiert erzeugt und damit bei jeder Vorführung identisch (45 Kunden Versicherung, 40 Banking, 12 Mandate, davon 19 handgebaute Personas). Zweisprachig, mit geführter Tour und Screen-Recorder.</p><p><em>Eine Einschränkung, die man vor dem Kundentermin kennen sollte:</em> Navigation und Schnellbefehle laufen ohne Backend. Nur <strong>freie</strong> Fragen an den Assistenten gehen an eine API; fehlt der Schlüssel, zeigt der Assistent eine Fehlermeldung und die Demo läuft weiter.</p>',
    salesRelevanz: 'Hoch', machbarkeit: 'Hoch', demoTauglichkeit: 'Hoch',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 10,
    schlagworte: ['Vertrieb', 'Beratung', 'Kundenakte', 'Cross-Selling', 'Cockpit', 'Versicherung', 'Banking', 'Asset Management'],
    ressourcen: { sourceCode: '', deployment: url('vertriebsarbeitsplatz/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },
  {
    titel: 'Vermittler-Scoring',
    bereich: 'Vertrieb & Beratung',
    kurzbeschreibung: 'Parametrisierbares Scoring-Modell für Versicherungsvermittler: Gewichte verstellen und sofort sehen, wie sich Ranking und Verteilung verschieben.',
    beschreibung: '<p>Links stehen die Parameter in sechs Abschnitten, rechts das Ergebnis: die Score-Verteilung über den Vermittlerbestand, Vertriebsstatistiken und eine Auswahl von Verbesserungsmaßnahmen mit Vorher-Nachher-Vergleich.</p><p>Der Kern der Vorführung ist die Gewichtung: Wer sie verändert, sieht die Verschiebung sofort — das macht ein Scoring-Modell verhandelbar, statt es als Black Box zu zeigen.</p>',
    salesRelevanz: 'Hoch', machbarkeit: 'Hoch', demoTauglichkeit: 'Mittel',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 20,
    schlagworte: ['Scoring', 'Vermittler', 'Risikobewertung', 'Simulation', 'Versicherung'],
    ressourcen: { sourceCode: '', deployment: url('Risikoscoring/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },
  {
    titel: 'Vergütungswerk',
    bereich: 'Vertrieb & Beratung',
    kurzbeschreibung: 'Zwei Werkzeuge auf einer Rechen-Engine: Was bewirkt eine Umstellung des Vergütungsmodells — und was heißt das je Institut in der Verhandlung?',
    beschreibung: '<p><strong>Makro:</strong> Links die Steuerung mit Beispielhaus-Auswahl, rechts Regler je Vergütungsbaustein, eine editierbare Bonusstaffel, Garantien und Zuschüsse. Daneben rechnen Wirkungstafeln sofort mit — Produktionszeilen Alt/Neu/Differenz, Staffelverteilung, Verlauf.</p><p><strong>Mikro:</strong> Je Sparkasse bzw. Vertriebspartner eine Institutsakte als Dashboard mit Stammdaten, gestapeltem Alt/Neu-Balken, Sparten-Säulen und Vergleichstabellen für Abschluss- und Folgeprovisionen.</p><p>Das tragende Prinzip: Makro ist <strong>exakt</strong> die Summe der Mikro-Ergebnisse, geprüft über eine Konsistenzfunktion. Drei Budget-Sichten (eingeschwungen, Barwert je Jahrgang, kassenwirksames Startjahr) werden immer gemeinsam gezeigt, die Gesamtampel ist stets das schlechteste der drei Signale. Marke, Farben und Sprachbausteine sind vollständig austauschbar.</p>',
    salesRelevanz: 'Hoch', machbarkeit: 'Mittel', demoTauglichkeit: 'Hoch',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 30,
    schlagworte: ['Vergütung', 'Provision', 'Simulation', 'White-Label', 'Verhandlung', 'Sparkasse'],
    ressourcen: { sourceCode: '', deployment: url('vertragsumstellung/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },
  {
    titel: 'KI-Kreditanalyse',
    bereich: 'Vertrieb & Beratung',
    kurzbeschreibung: 'Der komplette Firmenkredit-Prozess: von der Antragserfassung über KYC, Bilanzanalyse, Rating und Sicherheiten bis Entscheidung, Vertrag und Monitoring.',
    beschreibung: '<p>Vierzehn Ansichten hängen an einer Sidebar: Dashboard mit KPI-Karten und Kanban-Pipeline, Antragserfassung, KYC &amp; Compliance, §18-KWG-Analyse, Rating, Sicherheiten, Kreditvorlage, Entscheidung, Vertrag, Monitoring, Kundenantrag, Compliance, PDF-Upload und Kreditkonto-Neuanlage.</p><p>Hinterlegt sind Branchen-Benchmarks, Ratingskalen, Haircut-Sätze, eine Kompetenzmatrix und LMA-/Konsortialvertragsvorlagen; Muster-Verträge liegen als Assets bei. Ein angedocktes KI-Panel begleitet jeden Schritt.</p><p><em>Für die Vorführung angenehm:</em> Der Assistent fällt bei fehlendem API-Zugang automatisch auf vorbereitete Antworten zurück — die Demo funktioniert also auch ohne Schlüssel vollständig.</p>',
    salesRelevanz: 'Hoch', machbarkeit: 'Hoch', demoTauglichkeit: 'Hoch',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 40,
    schlagworte: ['Kredit', 'Firmenkunden', 'Rating', 'KYC', 'Sicherheiten', 'Monitoring', 'Banking'],
    ressourcen: { sourceCode: '', deployment: url('kreditanalyse/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },
  {
    titel: 'HR Pipeline',
    bereich: 'Vertrieb & Beratung',
    kurzbeschreibung: 'Recruiting aus drei Perspektiven in einer App: Bewerber-Portal, Recruiting-Dashboard und HR-Cockpit — mit KI-Analyse des Lebenslaufs gegen die Stellenanforderung.',
    beschreibung: '<p>Der Kopf schaltet zwischen <strong>Bewerber-Portal</strong>, <strong>Recruiting Dashboard</strong> und <strong>HR-Cockpit</strong>. Im Bewerberteil wird ein Profil eingereicht und analysiert — und die Analyse ist an die Quellstellen im Lebenslauf rückgebunden: Wer einen Befund anzweifelt, sieht die Stelle im Dokument, aus der er stammt.</p><p>Der Arbeitgeberteil führt Bewerberlisten und Stellen, das Cockpit zeigt die regionale Verteilung auf einer Deutschlandkarte. Rechts ist durchgängig ein KI-Assistent angedockt, mit getrennten Fragensets für Bewerber und Recruiter.</p>',
    salesRelevanz: 'Mittel', machbarkeit: 'Hoch', demoTauglichkeit: 'Hoch',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 50,
    schlagworte: ['Recruiting', 'Lebenslauf', 'Bewerbermanagement', 'Skill-Gap', 'HR'],
    ressourcen: { sourceCode: '', deployment: url('HRPipeline/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },

  // ------------------------------------------------------------------ //
  // Cockpits & Monitoring                                                //
  // ------------------------------------------------------------------ //
  {
    titel: 'VST-Cockpit',
    bereich: 'Cockpits & Monitoring',
    kurzbeschreibung: 'Neun-KPI-Cockpit der Vertriebssteuerung mit 12-Monats-Verlauf, Forecast und anklickbarer Deutschlandkarte nach Bundesländern.',
    beschreibung: '<p>Die Kennzahlen: Neugeschäftsvolumen, Bestandsvolumen, Stornoquote, NPS, Risikoscore, Combined Ratio, Gesamtergebnis, Underwriting-Qualität und Deckungsbeitrag. Gefiltert wird nach Jahr, Vertriebsweg (Ausschließlichkeit, Makler, Direktvertrieb, Banken) und Segment (Leben, Kranken, Schaden, Kfz); jede Kachel lässt sich vergrößern, ein Forecast zuschalten.</p><p>Dazu Vermittler-Ranking, Abschlussliste und die Bundesland-Karte mit Neugeschäft, Bestand, Storno und Vermittlerzahl je Land.</p><p><em>Unterschied zur gleichnamigen Ansicht im Vertriebsarbeitsplatz:</em> Diese Fassung erzeugt ihre Zahlen über einen Generator, der auf die letzten zwölf Monate ab dem aktuellen Datum rollt — die Demo veraltet also nicht.</p>',
    salesRelevanz: 'Hoch', machbarkeit: 'Hoch', demoTauglichkeit: 'Hoch',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 60,
    schlagworte: ['KPI', 'Cockpit', 'Vertriebssteuerung', 'Deutschlandkarte', 'Forecast', 'Storno', 'Versicherung'],
    ressourcen: { sourceCode: '', deployment: url('VSTCockpit/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },
  {
    titel: 'Active Equities',
    bereich: 'Cockpits & Monitoring',
    kurzbeschreibung: 'Der Morgen eines Portfolio Managers: AUM-Verlauf, Cashflows von gestern, Workflow-Status — und danach ein Dashboard, in dem Fonds-Workflows Schritt für Schritt abgearbeitet werden.',
    beschreibung: '<p>Der Ablauf ist dreistufig: Welcome-Screen (AUM der letzten fünf Tage, gestrige Cashflows, Workflow-Status, Highlights) → Ladeanimation → Dashboard.</p><p>Im Dashboard stehen Fonds wie „AI and Robotics Equity Fund" oder „Global Leaders Equity" mit AUM, Cash-Quote, Drift, Status (Trade required / PTC pending / SOD break) und offenen Orders. Jeder Workflow besteht aus Schritten, die teils automatisch laufen, teils eine Freigabe brauchen — „Wait for all cashflows", „Publish cash &amp; check portfolio", „Optimize portfolio", „Run PTC", „Approve PTC overrides", „Request 4-eye check", „Send orders". Ein Chatbot begleitet die Bearbeitung.</p><p><strong>Die Oberfläche ist englisch</strong> — als einzige der Demos.</p>',
    salesRelevanz: 'Mittel', machbarkeit: 'Hoch', demoTauglichkeit: 'Hoch',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 70,
    schlagworte: ['Portfolio Management', 'AUM', 'Pre-Trade Compliance', 'Workflow', 'Orders', 'Asset Management'],
    ressourcen: { sourceCode: '', deployment: url('asset-management/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },
  {
    titel: 'Banken-KPI',
    bereich: 'Cockpits & Monitoring',
    kurzbeschreibung: 'Monatlicher KPI-Bericht einer Bank über sieben Segmente — mit Ampelbewertung, Plan-/Ist-Abweichung und KI-Kommentar je Kennzahl, exportierbar als PowerPoint und PDF.',
    beschreibung: '<p>Umgeschaltet wird zwischen Konzern, Firmenkunden, Structured Finance, Immobilienkunden, Markets, Privatkunden und Treasury. Die Kennzahlen sind gruppiert: Neugeschäft und Neugeschäftsmarge, Volumina (Aktivvolumen/-marge, Einlagenvolumen/-marge) sowie Rentabilität und Produktivität (RWA, RWA-Produktivität, RoRaC, Cost/Income Ratio).</p><p>Jede Kachel zeigt einen 12-Monats-Verlauf und öffnet im Klick eine Detailansicht mit KI-Erläuterung; eine KPI-Matrix stellt die Segmente nebeneinander. Das Institut ist wählbar und wird gemerkt. Export als PowerPoint und über den Druckdialog als PDF.</p>',
    salesRelevanz: 'Hoch', machbarkeit: 'Hoch', demoTauglichkeit: 'Hoch',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 80,
    schlagworte: ['KPI', 'Banksteuerung', 'RWA', 'Cost/Income', 'Segmentbericht', 'PowerPoint-Export', 'Banking'],
    ressourcen: { sourceCode: '', deployment: url('BankenKPI/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },
  {
    titel: 'PMO Toolbox',
    bereich: 'Cockpits & Monitoring',
    kurzbeschreibung: 'Steuerungswerkzeug für ein Transformationsprogramm: Projekte und Workstreams, Programm-Timeline, Gantt-Chart der Meilensteine, Risiken und Entscheidungen.',
    beschreibung: '<p>Die Oberfläche führt Projekte je Workstream mit Statuserfassung, eine Programm-Timeline, ein Gantt-Chart der Meilensteine, ein Register „Risiken &amp; Herausforderungen", ein Entscheidungs-Management und ein Stichwortverzeichnis. Value-Map, KPI-Karten und ein Warnungs-Bereich kommen dazu; Details lassen sich in Modals öffnen und bearbeiten, Pakete exportieren.</p>',
    salesRelevanz: 'Mittel', machbarkeit: 'Hoch', demoTauglichkeit: 'Mittel',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 90,
    schlagworte: ['PMO', 'Transformation', 'Gantt', 'Meilensteine', 'Risiken', 'Entscheidungen'],
    ressourcen: { sourceCode: '', deployment: url('PMOToolbox/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },

  // ------------------------------------------------------------------ //
  // Backoffice-Prozesse                                                  //
  // ------------------------------------------------------------------ //
  {
    titel: 'Bestandsübertragung',
    bereich: 'Backoffice-Prozesse',
    kurzbeschreibung: 'Wie ein KI-Agent unstrukturierte Bestandsübertragungs-Mails samt Anhängen klassifiziert, Vollmachten erkennt und Multi-Verträge splittet — statt manueller Prüfung im Gruppenpostfach.',
    beschreibung: '<p>Elf Ansichten: Dashboard, Vorgänge, Makler, E-Mails, KI-Upload, Prüfstrecke, Simulation, Vorlagen sowie Architektur, Prozess und Projektplan.</p><p><strong>Der fachliche Hintergrund macht die Demo stark:</strong> Jährlich eine sechsstellige Zahl an Bestandsübertragungen, überwiegend als unstrukturierte E-Mails; heute laufen drei bis vier Prüfschritte manuell, die bestehende KI verarbeitet nur einfache Text-Mails ohne Anhänge. Zielbild ist ein KI-Agent früh in der Kette — Klassifikation, Vollmachtserkennung, Vertragssplitting, direkte JSON-Übergabe an Robotics. Durchlaufzeit von Tagen auf Minuten.</p><p>Vorgänge lassen sich als JSON exportieren und importieren; die Fachunterlagen (Wissensdokument, PDD, Prüfbot) liegen bei.</p>',
    salesRelevanz: 'Hoch', machbarkeit: 'Mittel', demoTauglichkeit: 'Hoch',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 100,
    schlagworte: ['Bestandsübertragung', 'Dokumentenverarbeitung', 'Vollmacht', 'Dunkelverarbeitung', 'RPA', 'Makler', 'Versicherung'],
    ressourcen: { sourceCode: '', deployment: url('Bestandsuebertragung/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },
  {
    titel: 'Abrechnungsprüfung',
    bereich: 'Backoffice-Prozesse',
    kurzbeschreibung: 'Abrechnungen prüfen und validieren.',
    beschreibung: '<p>Eigenständige Demo auf einem eigenen Host — <strong>nicht</strong> Teil des KI-Arbeitsplatzes und damit auch nicht hinter dessen Passwort. Der Quellcode liegt in einem anderen Repository.</p><p>Über den Kacheltext hinaus ist hier bewusst nichts beschrieben: Was die Demo im Einzelnen zeigt, sollte jemand ergänzen, der sie vorgeführt hat — eine ausgedachte Beschreibung wäre schlimmer als eine kurze.</p>',
    salesRelevanz: 'Mittel', machbarkeit: 'Hoch', demoTauglichkeit: 'Mittel',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 110,
    schlagworte: ['Abrechnung', 'Prüfung', 'Validierung', 'Backoffice', 'Versicherung'],
    ressourcen: { sourceCode: '', deployment: 'https://billingcheck.pages.dev/app', deploymentGuide: '', wiki: '', video: '' },
  },
  {
    titel: 'Provisionsvalidierung',
    bereich: 'Backoffice-Prozesse',
    kurzbeschreibung: 'Provisionen analysieren.',
    beschreibung: '<p>Eigenständige Demo auf einem eigenen Host — <strong>nicht</strong> Teil des KI-Arbeitsplatzes und damit auch nicht hinter dessen Passwort. Der Quellcode liegt in einem anderen Repository.</p><p>Wie bei der Abrechnungsprüfung steht hier absichtlich nur der Kacheltext: Die inhaltliche Beschreibung gehört von jemandem ergänzt, der die Demo kennt.</p>',
    salesRelevanz: 'Mittel', machbarkeit: 'Hoch', demoTauglichkeit: 'Mittel',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 120,
    schlagworte: ['Provision', 'Analyse', 'Validierung', 'Backoffice', 'Versicherung'],
    ressourcen: { sourceCode: '', deployment: 'https://provisions-analyzer.pages.dev/', deploymentGuide: '', wiki: '', video: '' },
  },
  {
    titel: 'Provisionsreklamation',
    bereich: 'Backoffice-Prozesse',
    kurzbeschreibung: 'Reklamation einer Provisionsabrechnung entlang dreier Rollen: Vermittler reicht ein, Backoffice prüft KI-gestützt, Admin verwaltet.',
    beschreibung: '<p>Drei Ansichten: <strong>Reklamation einreichen</strong> (Formular mit eigenem Monat/Jahr-Picker statt des nativen Feldes, damit die Monatsnamen garantiert deutsch sind), <strong>Backoffice</strong> (Prüfung der eingegangenen Fälle mit KI-Analyse und Antwortentwurf) und <strong>Admin</strong>.</p><p>Besonderheit für Termine ohne Netz: Es gibt einen zweiten Build als <strong>Offline-Fassung</strong>, die als Einzeldatei weitergegeben werden kann. Die Online-Fassung blendet dafür einen Download-Knopf ein.</p>',
    salesRelevanz: 'Mittel', machbarkeit: 'Hoch', demoTauglichkeit: 'Hoch',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 130,
    schlagworte: ['Provision', 'Reklamation', 'Backoffice', 'Antwortentwurf', 'Offline-Demo', 'Versicherung'],
    ressourcen: { sourceCode: '', deployment: url('provisionsreklamation/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },
  {
    titel: 'PDF-Editor',
    bereich: 'Backoffice-Prozesse',
    kurzbeschreibung: 'Text in bestehenden PDFs ändern, verschieben, löschen und ergänzen — mit der im PDF eingebetteten Originalschrift, sodass das Layout erhalten bleibt.',
    beschreibung: '<p>Ein eigener Font-Encoder baut aus Encoding, ToUnicode und Widths einen Encoder für die eingebettete Schrift; fehlt ein Zeichen in der eingebetteten Teilmenge, fällt der Block auf eine passende Standardschrift zurück — und der Nutzer bekommt einen Hinweis, statt dass es still anders aussieht.</p><p>Dazu ein <strong>Stapelmodus</strong>: Die Änderungen einer Vorlage werden über Seite, Position, Schriftgröße, Schrift und Text auf gleich aufgebaute PDFs übertragen und gesammelt als ZIP gespeichert; optional wird ein übergreifendes PDF (Anschreiben, AGB) an jedes Ergebnis angehängt.</p><p><strong>Alles läuft im Browser — die Datei verlässt den Rechner nicht.</strong> Das ist bei Kundendokumenten das eigentliche Verkaufsargument.</p>',
    salesRelevanz: 'Mittel', machbarkeit: 'Hoch', demoTauglichkeit: 'Hoch',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 140,
    schlagworte: ['PDF', 'Textbearbeitung', 'Schrifterhalt', 'Stapelverarbeitung', 'Serienbrief'],
    ressourcen: { sourceCode: '', deployment: url('pdfeditor/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },

  // ------------------------------------------------------------------ //
  // Risiko & IKS                                                         //
  // ------------------------------------------------------------------ //
  {
    titel: 'ControlHub',
    bereich: 'Risiko & IKS',
    kurzbeschreibung: 'Liest Prozessdokumente und Arbeitsanweisungen einer Bank ein und leitet daraus Risiken und Kontrollen ab — mit Analyse des Ist-IKS, Optimierungsvorschlägen und Soll-/Ist-Abgleich.',
    beschreibung: '<p>Fünf Bereiche: <strong>Navigator</strong> (Ableitung aus Dokumenten), <strong>Analyzer</strong>, <strong>Optimizer</strong>, <strong>Gap Analyzer</strong> (als Ausblick markiert) und <strong>Methodik &amp; Regelkreis</strong>.</p><p>Als Ausgangsmaterial liegen konkrete Beispieldokumente bereit — Prozessleitfaden Kreditvergabe Firmenkunden, Operations-Handbuch SEPA/SWIFT-Zahlungsverkehr, Kompetenzmatrix Kreditgeschäft, Sicherheiten-Bewertung, Zahlungsverkehr-Limits, Prozessbeschreibung Wertpapier-Settlement, Schaden-Operations-Handbuch. Hinterlegt sind ein Risikokatalog und mehrere Agenten-Definitionen; Ergebnisse lassen sich exportieren, ein KI-Chat steht dauerhaft rechts.</p>',
    salesRelevanz: 'Hoch', machbarkeit: 'Mittel', demoTauglichkeit: 'Hoch',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 150,
    schlagworte: ['IKS', 'Kontrollen', 'Risiko', 'Prozessdokumente', 'Gap-Analyse', 'Regelkreis', 'Banking'],
    ressourcen: { sourceCode: '', deployment: url('ControlHub/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },
  {
    titel: 'Collections Management',
    bereich: 'Risiko & IKS',
    kurzbeschreibung: 'Cockpit für das Forderungsmanagement mit KI-Tagesbriefing, Next-Best-Action-Queue und Agenten, die Vorschläge zur Freigabe vorlegen — plus Portfolio-Sicht und Stage-2/3-Monitoring.',
    beschreibung: '<p>Drei Sichten beantworten drei Fragen. <strong>Cockpit</strong> — „Was mache ich heute?": KI-Tagesbriefing, Next-Best-Action-Queue aus Wiedervorlagen und Agenten-Freigaben, Agenten-Status, Kompakt-KPIs. <strong>Portfolio</strong> — „Wie steht das Buch?": Portfolio-KPIs, Willingness/Ability-Matrix, neue Fälle, Zahlungseingänge, Kundenliste. <strong>Risiko</strong> — „Wohin bewegt es sich?": der Stage-Fluss 1 → 2 → 3 mit ECL, Migration und Risikotreibern bzw. NPL-Bestand, Trends und Maßnahmen.</p><p>Die Kundenakte ist zweistufig: Ein Klick öffnet den Schnellblick, „Akte öffnen" führt ins Vollprofil mit Stammdaten, Konten, Transaktionen, GuV/Haushalt, Open Finance, Kommunikation, Dokumenten, Workflow und einem KI-Dossier. Elf Akten sind voll ausgebaut; für weitere Fälle liefert die App automatisch eine Kurzakte, <strong>damit ein Klick in der Vorführung nie ins Leere läuft</strong>.</p>',
    salesRelevanz: 'Hoch', machbarkeit: 'Mittel', demoTauglichkeit: 'Hoch',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 160,
    schlagworte: ['Forderungsmanagement', 'NPL', 'Next Best Action', 'IFRS 9', 'ECL', 'KI-Agenten', 'Banking'],
    ressourcen: { sourceCode: '', deployment: url('Collections/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },

  // ------------------------------------------------------------------ //
  // Wissen                                                               //
  // ------------------------------------------------------------------ //
  {
    titel: 'Versicherungs-1×1',
    bereich: 'Wissen',
    kurzbeschreibung: 'Nachschlagewerk zum deutschen Versicherungsmarkt in neun Kapiteln — jedes mit Leitkennzahl, Diagrammen, drei Vertiefungsstufen und Quiz, dazu ein Glossar mit 153 Einträgen.',
    beschreibung: '<p>Die Kapitel: Der Markt im Überblick, Marktteilnehmer, Produktlandschaft, Vertriebswege, Vergütung und Provision, Regulatorik, Wertschöpfungskette, Zukunftsthemen, Verbände und Institutionen.</p><p>Jedes Kapitel öffnet mit einer Leitzahl — Kapitel 1 etwa „254 Mrd. € Beitragsvolumen des deutschen Versicherungsmarktes 2025, +6,6 % gegenüber 2024" — und drei Schnellzahlen (349 Versicherer, 501,5 Mio. Verträge, 1,9 Bio. € Kapitalanlagen). Der Inhalt ist <strong>dreistufig</strong> aufgebaut, sodass Einsteiger und Fachleute dieselbe Seite verschieden tief lesen; dazu Diagramme und je Kapitel ein Quiz mit Auflösungstext.</p>',
    salesRelevanz: 'Mittel', machbarkeit: 'Hoch', demoTauglichkeit: 'Mittel',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 170,
    schlagworte: ['Versicherungsmarkt', 'Grundlagenwissen', 'Regulatorik', 'Vertriebswege', 'Glossar', 'Quiz'],
    // Bewusst `dist/`: Die Kachel im KI-Arbeitsplatz zeigt auf
    // `Versicherungs1x1/index.html` — den Vite-Entwicklungs-Einstieg, der
    // `./src/main.jsx` laedt. Ein Browser fuehrt kein JSX aus; die Seite
    // bleibt weiss.
    ressourcen: { sourceCode: '', deployment: url('Versicherungs1x1/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },

  // ------------------------------------------------------------------ //
  // Workshops & Strategie                                                //
  // ------------------------------------------------------------------ //
  {
    titel: 'Bank der Zukunft',
    bereich: 'Workshops & Strategie',
    kurzbeschreibung: 'Ein KI-Chatbot „erzeugt" zum Auftakt das Strukturbild der Bank — danach ist das Bild eine begehbare Karte, die je Zone über Hypothesen bis in eine Live-Demo führt.',
    beschreibung: '<p>Der Ablauf: <strong>Intro</strong> (vorbereitete Frage, sichtbares Nachdenken, das Strukturbild baut sich auf) → <strong>Übersicht</strong> (Master-Bild füllt den Bildschirm, Zonen mit Hover-Glow) → <strong>Kameraflug</strong> in eine Zone → <strong>Kapitel</strong> (Bild links, Text rechts, Pager durch drei Hypothesen je Zone: Status quo ↔ Zielbild) → <strong>Absprung</strong>.</p><p>Neun Kapitel: Kunde, Vertrieb, Marktfolge, Corporate, IT, Compliance, Transformation, Projekt, Vorstand. Die Absprünge sind je Kapitel verschieden — eingebettete Live-Demos (unter anderem der Vertriebsarbeitsplatz selbst), ein Strategy Lab, ein Co-Creation-Canvas, ein Kredit-Dilemma. In derselben App stecken zwei weitere Werkzeuge: das <strong>EDEKABANK Holodeck 2031</strong> (fünf Steuerungsräume der KI-nativen Bank) und <strong>„Ein Tag in Banking Operations 2030"</strong>.</p>',
    salesRelevanz: 'Hoch', machbarkeit: 'Mittel', demoTauglichkeit: 'Hoch',
    status: 'Live', aufrufArt: 'fenster', reihenfolge: 180,
    schlagworte: ['Workshop', 'Zielbild', 'Strukturbild', 'Hypothesen', 'Vorstand', 'Transformation', 'Banking'],
    // Bewusst direkt auf die App: `BankDerZukunft/dist/index.html` ist nur
    // eine Huelle, die per ABSOLUTEM Pfad weiterleitet — sie bricht, sobald
    // das Repo nicht auf der Site-Wurzel liegt.
    ressourcen: { sourceCode: '', deployment: url('bank-der-zukunft-2035/dist/index.html'), deploymentGuide: '', wiki: '', video: '' },
  },
];
