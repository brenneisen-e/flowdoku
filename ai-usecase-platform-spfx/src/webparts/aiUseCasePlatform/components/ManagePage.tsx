/**
 * Use Case Studio — hier legen Use Case Organizer Use Cases an und pflegen sie.
 *
 * Der Pflegepfad ist genauso wichtig wie die Kachelwand: Die Demos entstehen
 * erst auf dem Hackathon, die Plattform muss sie danach aufnehmen koennen.
 * Wer hier einen Use Case anlegt, soll ihn in zwei Minuten fertig haben.
 *
 * Reihenfolge im Formular nach dem UI-Leitfaden: Pflicht → Optional → Fein.
 * Beschriftungen sind Fragen, Ja/Nein sind Schalter, Alternativen sind
 * Kacheln.
 *
 * v1.3: Das Kachelbild wird hochgeladen und zugeschnitten (vorher: eine
 * URL-Eingabe), die Betreuer sind eintragbar (vorher wurden sie angezeigt, aber
 * es gab kein Feld dafür), und das Speichern läuft über `saveUseCase` — die
 * Reihenfolge „Bild zuerst, altes Bild zuletzt" steht dort und nicht hier.
 */

import * as React from 'react';
import Modal from './Modal';
import ImageCropModal from './ImageCropModal';
import { cx, ensureDexUiStyles } from './dexUi';
import { Plus, Pencil, Trash2, Check, ImageIcon, X } from './Icons';
import { useUseCases } from '../context/UseCaseContext';
import { useLanguage } from '../context/LanguageContext';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useDialog } from '../context/DialogContext';
import { useHilfe } from '../context/HilfeContext';
import { UseCase, UseCaseStatus, Bewertung, AufrufArt, BildAenderung } from '../types';
import { START_USE_CASES } from '../data/startUseCases';

/** Die leeren Ressourcen als eigene Konstante: So braucht weder der
 *  Compiler ein Ausrufezeichen noch der Leser eine Annahme. */
const LEER_RESSOURCEN: UseCase['ressourcen'] = {
  sourceCode: '', deployment: '', deploymentGuide: '', wiki: '', video: '',
};

const LEER: Partial<UseCase> = {
  titel: '', kurzbeschreibung: '', beschreibung: '', bereich: '',
  status: 'Geplant', salesRelevanz: 'unbewertet', machbarkeit: 'unbewertet', demoTauglichkeit: 'unbewertet',
  aufrufArt: 'fenster',
  ressourcen: LEER_RESSOURCEN,
  bildUrl: '', reihenfolge: 100, betreuerEmails: [], betreuerNamen: [], schlagworte: [],
};

/** Eine Zeile des Betreuer-Editors. */
interface BetreuerZeile { name: string; email: string }

/** Größte Bilddatei, die der Zuschnitt annimmt — danach ist es ein 1280-px-JPEG. */
const MAX_BILD_MB = 15;

const EMAIL_FORMAT = /^[^\s@;]+@[^\s@;]+\.[^\s@;]+$/;

export default function ManagePage(props: { editId?: number }): React.ReactElement {
  ensureDexUiStyles();
  const { useCases, ladeStatus, letzterStatus, letzterFehler, reload, saveUseCase, remove } = useUseCases();
  const { t, isDe } = useLanguage();
  const { navigate } = useNavigation();
  const { isOrganizer } = useRoles();
  const { confirmDialog, showAlert } = useDialog();
  const { openKontakt } = useHilfe();

  const [entwurf, setEntwurf] = React.useState<Partial<UseCase> | null>(null);
  const [editId, setEditId] = React.useState<number | null>(null);
  const [speichert, setSpeichert] = React.useState(false);
  const [feinOffen, setFeinOffen] = React.useState(false);

  // Kachelbild: Was zum Speichern ansteht, getrennt vom Entwurf. Der Entwurf
  // geht als Zeile nach SharePoint; ein File gehört dort nicht hinein.
  const [bildDatei, setBildDatei] = React.useState<File | null>(null);
  const [bildVorschau, setBildVorschau] = React.useState('');
  const [bildEntfernt, setBildEntfernt] = React.useState(false);
  const [ursprungBild, setUrsprungBild] = React.useState('');
  const [cropQuelle, setCropQuelle] = React.useState('');
  const dateiRef = React.useRef<HTMLInputElement | null>(null);

  const [betreuer, setBetreuer] = React.useState<BetreuerZeile[]>([]);

  // Deep-Link aus der Detailseite: „Bearbeiten" oeffnet den Dialog direkt.
  const aufgerufenRef = React.useRef(false);
  React.useEffect(() => {
    if (aufgerufenRef.current || !props.editId || ladeStatus !== 'ok') return;
    const uc = useCases.filter(u => u.id === props.editId)[0];
    if (uc) { aufgerufenRef.current = true; oeffne(uc); }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.editId, ladeStatus, useCases]);

  if (!isOrganizer) {
    return (
      <div className="dex-ui-empty">
        <div className="dex-ui-empty-title">{t('Dafür fehlt dir die Berechtigung', 'You are not allowed to do this')}</div>
        <div className="dex-ui-empty-desc">
          {t('Use Cases anlegen und pflegen dürfen Use Case Organizer und Admins. Wenn du das brauchst, kannst du die Rolle anfragen.',
            'Use Case Organizers and admins may create and maintain use cases. If you need this, you can request the role.')}
        </div>
        <div className="dex-ui-inline dex-ui-empty-action" style={{ justifyContent: 'center', flexWrap: 'wrap', gap: 10 }}>
          <button type="button" className="btn btn-primary" onClick={() => openKontakt('organizer')}>
            {t('Organizer werden', 'Become an organizer')}
          </button>
          <button type="button" className="btn btn-secondary" onClick={() => navigate('start')}>
            {t('Zur Übersicht', 'Back to overview')}
          </button>
        </div>
      </div>
    );
  }

  function oeffne(uc?: UseCase): void {
    if (uc) {
      setEditId(uc.id);
      setEntwurf({ ...uc, ressourcen: { ...uc.ressourcen } });
      setBildVorschau(uc.bildUrl || '');
      setUrsprungBild(uc.bildUrl || '');
      // Namen und Adressen sind parallele Listen; fehlt ein Name, gilt die Adresse.
      setBetreuer(uc.betreuerEmails.map((email, i) => ({ email, name: uc.betreuerNamen[i] && uc.betreuerNamen[i] !== email ? uc.betreuerNamen[i] : '' })));
    } else {
      setEditId(null);
      setEntwurf({ ...LEER, ressourcen: { ...LEER_RESSOURCEN }, reihenfolge: (useCases.length + 1) * 10 });
      setBildVorschau('');
      setUrsprungBild('');
      setBetreuer([]);
    }
    setBildDatei(null);
    setBildEntfernt(false);
    setCropQuelle('');
    setFeinOffen(false);
  }

  function schliesse(): void {
    setEntwurf(null);
    setEditId(null);
    setBildDatei(null);
    setCropQuelle('');
  }

  const patch = (p: Partial<UseCase>): void => setEntwurf(prev => (prev ? { ...prev, ...p } : prev));
  const patchRes = (p: Partial<UseCase['ressourcen']>): void =>
    setEntwurf(prev => (prev ? { ...prev, ressourcen: { ...LEER_RESSOURCEN, ...prev.ressourcen, ...p } } : prev));

  /** Eine Bilddatei gewählt → erst prüfen, dann zum Zuschnitt. */
  function waehleDatei(e: React.ChangeEvent<HTMLInputElement>): void {
    const f = e.target.files && e.target.files[0];
    // Leeren, damit dieselbe Datei später noch einmal gewählt werden kann.
    e.target.value = '';
    if (!f) return;
    if (!/^image\/(png|jpe?g|webp)$/i.test(f.type)) {
      showAlert(t('Nimm bitte ein JPEG, PNG oder WebP.', 'Please use a JPEG, PNG or WebP.'), { variant: 'error' });
      return;
    }
    if (f.size > MAX_BILD_MB * 1024 * 1024) {
      showAlert(t(`Die Datei ist größer als ${MAX_BILD_MB} MB.`, `The file is larger than ${MAX_BILD_MB} MB.`), { variant: 'error' });
      return;
    }
    const r = new FileReader();
    r.onload = () => setCropQuelle(String(r.result || ''));
    r.onerror = () => showAlert(t('Die Datei ließ sich nicht lesen.', 'The file could not be read.'), { variant: 'error' });
    r.readAsDataURL(f);
  }

  function entferneBild(): void {
    setBildDatei(null);
    setBildVorschau('');
    // Ein noch nicht gespeichertes Bild wegzunehmen ändert nichts am Stand;
    // nur ein gespeichertes muss beim Speichern wirklich entfernt werden.
    setBildEntfernt(!!ursprungBild);
    patch({ bildUrl: '' });
  }

  async function speichern(): Promise<void> {
    if (!entwurf) return;
    const titel = (entwurf.titel || '').trim();
    if (!titel) { showAlert(t('Der Use Case braucht einen Titel — er steht auf der Kachel.', 'The use case needs a title — it goes on the tile.'), { variant: 'error' }); return; }
    // „Live" ohne Ziel waere eine Kachel, die nichts tut. Lieber hier
    // abfangen als den Nutzer im Kundentermin ins Leere klicken lassen.
    if (entwurf.status === 'Live' && !(entwurf.ressourcen?.deployment || '').trim()) {
      showAlert(t('„Live" heißt: aufrufbar. Dafür fehlt der Deployment-Link — trag ihn ein oder setz den Status auf „In Arbeit".',
        '"Live" means callable. The deployment link is missing — add it or set the status to "In progress".'), { variant: 'error' });
      return;
    }

    // Betreuer: Namen und Adressen sind PARALLELE Listen. `mapUseCase` filtert
    // leere Einträge je Liste einzeln — hätte eine Person einen Namen, aber
    // keine Adresse, verschöbe sich die Zuordnung aller folgenden. Deshalb
    // braucht jede Zeile eine Adresse; der Name ist freiwillig (dann gilt die
    // Adresse). Ein Semikolon würde die Liste zerschneiden — es wird ersetzt.
    const zeilen = betreuer.filter(b => b.email.trim() || b.name.trim());
    for (const z of zeilen) {
      if (!EMAIL_FORMAT.test(z.email.trim())) {
        showAlert(t(`Für den Betreuer „${z.name || z.email || '?'}" fehlt eine gültige E-Mail-Adresse.`, `The maintainer "${z.name || z.email || '?'}" needs a valid email address.`), { variant: 'error' });
        return;
      }
    }
    const betreuerEmails = zeilen.map(z => z.email.trim());
    const betreuerNamen = zeilen.map(z => (z.name.trim() || z.email.trim()).replace(/;/g, ','));

    const bild: BildAenderung = bildDatei
      ? { art: 'neu', datei: bildDatei }
      : bildEntfernt ? { art: 'entfernen' } : { art: 'unveraendert' };

    setSpeichert(true);
    try {
      const res = await saveUseCase(editId, { ...entwurf, betreuerEmails, betreuerNamen }, bild);
      if (!res.ok) {
        // Der Grund gehört in die Meldung: „Speichern hat nicht geklappt" lässt
        // raten, ob es an den Rechten, am Bild oder an der Verbindung lag.
        showAlert(
          res.grund === 'bild'
            ? t('Das Bild ist nicht angekommen — es wurde nichts geändert. Versuch es noch einmal oder wähl ein kleineres Bild.', 'The image did not arrive — nothing was changed. Try again or choose a smaller image.')
            : res.grund === 'rechte'
              ? t('Dafür fehlt dir die Berechtigung.', 'You are not allowed to do this.')
              : t('Das Speichern hat nicht geklappt. Versuch es noch einmal.', 'Saving did not work. Please try again.'),
          { variant: 'error' },
        );
        return;
      }
      schliesse();
      if (res.bildFehler) {
        // Gespeichert, aber ohne Bild — ehrlich sagen, nicht „erledigt".
        showAlert(t('Der Use Case ist gespeichert, aber das Bild ist nicht angekommen. Öffne ihn noch einmal und wähl das Bild erneut.',
          'The use case is saved, but the image did not arrive. Open it again and choose the image again.'), { variant: 'error' });
      }
    } finally {
      setSpeichert(false);
    }
  }

  async function loeschen(uc: UseCase): Promise<void> {
    const ja = await confirmDialog(
      t(`„${uc.titel}" löschen? Die Kachel verschwindet für alle. Ein Admin der Site kann sie 93 Tage lang aus dem Papierkorb zurückholen.`,
        `Delete "${uc.titel}"? The tile disappears for everyone. A site admin can restore it from the recycle bin for 93 days.`),
      { confirmLabel: t('Löschen', 'Delete'), danger: true },
    );
    if (!ja) return;
    if (!(await remove(uc.id))) {
      showAlert(t('Das Löschen hat nicht geklappt.', 'Deleting did not work.'), { variant: 'error' });
    }
  }

  /**
   * Den Startbestand anlegen — nacheinander, nicht parallel.
   *
   * Fuenf gleichzeitige POSTs gegen dieselbe Liste sind genau das Muster,
   * das SharePoint drosselt; in DEX hat ein `Promise.all` ueber alle Personen
   * unbemerkt Verweise liegen lassen (v29.2). Nacheinander dauert zwei
   * Sekunden laenger und geht durch.
   */
  async function startbestand(): Promise<void> {
    setSpeichert(true);
    let angelegt = 0;
    try {
      for (const uc of START_USE_CASES) {
        // eslint-disable-next-line no-await-in-loop
        const r = await saveUseCase(null, uc, { art: 'unveraendert' });
        if (r.ok) angelegt++;
      }
    } finally {
      setSpeichert(false);
    }
    if (angelegt === START_USE_CASES.length) {
      showAlert(t(`${angelegt} Use Cases angelegt. Trag jetzt die Links nach, sobald die Demos stehen.`,
        `${angelegt} use cases created. Add the links once the demos exist.`), { variant: 'success' });
    } else {
      // Teilerfolg BENENNEN. „Angelegt" bei drei von fuenf waere die Sorte
      // Meldung, der man beim naechsten Mal nicht mehr glaubt.
      showAlert(t(`Nur ${angelegt} von ${START_USE_CASES.length} Use Cases konnten angelegt werden. Der Rest fehlt — bitte noch einmal versuchen.`,
        `Only ${angelegt} of ${START_USE_CASES.length} use cases could be created. The rest are missing — please try again.`), { variant: 'error' });
    }
  }

  const STATUS: Array<{ w: UseCaseStatus; de: string; en: string; hilfe: string }> = [
    { w: 'Geplant', de: 'Geplant', en: 'Planned', hilfe: t('Sichtbar, aber ohne Demo.', 'Visible, but no demo.') },
    { w: 'InArbeit', de: 'In Arbeit', en: 'In progress', hilfe: t('Wird gerade gebaut.', 'Being built.') },
    { w: 'Live', de: 'Live', en: 'Live', hilfe: t('Aufrufbar. Braucht einen Deployment-Link.', 'Callable. Needs a deployment link.') },
    { w: 'Archiviert', de: 'Archiviert', en: 'Archived', hilfe: t('Nur noch für Organizer sichtbar.', 'Visible to organizers only.') },
  ];
  const BEWERTUNGEN: Bewertung[] = ['Hoch', 'Mittel', 'Niedrig', 'unbewertet'];

  return (
    <div>
      <div className="dex-ui-page-head">
        <div>
          <h1 className="dex-ui-page-head-title">Use Case Studio</h1>
          <p className="dex-ui-page-head-meta">
            {ladeStatus === 'ok'
              ? `${useCases.length} ${t('Einträge', 'entries')} · ${useCases.filter(u => u.status === 'Live').length} ${t('aufrufbar', 'callable')}`
              : ladeStatus === 'laedt' ? t('Wird geladen …', 'Loading …') : t('Nicht lesbar', 'Not readable')}
          </p>
        </div>
        <div className="dex-ui-page-head-actions">
          <button type="button" className="dex-ui-textbtn" onClick={() => navigate('protokoll')}>
            {t('Protokoll', 'Log')}
          </button>
          {/* Bei einem Lesefehler kennt niemand den Bestand — ein neuer Eintrag könnte
              eine Dublette sein. Erst wieder anbieten, wenn die Liste gelesen ist. */}
          <button type="button" className="btn btn-primary" disabled={ladeStatus !== 'ok'} onClick={() => oeffne()}>
            <Plus size={15} /> {t('Neuer Use Case', 'New use case')}
          </button>
        </div>
      </div>

      {/* Drei Zustände, drei Aussagen. Bis v1.2 zeigte diese Seite bei einem Lesefehler
          „0 Einträge · Noch nichts angelegt" und bot an, die Start-Use-Cases anzulegen —
          auf einer bereits befüllten Liste wären das Dubletten gewesen. Ein Lesefehler ist
          keine leere Liste (Sichtprüfung mit dem Harness, 29.09.2026). */}
      {ladeStatus === 'laedt' && (
        <div className="dex-ui-empty" role="status" aria-live="polite">
          <div className="dex-ui-progress dex-ui-progress--indeterminate"><div className="dex-ui-progress-bar" /></div>
          <div className="dex-ui-empty-desc" style={{ marginTop: 12 }}>{t('Use Cases werden geladen …', 'Loading use cases …')}</div>
        </div>
      )}

      {ladeStatus === 'fehler' && (
        <div className="dex-ui-callout dex-ui-callout--danger" role="alert" style={{ marginBottom: 16 }}>
          <span>
            <strong>{t('Die Use Cases konnten nicht gelesen werden.', 'The use cases could not be read.')}</strong>
            <br />
            {letzterStatus === 403
              ? t('Dir fehlt das Leserecht auf der Liste.', 'You do not have read access to the list.')
              : t('Das ist ein Lesefehler, keine leere Liste — es wurde nichts angelegt und nichts gelöscht. Versuch es gleich noch einmal.', 'This is a read error, not an empty list — nothing was created or deleted. Please try again shortly.')}
            {letzterStatus > 0 && <span className="dex-ui-muted"> (HTTP {letzterStatus})</span>}
            {letzterFehler && (<><br /><span className="dex-ui-muted" style={{ fontSize: '0.78rem' }}>{letzterFehler}</span></>)}
            <br />
            <button type="button" className="dex-ui-textbtn" style={{ marginTop: 8 }} onClick={() => { void reload(); }}>{t('Erneut versuchen', 'Try again')}</button>
          </span>
        </div>
      )}

      {ladeStatus === 'ok' && useCases.length === 0 ? (
        <div className="dex-ui-empty">
          <div className="dex-ui-empty-title">{t('Noch nichts angelegt', 'Nothing here yet')}</div>
          <div className="dex-ui-empty-desc">
            {t('Leg die Use Cases schon vor dem Hackathon an — die Teams tragen ihre Links später selbst nach.',
              'Create the use cases before the hackathon — the teams add their links later.')}
          </div>
          <button type="button" className="dex-ui-empty-action" onClick={() => oeffne()}>
            {t('Ersten Use Case anlegen', 'Create the first use case')}
          </button>
          {/* Der Startbestand. Nur solange die Liste leer ist — danach waere
              der Knopf ein Weg, versehentlich Dubletten anzulegen. */}
          <button type="button" className="dex-ui-textbtn" style={{ marginTop: 10 }} disabled={speichert} onClick={() => { void startbestand(); }}>
            {speichert
              ? t('Wird angelegt …', 'Creating …')
              : t('Die Start-Use-Cases anlegen', 'Create the starter use cases')}
          </button>
        </div>
      ) : useCases.length > 0 && (
        <div className="dex-ui-stack">
          {useCases.map(uc => (
            <div key={uc.id} className="dex-ui-row dex-ui-row--bordered">
              <span className="dex-ui-row-main">
                <span className="dex-ui-row-title dex-ui-row-title--wrap">
                  {uc.titel}
                  <span className={cx('dex-ui-pill', 'dex-ui-pill--sm', 'dex-ui-pill--trail', uc.status === 'Live' ? 'dex-ui-pill--green' : 'dex-ui-pill--gray')}>
                    {uc.status === 'InArbeit' ? t('In Arbeit', 'In progress') : uc.status}
                  </span>
                </span>
                <span className="dex-ui-row-sub">
                  {uc.bereich || t('kein Bereich', 'no area')}
                  {uc.ressourcen.deployment ? ` · ${t('Link hinterlegt', 'link stored')}` : ` · ${t('kein Deployment-Link', 'no deployment link')}`}
                </span>
              </span>
              <span className="dex-ui-row-actions">
                <button type="button" className="dex-ui-iconbtn" onClick={() => oeffne(uc)} title={t('Bearbeiten', 'Edit')} aria-label={t('Bearbeiten', 'Edit')}>
                  <Pencil size={15} />
                </button>
                <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" onClick={() => { void loeschen(uc); }} title={t('Löschen', 'Delete')} aria-label={t('Löschen', 'Delete')}>
                  <Trash2 size={15} />
                </button>
              </span>
            </div>
          ))}
        </div>
      )}

      {entwurf && (
        <Modal
          open
          onClose={schliesse}
          // Solange der Zuschnitt offen ist, bleibt dieser Dialog stehen: Beide
          // hören auf Escape, und ein Druck darauf würde sonst BEIDE schließen
          // und den Entwurf verwerfen.
          dismissable={!speichert && !cropQuelle}
          // Kein Schließen per Klick daneben: Wer einen Text markiert und die
          // Maus einen Millimeter neben der Karte loslässt, verliert sonst
          // alles Getippte (DEX v30.51).
          backdropClose={false}
          maxWidth={640}
          ariaLabel={editId ? t('Use Case bearbeiten', 'Edit use case') : t('Neuer Use Case', 'New use case')}
          title={editId ? t('Use Case bearbeiten', 'Edit use case') : t('Neuer Use Case', 'New use case')}
          subtitle={t('Titel und Kurzbeschreibung stehen auf der Kachel.', 'Title and short description appear on the tile.')}
          footer={<>
            <button type="button" className="btn btn-secondary" disabled={speichert} onClick={schliesse}>{t('Abbrechen', 'Cancel')}</button>
            <button type="button" className="btn btn-primary" disabled={speichert} onClick={() => { void speichern(); }}>
              {speichert ? t('Speichert …', 'Saving …') : t('Speichern', 'Save')}
            </button>
          </>}
        >
          {/* PFLICHT */}
          <div className="dex-ui-section" style={{ margin: 0 }}>
            <div className="dex-ui-section-title">{t('Was ist das?', 'What is it?')}</div>

            <label className="dex-ui-field">
              <span className="dex-ui-label">{t('Wie heißt der Use Case?', 'What is the use case called?')} <span className="dex-ui-label-required">*</span></span>
              <input id="uc-titel" type="text" className="dex-ui-input" value={entwurf.titel || ''} onChange={e => patch({ titel: e.target.value })} />
            </label>

            <label className="dex-ui-field">
              <span className="dex-ui-label">{t('Was macht die Demo?', 'What does the demo do?')}</span>
              <textarea id="uc-kurz" className="dex-ui-textarea" rows={3} value={entwurf.kurzbeschreibung || ''} onChange={e => patch({ kurzbeschreibung: e.target.value })} />
              <span className="dex-ui-help">{t('Ein bis zwei Sätze. Diese Zeilen stehen auf der Kachel.', 'One or two sentences. These lines go on the tile.')}</span>
            </label>

            <label className="dex-ui-field">
              <span className="dex-ui-label">{t('Zu welchem Bereich gehört er?', 'Which area does it belong to?')}</span>
              <input id="uc-bereich" type="text" className="dex-ui-input" list="uc-bereiche" value={entwurf.bereich || ''} onChange={e => patch({ bereich: e.target.value })} placeholder={t('z.B. Investment / Private Banking', 'e.g. Investment / Private Banking')} />
              <datalist id="uc-bereiche">
                {Array.from(new Set(useCases.map(u => u.bereich).filter(Boolean))).map(b => <option key={b} value={b} />)}
              </datalist>
            </label>

            <div className="dex-ui-field">
              <span className="dex-ui-label">{t('Wie weit ist er?', 'How far along is it?')}</span>
              <div style={{ display: 'grid', gap: 8 }}>
                {STATUS.map(s => (
                  <button key={s.w} type="button" className={cx('dex-ui-choice', entwurf.status === s.w && 'is-active')} style={{ padding: '10px 12px', alignItems: 'center' }} aria-pressed={entwurf.status === s.w} onClick={() => patch({ status: s.w })}>
                    <span className="dex-ui-choice-body" style={{ display: 'block' }}>
                      <span className="dex-ui-choice-title" style={{ display: 'block' }}>{isDe ? s.de : s.en}</span>
                      <span className="dex-ui-choice-desc" style={{ display: 'block' }}>{s.hilfe}</span>
                    </span>
                    <span className="dex-ui-choice-check">{entwurf.status === s.w && <Check size={12} />}</span>
                  </button>
                ))}
              </div>
            </div>

            {/* Das Kachelbild. Optional, aber sichtbar: Es ist das, was die
                Kachel von zwanzig anderen unterscheidet. */}
            <div className="dex-ui-field">
              <span className="dex-ui-label">{t('Bild für die Kachel', 'Tile image')} <span className="dex-ui-label-optional">{t('optional', 'optional')}</span></span>
              <div style={{ display: 'flex', gap: 14, alignItems: 'center', flexWrap: 'wrap' }}>
                <span
                  aria-hidden="true"
                  style={{
                    display: 'block', width: 176, aspectRatio: '16 / 9', flexShrink: 0, borderRadius: 10,
                    border: '1px solid var(--dex-gray-200, #e8e8e8)', position: 'relative', overflow: 'hidden',
                    background: bildVorschau
                      ? `center/cover no-repeat url("${bildVorschau.replace(/"/g, '%22')}")`
                      : 'linear-gradient(135deg, rgba(134,188,37,0.16), rgba(134,188,37,0.05))',
                  }}
                >
                  {!bildVorschau && (
                    <span style={{ position: 'absolute', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--dex-green-dark, #6b9a1e)' }}>
                      <ImageIcon size={26} />
                    </span>
                  )}
                </span>
                <span style={{ display: 'flex', flexDirection: 'column', gap: 6, alignItems: 'flex-start' }}>
                  <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => dateiRef.current && dateiRef.current.click()}>
                    <ImageIcon size={15} /> {bildVorschau ? t('Bild ändern', 'Change image') : t('Bild wählen', 'Choose image')}
                  </button>
                  {bildVorschau && (
                    <button type="button" className="dex-ui-textbtn dex-ui-textbtn--danger" onClick={entferneBild}>
                      <X size={13} /> {t('Bild entfernen', 'Remove image')}
                    </button>
                  )}
                </span>
                <input ref={dateiRef} type="file" accept="image/png,image/jpeg,image/webp" onChange={waehleDatei} style={{ display: 'none' }} />
              </div>
              <span className="dex-ui-help">
                {bildDatei
                  ? t('Das neue Bild wird beim Speichern hochgeladen.', 'The new image is uploaded when you save.')
                  : t('Wird auf das Format der Kachel (16:9) zugeschnitten. Ohne Bild zeigt die Kachel das Kürzel des Titels.', 'Cropped to the tile format (16:9). Without an image the tile shows the title initials.')}
              </span>
            </div>
          </div>

          {/* OPTIONAL — die Links */}
          <div className="dex-ui-section" style={{ margin: 0 }}>
            <div className="dex-ui-section-title">{t('Ressourcen', 'Resources')}</div>
            <label className="dex-ui-field">
              <span className="dex-ui-label">{t('Wo läuft die Demo?', 'Where does the demo run?')}</span>
              <input id="uc-deploy" type="url" className="dex-ui-input" value={entwurf.ressourcen?.deployment || ''} onChange={e => patchRes({ deployment: e.target.value })} placeholder="https://…" />
              <span className="dex-ui-help">{t('Das Ziel des Kachel-Klicks. Ohne diesen Link kann der Status nicht „Live" sein.', 'The target of the tile click. Without it the status cannot be "Live".')}</span>
            </label>
            {([
              { k: 'sourceCode' as const, label: t('Source Code (GitHub)', 'Source code (GitHub)') },
              { k: 'deploymentGuide' as const, label: t('Deployment-Guide', 'Deployment guide') },
              { k: 'wiki' as const, label: t('Wiki / Doku', 'Wiki / docs') },
              { k: 'video' as const, label: t('Use-Case-Video', 'Use case video') },
            ]).map(f => (
              <label key={f.k} className="dex-ui-field">
                <span className="dex-ui-label">{f.label} <span className="dex-ui-label-optional">{t('optional', 'optional')}</span></span>
                <input id={`uc-${f.k}`} type="url" className="dex-ui-input" value={entwurf.ressourcen?.[f.k] || ''} onChange={e => patchRes({ [f.k]: e.target.value })} placeholder="https://…" />
              </label>
            ))}
          </div>

          {/* FEIN — im Aufklapper, weil es die meisten nie brauchen */}
          <div>
            <button type="button" className={cx('dex-ui-disclosure', feinOffen && 'is-open')} onClick={() => setFeinOffen(o => !o)} aria-expanded={feinOffen}>
              <span className="dex-ui-disclosure-chevron">›</span>
              {t('Betreuer, Bewertung, Aufruf und Reihenfolge', 'Maintainers, assessment, launch and order')}
              <span className="dex-ui-disclosure-count">{t('optional', 'optional')}</span>
            </button>
            {feinOffen && (
              <div className="dex-ui-disclosure-body">
                {/* Wer betreut die Demo? Anzeige und Ansprechpartner — KEINE
                    Rechte. Ob ein Team seinen Eintrag selbst pflegen darf, ist
                    eine eigene Entscheidung (Zeilen-Sicherheit auf der Liste). */}
                <div className="dex-ui-field">
                  <span className="dex-ui-label">{t('Wer betreut die Demo?', 'Who maintains the demo?')} <span className="dex-ui-label-optional">{t('optional', 'optional')}</span></span>
                  <div style={{ display: 'grid', gap: 8 }}>
                    {betreuer.map((b, i) => (
                      <div key={i} style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                        <input
                          type="text"
                          className="dex-ui-input"
                          style={{ flex: '1 1 0', minWidth: 0 }}
                          value={b.name}
                          placeholder={t('Name (optional)', 'Name (optional)')}
                          aria-label={t('Name des Betreuers', 'Maintainer name')}
                          onChange={e => setBetreuer(prev => prev.map((x, j) => (j === i ? { ...x, name: e.target.value } : x)))}
                        />
                        <input
                          type="email"
                          className="dex-ui-input"
                          style={{ flex: '1.4 1 0', minWidth: 0 }}
                          value={b.email}
                          placeholder="vorname.nachname@deloitte.de"
                          aria-label={t('E-Mail-Adresse des Betreuers', 'Maintainer email address')}
                          onChange={e => setBetreuer(prev => prev.map((x, j) => (j === i ? { ...x, email: e.target.value } : x)))}
                        />
                        <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" onClick={() => setBetreuer(prev => prev.filter((_, j) => j !== i))} title={t('Entfernen', 'Remove')} aria-label={t('Betreuer entfernen', 'Remove maintainer')}>
                          <Trash2 size={15} />
                        </button>
                      </div>
                    ))}
                    <span>
                      <button type="button" className="dex-ui-textbtn" onClick={() => setBetreuer(prev => [...prev, { name: '', email: '' }])}>
                        <Plus size={14} /> {t('Betreuer hinzufügen', 'Add maintainer')}
                      </button>
                    </span>
                  </div>
                  <span className="dex-ui-help">
                    {t('Sie stehen auf der Detailseite als Ansprechpartner. Dadurch bekommen sie keine zusätzlichen Rechte.', 'They appear on the detail page as contacts. This does not give them any extra rights.')}
                  </span>
                </div>

                {([
                  { k: 'salesRelevanz' as const, label: t('Sales-Relevanz', 'Sales relevance') },
                  { k: 'machbarkeit' as const, label: t('Machbarkeit', 'Feasibility') },
                  { k: 'demoTauglichkeit' as const, label: t('Demo-Tauglichkeit', 'Demo readiness') },
                ]).map(dim => (
                  <label key={dim.k} className="dex-ui-field">
                    <span className="dex-ui-label">{dim.label}</span>
                    <select id={`uc-${dim.k}`} className="dex-ui-select" value={(entwurf[dim.k] as Bewertung) || 'unbewertet'} onChange={e => patch({ [dim.k]: e.target.value as Bewertung })}>
                      {BEWERTUNGEN.map(b => (
                        <option key={b} value={b}>{b === 'unbewertet' ? t('noch nicht bewertet', 'not assessed yet') : b}</option>
                      ))}
                    </select>
                  </label>
                ))}

                <div className="dex-ui-field">
                  <span className="dex-ui-label">{t('Wie soll die Demo öffnen?', 'How should the demo open?')}</span>
                  <div style={{ display: 'grid', gap: 8 }}>
                    {([
                      { w: 'fenster' as AufrufArt, titel: t('In neuem Tab', 'In a new tab'), desc: t('Funktioniert immer — auch bei Kamera, Mikrofon oder Zwischenablage.', 'Always works — including camera, microphone or clipboard.') },
                      { w: 'eingebettet' as AufrufArt, titel: t('Eingebettet auf der Seite', 'Embedded on the page'), desc: t('Sieht geschlossener aus. Scheitert, wenn die Demo Geräterechte braucht oder das Einbetten verbietet.', 'Looks more integrated. Fails if the demo needs device access or forbids embedding.') },
                    ]).map(a => (
                      <button key={a.w} type="button" className={cx('dex-ui-choice', entwurf.aufrufArt === a.w && 'is-active')} style={{ padding: '10px 12px', alignItems: 'center' }} aria-pressed={entwurf.aufrufArt === a.w} onClick={() => patch({ aufrufArt: a.w })}>
                        <span className="dex-ui-choice-body" style={{ display: 'block' }}>
                          <span className="dex-ui-choice-title" style={{ display: 'block' }}>{a.titel}</span>
                          <span className="dex-ui-choice-desc" style={{ display: 'block' }}>{a.desc}</span>
                        </span>
                        <span className="dex-ui-choice-check">{entwurf.aufrufArt === a.w && <Check size={12} />}</span>
                      </button>
                    ))}
                  </div>
                </div>

                <label className="dex-ui-field">
                  <span className="dex-ui-label">{t('Reihenfolge', 'Order')}</span>
                  <input id="uc-reihenfolge" type="number" className="dex-ui-input" value={entwurf.reihenfolge ?? 100} onChange={e => patch({ reihenfolge: parseInt(e.target.value, 10) || 0 })} />
                  <span className="dex-ui-help">{t('Kleiner steht weiter vorne.', 'Lower numbers come first.')}</span>
                </label>

                <label className="dex-ui-field">
                  <span className="dex-ui-label">{t('Schlagworte', 'Keywords')} <span className="dex-ui-label-optional">{t('optional', 'optional')}</span></span>
                  <input id="uc-schlagworte" type="text" className="dex-ui-input" value={(entwurf.schlagworte || []).join('; ')} onChange={e => patch({ schlagworte: e.target.value.split(';').map(s => s.trim()).filter(Boolean) })} placeholder={t('mit Semikolon trennen', 'separate with semicolons')} />
                  <span className="dex-ui-help">{t('Zusätzliche Wörter, über die man den Use Case findet.', 'Extra words that help find the use case.')}</span>
                </label>
              </div>
            )}
          </div>
        </Modal>
      )}

      <ImageCropModal
        open={!!cropQuelle}
        src={cropQuelle}
        onClose={() => setCropQuelle('')}
        onApply={(datei, vorschau) => {
          setBildDatei(datei);
          setBildVorschau(vorschau);
          setBildEntfernt(false);
          setCropQuelle('');
        }}
      />
    </div>
  );
}
