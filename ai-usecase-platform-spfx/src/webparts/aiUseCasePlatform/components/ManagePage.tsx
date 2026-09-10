/**
 * Use Cases pflegen — die Kurator-Seite.
 *
 * Der Pflegepfad ist genauso wichtig wie die Kachelwand: Die Demos entstehen
 * erst auf dem Hackathon, die Plattform muss sie danach aufnehmen koennen.
 * Wer hier einen Use Case anlegt, soll ihn in zwei Minuten fertig haben.
 *
 * Reihenfolge im Formular nach dem UI-Leitfaden: Pflicht → Optional → Fein.
 * Beschriftungen sind Fragen, Ja/Nein sind Schalter, Alternativen sind
 * Kacheln.
 */

import * as React from 'react';
import Modal from './Modal';
import { cx, ensureDexUiStyles } from './dexUi';
import { Plus, Pencil, Trash2, Check } from './Icons';
import { useUseCases } from '../context/UseCaseContext';
import { useLanguage } from '../context/LanguageContext';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useDialog } from '../context/DialogContext';
import { UseCase, UseCaseStatus, Bewertung, AufrufArt } from '../types';
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

export default function ManagePage(props: { editId?: number }): React.ReactElement {
  ensureDexUiStyles();
  const { useCases, ladeStatus, create, update, remove } = useUseCases();
  const { t, isDe } = useLanguage();
  const { navigate } = useNavigation();
  const { isKurator } = useRoles();
  const { confirmDialog, showAlert } = useDialog();

  const [entwurf, setEntwurf] = React.useState<Partial<UseCase> | null>(null);
  const [editId, setEditId] = React.useState<number | null>(null);
  const [speichert, setSpeichert] = React.useState(false);
  const [feinOffen, setFeinOffen] = React.useState(false);

  // Deep-Link aus der Detailseite: „Bearbeiten" oeffnet den Dialog direkt.
  const aufgerufenRef = React.useRef(false);
  React.useEffect(() => {
    if (aufgerufenRef.current || !props.editId || ladeStatus !== 'ok') return;
    const uc = useCases.filter(u => u.id === props.editId)[0];
    if (uc) { aufgerufenRef.current = true; oeffne(uc); }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.editId, ladeStatus, useCases]);

  if (!isKurator) {
    return (
      <div className="dex-ui-empty">
        <div className="dex-ui-empty-title">{t('Dafür fehlt dir die Berechtigung', 'You are not allowed to do this')}</div>
        <div className="dex-ui-empty-desc">
          {t('Use Cases pflegen dürfen Kuratoren und Admins. Wenn du das brauchst, melde dich bei einem Admin der Plattform.',
            'Curators and admins may manage use cases. If you need this, contact a platform admin.')}
        </div>
        <button type="button" className="dex-ui-empty-action" onClick={() => navigate('start')}>
          {t('Zur Übersicht', 'Back to overview')}
        </button>
      </div>
    );
  }

  function oeffne(uc?: UseCase): void {
    if (uc) {
      setEditId(uc.id);
      setEntwurf({ ...uc, ressourcen: { ...uc.ressourcen } });
    } else {
      setEditId(null);
      setEntwurf({ ...LEER, ressourcen: { ...LEER_RESSOURCEN }, reihenfolge: (useCases.length + 1) * 10 });
    }
    setFeinOffen(false);
  }

  const patch = (p: Partial<UseCase>): void => setEntwurf(prev => (prev ? { ...prev, ...p } : prev));
  const patchRes = (p: Partial<UseCase['ressourcen']>): void =>
    setEntwurf(prev => (prev ? { ...prev, ressourcen: { ...LEER_RESSOURCEN, ...prev.ressourcen, ...p } } : prev));

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
    setSpeichert(true);
    try {
      const ok = editId ? await update(editId, entwurf) : !!(await create(entwurf));
      if (!ok) { showAlert(t('Das Speichern hat nicht geklappt. Versuch es noch einmal.', 'Saving did not work. Please try again.'), { variant: 'error' }); return; }
      setEntwurf(null);
      setEditId(null);
    } finally {
      setSpeichert(false);
    }
  }

  async function loeschen(uc: UseCase): Promise<void> {
    const ja = await confirmDialog(
      t(`„${uc.titel}" löschen? Die Kachel verschwindet für alle, und die hinterlegten Links sind weg.`,
        `Delete "${uc.titel}"? The tile disappears for everyone and the stored links are gone.`),
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
        if (await create(uc)) angelegt++;
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
    { w: 'Archiviert', de: 'Archiviert', en: 'Archived', hilfe: t('Nur noch für Kuratoren sichtbar.', 'Visible to curators only.') },
  ];
  const BEWERTUNGEN: Bewertung[] = ['Hoch', 'Mittel', 'Niedrig', 'unbewertet'];

  return (
    <div>
      <div className="dex-ui-page-head">
        <div>
          <h1 className="dex-ui-page-head-title">{t('Use Cases pflegen', 'Manage use cases')}</h1>
          <p className="dex-ui-page-head-meta">
            {useCases.length} {t('Einträge', 'entries')} · {useCases.filter(u => u.status === 'Live').length} {t('aufrufbar', 'callable')}
          </p>
        </div>
        <div className="dex-ui-page-head-actions">
          <button type="button" className="btn btn-primary" onClick={() => oeffne()}>
            <Plus size={15} /> {t('Neuer Use Case', 'New use case')}
          </button>
        </div>
      </div>

      {useCases.length === 0 ? (
        <div className="dex-ui-empty">
          <div className="dex-ui-empty-title">{t('Noch nichts angelegt', 'Nothing here yet')}</div>
          <div className="dex-ui-empty-desc">
            {t('Leg die Use Cases schon vor dem Hackathon an — die Teams tragen ihre Links später selbst nach.',
              'Create the use cases before the hackathon — the teams add their links later.')}
          </div>
          <button type="button" className="dex-ui-empty-action" onClick={() => oeffne()}>
            {t('Ersten Use Case anlegen', 'Create the first use case')}
          </button>
          {/* Der Startbestand aus dem Konzept-Deck. Nur solange die Liste
              leer ist — danach waere der Knopf ein Weg, versehentlich
              Dubletten anzulegen. */}
          <button type="button" className="dex-ui-textbtn" style={{ marginTop: 10 }} disabled={speichert} onClick={() => { void startbestand(); }}>
            {speichert
              ? t('Wird angelegt …', 'Creating …')
              : t('Die fünf Use Cases aus dem Konzept anlegen', 'Create the five use cases from the concept')}
          </button>
        </div>
      ) : (
        <div className="dex-ui-stack">
          {useCases.map(uc => (
            <div key={uc.id} className="dex-ui-row dex-ui-row--bordered">
              <span className="dex-ui-row-main">
                <span className="dex-ui-row-title dex-ui-row-title--wrap">
                  {uc.titel}
                  <span className={cx('dex-ui-pill', 'dex-ui-pill--sm', uc.status === 'Live' ? 'dex-ui-pill--green' : 'dex-ui-pill--gray')}>
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
          onClose={() => setEntwurf(null)}
          maxWidth={640}
          ariaLabel={editId ? t('Use Case bearbeiten', 'Edit use case') : t('Neuer Use Case', 'New use case')}
          title={editId ? t('Use Case bearbeiten', 'Edit use case') : t('Neuer Use Case', 'New use case')}
          subtitle={t('Titel und Kurzbeschreibung stehen auf der Kachel.', 'Title and short description appear on the tile.')}
          footer={<>
            <button type="button" className="btn btn-secondary" onClick={() => setEntwurf(null)}>{t('Abbrechen', 'Cancel')}</button>
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
              {t('Bewertung, Aufruf und Reihenfolge', 'Assessment, launch and order')}
              <span className="dex-ui-disclosure-count">{t('optional', 'optional')}</span>
            </button>
            {feinOffen && (
              <div className="dex-ui-disclosure-body">
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
                  <span className="dex-ui-label">{t('Bild für die Kachel (URL)', 'Tile image (URL)')} <span className="dex-ui-label-optional">{t('optional', 'optional')}</span></span>
                  <input id="uc-bild" type="url" className="dex-ui-input" value={entwurf.bildUrl || ''} onChange={e => patch({ bildUrl: e.target.value })} placeholder="https://…" />
                  <span className="dex-ui-help">{t('Ohne Bild zeigt die Kachel das Kürzel des Titels.', 'Without an image the tile shows the title initials.')}</span>
                </label>

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
    </div>
  );
}
