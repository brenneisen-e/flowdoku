/**
 * v32.45 — Schutz beim Entfernen beantworteter Fragen und Hinweis bei neuen
 * Fragen für ein bestehendes Event mit Anmeldungen (Nutzer-Ansagen
 * 30.09.2026). Die Zählung steht in `logic/fieldAnswerGuard`.
 *
 * Die beiden Entfernen-Wege ersetzen `removeCustomField` bzw.
 * `removeSubEventCustomField` 1:1 in den Props des Fragen-Schritts — dieselbe
 * Signatur, damit keine Aufrufstelle in FieldsStep angefasst werden muss.
 */
import * as React from 'react';
import { DeloitteEvent } from '../../../types';
import { ConfirmOptions } from '../../../context/DialogContext';
import { SubEventDraft } from '../wizardTypes';
import { CustomFieldInput } from '../customFieldInput';
import { leseAntwortStand, AntwortStand } from '../logic/fieldAnswerGuard';
import { Mail } from '../../Icons';

export interface UseFieldAnswerGuardCtx {
  isEditMode: boolean;
  editEvent: DeloitteEvent | null | undefined;
  childEventsOf: (parentId: string) => DeloitteEvent[];
  customFields: CustomFieldInput[];
  subEvents: SubEventDraft[];
  currentStep: number;
  isDe: boolean;
  removeCustomField: (id: string) => void;
  removeSubEventCustomField: (subEventId: string, fieldId: string) => void;
  confirmDialog: (message: React.ReactNode, opts?: ConfirmOptions) => Promise<boolean>;
}

const FRAGEN_SCHRITT = 4;

export function useFieldAnswerGuard(ctx: UseFieldAnswerGuardCtx): {
  removeCustomFieldGuarded: (id: string) => void;
  removeSubEventCustomFieldGuarded: (subEventId: string, fieldId: string) => void;
  neueFragenHinweis: React.ReactNode;
} {
  const { isEditMode, editEvent, childEventsOf, customFields, subEvents, currentStep, isDe, removeCustomField, removeSubEventCustomField, confirmDialog } = ctx;
  const kinder = isEditMode && editEvent ? childEventsOf(editEvent.id) : [];
  // Frage des Hauptevents: Klammer- UND Termin-Listen (Antworten stehen dort,
  // wo angemeldet wurde). Frage eines Termins: nur dessen Liste.
  const alleUrls = [editEvent?.subsiteUrl || '', ...kinder.map(k => k.subsiteUrl || '')].filter(Boolean);
  const alteIdsHaupt = new Set((editEvent?.eventSpecificFields || []).map(f => f.id));
  const kindZuDraft = (draftId: string): DeloitteEvent | undefined => {
    const d = subEvents.find(s => s.id === draftId);
    return d && d.dbId ? kinder.find(k => String(k.id) === String(d.dbId)) : undefined;
  };

  // Beim Öffnen des Schritts einmal lesen — dann ist der Dialog beim Klick auf
  // „×" sofort da, und der Hinweis zu neuen Fragen kennt die Personenzahl.
  const [stand, setStand] = React.useState<AntwortStand | null>(null);
  const urlKey = alleUrls.join('|');
  React.useEffect(() => {
    if (!isEditMode || currentStep !== FRAGEN_SCHRITT || !urlKey) return undefined;
    let aktiv = true;
    leseAntwortStand(urlKey.split('|')).then(s => { if (aktiv) setStand(s); }).catch(() => { /* bleibt unbekannt */ });
    return () => { aktiv = false; };
  }, [isEditMode, currentStep, urlKey]);

  const fragen = async (urls: string[], fieldId: string, label: string): Promise<boolean> => {
    const s = await leseAntwortStand(urls).catch(() => null);
    const n = s ? (s.antworten[fieldId] || 0) : 0;
    if (s && s.gelesen && n === 0) return true;
    const frage = label.trim() ? `„${label.trim()}“` : (isDe ? 'Diese Frage' : 'This question');
    const text = s && s.gelesen
      ? (isDe
        ? <><p style={{ margin: '0 0 10px' }}><strong>{frage}</strong> wurde bereits von <strong>{n} {n === 1 ? 'Person' : 'Personen'}</strong> beantwortet.</p><p style={{ margin: 0 }}>Entfernst du die Frage, verschwinden diese Antworten beim Speichern aus Teilnehmerliste, Export und Mails. Über die App lassen sie sich nicht zurückholen – auch nicht, wenn du die Frage neu anlegst.</p></>
        : <><p style={{ margin: '0 0 10px' }}><strong>{frage}</strong> has already been answered by <strong>{n} {n === 1 ? 'person' : 'people'}</strong>.</p><p style={{ margin: 0 }}>If you remove it, these answers disappear from the participant list, export and mails when you save. They cannot be restored in the app – not even by adding the question again.</p></>)
      : (isDe
        ? <><p style={{ margin: '0 0 10px' }}>Ob <strong>{frage}</strong> schon beantwortet wurde, ließ sich nicht prüfen – die Teilnehmerliste war nicht lesbar.</p><p style={{ margin: 0 }}>Gibt es Antworten, verschwinden sie beim Speichern aus Teilnehmerliste, Export und Mails und lassen sich über die App nicht zurückholen.</p></>
        : <><p style={{ margin: '0 0 10px' }}>Could not check whether <strong>{frage}</strong> has been answered – the participant list could not be read.</p><p style={{ margin: 0 }}>Any answers will disappear from the participant list, export and mails when you save and cannot be restored in the app.</p></>);
    return confirmDialog(text, {
      title: isDe ? 'Beantwortete Frage entfernen?' : 'Remove answered question?',
      confirmLabel: isDe ? 'Trotzdem entfernen' : 'Remove anyway',
      danger: true,
    });
  };

  const removeCustomFieldGuarded = (id: string): void => {
    if (!isEditMode || !alteIdsHaupt.has(id) || alleUrls.length === 0) { removeCustomField(id); return; }
    const f = customFields.find(x => x.id === id);
    void fragen(alleUrls, id, (f && f.label) || '').then(ok => { if (ok) removeCustomField(id); });
  };

  const removeSubEventCustomFieldGuarded = (subEventId: string, fieldId: string): void => {
    const kind = isEditMode ? kindZuDraft(subEventId) : undefined;
    const alt = !!kind && (kind.eventSpecificFields || []).some(f => f.id === fieldId);
    if (!kind || !alt || !kind.subsiteUrl) { removeSubEventCustomField(subEventId, fieldId); return; }
    const d = subEvents.find(s => s.id === subEventId);
    const f = (d && d.customFields || []).find(x => x.id === fieldId);
    void fragen([kind.subsiteUrl], fieldId, (f && f.label) || '').then(ok => { if (ok) removeSubEventCustomField(subEventId, fieldId); });
  };

  // Neue Fragen bei einem Event, das schon Anmeldungen hat.
  let neu = customFields.filter(f => (f.label || '').trim() && !alteIdsHaupt.has(f.id)).length;
  subEvents.forEach(d => {
    const kind = kindZuDraft(d.id);
    if (!kind) return;
    const alt = new Set((kind.eventSpecificFields || []).map(f => f.id));
    neu += (d.customFields || []).filter(f => (f.label || '').trim() && !alt.has(f.id)).length;
  });
  const personen = stand ? stand.personen : 0;
  const zeigen = isEditMode && neu > 0 && !!stand && (personen > 0 || !stand.gelesen);
  const neueFragenHinweis = zeigen ? (
    <div className="dex-ui-callout dex-ui-callout--warn" style={{ margin: '0 0 16px' }}>
      <span className="dex-ui-callout-icon"><Mail size={16} /></span>
      <span>
        {isDe ? (
          <>
            <strong>{neu === 1 ? 'Neue Frage' : `${neu} neue Fragen`} bei bestehenden Anmeldungen.</strong>{' '}
            {stand && stand.gelesen
              ? <>Die {personen} bereits angemeldeten {personen === 1 ? 'Person hat' : 'Personen haben'} {neu === 1 ? 'sie' : 'sie'} noch nicht beantwortet.</>
              : <>Wer schon angemeldet ist, hat {neu === 1 ? 'sie' : 'sie'} noch nicht beantwortet (Teilnehmerliste gerade nicht lesbar).</>}{' '}
            Unter „Meine Events“ steht die Frage als „Noch offen“ und lässt sich über „Angaben ergänzen“ nachtragen. Schick den Bisherigen nach dem Speichern am besten eine Mail mit dieser Bitte (Organizer Center → E-Mail versenden).
          </>
        ) : (
          <>
            <strong>{neu === 1 ? 'New question' : `${neu} new questions`} with existing registrations.</strong>{' '}
            {stand && stand.gelesen
              ? <>The {personen} people already registered have not answered {neu === 1 ? 'it' : 'them'} yet.</>
              : <>People already registered have not answered {neu === 1 ? 'it' : 'them'} yet (participant list not readable right now).</>}{' '}
            Under “My events” the question shows as “Still open” and can be filled in via “Add details”. After saving, best send them a mail asking to do so (Organizer Center → Send email).
          </>
        )}
      </span>
    </div>
  ) : null;

  return { removeCustomFieldGuarded, removeSubEventCustomFieldGuarded, neueFragenHinweis };
}
