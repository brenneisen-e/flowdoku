/* PendingPeopleBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10452-10523 des
 * Stands vor dem Schnitt). Wer hier steht, entscheidet der Aufrufer.
 *
 * v31.3: Nach docs/ui-leitfaden.md (Abschnitt „Modale") umgebaut — Kopf und
 * Fuss kommen jetzt aus `Modal` (title/subtitle/icon/footer) statt aus einem
 * eigenen <h3> mit `padding: 0`; die Personen stehen als `dex-ui-person`-Zelle
 * (Foto, Name, darunter Position und Adresse) statt in Zebrastreifen, und der
 * Primärknopf sagt, was danach passiert.
 *
 * v31.26 (Nutzer-Ansage 11.09.2026: „hier sollte Vorname Nachname E-Mail und
 * nach Nachname sortiert sein und dann hinter jedem einzelnen dann auch
 * Reminder senden"):
 *
 *  - **Namen.** Es stand zweimal dieselbe Adresse untereinander. Der Grund
 *    liegt nicht in der Anzeige, sondern in der Quelle: Namen kommen aus der
 *    Verteiler-Abfrage, und wenn die Sichtbarkeit über
 *    `audienceResolvedEmails` aufgelöst wurde (eine blanke Adressliste), gibt
 *    es schlicht keine. Der Kasten holt sie jetzt beim Öffnen über Graph nach
 *    (`getPeopleByEmails`) — und weil das ein Nachtrag ist, rendert die Liste
 *    sofort mit den Adressen und tauscht die Namen ein, sobald sie da sind.
 *    Warten wäre der falsche Tausch: Der Kasten soll zeigen, WER fehlt, nicht
 *    auf hübschere Namen warten.
 *  - **Sortierung nach Nachname.** Erst mit Namen überhaupt möglich. Wo keiner
 *    bekannt ist, sortiert die Adresse — und diese Personen stehen hinten,
 *    nicht mitten in der Liste: Eine Adresse zwischen zwei Nachnamen liest
 *    sich wie ein Sortierfehler.
 *  - **Erinnern je Person.** Öffnet denselben Mail-Dialog wie der Sammel-Knopf,
 *    nur mit einer Adresse. Bewusst KEIN zweiter Versandweg: Ein eigener
 *    Ein-Personen-Pfad hätte eigene Vorlage, eigenes Protokoll und eigene
 *    Fehlerbehandlung — und die erste Abweichung fällt niemandem auf.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Mail, Users } from '../../Icons';
import { PersonContactHover } from '../../PersonContactHover';
import { AudiencePerson } from '../../admin/adminTypes';
import { EventService } from '../../../services/EventService';

export interface PendingPeopleBoxProps {
  isDe: boolean;
  openInviteModal: () => void;
  pendingPeople: { people: AudiencePerson[]; reachable: number; };
  setInviteAudienceOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setInviteCustomEmails: React.Dispatch<React.SetStateAction<string[]>>;
  setInviteTarget: React.Dispatch<React.SetStateAction<"organizer" | "audience" | "pending" | "uninvited">>;
  setPendingPeople: React.Dispatch<React.SetStateAction<{ people: AudiencePerson[]; reachable: number; }>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

/** Was wir über eine Person wissen — angereichert um das, was Graph liefert. */
interface Person {
  email: string;
  vorname: string;
  nachname: string;
  anzeige: string;
  zusatz: string;
  /** Sortierschlüssel: Nachname, sonst die Adresse (die dann hinten steht). */
  sortKey: string;
  nameBekannt: boolean;
}

export const PendingPeopleBox: React.FC<PendingPeopleBoxProps> = (p) => {
  const { isDe, openInviteModal, pendingPeople, setInviteAudienceOpen, setInviteCustomEmails, setInviteTarget, setPendingPeople, showAlert } = p;

  // Was Graph zusätzlich weiß. Leer = noch nicht geladen oder nichts gefunden.
  const [graph, setGraph] = React.useState<Record<string, { displayName: string; firstName: string; lastName: string; jobTitle: string; location: string }>>({});

  React.useEffect(() => {
    let weg = false;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    if (!ctx) return undefined;
    // Nur die Adressen nachschlagen, zu denen der Verteiler KEINEN Namen
    // geliefert hat — alles andere wäre eine Graph-Abfrage für nichts.
    const offen = pendingPeople.people.filter(x => !(x.displayName || '').trim()).map(x => x.email);
    if (offen.length === 0) return undefined;
    (async () => {
      try {
        const map = await new EventService(ctx).getPeopleByEmails(offen);
        if (!weg) setGraph(map);
      } catch { /* best-effort — ohne Namen bleibt die Adresse stehen */ }
    })().catch(() => { /* */ });
    return () => { weg = true; };
  }, [pendingPeople]);

  const personen = React.useMemo<Person[]>(() => {
    const liste = pendingPeople.people.map(x => {
      const g = graph[x.email];
      const vorname = (g?.firstName || '').trim();
      const nachname = (g?.lastName || '').trim();
      // Der Anzeigename aus dem Verteiler steht oft als „Nachname, Vorname" —
      // in dieser Liste soll aber „Vorname Nachname" stehen. Liegen beide
      // Teile getrennt vor, bauen wir sie selbst zusammen; sonst gilt, was da
      // ist, und zuletzt die Adresse.
      const anzeige = (vorname && nachname)
        ? `${vorname} ${nachname}`
        : ((x.displayName || '').trim() || (g?.displayName || '').trim() || x.email);
      const nameBekannt = anzeige !== x.email;
      const zusatz = [g?.jobTitle || x.jobTitle, g?.location || x.location].filter(Boolean).join(' · ');
      // Ohne Nachnamen hilft der Anzeigename: Steht dort „Nachname, Vorname",
      // ist der Teil vor dem Komma genau der Nachname.
      const ausAnzeige = nameBekannt && anzeige.indexOf(',') > 0 ? anzeige.split(',')[0].trim() : '';
      const key = (nachname || ausAnzeige || '').toLowerCase();
      return {
        email: x.email, vorname, nachname, anzeige, zusatz, nameBekannt,
        // `~` sortiert hinter allen Buchstaben — Adressen ohne Namen landen
        // damit geschlossen am Ende statt verstreut zwischen den Nachnamen.
        sortKey: key ? key : `~${x.email}`,
      };
    });
    return liste.sort((a, b) => a.sortKey.localeCompare(b.sortKey, isDe ? 'de' : 'en'));
  }, [pendingPeople, graph, isDe]);

  /** Den Mail-Dialog öffnen — für alle oder für genau eine Person. */
  const erinnern = (emails: string[]): void => {
    setPendingPeople(null);
    openInviteModal();
    setInviteTarget('pending');
    setInviteCustomEmails(emails);
    // Nur bei mehreren die Empfängerliste aufklappen. Bei einer Person ist
    // die Frage „an wen" schon beantwortet; das Feld aufzuklappen lenkt vom
    // Text ab, der hier das Eigentliche ist.
    setInviteAudienceOpen(emails.length > 1);
  };

  const offeneNamen = personen.filter(x => !x.nameBekannt).length;

  return (
          <Modal
            open={true}
            onClose={() => setPendingPeople(null)}
            maxWidth={720}
            ariaLabel={isDe ? 'Wer hat noch nicht geantwortet' : 'Who has not responded yet'}
            title={isDe ? 'Noch keine Rückmeldung' : 'No response yet'}
            icon={<Users size={20} />}
            subtitle={isDe
              ? <>Von <strong>{pendingPeople.reachable}</strong> Personen, die dieses Event sehen können, haben sich <strong>{pendingPeople.people.length}</strong> noch gar nicht geäußert: weder angemeldet noch abgemeldet und auch keine Absage. Das Organizer-Team ist nicht mitgezählt. Sortiert nach Nachname.</>
              : <>Of <strong>{pendingPeople.reachable}</strong> people who can see this event, <strong>{pendingPeople.people.length}</strong> have not responded at all: neither registered nor cancelled nor declined. The organizer team is not counted. Sorted by surname.</>}
            footer={<>
              {/* v31.3: Die beiden Nebenaktionen links, der Primärknopf rechts
                  aussen (Leitfaden „Modale") — vorher schob ein margin-left:auto
                  den Knopf weg und die Fusszeile hatte keine Trennlinie. */}
              <span className="dex-ui-modal-foot-left">
                <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => setPendingPeople(null)}>
                  {isDe ? 'Schließen' : 'Close'}
                </button>
                <button
                  type="button"
                  className="btn btn-secondary dex-ui-btn-sm"
                  onClick={() => {
                    // v31.26: In der Reihenfolge der Liste, nicht in der der
                    // Rohdaten — wer die Adressen neben die Liste legt, soll
                    // dieselbe Abfolge sehen.
                    const list = personen.map(x => x.email).join('; ');
                    void navigator.clipboard.writeText(list).then(
                      () => showAlert(isDe ? `${personen.length} Adressen kopiert.` : `${personen.length} addresses copied.`, { variant: 'success' }),
                      () => showAlert(isDe ? 'Kopieren nicht möglich.' : 'Copying failed.', { variant: 'error' })
                    );
                  }}
                >
                  {isDe ? 'Adressen kopieren' : 'Copy addresses'}
                </button>
              </span>
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => erinnern(personen.map(x => x.email))}
              >
                {isDe ? `Erinnerung an alle ${personen.length} schreiben` : `Write reminder to all ${personen.length}`}
              </button>
            </>}
          >
            {/* v31.3: Personen-Zelle statt Zebrastreifen (Leitfaden 5b). */}
            <div style={{ maxHeight: '48vh', overflowY: 'auto' }}>
              {personen.map(person => (
                <div
                  key={person.email}
                  className="dex-ui-person dex-ui-row--bordered"
                  style={{ display: 'flex', alignItems: 'center', padding: '8px 0', gap: 10 }}
                >
                  <PersonContactHover email={person.email} name={person.anzeige} size={34} subline={person.zusatz} isDe={isDe} />
                  <div style={{ minWidth: 0, flex: 1 }}>
                    <div className="dex-ui-person-name">{person.anzeige}</div>
                    {/* Die Adresse steht IMMER darunter — auch wenn der Name
                        bekannt ist. Sie ist das, wonach der Organizer im
                        Zweifel sucht, und der Nutzer hat sie ausdrücklich
                        neben dem Namen verlangt. Ist der Name unbekannt,
                        stünde sie zweimal; dann bleibt es bei einer. */}
                    {person.nameBekannt && (
                      <div className="dex-ui-person-sub">{person.zusatz ? `${person.zusatz} · ${person.email}` : person.email}</div>
                    )}
                    {!person.nameBekannt && person.zusatz && (
                      <div className="dex-ui-person-sub">{person.zusatz}</div>
                    )}
                  </div>
                  <button
                    type="button"
                    className="btn btn-secondary dex-ui-btn-sm"
                    style={{ flexShrink: 0 }}
                    onClick={() => erinnern([person.email])}
                    title={isDe
                      ? `Erinnerung nur an ${person.anzeige} — der Text lässt sich vorher noch ändern`
                      : `Reminder to ${person.anzeige} only — you can still edit the text`}
                  >
                    <Mail size={14} /> {isDe ? 'Erinnern' : 'Remind'}
                  </button>
                </div>
              ))}
              {/* Ehrlich benennen, was nicht aufgelöst werden konnte — sonst
                  liest sich eine Adresse ohne Namen wie ein Anzeigefehler. */}
              {offeneNamen > 0 && (
                <div className="dex-ui-muted" style={{ fontSize: '0.78rem', padding: '10px 0 2px' }}>
                  {isDe
                    ? `Zu ${offeneNamen} Adresse(n) kennt DEX keinen Namen — sie stehen am Ende der Liste. Das passiert, wenn die Sichtbarkeit über einen Verteiler läuft, der nur Adressen liefert.`
                    : `DEX does not know a name for ${offeneNamen} address(es) — they are listed last. This happens when visibility runs via a distribution list that only yields addresses.`}
                </div>
              )}
            </div>
          </Modal>
  );
};
