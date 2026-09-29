/**
 * Die Use Cases — laden, anlegen, aendern, loeschen.
 *
 * Der `ladeStatus` ist hier kein Beiwerk, sondern der Kern: Eine leere
 * Kachelwand und eine nicht lesbare Liste sehen gleich aus, und die
 * Kachelwand IST die App. In DEX hat genau diese Verwechslung eine
 * Organizerin ein Event mit 77 Anmeldungen als leer sehen lassen (v30.37).
 * Deshalb drei Zustaende: `laedt`, `ok`, `fehler` — und die Oberflaeche
 * benennt den dritten, statt „keine Use Cases" zu behaupten.
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { UseCase, BildAenderung, SpeicherErgebnis } from '../types';
import { useRoles } from './RoleContext';
import { START_USE_CASES } from '../data/startUseCases';
import { geaendertText } from '../utils/aenderungen';

export type LadeStatus = 'laedt' | 'ok' | 'fehler';

interface UseCaseContextType {
  useCases: UseCase[];
  ladeStatus: LadeStatus;
  /** HTTP-Status des letzten Lesens — fuer die Meldung („403" heisst Rechte, nicht kaputt). */
  letzterStatus: number;
  /** Was SharePoint im Klartext geantwortet hat. Ohne das raet man beim naechsten Mal wieder. */
  letzterFehler: string;
  /** Spalten, die beim Anlegen der Liste nicht entstanden sind. */
  fehlendeSpalten: string[];
  reload: () => Promise<void>;
  create: (uc: Partial<UseCase>) => Promise<number | null>;
  update: (id: number, uc: Partial<UseCase>) => Promise<boolean>;
  remove: (id: number) => Promise<boolean>;
  /**
   * v1.3: Einen Use Case samt Kachelbild speichern (`id === null` = neu).
   *
   * Die Reihenfolge steht HIER und nirgends sonst, weil sie das Bild schützt:
   * erst das neue Bild hochladen, dann die Zeile speichern, und das alte Bild
   * erst NACH erfolgreichem Speichern entfernen. Bricht irgendein Schritt ab,
   * bleibt der alte Stand vollständig erhalten — „erst anlegen, dann löschen"
   * (DEX v30.67). Beim Anlegen gibt es die Zeile erst, wenn sie ein Bild
   * aufnehmen kann; scheitert der Upload, ist der Use Case trotzdem da und
   * `bildFehler` sagt es.
   */
  saveUseCase: (id: number | null, uc: Partial<UseCase>, bild: BildAenderung) => Promise<SpeicherErgebnis>;
  /** Alle vorkommenden Bereiche, alphabetisch — fuer den Filter. */
  bereiche: string[];
  /**
   * Die fuenf Start-Use-Cases von Hand anlegen.
   *
   * Das Netz unter der automatischen Erstbefuellung: Die laeuft nur, wenn
   * BEIDE Listen lesbar sind — bei einem 403 auf das Protokoll bliebe die
   * Plattform sonst leer, ohne Weg sie zu fuellen. Genau dieser Zustand war
   * der Fehler, der v1.1 ausgeloest hat.
   *
   * Rueckgabe: wie viele angelegt wurden. 0 heisst, dass SharePoint sie
   * abgelehnt hat — der Grund steht dann in `letzterFehler`.
   */
  seedStartUseCases: () => Promise<number>;
}

const UseCaseContext = React.createContext<UseCaseContextType | undefined>(undefined);

export function UseCaseProvider(props: { context: WebPartContext; children: React.ReactNode }): React.ReactElement {
  const { service, isOrganizer } = useRoles();
  const [useCases, setUseCases] = React.useState<UseCase[]>([]);
  const [ladeStatus, setLadeStatus] = React.useState<LadeStatus>('laedt');
  const [letzterStatus, setLetzterStatus] = React.useState(0);
  const [letzterFehler, setLetzterFehler] = React.useState('');

  const reload = React.useCallback(async (): Promise<void> => {
    const rows = await service.getUseCases();
    setLetzterStatus(service.lastUseCasesReadStatus);
    setLetzterFehler(service.lastReadError);
    if (rows === null) {
      // NICHT auf [] setzen. Ein Lesefehler ist keine Aussage ueber die Daten.
      setLadeStatus('fehler');
      console.warn('[AIUC] Use Cases nicht lesbar — bestehender Stand bleibt.');
      return;
    }
    setUseCases(rows);
    setLadeStatus('ok');
  }, [service]);

  React.useEffect(() => {
    let abgebrochen = false;
    (async (): Promise<void> => {
      // Die Listen sicherstellen, bevor gelesen wird. Beim ersten Start der
      // App existiert noch nichts; ohne diesen Schritt sieht die erste Person
      // einen Fehler statt einer leeren Plattform.
      try {
        await service.ensureUseCaseList();
        await service.ensureLogList();
      } catch (e) {
        console.warn('[AIUC] Listen konnten nicht sichergestellt werden:', e);
      }

      /*
       * Erstbefuellung mit den fuenf Use Cases aus dem Konzept-Deck.
       *
       * v1.1 (Fehlerbehebung): Die Bedingung war `isNewlyCreated` — „die
       * Liste ist gerade erst entstanden". Das trug genau EINEN Start lang.
       * Bei der ersten Installation legte v1.0.0 die Liste an, konnte wegen
       * der fehlenden Spalten aber nichts hineinschreiben; beim naechsten
       * Start war die Liste dann nicht mehr neu, und die Befuellung lief nie
       * wieder. Die Plattform blieb leer, und es gab keinen Weg, das von
       * Hand nachzuholen — die Kacheln waren weg und der Knopf dafuer auch.
       *
       * Bedingung ist jetzt der Zustand statt des Zeitpunkts: leer **und**
       * noch nie befuellt. Der Merker steht im Protokoll; ohne ihn kaeme der
       * Bestand zurueck, sobald jemand alle Eintraege absichtlich geloescht
       * hat. Und weil die Richtung wichtig ist: Ist eine der beiden Listen
       * nicht lesbar, wird NICHT befuellt — ein Lesefehler ist keine leere
       * Plattform, und fuenf Kacheln neben einen unsichtbaren Bestand zu
       * legen waere der teurere Fehler.
       */
      if (!abgebrochen && await service.darfErstbefuellen()) {
        // Nacheinander, nicht parallel — fuenf gleichzeitige POSTs gegen
        // dieselbe Liste sind genau das Muster, das SharePoint drosselt.
        let angelegt = 0;
        for (const uc of START_USE_CASES) {
          // eslint-disable-next-line no-await-in-loop
          if (await service.createUseCase(uc)) angelegt++;
        }
        if (angelegt > 0) await service.merkeErstbefuellung(angelegt);
        else console.warn('[AIUC] Erstbefüllung: kein Eintrag angelegt — Grund steht oben in der Konsole.');
      }

      if (!abgebrochen) await reload();
    })().catch(() => setLadeStatus('fehler'));
    return () => { abgebrochen = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const create = React.useCallback(async (uc: Partial<UseCase>): Promise<number | null> => {
    if (!isOrganizer) return null;
    const id = await service.createUseCase(uc);
    if (id) {
      await service.log(id, 'angelegt', uc.titel || '');
      await reload();
    }
    return id;
  }, [service, reload, isOrganizer]);

  const update = React.useCallback(async (id: number, uc: Partial<UseCase>): Promise<boolean> => {
    if (!isOrganizer) return false;
    const alt = useCases.filter(u => u.id === id)[0];
    const ok = await service.updateUseCase(id, uc);
    if (ok) {
      // Was sich WIRKLICH geändert hat (v1.3) — vorher stand hier bei jedem
      // Speichern die Liste aller Schlüssel. Keine Änderung, keine Zeile.
      const text = geaendertText(alt, uc);
      if (text) await service.log(id, 'geaendert', text);
      await reload();
    }
    return ok;
  }, [service, reload, isOrganizer, useCases]);

  const saveUseCase = React.useCallback(async (
    id: number | null,
    uc: Partial<UseCase>,
    bild: BildAenderung,
  ): Promise<SpeicherErgebnis> => {
    if (!isOrganizer) return { ok: false, id, bildFehler: false, grund: 'rechte' };

    // --- Neu anlegen ---------------------------------------------------
    if (id === null) {
      // Ohne Bild-URL anlegen: Das Bild braucht die Zeile (der Anhang hängt an
      // ihr) und kommt danach.
      const neueId = await service.createUseCase({ ...uc, bildUrl: '' });
      if (!neueId) return { ok: false, id: null, bildFehler: false, grund: 'speichern' };
      await service.log(neueId, 'angelegt', uc.titel || '');
      let bildFehler = false;
      if (bild.art === 'neu') {
        const hoch = await service.uploadBild(neueId, bild.datei);
        if (hoch && await service.updateUseCase(neueId, { bildUrl: hoch.url })) {
          await service.log(neueId, 'geaendert', 'Bild');
        } else {
          bildFehler = true;
          // Ein Anhang ohne Verweis darauf ist Ballast — wegräumen.
          if (hoch) await service.entferneBilder(neueId, { nur: hoch.name });
        }
      }
      await reload();
      return { ok: true, id: neueId, bildFehler };
    }

    // --- Ändern --------------------------------------------------------
    const alt = useCases.filter(u => u.id === id)[0];

    // 1. Das neue Bild ZUERST. Scheitert es, ist noch nichts verändert.
    let neuesBild: { url: string; name: string } | null = null;
    if (bild.art === 'neu') {
      neuesBild = await service.uploadBild(id, bild.datei);
      if (!neuesBild) return { ok: false, id, bildFehler: true, grund: 'bild' };
    }

    // 2. Die Zeile speichern — mit der neuen Adresse (oder leer beim Entfernen).
    const daten: Partial<UseCase> = { ...uc };
    if (neuesBild) daten.bildUrl = neuesBild.url;
    else if (bild.art === 'entfernen') daten.bildUrl = '';
    const ok = await service.updateUseCase(id, daten);
    if (!ok) {
      // Rückbau: das gerade hochgeladene Bild gehört jetzt niemandem.
      if (neuesBild) await service.entferneBilder(id, { nur: neuesBild.name });
      return { ok: false, id, bildFehler: false, grund: 'speichern' };
    }

    const text = geaendertText(alt, daten);
    if (text) await service.log(id, 'geaendert', text);

    // 3. Erst jetzt das alte Bild — der unumkehrbare Schritt kommt zuletzt.
    if (bild.art !== 'unveraendert') {
      await service.entferneBilder(id, neuesBild ? { behalte: neuesBild.name } : {});
    }
    await reload();
    return { ok: true, id, bildFehler: false };
  }, [service, reload, isOrganizer, useCases]);

  const remove = React.useCallback(async (id: number): Promise<boolean> => {
    if (!isOrganizer) return false;
    // Protokoll VOR dem Loeschen — danach ist der Titel weg, und das
    // Protokoll waere die Nebenbuchhaltung, die nichts mehr belegt.
    // (Dieselbe Reihenfolge wie in DEX: pruefbare Nebenbuchhaltung zuerst,
    // der unumkehrbare Schritt zuletzt.)
    const uc = useCases.filter(u => u.id === id)[0];
    await service.log(id, 'geloescht', uc ? uc.titel : `Id ${id}`);
    const ok = await service.deleteUseCase(id);
    if (ok) await reload();
    // Das Protokoll steht VOR dem Löschen und behauptet „gelöscht". Schlägt
    // das Löschen fehl, muss die Gegenbuchung dastehen — sonst sagt das
    // Protokoll etwas, das nicht stimmt.
    else await service.log(id, 'loeschen-fehlgeschlagen', uc ? uc.titel : `Id ${id}`);
    return ok;
  }, [service, reload, useCases, isOrganizer]);

  const seedStartUseCases = React.useCallback(async (): Promise<number> => {
    let angelegt = 0;
    for (const uc of START_USE_CASES) {
      // eslint-disable-next-line no-await-in-loop
      if (await service.createUseCase(uc)) angelegt++;
    }
    setLetzterFehler(service.lastReadError);
    if (angelegt > 0) {
      await service.merkeErstbefuellung(angelegt);
      await reload();
    }
    return angelegt;
  }, [service, reload]);

  const bereiche = React.useMemo(() => {
    const set: Record<string, true> = {};
    useCases.forEach(u => { if (u.bereich) set[u.bereich] = true; });
    return Object.keys(set).sort((a, b) => a.localeCompare(b, 'de'));
  }, [useCases]);

  const value = React.useMemo<UseCaseContextType>(() => ({
    useCases, ladeStatus, letzterStatus, letzterFehler, fehlendeSpalten: service.fehlendeSpalten,
    reload, create, update, remove, saveUseCase, bereiche, seedStartUseCases,
  }), [useCases, ladeStatus, letzterStatus, letzterFehler, reload, create, update, remove, saveUseCase, bereiche, seedStartUseCases, service]);

  return React.createElement(UseCaseContext.Provider, { value }, props.children);
}

export function useUseCases(): UseCaseContextType {
  const ctx = React.useContext(UseCaseContext);
  if (!ctx) throw new Error('useUseCases muss innerhalb des UseCaseProvider stehen');
  return ctx;
}
