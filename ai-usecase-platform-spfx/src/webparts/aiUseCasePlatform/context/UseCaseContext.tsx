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
import { UseCase } from '../types';
import { useRoles } from './RoleContext';
import { START_USE_CASES } from '../data/startUseCases';

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
  const { service, isKurator } = useRoles();
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
    if (!isKurator) return null;
    const id = await service.createUseCase(uc);
    if (id) {
      await service.log(id, 'angelegt', uc.titel || '');
      await reload();
    }
    return id;
  }, [service, reload, isKurator]);

  const update = React.useCallback(async (id: number, uc: Partial<UseCase>): Promise<boolean> => {
    if (!isKurator) return false;
    const ok = await service.updateUseCase(id, uc);
    if (ok) {
      await service.log(id, 'geaendert', Object.keys(uc).join(', '));
      await reload();
    }
    return ok;
  }, [service, reload, isKurator]);

  const remove = React.useCallback(async (id: number): Promise<boolean> => {
    if (!isKurator) return false;
    // Protokoll VOR dem Loeschen — danach ist der Titel weg, und das
    // Protokoll waere die Nebenbuchhaltung, die nichts mehr belegt.
    // (Dieselbe Reihenfolge wie in DEX: pruefbare Nebenbuchhaltung zuerst,
    // der unumkehrbare Schritt zuletzt.)
    const uc = useCases.filter(u => u.id === id)[0];
    await service.log(id, 'geloescht', uc ? uc.titel : `Id ${id}`);
    const ok = await service.deleteUseCase(id);
    if (ok) await reload();
    return ok;
  }, [service, reload, useCases, isKurator]);

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
    reload, create, update, remove, bereiche, seedStartUseCases,
  }), [useCases, ladeStatus, letzterStatus, letzterFehler, reload, create, update, remove, bereiche, seedStartUseCases, service]);

  return React.createElement(UseCaseContext.Provider, { value }, props.children);
}

export function useUseCases(): UseCaseContextType {
  const ctx = React.useContext(UseCaseContext);
  if (!ctx) throw new Error('useUseCases muss innerhalb des UseCaseProvider stehen');
  return ctx;
}
