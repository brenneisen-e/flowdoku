/* v32.0.8: Rückmeldung nach „Live schalten".
 *
 * Nutzer-Ansage 28.09.2026: „Wenn man ein Event live geschaltet hat, soll
 * ein Modal kommen: Event ist nun live für die Verteilergruppe (nennen) —
 * und dass man nun gerne eine Einladungsmail über DEX versenden kann (mit
 * Button)." Die Zielgruppe steht hier mit Namen (Standorte, Verteiler bzw.
 * Personen aus `audienceFilter`), nicht nur als Anzahl wie in „Nächste
 * Schritte". Der Knopf öffnet denselben Einladungs-Dialog wie dort. */
import * as React from 'react';
import Modal from '../../Modal';
import { Check, Mail, Users } from '../../Icons';
import { DeloitteEvent } from '../../../types';

export function LiveGeschaltetModal(p: {
  event: DeloitteEvent | null;
  isDe: boolean;
  onClose: () => void;
  onInvite: () => void;
}): React.ReactElement | null {
  const { event, isDe, onClose, onInvite } = p;
  if (!event) return null;
  const locs = (event.locationAudience || []).map(s => (s || '').trim()).filter(Boolean);
  const auds = (event.audienceFilter || []).map(s => (s || '').trim()).filter(Boolean);
  const offen = locs.length === 0 && auds.length === 0;
  const undOder = event.filterMode === 'OR' ? (isDe ? 'oder' : 'or') : (isDe ? 'und' : 'and');
  const chip = (t: string): React.ReactNode => (
    <span key={t} className="dex-ui-chip" style={{ cursor: 'default' }}>{t}</span>
  );
  return (
    <Modal
      open
      onClose={onClose}
      backdropClose={false}
      maxWidth={540}
      icon={<Check size={20} />}
      title={isDe ? 'Dein Event ist jetzt live' : 'Your event is now live'}
      subtitle={event.title}
      footer={<>
        <button type="button" className="btn btn-secondary" onClick={onClose}>{isDe ? 'Später' : 'Later'}</button>
        <button type="button" className="btn btn-primary" onClick={onInvite}>
          <Mail size={14} /> {isDe ? 'Einladungsmail öffnen' : 'Open invitation email'}
        </button>
      </>}
    >
      <div className="dex-ui-stack">
        <div>
          <div className="dex-ui-section-title" style={{ marginBottom: 6 }}>
            <Users size={14} /> {isDe ? 'Sichtbar und buchbar für' : 'Visible and bookable for'}
          </div>
          {offen ? (
            <div>{isDe ? 'alle Mitarbeitenden von Deloitte Deutschland' : 'all Deloitte Germany employees'}</div>
          ) : (
            <div className="dex-ui-stack" style={{ gap: 6 }}>
              {locs.length > 0 && (
                <div className="dex-ui-inline" style={{ flexWrap: 'wrap', gap: 6 }}>
                  <span className="dex-ui-muted">{isDe ? (locs.length === 1 ? 'Standort' : 'Standorte') : (locs.length === 1 ? 'Location' : 'Locations')}</span>
                  {locs.map(chip)}
                </div>
              )}
              {locs.length > 0 && auds.length > 0 && <div className="dex-ui-muted">{undOder}</div>}
              {auds.length > 0 && (
                <div className="dex-ui-inline" style={{ flexWrap: 'wrap', gap: 6 }}>
                  <span className="dex-ui-muted">{isDe ? 'Verteiler bzw. Personen' : 'Lists or people'}</span>
                  {auds.slice(0, 8).map(chip)}
                  {auds.length > 8 && <span className="dex-ui-muted">{isDe ? `+ ${auds.length - 8} weitere` : `+ ${auds.length - 8} more`}</span>}
                </div>
              )}
            </div>
          )}
        </div>
        <div className="dex-ui-callout dex-ui-callout--info">
          <span className="dex-ui-callout-icon"><Mail size={16} /></span>
          <span>
            {isDe
              ? <>Als Nächstes kannst du über DEX eine <strong>Einladungsmail</strong> mit Anmelde-Link verschicken — an dich zum Weiterleiten oder direkt an den Verteiler.</>
              : <>Next, you can send an <strong>invitation email</strong> with the registration link via DEX — to yourself for forwarding or straight to the distribution list.</>}
          </span>
        </div>
      </div>
    </Modal>
  );
}
