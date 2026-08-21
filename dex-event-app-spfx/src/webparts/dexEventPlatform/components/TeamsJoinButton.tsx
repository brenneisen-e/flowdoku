/**
 * v29.39 — Teilnahme-Knopf für ein Teams-Meeting.
 *
 * Eine Komponente für beide Seiten (Organizer Center und „Meine Events"), damit
 * der Knopf nicht zweimal entsteht und beim nächsten Mal auseinanderläuft.
 * `target="_blank"` mit `rel="noopener"`: Der Teams-Link öffnet je nach Setup
 * die Desktop-App oder den Browser — in beiden Fällen soll die DEX-Seite offen
 * bleiben, sonst verliert der Teilnehmer seine Anmeldeansicht.
 */
import * as React from 'react';
import { Video } from './Icons';

export interface TeamsJoinButtonProps {
  url: string;
  isDe?: boolean;
  /** 'button' = grüner Knopf (Detail-Ansichten), 'link' = kompakte Zeile (Listen). */
  variant?: 'button' | 'link';
  style?: React.CSSProperties;
}

export function TeamsJoinButton({ url, isDe = true, variant = 'button', style }: TeamsJoinButtonProps): React.ReactElement | null {
  if (!url) return null;
  const label = isDe ? 'An Teams-Besprechung teilnehmen' : 'Join Teams meeting';
  if (variant === 'link') {
    return (
      <a
        href={url}
        target="_blank"
        rel="noopener noreferrer"
        onClick={(e) => e.stopPropagation()}
        style={{
          display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.8rem', fontWeight: 600,
          color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'none', ...style,
        }}
      >
        <Video size={14} /> {label}
      </a>
    );
  }
  return (
    <a
      href={url}
      target="_blank"
      rel="noopener noreferrer"
      onClick={(e) => e.stopPropagation()}
      className="btn btn-primary"
      style={{
        display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.82rem',
        padding: '7px 14px', textDecoration: 'none', ...style,
      }}
    >
      <Video size={15} /> {label}
    </a>
  );
}

export default TeamsJoinButton;
