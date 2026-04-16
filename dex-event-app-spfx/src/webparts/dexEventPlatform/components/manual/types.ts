/**
 * Datenstruktur fuer das In-App-Handbuch.
 *
 * Pro Workflow gibt es bis zu drei Perspektiven (Organizer / User / Admin),
 * jeweils mit nummerierten Schritten. Pro Schritt kann ein Live-Render
 * (echte Komponente mit Mock-Daten), ein SVG-Mockup oder ein Screenshot-Slot
 * eingebunden werden.
 */

import * as React from 'react';

export type ManualPerspective = 'organizer' | 'user' | 'admin';
export type ManualCategory = 'general' | 'organizer' | 'admin' | 'architecture';

export interface ManualStep {
  /** 1-basierte Schrittnummer (zeigt sich als grosse Zahl im UI) */
  number: number;
  title: string;
  /** Beschreibung — darf JSX enthalten (Code-Snippets, Links, Hervorhebungen) */
  description: React.ReactNode;
  /** Optional: ein Live-Render der echten App-Komponente mit Mock-Daten */
  liveDemo?: React.ReactNode;
  /** Optional: ein SVG/CSS-Mockup mit Click-Pfaden, Pfeilen, Highlights */
  mockup?: React.ReactNode;
  /** Optional: Pfad zu einem echten Screenshot — solange leer wird der Slot
   *  als gestylte Box mit "Screenshot folgt" Platzhalter angezeigt */
  screenshotPath?: string;
  /** Beschreibung was im Screenshot/Mockup zu sehen sein soll */
  visualHint?: string;
  /** Optional: gelber Tipp-Kasten unter dem Schritt */
  tip?: string;
  /** Optional: roter Warn-Kasten unter dem Schritt */
  warning?: string;
}

export interface ManualPerspectiveBlock {
  perspective: ManualPerspective;
  /** Optionaler Untertitel — Default: "Als Organizer:in" / "Als User:in" / "Als Admin:in" */
  title?: string;
  /** Kurze Einleitung vor den Schritten */
  intro?: React.ReactNode;
  steps: ManualStep[];
}

export interface ManualSection {
  /** Stable URL-Slug (fuer Anker-Links + Suche) */
  id: string;
  title: string;
  category: ManualCategory;
  /** 1-2 Saetze Intro (wird in der Sidebar als Subtitle gezeigt) */
  description: string;
  /** Welche Rollen sehen diese Sektion in der Sidebar */
  visibleFor: Array<'User' | 'Organizer' | 'Admin'>;
  perspectives: ManualPerspectiveBlock[];
}
