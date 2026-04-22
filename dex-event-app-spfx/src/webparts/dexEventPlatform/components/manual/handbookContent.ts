/**
 * Aggregator fuer alle Handbuch-Sektionen.
 *
 * Jede Sektion lebt in einer eigenen Datei unter ./sections/.
 * Reihenfolge hier = Reihenfolge in der Sidebar.
 */

import { ManualSection } from './types';
import { introSection } from './sections/intro';
import { findEventSection } from './sections/findEvent';
import { myEventsSection } from './sections/myEvents';
import { profileSection } from './sections/profile';
import { createEventSection } from './sections/createEvent';
import { editEventSection } from './sections/editEvent';
import { subEventsSection } from './sections/subEvents';
import { b2runSection } from './sections/b2run';
import { manageParticipantsSection } from './sections/manageParticipants';
import { registerForOtherSection } from './sections/registerForOther';
import { checkInSection } from './sections/checkIn';
import { quizSection } from './sections/quiz';
import { massMailSection } from './sections/massMail';
import { rolesSection } from './sections/roles';
import { templatesSection } from './sections/templates';
import { idReorderSection } from './sections/idReorder';
import { waitlistSection } from './sections/waitlist';
import { flowsSection } from './sections/flows';
import { faqSection } from './sections/faq';

export type Locale = 'de' | 'en';

export function getManualSections(locale: Locale): ManualSection[] {
  // Aktuell deutsch. English-Overrides kommen pro Sektion (locale-Parameter wird
  // weitergereicht, damit einzelne Sektionen zweisprachig zurueckgeben koennen).
  return [
    introSection(locale),
    findEventSection(locale),
    myEventsSection(locale),
    profileSection(locale),
    checkInSection(locale),
    createEventSection(locale),
    editEventSection(locale),
    subEventsSection(locale),
    b2runSection(locale),
    manageParticipantsSection(locale),
    registerForOtherSection(locale),
    quizSection(locale),
    massMailSection(locale),
    rolesSection(locale),
    templatesSection(locale),
    idReorderSection(locale),
    waitlistSection(locale),
    flowsSection(locale),
    faqSection(locale),
  ];
}
