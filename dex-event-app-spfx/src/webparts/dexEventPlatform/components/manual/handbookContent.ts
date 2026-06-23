/**
 * Aggregator für alle Handbuch-Sektionen.
 *
 * Jede Sektion lebt in einer eigenen Datei unter ./sections/.
 * Reihenfolge hier = Reihenfolge in der Sidebar.
 */

import { ManualSection } from './types';
import { introSection } from './sections/intro';
import { tutorialSection } from './sections/tutorial';
import { findEventSection } from './sections/findEvent';
import { myEventsSection } from './sections/myEvents';
import { profileSection } from './sections/profile';
import { createEventSection } from './sections/createEvent';
import { editEventSection } from './sections/editEvent';
import { subEventsSection } from './sections/subEvents';
import { b2runSection } from './sections/b2run';
import { manageParticipantsSection } from './sections/manageParticipants';
import { actionsSection } from './sections/actions';
import { registerForOtherSection } from './sections/registerForOther';
import { teamRegistrationSection } from './sections/teamRegistration';
import { checkInSection } from './sections/checkIn';
import { selfCheckInSection } from './sections/selfCheckIn';
import { quizSection } from './sections/quiz';
import { massMailSection } from './sections/massMail';
import { inviteMailSection } from './sections/inviteMail';
import { outlookUpdateSection } from './sections/outlookUpdate';
import { rolesSection } from './sections/roles';
import { templatesSection } from './sections/templates';
import { idReorderSection } from './sections/idReorder';
import { waitlistSection } from './sections/waitlist';
import { flowsSection } from './sections/flows';
import { faqSection } from './sections/faq';
import { demoImpersonationSection } from './sections/demoImpersonation';
import { templatesReseedSection } from './sections/templatesReseed';
import { peoplePickerScopeSection } from './sections/peoplePickerScope';

export type Locale = 'de' | 'en';

export function getManualSections(locale: Locale): ManualSection[] {
  // Aktuell deutsch. English-Overrides kommen pro Sektion (locale-Parameter wird
  // weitergereicht, damit einzelne Sektionen zweisprachig zurückgeben können).
  return [
    introSection(locale),
    tutorialSection(locale),
    findEventSection(locale),
    myEventsSection(locale),
    profileSection(locale),
    checkInSection(locale),
    selfCheckInSection(locale),
    createEventSection(locale),
    editEventSection(locale),
    subEventsSection(locale),
    b2runSection(locale),
    manageParticipantsSection(locale),
    actionsSection(locale),
    registerForOtherSection(locale),
    teamRegistrationSection(locale),
    quizSection(locale),
    massMailSection(locale),
    inviteMailSection(locale),
    outlookUpdateSection(locale),
    rolesSection(locale),
    peoplePickerScopeSection(locale),
    demoImpersonationSection(locale),
    templatesSection(locale),
    templatesReseedSection(locale),
    idReorderSection(locale),
    waitlistSection(locale),
    flowsSection(locale),
    faqSection(locale),
  ];
}
