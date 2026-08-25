/**
 * v28.94: Aus `EventCreationPage` herausgeloest. Vorschlagstexte für die
 * Event-Beschreibung — reine Daten, die im Editor-Dialog als Chips angeboten
 * werden (v28.7).
 */
export // v26.77: Fünf einladende Beschreibungs-Vorlagen als Starthilfe für Organizer.
// Bewusst OHNE Zeit, Ort, Organizer oder Kontaktperson — diese stehen bereits
// als eigene Felder auf der Anmeldemaske. Der Organizer kann eine Vorlage per
// Klick übernehmen und dann anpassen. Jede Vorlage in DE + EN.
const DESCRIPTION_TEMPLATES: Array<{ key: string; labelDe: string; labelEn: string; de: string; en: string }> = [
  {
    key: 'sport',
    labelDe: 'Sport & Lauf',
    labelEn: 'Sports & running',
    de: '<p>Lauf mit uns! Sei Teil des Deloitte-Teams und erlebe einen sportlichen Tag mit Kolleginnen und Kollegen. Egal ob du für die Bestzeit gehst oder einfach die Stimmung genießen möchtest – gemeinsam sind wir am Start. Wir freuen uns riesig auf dich!</p>',
    en: '<p>Run with us! Be part of the Deloitte team and enjoy a sporty day with your colleagues. Whether you are chasing a personal best or simply soaking up the atmosphere – we are at the start line together. We can&rsquo;t wait to see you!</p>',
  },
  {
    key: 'workshop',
    labelDe: 'Workshop & Weiterbildung',
    labelEn: 'Workshop & training',
    de: '<p>In diesem Workshop dreht sich alles um praxisnahe Impulse für deinen Arbeitsalltag. Du bekommst konkrete Einblicke, tauschst dich mit anderen aus und nimmst neue Ideen mit nach Hause. Bring gerne deine Fragen und Erfahrungen mit – wir gestalten den Tag gemeinsam.</p>',
    en: '<p>This workshop is all about hands-on impulses for your daily work. You will gain concrete insights, exchange ideas with others and take fresh inspiration home with you. Bring your questions and experiences along – we will shape the day together.</p>',
  },
  {
    key: 'network',
    labelDe: 'Netzwerken & After Work',
    labelEn: 'Networking & after work',
    de: '<p>Zeit zum Netzwerken! In entspannter Atmosphäre kommen wir zusammen, tauschen uns aus und lernen neue Gesichter kennen. Freu dich auf gute Gespräche, den ein oder anderen Snack und einen lockeren Ausklang. Komm einfach vorbei – wir freuen uns auf dich!</p>',
    en: '<p>Time to connect! In a relaxed setting we come together, share ideas and meet new faces. Look forward to good conversations, a few snacks and an easy-going get-together. Just drop by – we would love to see you!</p>',
  },
  {
    key: 'conference',
    labelDe: 'Konferenz & Fachevent',
    labelEn: 'Conference & expert event',
    de: '<p>Freu dich auf einen Tag voller spannender Vorträge, praxisnaher Sessions und wertvoller Begegnungen. Expertinnen und Experten teilen ihr Wissen, und du hast reichlich Gelegenheit, dich zu vernetzen und neue Perspektiven mitzunehmen.</p>',
    en: '<p>Look forward to a day full of inspiring talks, hands-on sessions and valuable encounters. Experts share their knowledge, and you will have plenty of opportunity to network and gain fresh perspectives.</p>',
  },
  {
    key: 'celebration',
    labelDe: 'Team-Event & Feier',
    labelEn: 'Team event & celebration',
    de: '<p>Wir haben etwas zu feiern – und du bist herzlich eingeladen! Freu dich auf einen geselligen Abend mit gutem Essen, netten Menschen und bester Stimmung. Lass uns gemeinsam eine richtig schöne Zeit verbringen. Wir freuen uns auf dich!</p>',
    en: '<p>We&rsquo;ve got something to celebrate – and you are warmly invited! Look forward to a sociable evening with good food, great people and a wonderful atmosphere. Let&rsquo;s enjoy some quality time together. We can&rsquo;t wait to see you!</p>',
  },
];
