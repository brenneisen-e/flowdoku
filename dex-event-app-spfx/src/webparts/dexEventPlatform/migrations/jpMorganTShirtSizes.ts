/**
 * One-Shot-Migration: T-Shirt-Größen für JP-Morgan-Event-Teilnehmer.
 *
 * Stand: 2026-04-22 — die T-Shirt-Größen wurden ursprünglich in einem
 * externen Tool (B2Run-Portal) gepflegt und sind in der DEX-Teilnehmerliste
 * nicht gelandet. Diese Migration läuft beim Admin-App-Start einmalig durch
 * und trägt die Größen pro Teilnehmer (E-Mail-Matching, case-insensitive)
 * als Custom-Field-Wert + CustomData-JSON-Eintrag nach.
 *
 * Idempotent: wenn ein Teilnehmer bereits einen Wert hat, wird er NICHT
 * überschrieben. Beim nächsten Admin-Start werden die noch-leeren Einträge
 * nachgetragen.
 *
 * Matching:
 *  - Event: Title matcht /jp\s*morgan|jpmorgan/i und hat eine Subsite
 *  - Custom-Field: label matcht /t-?shirt|shirt.*(groe|grö)sse/i → liefert spInternalName
 *  - Teilnehmer: ParticipantEmail (lowercase) == CSV-email
 */

import { EventService, SPRegistration } from '../services/EventService';
import { DeloitteEvent } from '../types';

interface TShirtEntry {
  email: string;
  size: string;
}

/** CSV-Daten, generiert aus der Teilnehmerliste-2026-04-22T131044.770.csv. */
const JP_MORGAN_TSHIRT_DATA: TShirtEntry[] = [
  { email: "nifelten@deloitte.de", size: "XS (Unisex)" },
  { email: "tluedke@deloitte.de", size: "Crew neck - L" },
  { email: "chgraf@deloitte.de", size: "Crew neck - L" },
  { email: "marold@deloitte.de", size: "Crew neck - M" },
  { email: "hafuhrmann@deloitte.de", size: "Crew neck - M" },
  { email: "nhquach@deloitte.de", size: "Crew neck - S" },
  { email: "sdenter@deloitte.de", size: "Crew neck - M" },
  { email: "jkissner@deloitte.de", size: "Crew neck - S" },
  { email: "juhummel@deloitte.de", size: "Crew neck - L" },
  { email: "yguan@deloitte.de", size: "Crew neck - M" },
  { email: "tarmbrust@deloitte.de", size: "Crew neck - XL" },
  { email: "ngilli@deloitte.de", size: "Crew neck - M" },
  { email: "ltrupp@deloitte.de", size: "Crew neck - S" },
  { email: "imilic@deloitte.de", size: "Crew neck - M" },
  { email: "dsumm@deloitte.de", size: "Crew neck - L" },
  { email: "skimmel@deloitte.de", size: "Crew neck - S" },
  { email: "ateusch@deloitte.de", size: "Crew neck - L" },
  { email: "lschuerger@deloitte.de", size: "Crew neck - L" },
  { email: "kwagner@deloitte.de", size: "Crew neck - S" },
  { email: "cboesch@deloitte.de", size: "Crew neck - L" },
  { email: "fkansoy@deloitte.de", size: "Crew neck - S" },
  { email: "tademi@deloitte.de", size: "Crew neck - M" },
  { email: "akurtulmus@deloitte.de", size: "Crew neck - S" },
  { email: "noffen@deloitte.de", size: "Crew neck - S" },
  { email: "bogconstantinescu@deloitte.de", size: "Crew neck - L" },
  { email: "vkraynyukov@deloitte.de", size: "Crew neck - XL" },
  { email: "shoffmeister@deloitte.de", size: "Crew neck - L" },
  { email: "dpalaciostorri@deloitte.de", size: "Crew neck - XL" },
  { email: "sfroemel@deloitte.de", size: "Crew neck - S" },
  { email: "ehauschild@deloitte.de", size: "Crew neck - L" },
  { email: "hoezcan@deloitte.de", size: "Crew neck - M" },
  { email: "smomberg@deloitte.de", size: "Crew neck - S" },
  { email: "cberndt@deloitte.de", size: "Crew neck - L" },
  { email: "thituynguyen@deloitte.de", size: "Crew neck - M" },
  { email: "anrapp@deloitte.de", size: "Crew neck - M" },
  { email: "fmeurer@deloitte.de", size: "Crew neck - XL" },
  { email: "michuynh@deloitte.de", size: "Crew neck - M" },
  { email: "milison@deloitte.de", size: "Crew neck - XL" },
  { email: "thundt@deloitte.de", size: "Crew neck - L" },
  { email: "dbonfadini@deloitte.de", size: "Crew neck - L" },
  { email: "aschiffner@deloitte.de", size: "Crew neck - L" },
  { email: "morschneider@deloitte.de", size: "Crew neck - L" },
  { email: "chuesing@deloitte.de", size: "Crew neck - M" },
  { email: "joherzig@deloitte.de", size: "Crew neck - S" },
  { email: "bermueller@deloitte.de", size: "Crew neck - M" },
  { email: "kahorvathova@deloitte.de", size: "Crew neck - M" },
  { email: "soodkim@deloitte.de", size: "Crew neck - L" },
  { email: "aweizmann@deloitte.de", size: "Crew neck - M" },
  { email: "kbassmann@deloitte.de", size: "Crew neck - M" },
  { email: "dvogler@deloitte.de", size: "Crew neck - M" },
  { email: "anburger@deloitte.de", size: "Crew neck - L" },
  { email: "aparshin@deloitte.de", size: "Crew neck - L" },
  { email: "skamps@deloitte.de", size: "Crew neck - XL" },
  { email: "yasilhan@deloitte.de", size: "Crew neck - M" },
  { email: "ianton@deloitte.de", size: "Crew neck - M" },
  { email: "tgoetze@deloitte.de", size: "Crew neck - L" },
  { email: "cstrauss@deloitte.de", size: "Crew neck - L" },
  { email: "astrek@deloitte.de", size: "Crew neck - M" },
  { email: "ledavies@deloitte.de", size: "Crew neck - M" },
  { email: "mtempetrini@deloitte.de", size: "Crew neck - S" },
  { email: "zaikemu@deloitte.de", size: "Crew neck - M" },
  { email: "ageis@deloitte.de", size: "Crew neck - XL" },
  { email: "tkiritharan@deloitte.de", size: "Crew neck - M" },
  { email: "senink@deloitte.de", size: "Crew neck - M" },
  { email: "tmerz@deloitte.de", size: "Crew neck - S" },
  { email: "akalvin@deloitte.de", size: "Crew neck - M" },
  { email: "nullrich@deloitte.de", size: "Crew neck - L" },
  { email: "lebeerboom@deloitte.de", size: "Crew neck - M" },
  { email: "joewunderlich@deloitte.de", size: "Crew neck - M" },
  { email: "sykorthof@deloitte.de", size: "Crew neck - L" },
  { email: "avigneswaran@deloitte.de", size: "Crew neck - XL" },
  { email: "babdouh@deloitte.de", size: "Crew neck - M" },
  { email: "nvonjagow@deloitte.de", size: "Crew neck - L" },
  { email: "tschulte@deloitte.de", size: "Crew neck - S" },
  { email: "daheck@deloitte.de", size: "Crew neck - XL" },
  { email: "davitran@deloitte.de", size: "Crew neck - L" },
  { email: "bzizah@deloitte.de", size: "Crew neck - L" },
  { email: "utanyildizi@deloitte.de", size: "Crew neck - M" },
  { email: "melidrissi@deloitte.de", size: "Crew neck - M" },
  { email: "mesmeenk@deloitte.de", size: "Crew neck - M" },
  { email: "kbazgar@deloitte.de", size: "Crew neck - M" },
  { email: "emartinezrengifo@deloitte.de", size: "Crew neck - S" },
  { email: "hangao@deloitte.de", size: "Crew neck - M" },
  { email: "bkiefer@deloitte.de", size: "Crew neck - L" },
  { email: "jenzel@deloitte.de", size: "Crew neck - S" },
  { email: "fhetterich@deloitte.de", size: "Crew neck - XL" },
  { email: "sel-menshawy@deloitte.de", size: "Crew neck - L" },
  { email: "andangra@deloitte.de", size: "Crew neck - S" },
  { email: "malbrecht@deloitte.de", size: "Crew neck - M" },
  { email: "jlembke@deloitte.de", size: "Crew neck - XL" },
  { email: "dknoebl@deloitte.de", size: "Crew neck - S" },
  { email: "rmeyer-gohde@deloitte.de", size: "Crew neck - S" },
  { email: "lreussner@deloitte.de", size: "Crew neck - M" },
  { email: "schami@deloitte.de", size: "Crew neck - S" },
  { email: "jpohlmann@deloitte.de", size: "Crew neck - M" },
  { email: "kebert@deloitte.de", size: "Crew neck - M" },
  { email: "salemi@deloitte.de", size: "Crew neck - M" },
  { email: "puzomah@deloitte.de", size: "Crew neck - S" },
  { email: "tsichler-reshamvala@deloitte.de", size: "Crew neck - S" },
  { email: "zizhao@deloitte.de", size: "Crew neck - M" },
  { email: "halt@deloitte.de", size: "Crew neck - M" },
  { email: "aakguen@deloitte.de", size: "Crew neck - M" },
  { email: "lhundt@deloitte.de", size: "Crew neck - M" },
  { email: "oabudaieh@deloitte.de", size: "Crew neck - XL" },
  { email: "krosenthalgabrol@deloitte.de", size: "Crew neck - M" },
  { email: "abinakaj@deloitte.de", size: "Crew neck - S" },
  { email: "alutsenko@deloitte.de", size: "Crew neck - M" },
  { email: "vgerlach@deloitte.de", size: "Crew neck - L" },
  { email: "mmoktad@deloitte.de", size: "Crew neck - S" },
  { email: "luriebeling@deloitte.de", size: "Crew neck - L" },
  { email: "kevhuynh@deloitte.de", size: "Crew neck - M" },
  { email: "llowak@deloitte.de", size: "Crew neck - XL" },
  { email: "jhertstein@deloitte.de", size: "Crew neck - XL" },
  { email: "jjusfin@deloitte.de", size: "Crew neck - M" },
  { email: "tohagemann@deloitte.de", size: "Crew neck - XL" },
  { email: "suzkingston@deloitte.de", size: "Crew neck - S" },
  { email: "mihofmann@deloitte.de", size: "Crew neck - M" },
  { email: "skirchhoff@deloitte.de", size: "Crew neck - XL" },
  { email: "nschindler@deloitte.de", size: "Crew neck - L" },
  { email: "sileckhardt@deloitte.de", size: "Crew neck - L" },
  { email: "jukretz@deloitte.de", size: "Crew neck - M" },
  { email: "kmengistu@deloitte.de", size: "Crew neck - L" },
  { email: "markremer@deloitte.de", size: "Crew neck - M" },
  { email: "ischefer@deloitte.de", size: "Crew neck - M" },
  { email: "jbeaufort@deloitte.de", size: "Crew neck - L" },
  { email: "shass@deloitte.de", size: "Crew neck - M" },
  { email: "sgruttmann@deloitte.de", size: "Crew neck - S" },
  { email: "anemmert@deloitte.de", size: "Crew neck - M" },
  { email: "fjohanna@deloitte.de", size: "Crew neck - S" },
  { email: "ljahn@deloitte.de", size: "Crew neck - XL" },
  { email: "dkarrenberg@deloitte.de", size: "Crew neck - M" },
  { email: "pkremer@deloitte.de", size: "Crew neck - XL" },
  { email: "sadammann@deloitte.de", size: "Crew neck - M" },
  { email: "jzeiss@deloitte.de", size: "Crew neck - L" },
  { email: "sversen@deloitte.de", size: "Crew neck - S" },
  { email: "jkim53@deloitte.de", size: "Crew neck - XL" },
  { email: "yaiqbal@deloitte.de", size: "Crew neck - M" },
  { email: "bnegussie@deloitte.de", size: "Crew neck - M" },
  { email: "gpaschalidis@deloitte.de", size: "Crew neck - S" },
  { email: "ylackas@deloitte.de", size: "Crew neck - M" },
  { email: "febert@deloitte.de", size: "Crew neck - M" },
  { email: "tfalkenstein@deloitte.de", size: "Crew neck - L" },
  { email: "jakammerer@deloitte.de", size: "Crew neck - M" },
  { email: "vrothenbaecher@deloitte.de", size: "Crew neck - M" },
  { email: "ahnenna@deloitte.de", size: "Crew neck - M" },
  { email: "vkoshevchuk@deloitte.de", size: "Crew neck - S" },
  { email: "minh-le@deloitte.de", size: "Crew neck - L" },
  { email: "nshahidi@deloitte.de", size: "Crew neck - M" },
  { email: "dmkorn@deloitte.de", size: "Crew neck - S" },
  { email: "kimueller@deloitte.de", size: "Crew neck - M" },
  { email: "shmirza@deloitte.de", size: "Crew neck - L" },
  { email: "akaletsch@deloitte.de", size: "Crew neck - S" },
  { email: "sischmidt@deloitte.de", size: "Crew neck - L" },
  { email: "jtafferner@deloitte.de", size: "Crew neck - L" },
  { email: "eyoumbinzetchouang@deloitte.de", size: "Crew neck - L" },
  { email: "parpandey@deloitte.de", size: "Crew neck - XL" },
  { email: "benjaschmitt@deloitte.de", size: "Crew neck - XL" },
  { email: "annaweber@deloitte.de", size: "Crew neck - M" },
  { email: "mscholand@deloitte.de", size: "Crew neck - S" },
  { email: "chrengel@deloitte.de", size: "Crew neck - L" },
  { email: "cjoosten@deloitte.de", size: "Crew neck - S" },
  { email: "hainzhang@deloitte.de", size: "Crew neck - L" },
  { email: "mroon@deloitte.de", size: "Crew neck - S" },
  { email: "thentschel@deloitte.de", size: "Crew neck - XL" },
  { email: "prdas@deloitte.de", size: "Crew neck - L" },
  { email: "lhaist@deloitte.de", size: "Crew neck - M" },
  { email: "peteng@deloitte.de", size: "Crew neck - M" },
  { email: "smatitawaer@deloitte.de", size: "Crew neck - XL" },
  { email: "fastein@deloitte.de", size: "Crew neck - XL" },
  { email: "lbilsing@deloitte.de", size: "Crew neck - XL" },
  { email: "pkehder@deloitte.de", size: "Crew neck - M" },
  { email: "kvasilakis@deloitte.de", size: "Crew neck - M" },
  { email: "lakat@deloitte.de", size: "Crew neck - S" },
  { email: "bgloede@deloitte.de", size: "Crew neck - L" },
  { email: "maxschneider@deloitte.de", size: "Crew neck - M" },
  { email: "nicweber@deloitte.de", size: "Crew neck - L" },
  { email: "gautalwar@deloitte.de", size: "Crew neck - XL" },
  { email: "dturano@deloitte.de", size: "Crew neck - S" },
  { email: "azhamharyan@deloitte.de", size: "Crew neck - M" },
  { email: "marischaefer@deloitte.de", size: "Crew neck - S" },
  { email: "svthomas@deloitte.de", size: "Crew neck - L" },
  { email: "haozzhao@deloitte.de", size: "Crew neck - L" },
  { email: "bschulz@deloitte.de", size: "Crew neck - M" },
  { email: "niweber@deloitte.de", size: "Crew neck - M" },
  { email: "mabuehler@deloitte.de", size: "Crew neck - M" },
  { email: "aivankovic@deloitte.de", size: "Crew neck - M" },
  { email: "tbigge@deloitte.de", size: "Crew neck - M" },
  { email: "pteichelmann@deloitte.de", size: "Crew neck - L" },
  { email: "npetersen@deloitte.de", size: "Crew neck - M" },
  { email: "cnagel@deloitte.de", size: "Crew neck - S" },
  { email: "lsailer@deloitte.de", size: "Crew neck - M" },
  { email: "lwildenhof@deloitte.de", size: "Crew neck - M" },
  { email: "mgrzyb@deloitte.de", size: "Crew neck - L" },
  { email: "trehnig@deloitte.de", size: "Crew neck - L" },
  { email: "ffriedel@deloitte.de", size: "Crew neck - M" },
  { email: "rmagasy@deloitte.de", size: "Crew neck - L" },
  { email: "anbeier@deloitte.de", size: "Crew neck - L" },
  { email: "fkoepf@deloitte.de", size: "Crew neck - L" },
  { email: "sahofmann@deloitte.de", size: "Crew neck - S" },
  { email: "ldankic@deloitte.de", size: "Crew neck - L" },
  { email: "agoeller@deloitte.de", size: "Crew neck - L" },
  { email: "rebergmann@deloitte.de", size: "Crew neck - M" },
  { email: "pamueller@deloitte.de", size: "Crew neck - S" },
  { email: "lasobottka@deloitte.de", size: "Crew neck - S" },
  { email: "ngross@deloitte.de", size: "Crew neck - M" },
  { email: "lujost@deloitte.de", size: "Crew neck - L" },
  { email: "szielinski@deloitte.de", size: "Crew neck - M" },
  { email: "hannjung@deloitte.de", size: "Crew neck - M" },
  { email: "mmoller-racke@deloitte.de", size: "Crew neck - L" },
  { email: "lkontag@deloitte.de", size: "Crew neck - S" },
  { email: "pdlugaj@deloitte.de", size: "Crew neck - L" },
  { email: "mhilgenhaus@deloitte.de", size: "Crew neck - XL" },
  { email: "nneukirch@deloitte.de", size: "Crew neck - L" },
  { email: "arnasarin@deloitte.de", size: "Crew neck - M" },
  { email: "leheld@deloitte.de", size: "Crew neck - M" },
  { email: "damarino@deloitte.de", size: "Crew neck - S" },
  { email: "vwotke@deloitte.de", size: "Crew neck - XL" },
  { email: "yscheu@deloitte.de", size: "Crew neck - L" },
  { email: "eguenay@deloitte.de", size: "Crew neck - M" },
  { email: "swelker@deloitte.de", size: "Crew neck - S" },
  { email: "spodeyn@deloitte.de", size: "Crew neck - S" },
  { email: "agrigutis@deloitte.de", size: "Crew neck - M" },
  { email: "chheinze@deloitte.de", size: "Crew neck - L" },
  { email: "demtran@deloitte.de", size: "Crew neck - S" },
  { email: "eholzapfel@deloitte.de", size: "Crew neck - S" },
  { email: "juwagner@deloitte.de", size: "Crew neck - XL" },
  { email: "rrodriguesdossantos@deloitte.de", size: "Crew neck - L" },
  { email: "mrockel@deloitte.de", size: "Crew neck - L" },
  { email: "sbolte@deloitte.de", size: "Crew neck - M" },
  { email: "patpaul@deloitte.de", size: "Crew neck - M" },
  { email: "lweber@deloitte.de", size: "Crew neck - M" },
  { email: "cburgard@deloitte.de", size: "Crew neck - M" },
  { email: "ssiebold@deloitte.de", size: "Crew neck - L" },
  { email: "swessling@deloitte.de", size: "Crew neck - S" },
  { email: "tgoeb@deloitte.de", size: "Crew neck - XL" },
  { email: "ahoti@deloitte.de", size: "Crew neck - XL" },
  { email: "ldenic@deloitte.de", size: "Crew neck - L" },
  { email: "etorun@deloitte.de", size: "Crew neck - S" },
  { email: "arodriguezgonzalez@deloitte.de", size: "Crew neck - L" },
  { email: "jkowasch@deloitte.de", size: "Crew neck - L" },
  { email: "lvasiljevic@deloitte.de", size: "Crew neck - L" },
  { email: "tschaefer@deloitte.de", size: "Crew neck - XL" },
  { email: "ckunze@deloitte.de", size: "Crew neck - L" },
  { email: "flueck@deloitte.de", size: "Crew neck - M" },
  { email: "epirdal@deloitte.de", size: "Crew neck - XL" },
  { email: "alkuehne@deloitte.de", size: "Crew neck - L" },
  { email: "jwickenhagen@deloitte.de", size: "Crew neck - L" },
  { email: "gshamkin@deloitte.de", size: "Crew neck - L" },
  { email: "najordan@deloitte.de", size: "Crew neck - S" },
  { email: "jreuss@deloitte.de", size: "Crew neck - M" },
  { email: "hklenk@deloitte.de", size: "Crew neck - S" },
  { email: "fardian@deloitte.de", size: "Crew neck - L" },
  { email: "ndoersam@deloitte.de", size: "Crew neck - L" },
  { email: "cjansenvanvuuren@deloitte.de", size: "Crew neck - M" },
  { email: "julikaiser@deloitte.de", size: "Crew neck - L" },
  { email: "jherborn@deloitte.de", size: "Crew neck - S" },
  { email: "aleite-gross@deloitte.de", size: "Crew neck - S" },
  { email: "jknauf@deloitte.de", size: "Crew neck - L" },
  { email: "tstay@deloitte.de", size: "Crew neck - XL" },
  { email: "qehring@deloitte.de", size: "Crew neck - XL" },
  { email: "ckressel@deloitte.de", size: "Crew neck - S" },
  { email: "bvoelker@deloitte.de", size: "Crew neck - L" },
  { email: "dmoldavski@deloitte.de", size: "Crew neck - M" },
  { email: "fdoenmez@deloitte.de", size: "Crew neck - L" },
  { email: "leasnyder@deloitte.de", size: "Crew neck - S" },
  { email: "marnoack@deloitte.de", size: "Crew neck - L" },
  { email: "ckiobel@deloitte.de", size: "Crew neck - M" },
  { email: "lfilobok@deloitte.de", size: "Crew neck - S" },
  { email: "jmundy@deloitte.de", size: "Crew neck - XL" },
  { email: "smausolf@deloitte.de", size: "Crew neck - S" },
  { email: "frabus@deloitte.de", size: "Crew neck - M" },
  { email: "lbaumert@deloitte.de", size: "Crew neck - M" },
  { email: "lbeck@deloitte.de", size: "Crew neck - XL" },
  { email: "arusche@deloitte.de", size: "Crew neck - M" },
  { email: "rkureck@deloitte.de", size: "Crew neck - S" },
  { email: "mglazkova@deloitte.de", size: "Crew neck - M" },
  { email: "lelanger@deloitte.de", size: "Crew neck - L" },
  { email: "pklemp@deloitte.de", size: "Crew neck - L" },
  { email: "bgoetz@deloitte.de", size: "Crew neck - L" },
  { email: "dbrunner@deloitte.de", size: "Crew neck - XL" },
  { email: "mpinnecke@deloitte.de", size: "Crew neck - S" },
  { email: "junfeng@deloitte.de", size: "Crew neck - S" },
  { email: "rflach@deloitte.de", size: "Crew neck - L" },
  { email: "tfranzke@deloitte.de", size: "Crew neck - M" },
  { email: "prkhairnar@deloitte.de", size: "Crew neck - M" },
  { email: "jheinen@deloitte.de", size: "Crew neck - L" },
  { email: "nschwiertz@deloitte.de", size: "Crew neck - S" },
  { email: "krimueller@deloitte.de", size: "Crew neck - S" },
  { email: "flormueller@deloitte.de", size: "Crew neck - L" },
  { email: "yascheve@deloitte.de", size: "Crew neck - XL" },
  { email: "habraham@deloitte.de", size: "Crew neck - S" },
  { email: "alhaupt@deloitte.de", size: "Crew neck - M" },
  { email: "ivogtmendez@deloitte.de", size: "Crew neck - S" },
  { email: "lkantim@deloitte.de", size: "Crew neck - L" },
  { email: "lflagmansky@deloitte.de", size: "Crew neck - L" },
  { email: "patheller@deloitte.de", size: "Crew neck - M" },
  { email: "jvondosky@deloitte.de", size: "Crew neck - L" },
  { email: "sschillinger@deloitte.de", size: "Crew neck - L" },
  { email: "sacbaid@deloitte.de", size: "Crew neck - S" },
  { email: "blinnertz@deloitte.de", size: "Crew neck - M" },
  { email: "mjandl@deloitte.de", size: "Crew neck - M" },
  { email: "jhass@deloitte.de", size: "Crew neck - XL" },
  { email: "soscheepers@deloitte.de", size: "Crew neck - S" },
  { email: "lderibo@deloitte.de", size: "Crew neck - L" },
  { email: "mtellezperez@deloitte.de", size: "Crew neck - L" },
  { email: "dabremmer@deloitte.de", size: "Crew neck - XL" },
  { email: "mfelline@deloitte.de", size: "Crew neck - M" },
  { email: "amastilovic@deloitte.de", size: "Crew neck - XL" },
  { email: "hatuhcic@deloitte.de", size: "Crew neck - L" },
  { email: "jreinholz@deloitte.de", size: "Crew neck - L" },
  { email: "dwessel@deloitte.de", size: "Crew neck - XL" },
  { email: "cblumenthal@deloitte.de", size: "Crew neck - L" },
  { email: "trickert@deloitte.de", size: "Crew neck - M" },
  { email: "qduong@deloitte.de", size: "Crew neck - S" },
  { email: "kthierbach@deloitte.de", size: "Crew neck - S" },
  { email: "akissner@deloitte.de", size: "Crew neck - M" },
  { email: "srupp@deloitte.de", size: "Crew neck - M" },
  { email: "tstay@deloitte.de", size: "Crew neck - XL" },
  { email: "ascholl@deloitte.de", size: "Crew neck - L" },
  { email: "mwondrak@deloitte.de", size: "Crew neck - S" },
  { email: "tpranjic@deloitte.de", size: "Crew neck - S" },
  { email: "mmorschett@deloitte.de", size: "Crew neck - S" },
  { email: "cengler@deloitte.de", size: "Crew neck - M" },
  { email: "kkalkar@deloitte.de", size: "Crew neck - L" },
  { email: "vsivarajahkumar@deloitte.de", size: "Crew neck - M" },
  { email: "jgathof@deloitte.de", size: "Crew neck - L" },
  { email: "mjeanneaux@deloitte.de", size: "Crew neck - M" },
  { email: "manbhutada@deloitte.de", size: "Crew neck - M" },
  { email: "fknoch@deloitte.de", size: "Crew neck - L" },
  { email: "sroghe@deloitte.de", size: "Crew neck - S" },
  { email: "janwind@deloitte.de", size: "Crew neck - M" },
  { email: "manraj@deloitte.de", size: "Crew neck - L" },
  { email: "aelter@deloitte.de", size: "Crew neck - L" },
  { email: "vjander@deloitte.de", size: "Crew neck - M" },
  { email: "svincent@deloitte.de", size: "Crew neck - S" },
  { email: "lkalaycioglu@deloitte.de", size: "Crew neck - S" },
  { email: "tstorz@deloitte.de", size: "Crew neck - L" },
  { email: "dpereirajimenez@deloitte.de", size: "Crew neck - XL" },
  { email: "saghili@deloitte.de", size: "Crew neck - L" },
  { email: "clclassen@deloitte.de", size: "Crew neck - L" },
  { email: "kruppert@deloitte.de", size: "Crew neck - L" },
  { email: "kaekholm@deloitte.de", size: "Crew neck - S" },
  { email: "lmelzer@deloitte.de", size: "Crew neck - S" },
  { email: "asinitsyn@deloitte.de", size: "Crew neck - M" },
  { email: "vlatincic@deloitte.de", size: "Crew neck - M" },
  { email: "mtaeger@deloitte.de", size: "Crew neck - XL" },
  { email: "kczernitzki@deloitte.de", size: "Crew neck - M" },
  { email: "nikprajapati@deloitte.de", size: "Crew neck - XL" },
  { email: "rolambert@deloitte.de", size: "Crew neck - L" },
  { email: "mweidesequeira@deloitte.de", size: "Crew neck - S" },
  { email: "juliuschneider@deloitte.de", size: "Crew neck - L" },
  { email: "hlappe@deloitte.de", size: "Crew neck - XL" },
  { email: "fspiegel@deloitte.de", size: "Crew neck - L" },
  { email: "ppropadalo@deloitte.de", size: "Crew neck - XL" },
  { email: "atkeil@deloitte.de", size: "Crew neck - S" },
  { email: "sfoell@deloitte.de", size: "Crew neck - M" },
  { email: "tgowin@deloitte.de", size: "Crew neck - XL" },
  { email: "nhinrichsen@deloitte.de", size: "Crew neck - S" },
  { email: "fodekerken@deloitte.de", size: "Crew neck - L" },
  { email: "cueberschaer@deloitte.de", size: "Crew neck - M" },
  { email: "lkornherr@deloitte.de", size: "Crew neck - L" },
  { email: "srothballer@deloitte.de", size: "Crew neck - M" },
  { email: "ehuke@deloitte.de", size: "Crew neck - L" },
  { email: "ayoussoufi@deloitte.de", size: "Crew neck - XL" },
  { email: "jhauffortega@deloitte.de", size: "Crew neck - L" },
  { email: "clengel@deloitte.de", size: "Crew neck - L" },
  { email: "luweber@deloitte.de", size: "Crew neck - L" },
  { email: "tyusofi@deloitte.de", size: "Crew neck - XL" },
  { email: "fohahn@deloitte.de", size: "Crew neck - XL" },
  { email: "lreinhold@deloitte.de", size: "Crew neck - S" },
  { email: "kabeyer@deloitte.de", size: "Crew neck - S" },
  { email: "evardar@deloitte.de", size: "Crew neck - M" },
  { email: "ninalbars@deloitte.de", size: "Crew neck - S" },
  { email: "antschmitt@deloitte.de", size: "Crew neck - M" },
  { email: "tjames-schulz@deloitte.de", size: "Crew neck - L" },
  { email: "vduchan@deloitte.de", size: "Crew neck - L" },
  { email: "rdavletov@deloitte.de", size: "Crew neck - L" },
  { email: "sgimplinger@deloitte.de", size: "Crew neck - S" },
  { email: "mrasch@deloitte.de", size: "Crew neck - L" },
  { email: "jseeliger@deloitte.de", size: "Crew neck - S" },
  { email: "mikaschneider@deloitte.de", size: "Crew neck - L" },
  { email: "sshah-cashmeh@deloitte.de", size: "Crew neck - S" },
  { email: "dferenz@deloitte.de", size: "Crew neck - XL" },
  { email: "nschubart@deloitte.de", size: "Crew neck - S" },
  { email: "fschuwirth@deloitte.de", size: "Crew neck - M" },
  { email: "lblochmann@deloitte.de", size: "Crew neck - S" },
  { email: "ktosun@deloitte.de", size: "Crew neck - M" },
  { email: "asmakolli@deloitte.de", size: "Crew neck - S" },
  { email: "aueckermann@deloitte.de", size: "Crew neck - L" },
  { email: "junterberg@deloitte.de", size: "Crew neck - XL" },
  { email: "lsandner@deloitte.de", size: "Crew neck - L" },
  { email: "jpeschers@deloitte.de", size: "Crew neck - XL" },
  { email: "hgross@deloitte.de", size: "Crew neck - S" },
  { email: "fkamga-zadi@deloitte.de", size: "Crew neck - M" },
  { email: "lhertel@deloitte.de", size: "Crew neck - S" },
  { email: "lerling@deloitte.de", size: "Crew neck - S" },
  { email: "gdengl@deloitte.de", size: "Crew neck - XL" },
  { email: "kgjikopulli@deloitte.de", size: "M (Unisex)" },
  { email: "nachieng@deloitte.de", size: "S (Unisex)" },
  { email: "qialu@deloitte.de", size: "S (Unisex)" },
  { email: "jdaub@deloitte.de", size: "L (Unisex)" },
  { email: "parpandey@deloitte.de", size: "XL (Unisex)" },
  { email: "rdavletov@deloitte.de", size: "L (Unisex)" },
  { email: "nshahidi@deloitte.de", size: "M (Unisex)" },
  { email: "ahnenna@deloitte.de", size: "M (Unisex)" },
  { email: "asinitsyn@deloitte.de", size: "M (Unisex)" },
  { email: "sacbaid@deloitte.de", size: "S (Unisex)" },
  { email: "nihle@deloitte.de", size: "XL (Unisex)" },
  { email: "vkoshevchuk@deloitte.de", size: "S (Unisex)" },
  { email: "sadammann@deloitte.de", size: "M (Unisex)" },
  { email: "mleyson@deloitte.de", size: "M (Unisex)" },
  { email: "cmachwirth@deloitte.de", size: "L (Unisex)" },
  { email: "mascheykin@deloitte.de", size: "S (Unisex)" },
  { email: "lfinster@deloitte.de", size: "M (Unisex)" },
  { email: "thommueller@deloitte.de", size: "L (Unisex)" },
  { email: "sbadalova@deloitte.de", size: "S (Unisex)" },
  { email: "jbiehl@deloitte.de", size: "M (Unisex)" },
  { email: "tpfeifer@deloitte.de", size: "XL (Unisex)" },
  { email: "skumari61@deloitte.de", size: "L (Unisex)" },
  { email: "btempel@deloitte.de", size: "L (Unisex)" },
  { email: "siddhashah@deloitte.de", size: "S (Unisex)" },
  { email: "subhasdas@deloitte.de", size: "M (Unisex)" },
  { email: "ntseytlina@deloitte.de", size: "M (Unisex)" },
  { email: "twugk@deloitte.de", size: "S (Unisex)" },
  { email: "jaschmitt@deloitte.de", size: "M (Unisex)" },
  { email: "narpatel@deloitte.de", size: "S (Unisex)" },
  { email: "thuhvu@deloitte.de", size: "M (Unisex)" },
  { email: "lekeck@deloitte.de", size: "L (Unisex)" },
  { email: "nwissmann@deloitte.de", size: "L (Unisex)" },
  { email: "bhsridhar@deloitte.de", size: "L (Unisex)" },
  { email: "khuseynli@deloitte.de", size: "XL (Unisex)" },
  { email: "fwicaksono@deloitte.de", size: "XL (Unisex)" },
  { email: "ccloes@deloitte.de", size: "L (Unisex)" },
  { email: "lisjiang@deloitte.de", size: "M (Unisex)" },
  { email: "azniazi@deloitte.de", size: "M (Unisex)" },
  { email: "bboukayoua@deloitte.de", size: "L (Unisex)" },
  { email: "jhempelmann@deloitte.de", size: "M (Unisex)" },
  { email: "nhasyim@deloitte.de", size: "S (Unisex)" },
  { email: "hmorozov@deloitte.de", size: "M (Unisex)" },
  { email: "hfurmli@deloitte.de", size: "M (Unisex)" },
  { email: "aneck@deloitte.de", size: "S (Unisex)" },
  { email: "ricbhardwaj@deloitte.de", size: "M (Unisex)" },
  { email: "sovieth@deloitte.de", size: "L (Unisex)" },
  { email: "jstindt@deloitte.de", size: "M (Unisex)" },
  { email: "jbujnowska@deloitte.de", size: "S (Unisex)" },
  { email: "achekuchenko@deloitte.de", size: "M (Unisex)" },
  { email: "abaermann@deloitte.de", size: "XS (Unisex)" },
  { email: "mniltop@deloitte.de", size: "M (Unisex)" },
  { email: "yaowu@deloitte.de", size: "M (Unisex)" },
  { email: "kruder@deloitte.de", size: "M (Unisex)" },
  { email: "ishemeshsarel@deloitte.de", size: "L (Unisex)" },
  { email: "oschmidbauer@deloitte.de", size: "M (Unisex)" },
  { email: "ppikula@deloitte.de", size: "L (Unisex)" },
  { email: "micsingh@deloitte.de", size: "L (Unisex)" },
  { email: "wmasysjuk@deloitte.de", size: "L (Unisex)" },
  { email: "mraja-ahmad@deloitte.de", size: "M (Unisex)" },
  { email: "lbutzbach@deloitte.de", size: "L (Unisex)" },
  { email: "mkalbe@deloitte.de", size: "S (Unisex)" },
  { email: "sovcharova@deloitte.de", size: "M (Unisex)" },
  { email: "ekeser@deloitte.de", size: "L (Unisex)" },
  { email: "jogold@deloitte.de", size: "S (Unisex)" },
  { email: "svolkeri@deloitte.de", size: "XL (Unisex)" },
  { email: "nramadhanti@deloitte.de", size: "M (Unisex)" },
  { email: "jgauger@deloitte.de", size: "L (Unisex)" },
  { email: "amehahn@deloitte.de", size: "S (Unisex)" },
  { email: "mdobben@deloitte.de", size: "XL (Unisex)" },
  { email: "nschauf@deloitte.de", size: "M (Unisex)" },
  { email: "atuerker@deloitte.de", size: "M (Unisex)" },
  { email: "admaslov@deloitte.de", size: "M (Unisex)" },
  { email: "tallam@deloitte.de", size: "M (Unisex)" },
  { email: "ayushimaheshwari@deloitte.de", size: "M (Unisex)" },
  { email: "kneuland@deloitte.de", size: "M (Unisex)" },
  { email: "ccoppenhagen@deloitte.de", size: "L (Unisex)" },
  { email: "jvacietis@deloitte.de", size: "XL (Unisex)" },
  { email: "toschmidt@deloitte.de", size: "XL (Unisex)" },
  { email: "lwaffenschmidt@deloitte.de", size: "S (Unisex)" },
  { email: "kacimac@deloitte.de", size: "S (Unisex)" },
  { email: "asamvelyan@deloitte.de", size: "XS (Unisex)" },
  { email: "oroot@deloitte.de", size: "L (Unisex)" },
  { email: "alfaust@deloitte.de", size: "XS (Unisex)" },
  { email: "mabramato@deloitte.de", size: "L (Unisex)" },
  { email: "akrostewitz@deloitte.de", size: "S (Unisex)" },
  { email: "creissfelder@deloitte.de", size: "L (Unisex)" },
  { email: "fdinges@deloitte.de", size: "L (Unisex)" },
  { email: "theberlein@deloitte.de", size: "L (Unisex)" },
  { email: "karoy@deloitte.de", size: "M (Unisex)" },
  { email: "nkoeninger@deloitte.de", size: "L (Unisex)" },
  { email: "sogerner@deloitte.de", size: "S (Unisex)" },
  { email: "aabdyramanova@deloitte.de", size: "XS (Unisex)" },
];

export interface MigrationResult {
  jpEventsFound: number;
  registrationsChecked: number;
  updated: number;
  skippedAlreadySet: number;
  skippedNoMatch: number;
  errors: number;
}

export type MigrationProgressPhase = 'loading-event' | 'loading-registrations' | 'updating' | 'done' | 'skipped' | 'error';

export interface MigrationProgress {
  phase: MigrationProgressPhase;
  message: string;
  current: number;   // Anzahl bereits verarbeiteter Teilnehmer in diesem Event
  total: number;     // Gesamt-Teilnehmer des aktuellen Events
  updated: number;   // bisher tatsächlich aktualisiert
  eventTitle: string;
}

/**
 * Führt die Migration aus. Liest alle JP-Morgan-Events, findet pro Event das
 * T-Shirt-Custom-Field, und trägt pro Teilnehmer (E-Mail-Match) die Größe nach.
 * Gibt eine Statistik zurück.
 *
 * Gibt `null` zurück, wenn kein JP-Morgan-Event gefunden wurde oder das
 * T-Shirt-Custom-Field fehlt (Nichts zu tun).
 */
export async function migrateJPMorganTShirtSizes(
  eventService: EventService,
  events: DeloitteEvent[],
  onProgress?: (p: MigrationProgress) => void
): Promise<MigrationResult | null> {
  const jpEvents = events.filter(e =>
    /jp\s*morgan|jpmorgan/i.test(e.title || '') && !!e.subsiteUrl
  );
  if (jpEvents.length === 0) return null;

  // Lookup-Map per E-Mail (lowercase)
  const emailToSize: Record<string, string> = {};
  for (const row of JP_MORGAN_TSHIRT_DATA) {
    emailToSize[row.email.toLowerCase().trim()] = row.size;
  }

  const result: MigrationResult = {
    jpEventsFound: jpEvents.length,
    registrationsChecked: 0,
    updated: 0,
    skippedAlreadySet: 0,
    skippedNoMatch: 0,
    errors: 0,
  };

  for (const event of jpEvents) {
    if (!event.subsiteUrl) continue;

    if (onProgress) onProgress({ phase: 'loading-event', message: `Event "${event.title}" vorbereiten…`, current: 0, total: 0, updated: result.updated, eventTitle: event.title });

    // T-Shirt-Custom-Field finden (via Label-Match)
    const tshirtField = event.eventSpecificFields.find(f =>
      /t-?shirt|shirt.*(groe|grö)sse/i.test(f.label || '')
    );
    if (!tshirtField) {
      console.warn('[DEX][MIGRATION] Kein T-Shirt-Custom-Field in Event gefunden:', event.title);
      if (onProgress) onProgress({ phase: 'skipped', message: `Kein T-Shirt-Feld in "${event.title}" gefunden — übersprungen.`, current: 0, total: 0, updated: result.updated, eventTitle: event.title });
      continue;
    }
    const spInternalName: string = (tshirtField as unknown as { spInternalName?: string }).spInternalName || '';
    if (!spInternalName) {
      console.warn('[DEX][MIGRATION] T-Shirt-Feld hat keinen spInternalName:', event.title, tshirtField.label);
      if (onProgress) onProgress({ phase: 'skipped', message: `T-Shirt-Feld "${tshirtField.label}" hat keine SP-Spalte — übersprungen.`, current: 0, total: 0, updated: result.updated, eventTitle: event.title });
      continue;
    }
    const cfId = tshirtField.id;

    if (onProgress) onProgress({ phase: 'loading-registrations', message: `Teilnehmerliste für "${event.title}" wird geladen…`, current: 0, total: 0, updated: result.updated, eventTitle: event.title });
    let regs: SPRegistration[] = [];
    try {
      regs = await eventService.getAllRegistrations(event.subsiteUrl);
    } catch (err) {
      console.warn('[DEX][MIGRATION] getAllRegistrations failed:', event.title, err);
      if (onProgress) onProgress({ phase: 'error', message: `Teilnehmerliste "${event.title}" konnte nicht geladen werden.`, current: 0, total: 0, updated: result.updated, eventTitle: event.title });
      continue;
    }

    const total = regs.length;
    let processed = 0;
    for (const reg of regs) {
      processed += 1;
      result.registrationsChecked += 1;
      const email = (reg.ParticipantEmail || '').toLowerCase().trim();
      if (!email) { result.skippedNoMatch += 1; if (onProgress) onProgress({ phase: 'updating', message: `Teilnehmer ohne E-Mail übersprungen (${processed}/${total})`, current: processed, total, updated: result.updated, eventTitle: event.title }); continue; }
      const newSize = emailToSize[email];
      if (!newSize) { result.skippedNoMatch += 1; if (onProgress) onProgress({ phase: 'updating', message: `${email}: nicht in CSV (${processed}/${total})`, current: processed, total, updated: result.updated, eventTitle: event.title }); continue; }

      // CustomData parsen — prüfen ob schon gesetzt
      let customData: Record<string, string> = {};
      try {
        if (reg.CustomData) customData = JSON.parse(reg.CustomData);
      } catch { /* ignore invalid JSON, behandeln als leer */ }

      if (customData[cfId] && customData[cfId].trim()) {
        result.skippedAlreadySet += 1;
        if (onProgress) onProgress({ phase: 'updating', message: `${email}: bereits gefüllt, übersprungen (${processed}/${total})`, current: processed, total, updated: result.updated, eventTitle: event.title });
        continue;
      }

      // Update: SP-Spalte (spInternalName) + CustomData-JSON
      customData[cfId] = newSize;
      const body: Record<string, unknown> = {
        [spInternalName]: newSize,
        CustomData: JSON.stringify(customData),
      };
      const ok = await eventService.mergeRegistrationFields(event.subsiteUrl, reg.Id, body);
      if (ok) { result.updated += 1; } else { result.errors += 1; }
      if (onProgress) onProgress({ phase: 'updating', message: ok ? `${email} → ${newSize} (${processed}/${total})` : `${email}: Update fehlgeschlagen (${processed}/${total})`, current: processed, total, updated: result.updated, eventTitle: event.title });
    }
  }

  if (onProgress) onProgress({
    phase: 'done',
    message: `Migration abgeschlossen. ${result.updated} aktualisiert, ${result.skippedAlreadySet} übersprungen (bereits gefüllt), ${result.skippedNoMatch} ohne CSV-Match, ${result.errors} Fehler.`,
    current: result.registrationsChecked,
    total: result.registrationsChecked,
    updated: result.updated,
    eventTitle: jpEvents.map(e => e.title).join(', '),
  });

  return result;
}
