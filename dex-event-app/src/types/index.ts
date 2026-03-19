// DEX Event Platform - Type Definitions
// Author: Eike Brenneisen
// Last updated: March 2026

export type EventType = 'B2Run' | 'JPMorgan' | 'Other';
export type EventStatus = 'Under Construction' | 'Active' | 'Completed' | 'Cancelled';
export type RegistrationStatus = 'Registered' | 'Waitlist' | 'Checked-In' | 'Cancelled';
export type Salutation = 'Herr' | 'Frau' | 'Divers';

export interface DeloitteEvent {
  id: string;
  title: string;
  type: EventType;
  status: EventStatus;
  organizers: string[];
  location: string;
  locationAudience: string[]; // welche Standorte das Event sehen koennen
  startDate: string;
  endDate: string;
  registrationDeadline: string;
  description: string;
  maxParticipants: number;
  currentParticipants: number;
  waitlistCount: number;
  imageUrl?: string; // TODO: Bildupload implementieren
  eventSpecificFields: EventSpecificField[];
}

// Dynamische Felder die pro Event konfiguriert werden koennen
// z.B. T-Shirt Groesse beim B2Run oder Essensauswahl bei Meetings
export interface EventSpecificField {
  id: string;
  label: string;
  type: 'text' | 'select' | 'number' | 'checkbox';
  required: boolean;
  options?: string[];
  helpText?: string;
}

export interface Registration {
  id: string;
  eventId: string;
  eventTitle: string;
  salutation: Salutation;
  firstName: string;
  surname: string;
  email: string;
  status: RegistrationStatus;
  registrationDate: string;
  cancellationDate?: string;
  waitlistPosition?: number;
  eventSpecificData: Record<string, string>;
}

export interface User {
  id: string;
  firstName: string;
  surname: string;
  email: string;
  isAdmin: boolean;
  location: string;
}
