// DEX Event Platform - Type Definitions
// Author: Eike Brenneisen

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
  locationAudience: string[];
  startDate: string;
  endDate: string;
  registrationDeadline: string;
  description: string;
  maxParticipants: number;
  currentParticipants: number;
  waitlistCount: number;
  imageUrl?: string;
  eventSpecificFields: EventSpecificField[];
}

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

export type UserRole = 'SuperAdmin' | 'EventAdmin' | 'User';

export interface User {
  id: string;
  firstName: string;
  surname: string;
  email: string;
  isAdmin: boolean;
  role: UserRole;
  location: string;
}

export interface RoleAssignment {
  id: number;
  userEmail: string;
  userName: string;
  role: UserRole;
  location: string;
  assignedBy: string;
  assignedDate: string;
}
