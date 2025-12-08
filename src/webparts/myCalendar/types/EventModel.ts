// export interface CalendarEvent {
//   id: string;
//   title: string;
//   startTime?: string; // ISO
//   endTime?: string;   // ISO
//   bgColor?: string;
//   type?: "booking" | "holiday" | "birthday" | "anniversary" | "other";
//   description?: string;
//   organizer?: string;
// }


export type EventType = "birthday" | "anniversary" | "booking" | "holiday";

export interface CalendarEvent {
  id: string;

  /** Event title */
  title: string;

  /** Always required */
  startTime: string | Date;

  /** Optional for non-booking events */
  endTime?: string | Date;

  /** For bookings and holidays */
  description?: string;

  /** For ConferenceBookings */
  organizer?: string;

  /** Event type */
  type: EventType;

  /** Calendar color */
  bgColor: string;
}
