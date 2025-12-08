// import * as React from "react";
// import { useEffect, useState } from "react";
// import styles from "./MyCalendar.module.scss";
// import ListService from "../services/ListService";
// import { IMyCalendarProps } from "./IMyCalendarProps";
// import { CalendarEvent } from "../types/EventModel";

// const MyCalendar: React.FC<IMyCalendarProps> = (props) => {
//   const [events, setEvents] = useState<CalendarEvent[]>([]);
//   const [selectedEvent, setSelectedEvent] = useState<CalendarEvent | null>(null);

//   const [loading, setLoading] = useState<boolean>(true);

//   const today = new Date();
//   const [viewMonth, setViewMonth] = useState<number>(today.getMonth());
//   const [viewYear, setViewYear] = useState<number>(today.getFullYear());

//   // ======================================================
//   // INIT + LOAD DATA
//   // ======================================================
//   useEffect(() => {
//     ListService.init(props.context);
//     ListService.ensureLists(); // Auto-create lists if missing
//     loadData();
//   }, []);

//   const loadData = async () => {
//     try {
//       setLoading(true);

//       const employees = await ListService.getEmployees();
//       const bookings = await ListService.getConferenceBookings();

//       const combined = ListService.mapToEvents(employees, bookings);
//       setEvents(combined);
//     } catch (err) {
//       console.error("Calendar load error:", err);
//     } finally {
//       setLoading(false);
//     }
//   };

//   // ======================================================
//   // MONTH CHANGE
//   // ======================================================
//   const changeMonth = (direction: number) => {
//     let newMonth = viewMonth + direction;
//     let newYear = viewYear;

//     if (newMonth < 0) {
//       newMonth = 11;
//       newYear--;
//     } else if (newMonth > 11) {
//       newMonth = 0;
//       newYear++;
//     }

//     setViewMonth(newMonth);
//     setViewYear(newYear);
//   };

//   const monthName = new Date(viewYear, viewMonth).toLocaleString("default", {
//     month: "long",
//   });

//   // ======================================================
//   // FORMAT DATE
//   // ======================================================
//   const formatDate = (d: string | Date) => {
//     const date = new Date(d);
//     return date.toLocaleDateString("en-US", {
//       day: "numeric",
//       month: "short",
//       year: "numeric"
//     });
//   };

//   // ======================================================
//   // RENDER CALENDAR DAYS
//   // ======================================================
//   const renderCalendarDays = () => {
//     const firstDay = new Date(viewYear, viewMonth, 1);
//     const lastDay = new Date(viewYear, viewMonth + 1, 0);

//     const daysInMonth = lastDay.getDate();
//     const startWeekDay = firstDay.getDay();

//     const dayCells: JSX.Element[] = [];

//     // Empty cells before month starts
//     for (let i = 0; i < startWeekDay; i++) {
//       dayCells.push(<div key={"empty-" + i}></div>);
//     }

//     // Days + events
//     for (let date = 1; date <= daysInMonth; date++) {
//       const fullDate = new Date(viewYear, viewMonth, date);

//       const dayEvents = events.filter((ev) => {
//         const evDate = new Date(ev.startTime);
//         return (
//           evDate.getFullYear() === fullDate.getFullYear() &&
//           evDate.getMonth() === fullDate.getMonth() &&
//           evDate.getDate() === fullDate.getDate()
//         );
//       });

//       dayCells.push(
//         <div className={styles["day-cell"]} key={date}>
//           <strong>{date}</strong>

//           {dayEvents.map((ev) => (
//             <div
//               key={ev.id}
//               className={styles["event-pill"]}
//               style={{ background: ev.bgColor }}
//               onClick={() => setSelectedEvent(ev)}
//             >
//               {ev.title}
//             </div>
//           ))}
//         </div>
//       );
//     }

//     return dayCells;
//   };

//   // ======================================================
//   // UI
//   // ======================================================
//   return (
//     <div className={styles["cal-container"]}>
//       {loading ? (
//         <p>Loading calendar...</p>
//       ) : (
//         <>
//           {/* HEADER */}
//           <div className={styles["cal-header"]}>
//             <button onClick={() => changeMonth(-1)}>‹</button>
//             <h2>
//               {monthName} {viewYear}
//             </h2>
//             <button onClick={() => changeMonth(+1)}>›</button>
//           </div>

//           {/* LEGEND */}
//           <div className={styles.legend}>
//             <span>
//               <span className={`${styles.dot} ${styles.birthday}`}></span> Birthday
//             </span>
//             <span>
//               <span className={`${styles.dot} ${styles.anniversary}`}></span> Anniversary
//             </span>
//             <span>
//               <span className={`${styles.dot} ${styles.booking}`}></span> Booking
//             </span>
//           </div>

//           {/* GRID */}
//           <div className={styles["cal-grid"]}>{renderCalendarDays()}</div>

//           {/* EVENT POPUP */}
//           {selectedEvent && (
//             <div className={styles["event-popup"]}>
//               <h3>{selectedEvent.title}</h3>

//               <p>
//                 <b>Date: </b>
//                 {formatDate(selectedEvent.startTime)}
//               </p>

//               {selectedEvent.organizer && (
//                 <p>
//                   <b>Organizer:</b> {selectedEvent.organizer}
//                 </p>
//               )}

//               {selectedEvent.description && (
//                 <p style={{ marginTop: "8px" }}>{selectedEvent.description}</p>
//               )}

//               <button onClick={() => setSelectedEvent(null)}>Close</button>
//             </div>
//           )}
//         </>
//       )}
//     </div>
//   );
// };

// export default MyCalendar;


import * as React from "react";
import { useEffect, useState } from "react";
import { Spinner, Dialog, DialogType, DefaultButton } from "@fluentui/react";
import styles from "./MyCalendar.module.scss";
import ListService from "../services/ListService";
import Calendar from "./Calendar";
import { CalendarEvent } from "../types/EventModel";
import { IMyCalendarProps } from "./IMyCalendarProps";
 
const monthName = (y: number, m: number): string =>
  new Date(y, m, 1).toLocaleString(undefined, { month: "long", year: "numeric" });
 
const normalizeBirthdayToYear = (iso: string, year: number): string => {
  const d = new Date(iso);
  const nd = new Date(year, d.getMonth(), d.getDate());
  return nd.toISOString();
};
 
const MyCalendar: React.FC<IMyCalendarProps> = (props: IMyCalendarProps): React.ReactElement => {
  const today = new Date();
  const [events, setEvents] = useState<CalendarEvent[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
 
  // view control (year/month) — start with current month
  const [viewYear, setViewYear] = useState<number>(today.getFullYear());
  const [viewMonth, setViewMonth] = useState<number>(today.getMonth());
 
  // selected event for dialog
  const [selectedEvent, setSelectedEvent] = useState<CalendarEvent | null>(null);
 
  useEffect(() => {
    ListService.init(props.context);
 
    const load = async (): Promise<void> => {
      setLoading(true);
      try {
        const employees = await ListService.getEmployees();
        const bookings = await ListService.getConferenceBookings();
 
        // If your ListService has getHolidays, you can fetch them too (optional)
        // const holidays = await ListService.getHolidays();
        // const combined = ListService.mapToEvents(employees, bookings, holidays);
 
        const combined = ListService.mapToEvents(employees, bookings);
 
        // normalize birthdays / anniversaries to current view year (so they appear in month)
        const normalized = combined.map((ev) => {
          if ((ev.type === "birthday" || ev.type === "anniversary") && ev.startTime) {
            ev.startTime = normalizeBirthdayToYear(ev.startTime, new Date().getFullYear());
            ev.endTime = undefined;
          }
          return ev;
        });
 
        setEvents(normalized as CalendarEvent[]);
      } catch (err) {
        // eslint-disable-next-line no-console
        console.error("Error loading calendar data:", err);
      } finally {
        setLoading(false);
      }
    };
 
    void load();
  }, [props.context]);
 
  const onEventClick = (ev: CalendarEvent): void => {
    setSelectedEvent(ev);
  };
 
  const onDialogDismiss = (): void => {
    setSelectedEvent(null);
  };
 
  const goPrev = (): void => {
    let m = viewMonth - 1;
    let y = viewYear;
    if (m < 0) {
      m = 11;
      y -= 1;
    }
    setViewMonth(m);
    setViewYear(y);
  };
 
  const goNext = (): void => {
    let m = viewMonth + 1;
    let y = viewYear;
    if (m > 11) {
      m = 0;
      y += 1;
    }
    setViewMonth(m);
    setViewYear(y);
  };
 
  const goToday = (): void => {
    const t = new Date();
    setViewYear(t.getFullYear());
    setViewMonth(t.getMonth());
  };
 
  if (loading) return <Spinner label="Loading calendar..." />;
 
  return (
<div className={styles.calendarWrapper}>
      {/* Header + navigation */}
<div className={styles.header}>
<div style={{ display: "flex", gap: 8, alignItems: "center" }}>
<button className={styles.navBtn} onClick={goPrev} aria-label="Previous month">◀</button>
<button className={styles.navBtn} onClick={goToday} aria-label="Today">Today</button>
<button className={styles.navBtn} onClick={goNext} aria-label="Next month">▶</button>
</div>
 
        <h2>{monthName(viewYear, viewMonth)}</h2>
 
        <div style={{ width: 120 }} /> {/* spacer to keep center title */}
</div>
 
      {/* Legend */}
<div className={styles.legend}>
<span><span className={styles.dot} style={{ background: "#FFDDDD" }}></span> Birthday</span>
<span><span className={styles.dot} style={{ background: "#D2F8D2" }}></span> Anniversary</span>
<span><span className={styles.dot} style={{ background: "#DCE7FF" }}></span> Booking</span>
</div>
 
      {/* Week days header */}
<div className={styles.weekRow}>
        {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map((d) => (
<div key={d} className={styles.weekDay}>{d}</div>
        ))}
</div>
 
      {/* Calendar grid component (grid drawing + overlay) */}
<Calendar
        year={viewYear}
        month={viewMonth}
        events={events}
        onEventClick={onEventClick}
        onMonthChange={(y, m) => { setViewYear(y); setViewMonth(m); }}
      />
 
      {/* Event dialog */}
<Dialog
        hidden={!selectedEvent}
        onDismiss={onDialogDismiss}
        dialogContentProps={{
          type: DialogType.normal,
          title: selectedEvent?.title ?? "Event details"
        }}
>
        {selectedEvent && (
<div>
<p><strong>When:</strong> {selectedEvent.startTime ? new Date(selectedEvent.startTime).toLocaleString() : "N/A"}{selectedEvent.endTime ? ` - ${new Date(selectedEvent.endTime).toLocaleString()}` : ""}</p>
            {selectedEvent.organizer && <p><strong>Organizer:</strong> {selectedEvent.organizer}</p>}
            {selectedEvent.description && <p>{selectedEvent.description}</p>}
<div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
<DefaultButton text="Close" onClick={onDialogDismiss} />
</div>
</div>
        )}
</Dialog>
</div>
  );
};
 
export default MyCalendar;