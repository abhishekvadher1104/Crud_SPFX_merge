import * as React from "react";
import { useEffect, useMemo, useRef, useState } from "react";
import styles from "./Calendar.module.scss";
import { CalendarEvent } from "../types/EventModel";

export interface CalendarProps {
  year?: number;
  month?: number;
  events: CalendarEvent[];
  onEventClick?: (ev: CalendarEvent) => void;
  onMonthChange?: (year: number, month: number) => void;
}

type PlacedEvent = CalendarEvent & {
  startDate: Date;
  endDate: Date;
  startIndex: number;
  durationDays: number;
  weekRow: number;
};

const MS_PER_DAY = 1000 * 60 * 60 * 24;

function dateOnly(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function daysBetween(a: Date, b: Date): number {
  const da = dateOnly(a).getTime();
  const db = dateOnly(b).getTime();
  return Math.round((db - da) / MS_PER_DAY);
}

export default function Calendar({
  year,
  month,
  events,
  onEventClick,
  onMonthChange,
}: CalendarProps): React.ReactElement {
  const today = new Date();
  const viewYear = year ?? today.getFullYear();
  const viewMonth = month ?? today.getMonth();

  const firstOfMonth = useMemo(() => new Date(viewYear, viewMonth, 1), [viewYear, viewMonth]);
  const daysInMonth = useMemo(() => new Date(viewYear, viewMonth + 1, 0).getDate(), [viewYear, viewMonth]);

  const startWeekday = firstOfMonth.getDay();
  const totalCells = Math.ceil((startWeekday + daysInMonth) / 7) * 7;

  const gridDates = useMemo(() => {
    const arr: Date[] = [];
    const gridStart = new Date(viewYear, viewMonth, 1 - startWeekday);
    for (let i = 0; i < totalCells; i++) {
      const d = new Date(gridStart.getFullYear(), gridStart.getMonth(), gridStart.getDate() + i);
      arr.push(d);
    }
    return arr;
  }, [viewYear, viewMonth, startWeekday, totalCells]);

  const gridRef = useRef<HTMLDivElement | null>(null);
  const [colWidth, setColWidth] = useState<number>(0);

  useEffect(() => {
    function measure(): void {
      if (!gridRef.current) return;
      const w = gridRef.current.clientWidth;
      setColWidth(Math.floor(w / 7));
    }
    measure();
    window.addEventListener("resize", measure);
    return () => window.removeEventListener("resize", measure);
  }, []);

  const placedEvents = useMemo(() => {
    const res: PlacedEvent[] = [];
    const monthStart = new Date(viewYear, viewMonth, 1);
    const monthEnd = new Date(viewYear, viewMonth, daysInMonth, 23, 59, 59);

    const clamp = (d: Date) => {
      if (d < monthStart) return monthStart;
      if (d > monthEnd) return monthEnd;
      return d;
    };

    events.forEach(e => {
      if (!e.startTime) return;

      const s = new Date(e.startTime);
      const en = e.endTime ? new Date(e.endTime) : new Date(e.startTime);

      if (en < monthStart || s > monthEnd) return;

      const start = clamp(s);
      const end = clamp(en);

      const gridStart = gridDates[0];
      const startIndex = daysBetween(gridStart, start);
      const durationDays = Math.max(1, daysBetween(start, end) + 1);
      const weekRow = Math.floor(startIndex / 7);

      res.push({
        ...e,
        startDate: start,
        endDate: end,
        startIndex,
        durationDays,
        weekRow,
      });
    });

    return res;
  }, [events, gridDates, viewMonth, viewYear, daysInMonth]);

  // navigation helpers - emit month change via callback
  const goPrev = (): void => {
    let m = viewMonth - 1;
    let y = viewYear;
    if (m < 0) {
      m = 11;
      y -= 1;
    }
    onMonthChange?.(y, m);
  };

  const goNext = (): void => {
    let m = viewMonth + 1;
    let y = viewYear;
    if (m > 11) {
      m = 0;
      y += 1;
    }
    onMonthChange?.(y, m);
  };

  const goToday = (): void => {
    const t = new Date();
    onMonthChange?.(t.getFullYear(), t.getMonth());
  };

  return (
    <div className={styles.calendarWrap}>
      <div className={styles.header}>
        <div>
          <div className={styles.monthTitle}>
            {firstOfMonth.toLocaleString(undefined, { month: "long", year: "numeric" })}
          </div>
        </div>

        <div className={styles.controls}>
          <button className={styles.button} onClick={goPrev}>◀</button>
          <button className={styles.button} onClick={goToday}>Today</button>
          <button className={styles.button} onClick={goNext}>▶</button>
        </div>
      </div>

      <div className={styles.weekDays}>
        {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map(d => (
          <div key={d} className={styles.weekDayCell}>{d}</div>
        ))}
      </div>

      <div className={styles.gridWrap} ref={gridRef}>
        <div className={styles.grid}>
          {gridDates.map((d, idx) => {
            const isCurrentMonth = d.getMonth() === viewMonth;
            return (
              <div key={idx} className={`${styles.cell} ${isCurrentMonth ? "" : styles.outside}`}>
                <div className={styles.dayNumber}>{d.getDate()}</div>
              </div>
            );
          })}
        </div>

        <div className={styles.overlay}>
          {placedEvents.map((ev, i) => {
            const left = ev.startIndex * colWidth;
            const width = ev.durationDays * colWidth - 6;
            const top = ev.weekRow * 36 + 8;

            return (
              <div
                key={ev.id ?? i}
                className={styles.eventBlock}
                style={{
                  left: left + 4,
                  top,
                  width: Math.max(80, width),
                  background: ev.bgColor ?? (ev.type === "booking" ? "#9ec8ff" : ev.type === "holiday" ? "#ffd57f" : "#f3a6a6")
                }}
                onClick={() => onEventClick?.(ev)}
                title={`${ev.title}${ev.startTime ? " — " + new Date(ev.startTime).toLocaleString() : ""}`}
              >
                <div className={styles.eventTitle}>{ev.title}</div>
                {ev.startTime && (
                  <div className={styles.eventTime}>
                    {new Date(ev.startTime).toLocaleString(undefined, { month: "short", day: "numeric", hour: "numeric", minute: "numeric" })}{ev.endTime ? ` - ${new Date(ev.endTime).toLocaleString(undefined, { hour: "numeric", minute: "numeric" })}` : ""}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>

      <div className={styles.legend}>
        <div className={styles.legendItem}><span className={styles.legendBox} style={{ background: "#9ec8ff" }}></span> Booking</div>
        <div className={styles.legendItem}><span className={styles.legendBox} style={{ background: "#ffd57f" }}></span> Holiday</div>
        <div className={styles.legendItem}><span className={styles.legendBox} style={{ background: "#f3a6a6" }}></span> Birthday / Anniversary</div>
      </div>
    </div>
  );
}
