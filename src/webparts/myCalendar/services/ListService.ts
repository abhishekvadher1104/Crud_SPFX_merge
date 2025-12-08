// import { WebPartContext } from "@microsoft/sp-webpart-base";

// // PnPjs v3 imports
// import { spfi, SPFI } from "@pnp/sp";
// import { SPFx } from "@pnp/sp/presets/all";

// // Import required PnP modules
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";

// export default class ListService {
//   private static _sp: SPFI;

//   public static init(context: WebPartContext): void {
//     this._sp = spfi().using(SPFx(context));
//   }

//   // ===========================================================
//   // GET EMPLOYEES LIST
//   // ===========================================================
//   public static async getEmployees(): Promise<any[]> {
//     return await this._sp.web.lists
//       .getByTitle("Employees")
//       .items.select("Id", "Title", "Email", "DOB", "JoiningDate")();
//   }

//   // ===========================================================
//   // GET CONFERENCE BOOKINGS LIST
//   // ===========================================================
//   public static async getConferenceBookings(): Promise<any[]> {
//     return await this._sp.web.lists
//       .getByTitle("Conference Bookings")
//       .items.select("Id", "Title", "Start", "End", "Description")();
//   }

//   // ===========================================================
//   // ✅ GET HOLIDAYS LIST  (NEW)
//   // ===========================================================
//   public static async getHolidays(): Promise<any[]> {
//     return await this._sp.web.lists
//       .getByTitle("Holidays")
//       .items.select("Id", "Title", "HolidayDate", "Description")();
//   }

//   // ===========================================================
//   // UPDATED mapToEvents — now supports 3rd parameter: holidays
//   // ===========================================================
//   public static mapToEvents(
//     employees: any[],
//     bookings: any[],
//     holidays: any[]
//   ) {
//     const events: any[] = [];

//     // 🎂 Birthdays
//     employees.forEach((emp) => {
//       if (emp.DOB) {
//         events.push({
//           id: "BD_" + emp.Id,
//           title: "🎂 Birthday: " + emp.Title,
//           startTime: emp.DOB,
//           type: "birthday",
//           bgColor: "#FFDDDD",
//         });
//       }
//     });

//     // 🎉 Work Anniversaries
//     employees.forEach((emp) => {
//       if (emp.JoiningDate) {
//         events.push({
//           id: "AN_" + emp.Id,
//           title: "🎉 Work Anniversary: " + emp.Title,
//           startTime: emp.JoiningDate,
//           type: "anniversary",
//           bgColor: "#D2F8D2",
//         });
//       }
//     });

//     // 📅 Conference Bookings
//     bookings.forEach((b) => {
//       events.push({
//         id: "BOOK_" + b.Id,
//         title: b.Title,
//         startTime: b.Start,
//         endTime: b.End,
//         type: "booking",
//         bgColor: "#DCE7FF",
//       });
//     });

//     // 🎊 Holidays  (NEW)
//     holidays.forEach((h) => {
//       if (h.HolidayDate) {
//         events.push({
//           id: "HOL_" + h.Id,
//           title: "🎊 Holiday: " + h.Title,
//           startTime: h.HolidayDate,
//           type: "holiday",
//           bgColor: "#FFF3CD",
//         });
//       }
//     });

//     return events;
//   }
// }



import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { FieldTypes } from "@pnp/sp/fields/types";

export default class ListService {
  private static _sp: SPFI;

  public static init(context: WebPartContext): void {
    this._sp = spfi().using(SPFx(context));
  }

  // =====================================================
  // AUTO CREATE LISTS
  // =====================================================
  public static async ensureLists(): Promise<void> {
    await this.ensureEmployeesList();
    await this.ensureConferenceBookingsList();
    console.log("✔ List creation check completed.");
  }

  // =====================================================
  // EMPLOYEES
  // =====================================================
  private static async ensureEmployeesList(): Promise<void> {
    const listTitle = "Employees";

    const existing = await this._sp.web.lists.filter(`Title eq '${listTitle}'`)();
    if (existing.length === 0) {
      await this._sp.web.lists.add(listTitle, "Employee master data", 100);
      console.log("✔ Employees list created");
    }

    const list = this._sp.web.lists.getByTitle(listTitle);

    const fields = [
      { name: "EmployeeID", type: FieldTypes.Text },
      { name: "FirstName", type: FieldTypes.Text },
      { name: "LastName", type: FieldTypes.Text },
      { name: "Email", type: FieldTypes.Text },
      { name: "Phone", type: FieldTypes.Text },
      { name: "Department", type: FieldTypes.Text },
      { name: "JobTitle", type: FieldTypes.Text },
      { name: "Location", type: FieldTypes.Text },
      { name: "Status", type: FieldTypes.Text },
      { name: "DOB", type: FieldTypes.DateTime },
      { name: "JoiningDate", type: FieldTypes.DateTime }
    ];

    for (const f of fields) {
      try {
        await list.fields.getByInternalNameOrTitle(f.name)();
      } catch {
        console.log("⚡ Creating field:", f.name);
        await list.fields.add(f.name, f.type);
      }
    }
  }

  // =====================================================
  // CONFERENCE BOOKINGS
  // =====================================================
  private static async ensureConferenceBookingsList(): Promise<void> {
    const listTitle = "ConferenceBookings";

    const existing = await this._sp.web.lists.filter(`Title eq '${listTitle}'`)();
    if (existing.length === 0) {
      await this._sp.web.lists.add(listTitle, "Meeting room booking", 100);
      console.log("✔ ConferenceBookings list created");
    }

    const list = this._sp.web.lists.getByTitle(listTitle);

    const fields = [
      { name: "Start", type: FieldTypes.DateTime },
      { name: "End", type: FieldTypes.DateTime },
      { name: "Description", type: FieldTypes.Text }
    ];

    for (const f of fields) {
      try {
        await list.fields.getByInternalNameOrTitle(f.name)();
      } catch {
        console.log("⚡ Creating field:", f.name);
        await list.fields.add(f.name, f.type);
      }
    }

    // Organizer - Person field
    try {
      await list.fields.getByInternalNameOrTitle("Organizer")();
    } catch {
      console.log("⚡ Creating Organizer person field");
      await list.fields.addUser("Organizer");
    }
  }

  // =====================================================
  // GET EMPLOYEES
  // =====================================================
  public static async getEmployees(): Promise<any[]> {
    return await this._sp.web.lists
      .getByTitle("Employees")
      .items.select("Id", "Title", "Email", "DOB", "JoiningDate")();
  }

  // =====================================================
  // GET BOOKINGS
  // =====================================================
  public static async getConferenceBookings(): Promise<any[]> {
    return await this._sp.web.lists
      .getByTitle("ConferenceBookings")
      .items.select(
        "Id",
        "Title",
        "Start",
        "End",
        "Description",
        "Organizer/Id",
        "Organizer/Title",
        "Organizer/Email"
      )
      .expand("Organizer")();
  }

  // =====================================================
  // MAP EVENTS
  // =====================================================
  public static mapToEvents(employees: any[], bookings: any[]) {
    const events: any[] = [];

    // Birthdays
    employees.forEach((emp) => {
      if (emp.DOB) {
        events.push({
          id: "BD_" + emp.Id,
          title: "🎂 " + emp.Title,
          startTime: emp.DOB,
          type: "birthday",
          bgColor: "#FFD1D1",
        });
      }
    });

    // Anniversaries
    employees.forEach((emp) => {
      if (emp.JoiningDate) {
        events.push({
          id: "AN_" + emp.Id,
          title: "🎉 Anniversary: " + emp.Title,
          startTime: emp.JoiningDate,
          type: "anniversary",
          bgColor: "#D8FFD8",
        });
      }
    });

    // Conference bookings
    bookings.forEach((b) => {
      events.push({
        id: "BK_" + b.Id,
        title: b.Title,
        startTime: b.Start,
        endTime: b.End,
        organizer: b.Organizer?.Title,
        type: "booking",
        bgColor: "#D8E5FF",
      });
    });

    return events;
  }
}

