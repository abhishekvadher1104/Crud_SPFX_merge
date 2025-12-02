import * as React from "react";
import type { ICrudProps } from "./ICrudProps";
import { getSP } from "../pnpConfig";
import {
  Checkbox,
  DetailsList,
  DetailsListLayoutMode,
  Dropdown,
  IColumn,
  IconButton,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  TextField,
} from "@fluentui/react";

interface IEmployee {
  ID?: number;
  Title?: string;
  Departments?: string;
  Experience?: number;
  IsActive?: boolean;
}

interface ICrudoperationsState {
  employees: IEmployee[];
  successMessage: string;
  newEmployee?: IEmployee;
  departmentChoices: string[];
  editEmployeeId?: number;
  editedEmployee?: IEmployee;
}

export default class Crudoperations extends React.Component<
  ICrudProps,
  ICrudoperationsState
> {
  private sp = getSP();

  constructor(props: ICrudProps) {
    super(props);
    this.state = {
      employees: [],
      successMessage: "",
      departmentChoices: [],
      editEmployeeId: undefined,
      editedEmployee: undefined,
    };
  }

  // ======================================
  // 🔥 AUTO-LIST CREATION (Correct version)
  // ======================================
private ensureListExists = async (): Promise<void> => {
  const listTitle = "EmployeesSpfx";
  const listRef = this.sp.web.lists.getByTitle(listTitle);

  let listExists = true;

  // ---------- CHECK IF LIST EXISTS ----------
  try {
    await listRef(); // OK → list exists
    console.log("✔ List already exists:", listTitle);
  } catch (e) {
    listExists = false;
    console.log("❌ List does NOT exist → creating:", listTitle);
    console.log(listExists);
    

    // ---------- CREATE LIST ----------
    await this.sp.web.lists.add(
      listTitle,
      "Auto-created list for SPFx CRUD",
      100 // Generic List
    );

    // Give SP time to provision fields
    await new Promise((resolve) => setTimeout(resolve, 500));
  }

  // ---------- ENSURE ALL FIELDS EXIST ----------
  const fields = listRef.fields;

  // Reusable helper function
  const ensureField = async (
    internalName: string,
    creator: () => Promise<any>
  ) => {
    try {
      await fields.getByInternalNameOrTitle(internalName)();
      console.log(`✔ Field '${internalName}' already exists`);
    } catch (err) {
      console.log(`➕ Creating missing field '${internalName}'`);
      await creator();
    }
  };

  // Title already exists → DO NOT recreate Title

  await ensureField("Departments", () =>
    fields.addChoice("Departments", {
      Choices: ["IT", "HR", "Finance", "Admin", "Sales"],
    })
  );

  await ensureField("Experience", () => fields.addNumber("Experience"));

  await ensureField("IsActive", () => fields.addBoolean("IsActive"));

  console.log("🎉 All fields ensured for list:", listTitle);
};


  // ======================================
  // INITIAL LOAD
  // ======================================
  public async componentDidMount(): Promise<void> {
    await this.ensureListExists(); // 🔥 Auto-create list + columns
    await this.loadDepartmentChoices();
    await this.loadEmployees();
  }

  // ======================================
  // LOAD EMPLOYEES
  // ======================================
  private loadEmployees = async (): Promise<void> => {
    try {
      const items = await this.sp.web.lists
        .getByTitle("EmployeesSpfx")
        .items.select("ID", "Title", "Departments", "Experience", "IsActive")();

      this.setState({
        employees: items,
      });
    } catch (error) {
      console.error("Error loading employees:", error);
    }
  };

  // ======================================
  // LOAD CHOICES
  // ======================================
  private loadDepartmentChoices = async (): Promise<void> => {
    try {
      const field = await this.sp.web.lists
        .getByTitle("EmployeesSpfx")
        .fields.getByInternalNameOrTitle("Departments")();

      if (field?.Choices) {
        this.setState({ departmentChoices: field.Choices });
      }
    } catch (error) {
      console.error("Error loading choices:", error);
    }
  };

  // ======================================
  // ADD NEW EMPLOYEE
  // ======================================
  private handleSaveNewEmployee = async (): Promise<void> => {
    const { newEmployee } = this.state;

    if (!newEmployee?.Title || !newEmployee.Departments) {
      alert("Please fill name and department");
      return;
    }

    await this.sp.web.lists.getByTitle("EmployeesSpfx").items.add(newEmployee);
    this.setState({ newEmployee: undefined });

    await this.loadEmployees();
  };

  private handleAddEmployee = () => {
    this.setState({
      newEmployee: {
        Title: "",
        Departments: "",
        Experience: 0,
        IsActive: true,
      },
      editedEmployee: undefined,
      editEmployeeId: undefined,
    });
  };

  // ======================================
  // EDIT EMPLOYEE
  // ======================================
  private handleEditEmployee = (emp: IEmployee) => {
    this.setState({
      editEmployeeId: emp.ID,
      editedEmployee: { ...emp },
      newEmployee: undefined,
    });
  };

  private handleSaveEditedEmployee = async () => {
    const { editedEmployee } = this.state;
    if (!editedEmployee?.ID) return;

    await this.sp.web.lists
      .getByTitle("EmployeesSpfx")
      .items.getById(editedEmployee.ID)
      .update(editedEmployee);

    this.setState({ editEmployeeId: undefined, editedEmployee: undefined });
    await this.loadEmployees();
  };

  // ======================================
  // DELETE EMPLOYEE
  // ======================================
  private handleDeleteEmployee = async (id?: number) => {
    if (!id) return;
    if (!confirm("Delete this employee?")) return;

    await this.sp.web.lists
      .getByTitle("EmployeesSpfx")
      .items.getById(id)
      .delete();
    await this.loadEmployees();
  };

  // ======================================
  // DETAILS LIST COLUMNS
  // ======================================
  private getColumns(): IColumn[] {
    const { newEmployee, editedEmployee, editEmployeeId, departmentChoices } =
      this.state;

    const deptOptions: IDropdownOption[] = departmentChoices.map((c) => ({
      key: c,
      text: c,
    }));

    return [
      {
        key: "name",
        name: "Name",
        fieldName: "Title",
        minWidth: 120,
        onRender: (item: IEmployee) => {
          if (newEmployee === item)
            return (
              <TextField
                value={item.Title}
                onChange={(_, val) =>
                  this.setState({
                    newEmployee: { ...item, Title: val || "" },
                  })
                }
              />
            );

          if (editEmployeeId === item.ID)
            return (
              <TextField
                value={editedEmployee?.Title}
                onChange={(_, val) =>
                  this.setState({
                    editedEmployee: { ...editedEmployee, Title: val || "" },
                  })
                }
              />
            );

          return item.Title;
        },
      },

      {
        key: "dept",
        name: "Department",
        fieldName: "Departments",
        minWidth: 120,
        onRender: (item: IEmployee) => {
          if (newEmployee === item)
            return (
              <Dropdown
                options={deptOptions}
                onChange={(_, opt) =>
                  this.setState({
                    newEmployee: { ...item, Departments: opt?.key as string },
                  })
                }
              />
            );

          if (editEmployeeId === item.ID)
            return (
              <Dropdown
                options={deptOptions}
                selectedKey={editedEmployee?.Departments}
                onChange={(_, opt) =>
                  this.setState({
                    editedEmployee: {
                      ...editedEmployee,
                      Departments: opt?.key as string,
                    },
                  })
                }
              />
            );

          return item.Departments;
        },
      },

      {
        key: "exp",
        name: "Experience",
        fieldName: "Experience",
        minWidth: 100,
        onRender: (item: IEmployee) => {
          if (newEmployee === item)
            return (
              <TextField
                type="number"
                value={(item.Experience ?? 0).toString()}
                onChange={(_, val) =>
                  this.setState({
                    newEmployee: {
                      ...item,
                      Experience: Number(val) || 0,
                    },
                  })
                }
              />
            );

          if (editEmployeeId === item.ID)
            return (
              <TextField
                type="number"
                value={(editedEmployee?.Experience ?? 0).toString()}
                onChange={(_, val) =>
                  this.setState({
                    editedEmployee: {
                      ...editedEmployee,
                      Experience: Number(val) || 0,
                    },
                  })
                }
              />
            );

          return item.Experience;
        },
      },

      {
        key: "active",
        name: "Active",
        fieldName: "IsActive",
        minWidth: 80,
        onRender: (item: IEmployee) => {
          if (newEmployee === item)
            return (
              <Checkbox
                checked={item.IsActive}
                onChange={(_, chk) =>
                  this.setState({
                    newEmployee: { ...item, IsActive: chk || false },
                  })
                }
              />
            );

          if (editEmployeeId === item.ID)
            return (
              <Checkbox
                checked={editedEmployee?.IsActive}
                onChange={(_, chk) =>
                  this.setState({
                    editedEmployee: {
                      ...editedEmployee,
                      IsActive: chk || false,
                    },
                  })
                }
              />
            );

          return <Checkbox checked={item.IsActive} disabled />;
        },
      },

      {
        key: "actions",
        name: "Actions",
        minWidth: 120,
        onRender: (item: IEmployee) => {
          if (newEmployee === item)
            return (
              <>
                <PrimaryButton
                  text="Save"
                  onClick={this.handleSaveNewEmployee}
                />
                <PrimaryButton
                  text="Cancel"
                  onClick={() => this.setState({ newEmployee: undefined })}
                  style={{ marginLeft: 6 }}
                />
              </>
            );

          if (editEmployeeId === item.ID)
            return (
              <>
                <PrimaryButton
                  text="Update"
                  onClick={this.handleSaveEditedEmployee}
                />
                <PrimaryButton
                  text="Cancel"
                  onClick={() =>
                    this.setState({
                      editEmployeeId: undefined,
                      editedEmployee: undefined,
                    })
                  }
                  style={{ marginLeft: 6 }}
                />
              </>
            );

          return (
            <>
              <IconButton
                iconProps={{ iconName: "Edit" }}
                onClick={() => this.handleEditEmployee(item)}
              />
              <IconButton
                iconProps={{ iconName: "Delete" }}
                onClick={() => this.handleDeleteEmployee(item.ID)}
              />
            </>
          );
        },
      },
    ];
  }

  // ======================================
  // RENDER
  // ======================================
  public render() {
    const { employees, newEmployee, successMessage } = this.state;
    const items = newEmployee ? [newEmployee, ...employees] : employees;

    return (
      <>
        {successMessage && (
          <MessageBar messageBarType={MessageBarType.success}>
            {successMessage}
          </MessageBar>
        )}

        <PrimaryButton
          text="Add Employee"
          onClick={this.handleAddEmployee}
          style={{ marginBottom: 10 }}
        />

        <DetailsList
          items={items}
          columns={this.getColumns()}
          layoutMode={DetailsListLayoutMode.fixedColumns}
        />
      </>
    );
  }
}
