import os
import tkinter as tk
from tkinter import messagebox

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment


# ------------- CONFIG -------------

DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
ROLES = ["admin", "shift", "area"]
GENDERS = ["M", "F"]

EXCEL_FILENAME = "db.xlsx"
SHEET_NAME = "Managers"


# ------------- EXCEL HELPERS -------------

def get_excel_path():
    """Always keep the Excel file in the same folder as the script/app."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, EXCEL_FILENAME)


def create_excel_if_missing():
    """Create a new nicely formatted Excel file if it does not exist."""
    path = get_excel_path()
    if os.path.exists(path):
        return

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    # Header row
    headers = ["ID", "Name", "Role", "Gender"]
    for day in DAYS:
        headers.append(f"{day}_start")
        headers.append(f"{day}_end")

    ws.append(headers)

    # Make header look nice
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        # Adjust column width a bit
        if header in ("ID", "Role", "Gender"):
            ws.column_dimensions[cell.column_letter].width = 10
        elif "start" in header or "end" in header:
            ws.column_dimensions[cell.column_letter].width = 12
        else:  # Name
            ws.column_dimensions[cell.column_letter].width = 25

    wb.save(path)


def load_all_managers():
    """
    Load all managers from Excel.
    Returns a list of dicts: [{row_index, id, name, role, gender, availability}, ...]
    availability is a dict: {"Mon": (start, end), ...} where start/end are ints or None.
    """
    path = get_excel_path()
    wb = load_workbook(path)
    ws = wb[SHEET_NAME]

    managers = []

    for row in range(2, ws.max_row + 1):
        id_val = ws.cell(row=row, column=1).value
        name = ws.cell(row=row, column=2).value or ""
        role = ws.cell(row=row, column=3).value or ""
        gender = ws.cell(row=row, column=4).value or ""

        availability = {}
        col = 5
        for day in DAYS:
            start_val = ws.cell(row=row, column=col).value
            end_val = ws.cell(row=row, column=col + 1).value
            availability[day] = (start_val, end_val)
            col += 2

        managers.append({
            "row_index": row,
            "id": id_val,
            "name": name,
            "role": role,
            "gender": gender,
            "availability": availability
        })

    wb.close()
    return managers


def write_manager_to_excel(row_index, manager_data):
    """Write/overwrite a single manager row in Excel."""
    path = get_excel_path()
    wb = load_workbook(path)
    ws = wb[SHEET_NAME]

    # If row_index is None, append a fresh row
    if row_index is None:
        row_index = ws.max_row + 1

    ws.cell(row=row_index, column=1, value=manager_data["id"])
    ws.cell(row=row_index, column=2, value=manager_data["name"])
    ws.cell(row=row_index, column=3, value=manager_data["role"])
    ws.cell(row=row_index, column=4, value=manager_data["gender"])

    col = 5
    for day in DAYS:
        start_val, end_val = manager_data["availability"].get(day, (None, None))
        ws.cell(row=row_index, column=col, value=start_val)
        ws.cell(row=row_index, column=col + 1, value=end_val)
        col += 2

    wb.save(path)
    wb.close()


def delete_row_from_excel(row_index):
    """Delete a row from Excel."""
    path = get_excel_path()
    wb = load_workbook(path)
    ws = wb[SHEET_NAME]
    ws.delete_rows(row_index, 1)
    wb.save(path)
    wb.close()


def get_next_id(managers):
    """Generate a simple incremental numeric ID."""
    max_id = 0
    for m in managers:
        try:
            mid = int(m["id"])
            if mid > max_id:
                max_id = mid
        except (TypeError, ValueError):
            continue
    return max_id + 1


# ------------- TKINTER APP -------------

class ManagersApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Managers (Excel database)")
        self.root.geometry("900x480")

        # Ensure Excel file exists
        create_excel_if_missing()

        # in-memory list of managers (dicts)
        self.managers = []

        # ---------- LEFT: list of managers ----------
        left_frame = tk.Frame(root)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        tk.Label(left_frame, text="Managers").pack()

        list_frame = tk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(list_frame, width=30)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.listbox.yview)

        self.listbox.bind("<<ListboxSelect>>", self.on_select)

        # ---------- RIGHT: form ----------
        right_frame = tk.Frame(root)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Name
        tk.Label(right_frame, text="Name (first + last):").grid(row=0, column=0, sticky="w")
        self.entry_name = tk.Entry(right_frame)
        self.entry_name.grid(row=0, column=1, columnspan=3, sticky="we", padx=5, pady=2)

        # Role
        tk.Label(right_frame, text="Role:").grid(row=1, column=0, sticky="w")
        self.role_var = tk.StringVar(value=ROLES[0])
        self.role_menu = tk.OptionMenu(right_frame, self.role_var, *ROLES)
        self.role_menu.grid(row=1, column=1, sticky="w", padx=5, pady=2)

        # Gender
        tk.Label(right_frame, text="Gender:").grid(row=1, column=2, sticky="w")
        self.gender_var = tk.StringVar(value=GENDERS[0])
        self.gender_menu = tk.OptionMenu(right_frame, self.gender_var, *GENDERS)
        self.gender_menu.grid(row=1, column=3, sticky="w", padx=5, pady=2)

        # Availability explanation
        tk.Label(
            right_frame,
            text="Availability (hours 0–23, empty = OFF)\nExample: Tue 13–23 → after 13:00"
        ).grid(row=2, column=0, columnspan=4, sticky="w", pady=(10, 2))

        # Spinboxes for each day
        self.avail_start_vars = {}
        self.avail_end_vars = {}

        for i, day in enumerate(DAYS):
            row = 3 + i
            tk.Label(right_frame, text=f"{day}:").grid(row=row, column=0, sticky="w")

            start_var = tk.StringVar()
            end_var = tk.StringVar()
            self.avail_start_vars[day] = start_var
            self.avail_end_vars[day] = end_var

            start_spin = tk.Spinbox(
                right_frame, from_=0, to=23, textvariable=start_var, width=5
            )
            end_spin = tk.Spinbox(
                right_frame, from_=0, to=23, textvariable=end_var, width=5
            )

            start_spin.grid(row=row, column=1, sticky="w", padx=5, pady=1)
            tk.Label(right_frame, text="to").grid(row=row, column=2, sticky="w")
            end_spin.grid(row=row, column=3, sticky="w", padx=5, pady=1)

        # Buttons
        btn_frame = tk.Frame(right_frame)
        btn_frame.grid(row=3 + len(DAYS), column=0, columnspan=4, pady=15)

        tk.Button(btn_frame, text="Add", width=10, command=self.add_manager).grid(row=0, column=0, padx=5)
        tk.Button(btn_frame, text="Update", width=10, command=self.update_manager).grid(row=0, column=1, padx=5)
        tk.Button(btn_frame, text="Delete", width=10, command=self.delete_manager).grid(row=0, column=2, padx=5)

        right_frame.columnconfigure(1, weight=1)
        right_frame.columnconfigure(3, weight=1)

        # Load data
        self.reload_from_excel()

    # ---------- Data <-> UI helpers ----------

    def reload_from_excel(self):
        self.managers = load_all_managers()
        self.refresh_listbox()

    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for m in self.managers:
            display = f"{m['name']} ({m['role']})"
            self.listbox.insert(tk.END, display)

    def get_selected_index(self):
        sel = self.listbox.curselection()
        if not sel:
            return None
        return sel[0]

    def clear_form(self):
        self.entry_name.delete(0, tk.END)
        self.role_var.set(ROLES[0])
        self.gender_var.set(GENDERS[0])
        for day in DAYS:
            self.avail_start_vars[day].set("")
            self.avail_end_vars[day].set("")

    def read_form(self):
        name = self.entry_name.get().strip()
        if not name:
            messagebox.showwarning("Input error", "Name cannot be empty.")
            return None

        role = self.role_var.get()
        gender = self.gender_var.get()

        availability = {}
        for day in DAYS:
            start_str = self.avail_start_vars[day].get().strip()
            end_str = self.avail_end_vars[day].get().strip()

            if start_str == "" and end_str == "":
                # OFF that day
                availability[day] = (None, None)
                continue

            # Basic validation: must be integers 0–23 if filled
            try:
                start_val = int(start_str) if start_str != "" else None
                end_val = int(end_str) if end_str != "" else None

                if start_val is not None and not (0 <= start_val <= 23):
                    raise ValueError
                if end_val is not None and not (0 <= end_val <= 23):
                    raise ValueError
                if start_val is not None and end_val is not None and start_val > end_val:
                    messagebox.showwarning(
                        "Input error",
                        f"For {day}, start hour cannot be greater than end hour."
                    )
                    return None

                availability[day] = (start_val, end_val)
            except ValueError:
                messagebox.showwarning(
                    "Input error",
                    f"For {day}, hours must be integers 0–23 or empty."
                )
                return None

        return {
            "name": name,
            "role": role,
            "gender": gender,
            "availability": availability
        }

    # ---------- Button actions ----------

    def add_manager(self):
        data = self.read_form()
        if data is None:
            return

        # Assign ID
        data["id"] = get_next_id(self.managers)

        # Append new row
        write_manager_to_excel(None, data)
        self.reload_from_excel()
        self.clear_form()

    def update_manager(self):
        idx = self.get_selected_index()
        if idx is None:
            messagebox.showwarning("Selection error", "Please select a manager to update.")
            return

        data = self.read_form()
        if data is None:
            return

        # Keep same ID & row
        selected = self.managers[idx]
        data["id"] = selected["id"]
        row_index = selected["row_index"]

        write_manager_to_excel(row_index, data)
        self.reload_from_excel()

    def delete_manager(self):
        idx = self.get_selected_index()
        if idx is None:
            messagebox.showwarning("Selection error", "Please select a manager to delete.")
            return

        selected = self.managers[idx]
        if not messagebox.askyesno("Confirm delete", f"Delete '{selected['name']}'?"):
            return

        delete_row_from_excel(selected["row_index"])
        self.reload_from_excel()
        self.clear_form()

    def on_select(self, event):
        idx = self.get_selected_index()
        if idx is None:
            return

        m = self.managers[idx]
        self.entry_name.delete(0, tk.END)
        self.entry_name.insert(0, m["name"])

        self.role_var.set(m["role"] if m["role"] in ROLES else ROLES[0])
        self.gender_var.set(m["gender"] if m["gender"] in GENDERS else GENDERS[0])

        for day in DAYS:
            start_val, end_val = m["availability"].get(day, (None, None))
            self.avail_start_vars[day].set("" if start_val is None else str(start_val))
            self.avail_end_vars[day].set("" if end_val is None else str(end_val))


if __name__ == "__main__":
    root = tk.Tk()
    app = ManagersApp(root)
    root.mainloop()
