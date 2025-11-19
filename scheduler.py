import tkinter as tk
from tkinter import messagebox

class NamesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Names Manager")
        self.root.geometry("400x300")

        # Internal storage for names
        self.names = ["Alice", "Bob", "Charlie"]  # starting data, can be empty

        # --- UI ELEMENTS ---

        # Listbox label
        self.label_list = tk.Label(root, text="Names:")
        self.label_list.pack(pady=(10, 0))

        # Listbox + scrollbar
        list_frame = tk.Frame(root)
        list_frame.pack(pady=5, fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(list_frame)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.listbox.yview)

        # Entry label
        self.label_entry = tk.Label(root, text="Name:")
        self.label_entry.pack()

        # Text entry
        self.entry = tk.Entry(root)
        self.entry.pack(fill=tk.X, padx=20)

        # Buttons frame
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        self.btn_add = tk.Button(btn_frame, text="Add", width=10, command=self.add_name)
        self.btn_add.grid(row=0, column=0, padx=5)

        self.btn_update = tk.Button(btn_frame, text="Update", width=10, command=self.update_name)
        self.btn_update.grid(row=0, column=1, padx=5)

        self.btn_delete = tk.Button(btn_frame, text="Delete", width=10, command=self.delete_name)
        self.btn_delete.grid(row=0, column=2, padx=5)

        # When you click on a name, load it into the entry box
        self.listbox.bind("<<ListboxSelect>>", self.on_select)

        # Load initial data
        self.refresh_listbox()

    # --- Helper methods ---

    def refresh_listbox(self):
        """Refresh Listbox contents from self.names."""
        self.listbox.delete(0, tk.END)
        for name in self.names:
            self.listbox.insert(tk.END, name)

    def get_selected_index(self):
        """Return selected index or None if nothing is selected."""
        selection = self.listbox.curselection()
        if not selection:
            return None
        return selection[0]

    # --- Button actions ---

    def add_name(self):
        new_name = self.entry.get().strip()
        if not new_name:
            messagebox.showwarning("Input error", "Name cannot be empty.")
            return

        self.names.append(new_name)
        self.refresh_listbox()
        self.entry.delete(0, tk.END)

    def update_name(self):
        idx = self.get_selected_index()
        if idx is None:
            messagebox.showwarning("Selection error", "Please select a name to update.")
            return

        new_name = self.entry.get().strip()
        if not new_name:
            messagebox.showwarning("Input error", "Name cannot be empty.")
            return

        self.names[idx] = new_name
        self.refresh_listbox()

    def delete_name(self):
        idx = self.get_selected_index()
        if idx is None:
            messagebox.showwarning("Selection error", "Please select a name to delete.")
            return

        name = self.names[idx]
        confirm = messagebox.askyesno("Confirm delete", f"Delete '{name}'?")
        if confirm:
            del self.names[idx]
            self.refresh_listbox()
            self.entry.delete(0, tk.END)

    def on_select(self, event):
        idx = self.get_selected_index()
        if idx is not None:
            selected_name = self.names[idx]
            self.entry.delete(0, tk.END)
            self.entry.insert(0, selected_name)


if __name__ == "__main__":
    root = tk.Tk()
    app = NamesApp(root)
    root.mainloop()
