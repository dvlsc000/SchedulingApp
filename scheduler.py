import tkinter as tk
from tkinter import messagebox

import firebase_admin
from firebase_admin import credentials, firestore

# -------- Firestore setup --------

# Path to your service account key JSON file
SERVICE_ACCOUNT_PATH = "serviceKey.json"

try:
    # Only initialize once
    if not firebase_admin._apps:
        cred = credentials.Certificate(SERVICE_ACCOUNT_PATH)
        firebase_admin.initialize_app(cred)

    db = firestore.client()
except Exception as e:
    print("Error initializing Firebase:", e)
    db = None


class NamesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Names Manager (Firestore)")
        self.root.geometry("400x300")

        if db is None:
            messagebox.showerror(
                "Firestore error",
                "Could not initialize Firestore. Check console for details."
            )
            self.root.destroy()
            return

        # Will store list of dicts: { "id": <doc_id>, "name": <name> }
        self.names_docs = []

        # --- UI elements ---

        self.label_list = tk.Label(root, text="Names from Firestore:")
        self.label_list.pack(pady=(10, 0))

        list_frame = tk.Frame(root)
        list_frame.pack(pady=5, fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(list_frame)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.listbox.yview)

        self.label_entry = tk.Label(root, text="Name:")
        self.label_entry.pack()

        self.entry = tk.Entry(root)
        self.entry.pack(fill=tk.X, padx=20)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        self.btn_add = tk.Button(btn_frame, text="Add", width=10, command=self.add_name)
        self.btn_add.grid(row=0, column=0, padx=5)

        self.btn_update = tk.Button(btn_frame, text="Update", width=10, command=self.update_name)
        self.btn_update.grid(row=0, column=1, padx=5)

        self.btn_delete = tk.Button(btn_frame, text="Delete", width=10, command=self.delete_name)
        self.btn_delete.grid(row=0, column=2, padx=5)

        self.listbox.bind("<<ListboxSelect>>", self.on_select)

        # Load from Firestore
        self.load_from_firestore()

    # -------- Firestore interaction methods --------

    def load_from_firestore(self):
        """Load all names from Firestore into self.names_docs and refresh the UI."""
        try:
            # query collection "names", order by 'name' field if you like
            docs = db.collection("names").order_by("name").stream()
            self.names_docs = []

            for d in docs:
                data = d.to_dict() or {}
                name = data.get("name", "")
                self.names_docs.append({"id": d.id, "name": name})

            self.refresh_listbox()

        except Exception as e:
            messagebox.showerror("Firestore error", f"Error loading data:\n{e}")

    def add_to_firestore(self, name: str):
        """Add a new name document to Firestore."""
        try:
            db.collection("names").add({"name": name})
        except Exception as e:
            messagebox.showerror("Firestore error", f"Error adding name:\n{e}")

    def update_firestore(self, doc_id: str, new_name: str):
        """Update an existing name document."""
        try:
            db.collection("names").document(doc_id).update({"name": new_name})
        except Exception as e:
            messagebox.showerror("Firestore error", f"Error updating name:\n{e}")

    def delete_from_firestore(self, doc_id: str):
        """Delete a name document from Firestore."""
        try:
            db.collection("names").document(doc_id).delete()
        except Exception as e:
            messagebox.showerror("Firestore error", f"Error deleting name:\n{e}")

    # -------- UI helpers --------

    def refresh_listbox(self):
        """Refresh Listbox contents from self.names_docs."""
        self.listbox.delete(0, tk.END)
        for item in self.names_docs:
            self.listbox.insert(tk.END, item["name"])

    def get_selected_index(self):
        """Return selected index or None if nothing is selected."""
        selection = self.listbox.curselection()
        if not selection:
            return None
        return selection[0]

    # -------- Button actions --------

    def add_name(self):
        new_name = self.entry.get().strip()
        if not new_name:
            messagebox.showwarning("Input error", "Name cannot be empty.")
            return

        # Add to Firestore
        self.add_to_firestore(new_name)
        # Reload list
        self.load_from_firestore()
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

        doc = self.names_docs[idx]
        doc_id = doc["id"]

        self.update_firestore(doc_id, new_name)
        self.load_from_firestore()

    def delete_name(self):
        idx = self.get_selected_index()
        if idx is None:
            messagebox.showwarning("Selection error", "Please select a name to delete.")
            return

        doc = self.names_docs[idx]
        name = doc["name"]
        doc_id = doc["id"]

        confirm = messagebox.askyesno("Confirm delete", f"Delete '{name}'?")
        if confirm:
            self.delete_from_firestore(doc_id)
            self.load_from_firestore()
            self.entry.delete(0, tk.END)

    def on_select(self, event):
        idx = self.get_selected_index()
        if idx is not None and 0 <= idx < len(self.names_docs):
            selected_name = self.names_docs[idx]["name"]
            self.entry.delete(0, tk.END)
            self.entry.insert(0, selected_name)


if __name__ == "__main__":
    root = tk.Tk()
    app = NamesApp(root)
    root.mainloop()
