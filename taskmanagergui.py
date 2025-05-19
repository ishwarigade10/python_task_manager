import tkinter as tk
from tkinter import messagebox, ttk
from tkcalendar import DateEntry
from datetime import datetime
import winsound

class TaskManagerGUI:
    def _init_(self, root):
        self.root = root
        self.root.title("Task Manager")
        self.tasks = {}

        # Task form
        form_frame = tk.LabelFrame(root, text="Task Details", padx=10, pady=10)
        form_frame.pack(padx=10, pady=5, fill="x")

        tk.Label(form_frame, text="Name:").grid(row=0, column=0, sticky="e")
        tk.Label(form_frame, text="Description:").grid(row=1, column=0, sticky="e")
        tk.Label(form_frame, text="Start Date:").grid(row=0, column=2, sticky="e")
        tk.Label(form_frame, text="Start Time (HH:MM AM/PM):").grid(row=1, column=2, sticky="e")
        tk.Label(form_frame, text="Deadline Date:").grid(row=0, column=4, sticky="e")
        tk.Label(form_frame, text="Deadline Time (HH:MM AM/PM):").grid(row=1, column=4, sticky="e")
        tk.Label(form_frame, text="Priority:").grid(row=0, column=6, sticky="e")

        self.name_entry = tk.Entry(form_frame, width=20)
        self.desc_entry = tk.Entry(form_frame, width=30)
        self.start_date_entry = DateEntry(form_frame, date_pattern='yyyy-mm-dd')
        self.start_time_entry = tk.Entry(form_frame, width=10)
        self.deadline_date_entry = DateEntry(form_frame, date_pattern='yyyy-mm-dd')
        self.deadline_time_entry = tk.Entry(form_frame, width=10)
        self.priority_combo = ttk.Combobox(form_frame, values=["High", "Medium", "Low"], width=10, state="readonly")
        self.priority_combo.current(1)

        self.name_entry.grid(row=0, column=1, padx=5)
        self.desc_entry.grid(row=1, column=1, padx=5)
        self.start_date_entry.grid(row=0, column=3, padx=5)
        self.start_time_entry.grid(row=1, column=3, padx=5)
        self.deadline_date_entry.grid(row=0, column=5, padx=5)
        self.deadline_time_entry.grid(row=1, column=5, padx=5)
        self.priority_combo.grid(row=0, column=7, padx=5)

        # Buttons
        tk.Button(form_frame, text="Add Task", command=self.add_task).grid(row=2, column=1, pady=10)
        tk.Button(form_frame, text="Update Task", command=self.update_task).grid(row=2, column=2, pady=10)
        tk.Button(form_frame, text="Delete Task", command=self.delete_task).grid(row=2, column=3, pady=10)
        tk.Button(form_frame, text="Mark Completed", command=self.mark_completed).grid(row=2, column=4, pady=10)
        tk.Button(form_frame, text="Refresh", command=self.display_tasks).grid(row=2, column=5, pady=10)

        # Treeview for displaying tasks
        self.tree = ttk.Treeview(root, columns=("Priority", "Description", "Start", "Deadline", "Time Left", "Status"), show='headings')
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=140, anchor='center')
        self.tree.pack(padx=10, pady=10, fill="x")

    def make_sound(self):
        winsound.Beep(600, 1000)

    def get_datetime_string(self, date_widget, time_widget):
        return f"{date_widget.get()} {time_widget.get()}"

    def add_task(self):
        name = self.name_entry.get().strip()
        if not name or name in self.tasks:
            messagebox.showwarning("Duplicate or Missing", "Task name is either missing or already exists.")
            return
        self._save_task(name)

    def update_task(self):
        name = self.name_entry.get().strip()
        if name not in self.tasks:
            messagebox.showerror("Not Found", "Task not found for update.")
            return
        self._save_task(name)

    def _save_task(self, name):
        desc = self.desc_entry.get().strip()
        start = self.get_datetime_string(self.start_date_entry, self.start_time_entry)
        deadline = self.get_datetime_string(self.deadline_date_entry, self.deadline_time_entry)
        priority = self.priority_combo.get()

        if not all([name, desc, start, deadline, priority]):
            messagebox.showwarning("Missing Info", "Please fill all fields.")
            return

        self.tasks[name] = {
            "description": desc,
            "start_date": start,
            "deadline_date": deadline,
            "status": self.tasks.get(name, {}).get("status", "Pending"),
            "priority": priority
        }

        self.clear_form()
        self.display_tasks()

    def delete_task(self):
        name = self.name_entry.get().strip()
        if name in self.tasks:
            del self.tasks[name]
            self.clear_form()
            self.display_tasks()
        else:
            messagebox.showerror("Not Found", "Task not found.")

    def mark_completed(self):
        name = self.name_entry.get().strip()
        if name in self.tasks:
            self.tasks[name]['status'] = "Completed"
            self.display_tasks()
        else:
            messagebox.showerror("Not Found", "Task not found.")

    def clear_form(self):
        self.name_entry.delete(0, tk.END)
        self.desc_entry.delete(0, tk.END)
        self.start_time_entry.delete(0, tk.END)
        self.deadline_time_entry.delete(0, tk.END)
        self.priority_combo.current(1)

    def display_tasks(self):
        self.tree.delete(*self.tree.get_children())
        for task, details in self.tasks.items():
            try:
                deadline = datetime.strptime(details['deadline_date'], "%Y-%m-%d %I:%M %p")
                time_left = deadline - datetime.now()

                if 0 < time_left.total_seconds() <= 300:
                    self.make_sound()
                    time_str = f" {int(time_left.total_seconds() // 60)}m left!"
                elif time_left.total_seconds() > 0:
                    days, seconds = divmod(time_left.total_seconds(), 86400)
                    hours, seconds = divmod(seconds, 3600)
                    minutes, seconds = divmod(seconds, 60)
                    time_str = f"{int(days)}d {int(hours)}h {int(minutes)}m {int(seconds)}s"
                else:
                    time_str = "Deadline Passed"
            except:
                time_str = "Invalid Date"

            self.tree.insert("", "end", values=(
                details["priority"],
                details["description"],
                details["start_date"],
                details["deadline_date"],
                time_str,
                details["status"]
            ))

# Run the app
if _name_ == "_main_":
    root = tk.Tk()
    app = TaskManagerGUI(root)
    root.mainloop()