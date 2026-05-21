"""A lightweight desktop todo list application using Python + tkinter."""

import json
import os
import sys
from dataclasses import dataclass, asdict
from datetime import datetime
from tkinter import ttk, messagebox, simpledialog
import tkinter as tk

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
APP_NAME = "TodoApp"
DATA_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), APP_NAME)
DATA_FILE = os.path.join(DATA_DIR, "tasks.json")

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class Task:
    id: int
    description: str
    completed: bool = False
    created_at: str = ""

# ---------------------------------------------------------------------------
# Global state
# ---------------------------------------------------------------------------
_tasks: list[Task] = []
_task_listbox: tk.Listbox | None = None
_entry: ttk.Entry | None = None
_toggle_btn: ttk.Button | None = None
_edit_btn: ttk.Button | None = None
_delete_btn: ttk.Button | None = None

# ---------------------------------------------------------------------------
# Persistence
# ---------------------------------------------------------------------------

def _load_tasks() -> list[Task]:
    if not os.path.exists(DATA_FILE):
        return []
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return [Task(**t) for t in data.get("tasks", [])]
    except (json.JSONDecodeError, TypeError) as e:
        print(f"Warning: failed to load tasks file — {e}", file=sys.stderr)
        return []


def _save_tasks() -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump({"tasks": [asdict(t) for t in _tasks]}, f, ensure_ascii=False, indent=2)

# ---------------------------------------------------------------------------
# UI helpers
# ---------------------------------------------------------------------------

def _refresh_listbox() -> None:
    assert _task_listbox is not None
    selected_idx = _task_listbox.curselection()
    selected_id: int | None = None
    if selected_idx:
        selected_id = _tasks[selected_idx[0]].id

    _task_listbox.delete(0, tk.END)
    for task in _tasks:
        prefix = "[x]" if task.completed else "[ ]"
        _task_listbox.insert(tk.END, f"{prefix} {task.description}")
        if task.completed:
            color = "#888888"
            _task_listbox.itemconfig(tk.END, fg=color)
        else:
            _task_listbox.itemconfig(tk.END, fg="black")

    # Restore selection
    if selected_id is not None:
        for i, task in enumerate(_tasks):
            if task.id == selected_id:
                _task_listbox.selection_set(i)
                break

    _update_button_states()


def _update_button_states() -> None:
    assert _task_listbox is not None
    assert _toggle_btn is not None
    assert _edit_btn is not None
    assert _delete_btn is not None

    selection = _task_listbox.curselection()
    if not selection:
        _toggle_btn.config(text="Mark Done", state=tk.DISABLED)
        _edit_btn.config(state=tk.DISABLED)
        _delete_btn.config(state=tk.DISABLED)
        return

    task = _tasks[selection[0]]
    _toggle_btn.config(
        text="Mark Undone" if task.completed else "Mark Done",
        state=tk.NORMAL,
    )
    _edit_btn.config(state=tk.NORMAL)
    _delete_btn.config(state=tk.NORMAL)


def _get_selected_task() -> Task | None:
    assert _task_listbox is not None
    selection = _task_listbox.curselection()
    if not selection:
        return None
    return _tasks[selection[0]]

# ---------------------------------------------------------------------------
# Actions
# ---------------------------------------------------------------------------

def _add_task(event: tk.Event | None = None) -> None:
    assert _entry is not None
    desc = _entry.get().strip()
    if not desc:
        messagebox.showwarning("Empty Task", "Task description cannot be empty.")
        return

    new_id = max((t.id for t in _tasks), default=0) + 1
    task = Task(id=new_id, description=desc, created_at=datetime.now().isoformat())
    _tasks.append(task)
    _entry.delete(0, tk.END)
    _save_tasks()
    _refresh_listbox()


def _toggle_complete(event: tk.Event | None = None) -> None:
    task = _get_selected_task()
    if task is None:
        return
    task.completed = not task.completed
    _save_tasks()
    _refresh_listbox()


def _edit_task(event: tk.Event | None = None) -> None:
    task = _get_selected_task()
    if task is None:
        return

    new_desc = simpledialog.askstring("Edit Task", "Modify description:", initialvalue=task.description)
    if new_desc is None:  # cancelled
        return
    new_desc = new_desc.strip()
    if not new_desc:
        messagebox.showwarning("Empty Task", "Task description cannot be empty.")
        return

    task.description = new_desc
    _save_tasks()
    _refresh_listbox()


def _delete_task(event: tk.Event | None = None) -> None:
    task = _get_selected_task()
    if task is None:
        return

    confirmed = messagebox.askyesno("Confirm Delete", f'Delete "{task.description}"?')
    if not confirmed:
        return

    _tasks.remove(task)
    _save_tasks()
    _refresh_listbox()


def _clear_done() -> None:
    global _tasks
    remaining = [t for t in _tasks if not t.completed]
    if len(remaining) == len(_tasks):
        return
    _tasks = remaining
    _save_tasks()
    _refresh_listbox()

# ---------------------------------------------------------------------------
# UI construction
# ---------------------------------------------------------------------------

def _build_ui(root: tk.Tk) -> None:
    global _task_listbox, _entry, _toggle_btn, _edit_btn, _delete_btn

    root.title("Todo List")
    root.geometry("500x420")
    root.minsize(350, 280)
    root.option_add("*Font", ("Segoe UI", 10))

    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    # --- Top frame: entry + add button ---
    top_frame = ttk.Frame(root, padding=(8, 8, 8, 4))
    top_frame.grid(row=0, column=0, sticky="ew")
    top_frame.columnconfigure(0, weight=1)

    _entry = ttk.Entry(top_frame)
    _entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
    _entry.bind("<Return>", _add_task)

    add_btn = ttk.Button(top_frame, text="Add Task", command=_add_task)
    add_btn.grid(row=0, column=1)

    # --- Middle frame: listbox + scrollbar ---
    list_frame = ttk.Frame(root, padding=(8, 4))
    list_frame.grid(row=1, column=0, sticky="nsew")
    list_frame.rowconfigure(0, weight=1)
    list_frame.columnconfigure(0, weight=1)

    _task_listbox = tk.Listbox(
        list_frame,
        selectmode=tk.BROWSE,
        activestyle="dotbox",
        relief=tk.FLAT,
        highlightthickness=1,
        highlightbackground="#c0c0c0",
        highlightcolor="#4a86e8",
    )
    _task_listbox.grid(row=0, column=0, sticky="nsew")
    _task_listbox.bind("<<ListboxSelect>>", lambda e: _update_button_states())
    _task_listbox.bind("<Double-Button-1>", lambda e: _toggle_complete())
    _task_listbox.bind("<Delete>", _delete_task)
    _task_listbox.bind("<Control-e>", _edit_task)

    scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=_task_listbox.yview)
    scrollbar.grid(row=0, column=1, sticky="ns")
    _task_listbox.config(yscrollcommand=scrollbar.set)

    # --- Bottom frame: action buttons ---
    bottom_frame = ttk.Frame(root, padding=(8, 4, 8, 8))
    bottom_frame.grid(row=2, column=0, sticky="ew")

    _toggle_btn = ttk.Button(bottom_frame, text="Mark Done", command=_toggle_complete, state=tk.DISABLED)
    _toggle_btn.pack(side=tk.LEFT, padx=(0, 4))

    _edit_btn = ttk.Button(bottom_frame, text="Edit", command=_edit_task, state=tk.DISABLED)
    _edit_btn.pack(side=tk.LEFT, padx=(0, 4))

    _delete_btn = ttk.Button(bottom_frame, text="Delete", command=_delete_task, state=tk.DISABLED)
    _delete_btn.pack(side=tk.LEFT, padx=(0, 4))

    clear_btn = ttk.Button(bottom_frame, text="Clear Done", command=_clear_done)
    clear_btn.pack(side=tk.LEFT)

    # --- Window close handler ---
    root.protocol("WM_DELETE_WINDOW", lambda: (_save_tasks(), root.destroy()))

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    global _tasks
    _tasks = _load_tasks()

    root = tk.Tk()
    _build_ui(root)
    _refresh_listbox()
    root.mainloop()


if __name__ == "__main__":
    main()
