import csv
from datetime import datetime
from difflib import SequenceMatcher
import hashlib
import json
import math
import os
import re
import shutil
import threading
import queue
import zipfile
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from xml.etree import ElementTree


APP_NAME = "Enterprise File Manager"
APP_VERSION = "3.3"
META_NAME = "filedata.efm"
SETTINGS_NAME = "settings.efs"
DUPLICATES_NAME = "duplicates.efs"
LEGACY_SETTINGS_NAME = "settings.efm"
LEGACY_GROUPS_NAME = "groups.efm"
DEFAULT_STATUS_OPTIONS = ("Current", "Legacy (Still Relevant)", "Historic (Archived)")
METADATA_COLUMNS = ["FileName", "DocumentTitle", "Type", "Number", "IssueDate", "SheetNo", "TotalSheets", "Version", "Status", "Groups", "AssociatedDocs", "Attachments", "VersionHistory"]


def human_size(size):
    value = float(size)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if value < 1024 or unit == "TB":
            return f"{value:.0f} {unit}" if unit == "B" else f"{value:.2f} {unit}"
        value /= 1024


class EFM(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("1240x760")
        self.minsize(1040, 650)
        self.configure(bg="#eef2f5")
        self.folder = None
        self.folder_files = []
        self.files = []
        self.filtered_files = []
        self.data = {}
        self.duplicate_audit = {"format": "EFM duplicate audit", "version": 1, "records": []}
        self.selected_name = None
        self.navigation_history = []
        self.groups = []
        self.defined_groups = []
        self.group_descriptions = {}
        self.defined_statuses = list(DEFAULT_STATUS_OPTIONS)
        self.status_descriptions = {}
        self.tooltip_window = None
        self.tooltip_key = None
        self.associations = []
        self.attachments = []
        self.version_history = []
        self.settings_window = None
        self._style()
        self._build()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame", background="#eef2f5")
        s.configure("Card.TFrame", background="#ffffff")
        s.configure("TLabel", background="#eef2f5", foreground="#243746", font=("Segoe UI", 9))
        s.configure("Card.TLabel", background="#ffffff", foreground="#243746", font=("Segoe UI", 9))
        s.configure("Title.TLabel", background="#ffffff", foreground="#17324d", font=("Segoe UI", 12, "bold"))
        s.configure("Muted.TLabel", background="#ffffff", foreground="#71808b", font=("Segoe UI", 9))
        s.configure("Header.TLabel", background="#17324d", foreground="#ffffff", font=("Segoe UI", 18, "bold"))
        s.configure("SubHeader.TLabel", background="#17324d", foreground="#d7e5ef", font=("Segoe UI", 9))
        s.configure("TButton", font=("Segoe UI", 9), padding=(11, 7))
        s.configure("Primary.TButton", background="#087e8b", foreground="#ffffff")
        s.map("Primary.TButton", background=[("active", "#066b76"), ("disabled", "#91adb1")])
        s.configure("Secondary.TButton", background="#dce6ed", foreground="#17324d")
        s.map("Secondary.TButton", background=[("active", "#cbdce5")])
        s.configure("EFM.TNotebook", background="#ffffff", borderwidth=0)
        s.configure("EFM.TNotebook.Tab", font=("Segoe UI", 9, "bold"), padding=(14, 8), background="#dce6ed", foreground="#17324d")
        s.map("EFM.TNotebook.Tab", background=[("selected", "#087e8b"), ("active", "#cbdce5")], foreground=[("selected", "#ffffff")])
        s.configure("Treeview", font=("Segoe UI", 9), rowheight=28, background="#ffffff", fieldbackground="#ffffff", borderwidth=0)
        s.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#dce6ed", foreground="#17324d")
        s.map("Treeview", background=[("selected", "#b8dce2")], foreground=[("selected", "#102a3e")])

    def _build(self):
        header = tk.Frame(self, bg="#17324d", height=76)
        header.pack(fill="x")
        header.pack_propagate(False)
        ttk.Label(header, text=APP_NAME, style="Header.TLabel").pack(anchor="w", padx=20, pady=(12, 0))
        self.path_var = tk.StringVar(value="Select a document folder to begin")
        ttk.Label(header, textvariable=self.path_var, style="SubHeader.TLabel").pack(anchor="w", padx=21)
        settings_button = tk.Button(
            header,
            text="⚙",
            command=self.open_settings,
            font=("Segoe UI Symbol", 17),
            bg="#17324d",
            fg="#ffffff",
            activebackground="#244b6d",
            activeforeground="#ffffff",
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            cursor="hand2",
            width=2,
            takefocus=True,
        )
        settings_button.place(relx=1.0, x=-18, y=14, anchor="ne")

        toolbar = ttk.Frame(self, padding=(16, 12, 16, 8))
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="Select Folder", style="Primary.TButton", command=self.select_folder).pack(side="left")
        self.refresh_button = ttk.Button(toolbar, text="Refresh List", style="Secondary.TButton", command=self.load_files, state="disabled")
        self.refresh_button.pack(side="left", padx=(7, 0))
        self.manage_button = ttk.Button(toolbar, text="Manage List", style="Secondary.TButton", command=self.open_document_management, state="disabled")
        self.manage_button.pack(side="left", padx=(7, 0))
        self.open_button = ttk.Button(toolbar, text="Open File", style="Secondary.TButton", command=self.open_file, state="disabled")
        self.open_button.pack(side="left", padx=(7, 0))
        self.back_button = ttk.Button(toolbar, text="Back", style="Secondary.TButton", command=self.go_back, state="disabled")
        self.back_button.pack(side="left", padx=(7, 0))
        self.summary_var = tk.StringVar(value="No folder selected")
        ttk.Label(toolbar, textvariable=self.summary_var).pack(side="right", padx=4)

        paned = ttk.Panedwindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        left = ttk.Frame(paned, style="Card.TFrame", padding=12)
        right = ttk.Frame(paned, style="Card.TFrame", padding=12)
        paned.add(left, weight=2)
        paned.add(right, weight=3)

        ttk.Label(left, text="Documents", style="Title.TLabel").pack(anchor="w")
        search_row = ttk.Frame(left, style="Card.TFrame")
        search_row.pack(fill="x", pady=(8, 9))
        ttk.Label(search_row, text="Search", style="Card.TLabel").pack(side="left", padx=(0, 8))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.apply_filter())
        ttk.Entry(search_row, textvariable=self.search_var).pack(side="left", fill="x", expand=True)

        filter_row = ttk.Frame(left, style="Card.TFrame")
        filter_row.pack(fill="x", pady=(0, 9))
        ttk.Label(filter_row, text="Status", style="Card.TLabel").pack(side="left", padx=(0, 10))
        self.status_filter_var = tk.StringVar(value="All statuses")
        self.status_filter_var.trace_add("write", lambda *_: self.apply_filter())
        self.status_filter_entry = ttk.Combobox(
            filter_row,
            textvariable=self.status_filter_var,
            state="readonly",
            values=("All statuses", *self.defined_statuses),
        )
        self.status_filter_entry.pack(side="left", fill="x", expand=True)
        self.bind_value_tooltip(self.status_filter_entry, self.status_filter_var, lambda: self.status_descriptions)

        self.tree = ttk.Treeview(left, columns=("title", "type", "status"), show="tree headings", selectmode="browse")
        self.tree.heading("#0", text="Name", anchor="w")
        self.tree.heading("title", text="Document Title", anchor="w")
        self.tree.heading("type", text="Filetype", anchor="w")
        self.tree.heading("status", text="Status", anchor="w")
        self.tree.column("#0", width=210, minwidth=130)
        self.tree.column("title", width=220, minwidth=130)
        self.tree.column("type", width=75, minwidth=60, stretch=False)
        self.tree.column("status", width=145, minwidth=100, stretch=False)
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", lambda _e: self.open_file())
        self.bind_tree_tooltip(self.tree, "status", lambda: self.status_descriptions)

        ttk.Label(right, text="Document Details", style="Title.TLabel").pack(anchor="w")
        self.details_notebook = ttk.Notebook(right, style="EFM.TNotebook")
        self.details_notebook.pack(fill="both", expand=True, pady=(8, 0))
        self.metadata_tab = ttk.Frame(self.details_notebook, style="Card.TFrame", padding=12)
        self.details_notebook.add(self.metadata_tab, text="Metadata")
        groups_tab = ttk.Frame(self.details_notebook, style="Card.TFrame", padding=12)
        association_tab = ttk.Frame(self.details_notebook, style="Card.TFrame", padding=12)
        attachments_tab = ttk.Frame(self.details_notebook, style="Card.TFrame", padding=12)
        network_tab = ttk.Frame(self.details_notebook, style="Card.TFrame", padding=12)
        history_tab = ttk.Frame(self.details_notebook, style="Card.TFrame", padding=12)
        self.details_notebook.add(groups_tab, text="Groups")
        self.details_notebook.add(association_tab, text="Associations")
        self.details_notebook.add(attachments_tab, text="Attachments")
        self.details_notebook.add(network_tab, text="Network")
        self.details_notebook.add(history_tab, text="Version History")

        info = ttk.Frame(self.metadata_tab, style="Card.TFrame")
        info.pack(fill="x", pady=(0, 10))
        self.file_info_vars = {k: tk.StringVar(value="—") for k in ("Name", "Filetype", "Size")}
        for row, key in enumerate(self.file_info_vars):
            ttk.Label(info, text=f"{key}:", style="Card.TLabel", width=14).grid(row=row, column=0, sticky="w", pady=2)
            ttk.Label(info, textvariable=self.file_info_vars[key], style="Muted.TLabel").grid(row=row, column=1, sticky="w", pady=2)

        ttk.Separator(self.metadata_tab).pack(fill="x", pady=(0, 12))
        form = ttk.Frame(self.metadata_tab, style="Card.TFrame")
        form.pack(fill="x")
        self.fields = {}
        labels = [("Document Title", "DocumentTitle"), ("Type", "Type"), ("Number", "Number"), ("Sheet number", "SheetNo"), ("Total sheets", "TotalSheets"), ("Version", "Version"), ("Issue date", "IssueDate")]
        for row, (label, key) in enumerate(labels):
            ttk.Label(form, text=f"{label}:", style="Card.TLabel", width=16).grid(row=row, column=0, sticky="w", pady=5)
            var = tk.StringVar()
            ttk.Entry(form, textvariable=var).grid(row=row, column=1, sticky="ew", pady=5)
            self.fields[key] = var
        ttk.Label(form, text="Status:", style="Card.TLabel", width=16).grid(row=7, column=0, sticky="w", pady=5)
        self.fields["Status"] = tk.StringVar(value=self.default_status())
        self.status_entry = ttk.Combobox(form, textvariable=self.fields["Status"], state="readonly", values=self.defined_statuses)
        self.status_entry.grid(row=7, column=1, sticky="ew", pady=5)
        self.bind_value_tooltip(self.status_entry, self.fields["Status"], lambda: self.status_descriptions)
        form.columnconfigure(1, weight=1)
        actions = ttk.Frame(self.metadata_tab, style="Card.TFrame")
        actions.pack(fill="x", pady=(12, 0))
        self.save_buttons = []
        save_button = ttk.Button(actions, text="Update Details", style="Primary.TButton", command=self.save_current, state="disabled")
        save_button.pack(side="right")
        self.save_buttons.append(save_button)

        self._build_groups_tab(groups_tab)
        self._build_link_tab(association_tab, "association")
        self._build_attachments_tab(attachments_tab)
        self._build_network_tab(network_tab)
        self._build_link_tab(history_tab, "history")

    def open_settings(self):
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.lift()
            self.settings_window.focus_force()
            return
        self.settings_window = tk.Toplevel(self)
        self.settings_window.title("Settings")
        self.settings_window.transient(self)
        self.settings_window.geometry("560x400")
        self.settings_window.minsize(420, 300)
        self.settings_window.configure(bg="#ffffff")
        settings_content = ttk.Frame(self.settings_window, style="Card.TFrame", padding=18)
        settings_content.pack(fill="both", expand=True)
        ttk.Label(settings_content, text="Settings", style="Title.TLabel").pack(anchor="w", pady=(0, 14))
        manager_buttons = ttk.Frame(settings_content, style="Card.TFrame")
        manager_buttons.pack(fill="x", anchor="w")
        ttk.Button(manager_buttons, text="Manage Groups", style="Secondary.TButton", command=self.open_group_manager).pack(side="left")
        ttk.Button(manager_buttons, text="Manage Status", style="Secondary.TButton", command=self.open_status_manager).pack(side="left", padx=(8, 0))

        def close_settings():
            self.settings_window.destroy()
            self.settings_window = None

        self.settings_window.protocol("WM_DELETE_WINDOW", close_settings)

    def show_tooltip(self, text, x, y, key=None):
        if not text:
            self.hide_tooltip()
            return
        if self.tooltip_window is not None and self.tooltip_key == key:
            return
        self.hide_tooltip()
        self.tooltip_key = key
        self.tooltip_window = tk.Toplevel(self)
        self.tooltip_window.overrideredirect(True)
        self.tooltip_window.attributes("-topmost", True)
        self.tooltip_window.geometry(f"+{x + 12}+{y + 14}")
        tk.Label(
            self.tooltip_window,
            text=text,
            justify="left",
            wraplength=340,
            bg="#fff9d9",
            fg="#243746",
            relief="solid",
            borderwidth=1,
            padx=7,
            pady=5,
            font=("Segoe UI", 9),
        ).pack()

    def hide_tooltip(self, _event=None):
        if self.tooltip_window is not None:
            self.tooltip_window.destroy()
            self.tooltip_window = None
        self.tooltip_key = None

    def bind_value_tooltip(self, widget, variable, descriptions_getter):
        def show(event):
            value = variable.get().strip()
            description = descriptions_getter().get(value.casefold(), "") if value else ""
            self.show_tooltip(description, event.x_root, event.y_root, key=(str(widget), value.casefold()))

        widget.bind("<Enter>", show, add="+")
        widget.bind("<Leave>", self.hide_tooltip, add="+")

    def bind_tree_tooltip(self, tree, column, descriptions_getter):
        def show(event):
            item = tree.identify_row(event.y)
            value = tree.set(item, column).strip() if item else ""
            description = descriptions_getter().get(value.casefold(), "") if value else ""
            self.show_tooltip(description, event.x_root, event.y_root, key=(str(tree), item))

        tree.bind("<Motion>", show, add="+")
        tree.bind("<Leave>", self.hide_tooltip, add="+")

    def open_group_manager(self):
        self.open_definition_manager("group")

    def open_status_manager(self):
        self.open_definition_manager("status")

    def open_definition_manager(self, kind):
        if not self.folder:
            messagebox.showinfo(APP_NAME, "Select a document folder before managing its settings.", parent=self.settings_window)
            return
        is_group = kind == "group"
        title = "Manage Groups" if is_group else "Manage Status"
        singular = "Group" if is_group else "Status"
        manager = tk.Toplevel(self.settings_window)
        manager.title(title)
        manager.transient(self.settings_window)
        manager.grab_set()
        manager.geometry("700x500")
        manager.minsize(560, 400)
        ttk.Label(manager, text=title, padding=(16, 16, 16, 2), font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(manager, text=f"Define the {singular.lower()} options and supporting descriptions available to users.", padding=(16, 0, 16, 10)).pack(anchor="w")

        definition_tree = ttk.Treeview(manager, columns=("name", "description"), show="headings", selectmode="browse")
        definition_tree.heading("name", text=singular, anchor="w")
        definition_tree.heading("description", text="Description", anchor="w")
        definition_tree.column("name", width=210, minwidth=140)
        definition_tree.column("description", width=420, minwidth=220)
        definition_tree.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        editor = ttk.Frame(manager, padding=(16, 0, 16, 10))
        editor.pack(fill="x")
        editor.columnconfigure(1, weight=1)
        name_var = tk.StringVar()
        description_var = tk.StringVar()
        ttk.Label(editor, text=f"{singular}:", width=12).grid(row=0, column=0, sticky="w", pady=3)
        name_entry = ttk.Entry(editor, textvariable=name_var)
        name_entry.grid(row=0, column=1, sticky="ew", pady=3)
        ttk.Label(editor, text="Description:", width=12).grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(editor, textvariable=description_var).grid(row=1, column=1, sticky="ew", pady=3)
        editor_buttons = ttk.Frame(editor)
        editor_buttons.grid(row=0, column=2, rowspan=2, sticky="e", padx=(8, 0))
        save_entry_button = ttk.Button(editor_buttons, text=f"Add {singular}", style="Secondary.TButton")
        save_entry_button.pack(side="left")
        new_entry_button = ttk.Button(editor_buttons, text="New", style="Secondary.TButton", state="disabled")
        new_entry_button.pack(side="left", padx=(7, 0))
        remove_button = ttk.Button(editor_buttons, text="Remove", style="Secondary.TButton", state="disabled")
        remove_button.pack(side="left", padx=(7, 0))
        selected_name = [None]

        def definitions():
            return self.defined_groups if is_group else self.defined_statuses

        def descriptions():
            return self.group_descriptions if is_group else self.status_descriptions

        def refresh_manager():
            definition_tree.delete(*definition_tree.get_children())
            for index, value in enumerate(sorted(definitions(), key=str.casefold)):
                definition_tree.insert("", "end", iid=str(index), values=(value, descriptions().get(value.casefold(), "")))
            remove_button.configure(state="disabled")
            self.refresh_groups()
            self.refresh_status_options()

        def clear_editor():
            selected_name[0] = None
            definition_tree.selection_remove(*definition_tree.selection())
            name_entry.configure(state="normal")
            name_var.set("")
            description_var.set("")
            save_entry_button.configure(text=f"Add {singular}")
            new_entry_button.configure(state="disabled")
            remove_button.configure(state="disabled")
            name_entry.focus_set()

        def save_entry(_event=None):
            value = name_var.get().strip()
            description = description_var.get().strip()
            if selected_name[0]:
                descriptions()[selected_name[0].casefold()] = description
            else:
                if not value or any(existing.casefold() == value.casefold() for existing in definitions()):
                    return
                definitions().append(value)
                descriptions()[value.casefold()] = description
            if self.save_definitions():
                refresh_manager()
                clear_editor()
            else:
                self.load_definitions()
                refresh_manager()
                clear_editor()

        def select_entry(_event=None):
            selected = definition_tree.selection()
            if not selected:
                return
            value, description = definition_tree.item(selected[0], "values")
            selected_name[0] = value
            name_var.set(value)
            description_var.set(description)
            name_entry.configure(state="disabled")
            save_entry_button.configure(text="Update Description")
            new_entry_button.configure(state="normal")
            remove_button.configure(state="normal")

        def remove_definition():
            value = selected_name[0]
            if not value:
                return
            if not is_group and len(self.defined_statuses) == 1:
                messagebox.showinfo(APP_NAME, "At least one Status definition is required.", parent=manager)
                return
            if is_group:
                affected_documents = [
                    name for name, metadata in self.data.items()
                    if any(
                        item.casefold() == value.casefold()
                        for item in self.decode_groups(metadata.get("Groups", ""))
                    )
                ]
            else:
                affected_documents = [
                    name for name, metadata in self.data.items()
                    if (metadata.get("Status") or self.default_status()).casefold() == value.casefold()
                ]
            if affected_documents:
                preview = "\n".join(f"• {name}" for name in sorted(affected_documents, key=str.casefold)[:8])
                if len(affected_documents) > 8:
                    preview += f"\n• …and {len(affected_documents) - 8} more"
                prompt = (
                    f"Warning: '{value}' is currently used by {len(affected_documents)} document(s):\n\n"
                    f"{preview}\n\nRemoving it from the available {singular.lower()} options will not remove "
                    "the value from those documents. Do you want to continue?"
                )
                icon = "warning"
            else:
                prompt = f"Remove '{value}' from the available {singular.lower()} options?"
                icon = "question"
            if not messagebox.askyesno(
                f"Remove {singular} Definition",
                prompt,
                icon=icon,
                parent=manager,
            ):
                return
            old_values = list(definitions())
            old_description = descriptions().get(value.casefold(), "")
            updated = [item for item in definitions() if item.casefold() != value.casefold()]
            if is_group:
                self.defined_groups = updated
            else:
                self.defined_statuses = updated
            descriptions().pop(value.casefold(), None)
            if self.save_definitions():
                refresh_manager()
                clear_editor()
            else:
                if is_group:
                    self.defined_groups = old_values
                else:
                    self.defined_statuses = old_values
                descriptions()[value.casefold()] = old_description

        save_entry_button.configure(command=save_entry)
        new_entry_button.configure(command=clear_editor)
        remove_button.configure(command=remove_definition)
        name_entry.bind("<Return>", save_entry)
        definition_tree.bind("<<TreeviewSelect>>", select_entry)
        self.bind_tree_tooltip(definition_tree, "name", descriptions)
        actions = ttk.Frame(manager, padding=(16, 0, 16, 16))
        actions.pack(fill="x")
        ttk.Button(actions, text="Close", style="Primary.TButton", command=manager.destroy).pack(side="right")
        refresh_manager()
        clear_editor()
        self.wait_window(manager)

    def _build_groups_tab(self, tab):
        ttk.Label(tab, text="Document groups", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            tab,
            text="Assign this document to one or more groups defined in Settings.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(2, 9))
        self.groups_tree = ttk.Treeview(
            tab,
            columns=("group", "description"),
            show="headings",
            selectmode="browse",
        )
        self.groups_tree.heading("group", text="Group", anchor="w")
        self.groups_tree.heading("description", text="Description", anchor="w")
        self.groups_tree.column("group", width=190, minwidth=120)
        self.groups_tree.column("description", width=360, minwidth=180)
        self.groups_tree.pack(fill="both", expand=True)
        self.groups_tree.bind("<<TreeviewSelect>>", self._on_group_select)
        self.bind_tree_tooltip(self.groups_tree, "group", lambda: self.group_descriptions)

        entry_row = ttk.Frame(tab, style="Card.TFrame")
        entry_row.pack(fill="x", pady=(10, 0))
        self.group_var = tk.StringVar()
        self.group_entry = ttk.Combobox(entry_row, textvariable=self.group_var, state="readonly")
        self.group_entry.pack(side="left", fill="x", expand=True)
        self.group_entry.bind("<<ComboboxSelected>>", self._on_group_choice)
        self.bind_value_tooltip(self.group_entry, self.group_var, lambda: self.group_descriptions)
        self.add_group_button = ttk.Button(entry_row, text="Assign Group", style="Secondary.TButton", command=self.add_group, state="disabled")
        self.add_group_button.pack(side="left", padx=(7, 0))
        self.remove_group_button = ttk.Button(entry_row, text="Remove", style="Secondary.TButton", command=self.remove_group, state="disabled")
        self.remove_group_button.pack(side="left", padx=(7, 0))
        save_button = ttk.Button(entry_row, text="Update Details", style="Primary.TButton", command=self.save_current, state="disabled")
        save_button.pack(side="right", padx=(7, 0))
        self.save_buttons.append(save_button)

    def _build_network_tab(self, tab):
        ttk.Label(tab, text="Association network", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            tab,
            text="Only direct associations are shown. Click a document node to open its details.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(2, 9))
        self.network_canvas = tk.Canvas(tab, background="#f7fafc", highlightthickness=1, highlightbackground="#dce6ed")
        self.network_canvas.pack(fill="both", expand=True)
        self.network_canvas.bind("<Configure>", lambda _event: self.refresh_network())
        controls = ttk.Frame(tab, style="Card.TFrame")
        controls.pack(fill="x", pady=(10, 0))
        save_button = ttk.Button(controls, text="Update Details", style="Primary.TButton", command=self.save_current, state="disabled")
        save_button.pack(side="right")
        self.save_buttons.append(save_button)

    def _build_attachments_tab(self, tab):
        ttk.Label(tab, text="Document attachments", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            tab,
            text="Attached files remain in the selected folder but are hidden from the Documents list.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(2, 9))
        self.attachment_tree = ttk.Treeview(
            tab,
            columns=("name", "type"),
            show="headings",
            selectmode="extended",
        )
        self.attachment_tree.heading("name", text="Attachment Name", anchor="w")
        self.attachment_tree.heading("type", text="Filetype", anchor="w")
        self.attachment_tree.column("name", width=480, minwidth=220)
        self.attachment_tree.column("type", width=100, minwidth=70, stretch=False)
        self.attachment_tree.pack(fill="both", expand=True)
        controls = ttk.Frame(tab, style="Card.TFrame")
        controls.pack(fill="x", pady=(10, 0))
        ttk.Button(
            controls,
            text="Add Attachments…",
            style="Primary.TButton",
            command=self.add_attachments,
        ).pack(side="left")
        self.remove_attachment_button = ttk.Button(
            controls,
            text="Remove",
            style="Secondary.TButton",
            command=self.remove_attachments,
            state="disabled",
        )
        self.remove_attachment_button.pack(side="left", padx=(7, 0))
        self.open_attachment_button = ttk.Button(
            controls,
            text="Open Attachment",
            style="Secondary.TButton",
            command=self.open_attachment,
            state="disabled",
        )
        self.open_attachment_button.pack(side="left", padx=(7, 0))
        save_button = ttk.Button(
            controls,
            text="Update Details",
            style="Primary.TButton",
            command=self.save_current,
            state="disabled",
        )
        save_button.pack(side="right")
        self.save_buttons.append(save_button)
        self.attachment_tree.bind("<<TreeviewSelect>>", self._on_attachment_select)
        self.attachment_tree.bind("<Double-1>", lambda _event: self.open_attachment())

    def _build_link_tab(self, tab, kind):
        title = "Associated documents" if kind == "association" else "Earlier or superseded document versions"
        ttk.Label(tab, text=title, style="Title.TLabel").pack(anchor="w")
        hint = "Double-click a row to open that document in EFM." if kind == "association" else "Add linked documents to create a version trail. Double-click to open one in EFM."
        ttk.Label(tab, text=hint, style="Muted.TLabel").pack(anchor="w", pady=(2, 9))
        tree = ttk.Treeview(tab, columns=("name", "title", "version", "date"), show="headings", selectmode="browse")
        tree._heading_labels = {"name": "Document Name", "title": "Document Title", "version": "Version", "date": "Issue Date"}
        tree._sort_column = None
        tree._sort_descending = False
        for column, heading in tree._heading_labels.items():
            tree.heading(column, text=heading, anchor="w", command=lambda value=column, source=tree: self.sort_link_tree(source, value))
        tree.column("name", width=200, minwidth=120)
        tree.column("title", width=240, minwidth=130)
        tree.column("version", width=110, minwidth=75, stretch=False)
        tree.column("date", width=110, minwidth=85, stretch=False)
        tree.pack(fill="both", expand=True)
        tree.bind("<Double-1>", lambda _event, source=tree: self.open_link(source))
        controls = ttk.Frame(tab, style="Card.TFrame")
        controls.pack(fill="x", pady=(10, 0))
        ttk.Button(controls, text="Add Document…", style="Primary.TButton", command=lambda value=kind: self.add_link(value)).pack(side="left")
        remove = ttk.Button(controls, text="Remove", style="Secondary.TButton", command=lambda value=kind: self.remove_link(value), state="disabled")
        remove.pack(side="left", padx=(7, 0))
        save_button = ttk.Button(controls, text="Update Details", style="Primary.TButton", command=self.save_current, state="disabled")
        save_button.pack(side="right")
        self.save_buttons.append(save_button)
        tree.bind("<<TreeviewSelect>>", lambda _event, button=remove: button.configure(state="normal" if tree.selection() else "disabled"))
        if kind == "association":
            self.association_tree, self.remove_association_button = tree, remove
        else:
            self.history_tree, self.remove_history_button = tree, remove

    @staticmethod
    def link_sort_key(value, column):
        text = str(value).strip()
        if column == "date":
            for date_format in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d %b %Y", "%d %B %Y"):
                try:
                    return (0, datetime.strptime(text, date_format).toordinal())
                except ValueError:
                    pass
        natural = tuple((0, int(part)) if part.isdigit() else (1, part.casefold()) for part in re.split(r"(\d+)", text))
        return (1, natural)

    def sort_link_tree(self, tree, column):
        if tree._sort_column == column:
            tree._sort_descending = not tree._sort_descending
        else:
            tree._sort_column = column
            tree._sort_descending = False
        self.apply_link_sort(tree)

    def apply_link_sort(self, tree):
        column = tree._sort_column
        if not column:
            return
        populated, empty = [], []
        for item in tree.get_children(""):
            value = tree.set(item, column)
            (populated if str(value).strip() else empty).append((value, item))
        populated.sort(key=lambda pair: self.link_sort_key(pair[0], column), reverse=tree._sort_descending)
        for position, (_value, item) in enumerate(populated + empty):
            tree.move(item, "", position)
        for heading_column, heading in tree._heading_labels.items():
            marker = " ▼" if tree._sort_descending else " ▲"
            tree.heading(heading_column, text=heading + (marker if heading_column == column else ""))

    @property
    def meta_path(self):
        return self.folder / META_NAME

    @property
    def settings_path(self):
        return self.folder / SETTINGS_NAME

    @property
    def duplicates_path(self):
        return self.folder / DUPLICATES_NAME

    @property
    def legacy_groups_path(self):
        return self.folder / LEGACY_GROUPS_NAME

    @property
    def legacy_settings_path(self):
        return self.folder / LEGACY_SETTINGS_NAME

    def select_folder(self):
        selected = filedialog.askdirectory(title="Select document folder")
        if not selected:
            return
        self.folder = Path(selected)
        self.path_var.set(str(self.folder))
        self.refresh_button.configure(state="normal")
        self.manage_button.configure(state="normal")
        self.load_metadata()
        self.load_duplicate_audit()
        self.load_definitions()
        self.load_files()

    @staticmethod
    def parse_definitions(values):
        names, descriptions = [], {}
        seen = set()
        for item in values if isinstance(values, list) else []:
            if isinstance(item, dict):
                name = str(item.get("name", "")).strip()
                description = str(item.get("description", "")).strip()
            else:
                name, description = str(item).strip(), ""
            if name and name.casefold() not in seen:
                seen.add(name.casefold())
                names.append(name)
                descriptions[name.casefold()] = description
        return sorted(names, key=str.casefold), descriptions

    def load_definitions(self):
        inferred = {}
        for metadata in self.data.values():
            for group in self.decode_groups(metadata.get("Groups", "")):
                inferred.setdefault(group.casefold(), group)
        payload = None
        source_path = None
        try:
            source_path = next(
                (path for path in (self.settings_path, self.legacy_settings_path, self.legacy_groups_path) if path.exists()),
                None,
            )
            if source_path is not None:
                with source_path.open("r", encoding="utf-8") as stream:
                    payload = json.load(stream)
        except (OSError, json.JSONDecodeError, TypeError) as exc:
            messagebox.showerror(APP_NAME, f"Could not read folder settings:\n\n{exc}")
        if isinstance(payload, dict) and "groups" in payload:
            self.defined_groups, self.group_descriptions = self.parse_definitions(payload.get("groups"))
        elif isinstance(payload, list):
            self.defined_groups, self.group_descriptions = self.parse_definitions(payload)
        else:
            self.defined_groups = sorted(inferred.values(), key=str.casefold)
            self.group_descriptions = {}
        if isinstance(payload, dict) and "statuses" in payload:
            self.defined_statuses, self.status_descriptions = self.parse_definitions(payload.get("statuses"))
        else:
            self.defined_statuses = list(DEFAULT_STATUS_OPTIONS)
            self.status_descriptions = {}
        if not self.defined_statuses:
            self.defined_statuses = ["Current"]
        self.refresh_groups()
        self.refresh_status_options()
        if source_path is not None and source_path != self.settings_path and self.save_definitions():
            try:
                source_path.unlink()
            except OSError as exc:
                messagebox.showwarning(
                    APP_NAME,
                    f"Folder settings were migrated to {SETTINGS_NAME}, but the old file could not be removed:\n\n{exc}",
                )

    def save_definitions(self):
        temp_path = self.settings_path.with_suffix(".efs.tmp")
        payload = {
            "groups": [
                {"name": value, "description": self.group_descriptions.get(value.casefold(), "")}
                for value in sorted(self.defined_groups, key=str.casefold)
            ],
            "statuses": [
                {"name": value, "description": self.status_descriptions.get(value.casefold(), "")}
                for value in sorted(self.defined_statuses, key=str.casefold)
            ],
        }
        try:
            with temp_path.open("w", encoding="utf-8") as stream:
                json.dump(payload, stream, ensure_ascii=False, indent=2)
            os.replace(temp_path, self.settings_path)
        except OSError as exc:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass
            messagebox.showerror(APP_NAME, f"Could not save {SETTINGS_NAME}:\n\n{exc}")
            return False
        return True

    def load_metadata(self):
        self.data = {}
        if not self.meta_path.exists():
            return
        try:
            with self.meta_path.open("r", encoding="utf-8-sig", newline="") as stream:
                for row in csv.DictReader(stream):
                    name = row.get("FileName", "")
                    if name:
                        self.data[name] = row
        except (OSError, csv.Error) as exc:
            messagebox.showerror(APP_NAME, f"Could not read {META_NAME}:\n\n{exc}")

    def load_duplicate_audit(self):
        self.duplicate_audit = {"format": "EFM duplicate audit", "version": 1, "records": []}
        if not self.duplicates_path.exists():
            self.save_duplicate_audit()
            return
        try:
            with self.duplicates_path.open("r", encoding="utf-8") as stream:
                payload = json.load(stream)
            if not isinstance(payload, dict) or not isinstance(payload.get("records"), list):
                raise ValueError("The audit file does not contain a valid records list.")
            self.duplicate_audit = payload
        except (OSError, json.JSONDecodeError, TypeError, ValueError) as exc:
            messagebox.showerror(APP_NAME, f"Could not read {DUPLICATES_NAME}:\n\n{exc}")

    def save_duplicate_audit(self):
        temp_path = self.duplicates_path.with_suffix(".efs.tmp")
        try:
            with temp_path.open("w", encoding="utf-8") as stream:
                json.dump(self.duplicate_audit, stream, ensure_ascii=False, indent=2)
            os.replace(temp_path, self.duplicates_path)
        except OSError as exc:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass
            messagebox.showerror(APP_NAME, f"Could not save {DUPLICATES_NAME}:\n\n{exc}")
            return False
        return True

    @staticmethod
    def file_digest(path):
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            while chunk := stream.read(1024 * 1024):
                digest.update(chunk)
        return digest.hexdigest()

    @staticmethod
    def extract_comparison_text(path):
        suffix = path.suffix.casefold()
        try:
            if suffix == ".docx":
                with zipfile.ZipFile(path) as archive:
                    xml = archive.read("word/document.xml")
                text = " ".join(ElementTree.fromstring(xml).itertext())
            elif suffix in {".txt", ".csv", ".md", ".rtf", ".xml", ".html", ".htm"}:
                text = path.read_text(encoding="utf-8", errors="ignore")
            elif suffix == ".pdf":
                with path.open("rb") as stream:
                    raw = stream.read(4 * 1024 * 1024)
                text = " ".join(re.findall(r"[A-Za-z0-9][A-Za-z0-9 ,.;:'()/_-]{5,}", raw.decode("latin-1", errors="ignore")))
            else:
                return ""
        except (OSError, KeyError, zipfile.BadZipFile, ElementTree.ParseError):
            return ""
        return re.sub(r"\s+", " ", text).strip().casefold()[:50000]

    def duplicate_fingerprint(self, path):
        stat = path.stat()
        text = self.extract_comparison_text(path)
        metadata = self.data.get(path.name, {})
        return {
            "name": path.name,
            "path": path,
            "size": stat.st_size,
            "hash": self.file_digest(path),
            "text": text,
            "text_hash": hashlib.sha256(text.encode("utf-8")).hexdigest() if text else "",
            "title": metadata.get("DocumentTitle", "").strip(),
            "number": metadata.get("Number", "").strip(),
            "version": metadata.get("Version", "").strip(),
        }

    @staticmethod
    def duplicate_name_key(name):
        tokens = re.findall(r"[a-z0-9]+", Path(name).stem.casefold())
        ignored = {"copy", "duplicate", "duplicated", "final", "document", "file"}
        return " ".join(token for token in tokens if token not in ignored)

    @staticmethod
    def content_simhash(text):
        if not text:
            return None
        tokens = re.findall(r"[a-z0-9]{3,}", text[:30000])
        if not tokens:
            return None
        weights = [0] * 64
        for token in set(tokens[:5000]):
            value = int.from_bytes(hashlib.blake2b(token.encode("utf-8"), digest_size=8).digest(), "big")
            for bit in range(64):
                weights[bit] += 1 if value & (1 << bit) else -1
        result = 0
        for bit, weight in enumerate(weights):
            if weight >= 0:
                result |= 1 << bit
        return result

    @staticmethod
    def duplicate_score(first, second):
        if first["hash"] == second["hash"]:
            return 100, ["identical SHA-256 hash"]
        reasons = []
        score = 0.0
        name_similarity = SequenceMatcher(None, Path(first["name"]).stem.casefold(), Path(second["name"]).stem.casefold()).ratio()
        if name_similarity >= 0.45:
            score += name_similarity * 12
            reasons.append(f"filename {name_similarity:.0%}")
        size_ratio = min(first["size"], second["size"]) / max(first["size"], second["size"], 1)
        if size_ratio >= 0.65:
            score += size_ratio * 10
            reasons.append(f"size {size_ratio:.0%}")
        if first["number"] and first["number"].casefold() == second["number"].casefold():
            score += 22
            reasons.append("document number matches")
        if first["title"] and second["title"]:
            title_similarity = SequenceMatcher(None, first["title"].casefold(), second["title"].casefold()).ratio()
            if title_similarity >= 0.55:
                score += title_similarity * 20
                reasons.append(f"title {title_similarity:.0%}")
        if first["text_hash"] and first["text_hash"] == second["text_hash"]:
            score = max(score, 94)
            reasons.insert(0, "normalized document content matches")
        elif first["text"] and second["text"] and (name_similarity >= 0.35 or score >= 20):
            text_similarity = SequenceMatcher(None, first["text"][:20000], second["text"][:20000]).ratio()
            if text_similarity >= 0.45:
                score += text_similarity * 52
                reasons.insert(0, f"document content {text_similarity:.0%}")
        if first["version"] and second["version"] and first["version"].casefold() != second["version"].casefold():
            score -= 15
            reasons.append("versions differ")
        return max(0, min(99, round(score))), reasons

    def find_duplicate_candidates(self, progress=None, cancel_event=None):
        def cancelled():
            return cancel_event is not None and cancel_event.is_set()

        def report(message):
            if progress is not None:
                progress(message)

        fingerprints = []
        total = len(self.files)
        for index, path in enumerate(self.files, start=1):
            if cancelled():
                return None
            report(f"Stage 1 of 3 — fingerprinting {index} of {total}: {path.name}")
            try:
                fingerprint = self.duplicate_fingerprint(path)
                fingerprint["name_key"] = self.duplicate_name_key(path.name)
                fingerprint["simhash"] = self.content_simhash(fingerprint["text"])
                fingerprints.append(fingerprint)
            except OSError:
                continue

        report("Stage 2 of 3 — building focused candidate groups")
        candidate_pairs = set()
        hash_groups = {}
        content_groups = {}
        blocks = {}
        for index, item in enumerate(fingerprints):
            hash_groups.setdefault(item["hash"], []).append(index)
            if item["text_hash"]:
                content_groups.setdefault(item["text_hash"], []).append(index)
            keys = []
            if item["number"]:
                keys.append(("number", item["number"].casefold()))
            if item["title"]:
                title_key = re.sub(r"\W+", "", item["title"].casefold())
                if len(title_key) >= 5:
                    keys.append(("title", title_key))
            if len(item["name_key"]) >= 4:
                keys.append(("name", item["name_key"]))
            size_bucket = round(math.log(max(item["size"], 1), 1.08))
            extension = Path(item["name"]).suffix.casefold()
            for nearby in (size_bucket - 1, size_bucket, size_bucket + 1):
                keys.append(("size", extension, nearby))
            if item["simhash"] is not None:
                for chunk in range(4):
                    keys.append(("content", chunk, (item["simhash"] >> (chunk * 16)) & 0xFFFF))
            for key in keys:
                blocks.setdefault(key, []).append(index)

        # Exact groups need only a star of comparisons, not every combination.
        for groups in (hash_groups, content_groups):
            for members in groups.values():
                if len(members) > 1:
                    candidate_pairs.update((min(members[0], other), max(members[0], other)) for other in members[1:])
        skipped_broad_blocks = 0
        for members in blocks.values():
            unique_members = sorted(set(members))
            if len(unique_members) > 120:
                skipped_broad_blocks += 1
                continue
            for position, first_index in enumerate(unique_members):
                candidate_pairs.update((first_index, second_index) for second_index in unique_members[position + 1:])
            if cancelled():
                return None

        pair_list = sorted(candidate_pairs)
        candidates = []
        total_pairs = len(pair_list)
        for checked, (first_index, second_index) in enumerate(pair_list, start=1):
            if cancelled():
                return None
            if checked == 1 or checked % 100 == 0:
                report(f"Stage 3 of 3 — scoring candidate {checked} of {total_pairs}")
            first, second = fingerprints[first_index], fingerprints[second_index]
            score, reasons = self.duplicate_score(first, second)
            if score >= 55:
                candidates.append({"first": first["name"], "second": second["name"], "confidence": score, "reasons": reasons})
        report(f"Complete — scored {total_pairs} focused candidate pair(s); skipped {skipped_broad_blocks} overly broad group(s)")
        return sorted(candidates, key=lambda item: (-item["confidence"], item["first"].casefold(), item["second"].casefold()))

    def open_document_management(self):
        if not self.folder:
            return
        dialog = tk.Toplevel(self)
        dialog.title("Document Management")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("980x620")
        dialog.minsize(760, 480)
        ttk.Label(dialog, text="Document Management", padding=(16, 16, 16, 2), font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(dialog, text="Find possible duplicates, review the evidence, and explicitly choose which file to retain.", padding=(16, 0, 16, 10)).pack(anchor="w")
        notebook = ttk.Notebook(dialog)
        notebook.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        matches_tab = ttk.Frame(notebook, padding=10)
        audit_tab = ttk.Frame(notebook, padding=10)
        notebook.add(matches_tab, text="Possible Duplicates")
        notebook.add(audit_tab, text="Resolved Duplicates")
        match_tree = ttk.Treeview(matches_tab, columns=("confidence", "first", "second", "evidence"), show="headings", selectmode="browse")
        for column, label, width in (("confidence", "Confidence", 90), ("first", "File A", 230), ("second", "File B", 230), ("evidence", "Evidence", 330)):
            match_tree.heading(column, text=label, anchor="w")
            match_tree.column(column, width=width, minwidth=70)
        match_tree.pack(fill="both", expand=True)
        progress_var = tk.StringVar(value="Choose Find Possible Duplicates to analyse the current Documents list.")
        ttk.Label(matches_tab, textvariable=progress_var).pack(anchor="w", pady=(8, 0))
        match_actions = ttk.Frame(matches_tab)
        match_actions.pack(fill="x", pady=(8, 0))
        scan_button = ttk.Button(match_actions, text="Find Possible Duplicates", style="Primary.TButton")
        scan_button.pack(side="left")
        cancel_scan_button = ttk.Button(match_actions, text="Cancel Analysis", style="Secondary.TButton", state="disabled")
        cancel_scan_button.pack(side="left", padx=(7, 0))
        review_button = ttk.Button(match_actions, text="Review Selected", style="Secondary.TButton", state="disabled")
        review_button.pack(side="left", padx=(7, 0))
        candidates = []
        scan_messages = queue.Queue()
        scan_cancel = [None]
        scan_thread = [None]

        audit_tree = ttk.Treeview(audit_tab, columns=("date", "canonical", "duplicate", "confidence", "status"), show="headings", selectmode="browse")
        for column, label, width in (("date", "Resolved", 150), ("canonical", "Kept File", 220), ("duplicate", "Moved File", 220), ("confidence", "Confidence", 90), ("status", "Status", 90)):
            audit_tree.heading(column, text=label, anchor="w")
            audit_tree.column(column, width=width, minwidth=70)
        audit_tree.pack(fill="both", expand=True)
        audit_actions = ttk.Frame(audit_tab)
        audit_actions.pack(fill="x", pady=(8, 0))
        restore_button = ttk.Button(audit_actions, text="Restore Selected", style="Secondary.TButton", state="disabled")
        restore_button.pack(side="left")

        def refresh_audit():
            audit_tree.delete(*audit_tree.get_children())
            for index, record in enumerate(self.duplicate_audit.get("records", [])):
                audit_tree.insert("", "end", iid=str(index), values=(record.get("resolved_at", ""), record.get("canonical", ""), record.get("duplicate", ""), f"{record.get('confidence', 0)}%", "Resolved" if record.get("active", True) else "Restored"))
            restore_button.configure(state="disabled")

        def run_scan():
            if scan_thread[0] is not None and scan_thread[0].is_alive():
                return
            scan_button.configure(state="disabled")
            cancel_scan_button.configure(state="normal")
            review_button.configure(state="disabled")
            match_tree.delete(*match_tree.get_children())
            progress_var.set("Starting duplicate analysis…")
            scan_cancel[0] = threading.Event()

            def worker():
                try:
                    result = self.find_duplicate_candidates(
                        lambda message: scan_messages.put(("progress", message)),
                        scan_cancel[0],
                    )
                    scan_messages.put(("complete", result))
                except Exception as exc:
                    scan_messages.put(("error", str(exc)))

            scan_thread[0] = threading.Thread(target=worker, name="EFM duplicate analysis", daemon=True)
            scan_thread[0].start()
            dialog.after(100, poll_scan)

        def poll_scan():
            nonlocal candidates
            complete = False
            try:
                while True:
                    message_type, payload = scan_messages.get_nowait()
                    if message_type == "progress":
                        progress_var.set(payload)
                    elif message_type == "complete":
                        complete = True
                        candidates = payload or []
                        if payload is None:
                            progress_var.set("Analysis cancelled. No files were changed.")
                        else:
                            for index, item in enumerate(candidates):
                                match_tree.insert("", "end", iid=str(index), values=(f"{item['confidence']}%", item["first"], item["second"], "; ".join(item["reasons"])))
                            progress_var.set(f"Found {len(candidates)} possible duplicate pair(s). No files have been changed.")
                    elif message_type == "error":
                        complete = True
                        progress_var.set("Analysis stopped because an error occurred.")
                        messagebox.showerror(APP_NAME, f"Duplicate analysis could not be completed:\n\n{payload}", parent=dialog)
            except queue.Empty:
                pass
            if complete:
                scan_button.configure(state="normal")
                cancel_scan_button.configure(state="disabled")
                scan_thread[0] = None
            elif dialog.winfo_exists():
                dialog.after(100, poll_scan)

        def cancel_scan():
            if scan_cancel[0] is not None:
                scan_cancel[0].set()
                cancel_scan_button.configure(state="disabled")
                progress_var.set("Cancelling analysis…")

        def review_selected(_event=None):
            selected = match_tree.selection()
            if selected:
                self.review_duplicate_candidate(dialog, candidates[int(selected[0])], run_scan, refresh_audit)

        def restore_selected():
            selected = audit_tree.selection()
            if selected and self.restore_duplicate_record(int(selected[0]), dialog):
                refresh_audit()
                run_scan()

        scan_button.configure(command=run_scan)
        cancel_scan_button.configure(command=cancel_scan)
        review_button.configure(command=review_selected)
        restore_button.configure(command=restore_selected)
        match_tree.bind("<<TreeviewSelect>>", lambda _event: review_button.configure(state="normal" if match_tree.selection() else "disabled"))
        match_tree.bind("<Double-1>", review_selected)
        audit_tree.bind("<<TreeviewSelect>>", lambda _event: restore_button.configure(state="normal" if audit_tree.selection() and self.duplicate_audit["records"][int(audit_tree.selection()[0])].get("active", True) else "disabled"))
        refresh_audit()
        def close_dialog():
            cancel_scan()
            dialog.destroy()

        ttk.Button(dialog, text="Close", style="Primary.TButton", command=close_dialog).pack(side="right", padx=16, pady=(0, 14))
        dialog.protocol("WM_DELETE_WINDOW", close_dialog)

    def review_duplicate_candidate(self, parent, candidate, on_resolved, on_audit_changed):
        names = (candidate["first"], candidate["second"])
        if any(not (self.folder / name).is_file() for name in names):
            messagebox.showwarning(APP_NAME, "One of these files is no longer available in the selected folder.", parent=parent)
            return
        review = tk.Toplevel(parent)
        review.title("Review Possible Duplicate")
        review.transient(parent)
        review.grab_set()
        review.geometry("820x600")
        review.minsize(680, 500)
        ttk.Label(review, text=f"Duplicate confidence: {candidate['confidence']}%", padding=(16, 16, 16, 2), font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(review, text="Evidence: " + "; ".join(candidate["reasons"]), padding=(16, 0, 16, 10), wraplength=760).pack(anchor="w")
        canonical_var = tk.StringVar(value=names[0])
        choice = ttk.LabelFrame(review, text="Choose the file to keep in the selected folder", padding=10)
        choice.pack(fill="x", padx=16, pady=(0, 10))
        for name in names:
            ttk.Radiobutton(choice, text=name, variable=canonical_var, value=name).pack(anchor="w", pady=2)
        comparison = ttk.Treeview(review, columns=("field", "first", "second"), show="headings")
        comparison.heading("field", text="Metadata Field", anchor="w")
        comparison.heading("first", text=names[0], anchor="w")
        comparison.heading("second", text=names[1], anchor="w")
        comparison.column("field", width=150, minwidth=100)
        comparison.column("first", width=300, minwidth=180)
        comparison.column("second", width=300, minwidth=180)
        comparison.pack(fill="both", expand=True, padx=16)
        fields = (("Document Title", "DocumentTitle"), ("Type", "Type"), ("Number", "Number"), ("Version", "Version"), ("Issue Date", "IssueDate"), ("Status", "Status"), ("Groups", "Groups"))
        for index, (label, key) in enumerate(fields):
            values = []
            for name in names:
                value = self.data.get(name, {}).get(key, "")
                if key == "Groups":
                    value = ", ".join(self.decode_groups(value))
                values.append(value or "—")
            comparison.insert("", "end", iid=str(index), values=(label, *values))
        ttk.Label(
            review,
            text="The other file will be moved to the duplicates folder. Blank canonical metadata will be filled where possible, and document references will be redirected.",
            padding=(16, 10, 16, 8),
            wraplength=760,
        ).pack(anchor="w")
        buttons = ttk.Frame(review, padding=(16, 0, 16, 16))
        buttons.pack(fill="x")

        def confirm():
            canonical = canonical_var.get()
            duplicate = names[1] if canonical == names[0] else names[0]
            if not messagebox.askyesno(
                "Confirm Duplicate Resolution",
                f"Keep:\n{canonical}\n\nMove to duplicates:\n{duplicate}\n\nThis will move the file and update EFM metadata references. Continue?",
                icon="warning",
                parent=review,
            ):
                return
            if self.resolve_duplicate(canonical, duplicate, candidate, review):
                review.destroy()
                on_audit_changed()
                on_resolved()

        ttk.Button(buttons, text="Confirm and Move Duplicate", style="Primary.TButton", command=confirm).pack(side="right")
        ttk.Button(buttons, text="Cancel", style="Secondary.TButton", command=review.destroy).pack(side="right", padx=(0, 7))

    @staticmethod
    def merge_unique(values):
        merged = {}
        for value in values:
            text = str(value).strip()
            if text:
                merged.setdefault(text.casefold(), text)
        return list(merged.values())

    def consolidate_metadata(self, canonical, duplicate):
        canonical_row = dict(self.data.get(canonical, {}))
        duplicate_row = dict(self.data.get(duplicate, {}))
        canonical_row["FileName"] = canonical
        canonical_row.setdefault("Status", self.default_status())
        special = {"FileName", "Groups", "AssociatedDocs", "Attachments", "VersionHistory"}
        for key in METADATA_COLUMNS:
            if key not in special and not canonical_row.get(key) and duplicate_row.get(key):
                canonical_row[key] = duplicate_row[key]
        groups = self.merge_unique(self.decode_groups(canonical_row.get("Groups", "")) + self.decode_groups(duplicate_row.get("Groups", "")))
        canonical_row["Groups"] = json.dumps(groups, ensure_ascii=False)
        for field in ("AssociatedDocs", "VersionHistory"):
            records = self.decode_links(canonical_row.get(field, "")) + self.decode_links(duplicate_row.get(field, ""))
            unique = {}
            for record in records:
                target = record.get("FileName", "")
                if target.casefold() in {canonical.casefold(), duplicate.casefold()}:
                    continue
                key = target.casefold() if target else "legacy:" + record.get("LegacyText", "").casefold()
                if key:
                    unique.setdefault(key, record)
            canonical_row[field] = json.dumps(list(unique.values()), ensure_ascii=False)
        attachments = self.merge_unique(self.decode_attachments(canonical_row.get("Attachments", "")) + self.decode_attachments(duplicate_row.get("Attachments", "")))
        canonical_row["Attachments"] = json.dumps([name for name in attachments if name.casefold() not in {canonical.casefold(), duplicate.casefold()}], ensure_ascii=False)
        return canonical_row

    def redirect_duplicate_references(self, canonical, duplicate):
        for source, row in list(self.data.items()):
            if source == duplicate:
                continue
            updated = dict(row)
            for field in ("AssociatedDocs", "VersionHistory"):
                unique = {}
                for record in self.decode_links(updated.get(field, "")):
                    item = dict(record)
                    target = item.get("FileName", "")
                    if target.casefold() == duplicate.casefold():
                        item["FileName"] = canonical
                        target = canonical
                    if target.casefold() == source.casefold():
                        continue
                    key = target.casefold() if target else "legacy:" + item.get("LegacyText", "").casefold()
                    if key:
                        unique.setdefault(key, item)
                updated[field] = json.dumps(list(unique.values()), ensure_ascii=False)
            attachments = [canonical if name.casefold() == duplicate.casefold() else name for name in self.decode_attachments(updated.get("Attachments", ""))]
            updated["Attachments"] = json.dumps(self.merge_unique(name for name in attachments if name.casefold() != source.casefold()), ensure_ascii=False)
            self.data[source] = updated

    def resolve_duplicate(self, canonical, duplicate, candidate, parent):
        source = (self.folder / duplicate).resolve()
        root = self.folder.resolve()
        duplicate_dir = (root / "duplicates").resolve()
        destination = (duplicate_dir / duplicate).resolve()
        if source.parent != root or destination.parent != duplicate_dir or not source.is_file():
            messagebox.showerror(APP_NAME, "The duplicate file path is not valid for this selected folder.", parent=parent)
            return False
        if destination.exists():
            messagebox.showerror(APP_NAME, f"A file with this name already exists in the duplicates folder:\n\n{duplicate}", parent=parent)
            return False
        original_data = {name: dict(row) for name, row in self.data.items()}
        original_records = list(self.duplicate_audit.get("records", []))
        try:
            duplicate_dir.mkdir(exist_ok=True)
            shutil.move(str(source), str(destination))
            self.data[canonical] = self.consolidate_metadata(canonical, duplicate)
            self.data.pop(duplicate, None)
            self.redirect_duplicate_references(canonical, duplicate)
            changed_before = {
                name: original_data.get(name)
                for name in set(original_data) | set(self.data)
                if original_data.get(name) != self.data.get(name)
            }
            if not self.write_metadata():
                raise OSError("The metadata update could not be saved.")
            record = {
                "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                "resolved_at": datetime.now().isoformat(timespec="seconds"),
                "canonical": canonical,
                "duplicate": duplicate,
                "moved_to": str(Path("duplicates") / duplicate),
                "confidence": candidate.get("confidence", 0),
                "evidence": candidate.get("reasons", []),
                "affected_metadata_before": changed_before,
                "active": True,
            }
            self.duplicate_audit.setdefault("records", []).append(record)
            if not self.save_duplicate_audit():
                raise OSError("The duplicate audit record could not be saved.")
        except OSError as exc:
            self.data = original_data
            self.duplicate_audit["records"] = original_records
            self.write_metadata()
            try:
                if destination.exists() and not source.exists():
                    shutil.move(str(destination), str(source))
            except OSError:
                pass
            messagebox.showerror(APP_NAME, f"The duplicate could not be resolved safely:\n\n{exc}", parent=parent)
            return False
        self.load_files()
        return True

    def restore_duplicate_record(self, index, parent):
        records = self.duplicate_audit.get("records", [])
        if index < 0 or index >= len(records) or not records[index].get("active", True):
            return False
        record = records[index]
        source = (self.folder / record.get("moved_to", "")).resolve()
        destination = (self.folder / record.get("duplicate", "")).resolve()
        root = self.folder.resolve()
        if source.parent != (root / "duplicates").resolve() or destination.parent != root or not source.is_file():
            messagebox.showerror(APP_NAME, "The archived duplicate file is not available at its recorded location.", parent=parent)
            return False
        if destination.exists():
            messagebox.showerror(APP_NAME, f"The file cannot be restored because its original name is already in use:\n\n{destination.name}", parent=parent)
            return False
        if not messagebox.askyesno("Restore Duplicate", f"Restore {destination.name} to the selected folder and reinstate the affected metadata records?", icon="warning", parent=parent):
            return False
        original_data = {name: dict(row) for name, row in self.data.items()}
        original_record = dict(record)
        try:
            shutil.move(str(source), str(destination))
            for name, prior in record.get("affected_metadata_before", {}).items():
                if prior is None:
                    self.data.pop(name, None)
                else:
                    self.data[name] = prior
            if not self.write_metadata():
                raise OSError("The restored metadata could not be saved.")
            record["active"] = False
            record["restored_at"] = datetime.now().isoformat(timespec="seconds")
            if not self.save_duplicate_audit():
                raise OSError("The restoration audit could not be saved.")
        except OSError as exc:
            self.data = original_data
            records[index] = original_record
            self.write_metadata()
            try:
                if destination.exists() and not source.exists():
                    shutil.move(str(destination), str(source))
            except OSError:
                pass
            messagebox.showerror(APP_NAME, f"The duplicate could not be restored safely:\n\n{exc}", parent=parent)
            return False
        self.load_files()
        return True

    def load_files(self):
        if not self.folder:
            return
        if not self.scan_folder_files():
            return
        self.show_reconciliation_summary()
        self.selected_name = None
        self.navigation_history.clear()
        self.back_button.configure(state="disabled")
        self.apply_filter()
        self.clear_details()

    def all_attachment_names(self):
        names = set()
        for metadata in self.data.values():
            names.update(name.casefold() for name in self.decode_attachments(metadata.get("Attachments", "")))
        return names

    def scan_folder_files(self):
        try:
            managed_names = {
                META_NAME.casefold(),
                SETTINGS_NAME.casefold(),
                DUPLICATES_NAME.casefold(),
                LEGACY_SETTINGS_NAME.casefold(),
                LEGACY_GROUPS_NAME.casefold(),
            }
            self.folder_files = sorted(
                (p for p in self.folder.iterdir() if p.is_file() and p.name.casefold() not in managed_names),
                key=lambda p: p.name.casefold(),
            )
            attachment_names = self.all_attachment_names()
            self.files = [path for path in self.folder_files if path.name.casefold() not in attachment_names]
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Could not read the selected folder:\n\n{exc}")
            return False
        return True

    def show_reconciliation_summary(self):
        def discrepancies():
            folder_names = {path.name.casefold(): path.name for path in self.files}
            attachment_names = self.all_attachment_names()
            metadata_names = {name.casefold(): name for name in self.data if name.casefold() not in attachment_names}
            files_without_metadata = sorted(
                (actual for key, actual in folder_names.items() if key not in metadata_names),
                key=str.casefold,
            )
            metadata_without_files = sorted(
                (actual for key, actual in metadata_names.items() if key not in folder_names),
                key=str.casefold,
            )
            return files_without_metadata, metadata_without_files

        files_without_metadata, metadata_without_files = discrepancies()
        if not files_without_metadata and not metadata_without_files:
            return

        dialog = tk.Toplevel(self)
        dialog.title("File Reconciliation Review")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("820x520")
        dialog.minsize(640, 400)
        ttk.Label(dialog, text="File reconciliation review", padding=(16, 16, 16, 2), font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(
            dialog,
            text="No changes have been made. Review the discrepancies between the selected folder and filedata.efm.",
            padding=(16, 0, 16, 12),
        ).pack(anchor="w")

        content = ttk.Frame(dialog, padding=(16, 0, 16, 10))
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=1)
        content.columnconfigure(2, weight=1)
        content.rowconfigure(1, weight=1)

        files_label = ttk.Label(content, font=("Segoe UI", 10, "bold"))
        files_label.grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 6))
        metadata_label = ttk.Label(content, font=("Segoe UI", 10, "bold"))
        metadata_label.grid(row=0, column=2, sticky="w", padx=(8, 0), pady=(0, 6))

        file_tree = ttk.Treeview(content, columns=("name",), show="headings", selectmode="browse")
        file_tree.heading("name", text="Document Name", anchor="w")
        file_tree.column("name", width=330, minwidth=180)
        file_tree.grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        file_scroll = ttk.Scrollbar(content, orient="vertical", command=file_tree.yview)
        file_scroll.grid(row=1, column=1, sticky="ns")
        file_tree.configure(yscrollcommand=file_scroll.set)
        metadata_tree = ttk.Treeview(content, columns=("name",), show="headings", selectmode="browse")
        metadata_tree.heading("name", text="Document Name", anchor="w")
        metadata_tree.column("name", width=330, minwidth=180)
        metadata_tree.grid(row=1, column=2, sticky="nsew", padx=(8, 8))
        metadata_scroll = ttk.Scrollbar(content, orient="vertical", command=metadata_tree.yview)
        metadata_scroll.grid(row=1, column=3, sticky="ns")
        metadata_tree.configure(yscrollcommand=metadata_scroll.set)
        actions = ttk.Frame(dialog, padding=(16, 0, 16, 16))
        actions.pack(fill="x")
        review_button = ttk.Button(actions, text="Review Selected File", style="Primary.TButton", state="disabled")
        review_button.pack(side="left")
        ttk.Button(actions, text="Close Review", style="Primary.TButton", command=dialog.destroy).pack(side="right")

        def refresh_review():
            current_files, current_metadata = discrepancies()
            file_tree.delete(*file_tree.get_children())
            metadata_tree.delete(*metadata_tree.get_children())
            for index, name in enumerate(current_files):
                file_tree.insert("", "end", iid=str(index), values=(name,))
            for index, name in enumerate(current_metadata):
                metadata_tree.insert("", "end", iid=str(index), values=(name,))
            files_label.configure(text=f"Files without Metadata ({len(current_files)})")
            metadata_label.configure(text=f"Metadata without a File ({len(current_metadata)})")
            review_button.configure(state="disabled")

        def selected_file_name():
            selected = file_tree.selection()
            return file_tree.item(selected[0], "values")[0] if selected else None

        def review_selected(_event=None):
            file_name = selected_file_name()
            _files, available_metadata = discrepancies()
            if file_name and available_metadata:
                self.review_metadata_proposal(dialog, file_name, available_metadata, refresh_review)
            elif file_name:
                messagebox.showinfo(APP_NAME, "There are no unmatched metadata records available to assign.", parent=dialog)

        def update_review_button(_event=None):
            _files, available_metadata = discrepancies()
            review_button.configure(state="normal" if selected_file_name() and available_metadata else "disabled")

        review_button.configure(command=review_selected)
        file_tree.bind("<<TreeviewSelect>>", update_review_button)
        file_tree.bind("<Double-1>", review_selected)
        refresh_review()
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        self.wait_window(dialog)

    def review_metadata_proposal(self, parent, file_name, metadata_names, on_applied):
        proposal = tk.Toplevel(parent)
        proposal.title("Assign Metadata to File")
        proposal.transient(parent)
        proposal.grab_set()
        proposal.geometry("650x540")
        proposal.minsize(560, 470)
        ttk.Label(proposal, text="Assign metadata to file", padding=(16, 16, 16, 2), font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(proposal, text=f"File: {file_name}", padding=(16, 0, 16, 10), font=("Segoe UI", 10, "bold")).pack(anchor="w")

        normalized_file = Path(file_name).stem.casefold()
        recommended = max(
            metadata_names,
            key=lambda name: SequenceMatcher(None, normalized_file, Path(name).stem.casefold()).ratio(),
        )
        choice_row = ttk.Frame(proposal, padding=(16, 0, 16, 12))
        choice_row.pack(fill="x")
        ttk.Label(choice_row, text="Proposed metadata:").pack(side="left", padx=(0, 8))
        metadata_var = tk.StringVar(value=recommended)
        selector = ttk.Combobox(choice_row, textvariable=metadata_var, state="readonly", values=tuple(metadata_names))
        selector.pack(side="left", fill="x", expand=True)

        ttk.Label(proposal, text="Metadata preview", padding=(16, 0, 16, 6), font=("Segoe UI", 10, "bold")).pack(anchor="w")
        preview = ttk.Frame(proposal, padding=(16, 0, 16, 8))
        preview.pack(fill="both", expand=True)
        preview.columnconfigure(1, weight=1)
        preview_fields = [
            ("Document Title", "DocumentTitle"),
            ("Type", "Type"),
            ("Number", "Number"),
            ("Issue date", "IssueDate"),
            ("Sheet number", "SheetNo"),
            ("Total sheets", "TotalSheets"),
            ("Version", "Version"),
            ("Status", "Status"),
        ]
        preview_vars = {}
        for row_index, (label, key) in enumerate(preview_fields):
            ttk.Label(preview, text=f"{label}:", width=16).grid(row=row_index, column=0, sticky="w", pady=4)
            variable = tk.StringVar()
            ttk.Label(preview, textvariable=variable).grid(row=row_index, column=1, sticky="w", pady=4)
            preview_vars[key] = variable

        def update_preview(_event=None):
            row = self.data.get(metadata_var.get(), {})
            for key, variable in preview_vars.items():
                variable.set(row.get(key, "") or (self.default_status() if key == "Status" else "—"))

        def apply_proposal():
            old_name = metadata_var.get()
            confirmed = messagebox.askyesno(
                "Confirm Metadata Assignment",
                f"Assign the metadata currently stored for:\n\n{old_name}\n\nto:\n\n{file_name}\n\nAssociation and version-history references will also be updated.",
                parent=proposal,
            )
            if confirmed and self.assign_metadata_to_file(old_name, file_name):
                proposal.destroy()
                on_applied()

        selector.bind("<<ComboboxSelected>>", update_preview)
        update_preview()
        buttons = ttk.Frame(proposal, padding=(16, 8, 16, 16))
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Assign Metadata", style="Primary.TButton", command=apply_proposal).pack(side="right")
        ttk.Button(buttons, text="Cancel", style="Secondary.TButton", command=proposal.destroy).pack(side="right", padx=(0, 7))
        proposal.protocol("WM_DELETE_WINDOW", proposal.destroy)
        self.wait_window(proposal)

    def assign_metadata_to_file(self, old_name, new_name):
        if old_name not in self.data or new_name in self.data:
            return False
        original_data = {name: dict(row) for name, row in self.data.items()}
        row = self.data.pop(old_name)
        row["FileName"] = new_name
        self.data[new_name] = row
        for metadata in self.data.values():
            for field in ("AssociatedDocs", "VersionHistory"):
                records = self.decode_links(metadata.get(field, ""))
                changed = False
                for record in records:
                    if record.get("FileName") == old_name:
                        record["FileName"] = new_name
                        changed = True
                if changed:
                    metadata[field] = json.dumps(records, ensure_ascii=False)
            attachments = self.decode_attachments(metadata.get("Attachments", ""))
            if any(name.casefold() == old_name.casefold() for name in attachments):
                metadata["Attachments"] = json.dumps(
                    [new_name if name.casefold() == old_name.casefold() else name for name in attachments],
                    ensure_ascii=False,
                )
        if not self.write_metadata():
            self.data = original_data
            return False
        return True

    def apply_filter(self):
        term = self.search_var.get().strip().casefold()
        selected_status = self.status_filter_var.get()
        self.filtered_files = [
            p for p in self.files
            if (term in p.name.casefold() or term in self.data.get(p.name, {}).get("DocumentTitle", "").casefold())
            and (
                selected_status == "All statuses"
                or (self.data.get(p.name, {}).get("Status") or self.default_status()) == selected_status
            )
        ]
        self.tree.delete(*self.tree.get_children())
        for path in self.filtered_files:
            metadata = self.data.get(path.name, {})
            title = metadata.get("DocumentTitle", "")
            status = metadata.get("Status") or self.default_status()
            self.tree.insert("", "end", iid=path.name, text=path.name, values=(title, path.suffix[1:].upper() or "FILE", status))
        self.summary_var.set(f"{len(self.filtered_files)} of {len(self.files)} document{'s' if len(self.files) != 1 else ''}")

    def on_select(self, _event=None):
        selection = self.tree.selection()
        if not selection:
            return
        new_name = selection[0]
        if new_name != self.selected_name:
            self.details_notebook.select(self.metadata_tab)
        self.selected_name = new_name
        path = self.folder / self.selected_name
        try:
            stat = path.stat()
            self.file_info_vars["Name"].set(path.name)
            self.file_info_vars["Filetype"].set(path.suffix[1:].upper() or "File")
            self.file_info_vars["Size"].set(human_size(stat.st_size))
        except OSError:
            self.file_info_vars["Name"].set(path.name)
            self.file_info_vars["Filetype"].set("—")
            self.file_info_vars["Size"].set("—")
        self.load_fields(path.name)
        self.open_button.configure(state="normal")
        for button in self.save_buttons:
            button.configure(state="normal")

    def load_fields(self, name):
        row = self.data.get(name)
        if row:
            for key, var in self.fields.items():
                var.set(row.get(key, "") or (self.default_status() if key == "Status" else ""))
            self.groups = self.decode_groups(row.get("Groups", ""))
            self.associations = self.decode_links(row.get("AssociatedDocs", ""))
            self.attachments = self.decode_attachments(row.get("Attachments", ""))
            self.version_history = self.decode_links(row.get("VersionHistory", ""))
        else:
            match = re.match(r"^([A-Za-z]+)([0-9]{2,4})([A-Za-z]{1,2})?", name)
            values = {"DocumentTitle": "", "Type": "", "Number": "", "IssueDate": "", "SheetNo": "", "TotalSheets": "", "Version": "", "Status": self.default_status()}
            if match:
                values.update(Type=match.group(1), Number=match.group(2), Version=match.group(3) or "")
            for key, var in self.fields.items():
                var.set(values[key])
            self.groups, self.associations, self.attachments, self.version_history = [], [], [], []
        self.refresh_status_options()
        self.refresh_groups()
        self.refresh_links()
        self.refresh_attachments()

    def clear_details(self):
        for var in self.file_info_vars.values():
            var.set("—")
        for key, var in self.fields.items():
            var.set(self.default_status() if key == "Status" else "")
        self.groups, self.associations, self.attachments, self.version_history = [], [], [], []
        self.refresh_status_options()
        self.refresh_groups()
        self.refresh_links()
        self.refresh_attachments()
        self.open_button.configure(state="disabled")
        for button in self.save_buttons:
            button.configure(state="disabled")

    @staticmethod
    def decode_links(value):
        """Read current JSON links and retain legacy free-text associations."""
        if not value:
            return []
        try:
            parsed = json.loads(value)
            if isinstance(parsed, list):
                return [item for item in parsed if isinstance(item, dict) and (item.get("FileName") or item.get("LegacyText"))]
        except (TypeError, json.JSONDecodeError):
            pass
        return [{"FileName": "", "LegacyText": str(value)}]

    @staticmethod
    def decode_groups(value):
        if not value:
            return []
        try:
            parsed = json.loads(value)
            if isinstance(parsed, list):
                values = [str(item).strip() for item in parsed if str(item).strip()]
            else:
                values = [str(parsed).strip()]
        except (TypeError, json.JSONDecodeError):
            values = [item.strip() for item in str(value).split(";") if item.strip()]
        unique = {}
        for group in values:
            unique.setdefault(group.casefold(), group)
        return list(unique.values())

    @staticmethod
    def decode_attachments(value):
        if not value:
            return []
        try:
            parsed = json.loads(value)
            values = parsed if isinstance(parsed, list) else [parsed]
        except (TypeError, json.JSONDecodeError):
            values = str(value).split(";")
        unique = {}
        for item in values:
            name = str(item).strip()
            if name:
                unique.setdefault(name.casefold(), name)
        return list(unique.values())

    def default_status(self):
        return self.defined_statuses[0] if self.defined_statuses else "Current"

    def refresh_status_options(self):
        current = self.fields["Status"].get().strip() if hasattr(self, "fields") and "Status" in self.fields else ""
        assignment_values = list(self.defined_statuses)
        if current and not any(value.casefold() == current.casefold() for value in assignment_values):
            assignment_values.append(current)
        if hasattr(self, "status_entry"):
            self.status_entry.configure(values=tuple(assignment_values))

        filter_values = list(self.defined_statuses)
        for metadata in self.data.values():
            value = (metadata.get("Status") or self.default_status()).strip()
            if value and not any(existing.casefold() == value.casefold() for existing in filter_values):
                filter_values.append(value)
        if hasattr(self, "status_filter_entry"):
            available = ("All statuses", *filter_values)
            self.status_filter_entry.configure(values=available)
            if self.status_filter_var.get() not in available:
                self.status_filter_var.set("All statuses")

    def refresh_groups(self):
        self.groups_tree.delete(*self.groups_tree.get_children())
        for index, group in enumerate(sorted(self.groups, key=str.casefold)):
            self.groups_tree.insert(
                "",
                "end",
                iid=str(index),
                values=(group, self.group_descriptions.get(group.casefold(), "")),
            )
        assigned = {group.casefold() for group in self.groups}
        available = sorted((group for group in self.defined_groups if group.casefold() not in assigned), key=str.casefold)
        self.group_var.set("")
        self.group_entry.configure(values=tuple(available))
        self.add_group_button.configure(state="disabled")
        self.remove_group_button.configure(state="disabled")

    def _on_group_choice(self, _event=None):
        self.add_group_button.configure(state="normal" if self.selected_name and self.group_var.get() else "disabled")

    def _on_group_select(self, _event=None):
        self.remove_group_button.configure(state="normal" if self.groups_tree.selection() else "disabled")

    def add_group(self):
        if not self.selected_name:
            return
        group = self.group_var.get().strip()
        if (
            not group
            or not any(defined.casefold() == group.casefold() for defined in self.defined_groups)
            or any(existing.casefold() == group.casefold() for existing in self.groups)
        ):
            return
        self.groups.append(group)
        self.group_var.set("")
        self.refresh_groups()
        self.save_current()

    def remove_group(self):
        selected = self.groups_tree.selection()
        if not selected:
            return
        group = self.groups_tree.item(selected[0], "values")[0]
        self.groups = [existing for existing in self.groups if existing != group]
        self.refresh_groups()
        self.save_current()

    def refresh_attachments(self):
        self.attachment_tree.delete(*self.attachment_tree.get_children())
        for index, name in enumerate(sorted(self.attachments, key=str.casefold)):
            filetype = Path(name).suffix[1:].upper() or "FILE"
            self.attachment_tree.insert("", "end", iid=str(index), values=(name, filetype))
        self.remove_attachment_button.configure(state="disabled")
        self.open_attachment_button.configure(state="disabled")

    def _on_attachment_select(self, _event=None):
        selected = self.attachment_tree.selection()
        self.remove_attachment_button.configure(state="normal" if selected else "disabled")
        self.open_attachment_button.configure(state="normal" if len(selected) == 1 else "disabled")

    def add_attachments(self):
        if not self.selected_name:
            return
        attached_elsewhere = self.all_attachment_names()
        current_names = {name.casefold() for name in self.attachments}
        candidates = [
            path for path in self.folder_files
            if path.name != self.selected_name
            and path.name.casefold() not in attached_elsewhere
            and path.name.casefold() not in current_names
        ]
        if not candidates:
            messagebox.showinfo(APP_NAME, "There are no available files in the selected folder to attach.")
            return
        dialog = tk.Toplevel(self)
        dialog.title("Add Attachments")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("680x440")
        dialog.minsize(520, 320)
        ttk.Label(dialog, text="Select one or more attachment files", padding=(14, 14, 14, 2), font=("Segoe UI", 11, "bold")).pack(anchor="w")
        ttk.Label(
            dialog,
            text="Use Ctrl or Shift for multiple files. Added files will be hidden from the Documents list.",
            padding=(14, 0, 14, 8),
        ).pack(anchor="w")
        table = ttk.Treeview(dialog, columns=("type", "title"), show="tree headings", selectmode="extended")
        table.heading("#0", text="File Name", anchor="w")
        table.heading("type", text="Filetype", anchor="w")
        table.heading("title", text="Existing Document Title", anchor="w")
        table.column("#0", width=300, minwidth=180)
        table.column("type", width=90, minwidth=65, stretch=False)
        table.column("title", width=250, minwidth=150)
        table.pack(fill="both", expand=True, padx=14, pady=(0, 10))
        for path in candidates:
            table.insert(
                "",
                "end",
                iid=path.name,
                text=path.name,
                values=(path.suffix[1:].upper() or "FILE", self.data.get(path.name, {}).get("DocumentTitle", "")),
            )
        buttons = ttk.Frame(dialog, padding=(14, 0, 14, 14))
        buttons.pack(fill="x")
        choose = ttk.Button(buttons, text="Add Selected", style="Primary.TButton", state="disabled")
        choose.pack(side="right")
        ttk.Button(buttons, text="Cancel", style="Secondary.TButton", command=dialog.destroy).pack(side="right", padx=(0, 7))

        def add_selected(_event=None):
            selected = table.selection()
            if not selected:
                return
            self.attachments.extend(selected)
            self.refresh_attachments()
            self.save_current()
            dialog.destroy()

        choose.configure(command=add_selected)
        table.bind("<<TreeviewSelect>>", lambda _event: choose.configure(state="normal" if table.selection() else "disabled"))
        table.bind("<Double-1>", add_selected)

    def remove_attachments(self):
        selected = self.attachment_tree.selection()
        if not selected:
            return
        removed = {self.attachment_tree.item(item, "values")[0].casefold() for item in selected}
        self.attachments = [name for name in self.attachments if name.casefold() not in removed]
        self.refresh_attachments()
        self.save_current()

    def open_attachment(self):
        selected = self.attachment_tree.selection()
        if len(selected) != 1:
            return
        name = self.attachment_tree.item(selected[0], "values")[0]
        path = self.folder / name
        if not path.exists():
            messagebox.showwarning(APP_NAME, f"The attachment is no longer in the selected folder:\n\n{name}")
            return
        try:
            os.startfile(path)
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Could not open the attachment:\n\n{exc}")

    def update_reciprocal_association(self, target_name, add):
        """Add or remove the current document in another document's associations."""
        row = dict(self.data.get(target_name, {}))
        row.setdefault("FileName", target_name)
        row.setdefault("Status", self.default_status())
        records = self.decode_links(row.get("AssociatedDocs", ""))
        if add:
            if not any(record.get("FileName") == self.selected_name for record in records):
                records.append({"FileName": self.selected_name})
        else:
            records = [record for record in records if record.get("FileName") != self.selected_name]
        row["AssociatedDocs"] = json.dumps(records, ensure_ascii=False)
        self.data[target_name] = row

    def refresh_links(self):
        for tree, records in ((self.association_tree, self.associations), (self.history_tree, self.version_history)):
            tree.delete(*tree.get_children())
            for index, record in enumerate(records):
                name = record.get("FileName", "")
                if not name:
                    tree.insert("", "end", iid=str(index), values=("", record.get("LegacyText", "Legacy association"), "", ""))
                    continue
                metadata = self.data.get(name, {})
                title = metadata.get("DocumentTitle") or name
                tree.insert("", "end", iid=str(index), values=(name, title, metadata.get("Version", ""), metadata.get("IssueDate", "")))
            self.apply_link_sort(tree)
        self.remove_association_button.configure(state="disabled")
        self.remove_history_button.configure(state="disabled")
        self.refresh_network()

    def refresh_network(self):
        if not hasattr(self, "network_canvas"):
            return
        canvas = self.network_canvas
        canvas.delete("all")
        if not self.selected_name:
            canvas.create_text(20, 20, anchor="nw", text="Select a document to view its association network.", fill="#71808b", font=("Segoe UI", 10))
            return

        direct_associations = {}
        for record in self.associations:
            target = str(record.get("FileName", "")).strip()
            if target and target.casefold() != self.selected_name.casefold():
                direct_associations.setdefault(target.casefold(), target)
        targets = sorted(direct_associations.values(), key=str.casefold)
        connected = [self.selected_name, *targets]
        edges = [(self.selected_name, target) for target in targets]

        width = max(canvas.winfo_width(), 560)
        height = max(canvas.winfo_height(), 360)
        positions = self.layout_network(connected, edges, self.selected_name, width, height)

        for source, target in edges:
            canvas.create_line(*positions[source], *positions[target], fill="#8ba6b5", width=2)

        for index, name in enumerate(connected):
            x, y = positions[name]
            selected = name == self.selected_name
            fill = "#087e8b" if selected else "#ffffff"
            outline = "#066b76" if selected else "#6f8c9d"
            text_fill = "#ffffff" if selected else "#17324d"
            node_tag = f"network_node_{index}"
            canvas.create_rectangle(
                x - 68, y - 24, x + 68, y + 24,
                fill=fill, outline=outline, width=2,
                tags=(node_tag, "network_node"),
            )
            canvas.create_text(
                x, y, text=name, width=126, fill=text_fill,
                font=("Segoe UI", 8, "bold" if selected else "normal"),
                tags=(node_tag, "network_node"),
            )
            canvas.tag_bind(node_tag, "<Button-1>", lambda _event, target=name: self.open_network_document(target))
            canvas.tag_bind(node_tag, "<Enter>", lambda _event: canvas.configure(cursor="hand2"))
            canvas.tag_bind(node_tag, "<Leave>", lambda _event: canvas.configure(cursor=""))

    def open_network_document(self, target):
        if not target or target == self.selected_name:
            return
        if target not in self.tree.get_children():
            self.search_var.set("")
            self.status_filter_var.set("All statuses")
        if target not in self.tree.get_children():
            messagebox.showwarning(
                APP_NAME,
                f"The associated document is not available in the Documents list:\n\n{target}",
            )
            return
        self.navigation_history.append(self.selected_name)
        self.back_button.configure(state="normal")
        self.tree.selection_set(target)
        self.tree.focus(target)
        self.tree.see(target)
        self.on_select()

    @staticmethod
    def layout_network(nodes, edges, selected_name, width, height):
        """Lay out a connected graph while discouraging node and line overlap."""
        centre_x, centre_y = width / 2, height / 2
        if len(nodes) == 1:
            return {nodes[0]: (centre_x, centre_y)}

        positions = {selected_name: [centre_x, centre_y]}
        golden_angle = math.pi * (3 - math.sqrt(5))
        for index, name in enumerate((node for node in nodes if node != selected_name), start=1):
            radius = 85 + 34 * math.sqrt(index)
            angle = index * golden_angle
            positions[name] = [centre_x + radius * math.cos(angle), centre_y + radius * math.sin(angle)]

        area = max((width - 160) * (height - 100), 1)
        ideal = max(95, min(190, math.sqrt(area / len(nodes)) * 0.72))
        margin_x, margin_y = 78, 34
        for iteration in range(90):
            displacement = {name: [0.0, 0.0] for name in nodes}
            for index, first in enumerate(nodes):
                for second in nodes[index + 1:]:
                    dx = positions[first][0] - positions[second][0]
                    dy = positions[first][1] - positions[second][1]
                    distance = max(math.hypot(dx, dy), 0.1)
                    force = ideal * ideal / distance
                    fx, fy = dx / distance * force, dy / distance * force
                    displacement[first][0] += fx
                    displacement[first][1] += fy
                    displacement[second][0] -= fx
                    displacement[second][1] -= fy
            for first, second in edges:
                dx = positions[first][0] - positions[second][0]
                dy = positions[first][1] - positions[second][1]
                distance = max(math.hypot(dx, dy), 0.1)
                force = distance * distance / ideal
                fx, fy = dx / distance * force, dy / distance * force
                displacement[first][0] -= fx
                displacement[first][1] -= fy
                displacement[second][0] += fx
                displacement[second][1] += fy
            for name in nodes:
                gravity = 0.12 if name == selected_name else 0.025
                displacement[name][0] += (centre_x - positions[name][0]) * gravity
                displacement[name][1] += (centre_y - positions[name][1]) * gravity
                distance = max(math.hypot(*displacement[name]), 0.1)
                temperature = max(2, min(width, height) * 0.11 * (1 - iteration / 90))
                positions[name][0] += displacement[name][0] / distance * min(distance, temperature)
                positions[name][1] += displacement[name][1] / distance * min(distance, temperature)
                positions[name][0] = min(max(positions[name][0], margin_x), width - margin_x)
                positions[name][1] = min(max(positions[name][1], margin_y), height - margin_y)

        # Move nodes away from unrelated edge paths so a line cannot imply a false link.
        for _ in range(18):
            for first, second in edges:
                ax, ay = positions[first]
                bx, by = positions[second]
                edge_dx, edge_dy = bx - ax, by - ay
                edge_length_sq = max(edge_dx * edge_dx + edge_dy * edge_dy, 0.1)
                for name in nodes:
                    if name in (first, second):
                        continue
                    px, py = positions[name]
                    ratio = max(0, min(1, ((px - ax) * edge_dx + (py - ay) * edge_dy) / edge_length_sq))
                    closest_x, closest_y = ax + ratio * edge_dx, ay + ratio * edge_dy
                    dx, dy = px - closest_x, py - closest_y
                    distance = math.hypot(dx, dy)
                    if distance < 74:
                        if distance < 0.1:
                            dx, dy, distance = -edge_dy, edge_dx, max(math.hypot(edge_dx, edge_dy), 0.1)
                        push = (74 - distance) * 0.45
                        positions[name][0] += dx / distance * push
                        positions[name][1] += dy / distance * push
                        positions[name][0] = min(max(positions[name][0], margin_x), width - margin_x)
                        positions[name][1] = min(max(positions[name][1], margin_y), height - margin_y)

        min_x = min(point[0] for point in positions.values())
        max_x = max(point[0] for point in positions.values())
        min_y = min(point[1] for point in positions.values())
        max_y = max(point[1] for point in positions.values())
        shift_x = centre_x - (min_x + max_x) / 2
        shift_y = centre_y - (min_y + max_y) / 2
        return {
            name: (
                min(max(point[0] + shift_x, margin_x), width - margin_x),
                min(max(point[1] + shift_y, margin_y), height - margin_y),
            )
            for name, point in positions.items()
        }

    def add_link(self, kind):
        if not self.selected_name:
            return
        dialog = tk.Toplevel(self)
        dialog.title("Add Associated Documents" if kind == "association" else "Add Version History Documents")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("650x420")
        dialog.minsize(500, 320)
        ttk.Label(dialog, text="Select one or more documents to link", padding=(14, 14, 14, 2), font=("Segoe UI", 11, "bold")).pack(anchor="w")
        ttk.Label(dialog, text="Use Ctrl or Shift to select multiple documents.", padding=(14, 0, 14, 8)).pack(anchor="w")
        table = ttk.Treeview(dialog, columns=("title",), show="tree headings", selectmode="extended")
        table.heading("#0", text="File name", anchor="w")
        table.heading("title", text="Document Title", anchor="w")
        table.column("#0", width=300, minwidth=180)
        table.column("title", width=300, minwidth=180)
        table.pack(fill="both", expand=True, padx=14, pady=(0, 10))
        existing = self.associations if kind == "association" else self.version_history
        existing_names = {entry.get("FileName") for entry in existing}
        for path in self.files:
            if path.name != self.selected_name and path.name not in existing_names:
                table.insert("", "end", iid=path.name, text=path.name, values=(self.data.get(path.name, {}).get("DocumentTitle", ""),))
        buttons = ttk.Frame(dialog, padding=(14, 0, 14, 14))
        buttons.pack(fill="x")
        choose = ttk.Button(buttons, text="Add Selected", style="Primary.TButton", state="disabled")
        choose.pack(side="right")
        ttk.Button(buttons, text="Cancel", style="Secondary.TButton", command=dialog.destroy).pack(side="right", padx=(0, 7))

        def add_selected(_event=None):
            selected = table.selection()
            if not selected:
                return
            target = self.associations if kind == "association" else self.version_history
            for target_name in selected:
                target.append({"FileName": target_name})
                if kind == "association":
                    self.update_reciprocal_association(target_name, True)
            self.refresh_links()
            self.save_current()
            dialog.destroy()

        choose.configure(command=add_selected)
        table.bind("<<TreeviewSelect>>", lambda _event: choose.configure(state="normal" if table.selection() else "disabled"))
        table.bind("<Double-1>", add_selected)

    def remove_link(self, kind):
        tree = self.association_tree if kind == "association" else self.history_tree
        selected = tree.selection()
        if not selected:
            return
        records = self.associations if kind == "association" else self.version_history
        removed = records.pop(int(selected[0]))
        if kind == "association" and removed.get("FileName"):
            self.update_reciprocal_association(removed["FileName"], False)
        self.refresh_links()
        self.save_current()

    def open_link(self, tree):
        selected = tree.selection()
        if not selected:
            return
        records = self.associations if tree is self.association_tree else self.version_history
        target = records[int(selected[0])].get("FileName")
        if not target:
            messagebox.showinfo(APP_NAME, "This legacy association does not identify a document that can be opened.")
            return
        if target not in self.tree.get_children():
            self.search_var.set("")
            self.status_filter_var.set("All statuses")
        if target in self.tree.get_children():
            self.navigation_history.append(self.selected_name)
            self.back_button.configure(state="normal")
            self.tree.selection_set(target)
            self.tree.focus(target)
            self.tree.see(target)
            self.on_select()
        else:
            messagebox.showwarning(APP_NAME, f"The linked document is no longer in the selected folder:\n\n{target}")

    def go_back(self):
        while self.navigation_history:
            target = self.navigation_history.pop()
            if target in self.tree.get_children():
                self.tree.selection_set(target)
                self.tree.focus(target)
                self.tree.see(target)
                self.on_select()
                break
        self.back_button.configure(state="normal" if self.navigation_history else "disabled")

    def save_current(self):
        if not self.selected_name:
            return
        self.data[self.selected_name] = {
            "FileName": self.selected_name,
            **{key: var.get().strip() for key, var in self.fields.items()},
            "Groups": json.dumps(self.groups, ensure_ascii=False),
            "AssociatedDocs": json.dumps(self.associations, ensure_ascii=False),
            "Attachments": json.dumps(self.attachments, ensure_ascii=False),
            "VersionHistory": json.dumps(self.version_history, ensure_ascii=False),
        }
        if not self.write_metadata():
            return
        if not self.scan_folder_files():
            return
        self.apply_filter()
        if self.selected_name in self.tree.get_children():
            self.tree.selection_set(self.selected_name)
            self.tree.focus(self.selected_name)
        self.summary_var.set(f"Saved details for {self.selected_name}")

    def write_metadata(self):
        temp_path = self.meta_path.with_suffix(".efm.tmp")
        try:
            with temp_path.open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.DictWriter(stream, fieldnames=METADATA_COLUMNS, extrasaction="ignore")
                writer.writeheader()
                for name in sorted(self.data, key=str.casefold):
                    writer.writerow(self.data[name])
            os.replace(temp_path, self.meta_path)
        except OSError as exc:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass
            messagebox.showerror(APP_NAME, f"Could not save {META_NAME}:\n\n{exc}")
            return False
        return True

    def open_file(self):
        if not self.selected_name:
            return
        path = self.folder / self.selected_name
        try:
            os.startfile(path)
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Could not open the file:\n\n{exc}")


if __name__ == "__main__":
    try:
        EFM().mainloop()
    except Exception as exc:
        messagebox.showerror(APP_NAME, f"The application encountered an unexpected error:\n\n{exc}")
