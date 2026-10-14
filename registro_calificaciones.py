import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None

APP_TITLE = "Aula | Registro de calificaciones"
DB_PATH = Path(__file__).with_name("calificaciones.db")
DEFAULT_GROUPS = {
    "A2": ["Historia", "Español", "Biología", "Álgebra"],
    "A3": ["Literatura", "Geografía", "Química", "Aritmética"],
    "A4": ["Inglés", "Física", "Civismo", "Estadística"],
    "A5": ["Filosofía", "Ecología", "Trigonometría", "Economía"],
}
BG = "#F5F7F6"
CARD = "#FFFFFF"
INK = "#25343B"
MUTED = "#7D8A8F"
GREEN = "#6C9A8B"
GREEN_DARK = "#527D70"
GREEN_PALE = "#EAF3EF"
LINE = "#E5EBE8"
FIELD = "#F8FAF9"


class GradebookApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1240x790")
        self.minsize(1050, 680)
        self.configure(bg=BG)
        self.editing_id = None
        self._init_db()
        self._style()
        self._build_ui()
        self._refresh_group_choices()
        self.refresh_rows()

    def _init_db(self):
        with sqlite3.connect(DB_PATH) as con:
            con.execute("""CREATE TABLE IF NOT EXISTS grades (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student TEXT NOT NULL,
                subject TEXT NOT NULL,
                group_name TEXT NOT NULL,
                score REAL NOT NULL CHECK(score >= 0 AND score <= 100),
                date TEXT NOT NULL,
                notes TEXT DEFAULT ''
            )""")
            con.execute("""CREATE TABLE IF NOT EXISTS school_groups (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE
            )""")
            con.execute("""CREATE TABLE IF NOT EXISTS group_subjects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_id INTEGER NOT NULL REFERENCES school_groups(id) ON DELETE CASCADE,
                name TEXT NOT NULL,
                UNIQUE(group_id, name)
            )""")
            count = con.execute("SELECT COUNT(*) FROM school_groups").fetchone()[0]
            if count == 0:
                for group_name, subjects in DEFAULT_GROUPS.items():
                    cur = con.execute("INSERT INTO school_groups(name) VALUES (?)", (group_name,))
                    group_id = cur.lastrowid
                    con.executemany("INSERT INTO group_subjects(group_id, name) VALUES (?, ?)",
                                    [(group_id, name) for name in subjects])

    def _style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview", background=CARD, foreground=INK, fieldbackground=CARD,
                        rowheight=38, font=("Segoe UI", 10), borderwidth=0)
        style.configure("Treeview.Heading", background="#F0F5F2", foreground="#52635D",
                        font=("Segoe UI Semibold", 9), padding=(10, 11), relief="flat")
        style.map("Treeview", background=[("selected", "#DDECE5")], foreground=[("selected", INK)])
        style.configure("Soft.TCombobox", padding=8, font=("Segoe UI", 10), fieldbackground=FIELD,
                        background=FIELD, foreground=INK, arrowcolor=GREEN_DARK)
        style.map("Soft.TCombobox", fieldbackground=[("readonly", FIELD)],
                  selectbackground=[("readonly", FIELD)], selectforeground=[("readonly", INK)])

    def _build_ui(self):
        header = tk.Frame(self, bg=GREEN, height=128)
        header.pack(fill="x")
        header.pack_propagate(False)
        title_area = tk.Frame(header, bg=GREEN)
        title_area.pack(side="left", fill="y")
        tk.Label(title_area, text="A U L A   /   GESTIÓN ACADÉMICA", bg=GREEN, fg="#E3F0EA",
                 font=("Segoe UI Semibold", 9)).pack(anchor="w", padx=32, pady=(19, 4))
        tk.Label(title_area, text="Registro de calificaciones", bg=GREEN, fg="white",
                 font=("Segoe UI Semibold", 24)).pack(anchor="w", padx=32)
        tk.Label(title_area, text="Un espacio sencillo y ordenado para dar seguimiento al grupo.", bg=GREEN,
                 fg="#E4EFEB", font=("Segoe UI", 10)).pack(anchor="w", padx=33, pady=(3, 0))
        self._button(header, "Ver materias", self.open_subject_manager, "#F1F7F4", GREEN_DARK).pack(
            side="right", padx=30, pady=42)

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=28, pady=22)
        cards = tk.Frame(body, bg=BG)
        cards.pack(fill="x", pady=(0, 17))
        self.stat_total = self._stat_card(cards, "REGISTROS", "0", "Calificaciones guardadas")
        self.stat_avg = self._stat_card(cards, "PROMEDIO GENERAL", "—", "Escala de 0 a 100")
        self.stat_pass = self._stat_card(cards, "APROBACIÓN", "—", "Calificaciones de 60 o más")

        content = tk.Frame(body, bg=BG)
        content.pack(fill="both", expand=True)
        form = tk.Frame(content, bg=CARD, highlightthickness=1, highlightbackground=LINE, width=310)
        form.pack(side="left", fill="y", padx=(0, 17))
        form.pack_propagate(False)
        tk.Label(form, text="Datos del registro", bg=CARD, fg=INK,
                 font=("Segoe UI Semibold", 15)).pack(anchor="w", padx=21, pady=(20, 3))
        tk.Label(form, text="Completa la información del estudiante.", bg=CARD, fg=MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=21, pady=(0, 17))
        self.vars = {k: tk.StringVar() for k in ("student", "group", "subject", "score", "date")}
        self.vars["date"].set(datetime.now().strftime("%Y-%m-%d"))
        self._label(form, "Nombre del estudiante")
        self._entry(form, self.vars["student"])
        self._label(form, "Grupo")
        self.group_input = ttk.Combobox(form, textvariable=self.vars["group"], state="readonly", style="Soft.TCombobox")
        self.group_input.pack(fill="x", padx=20, ipady=3, pady=(0, 12))
        self.group_input.bind("<<ComboboxSelected>>", self._on_group_selected)
        self._label(form, "Materia")
        self.subject_input = ttk.Combobox(form, textvariable=self.vars["subject"], state="readonly", style="Soft.TCombobox")
        self.subject_input.pack(fill="x", padx=20, ipady=3, pady=(0, 12))
        self._label(form, "Calificación (0–100)")
        self._entry(form, self.vars["score"])
        self._label(form, "Fecha (AAAA-MM-DD)")
        self._entry(form, self.vars["date"])
        self._label(form, "Observaciones")
        self.notes = tk.Text(form, height=3, wrap="word", relief="flat", bg=FIELD, fg=INK,
                             font=("Segoe UI", 10), highlightthickness=1, highlightbackground=LINE,
                             highlightcolor=GREEN)
        self.notes.pack(fill="x", padx=20, pady=(0, 14))
        self.save_button = self._button(form, "Guardar calificación", self.save_grade, GREEN, "white")
        self.save_button.pack(fill="x", padx=20, pady=(0, 9), ipady=3)
        self.cancel_button = self._button(form, "Cancelar edición", self.cancel_edit, "#F0F4F2", GREEN_DARK)
        self.cancel_button.pack(fill="x", padx=20, pady=(0, 9), ipady=3)
        self.cancel_button.pack_forget()
        self._button(form, "Limpiar formulario", self.clear_form, "#F5F7F6", "#64736D").pack(fill="x", padx=20, pady=(0, 19), ipady=3)

        panel = tk.Frame(content, bg=CARD, highlightthickness=1, highlightbackground=LINE)
        panel.pack(side="left", fill="both", expand=True)
        top = tk.Frame(panel, bg=CARD)
        top.pack(fill="x", padx=20, pady=(18, 12))
        tk.Label(top, text="Calificaciones registradas", bg=CARD, fg=INK,
                 font=("Segoe UI Semibold", 15)).pack(anchor="w")
        tk.Label(top, text="Busca, filtra y organiza los registros del grupo.", bg=CARD, fg=MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(2, 11))
        filters = tk.Frame(top, bg=CARD)
        filters.pack(fill="x")
        self.search_var = tk.StringVar()
        search = tk.Entry(filters, textvariable=self.search_var, relief="flat", bg=FIELD, fg=INK,
                          font=("Segoe UI", 10), highlightthickness=1, highlightbackground=LINE,
                          highlightcolor=GREEN)
        search.pack(side="left", fill="x", expand=True, ipady=8, padx=(0, 8))
        search.insert(0, "Buscar por estudiante o materia")
        self.search_placeholder = "Buscar por estudiante o materia"
        search.bind("<FocusIn>", self._clear_search_placeholder)
        search.bind("<FocusOut>", self._restore_search_placeholder)
        search.bind("<KeyRelease>", lambda _e: self.refresh_rows())
        self.subject_filter_var = tk.StringVar(value="Todas las materias")
        self.subject_filter = ttk.Combobox(filters, textvariable=self.subject_filter_var,
                                           state="readonly", width=18, style="Soft.TCombobox")
        self.subject_filter.pack(side="left", padx=(0, 8))
        self.subject_filter.bind("<<ComboboxSelected>>", lambda _e: self.refresh_rows())
        self.sort_var = tk.StringVar(value="Más recientes")
        sort_filter = ttk.Combobox(filters, textvariable=self.sort_var,
                                   values=("Más recientes", "Nombre A–Z", "Nombre Z–A"),
                                   state="readonly", width=15, style="Soft.TCombobox")
        sort_filter.pack(side="left", padx=(0, 8))
        sort_filter.bind("<<ComboboxSelected>>", lambda _e: self.refresh_rows())
        self._button(filters, "Exportar Excel", self.export_excel, GREEN_PALE, GREEN_DARK).pack(side="left", padx=(0, 8), ipady=2)
        actions = tk.Frame(panel, bg=CARD)
        actions.pack(fill="x", padx=20, pady=(0, 10))
        self._button(actions, "Editar seleccionado", self.edit_selected, "#EEF3F7", "#526B78").pack(side="left", padx=(0, 8))
        self._button(actions, "Eliminar", self.delete_selected, "#FAEEEE", "#A75D5D").pack(side="left")
        table_wrap = tk.Frame(panel, bg=CARD)
        table_wrap.pack(fill="both", expand=True, padx=20, pady=(0, 8))
        columns = ("student", "subject", "group", "score", "date")
        self.tree = ttk.Treeview(table_wrap, columns=columns, show="headings", selectmode="browse")
        headings = {"student": "ESTUDIANTE", "subject": "MATERIA", "group": "GRUPO", "score": "CALIFICACIÓN", "date": "FECHA"}
        widths = {"student": 200, "subject": 145, "group": 90, "score": 110, "date": 110}
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor="w" if col in ("student", "subject") else "center")
        scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.tag_configure("even", background=CARD)
        self.tree.tag_configure("odd", background="#FAFCFB")
        self.tree.bind("<Double-1>", lambda _e: self.edit_selected())
        self.footer = tk.Label(panel, text="Los datos se guardan automáticamente en este equipo.", bg=CARD,
                               fg=MUTED, font=("Segoe UI", 9), anchor="w")
        self.footer.pack(fill="x", padx=20, pady=(4, 14))

    def _label(self, parent, text):
        tk.Label(parent, text=text, bg=CARD, fg="#52615B",
                 font=("Segoe UI Semibold", 9)).pack(anchor="w", padx=20, pady=(0, 5))

    def _entry(self, parent, variable):
        entry = tk.Entry(parent, textvariable=variable, relief="flat", bg=FIELD, fg=INK,
                         insertbackground=INK, font=("Segoe UI", 10), highlightthickness=1,
                         highlightbackground=LINE, highlightcolor=GREEN)
        entry.pack(fill="x", padx=20, ipady=9, pady=(0, 12))
        return entry

    def _stat_card(self, parent, title, value, subtitle):
        card = tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=LINE)
        card.pack(side="left", fill="x", expand=True, padx=(0, 12))
        tk.Label(card, text=title, bg=CARD, fg=MUTED, font=("Segoe UI Semibold", 8)).pack(anchor="w", padx=17, pady=(12, 2))
        value_label = tk.Label(card, text=value, bg=CARD, fg=GREEN_DARK, font=("Segoe UI Semibold", 21))
        value_label.pack(anchor="w", padx=17)
        tk.Label(card, text=subtitle, bg=CARD, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", padx=17, pady=(0, 12))
        return value_label

    def _button(self, parent, text, command, bg, fg):
        return tk.Button(parent, text=text, command=command, bg=bg, fg=fg,
                         activebackground=bg, activeforeground=fg, relief="flat", cursor="hand2",
                         font=("Segoe UI Semibold", 9), padx=12, pady=8, bd=0)

    def _group_rows(self):
        with sqlite3.connect(DB_PATH) as con:
            return con.execute("SELECT id, name FROM school_groups ORDER BY id").fetchall()

    def _subjects_for_group(self, group_name):
        with sqlite3.connect(DB_PATH) as con:
            return [row[0] for row in con.execute(
                "SELECT s.name FROM group_subjects s JOIN school_groups g ON g.id=s.group_id WHERE g.name=? ORDER BY s.name COLLATE NOCASE",
                (group_name,)).fetchall()]

    def _refresh_group_choices(self, keep_group=None):
        groups = self._group_rows()
        names = [row[1] for row in groups]
        self.group_input["values"] = names
        if keep_group in names:
            self.vars["group"].set(keep_group)
        elif self.vars["group"].get() not in names:
            self.vars["group"].set(names[0] if names else "")
        self._on_group_selected()
        all_subjects = sorted({subject for name in names for subject in self._subjects_for_group(name)}, key=str.casefold)
        self.subject_filter["values"] = ("Todas las materias",) + tuple(all_subjects)
        if self.subject_filter_var.get() not in self.subject_filter["values"]:
            self.subject_filter_var.set("Todas las materias")

    def _on_group_selected(self, _event=None):
        subjects = self._subjects_for_group(self.vars["group"].get())
        self.subject_input["values"] = subjects
        if self.vars["subject"].get() not in subjects:
            self.vars["subject"].set(subjects[0] if subjects else "")

    def _clear_search_placeholder(self, _event=None):
        if self.search_var.get() == self.search_placeholder:
            self.search_var.set("")

    def _restore_search_placeholder(self, _event=None):
        if not self.search_var.get().strip():
            self.search_var.set(self.search_placeholder)
            self.refresh_rows()

    def save_grade(self):
        student = self.vars["student"].get().strip()
        group = self.vars["group"].get().strip()
        subject = self.vars["subject"].get().strip()
        try:
            score = float(self.vars["score"].get().strip().replace(",", "."))
            if not 0 <= score <= 100:
                raise ValueError
        except ValueError:
            messagebox.showerror("Dato inválido", "La calificación debe ser un número entre 0 y 100.", parent=self)
            return
        date = self.vars["date"].get().strip()
        try:
            datetime.strptime(date, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Fecha inválida", "Usa el formato AAAA-MM-DD, por ejemplo 2026-10-13.", parent=self)
            return
        if not student or not group or not subject:
            messagebox.showerror("Campos obligatorios", "Completa el nombre y selecciona un grupo y una materia.", parent=self)
            return
        if subject not in self._subjects_for_group(group):
            messagebox.showerror("Materia no disponible", "La materia debe pertenecer al grupo seleccionado.", parent=self)
            return
        notes = self.notes.get("1.0", "end").strip()
        with sqlite3.connect(DB_PATH) as con:
            if self.editing_id is None:
                con.execute("INSERT INTO grades(student, subject, group_name, score, date, notes) VALUES (?, ?, ?, ?, ?, ?)",
                            (student, subject, group, score, date, notes))
                notice = f"Calificación guardada · {student}"
            else:
                con.execute("UPDATE grades SET student=?, subject=?, group_name=?, score=?, date=?, notes=? WHERE id=?",
                            (student, subject, group, score, date, notes, self.editing_id))
                notice = f"Cambios guardados · {student}"
        self.cancel_edit(clear=True)
        self.refresh_rows()
        self.footer.config(text=notice)

    def edit_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Selecciona un registro", "Selecciona una fila y pulsa Editar seleccionado.", parent=self)
            return
        with sqlite3.connect(DB_PATH) as con:
            row = con.execute("SELECT id, student, subject, group_name, score, date, notes FROM grades WHERE id=?",
                              (int(selected[0]),)).fetchone()
        if not row:
            self.refresh_rows()
            return
        self.editing_id = row[0]
        self.vars["student"].set(row[1])
        self.vars["group"].set(row[3])
        self.subject_input["values"] = self._subjects_for_group(row[3])
        # Permite conservar y editar registros antiguos si su materia fue retirada de la lista.
        current_subjects = list(self.subject_input["values"])
        if row[2] not in current_subjects:
            self.subject_input["values"] = current_subjects + [row[2]]
        self.vars["subject"].set(row[2])
        self.vars["score"].set(str(row[4]))
        self.vars["date"].set(row[5])
        self.notes.delete("1.0", "end")
        self.notes.insert("1.0", row[6] or "")
        self.save_button.config(text="Guardar cambios")
        self.cancel_button.pack(fill="x", padx=20, pady=(0, 9), ipady=3)
        self.footer.config(text=f"Editando registro de {row[1]} · actualiza los datos y guarda los cambios")

    def cancel_edit(self, clear=True):
        self.editing_id = None
        self.save_button.config(text="Guardar calificación")
        self.cancel_button.pack_forget()
        if clear:
            self.clear_form()

    def clear_form(self):
        if self.editing_id is not None:
            self.cancel_edit(clear=False)
        for key in ("student", "subject", "score"):
            self.vars[key].set("")
        self.vars["date"].set(datetime.now().strftime("%Y-%m-%d"))
        self.notes.delete("1.0", "end")
        self._refresh_group_choices()

    def refresh_rows(self):
        if not hasattr(self, "tree"):
            return
        for item in self.tree.get_children():
            self.tree.delete(item)
        query = self.search_var.get().strip() if hasattr(self, "search_var") else ""
        if query == getattr(self, "search_placeholder", None):
            query = ""
        subject = self.subject_filter_var.get() if hasattr(self, "subject_filter_var") else "Todas las materias"
        sort = self.sort_var.get() if hasattr(self, "sort_var") else "Más recientes"
        with sqlite3.connect(DB_PATH) as con:
            all_rows = con.execute("SELECT id, student, subject, group_name, score, date FROM grades").fetchall()
        rows = [r for r in all_rows if (not query or query.casefold() in (r[1] + " " + r[2]).casefold())
                and (subject == "Todas las materias" or r[2] == subject)]
        if sort == "Nombre A–Z":
            rows.sort(key=lambda r: (r[1].casefold(), r[2].casefold(), r[5]))
        elif sort == "Nombre Z–A":
            rows.sort(key=lambda r: (r[1].casefold(), r[2].casefold(), r[5]), reverse=True)
        else:
            rows.sort(key=lambda r: (r[5], r[0]), reverse=True)
        for i, row in enumerate(rows):
            self.tree.insert("", "end", iid=str(row[0]), values=(row[1], row[2], row[3], f"{row[4]:g}", row[5]),
                              tags=("even" if i % 2 == 0 else "odd",))
        total = len(all_rows)
        average = sum(r[4] for r in all_rows) / total if total else 0
        passed = sum(1 for r in all_rows if r[4] >= 60)
        self.stat_total.config(text=str(total))
        self.stat_avg.config(text=f"{average:.1f}" if total else "—")
        self.stat_pass.config(text=f"{passed / total * 100:.0f}%" if total else "—")
        self.footer.config(text=f"Mostrando {len(rows)} de {total} registros · Guardado local automático")

    def open_subject_manager(self):
        manager = tk.Toplevel(self)
        manager.title("Ver materias y grupos")
        manager.geometry("800x680")
        manager.minsize(690, 540)
        manager.configure(bg=BG)
        manager.transient(self)
        manager.grab_set()
        tk.Label(manager, text="Grupos y materias", bg=BG, fg=INK,
                 font=("Segoe UI Semibold", 20)).pack(anchor="w", padx=24, pady=(20, 3))
        tk.Label(manager, text="Cambia el nombre de cada grupo y ajusta sus materias. Escribe una materia por línea.",
                 bg=BG, fg=MUTED, font=("Segoe UI", 10), wraplength=740, justify="left").pack(anchor="w", padx=25, pady=(0, 16))
        scroller = tk.Frame(manager, bg=BG)
        scroller.pack(fill="both", expand=True, padx=22)
        canvas = tk.Canvas(scroller, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroller, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=BG)
        inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw", width=735)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        fields = {}
        with sqlite3.connect(DB_PATH) as con:
            groups = con.execute("SELECT id, name FROM school_groups ORDER BY id").fetchall()
            for group_id, group_name in groups:
                subjects = [r[0] for r in con.execute("SELECT name FROM group_subjects WHERE group_id=? ORDER BY name COLLATE NOCASE", (group_id,)).fetchall()]
                card = tk.Frame(inner, bg=CARD, highlightthickness=1, highlightbackground=LINE)
                card.pack(fill="x", pady=(0, 12))
                name_var = tk.StringVar(value=group_name)
                tk.Label(card, text="NOMBRE DEL GRUPO", bg=CARD, fg=MUTED,
                         font=("Segoe UI Semibold", 8)).pack(anchor="w", padx=16, pady=(12, 4))
                name_entry = tk.Entry(card, textvariable=name_var, relief="flat", bg=FIELD, fg=INK,
                                      font=("Segoe UI", 10), highlightthickness=1, highlightbackground=LINE)
                name_entry.pack(fill="x", padx=16, ipady=7, pady=(0, 9))
                tk.Label(card, text="MATERIAS (una por línea)", bg=CARD, fg=MUTED,
                         font=("Segoe UI Semibold", 8)).pack(anchor="w", padx=16, pady=(0, 4))
                subject_text = tk.Text(card, height=4, wrap="word", relief="flat", bg=FIELD, fg=INK,
                                       font=("Segoe UI", 10), highlightthickness=1, highlightbackground=LINE)
                subject_text.pack(fill="x", padx=16, pady=(0, 14))
                subject_text.insert("1.0", "\n".join(subjects))
                fields[group_id] = (group_name, name_var, subject_text)
        buttons = tk.Frame(manager, bg=BG)
        buttons.pack(fill="x", padx=24, pady=16)
        self._button(buttons, "Guardar cambios", lambda: self.save_group_config(manager, fields), GREEN, "white").pack(side="right")
        self._button(buttons, "Cancelar", manager.destroy, "#EEF3F0", GREEN_DARK).pack(side="right", padx=(0, 9))

    def save_group_config(self, window, fields):
        updated = []
        names = []
        for group_id, (old_name, name_var, subject_text) in fields.items():
            new_name = name_var.get().strip()
            subjects = [line.strip() for line in subject_text.get("1.0", "end").splitlines() if line.strip()]
            normalized = [subject.casefold() for subject in subjects]
            if not new_name:
                messagebox.showerror("Nombre requerido", "Todos los grupos necesitan un nombre.", parent=window)
                return
            if not subjects:
                messagebox.showerror("Materias requeridas", f"Agrega al menos una materia para el grupo {new_name}.", parent=window)
                return
            if len(normalized) != len(set(normalized)):
                messagebox.showerror("Materia repetida", f"Hay materias repetidas en el grupo {new_name}.", parent=window)
                return
            names.append(new_name.casefold())
            updated.append((group_id, old_name, new_name, subjects))
        if len(names) != len(set(names)):
            messagebox.showerror("Nombre repetido", "Cada grupo debe tener un nombre diferente.", parent=window)
            return
        try:
            with sqlite3.connect(DB_PATH) as con:
                for group_id, old_name, new_name, subjects in updated:
                    con.execute("UPDATE school_groups SET name=? WHERE id=?", (new_name, group_id))
                    if old_name != new_name:
                        con.execute("UPDATE grades SET group_name=? WHERE group_name=?", (new_name, old_name))
                    con.execute("DELETE FROM group_subjects WHERE group_id=?", (group_id,))
                    con.executemany("INSERT INTO group_subjects(group_id, name) VALUES (?, ?)",
                                    [(group_id, subject) for subject in subjects])
        except sqlite3.IntegrityError:
            messagebox.showerror("No se pudieron guardar los cambios", "Revisa que no haya nombres duplicados.", parent=window)
            return
        active_group = self.vars["group"].get()
        for _, old_name, new_name, _ in updated:
            if active_group == old_name:
                active_group = new_name
        self._refresh_group_choices(keep_group=active_group)
        self.refresh_rows()
        window.destroy()
        messagebox.showinfo("Configuración guardada", "Se actualizaron los grupos y sus materias.", parent=self)

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Selecciona un registro", "Selecciona una fila de la tabla para eliminarla.", parent=self)
            return
        if not messagebox.askyesno("Confirmar eliminación", "¿Eliminar la calificación seleccionada?", parent=self):
            return
        selected_id = int(selected[0])
        if self.editing_id == selected_id:
            self.cancel_edit()
        with sqlite3.connect(DB_PATH) as con:
            con.execute("DELETE FROM grades WHERE id = ?", (selected_id,))
        self.refresh_rows()

    def export_excel(self):
        if Workbook is None:
            messagebox.showerror("Falta una dependencia", "Instala openpyxl con: pip install openpyxl", parent=self)
            return
        with sqlite3.connect(DB_PATH) as con:
            rows = con.execute("SELECT student, subject, group_name, score, date, notes FROM grades ORDER BY group_name, student COLLATE NOCASE, subject").fetchall()
        if not rows:
            messagebox.showinfo("Sin registros", "Agrega calificaciones antes de exportar.", parent=self)
            return
        path = filedialog.asksaveasfilename(parent=self, title="Guardar reporte de calificaciones", defaultextension=".xlsx",
                                            initialfile=f"calificaciones_{datetime.now():%Y-%m-%d}.xlsx",
                                            filetypes=[("Libro de Excel", "*.xlsx")])
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "Calificaciones"
        ws.append(["REGISTRO DE CALIFICACIONES"])
        ws.merge_cells("A1:F1")
        ws["A1"].font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
        ws["A1"].fill = PatternFill("solid", fgColor="6C9A8B")
        ws["A1"].alignment = Alignment(vertical="center")
        ws.row_dimensions[1].height = 32
        ws.append(["Estudiante", "Materia", "Grupo", "Calificación", "Fecha", "Observaciones"])
        for cell in ws[2]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="527D70")
            cell.alignment = Alignment(horizontal="center")
        for row in rows:
            ws.append(list(row))
        for row in ws.iter_rows(min_row=3):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = Border(bottom=Side(style="hair", color="E5EBE8"))
        for i, width in enumerate([28, 24, 14, 16, 16, 42], start=1):
            ws.column_dimensions[get_column_letter(i)].width = width
        ws.freeze_panes = "A3"
        ws.auto_filter.ref = f"A2:F{ws.max_row}"
        ws.sheet_view.showGridLines = False
        try:
            wb.save(path)
        except OSError as exc:
            messagebox.showerror("No se pudo guardar", f"Revisa que el archivo no esté abierto y tengas permiso.\n\n{exc}", parent=self)
            return
        messagebox.showinfo("Exportación completada", f"Se exportaron {len(rows)} calificaciones a:\n{path}", parent=self)


if __name__ == "__main__":
    GradebookApp().mainloop()
