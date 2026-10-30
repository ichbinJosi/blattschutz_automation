import os
import subprocess
import ctypes
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import filedialog, messagebox

from add_blattschutz import make_non_yellow_cells_readonly

# ---------------------------------------------------------
# High-DPI / Schärfe-Fix für Windows
# ---------------------------------------------------------
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Process_Per_Monitor_DPI_Aware
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# App-Theme
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")


class EyeCanvas(ctk.CTkCanvas):
    """Zeichnet ein gestochen scharfes, minimalistisches Augensymbol."""
    def __init__(self, parent, size=20, color="#374151", **kwargs):
        super().__init__(
            parent,
            width=size,
            height=size,
            bg=parent.cget("fg_color") if hasattr(parent, "cget") else "#F3F4F6",
            highlightthickness=0,
            bd=0,
            **kwargs
        )
        self.size = size
        self.color = color
        self.draw_eye(closed=False)

    def draw_eye(self, closed=False):
        self.delete("all")
        w, h = self.size, self.size
        pad = 2

        if not closed:
            # Offenes Auge (Bogen oben & unten + Iris)
            self.create_arc(pad, pad-2, w-pad, h-pad+2, start=25, extent=130, style="arc", outline=self.color, width=2)
            self.create_arc(pad, pad-2, w-pad, h-pad+2, start=205, extent=130, style="arc", outline=self.color, width=2)
            self.create_oval(w/2-3, h/2-3, w/2+3, h/2+3, fill=self.color, outline="")
        else:
            # Geschlossenes Auge (Bogen unten + Striche für Wimpern)
            self.create_arc(pad, pad-4, w-pad, h-pad, start=200, extent=140, style="arc", outline=self.color, width=2)
            # Diagonaler Durchstreich-Strich
            self.create_line(pad+2, pad+2, w-pad-2, h-pad-2, fill=self.color, width=2)


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("Excel Schreibschutz-Tool")
        self.geometry("540x620")
        self.minsize(460, 560)  # Mindestgröße festlegen
        self.resizable(True, True)  # Fenstergröße wieder anpassbar machen

        self._create_variables()
        self._create_widgets()
        self._bind_events()

    def _create_variables(self):
        self.file_path = None
        self.output_path = None
        self.processing = False
        self.show_password = False

    def _create_widgets(self):
        # Haupt-Container dehnt sich bei Fenstervergrößerung mit aus
        self.main_container = ctk.CTkFrame(self, fg_color="#F8F9FA")
        self.main_container.pack(fill="both", expand=True, padx=25, pady=25)

        # 1. Header
        self.header_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.header_frame.pack(fill="x", pady=(0, 15))

        self.lbl_title = ctk.CTkLabel(
            self.header_frame,
            text="Excel Schreibschutz",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="#1F2937",
        )
        self.lbl_title.pack(anchor="w")

        self.lbl_subtitle = ctk.CTkLabel(
            self.header_frame,
            text="Schützt automatisiert alle Zellen ohne gelbe Markierung",
            font=ctk.CTkFont(size=12),
            text_color="#6B7280",
        )
        self.lbl_subtitle.pack(anchor="w", pady=(2, 0))

        # 2. Upload / Drop-Card
        self.drop_card = ctk.CTkFrame(
            self.main_container,
            height=130,
            corner_radius=12,
            border_width=1,
            border_color="#D1D5DB",
            fg_color="#FFFFFF",
            cursor="hand2",
        )
        self.drop_card.pack(fill="x", pady=(0, 15))
        self.drop_card.pack_propagate(False)

        self.drop_icon = ctk.CTkLabel(
            self.drop_card,
            text="📄",
            font=ctk.CTkFont(size=30),
            cursor="hand2",
        )
        self.drop_icon.pack(pady=(20, 2))

        self.drop_label = ctk.CTkLabel(
            self.drop_card,
            text="Excel-Datei hierher ziehen oder klicken",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#374151",
            cursor="hand2",
        )
        self.drop_label.pack()

        self.drop_sublabel = ctk.CTkLabel(
            self.drop_card,
            text="Nur .xlsx Dateien unterstützt",
            font=ctk.CTkFont(size=11),
            text_color="#9CA3AF",
            cursor="hand2",
        )
        self.drop_sublabel.pack(pady=(2, 0))

        # 3. Formularbereich
        self.form_card = ctk.CTkFrame(
            self.main_container,
            corner_radius=12,
            border_width=1,
            border_color="#E5E7EB",
            fg_color="#FFFFFF",
        )
        self.form_card.pack(fill="x", pady=(0, 15), ipadx=15, ipady=15)

        # Feld 1: Dateiname
        self.lbl_name_tag = ctk.CTkLabel(
            self.form_card,
            text="Name des neuen Dokuments",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#4B5563",
        )
        self.lbl_name_tag.pack(anchor="w", pady=(0, 4))

        self.name_box = ctk.CTkFrame(self.form_card, fg_color="transparent")
        self.name_box.pack(fill="x", pady=(0, 14))

        self.name_entry = ctk.CTkEntry(
            self.name_box,
            placeholder_text="z. B. Muster_geschuetzt",
            height=38,
            corner_radius=8,
            fg_color="#F9FAFB",
            text_color="#111827",
            placeholder_text_color="#9CA3AF",
            border_color="#D1D5DB",
            border_width=1,
        )
        self.name_entry.pack(side="left", fill="x", expand=True)

        self.ext_badge = ctk.CTkLabel(
            self.name_box,
            text=".xlsx",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#6B7280",
            width=40,
        )
        self.ext_badge.pack(side="right", padx=(8, 0))

        # Feld 2: Passwort
        self.lbl_pass_tag = ctk.CTkLabel(
            self.form_card,
            text="Passwort für Blattschutz",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#4B5563",
        )
        self.lbl_pass_tag.pack(anchor="w", pady=(0, 4))

        self.pass_box = ctk.CTkFrame(self.form_card, fg_color="transparent")
        self.pass_box.pack(fill="x")

        self.pass_entry = ctk.CTkEntry(
            self.pass_box,
            placeholder_text="Passwort eingeben",
            show="*",
            height=38,
            corner_radius=8,
            fg_color="#F9FAFB",
            text_color="#111827",
            placeholder_text_color="#9CA3AF",
            border_color="#D1D5DB",
            border_width=1,
        )
        self.pass_entry.pack(side="left", fill="x", expand=True)

        # Eleganter Button mit gezeichnetem Vektor-Augen-Icon
        self.btn_show_pass = ctk.CTkButton(
            self.pass_box,
            text="",
            width=42,
            height=38,
            corner_radius=8,
            fg_color="#F3F4F6",
            hover_color="#E5E7EB",
            command=self._toggle_password,
        )
        self.btn_show_pass.pack(side="right", padx=(8, 0))

        # Vektor-Auge platzieren
        self.eye_icon = EyeCanvas(self.btn_show_pass, size=20, color="#374151")
        self.eye_icon.place(relx=0.5, rely=0.5, anchor="center")
        self.eye_icon.bind("<Button-1>", lambda e: self._toggle_password())

        # Progress Bar
        self.progress = ctk.CTkProgressBar(
            self.main_container, height=4, corner_radius=2, progress_color="#2563EB"
        )
        self.progress.pack(fill="x", pady=(0, 12))
        self.progress.set(0)

        # Action Button
        self.btn_process = ctk.CTkButton(
            self.main_container,
            text="Schreibschutz anwenden",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=44,
            corner_radius=8,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            command=self._process_file,
            state="disabled",
        )
        self.btn_process.pack(fill="x")

        # Status Nachricht
        self.status_label = ctk.CTkLabel(
            self.main_container,
            text="",
            font=ctk.CTkFont(size=12),
            wraplength=450,
        )
        self.status_label.pack(pady=(10, 0))

        # 4. Success Overlay Screen
        self.success_card = ctk.CTkFrame(
            self.main_container,
            corner_radius=12,
            border_width=1,
            border_color="#10B981",
            fg_color="#FFFFFF",
        )

        self.succ_icon = ctk.CTkLabel(
            self.success_card, text="✅", font=ctk.CTkFont(size=36)
        )
        self.succ_icon.pack(pady=(20, 5))

        self.succ_title = ctk.CTkLabel(
            self.success_card,
            text="Erfolgreich geschützt!",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#059669",
        )
        self.succ_title.pack()

        self.succ_path_label = ctk.CTkLabel(
            self.success_card,
            text="",
            font=ctk.CTkFont(size=11),
            text_color="#6B7280",
            wraplength=420,
        )
        self.succ_path_label.pack(pady=(8, 15))

        self.btn_open_folder = ctk.CTkButton(
            self.success_card,
            text="📁 Ordner öffnen",
            font=ctk.CTkFont(weight="bold"),
            height=36,
            fg_color="#F3F4F6",
            hover_color="#E5E7EB",
            text_color="#1F2937",
            command=self._open_output_folder,
        )
        self.btn_open_folder.pack(pady=(0, 10))

        self.btn_reset = ctk.CTkButton(
            self.success_card,
            text="Weitere Datei bearbeiten",
            font=ctk.CTkFont(weight="bold"),
            height=38,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            command=self._reset_form,
        )
        self.btn_reset.pack(pady=(0, 20))

    def _bind_events(self):
        targets = [self.drop_card, self.drop_label, self.drop_icon, self.drop_sublabel]
        for w in targets:
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<DropEnter>>", self._on_drag_enter)
            w.dnd_bind("<<DropLeave>>", self._on_drag_leave)
            w.dnd_bind("<<Drop>>", self._on_drop)
            w.bind("<Button-1>", lambda e: self._browse_file())

        self.name_entry.bind("<KeyRelease>", lambda e: self._validate_inputs())
        self.pass_entry.bind("<KeyRelease>", lambda e: self._validate_inputs())
        self.bind("<Return>", self._on_enter_pressed)

    def _toggle_password(self):
        self.show_password = not self.show_password
        self.pass_entry.configure(show="" if self.show_password else "*")
        self.eye_icon.draw_eye(closed=self.show_password)

    def _on_drag_enter(self, event):
        if not self.processing:
            self.drop_card.configure(
                fg_color="#EFF6FF",
                border_color="#2563EB",
            )

    def _on_drag_leave(self, event=None):
        self.drop_card.configure(
            fg_color="#FFFFFF",
            border_color="#D1D5DB",
        )

    def _on_drop(self, event):
        self._on_drag_leave()
        if self.processing:
            return

        file_path = event.data.strip("{}")
        if file_path.lower().endswith(".xlsx"):
            self._set_selected_file(file_path)
        else:
            self.file_path = None
            self.status_label.configure(
                text="❌ Nur .xlsx Dateien werden unterstützt.", text_color="#EF4444"
            )
            self._validate_inputs()

    def _browse_file(self):
        if self.processing:
            return
        file_path = filedialog.askopenfilename(filetypes=[("Excel Dateien", "*.xlsx")])
        if file_path:
            self._set_selected_file(file_path)

    def _set_selected_file(self, file_path):
        self.file_path = file_path
        filename = os.path.basename(file_path)

        self.drop_icon.configure(text="📊")
        self.drop_label.configure(text=filename, text_color="#059669")
        self.drop_sublabel.configure(text="Klicke zum Ändern der Datei")

        if not self.name_entry.get().strip():
            suggested_name = os.path.splitext(filename)[0] + "_geschuetzt"
            self.name_entry.insert(0, suggested_name)

        self.status_label.configure(text="")
        self._validate_inputs()

    def _validate_inputs(self):
        has_file = self.file_path is not None
        has_name = bool(self.name_entry.get().strip())
        has_pass = bool(self.pass_entry.get().strip())

        if has_file and has_name and has_pass and not self.processing:
            self.btn_process.configure(state="normal")
        else:
            self.btn_process.configure(state="disabled")

    def _on_enter_pressed(self, event):
        if self.btn_process.cget("state") == "normal" and not self.processing:
            self._process_file()

    def _reset_form(self):
        self.file_path = None
        self.output_path = None
        self.processing = False

        self._on_drag_leave()
        self.drop_icon.configure(text="📄")
        self.drop_label.configure(
            text="Excel-Datei hierher ziehen oder klicken",
            text_color="#374151",
        )
        self.drop_sublabel.configure(text="Nur .xlsx Dateien unterstützt")

        self.name_entry.delete(0, "end")
        self.pass_entry.delete(0, "end")
        self.show_password = False
        self.pass_entry.configure(show="*")
        self.eye_icon.draw_eye(closed=False)

        self.status_label.configure(text="")
        self.progress.set(0)

        self.success_card.pack_forget()
        self.drop_card.pack(fill="x", pady=(0, 15))
        self.form_card.pack(fill="x", pady=(0, 15), ipadx=15, ipady=15)
        self.progress.pack(fill="x", pady=(0, 12))
        self.btn_process.pack(fill="x")
        self._validate_inputs()

    def _open_output_folder(self):
        if self.output_path and os.path.exists(self.output_path):
            folder = os.path.dirname(self.output_path)
            if os.name == "nt":
                os.startfile(folder)
            else:
                subprocess.run(["open" if os.name == "posix" else "xdg-open", folder])

    def _process_file(self):
        output_name = self.name_entry.get().strip()
        password = self.pass_entry.get().strip()

        if not output_name.lower().endswith(".xlsx"):
            output_name += ".xlsx"

        dir_name, _ = os.path.split(self.file_path)
        self.output_path = os.path.join(dir_name, output_name)

        if os.path.exists(self.output_path):
            answer = messagebox.askyesno(
                "Datei überschreiben?",
                f"Die Datei '{output_name}' existiert bereits.\nSoll sie überschrieben werden?",
            )
            if not answer:
                return

        self.processing = True
        self.btn_process.configure(state="disabled", text="Verarbeite Schreibschutz...")
        self.progress.set(0.4)
        self.update()

        try:
            make_non_yellow_cells_readonly(self.file_path, self.output_path, password)

            self.progress.set(1.0)
            self.update()

            self.drop_card.pack_forget()
            self.form_card.pack_forget()
            self.progress.pack_forget()
            self.btn_process.pack_forget()
            self.status_label.configure(text="")

            self.succ_path_label.configure(
                text=f"Gespeichert unter:\n{self.output_path}"
            )
            self.success_card.pack(fill="x", pady=10, ipadx=10)

        except Exception as e:
            self.progress.set(0)
            self.status_label.configure(
                text=f"❌ Fehler: {str(e)}", text_color="#EF4444"
            )
            self.processing = False
            self.btn_process.configure(text="Schreibschutz anwenden")
            self._validate_inputs()


if __name__ == "__main__":
    app = App()
    app.mainloop()