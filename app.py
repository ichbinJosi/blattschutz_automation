import os
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES

from add_blattschutz import make_non_yellow_cells_readonly

# CustomTkinter Design-Einstellungen
ctk.set_appearance_mode("System")  # "System", "Dark", "Light"
ctk.set_default_color_theme("blue")


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("Excel Schreibschutz-Tool")
        self.geometry("520x480")
        self.resizable(False, False)

        self.file_path = None

        # Haupt-Frame für CustomTkinter Styling
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True)

        # UI Komponenten aufbauen
        self._create_widgets()

    def _create_widgets(self):
        # Titel
        self.label_title = ctk.CTkLabel(
            self.main_frame, text="Excel Schreibschutz", font=ctk.CTkFont(size=22, weight="bold")
        )
        self.label_title.pack(pady=(20, 10))

        # Drag & Drop Zone
        self.drop_frame = ctk.CTkFrame(
            self.main_frame,
            width=460,
            height=110,
            corner_radius=10,
            border_width=0,  # Start with no border
        )
        self.drop_frame.pack(pady=10, padx=20)
        self.drop_frame.pack_propagate(False)

        self.drop_label = ctk.CTkLabel(
            self.drop_frame,
            text="Excel-Datei (.xlsx) hierher ziehen\noder Drag & Drop nutzen",
            font=ctk.CTkFont(size=14),
        )
        self.drop_label.place(relx=0.5, rely=0.5, anchor="center")

        # Drag & Drop Events registrieren
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<DropEnter>>", self._on_drag_enter)
        self.drop_frame.dnd_bind("<<DropLeave>>", self._on_drag_leave)
        self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)

        # Eingabe: Name der neuen Datei + .xlsx Suffix Container
        self.name_container = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.name_container.pack(pady=8)

        self.name_entry = ctk.CTkEntry(
            self.name_container,
            placeholder_text="Name des geschützten Dokuments",
            width=250,
        )
        self.name_entry.pack(side="left", padx=(0, 5))
        self.name_entry.bind("<KeyRelease>", lambda e: self._validate_inputs())

        self.ext_label = ctk.CTkLabel(
            self.name_container,
            text=".xlsx",
            text_color="gray",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.ext_label.pack(side="left")

        # Eingabe: Passwort für Blattschutz (Pflichtfeld)
        self.pass_entry = ctk.CTkEntry(
            self.main_frame,
            placeholder_text="Passwort für Blattschutz (Erforderlich)",
            show="*",
            width=286,
        )
        self.pass_entry.pack(pady=8)
        self.pass_entry.bind("<KeyRelease>", lambda e: self._validate_inputs())

        # Start Button (Standardmäßig deaktiviert)
        self.btn_process = ctk.CTkButton(
            self.main_frame,
            text="Schreibschutz anwenden",
            command=self._process_file,
            font=ctk.CTkFont(weight="bold"),
            height=35,
            state="disabled",
        )
        self.btn_process.pack(pady=10)

        # Status / Erfolgsnachricht (Klickbar für Reset)
        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="",
            font=ctk.CTkFont(size=12),
            wraplength=460,
            cursor="hand2",
        )
        self.status_label.pack(pady=10)
        self.status_label.bind("<Button-1>", self._reset_form)

    def _on_drag_enter(self, event):
        self.drop_frame.configure(
            border_width=2,
            border_color="#3B8ED0"
        )

    def _on_drag_leave(self, event):
        self.drop_frame.configure(border_width=0)


    def _on_drop(self, event):
        self._on_drag_leave(event)
        file_path = event.data.strip("{}")

        if file_path.lower().endswith(".xlsx"):
            self.file_path = file_path
            filename = os.path.basename(file_path)
            self.drop_label.configure(
                text=f"Ausgewählt:\n{filename}", text_color="#2FA572"
            )
            self.status_label.configure(text="")
        else:
            self.file_path = None
            self.status_label.configure(
                text="Bitte nur .xlsx Dateien ablegen!", text_color="red"
            )

        self._validate_inputs()

    def _validate_inputs(self):
        """Aktiviert den Button nur, wenn alle Pflichtfelder vorhanden sind."""
        has_file = self.file_path is not None
        has_name = bool(self.name_entry.get().strip())
        has_password = bool(self.pass_entry.get().strip())

        if has_file and has_name and has_password:
            self.btn_process.configure(state="normal")
        else:
            self.btn_process.configure(state="disabled")

    def _reset_form(self, event=None):
        """Setzt die Oberfläche zurück, um ein neues Dokument zu bearbeiten."""
        self.file_path = None
        self.drop_label.configure(
            text="Excel-Datei (.xlsx) hierher ziehen\noder Drag & Drop nutzen",
            text_color=("black", "white"),
        )
        self.name_entry.delete(0, "end")
        self.pass_entry.delete(0, "end")
        self.status_label.configure(text="")
        self._validate_inputs()

    def _process_file(self):
        output_name = self.name_entry.get().strip()
        password = self.pass_entry.get().strip()

        # Dateiendung .xlsx automatisch anfügen, falls nicht eingegeben
        if not output_name.lower().endswith(".xlsx"):
            output_name += ".xlsx"

        try:
            dir_name, _ = os.path.split(self.file_path)
            output_path = os.path.join(dir_name, output_name)

            # Aufruf deiner Funktion aus add_blattschutz.py
            make_non_yellow_cells_readonly(self.file_path, output_path, password)

            # Erfolgsnachricht
            self.status_label.configure(
                text=f"✅ Erfolgreich geschützt!\n\nGespeichert unter:\n{output_path}\n\n👉 Klicke hier, um eine weitere Datei zu bearbeiten.",
                text_color="#2FA572",
            )

        except Exception as e:
            self.status_label.configure(
                text=f"Fehler: {str(e)}", text_color="red"
            )


if __name__ == "__main__":
    app = App()
    app.mainloop()