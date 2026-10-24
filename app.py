import os
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES

from add_blattschutz import make_non_yellow_cells_readonly

# CustomTkinter Design-Einstellungen
ctk.set_appearance_mode("System")  # "System", "Dark", "Light"
ctk.set_default_color_theme("blue")


# Vererbung von TkinterDnD.Tk zur korrekten Drag-and-Drop-Initialisierung
class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("Excel Schreibschutz-Tool")
        self.geometry("520x460")
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
        self.drop_frame = ctk.CTkFrame(self.main_frame, width=460, height=120, corner_radius=10)
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
        self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)

        # Eingabe: Name der neuen Datei (ohne Passwort-Maskierung)
        self.name_entry = ctk.CTkEntry(
            self.main_frame,
            placeholder_text="Name des geschützten Dokuments",
            width=300,
        )
        self.name_entry.pack(pady=10)

        # Eingabe: Passwort für Blattschutz (mit show="*" maskiert)
        self.pass_entry = ctk.CTkEntry(
            self.main_frame,
            placeholder_text="Passwort für Blattschutz (optional)",
            show="*",
            width=300,
        )
        self.pass_entry.pack(pady=10)

        # Start Button
        self.btn_process = ctk.CTkButton(
            self.main_frame,
            text="Schreibschutz anwenden",
            command=self._process_file,
            font=ctk.CTkFont(weight="bold"),
            height=35,
        )
        self.btn_process.pack(pady=10)

        # Status Anzeigelabel
        self.status_label = ctk.CTkLabel(
            self.main_frame, text="", font=ctk.CTkFont(size=12), wraplength=460
        )
        self.status_label.pack(pady=10)

    def _on_drop(self, event):
        file_path = event.data.strip("{}")

        if file_path.lower().endswith(".xlsx"):
            self.file_path = file_path
            filename = os.path.basename(file_path)
            self.drop_label.configure(
                text=f"Ausgewählt:\n{filename}", text_color="#2FA572"
            )
            self.status_label.configure(text="")
        else:
            self.status_label.configure(
                text="Bitte nur .xlsx Dateien ablegen!", text_color="red"
            )

    def _process_file(self):
        if not self.file_path:
            self.status_label.configure(
                text="Bitte zuerst eine Datei hineinziehen!", text_color="red"
            )
            return

        output_name = self.name_entry.get().strip()
        if not output_name:
            self.status_label.configure(
                text="Bitte einen Dokumentnamen für das geschützte Dokument angeben.",
                text_color="red",
            )
            return

        password = self.pass_entry.get().strip()

        # Dateiendung .xlsx automatisch anfügen, falls nicht eingegeben
        if not output_name.lower().endswith(".xlsx"):
            output_name += ".xlsx"

        try:
            dir_name, _ = os.path.split(self.file_path)
            output_path = os.path.join(dir_name, output_name)

            make_non_yellow_cells_readonly(self.file_path, output_path, password)

            self.status_label.configure(
                text=f"Gespeichert unter:\n{output_path}",
                text_color="green",
            )

        except Exception as e:
            self.status_label.configure(
                text=f"Fehler: {str(e)}", text_color="red"
            )


if __name__ == "__main__":
    app = App()
    app.mainloop()