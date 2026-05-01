import tkinter
import tkinter.filedialog
import customtkinter
import os
import zipfile
import threading
from icecream import ic
import doc_engine

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light")
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue")


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("Performance Testing")
        self.geometry(f"{1100}x{580}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)
        self.grid_columnconfigure(3, weight=2)
        self.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=7, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Performance Testing", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

        # create file input button and text field
        self.file_button = customtkinter.CTkButton(self, text="Select File", command=self.select_file)
        self.file_button.grid(row=0, column=1, padx=(20, 0), pady=(20, 10), sticky="ew")
        self.file_label = customtkinter.CTkEntry(self, placeholder_text="No file selected")
        self.file_label.grid(row=0, column=2, padx=(20, 0), pady=(20, 10), sticky="ew")

        # create numerical inputs for file count and record count
        self.file_count_label = customtkinter.CTkLabel(self, text="File Count:")
        self.file_count_label.grid(row=1, column=1, padx=(20, 0), pady=(10, 10), sticky="w")
        self.file_count_entry = customtkinter.CTkEntry(self, validate="key", validatecommand=(self.register(self.validate_integer), '%P'))
        self.file_count_entry.grid(row=1, column=2, padx=(20, 0), pady=(10, 10), sticky="ew")

        self.record_count_label = customtkinter.CTkLabel(self, text="Record Count:")
        self.record_count_label.grid(row=2, column=1, padx=(20, 0), pady=(10, 10), sticky="w")
        self.record_count_entry = customtkinter.CTkEntry(self, placeholder_text="10, 50, 100, 1000", placeholder_text_color=("gray60","gray40"), validate="key", validatecommand=(self.register(self.validate_record_count), '%P'))
        self.record_count_entry.grid(row=2, column=2, padx=(20, 0), pady=(10, 10), sticky="ew")

        # create reference column input
        self.ref_column_label = customtkinter.CTkLabel(self, text="Reference Column:")
        self.ref_column_label.grid(row=3, column=1, padx=(20, 0), pady=(10, 10), sticky="w")
        self.ref_column_entry = customtkinter.CTkEntry(self, placeholder_text="DocumentNo")
        self.ref_column_entry.grid(row=3, column=2, padx=(20, 0), pady=(10, 10), sticky="ew")
        self.ref_column_entry.insert(0, "DocumentNo")

        # create folder select input
        self.folder_button = customtkinter.CTkButton(self, text="Select Destination Folder", command=self.select_folder)
        self.folder_button.grid(row=4, column=1, padx=(20, 0), pady=(20, 10), sticky="ew")
        self.folder_label = customtkinter.CTkEntry(self, placeholder_text="No folder selected")
        self.folder_label.grid(row=4, column=2, padx=(20, 0), pady=(20, 10), sticky="ew")

        # create process button
        self.process_button = customtkinter.CTkButton(self, text="Process", command=self.start_process_thread)
        self.process_button.grid(row=5, column=1, padx=(20, 0), pady=(20, 10), sticky="ew")

        # create create zip file button
        self.zip_button = customtkinter.CTkButton(self, text="Create Zip File", command=self.create_zip)
        self.zip_button.grid(row=5, column=2, padx=(20, 0), pady=(20, 10), sticky="ew")

        # create clear button
        self.clear_button = customtkinter.CTkButton(self, text="Clear Logs", command=self.clear_logs)
        self.clear_button.grid(row=6, column=1, padx=(20, 0), pady=(20, 10), sticky="ew")

        # create clear folder button
        self.clear_folder_button = customtkinter.CTkButton(self, text="Clear Folder", command=self.clear_folder)
        self.clear_folder_button.grid(row=6, column=2, padx=(20, 0), pady=(20, 10), sticky="ew")

        # create large text output area for messages
        self.text_output = customtkinter.CTkTextbox(self, width=250)
        self.text_output.grid(row=0, column=3, rowspan=7, padx=(20, 20), pady=(20, 20), sticky="nsew")

    def select_file(self):
        file_path = tkinter.filedialog.askopenfilename()
        if file_path:
            self.file_label.delete(0, tkinter.END)
            self.file_label.insert(0, file_path)

    def select_folder(self):
        folder_path = tkinter.filedialog.askdirectory()
        if folder_path:
            self.folder_label.delete(0, tkinter.END)
            self.folder_label.insert(0, folder_path)

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def validate_integer(self, value_if_allowed):
        if value_if_allowed == "" or value_if_allowed.isdigit():
            return True
        else:
            return False

    def validate_record_count(self, value_if_allowed):
        """Allow digits, commas, and spaces for comma-separated record counts."""
        if value_if_allowed == "":
            return True
        return all(c.isdigit() or c in (",", " ") for c in value_if_allowed)

    def start_process_thread(self):
        process_thread = threading.Thread(target=self.process)
        process_thread.start()

    def _log(self, message: str):
        """Append a message to the text output (thread-safe via self.after)."""
        def _append():
            self.text_output.insert("end", message + "\n")
            self.text_output.see("end")
        self.after(0, _append)

    def _set_processing_state(self, is_processing: bool):
        """Enable/disable the Process button from any thread."""
        state = "disabled" if is_processing else "normal"
        self.after(0, lambda: self.process_button.configure(state=state))

    def process(self):
        input_file = self.file_label.get()
        output_folder = self.folder_label.get()
        file_count = self.file_count_entry.get()
        record_count_raw = self.record_count_entry.get().strip()
        ref_column_name = self.ref_column_entry.get().strip()

        if not input_file or not output_folder or not file_count or not record_count_raw or not ref_column_name:
            self._log("⚠ Please fill in all fields before processing.")
            return

        try:
            file_count = int(file_count)
        except ValueError:
            self._log("⚠ File Count must be an integer.")
            return

        # Parse record_count: supports single int or comma-separated list
        try:
            if "," in record_count_raw:
                record_counts = [int(x.strip()) for x in record_count_raw.split(",") if x.strip()]
            else:
                record_counts = [int(record_count_raw)]
        except ValueError:
            self._log("⚠ Record Count must be integers (e.g. 100 or 10, 50, 100, 1000).")
            return

        if not os.path.exists(input_file):
            self._log(f"⚠ File not found: {input_file}")
            return

        # Disable the button and run the engine
        self._set_processing_state(True)
        self._log("═══ Starting file generation ═══")

        try:
            summary = doc_engine.run(
                input_file=input_file,
                output_folder=output_folder,
                file_count=file_count,
                record_counts=record_counts,
                ref_column_name=ref_column_name,
                progress_callback=lambda done, total, msg: self._log(msg),
            )
            ic(summary)
        except Exception as e:
            self._log(f"✗ Engine error: {e}")
            ic(f"Engine error: {e}")
        finally:
            self._set_processing_state(False)

    def create_zip(self):
        output_folder = self.folder_label.get()
        if not output_folder:
            ic("Please select a destination folder.")
            return

        zip_filename = os.path.join(os.path.dirname(output_folder), f"{os.path.basename(output_folder)}.zip")
        try:
            with zipfile.ZipFile(zip_filename, 'w') as zipf:
                for root, dirs, files in os.walk(output_folder):
                    for file in files:
                        file_path = os.path.join(root, file)
                        zipf.write(file_path, os.path.relpath(file_path, os.path.dirname(output_folder)))
            ic(f"Zip file created: {zip_filename}")
        except Exception as e:
            ic(f"An error occurred while creating the zip file: {e}")

    def clear_logs(self):
        self.text_output.delete(1.0, tkinter.END)

    def clear_folder(self):
        output_folder = self.folder_label.get()
        if not output_folder:
            ic("Please select a destination folder.")
            return

        try:
            for root, dirs, files in os.walk(output_folder):
                for file in files:
                    os.remove(os.path.join(root, file))
            ic(f"All files in {output_folder} have been deleted.")
        except Exception as e:
            ic(f"An error occurred while clearing the folder: {e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()