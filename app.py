import threading
from pathlib import Path
from tkinter import filedialog, messagebox
import zipfile

import customtkinter as ctk

from converter import extract_pdf_to_excel, process_pdf, write_excel, process_zip


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class PDFConverterApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("Finance PDF → Excel")
        self.geometry("700x650")
        self.resizable(False, False)

        self.input_path = None
        self.input_mode = "single"
        self.output_mode = "merge_all"
        self.group_size = 2

        self.build_ui()

    def build_ui(self):

        self.container = ctk.CTkFrame(self, corner_radius=20)
        self.container.pack(fill="both", expand=True, padx=30, pady=30)

        self.title_label = ctk.CTkLabel(
            self.container,
            text="Finance PDF → Excel",
            font=ctk.CTkFont(size=28, weight="bold"),
        )
        self.title_label.pack(pady=(25, 5))

        self.subtitle_label = ctk.CTkLabel(
            self.container,
            text="Settlement Report Converter",
            font=ctk.CTkFont(size=14),
        )
        self.subtitle_label.pack(pady=(0, 18))

        ctk.CTkLabel(
            self.container,
            text="Input",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(pady=(0, 5))

        mode_frame = ctk.CTkFrame(self.container, fg_color="transparent")
        mode_frame.pack()

        self.input_var = ctk.StringVar(value="single")

        ctk.CTkRadioButton(
            mode_frame,
            text="Single PDF",
            variable=self.input_var,
            value="single",
            command=self.change_input_mode,
        ).grid(row=0, column=0, padx=15)

        ctk.CTkRadioButton(
            mode_frame,
            text="Multiple PDFs (ZIP)",
            variable=self.input_var,
            value="zip",
            command=self.change_input_mode,
        ).grid(row=0, column=1, padx=15)

        self.file_label = ctk.CTkLabel(
            self.container,
            text="No PDF selected",
            font=ctk.CTkFont(size=13),
        )
        self.file_label.pack(pady=(15, 7))

        self.select_button = ctk.CTkButton(
            self.container,
            text="Select PDF",
            width=220,
            height=40,
            corner_radius=10,
            command=self.select_input,
        )
        self.select_button.pack(pady=7)

        self.output_frame = ctk.CTkFrame(self.container, fg_color="transparent")

        ctk.CTkLabel(
            self.output_frame,
            text="Output",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(pady=(10, 5))

        self.output_var = ctk.StringVar(value="merge_all")

        ctk.CTkRadioButton(
            self.output_frame,
            text="Separate Excel files",
            variable=self.output_var,
            value="separate",
            command=self.change_output_mode,
        ).pack(anchor="w", padx=85, pady=2)

        ctk.CTkRadioButton(
            self.output_frame,
            text="Merge all PDFs",
            variable=self.output_var,
            value="merge_all",
            command=self.change_output_mode,
        ).pack(anchor="w", padx=85, pady=2)

        group_row = ctk.CTkFrame(self.output_frame, fg_color="transparent")
        group_row.pack(anchor="w", padx=85, pady=2)

        self.group_radio = ctk.CTkRadioButton(
            group_row,
            text="Group PDFs",
            variable=self.output_var,
            value="group",
            command=self.change_output_mode,
        )
        self.group_radio.pack(side="left")

        self.group_entry = ctk.CTkEntry(group_row, width=65)
        self.group_entry.insert(0, "2")
        self.group_entry.pack(side="left", padx=10)

        ctk.CTkLabel(
            group_row,
            text="PDFs per Excel",
        ).pack(side="left")

        self.output_frame.pack(pady=3)

        self.convert_button = ctk.CTkButton(
            self.container,
            text="Convert to Excel",
            width=220,
            height=40,
            corner_radius=10,
            command=self.convert,
            state="disabled",
        )
        self.convert_button.pack(pady=12)

        self.status_label = ctk.CTkLabel(
            self.container,
            text="",
            font=ctk.CTkFont(size=13),
        )
        self.status_label.pack(pady=(8, 3))

        self.progress_bar = ctk.CTkProgressBar(
            self.container,
            width=450,
            height=12,
            corner_radius=6,
        )
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=7)

        self.page_label = ctk.CTkLabel(
            self.container,
            text="",
            font=ctk.CTkFont(size=12),
        )
        self.page_label.pack(pady=3)

        self.change_output_mode()

    def change_input_mode(self):
        self.input_mode = self.input_var.get()
        self.input_path = None
        self.file_label.configure(
            text="No ZIP selected" if self.input_mode == "zip" else "No PDF selected"
        )
        self.select_button.configure(
            text="Select ZIP" if self.input_mode == "zip" else "Select PDF"
        )
        self.convert_button.configure(state="disabled")
        self.status_label.configure(text="")
        self.page_label.configure(text="")
        self.progress_bar.set(0)

    def change_output_mode(self):
        self.output_mode = self.output_var.get()

    def select_input(self):
        if self.input_mode == "zip":
            path = filedialog.askopenfilename(
                title="Select ZIP file",
                filetypes=[("ZIP files", "*.zip")],
            )
        else:
            path = filedialog.askopenfilename(
                title="Select PDF",
                filetypes=[("PDF files", "*.pdf")],
            )

        if not path:
            return

        self.input_path = Path(path)
        self.file_label.configure(text=self.input_path.name)
        self.convert_button.configure(state="normal")
        self.status_label.configure(text="")
        self.page_label.configure(text="")
        self.progress_bar.set(0)

    def update_progress(self, page, total_pages, status):
        self.after(0, self._update_progress, page, total_pages, status)

    def _update_progress(self, page, total_pages, status):
        progress = page / total_pages if total_pages else 0
        self.progress_bar.set(progress)
        self.status_label.configure(text=status)
        self.page_label.configure(text=f"Page {page} / {total_pages}")

    def convert(self):
        if not self.input_path:
            return

        if self.input_mode == "single":
            output_path = self.input_path.with_suffix(".xlsx")
            target = self.run_single
            args = (output_path,)
        else:
            if self.output_mode == "group":
                try:
                    self.group_size = int(self.group_entry.get())
                    if self.group_size < 1:
                        raise ValueError
                except ValueError:
                    messagebox.showerror(
                        "Invalid group size",
                        "Enter a whole number greater than 0.",
                    )
                    return
            target = self.run_zip
            args = ()

        self.set_busy(True)
        self.progress_bar.set(0)
        self.status_label.configure(text="Starting...")
        self.page_label.configure(text="")

        threading.Thread(
            target=target,
            args=args,
            daemon=True,
        ).start()

    def run_single(self, output_path):
        try:
            result = extract_pdf_to_excel(
                str(self.input_path),
                str(output_path),
                progress_callback=self.update_progress,
            )
            self.after(0, self.single_finished, result)
        except Exception as exc:
            self.after(0, self.conversion_failed, str(exc))

    def run_zip(self):
        try:
            result = process_zip(
                str(self.input_path),
                self.output_mode,
                self.group_size,
                progress_callback=self.update_progress,
            )
            self.after(0, self.batch_finished, result)
        except Exception as exc:
            self.after(0, self.conversion_failed, str(exc))

    def single_finished(self, result):
        self.progress_bar.set(1)
        self.status_label.configure(text="Conversion completed ✓")
        self.page_label.configure(text=f"{result['pages']} pages processed")

        validation_message = (
            "Purchase validation passed ✓"
            if result["purchase_valid"]
            else "Purchase validation warning"
        )

        messagebox.showinfo(
            "Conversion Completed",
            (
                f"Excel file created successfully.\n\n"
                f"Pages: {result['pages']}\n"
                f"Sections: {result['sections']}\n"
                f"Total rows: {result['rows']:,}\n\n"
                f"{validation_message}\n"
                f"Purchase rows: {result['purchase_rows']:,}\n"
                f"Expected: {result['expected_purchase_rows']:,}\n\n"
                f"File:\n{result['output']}"
            ),
        )
        self.set_busy(False)

    def batch_finished(self, result):
        self.progress_bar.set(1)
        self.status_label.configure(
            text="Batch completed ✓" if result["warnings"] == 0
            else "Batch completed with warnings"
        )
        self.page_label.configure(
            text=f"{result['files']} PDF files processed"
        )

        warning_text = ""
        if result["warnings"]:
            warning_text = "\n\nWarnings:\n" + "\n".join(result["warning_messages"][:10])
            if result["warnings"] > 10:
                warning_text += f"\n...and {result['warnings'] - 10} more."

        messagebox.showinfo(
            "Batch Completed",
            (
                f"PDF files processed: {result['files']}\n"
                f"Successful validations: {result['passed']}\n"
                f"Warnings: {result['warnings']}\n"
                f"Total rows: {result['rows']:,}\n\n"
                f"Output files: {len(result['outputs'])}\n"
                f"First output:\n{result['outputs'][0]}"
                f"{warning_text}"
            ),
        )
        self.set_busy(False)

    def conversion_failed(self, error_message):
        self.progress_bar.set(0)
        self.status_label.configure(text="Conversion failed")
        self.page_label.configure(text="")
        messagebox.showerror("Conversion Error", error_message)
        self.set_busy(False)

    def set_busy(self, busy):
        state = "disabled" if busy else "normal"
        self.select_button.configure(state=state)
        self.convert_button.configure(state=state)


if __name__ == "__main__":
    app = PDFConverterApp()
    app.mainloop()
