import threading
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from converter import extract_pdf_to_excel

import converter
print("USING CONVERTER:", converter.__file__)

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class PDFConverterApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("Finance PDF → Excel")
        self.geometry("650x500")
        self.resizable(False, False)

        self.pdf_path = None

        self.build_ui()

    def build_ui(self):

        # --------------------------------
        # Main container
        # --------------------------------

        self.container = ctk.CTkFrame(
            self,
            corner_radius=20,
        )

        self.container.pack(
            fill="both",
            expand=True,
            padx=30,
            pady=30,
        )

        # --------------------------------
        # Title
        # --------------------------------

        self.title_label = ctk.CTkLabel(
            self.container,
            text="Finance PDF → Excel",
            font=ctk.CTkFont(
                size=28,
                weight="bold",
            ),
        )

        self.title_label.pack(
            pady=(35, 5)
        )

        self.subtitle_label = ctk.CTkLabel(
            self.container,
            text="Settlement Report Converter",
            font=ctk.CTkFont(
                size=14
            ),
        )

        self.subtitle_label.pack(
            pady=(0, 30)
        )

        # --------------------------------
        # File label
        # --------------------------------

        self.file_label = ctk.CTkLabel(
            self.container,
            text="No PDF selected",
            font=ctk.CTkFont(
                size=14
            ),
        )

        self.file_label.pack(
            pady=10
        )

        # --------------------------------
        # Select button
        # --------------------------------

        self.select_button = ctk.CTkButton(
            self.container,
            text="Select PDF",
            width=220,
            height=40,
            corner_radius=10,
            command=self.select_pdf,
        )

        self.select_button.pack(
            pady=10
        )

        # --------------------------------
        # Convert button
        # --------------------------------

        self.convert_button = ctk.CTkButton(
            self.container,
            text="Convert to Excel",
            width=220,
            height=40,
            corner_radius=10,
            command=self.convert_pdf,
            state="disabled",
        )

        self.convert_button.pack(
            pady=10
        )

        # --------------------------------
        # Status
        # --------------------------------

        self.status_label = ctk.CTkLabel(
            self.container,
            text="",
            font=ctk.CTkFont(
                size=13
            ),
        )

        self.status_label.pack(
            pady=(20, 5)
        )

        # --------------------------------
        # Progress bar
        # --------------------------------

        self.progress_bar = ctk.CTkProgressBar(
            self.container,
            width=450,
            height=12,
            corner_radius=6,
        )

        self.progress_bar.set(0)

        self.progress_bar.pack(
            pady=10
        )

        # --------------------------------
        # Page counter
        # --------------------------------

        self.page_label = ctk.CTkLabel(
            self.container,
            text="",
            font=ctk.CTkFont(
                size=12
            ),
        )

        self.page_label.pack(
            pady=5
        )

    def select_pdf(self):

        path = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[
                ("PDF files", "*.pdf")
            ],
        )

        if not path:
            return

        self.pdf_path = Path(path)

        self.file_label.configure(
            text=self.pdf_path.name
        )

        self.convert_button.configure(
            state="normal"
        )

        self.status_label.configure(
            text=""
        )

        self.page_label.configure(
            text=""
        )

        self.progress_bar.set(0)

    def update_progress(
        self,
        page,
        total_pages,
        status,
    ):
        """
        Update the GUI safely from the
        background processing thread.
        """

        self.after(
            0,
            self._update_progress,
            page,
            total_pages,
            status,
        )

    def _update_progress(
        self,
        page,
        total_pages,
        status,
    ):

        progress = page / total_pages

        self.progress_bar.set(
            progress
        )

        self.status_label.configure(
            text=status
        )

        self.page_label.configure(
            text=(
                f"Page {page} / "
                f"{total_pages}"
            )
        )

    def convert_pdf(self):

        if not self.pdf_path:
            return

        output_path = (
            self.pdf_path.with_suffix(".xlsx")
        )

        self.select_button.configure(
            state="disabled"
        )

        self.convert_button.configure(
            state="disabled"
        )

        self.progress_bar.set(0)

        self.status_label.configure(
            text="Starting..."
        )

        self.page_label.configure(
            text=""
        )

        # Run conversion in background.
        thread = threading.Thread(
            target=self.run_conversion,
            args=(output_path,),
            daemon=True,
        )

        thread.start()

    def run_conversion(
        self,
        output_path,
    ):

        try:

            result = extract_pdf_to_excel(
                str(self.pdf_path),
                str(output_path),
                progress_callback=self.update_progress,
            )

            self.after(
                0,
                self.conversion_finished,
                result,
            )

        except Exception as exc:

            self.after(
                0,
                self.conversion_failed,
                str(exc),
            )

    def conversion_finished(
        self,
        result,
    ):

        self.progress_bar.set(1)

        self.status_label.configure(
            text="Conversion completed ✓"
        )

        self.page_label.configure(
            text=(
                f"{result['pages']} pages processed"
            )
        )

        if result["purchase_valid"]:

            validation_message = (
                "Purchase validation passed ✓\n\n"
                f"Purchase rows: "
                f"{result['purchase_rows']:,}\n"
                f"Expected: "
                f"{result['expected_purchase_rows']:,}"
            )

        else:

            validation_message = (
                "Purchase validation warning\n\n"
                f"Purchase rows found: "
                f"{result['purchase_rows']:,}\n"
                f"Expected: "
                f"{result['expected_purchase_rows']:,}"
            )

        messagebox.showinfo(
            "Conversion Completed",
            (
                f"Excel file created successfully.\n\n"
                f"Pages: {result['pages']}\n"
                f"Sections: {result['sections']}\n"
                f"Total rows: {result['rows']:,}\n\n"
                f"{validation_message}\n\n"
                f"File:\n{result['output']}"
            ),
        )

        self.select_button.configure(
            state="normal"
        )

        self.convert_button.configure(
            state="normal"
        )

    def conversion_failed(
        self,
        error_message,
    ):

        self.progress_bar.set(0)

        self.status_label.configure(
            text="Conversion failed"
        )

        self.page_label.configure(
            text=""
        )

        messagebox.showerror(
            "Conversion Error",
            error_message,
        )

        self.select_button.configure(
            state="normal"
        )

        self.convert_button.configure(
            state="normal"
        )


if __name__ == "__main__":
    app = PDFConverterApp()
    app.mainloop()