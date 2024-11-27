import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import pyautogui
import io
from openpyxl.drawing.image import Image as ExcelImage
from docx import Document



class ScreenshotApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SnapIt")
        self.root.attributes('-topmost', True)  # Make the window stay on top

        self.screenshots = []

        # Status label
        self.status_label = tk.Label(root, text="Status: Ready.")
        self.status_label.grid(row=0, column=0, columnspan=4, pady=5)

        # Screenshot and Reset buttons in the first row
        self.screenshot_btn = tk.Button(root, text="Screenshot", command=self.take_screenshot)
        self.screenshot_btn.grid(row=1, column=0, padx=5, pady=5)

        # Save to Excel and Save to Word buttons in the second row
        self.save_to_excel_btn = tk.Button(root, text="Save to Excel", command=self.save_to_excel)
        self.save_to_excel_btn.grid(row=1, column=1, padx=5, pady=5)

        self.save_existing_excel_btn = tk.Button(root, text="Save to Existing Excel", command=self.save_to_existing_excel)
        self.save_existing_excel_btn.grid(row=1, column=2, padx=5, pady=5)

        self.save_to_word_btn = tk.Button(root, text="Save to Word", command=self.save_to_word)
        self.save_to_word_btn.grid(row=1, column=3, padx=5, pady=5)

        self.reset_btn = tk.Button(root, text="Reset", command=self.reset_screenshots)
        self.reset_btn.grid(row=1, column=4, padx=5, pady=5)

    def take_screenshot(self):
        # Temporarily hide the window
        self.root.attributes('-alpha', 0)  # Make the window completely transparent

        # Take the screenshot
        screenshot = pyautogui.screenshot()

        # Restore the window
        self.root.attributes('-alpha', 1)  # Restore visibility

        # Store the screenshot
        self.screenshots.append(screenshot)
        self.status_label.config(text=f"Status: Screenshot {len(self.screenshots)} taken")

    def save_to_excel(self):
        if not self.screenshots:
            self.status_label.config(text="Status: No screenshots to save.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Screenshots"
            row_offset = 4

            for idx, screenshot in enumerate(self.screenshots):
                img_io = io.BytesIO()
                screenshot.save(img_io, format="PNG")
                img_io.seek(0)
                img = ExcelImage(img_io)
                img.anchor = f"A{row_offset}"
                ws.add_image(img)
                ws.row_dimensions[row_offset].height = screenshot.height // 15
                ws.column_dimensions['A'].width = screenshot.width // 15
                row_offset += int(screenshot.height / 15)

            wb.save(file_path)
            self.status_label.config(text=f"Status: Excel file saved!")
        except PermissionError:
            messagebox.showerror("Error", "The Excel file is open. Please close it and try again.")

    def save_to_existing_excel(self):
        if not self.screenshots:
            self.status_label.config(text="No screenshots to save.")
            return

        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        try:
            try:
                wb = openpyxl.load_workbook(file_path)
            except Exception as e:
                self.status_label.config(text="Failed to load the Excel file.")
                return

            sheet_name = f"Screenshot{len(wb.sheetnames) + 1}"
            ws = wb.create_sheet(title=sheet_name)

            row_offset = 4
            for idx, screenshot in enumerate(self.screenshots):
                img_io = io.BytesIO()
                screenshot.save(img_io, format="PNG")
                img_io.seek(0)
                img = ExcelImage(img_io)
                img.anchor = f"A{row_offset}"
                ws.add_image(img)
                ws.row_dimensions[row_offset].height = screenshot.height // 15
                ws.column_dimensions['A'].width = screenshot.width // 15
                row_offset += int(screenshot.height / 15)

            wb.save(file_path)
            self.status_label.config(text=f"Screenshots saved to existing {sheet_name} in {file_path}!")
        except PermissionError:
            messagebox.showerror("Error", "The Excel file is open. Please close it and try again.")

    def save_to_word(self):
        if not self.screenshots:
            self.status_label.config(text="Status: No screenshots to save.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                                 filetypes=[("Word files", "*.docx")])
        if not file_path:
            return

        doc = Document()
        # Add the heading without "Screenshot"
        doc.add_heading("", level=1)

        try:
            doc = Document()
            for idx, screenshot in enumerate(self.screenshots):
                img_io = io.BytesIO()
                screenshot.save(img_io, format="PNG")
                img_io.seek(0)
                doc.add_picture(img_io)
                doc.add_paragraph()  # Add space between screenshots

            doc.save(file_path)
            self.status_label.config(text=f"Status: Word file saved!")
        except PermissionError:
            messagebox.showerror("Error", "The Word file is open. Please close it and try again.")

    def reset_screenshots(self):
        self.screenshots = []  # Clear all stored screenshots
        self.status_label.config(text="Status: Screenshots reset.")


if __name__ == "__main__":
    root = tk.Tk()
    app = ScreenshotApp(root)
    root.mainloop()