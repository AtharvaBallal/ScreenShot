import tkinter as tk
from tkinter import filedialog
import pyautogui
from PIL import ImageGrab, Image
import io
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from docx import Document
from docx.shared import Inches


class ScreenshotApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Screenshot Tool")
        self.root.attributes('-topmost', True)  # Make the window stay on top

        self.screenshots = []

        # Status label
        self.status_label = tk.Label(root, text="Status: Ready")
        self.status_label.grid(row=0, column=0, columnspan=2, pady=10)

        # Screenshot and Reset buttons in the first row
        self.screenshot_btn = tk.Button(root, text="Screenshot", command=self.take_screenshot)
        self.screenshot_btn.grid(row=1, column=0, padx=10, pady=10)

        self.reset_btn = tk.Button(root, text="Reset", command=self.reset_screenshots)
        self.reset_btn.grid(row=1, column=1, padx=10, pady=10)

        # Save to Excel and Save to Word buttons in the second row
        self.save_to_excel_btn = tk.Button(root, text="Save to Excel", command=self.save_to_excel)
        self.save_to_excel_btn.grid(row=2, column=0, padx=10, pady=10)

        self.save_to_word_btn = tk.Button(root, text="Save to Word", command=self.save_to_word)
        self.save_to_word_btn.grid(row=2, column=1, padx=10, pady=10)

    def take_screenshot(self):
        screenshot = pyautogui.screenshot()
        self.screenshots.append(screenshot)
        self.status_label.config(text=f"Status: Screenshot {len(self.screenshots)} taken")

    def save_to_excel(self):
        if not self.screenshots:
            self.status_label.config(text="Status: No screenshots to save")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Screenshots"

        # Start saving images from row 4
        row_offset = 4
        for idx, screenshot in enumerate(self.screenshots):
            img_io = io.BytesIO()
            screenshot.save(img_io, format="PNG")
            img_io.seek(0)
            img = ExcelImage(img_io)

            # Anchor images to specific rows
            img.anchor = f"A{row_offset}"
            ws.add_image(img)

            # Increase row height for visibility
            ws.row_dimensions[row_offset].height = screenshot.height // 15
            ws.column_dimensions['A'].width = screenshot.width // 15

            # Move to the next row with a gap
            row_offset += int(screenshot.height / 15) + 5

        wb.save(file_path)
        self.status_label.config(text=f"Status: Screenshots saved to {file_path}")

    def save_to_word(self):
        if not self.screenshots:
            self.status_label.config(text="Status: No screenshots to save")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                                 filetypes=[("Word files", "*.docx")])
        if not file_path:
            return

        doc = Document()
        # Add the heading without "Screenshot"
        doc.add_heading("", level=1)

        for idx, screenshot in enumerate(self.screenshots):
            img_io = io.BytesIO()
            screenshot.save(img_io, format="PNG")
            img_io.seek(0)
            screenshot_image = Image.open(img_io)


            img_io.seek(0)  # Reset the IO stream for Word
            doc.add_picture(img_io, width=Inches(6))

            # Add a blank line after the image
            doc.add_paragraph("")

        doc.save(file_path)
        self.status_label.config(text=f"Status: Screenshots saved to {file_path}")

    def reset_screenshots(self):
        self.screenshots = []  # Clear all stored screenshots
        self.status_label.config(text="Status: Screenshots reset")


if __name__ == "__main__":
    root = tk.Tk()
    app = ScreenshotApp(root)
    root.mainloop()
