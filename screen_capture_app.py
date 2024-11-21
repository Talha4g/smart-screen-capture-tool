import tkinter as tk
from tkinter import messagebox
from PIL import ImageGrab, Image, ImageTk
import sys
import os
import re
import subprocess

class ScreenCaptureApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Talha's Screen Capture App")
        self.master.geometry("300x200")
        self.center_window(self.master)

        # Check Tesseract installation first
        self.tesseract_path = self.find_tesseract()
        if not self.tesseract_path:
            messagebox.showerror("Error", "Tesseract OCR is not found. Please install Tesseract OCR and try again.")
            master.destroy()
            return

        # Import pytesseract only after confirming Tesseract installation
        try:
            global pytesseract
            import pytesseract
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
        except ImportError:
            messagebox.showerror("Error", "pytesseract module is not installed. Please install it using: pip install pytesseract")
            master.destroy()
            return

        self.start_x = None
        self.start_y = None
        self.rect_id = None

        # Create GUI elements
        self.create_gui()

    def find_tesseract(self):
        """Find Tesseract executable in common installation locations."""
        possible_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            r"/usr/bin/tesseract",
            r"/usr/local/bin/tesseract",
            os.path.join(os.getenv('ProgramFiles', ''), 'Tesseract-OCR', 'tesseract.exe'),
            os.path.join(os.getenv('ProgramFiles(x86)', ''), 'Tesseract-OCR', 'tesseract.exe'),
        ]

        # Check if tesseract is in PATH
        try:
            if sys.platform.startswith('win'):
                result = subprocess.run(['where', 'tesseract'], capture_output=True, text=True)
                if result.returncode == 0:
                    return result.stdout.strip().split('\n')[0]
            else:
                result = subprocess.run(['which', 'tesseract'], capture_output=True, text=True)
                if result.returncode == 0:
                    return result.stdout.strip()
        except:
            pass

        # Check common installation paths
        for path in possible_paths:
            if os.path.isfile(path):
                return path

        return None

    def create_gui(self):
        """Create the GUI elements."""
        # Start button for text extraction
        self.start_button = tk.Button(self.master, text="Start Drawing", command=self.activate_drawing)
        self.start_button.pack(pady=10)

        # Calculate Sum button
        self.calculate_sum_button = tk.Button(self.master, text="Calculate Sum", command=self.activate_calculation)
        self.calculate_sum_button.pack(pady=10)

        # Status label
        self.status_label = tk.Label(self.master, text="OCR Ready", fg="green")
        self.status_label.pack(pady=10)

        self.master.attributes("-topmost", True)

    def activate_drawing(self):
        """Activates the drawing mode for text extraction."""
        self.master.withdraw()
        self.draw_box_window(action="extract_text")

    def activate_calculation(self):
        """Activates the drawing mode for calculating sum."""
        self.master.withdraw()
        self.draw_box_window(action="calculate_sum")

    def draw_box_window(self, action):
        """Draws a box to capture an area and process based on the action."""
        self.action = action
        self.top = tk.Toplevel()
        self.top.attributes("-fullscreen", True)
        self.top.attributes("-alpha", 0.3)
        self.top.configure(cursor="cross")
        self.top.bind("<ButtonPress-1>", self.on_button_press)
        self.top.bind("<B1-Motion>", self.on_mouse_drag)
        self.top.bind("<ButtonRelease-1>", self.on_button_release)

        self.canvas = tk.Canvas(self.top, bg="black", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.configure(background="", highlightbackground="")

    def on_button_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        self.rect_id = None

    def on_mouse_drag(self, event):
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(self.start_x, self.start_y, event.x, event.y, outline="red", width=2)

    def on_button_release(self, event):
        end_x, end_y = event.x, event.y
        x1, x2 = min(self.start_x, end_x), max(self.start_x, end_x)
        y1, y2 = min(self.start_y, end_y), max(self.start_y, end_y)
        
        if self.action == "extract_text":
            self.extract_text(x1, y1, x2, y2)
        elif self.action == "calculate_sum":
            self.calculate_sum(x1, y1, x2, y2)
        
        self.top.destroy()

    def extract_text(self, x1, y1, x2, y2):
        """Extracts text from the selected area with error handling."""
        try:
            self.master.deiconify()
            image = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            
            try:
                text = pytesseract.image_to_string(image)
            except Exception as e:
                messagebox.showerror("OCR Error", f"Failed to process image: {str(e)}")
                return

            result_window = tk.Toplevel(self.master)
            result_window.title("Extracted Text")
            result_window.geometry("400x300")
            
            copy_button = tk.Button(result_window, text="Copy to Clipboard", 
                                  command=lambda: self.copy_to_clipboard(text))
            copy_button.pack(pady=5)

            text_display = tk.Text(result_window, wrap="word", font=("Helvetica", 12))
            text_display.insert("1.0", text)
            text_display.config(state="disabled")
            text_display.pack(expand=True, fill="both")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def calculate_sum(self, x1, y1, x2, y2):
        """Calculates sum with improved error handling and number detection."""
        try:
            self.master.deiconify()
            image = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            
            try:
                text = pytesseract.image_to_string(
                    image,
                    config='--psm 6 -c tessedit_char_whitelist=0123456789.'
                )
            except Exception as e:
                messagebox.showerror("OCR Error", f"Failed to process image: {str(e)}")
                return

            numbers = re.findall(r'-?\d*\.?\d+', text)
            if not numbers:
                messagebox.showinfo("No Numbers", "No numbers were found in the selected area.")
                return

            try:
                numbers = [float(n) for n in numbers]
                total_sum = sum(numbers)
            except ValueError as e:
                messagebox.showerror("Calculation Error", "Failed to process numbers.")
                return

            result_window = tk.Toplevel(self.master)
            result_window.title("Calculation Result")
            result_window.attributes("-topmost", True)

            numbers_text = "Numbers found:\n" + "\n".join([f"{n:,.2f}" for n in numbers])
            numbers_label = tk.Label(result_window, text=numbers_text, justify=tk.LEFT)
            numbers_label.pack(padx=20, pady=10)

            result_label = tk.Label(result_window, 
                                  text=f"Total Sum: {total_sum:,.2f}", 
                                  font=("Helvetica", 16, "bold"))
            result_label.pack(padx=20, pady=10)

            copy_button = tk.Button(result_window, text="Copy Sum to Clipboard",
                                  command=lambda: self.copy_to_clipboard(total_sum))
            copy_button.pack(pady=10)

            self.center_window(result_window, width=300, height=400)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def copy_to_clipboard(self, text):
        """Copy the provided text to the clipboard."""
        self.master.clipboard_clear()
        self.master.clipboard_append(str(text))
        self.master.update()

    def center_window(self, window, width=300, height=200):
        """Helper function to center a window on the screen."""
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")

def main():
    try:
        root = tk.Tk()
        app = ScreenCaptureApp(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Fatal Error", f"Application failed to start: {str(e)}")

if __name__ == "__main__":
    main()