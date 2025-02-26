import customtkinter as ctk
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook

# Import your generators
from ReportGenerator import DataExtractionApp
from LocateSheet import LocateSheetGenerator
from NewFraming import NewFramingGenerator
from SteelPole import SteelPoleGenerator
from LongLeadIS import LongLeadGenerator
from PullingSectionTracker import PullingSectionTracker

# Import mergeXML application
from mergeXML import XMLTagExtractorApp  # Ensure mergeXML.py is in the same directory or update the import path

class ModernApp:
    def __init__(self):
        self.app = ctk.CTk()
        self.app.title("Data Extraction Tool")
        self.app.geometry("1100x700")
        
        # Set the theme and color
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        # Initialize variables
        self.workbook = None
        self.file_path = None
        self.wo_number = None
        self.county = None
        self.city_place = None
        
        # Create the main container
        self.setup_gui()
        
    def setup_gui(self):
        # Create main frame
        self.main_frame = ctk.CTkFrame(self.app, corner_radius=0)
        self.main_frame.pack(fill="both", expand=True)
        
        # Title
        title_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        title_frame.pack(pady=20, padx=20, fill="x")
        
        ctk.CTkLabel(
            title_frame,
            text="Data Extraction Tool",
            font=ctk.CTkFont(size=24, weight="bold")
        ).pack()

        # Create left sidebar
        self.sidebar = ctk.CTkFrame(self.main_frame, width=200)
        self.sidebar.pack(side="left", fill="y", padx=20, pady=20)
        
        # File upload section in sidebar
        upload_label = ctk.CTkLabel(
            self.sidebar,
            text="File Upload",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        upload_label.pack(pady=(0, 10))
        
        self.upload_button = ctk.CTkButton(
            self.sidebar,
            text="Select Excel File",
            command=self.upload_file
        )
        self.upload_button.pack(pady=10, padx=20)
        
        # Status indicator
        self.status_label = ctk.CTkLabel(
            self.sidebar,
            text="No file loaded",
            font=ctk.CTkFont(size=12)
        )
        self.status_label.pack(pady=(10, 20))
        
        # Separator
        separator = ctk.CTkFrame(self.sidebar, height=2, fg_color=("gray70", "gray30"))
        separator.pack(fill="x", padx=15, pady=10)

        # Main content area
        self.content = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.content.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        
        # Options Frame
        options_frame = ctk.CTkFrame(self.content)
        options_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Options Label
        ctk.CTkLabel(
            options_frame,
            text="Automation Options",
            font=ctk.CTkFont(size=20, weight="bold")
        ).pack(pady=20)
        
        # Action Buttons with modern styling
        actions = [
            ("Generate New Framing Sheet", self.generate_new_framing_sheet),
            ("Generate Locate Sheet", self.generate_locate_sheet),
            ("Generate Steel Pole Information", self.generate_steel_pole_information),
            ("Generate Long Lead Item Sheet", self.generate_long_lead_sheet),
            ("Generate Pulling Section Tracker", self.generate_pulling_section_tracker)
        ]
        
        for text, command in actions:
            btn = ctk.CTkButton(
                options_frame,
                text=text,
                command=command,
                height=40,
                corner_radius=8
            )
            btn.pack(pady=10, padx=40, fill="x")

        self.launch_data_extraction_app_button = ctk.CTkButton(
            self.sidebar,
            text="Launch Data Extraction Tool",
            command=self.launch_data_extraction_app
        )
        self.launch_data_extraction_app_button.pack(pady=10, padx=20)

        # Add the new button for testing large XML
        self.test_large_xml_button = ctk.CTkButton(
            self.sidebar,
            text="Testing Large XML",
            command=self.launch_merge_xml_app
        )
        self.test_large_xml_button.pack(pady=10, padx=20)

    def launch_merge_xml_app(self):
        """Launches the Merge XML Tool in a new window."""
        new_window = tk.Toplevel(self.app)
        new_window.title("Large XML Testing")
        new_window.geometry("800x600")
        
        # Make window modal
        new_window.transient(self.app)
        new_window.grab_set()
        
        # Center the window relative to parent
        x = self.app.winfo_x() + (self.app.winfo_width() // 2) - 400
        y = self.app.winfo_y() + (self.app.winfo_height() // 2) - 300
        new_window.geometry(f"+{x}+{y}")
        
        # Initialize the XML extractor app
        xml_app = XMLTagExtractorApp(new_window)
        
        # Wait for this window to be closed before allowing interaction with parent
        self.app.wait_window(new_window)

    def launch_data_extraction_app(self):
        """Launches the Data Extraction Report App in a new window."""
        new_window = tk.Toplevel(self.app)
        new_window.title("Data Extraction Report Generator")
        new_window.geometry("1100x700")  # More reasonable initial size
        
        # Make window modal
        new_window.transient(self.app)
        new_window.grab_set()
        
        # Center the window relative to parent
        x = self.app.winfo_x() + (self.app.winfo_width() // 2) - 550
        y = self.app.winfo_y() + (self.app.winfo_height() // 2) - 350
        new_window.geometry(f"+{x}+{y}")
        
        # Initialize the data extraction app
        data_app = DataExtractionApp(new_window)
        
        # Ensure this window stays on top of parent
        new_window.lift()
        new_window.focus_force()
        
        # Wait for this window to be closed before allowing interaction with parent
        self.app.wait_window(new_window)

    def create_dialog(self, title, message):
        """Create a modern custom dialog"""
        dialog = ctk.CTkInputDialog(
            text=message,
            title=title
        )
        return dialog.get_input()

    def upload_file(self):
        """Handle file upload"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.process_upload(file_path)

    def process_upload(self, file_path):
        """Process the uploaded file"""
        try:
            self.workbook = load_workbook(file_path)
            self.file_path = file_path
            
            # Get the file name for display
            file_name = os.path.basename(file_path)
            
            # Collect additional information
            self.wo_number = self.create_dialog("Work Order Information", "Please enter the WO#:")
            if self.wo_number:
                self.county = self.create_dialog("Location Information", "Please enter the County:")
                if self.county:
                    self.city_place = self.create_dialog("Location Information", "Please enter the City/Place:")
                    if self.city_place:
                        self.status_label.configure(text=f"Loaded: {file_name}")
                        self.show_success_toast("File uploaded successfully!")
                        return
                        
            # If we get here, user cancelled one of the dialogs
            self.reset_state()
            self.status_label.configure(text="Upload cancelled")
            
        except Exception as e:
            self.reset_state()
            self.show_error("Failed to upload file", str(e))
            self.status_label.configure(text="Upload failed")

    def reset_state(self):
        """Resets the application state"""
        self.workbook = None
        self.file_path = None
        self.wo_number = None
        self.county = None
        self.city_place = None

    def check_file_loaded(self):
        """Checks if a file is loaded and required information is present"""
        if not all([self.workbook, self.wo_number, self.county, self.city_place]):
            self.show_error("Error", "Please upload an Excel file and provide all required information first.")
            return False
        return True

    def show_success_toast(self, message):
        """Show a modern success toast notification"""
        toast = ctk.CTkToplevel()
        toast.geometry("300x100")
        toast.title("")
        toast.attributes('-topmost', True)
        
        # Center the toast
        x = self.app.winfo_x() + (self.app.winfo_width() // 2) - 150
        y = self.app.winfo_y() + (self.app.winfo_height() // 2) - 50
        toast.geometry(f"+{x}+{y}")
        
        ctk.CTkLabel(
            toast,
            text="✓",
            font=ctk.CTkFont(size=30)
        ).pack(pady=(10, 0))
        
        ctk.CTkLabel(
            toast,
            text=message,
            font=ctk.CTkFont(size=14)
        ).pack(pady=5)
        
        toast.after(2000, toast.destroy)

    def show_error(self, title, message):
        """Show error dialog"""
        messagebox.showerror(title, message)

    def generate_new_framing_sheet(self):
        """Handles new framing sheet generation"""
        if not self.check_file_loaded():
            return
            
        try:
            self.status_label.configure(text="Generating framing sheet...")
            generator = NewFramingGenerator(self.workbook)
            generator.generate_sheet()
            self.status_label.configure(text="Framing sheet generated")
            self.show_success_toast("Framing sheet generated successfully!")
        except Exception as e:
            self.show_error("Error", f"Failed to generate framing sheet: {str(e)}")
            self.status_label.configure(text="Generation failed")

    def generate_locate_sheet(self):
        """Handles locate sheet generation"""
        if not self.check_file_loaded():
            return
            
        try:
            self.status_label.configure(text="Generating locate sheet...")
            generator = LocateSheetGenerator(
                self.workbook,
                self.wo_number,
                self.county,
                self.city_place
            )
            generator.generate_sheet()
            self.status_label.configure(text="Locate sheet generated")
            self.show_success_toast("Locate sheet generated successfully!")
        except Exception as e:
            self.show_error("Error", f"Failed to generate locate sheet: {str(e)}")
            self.status_label.configure(text="Generation failed")

    def generate_steel_pole_information(self):
        """Handles steel pole information generation"""
        if not self.check_file_loaded():
            return
            
        try:
            self.status_label.configure(text="Generating steel pole info...")
            generator = SteelPoleGenerator(self.workbook)
            generator.generate_sheet()
            self.status_label.configure(text="Steel pole info generated")
            self.show_success_toast("Steel pole information generated successfully!")
        except Exception as e:
            self.show_error("Error", f"Failed to generate steel pole information: {str(e)}")
            self.status_label.configure(text="Generation failed")

    def generate_long_lead_sheet(self):
        """Handles long lead item sheet generation"""
        if not self.check_file_loaded():
            return
            
        try:
            self.status_label.configure(text="Generating long lead item sheet...")
            generator = LongLeadGenerator(
                self.workbook,
                self.wo_number,
                self.county,
                self.city_place
            )
            output_path = generator.generate_sheet()
            self.status_label.configure(text="Long lead item sheet generated")
            self.show_success_toast(f"Long lead item sheet generated successfully!\nSaved to: {output_path}")
        except Exception as e:
            self.show_error("Error", f"Failed to generate long lead item sheet: {str(e)}")
            self.status_label.configure(text="Generation failed")

    def generate_pulling_section_tracker(self):
        """Handles pulling section tracker generation"""
        if not self.check_file_loaded():
            return
            
        try:
            self.status_label.configure(text="Generating pulling section tracker...")
            tracker = PullingSectionTracker(
                self.workbook,
                self.wo_number,
                self.county,
                self.city_place
            )
            output_path = tracker.generate_sheet()
            self.status_label.configure(text="Pulling section tracker generated")
            self.show_success_toast(f"Pulling section tracker generated successfully!\nSaved to: {output_path}")
        except Exception as e:
            self.show_error("Error", f"Failed to generate pulling section tracker: {str(e)}")
            self.status_label.configure(text="Generation failed")

    def run(self):
        self.app.mainloop()

if __name__ == "__main__":
    app = ModernApp()
    app.run()
