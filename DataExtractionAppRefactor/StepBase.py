import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from abc import ABC, abstractmethod

class StepBase(ABC):
    def __init__(self):
        self.file_path = None
        self.data = {}
        self.columns = []
        self.selected_columns = []
        self.output_data = []
        self.max_force_data = {}
        self.soil_class_data = {}

    @abstractmethod
    def setup_widgets(self, parent_frame):
        """Set up the widgets specific to this step"""
        pass

    @abstractmethod
    def save_data(self):
        """Save the data collected in this step"""
        pass

    def create_upload_widgets(self, parent_frame, button_text, file_types):
        """Create standard file upload widgets"""
        upload_btn = ttk.Button(parent_frame, text=button_text,
                              command=lambda: self.upload_file(file_types))
        upload_btn.pack(pady=10)

    def create_paste_widgets(self, parent_frame):
        """Create standard paste text widgets"""
        self.paste_label = ttk.Label(parent_frame, 
                                   text="Paste Data Here", 
                                   font=("Arial", 14))
        self.paste_label.pack(pady=10)
        
        self.text_frame = ttk.Frame(parent_frame)
        self.text_frame.pack(pady=10)
        
        self.paste_text = tk.Text(self.text_frame, 
                                wrap=tk.WORD, 
                                height=15, 
                                font=("Arial", 12))
        self.text_scrollbar = ttk.Scrollbar(self.text_frame, 
                                          orient=tk.VERTICAL, 
                                          command=self.paste_text.yview)
        
        self.paste_text.config(yscrollcommand=self.text_scrollbar.set)
        self.paste_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Hide by default
        self.text_frame.pack_forget()
        self.paste_label.pack_forget()

    def create_listbox_widgets(self, parent_frame):
        """Create standard listbox widgets"""
        self.listbox_frame = ttk.Frame(parent_frame)
        self.listbox_frame.pack(pady=10)
        self.column_listbox = tk.Listbox(self.listbox_frame, 
                                       selectmode=tk.MULTIPLE,
                                       font=("Arial", 12),
                                       width=50,
                                       height=10)
        self.column_listbox.pack(side=tk.LEFT, fill=tk.BOTH)
        self.scrollbar = ttk.Scrollbar(self.listbox_frame,
                                     orient=tk.VERTICAL,
                                     command=self.column_listbox.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.column_listbox.config(yscrollcommand=self.scrollbar.set)

    def upload_file(self, file_types):
        """Handle file upload"""
        self.file_path = filedialog.askopenfilename(filetypes=file_types)
        if self.file_path:
            try:
                self.process_file()
                messagebox.showinfo("Success", "File uploaded successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process file: {e}")
    
    def next_step(self):
        """Navigate to next step"""
        if hasattr(self, 'app'):
            self.app.next_step()

    @abstractmethod
    def process_file(self):
        """Process the uploaded file"""
        pass