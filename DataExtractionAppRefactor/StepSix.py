from StepBase import StepBase
import tkinter as tk
from tkinter import ttk, messagebox
import xml.etree.ElementTree as ET

class StepSix(StepBase):
    def __init__(self):
        super().__init__()
        self.step = 6
        self.output_data = []

    def setup_widgets(self, parent_frame):
        # Create step label
        self.step_label = ttk.Label(parent_frame, 
                                  text="Step 6: Upload your Structure Usage Report", 
                                  font=("Arial", 14))
        self.step_label.pack(pady=10)

        # Create file upload widgets
        self.create_upload_widgets(parent_frame, 
                                 "Upload Structure Usage Report",
                                 [("XML files", "*.xml")])

        # Add process and skip buttons
        self.process_btn = ttk.Button(parent_frame, 
                                    text="Parse Data and Move to Next Step",
                                    command=self.save_data)
        self.process_btn.pack(pady=10)

        self.skip_btn = ttk.Button(parent_frame, 
                                 text="Skip This Step",
                                 command=self.next_step)
        self.skip_btn.pack(pady=10)

    def process_file(self):
        try:
            tree = ET.parse(self.file_path)
            root = tree.getroot()
            self.output_data = []

            # Extract data from XML for guy wire and cable elements
            for report in root.findall('.//summary_of_maximum_element_usages_for_structure_range'):
                seq_no = report.find('str_no').text
                element_label = report.find('element_label').text
                element_type = report.find('element_type').text
                max_usage = report.find('maximum_usage').text
                
                if element_type == "Guy" or element_type == "Cable":
                    self.output_data.append((
                        seq_no, 
                        element_label, 
                        element_type, 
                        max_usage
                    ))

            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process XML file: {e}")
            return False

    def save_data(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please upload a file first!")
            return False

        if not self.process_file():
            return False

        try:
            # Save extracted data to formatted text file
            with open("extractGuyUsage_seq_elementType_usage.txt", "w") as file:
                # Calculate max lengths for formatting
                max_lengths = {
                    'sequence': max(len(row[0]) for row in self.output_data),
                    'element_label': max(len(row[1]) for row in self.output_data),
                    'element_type': max(len(row[2]) for row in self.output_data),
                    'max_usage': max(len(row[3]) for row in self.output_data)
                }

                headers = [
                    ("Sequence #", max_lengths['sequence']),
                    ("Element Label", max_lengths['element_label']),
                    ("Element Type", max_lengths['element_type']),
                    ("Maximum Usage", max_lengths['max_usage'])
                ]

                # Write headers
                header_row = " | ".join(f"{header[0]:<{header[1]}}" for header in headers)
                file.write(header_row + "\n")
                file.write("-" * len(header_row) + "\n")

                # Write data rows
                for row in self.output_data:
                    formatted_row = [
                        f"{row[0]:<{max_lengths['sequence']}}",
                        f"{row[1]:<{max_lengths['element_label']}}",
                        f"{row[2]:<{max_lengths['element_type']}}",
                        f"{row[3]:<{max_lengths['max_usage']}}"
                    ]
                    file.write(" | ".join(formatted_row) + "\n")

            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")
            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {e}")
            return False