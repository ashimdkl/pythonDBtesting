from StepBase import StepBase
import tkinter as tk
from tkinter import ttk, messagebox
import xml.etree.ElementTree as ET
import re

class StepSeven(StepBase):
    def __init__(self):
        super().__init__()
        self.step = 7
        self.soil_class_data = {}
        self.max_force_data = {}

    def setup_widgets(self, parent_frame):
        # Store parent frame reference
        self.parent_frame = parent_frame
        
        # Create step label
        self.step_label = ttk.Label(parent_frame, 
                                  text="Step 7: Copy and Paste Soil Class Data", 
                                  font=("Arial", 14))
        self.step_label.pack(pady=10)

        # Create paste widgets
        self.create_paste_widgets(parent_frame)
        self.paste_label.config(text="Paste Soil Class Data Here")
        self.paste_label.pack(pady=10)
        self.text_frame.pack(pady=10)

        # Create XML upload button
        self.upload_btn = ttk.Button(parent_frame, 
                                   text="Upload Joint Support XML and Parse",
                                   command=lambda: self.upload_file([("XML files", "*.xml")]))
        self.upload_btn.pack(pady=10)

        # Process button for soil class data
        self.process_btn = ttk.Button(parent_frame, 
                                    text="Parse Soil Class Data",
                                    command=self._parse_soil_class_data)
        self.process_btn.pack(pady=10)

        # Generate Report button
        self.gen_report_btn = ttk.Button(parent_frame,
                                       text="Generate Report",
                                       command=self.generate_report)
        self.gen_report_btn.pack(pady=10)

    def process_file(self):
        try:
            tree = ET.parse(self.file_path)
            root = tree.getroot()
            
            # Process XML data
            for report in root.findall('.//summary_of_joint_support_reactions_for_all_load_cases_for_structure_range'):
                seq_no = report.find('str_no').text
                shear_force = float(report.find('shear_force').text)
                bending_moment = float(report.find('bending_moment').text)
                max_force = max(shear_force, bending_moment)

                if seq_no not in self.max_force_data:
                    self.max_force_data[seq_no] = max_force
                else:
                    self.max_force_data[seq_no] = max(self.max_force_data[seq_no], max_force)

            # Save initial data
            self._save_initial_max_force_data()
            messagebox.showinfo("Success", "Joint support data from XML processed successfully.")
            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process XML file: {e}")
            return False

    def _save_initial_max_force_data(self):
        try:
            with open("extractMAX_sequence_MaxForce.txt", "w") as file:
                max_lengths = {
                    'sequence': max(len(seq) for seq in self.max_force_data),
                    'max_force': max(len(f"{force:.2f}") for force in self.max_force_data.values())
                }

                file.write("Sequence | Max Force | Soil Class\n")
                file.write("-" * 40 + "\n")

                for seq, force in sorted(self.max_force_data.items()):
                    row = [
                        f"{seq:<{max_lengths['sequence']}}",
                        f"{force:<{max_lengths['max_force']}.2f}",
                        "N/A"  # Initial soil class placeholder
                    ]
                    file.write(" | ".join(row) + "\n")

            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save initial max force data: {e}")
            return False

    def _parse_soil_class_data(self):
        pasted_data = self.paste_text.get("1.0", tk.END).strip()
        if not pasted_data:
            messagebox.showerror("Error", "Please paste soil class data into the text box.")
            return False

        try:
            lines = pasted_data.split('\n')
            pattern = r'(\d+)(?:-(\d+))?\s+(\d+)\s+(.*)'
            
            for line in lines:
                match = re.match(pattern, line.strip())
                if match:
                    start_seq, end_seq, soil_class, description = match.groups()
                    start_seq = int(start_seq)
                    end_seq = int(end_seq) if end_seq else start_seq

                    for seq in range(start_seq, end_seq + 1):
                        self.soil_class_data[seq] = {
                            'soil_class': soil_class, 
                            'description': description
                        }

            self._update_max_force_file()
            messagebox.showinfo("Success", "Soil class data processed and saved successfully.")
            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse soil class data: {e}")
            return False

    def _update_max_force_file(self):
        try:
            with open("extractMAX_sequence_MaxForce.txt", "r") as file:
                lines = file.readlines()

            updated_lines = [
                lines[0],
                "Sequence | Max Force | Soil Class\n",
                "-" * 40 + "\n"
            ]

            for line in lines[2:]:
                parts = line.strip().split('|')
                sequence = int(parts[0].strip())
                max_force = parts[1].strip()
                soil_class = self.soil_class_data.get(sequence, {}).get('soil_class', 'N/A')
                updated_lines.append(f"{sequence:4d} | {max_force:8} | {soil_class}\n")

            with open("extractMAX_sequence_MaxForce.txt", "w") as file:
                file.writelines(updated_lines)

            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to update max force file: {e}")
            return False

    def save_data(self):
        return True

    def generate_report(self):
        """Generate final report"""
        try:
            if hasattr(self, 'app'):
                self.app.generate_report()
            else:
                messagebox.showerror("Error", "Application reference not found.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {e}")