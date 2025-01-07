from StepBase import StepBase
import tkinter as tk
from tkinter import ttk, messagebox
import xml.etree.ElementTree as ET
import re

class StepFour(StepBase):
    def __init__(self):
        super().__init__()
        self.step = 4
        self.output_data = []

    def setup_widgets(self, parent_frame):
        # Create step label
        self.step_label = ttk.Label(parent_frame, 
                                  text="Step 4: Copy and Paste your Stringing Chart - Neutral and Span Guy", 
                                  font=("Arial", 14))
        self.step_label.pack(pady=10)

        # Create paste widgets
        self.create_paste_widgets(parent_frame)
        self.paste_label.config(text="Paste Stringing Chart Data Here")
        self.paste_label.pack(pady=10)
        self.text_frame.pack(pady=10)

        # Create file upload button for XML
        self.upload_btn = ttk.Button(parent_frame, 
                                   text="Upload Span Guy XML",
                                   command=lambda: self.upload_file([("XML files", "*.xml")]))
        self.upload_btn.pack(pady=10)

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

            for section in root.findall('.//section_sagging_data'):
                circuit_type = section.find('circuit').text.strip()
                section_num = section.find('sec_no').text.strip()
                start_seq = section.find('from_str').text.strip()
                end_seq = section.find('to_str').text.strip()
                ruling_span = section.find('ruling_span').text.strip()

                sequences = f"{start_seq} - {end_seq}"
                total_span_length = float(ruling_span) if ruling_span else 0.0
                self.output_data.append((
                    section_num, 
                    sequences, 
                    total_span_length, 
                    circuit_type
                ))

            self.output_data.sort(key=lambda x: int(x[0]))
            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse XML file: {e}")
            return False

    def _parse_pasted_data(self):
        pasted_data = self.paste_text.get("1.0", tk.END).strip()
        if not pasted_data:
            messagebox.showerror("Error", "Please paste data into the text box.")
            return False

        try:
            sections = re.findall(
                r"Stringing Chart Report\n\nCircuit '(.*?)' Section #(.*?) "
                r"from structure #(.*?) to structure #(.*?),.*?Span\n(.*?)\n\n", 
                pasted_data, 
                re.DOTALL
            )
            
            self.output_data = []

            for section in sections:
                circuit_type, section_num, start_seq, end_seq, spans_data = section
                if "Span Guy" in circuit_type:
                    continue

                spans = re.findall(r"\n\s+(\d+\.\d+)\s+", spans_data)
                total_span_length = sum(map(float, spans))
                sequences = f"{start_seq} - {end_seq}"
                self.output_data.append((
                    section_num, 
                    sequences, 
                    total_span_length, 
                    circuit_type
                ))

            messagebox.showinfo("Success", "Pasted data parsed successfully. Please upload the Span Guy XML file.")
            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse pasted data: {e}")
            return False

    def save_data(self):
        if not self._parse_pasted_data():
            return

        try:
            with open("extractStringingChartNeutralSpan_section_seq_totalSpanLength_circuitType.txt", "w") as file:
                # Calculate max lengths for formatting
                max_lengths = {
                    'section_num': max(len(str(row[0])) for row in self.output_data),
                    'sequences': max(len(row[1]) for row in self.output_data),
                    'total_span_length': max(len(f"{row[2]:.2f}") for row in self.output_data),
                    'circuit_type': max(len(row[3]) for row in self.output_data)
                }

                headers = [
                    ("Section #", max_lengths['section_num']),
                    ("Sequence #s", max_lengths['sequences']),
                    ("Total Span Length", max_lengths['total_span_length']),
                    ("Circuit Type", max_lengths['circuit_type'])
                ]

                # Write headers
                header_row = " | ".join(f"{header[0]:<{header[1]}}" for header in headers)
                file.write(header_row + "\n")
                file.write("-" * len(header_row) + "\n")

                # Write data rows
                for row in self.output_data:
                    formatted_row = [
                        f"{row[0]:<{max_lengths['section_num']}}",
                        f"{row[1]:<{max_lengths['sequences']}}",
                        f"{row[2]:<{max_lengths['total_span_length']}.2f}",
                        f"{row[3]:<{max_lengths['circuit_type']}}"
                    ]
                    file.write(" | ".join(formatted_row) + "\n")

            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {e}")