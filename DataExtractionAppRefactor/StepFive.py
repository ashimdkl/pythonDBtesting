from StepBase import StepBase
import tkinter as tk
from tkinter import ttk, messagebox
import re

class StepFive(StepBase):
    def __init__(self):
        super().__init__()
        self.step = 5
        self.output_data = []

    def setup_widgets(self, parent_frame):
        # Create step label
        self.step_label = ttk.Label(parent_frame, 
                                  text="Step 5: Copy and Paste your Stringing Chart - Primary Conductor", 
                                  font=("Arial", 14))
        self.step_label.pack(pady=10)

        # Create paste widgets
        self.create_paste_widgets(parent_frame)
        self.paste_label.config(text="Paste Primary Conductor Stringing Chart Data Here")
        self.paste_label.pack(pady=10)
        self.text_frame.pack(pady=10)

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
        pass  # Not used in Step 5

    def save_data(self):
        if not self._parse_primary_conductor_data():
            return

        try:
            with open("extractStringingChartPrimary_section_struct_circuitType_spanLength_total.txt", "w") as file:
                max_lengths = {
                    'section_num': max(len(str(row[0])) for row in self.output_data),
                    'structure_to_structure': max(len(row[1]) for row in self.output_data),
                    'circuit_type': max(len(row[2]) for row in self.output_data),
                    'circuit_value': max(len(str(row[3])) for row in self.output_data),
                    'total_span_length': max(len(f"{row[4]:.2f}") for row in self.output_data),
                    'result': max(len(f"{row[5]:.2f}") for row in self.output_data),
                    'sequences': max(len(row[6]) for row in self.output_data),
                }

                headers = [
                    ("Section #", max_lengths['section_num']),
                    ("Structure -> Structure", max_lengths['structure_to_structure']),
                    ("Circuit Type", max_lengths['circuit_type']),
                    ("Circuit Value", max_lengths['circuit_value']),
                    ("Span Length", max_lengths['total_span_length']),
                    ("Result", max_lengths['result']),
                    ("Sequences", max_lengths['sequences'])
                ]

                # Write headers
                header_row = " | ".join(f"{header[0]:<{header[1]}}" for header in headers)
                file.write(header_row + "\n")
                file.write("-" * len(header_row) + "\n")

                # Write data rows
                for row in self.output_data:
                    formatted_row = [
                        f"{row[0]:<{max_lengths['section_num']}}",
                        f"{row[1]:<{max_lengths['structure_to_structure']}}",
                        f"{row[2]:<{max_lengths['circuit_type']}}",
                        f"{row[3]:<{max_lengths['circuit_value']}}",
                        f"{row[4]:<{max_lengths['total_span_length']}.2f}",
                        f"{row[5]:<{max_lengths['result']}.2f}",
                        f"{row[6]:<{max_lengths['sequences']}}"
                    ]
                    file.write(" | ".join(formatted_row) + "\n")

            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {e}")

    def _parse_primary_conductor_data(self):
        """Parse the primary conductor stringing chart data from pasted text"""
        pasted_data = self.paste_text.get("1.0", tk.END).strip()
        if not pasted_data:
            messagebox.showerror("Error", "Please paste data into the text box.")
            return False

        try:
            # Extract sections using regex pattern
            sections = re.findall(
                r"Stringing Chart Report\n\nCircuit '(.*?)' Section #(.*?) "
                r"from structure #(.*?) to structure #(.*?),.*?Span\n(.*?)\n\n", 
                pasted_data, 
                re.DOTALL
            )
            
            self.output_data = []

            # Process each section
            for section in sections:
                circuit_type, section_num, start_seq, end_seq, spans_data = section
                
                # Extract and calculate span data
                spans = re.findall(r"\n\s+(\d+\.\d+)\s+", spans_data)
                total_span_length = sum(map(float, spans))
                
                # Extract sequence numbers and create formatted strings
                sequences = re.findall(r"\d{4}", spans_data)
                sequences_str = ", ".join(sequences)
                structure_to_structure = f"{start_seq} -> {end_seq}"
                
                # Extract circuit value and calculate result
                circuit_value = int(re.search(r'(\d+)PH', circuit_type).group(1))
                result = total_span_length * circuit_value
                
                # Store processed data
                self.output_data.append((
                    section_num,
                    structure_to_structure,
                    circuit_type,
                    circuit_value,
                    total_span_length,
                    result,
                    sequences_str
                ))

            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse primary conductor data: {e}")
            return False