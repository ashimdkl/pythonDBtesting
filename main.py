import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
import xml.etree.ElementTree as ET
import math

class DataExtractionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Extraction App")
        self.root.geometry('800x600')

        self.file_path = None
        self.columns = []
        self.selected_columns = []
        self.step = 1
        self.output_data = []
        self.setup_gui()

    def setup_gui(self):
        self.intro_frame = tk.Frame(self.root, bg="white")
        self.intro_frame.pack(expand=True, fill="both")

        self.start_btn = tk.Button(self.intro_frame, text="Click to Begin Analysis!", command=self.start_analysis, font=("Arial", 14), bg="#f0f0f0")
        self.start_btn.pack(pady=20)

        self.main_frame = tk.Frame(self.root, bg="white")

        self.step_label = tk.Label(self.main_frame, text="Step 1: Upload your Hendrix Input Sheet (HIS)", font=("Arial", 14), bg="white")
        self.step_label.pack(pady=10)

        self.upload_btn = tk.Button(self.main_frame, text="Upload HIS Excel File", command=self.upload_file, font=("Arial", 14), bg="#f0f0f0")
        self.upload_btn.pack(pady=10)

        self.column_label = tk.Label(self.main_frame, text="Select Columns to Keep (including Sequence #)", font=("Arial", 14), bg="white")
        self.column_label.pack(pady=10)

        self.column_listbox = tk.Listbox(self.main_frame, selectmode=tk.MULTIPLE, font=("Arial", 12))
        self.column_listbox.pack(pady=10)

        self.process_btn = tk.Button(self.main_frame, text="Parse Data", command=self.parse_data, font=("Arial", 12), bg="#f0f0f0")
        self.process_btn.pack(pady=10)

        self.next_btn = tk.Button(self.main_frame, text="Next Step", command=self.next_step, font=("Arial", 12), bg="#f0f0f0")
        self.next_btn.pack(pady=10)

        self.paste_label = tk.Label(self.main_frame, text="Paste Fusing Coordination Data Here", font=("Arial", 14), bg="white")
        self.paste_text = tk.Text(self.main_frame, wrap=tk.WORD, height=15, font=("Arial", 12))
        self.paste_label.pack(pady=10)
        self.paste_text.pack(pady=10)
        self.paste_label.pack_forget()
        self.paste_text.pack_forget()

    def start_analysis(self):
        self.intro_frame.pack_forget()
        self.main_frame.pack(expand=True, fill="both")

    def upload_file(self):
        filetypes = [("Excel files", "*.xlsx *.xls")] if self.step == 1 else [("XML files", "*.xml"), ("Text files", "*.txt")]
        self.file_path = filedialog.askopenfilename(filetypes=filetypes)
        if self.file_path:
            if self.step == 1:
                self.load_columns_from_file()
            elif self.step == 3:
                self.parse_step3_xml()
            elif self.step == 4:
                self.parse_span_guy_xml()
            elif self.step == 6:
                self.parse_step6_structure_usage()
            messagebox.showinfo("File Uploaded", "File uploaded successfully.")

    def load_columns_from_file(self):
        try:
            workbook = load_workbook(self.file_path)
            sheet = workbook.active
            self.columns = [cell.value for cell in sheet[1]]
            self.column_listbox.delete(0, tk.END)
            for column in self.columns:
                self.column_listbox.insert(tk.END, column)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {e}")

    def parse_data(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please upload a file first!")
            return

        self.selected_columns = [self.column_listbox.get(i) for i in self.column_listbox.curselection()]
        if not self.selected_columns:
            messagebox.showerror("Error", "Please select at least one column to keep!")
            return

        try:
            if self.step == 1:
                workbook = load_workbook(self.file_path)
                sheet = workbook.active

                with open(f"step{self.step}.txt", "w") as file:
                    file.write("\t".join(self.selected_columns) + "\n")
                    for row in sheet.iter_rows(min_row=2, values_only=True):
                        if row[0] is not None:  # Assuming sequence number is the first column
                            row_data = [str(row[self.columns.index(col)]) for col in self.selected_columns]
                            file.write("\t".join(row_data) + "\n")

            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse file: {e}")

    def next_step(self):
        self.step += 1
        self.file_path = None
        self.columns = []
        self.selected_columns = []
        self.column_listbox.delete(0, tk.END)

        if self.step == 2:
            self.step_label.config(text="Step 2: Copy and Paste your Fusing Coordination Report")
            self.upload_btn.pack_forget()
            self.column_label.pack_forget()
            self.column_listbox.pack_forget()
            self.process_btn.config(text="Parse Pasted Data", command=self.parse_pasted_data)
            self.paste_label.pack(pady=10)
            self.paste_text.pack(pady=10)
            self.process_btn.pack(pady=10)
        elif self.step == 3:
            self.step_label.config(text="Step 3: Upload your Construction Staking Report")
            self.upload_btn.config(text="Upload Construction Staking Report", command=self.upload_file)
            self.paste_label.pack_forget()
            self.paste_text.pack_forget()
            self.upload_btn.pack(pady=10)
            self.column_listbox.pack_forget()
            self.process_btn.config(text="Parse Data", command=self.parse_step3_xml)
            self.process_btn.pack(pady=10)
        elif self.step == 4:
            self.step_label.config(text="Step 4: Copy and Paste your Stringing Chart - Neutral and Span Guy")
            self.upload_btn.pack_forget()
            self.paste_label.config(text="Paste Stringing Chart Data Here")
            self.paste_label.pack(pady=10)
            self.paste_text.pack(pady=10)
            self.process_btn.config(text="Parse and Continue", command=self.parse_and_continue_stringing_chart)
            self.process_btn.pack(pady=10)
        elif self.step == 5:
            self.step_label.config(text="Step 5: Copy and Paste your Stringing Chart - Primary Conductor")
            self.upload_btn.pack_forget()
            self.paste_label.config(text="Paste Primary Conductor Stringing Chart Data Here")
            self.paste_label.pack(pady=10)
            self.paste_text.pack(pady=10)
            self.process_btn.config(text="Parse Primary Conductor Data", command=self.parse_primary_conductor_data)
            self.process_btn.pack(pady=10)
        elif self.step == 6:
            self.step_label.config(text="Step 6: Upload your Structure Usage Report")
            self.upload_btn.config(text="Upload Structure Usage Report", command=self.upload_file)
            self.paste_label.pack_forget()
            self.paste_text.pack_forget()
            self.upload_btn.pack(pady=10)
            self.process_btn.config(text="Parse Data", command=self.parse_step6_structure_usage)
            self.process_btn.pack(pady=10)
        else:
            messagebox.showinfo("Completed", "All steps completed. Now you can merge and download the data.")

    def parse_pasted_data(self):
        pasted_data = self.paste_text.get("1.0", tk.END).strip()
        if not pasted_data:
            messagebox.showerror("Error", "Please paste data into the text box.")
            return

        try:
            lines = pasted_data.split("\n")
            header = lines[0].split("\t")
            data = []
            for line in lines[1:]:
                fields = line.split("\t")
                sequence = fields[0][:4]  # Assuming sequence is the first 4 characters
                existing = fields[2]  # Assuming existing value is the third field
                data.append([sequence, existing])

            with open(f"step{self.step}.txt", "w") as file:
                file.write("Sequence\tExisting\n")
                for row in data:
                    file.write("\t".join(row) + "\n")

            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse pasted data: {e}")

    def parse_step3_xml(self):
        try:
            tree = ET.parse(self.file_path)
            root = tree.getroot()
            data = {}
            pole_types = {}
            for report in root.findall('.//construction_staking_report'):
                sequence = report.find('structure_number').text or ''
                framing = report.find('structure_name').text or ''
                latitude = report.find('latitude').text or ''
                longitude = report.find('longitude').text or ''
                x_easting = report.find('x_easting').text or ''
                y_northing = report.find('y_northing').text or ''
                stake_description = report.find('stake_description').text or ''

                if "P1" in stake_description:
                    pole_type = report.find('pole_property_label').text or ''
                    pole_types[sequence] = pole_type

                framing_parts = framing.split(" ", 2)
                if len(framing_parts) > 2:
                    framing = framing_parts[-1]
                    framing = " ".join(framing.split()[:-1])

                if sequence not in data:
                    data[sequence] = []

                data[sequence].append({
                    'framing': framing,
                    'latitude': latitude,
                    'longitude': longitude,
                    'x_easting': x_easting,
                    'y_northing': y_northing,
                    'stake_description': stake_description
                })

            anchor_data = []
            guy_types = ["P2", "PG", "SE", "NG", "CM", "FG"]
            for sequence, points in data.items():
                p1_point = None
                for point in points:
                    if "P1" in point['stake_description']:
                        p1_point = point
                        break

                if p1_point:
                    x_origin = float(p1_point['x_easting'])
                    y_origin = float(p1_point['y_northing'])
                    stake_description_set = set()
                    for point in points:
                        for guy_type in guy_types:
                            if guy_type in point['stake_description']:
                                x_next = float(point['x_easting'])
                                y_next = float(point['y_northing'])
                                lead_length = math.sqrt((x_next - x_origin)**2 + (y_next - y_origin)**2)
                                theta = math.degrees(math.atan2(y_next - y_origin, x_next - x_origin))
                                direction = self.get_cardinal_direction(theta)
                                descriptions = point['stake_description'].split(',')
                                for description in descriptions:
                                    if description.strip() not in stake_description_set:
                                        stake_description_set.add(description.strip())
                                        anchor_data.append({
                                            'sequence': sequence,
                                            'type': f"P1 to {description.strip()}",
                                            'latitude': point['latitude'],
                                            'longitude': point['longitude'],
                                            'framing': point['framing'],
                                            'anchor_direction': direction,
                                            'lead_length': lead_length
                                        })

            anchor_data.sort(key=lambda x: x['sequence'])

            with open(f"step{self.step}.txt", "w") as file:
                max_lengths = {
                    'sequence': max(len(item['sequence']) for item in anchor_data),
                    'type': max(len(item['type']) for item in anchor_data),
                    'latitude': max(len(str(item['latitude'])) for item in anchor_data),
                    'longitude': max(len(str(item['longitude'])) for item in anchor_data),
                    'framing': max(len(item['framing']) for item in anchor_data),
                    'anchor_direction': max(len(item['anchor_direction']) for item in anchor_data),
                    'lead_length': max(len(f"{item['lead_length']:.2f}") for item in anchor_data)
                }
                headers = [
                    ("Sequence", max_lengths['sequence']),
                    ("Type", max_lengths['type']),
                    ("Latitude", max_lengths['latitude']),
                    ("Longitude", max_lengths['longitude']),
                    ("Framing", max_lengths['framing']),
                    ("Anchor Direction", max_lengths['anchor_direction']),
                    ("Lead Length", max_lengths['lead_length'])
                ]

                header_row = " | ".join(f"{header[0]:<{header[1]}}" for header in headers)
                file.write(header_row + "\n")
                file.write("-" * len(header_row) + "\n")

                for item in anchor_data:
                    row = [
                        f"{item['sequence']:<{max_lengths['sequence']}}",
                        f"{item['type']:<{max_lengths['type']}}",
                        f"{item['latitude']:<{max_lengths['latitude']}}",
                        f"{item['longitude']:<{max_lengths['longitude']}}",
                        f"{item['framing']:<{max_lengths['framing']}}",
                        f"{item['anchor_direction']:<{max_lengths['anchor_direction']}}",
                        f"{item['lead_length']:<{max_lengths['lead_length']}.2f}"
                    ]
                    file.write(" | ".join(row) + "\n")

            with open("step3types.txt", "w") as file:
                file.write("Sequence\tPole Type\n")
                for sequence, pole_type in sorted(pole_types.items()):
                    file.write(f"{sequence}\t{pole_type}\n")

            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse XML file: {e}")

    def get_cardinal_direction(self, angle):
        if -22.5 < angle <= 22.5:
            return 'E'
        elif 22.5 < angle <= 67.5:
            return 'NE'
        elif 67.5 < angle <= 112.5:
            return 'N'
        elif 112.5 < angle <= 157.5:
            return 'NW'
        elif -67.5 < angle <= -22.5:
            return 'SE'
        elif -112.5 < angle <= -67.5:
            return 'S'
        elif -157.5 < angle <= -112.5:
            return 'SW'
        else:
            return 'W'

    def parse_and_continue_stringing_chart(self):
        pasted_data = self.paste_text.get("1.0", tk.END).strip()
        if not pasted_data:
            messagebox.showerror("Error", "Please paste data into the text box.")
            return

        try:
            sections = re.findall(r"Stringing Chart Report\n\nCircuit '(.*?)' Section #(.*?) from structure #(.*?) to structure #(.*?),.*?Span\n(.*?)\n\n", pasted_data, re.DOTALL)
            self.output_data = []

            for section in sections:
                circuit_type, section_num, start_seq, end_seq, spans_data = section
                if "Span Guy" in circuit_type:
                    continue  # Skip Span Guy entries from the pasted data

                spans = re.findall(r"\n\s+(\d+\.\d+)\s+", spans_data)
                if spans:
                    total_span_length = sum(map(float, spans))
                else:
                    total_span_length_match = re.search(r"Ruling span \(ft\) (\d+\.\d+)", spans_data)
                    if total_span_length_match:
                        total_span_length = float(total_span_length_match.group(1))
                    else:
                        total_span_length = 0.0
                sequences = f"{start_seq} - {end_seq}"
                self.output_data.append((section_num, sequences, total_span_length, circuit_type))

            messagebox.showinfo("Success", f"Pasted data parsed successfully. Please upload the Span Guy XML file.")
            self.upload_file()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse stringing chart data: {e}")

    def parse_span_guy_xml(self):
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
                self.output_data.append((section_num, sequences, total_span_length, circuit_type))

            self.output_data.sort(key=lambda x: int(x[0]))

            with open(f"step{self.step}.txt", "w") as file:
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

                header_row = " | ".join(f"{header[0]:<{header[1]}}" for header in headers)
                file.write(header_row + "\n")
                file.write("-" * len(header_row) + "\n")

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
            messagebox.showerror("Error", f"Failed to parse Span Guy XML file: {e}")

    def parse_primary_conductor_data(self):
        pasted_data = self.paste_text.get("1.0", tk.END).strip()
        if not pasted_data:
            messagebox.showerror("Error", "Please paste data into the text box.")
            return

        try:
            sections = re.findall(r"Stringing Chart Report\n\nCircuit '(.*?)' Section #(.*?) from structure #(.*?) to structure #(.*?),.*?Span\n(.*?)\n\n", pasted_data, re.DOTALL)
            output_data = []

            for section in sections:
                circuit_type, section_num, start_seq, end_seq, spans_data = section
                spans = re.findall(r"\n\s+(\d+\.\d+)\s+", spans_data)
                total_span_length = sum(map(float, spans))
                sequences = f"{start_seq} -> {end_seq}"
                circuit_value = int(re.search(r'(\d+)PH', circuit_type).group(1))
                result = total_span_length * circuit_value
                output_data.append((section_num, sequences, circuit_type, circuit_value, total_span_length, result))

            with open(f"step{self.step}.txt", "w") as file:
                max_lengths = {
                    'section_num': max(len(str(row[0])) for row in output_data),
                    'sequences': max(len(row[1]) for row in output_data),
                    'circuit_type': max(len(row[2]) for row in output_data),
                    'circuit_value': max(len(str(row[3])) for row in output_data),
                    'total_span_length': max(len(f"{row[4]:.2f}") for row in output_data),
                    'result': max(len(f"{row[5]:.2f}") for row in output_data),
                }
                headers = [
                    ("Section #", max_lengths['section_num']),
                    ("Structure -> Structure", max_lengths['sequences']),
                    ("Circuit Type", max_lengths['circuit_type']),
                    ("Circuit Value", max_lengths['circuit_value']),
                    ("Span Length", max_lengths['total_span_length']),
                    ("Result", max_lengths['result'])
                ]

                header_row = " | ".join(f"{header[0]:<{header[1]}}" for header in headers)
                file.write(header_row + "\n")
                file.write("-" * len(header_row) + "\n")

                for row in output_data:
                    formatted_row = [
                        f"{row[0]:<{max_lengths['section_num']}}",
                        f"{row[1]:<{max_lengths['sequences']}}",
                        f"{row[2]:<{max_lengths['circuit_type']}}",
                        f"{row[3]:<{max_lengths['circuit_value']}}",
                        f"{row[4]:<{max_lengths['total_span_length']}.2f}",
                        f"{row[5]:<{max_lengths['result']}.2f}",
                    ]
                    file.write(" | ".join(formatted_row) + "\n")

            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse primary conductor data: {e}")

    def parse_step6_structure_usage(self):
        try:
            tree = ET.parse(self.file_path)
            root = tree.getroot()
            output_data = []
            for report in root.findall('.//summary_of_maximum_element_usages_for_structure_range'):
                seq_no = report.find('str_no').text
                element_label = report.find('element_label').text
                element_type = report.find('element_type').text
                max_usage = report.find('maximum_usage').text
                if element_type == "Guy" or element_type == "Cable":
                    output_data.append((seq_no, element_label, element_type, max_usage))

            with open(f"step{self.step}.txt", "w") as file:
                max_lengths = {
                    'sequence': max(len(row[0]) for row in output_data),
                    'element_label': max(len(row[1]) for row in output_data),
                    'element_type': max(len(row[2]) for row in output_data),
                    'max_usage': max(len(row[3]) for row in output_data)
                }
                headers = [
                    ("Sequence #", max_lengths['sequence']),
                    ("Element Label", max_lengths['element_label']),
                    ("Element Type", max_lengths['element_type']),
                    ("Maximum Usage", max_lengths['max_usage'])
                ]

                header_row = " | ".join(f"{header[0]:<{header[1]}}" for header in headers)
                file.write(header_row + "\n")
                file.write("-" * len(header_row) + "\n")

                for row in output_data:
                    formatted_row = [
                        f"{row[0]:<{max_lengths['sequence']}}",
                        f"{row[1]:<{max_lengths['element_label']}}",
                        f"{row[2]:<{max_lengths['element_type']}}",
                        f"{row[3]:<{max_lengths['max_usage']}}"
                    ]
                    file.write(" | ".join(formatted_row) + "\n")

            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse XML file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DataExtractionApp(root)
    root.mainloop()