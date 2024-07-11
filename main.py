import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
import xml.etree.ElementTree as ET
import re
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
        elif self.step == 3:
            self.step_label.config(text="Step 3: Upload your Construction Staking Report")
            self.upload_btn.config(text="Upload Construction Staking Report", command=self.upload_file)
            self.paste_label.pack_forget()
            self.paste_text.pack_forget()
            self.upload_btn.pack(pady=10)
            self.column_listbox.pack_forget()
            self.process_btn.pack_forget()
        elif self.step == 4:
            self.step_label.config(text="Step 4: Copy and Paste your Stringing Chart - Neutral and Span Guy")
            self.upload_btn.pack_forget()
            self.paste_label.config(text="Paste Stringing Chart Data Here")
            self.paste_label.pack(pady=10)
            self.paste_text.pack(pady=10)
            self.process_btn.config(text="Parse Stringing Chart Data", command=self.parse_stringing_chart_data)
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
            self.process_btn.config(text="")
            self.process_btn.pack_forget()
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
            for report in root.findall('.//construction_staking_report'):
                sequence = report.find('structure_number').text or ''
                framing = report.find('structure_name').text or ''
                latitude = report.find('latitude').text or ''
                longitude = report.find('longitude').text or ''
                x_easting = report.find('x_easting').text or ''
                y_northing = report.find('y_northing').text or ''
                stake_description = report.find('stake_description').text or ''
                
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
            for sequence, points in data.items():
                for i in range(len(points) - 1):
                    current = points[i]
                    next_point = points[i + 1]
                    x_origin = float(current['x_easting'])
                    y_origin = float(current['y_northing'])
                    x_next = float(next_point['x_easting'])
                    y_next = float(next_point['y_northing'])
                    # calculate lead length and direction.
                    lead_length = math.sqrt((x_next - x_origin)**2 + (y_next - y_origin)**2)
                    theta = math.degrees(math.atan2(y_next - y_origin, x_next - x_origin))
                    direction = self.get_cardinal_direction(theta)
                    anchor_data.append({
                        'sequence': f"{sequence} {current['stake_description']} to {next_point['stake_description']}",
                        'latitude': current['latitude'],
                        'longitude': current['longitude'],
                        'framing': current['framing'],
                        'anchor_direction': direction,
                        'lead_length': lead_length
                    })
            
            # Calculate maximum widths for each column
            max_lengths = {
                'sequence': max(len(item['sequence']) for item in anchor_data) + 2,
                'latitude': max(len(item['latitude']) for item in anchor_data) + 2,
                'longitude': max(len(item['longitude']) for item in anchor_data) + 2,
                'framing': max(len(item['framing']) for item in anchor_data) + 2,
                'anchor_direction': max(len(item['anchor_direction']) for item in anchor_data) + 2,
                'lead_length': max(len(f"{item['lead_length']:.2f}") for item in anchor_data) + 2
            }
            
            # Write the aligned data to a file
            with open(f"step{self.step}.txt", "w") as file:
                header = f"{'Sequence'.ljust(max_lengths['sequence'])}{'Latitude'.ljust(max_lengths['latitude'])}{'Longitude'.ljust(max_lengths['longitude'])}{'Framing'.ljust(max_lengths['framing'])}{'Anchor Direction'.ljust(max_lengths['anchor_direction'])}{'Lead Length'.ljust(max_lengths['lead_length'])}\n"
                file.write(header)
                for item in anchor_data:
                    line = f"{item['sequence'].ljust(max_lengths['sequence'])}{item['latitude'].ljust(max_lengths['latitude'])}{item['longitude'].ljust(max_lengths['longitude'])}{item['framing'].ljust(max_lengths['framing'])}{item['anchor_direction'].ljust(max_lengths['anchor_direction'])}{f'{item['lead_length']:.2f}'.ljust(max_lengths['lead_length'])}\n"
                    file.write(line)
            
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
    
    def parse_stringing_chart_data(self):
        pasted_data = self.paste_text.get("1.0", tk.END).strip()
        if not pasted_data:
            messagebox.showerror("Error", "Please paste data into the text box.")
            return
        
        try:
            self.process_stringing_chart(pasted_data)
            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse stringing chart data: {e}")
    
    def parse_primary_conductor_data(self):
        pasted_data = self.paste_text.get("1.0", tk.END).strip()
        if not pasted_data:
            messagebox.showerror("Error", "Please paste data into the text box.")
            return
        
        try:
            self.process_stringing_chart(pasted_data)
            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse primary conductor data: {e}")
    
    def process_stringing_chart(self, data):
        sections = re.findall(r"Stringing Chart Report\n\nCircuit '(.*?)' Section #(.*?) from structure #(.*?) to structure #(.*?),.*?Span\n(.*?)\n\n", data, re.DOTALL)
        output_data = []
        
        for section in sections:
            circuit_type, section_num, start_seq, end_seq, spans_data = section
            spans = re.findall(r"\s+(\d+\.\d+)\s+", spans_data)
            total_span_length = sum(map(float, spans))
            sequences = f"{start_seq} - {end_seq}"
            output_data.append((section_num, sequences, total_span_length, circuit_type))
        
        with open(f"step{self.step}.txt", "w") as file:
            file.write("Section #\tSequence #s\tTotal Span Length\tCircuit Type\n")
            for row in output_data:
                file.write(f"{row[0]}\t{row[1]}\t{row[2]}\t{row[3]}\n")
    
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
                if element_type == "Guy":
                    output_data.append((seq_no, element_label, element_type, max_usage))
            
            with open(f"step{self.step}.txt", "w") as file:
                file.write("Sequence #\tElement Label\tElement Type\tMaximum Usage\n")
                for row in output_data:
                    file.write(f"{row[0]}\t{row[1]}\t{row[2]}\t{row[3]}\n")
            
            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse XML file: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DataExtractionApp(root)
    root.mainloop()
