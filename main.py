# import tkinter as tk, filedialog, messagebox, ttk from tkinter, load_workbook from openpyxl, ET from xml.etree.ElementTree, re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
import xml.etree.ElementTree as ET
import re


# class DataExtractionApp being defined.
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
        
    # create initial frame for the app with the appropriate buttons and labels.
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
        
    # function to start the analysis by hiding the intro frame and displaying the main frame.
    def start_analysis(self):
        self.intro_frame.pack_forget()
        self.main_frame.pack(expand=True, fill="both")
    
    # function to upload a file and load the columns from the file.
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
    
    # function to load the columns from the uploaded file.
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
    
    # function to parse the data from the uploaded file.
    def parse_data(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please upload a file first!")
            return
        
        # Get the selected columns
        self.selected_columns = [self.column_listbox.get(i) for i in self.column_listbox.curselection()]
        if not self.selected_columns:
            messagebox.showerror("Error", "Please select at least one column to keep!")
            return
        
        # Parse the data based on the selected columns
        try:
            if self.step == 1:
                workbook = load_workbook(self.file_path)
                sheet = workbook.active
                
                # Write the selected columns to a new file by iterating over the rows and columns and then using the selected columns to write the data.
                with open(f"step{self.step}.txt", "w") as file:
                    file.write("\t".join(self.selected_columns) + "\n")
                    for row in sheet.iter_rows(min_row=2, values_only=True):
                        if row[0] is not None:  # Assuming sequence number is the first column
                            row_data = [str(row[self.columns.index(col)]) for col in self.selected_columns]
                            file.write("\t".join(row_data) + "\n")
            
            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse file: {e}")
    
    # function to move to the next step in the process.
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
    
    # function to parse the pasted data from the text box.
    def parse_pasted_data(self):
        pasted_data = self.paste_text.get("1.0", tk.END).strip()
        if not pasted_data:
            messagebox.showerror("Error", "Please paste data into the text box.")
            return
        
        # Parse the pasted data and write it to a file by splitting the data into lines and then into fields. Write the data to a new file.
        try:
            lines = pasted_data.split("\n")
            header = lines[0].split("\t")
            data = []
            for line in lines[1:]:
                fields = line.split("\t")
                sequence = fields[0][:4]  # Assuming sequence is the first 4 characters
                existing = fields[2]  # Assuming existing value is the third field
                data.append([sequence, existing])
            
            # Write the parsed data to a new file
            with open(f"step{self.step}.txt", "w") as file:
                file.write("Sequence\tExisting\n")
                for row in data:
                    file.write("\t".join(row) + "\n")
            
            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")
        
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse pasted data: {e}")
    
    # function to parse the XML data from the uploaded file.
    def parse_step3_xml(self):
        # Parse the XML data and write it to a file by iterating over the XML elements and extracting the required data.
        try:
            tree = ET.parse(self.file_path)
            root = tree.getroot()
            data = {}
            # Extract the required data from the XML file
            for report in root.findall('.//construction_staking_report'):
                sequence = report.find('structure_number').text or ''
                framing = report.find('structure_name').text or ''
                latitude = report.find('latitude').text or ''
                longitude = report.find('longitude').text or ''
                if report.find('stake_description').text == 'P1':
                    pole_type = report.find('pole_property_label').text or ''
                else:
                    pole_type = ""
                
                # Remove the "SEQ XXXX " and the last identifier part from the framing string
                framing_parts = framing.split(" ", 2)
                if len(framing_parts) > 2:
                    framing = framing_parts[-1]
                    framing = " ".join(framing.split()[:-1])
                
                # Store the data in a dictionary
                if sequence in data:
                    data[sequence]['framing'] = framing
                    data[sequence]['latitude'] = latitude
                    data[sequence]['longitude'] = longitude
                    if pole_type:
                        data[sequence]['pole_type'] = pole_type
                else:
                    data[sequence] = {
                        'framing': framing,
                        'latitude': latitude,
                        'longitude': longitude,
                        'pole_type': pole_type
                    }
            
            # Write the extracted data to a new file
            with open(f"step{self.step}.txt", "w") as file:
                file.write("Sequence\tLatitude\tLongitude\tFraming\tPole Type\n")
                for seq, values in data.items():
                    file.write(f"{seq}\t{values['latitude']}\t{values['longitude']}\t{values['framing']}\t{values['pole_type']}\n")
            
            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse XML file: {e}")

    # function to parse the stringing chart data from the text box.
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
    
    # function to parse the primary conductor data from the text box.
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
    
    # function to process the stringing chart data by extracting the required information.
    def process_stringing_chart(self, data):
        # we want to use regular expressions to extract the required information from the stringing chart data.
        sections = re.findall(r"Stringing Chart Report\n\nCircuit '(.*?)' Section #(.*?) from structure #(.*?) to structure #(.*?),.*?Span\n(.*?)\n\n", data, re.DOTALL)
        output_data = []
        
        # Extract the required information from the sections and write it to a new file.
        for section in sections:
            # Extract the section data
            circuit_type, section_num, start_seq, end_seq, spans_data = section
            # Extract the span lengths and calculate the total span length
            spans = re.findall(r"\s+(\d+\.\d+)\s+", spans_data)
            # Calculate the total span length
            total_span_length = sum(map(float, spans))
            # Store the data in the output list
            sequences = f"{start_seq} - {end_seq}"
            # Determine the circuit type based on the section number
            output_data.append((section_num, sequences, total_span_length, circuit_type))
        
        # Write the extracted data to a new file
        with open(f"step{self.step}.txt", "w") as file:
            file.write("Section #\tSequence #s\tTotal Span Length\tCircuit Type\n")
            for row in output_data:
                file.write(f"{row[0]}\t{row[1]}\t{row[2]}\t{row[3]}\n")
    
    # function to parse the structure usage data from the uploaded file.
    def parse_step6_structure_usage(self):
        try:
            tree = ET.parse(self.file_path)
            root = tree.getroot()
            output_data = []
            # Extract the required information from the XML file
            for report in root.findall('.//summary_of_maximum_element_usages_for_structure_range'):
                seq_no = report.find('str_no').text
                element_label = report.find('element_label').text
                element_type = report.find('element_type').text
                max_usage = report.find('maximum_usage').text
                # Store the data in the output list only if the element type is "Guy"
                if element_type == "Guy":
                    output_data.append((seq_no, element_label, element_type, max_usage))
            
            # Write the extracted data to a new file
            with open(f"step{self.step}.txt", "w") as file:
                file.write("Sequence #\tElement Label\tElement Type\tMaximum Usage\n")
                for row in output_data:
                    file.write(f"{row[0]}\t{row[1]}\t{row[2]}\t{row[3]}\n")
            
            # Show success message
            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse XML file: {e}")

# main function to run the application.
if __name__ == "__main__":
    root = tk.Tk()
    app = DataExtractionApp(root)
    root.mainloop()
