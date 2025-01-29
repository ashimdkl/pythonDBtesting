import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import xml.etree.ElementTree as ET
import math
import re


class XMLTagExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XML Tag Extractor")
        self.root.geometry("600x400")

        # Frame for GUI layout
        self.frame = ttk.Frame(self.root, padding=10)
        self.frame.pack(expand=True, fill="both")

        # Label
        self.label = ttk.Label(
            self.frame, text="Upload a large XML file to extract tags", font=("Helvetica", 14)
        )
        self.label.pack(pady=10)

        # Upload button
        self.upload_btn = ttk.Button(
            self.frame, text="Upload XML File", command=self.upload_file
        )
        self.upload_btn.pack(pady=10)

        # Status label
        self.status_label = ttk.Label(self.frame, text="", font=("Helvetica", 10))
        self.status_label.pack(pady=10)

    def upload_file(self):
        # Open file dialog to select an XML file
        file_path = filedialog.askopenfilename(
            filetypes=[("XML Files", "*.xml")], title="Select XML File"
        )
        if file_path:
            self.status_label.config(text="Processing...")
            self.extract_tags(file_path)
            self.status_label.config(text="Extraction complete! Files saved.")

    def extract_tags(self, file_path):
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()

            # Extract Step 3 tags
            self.extract_step3(root)

            # Extract Step 6 tags
            self.extract_step6(root)

            # Extract Step 7 tags
            self.extract_step7(root)

            # Extract grades and write to grade.txt
            self.extract_loading_grades(root)

            # Extract additional steps as needed
            self.extract_step4_span_guy(root)

            messagebox.showinfo(
                "Success",
                "Data extracted successfully! Files saved for all steps."
            )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse XML file: {e}")

    def get_cardinal_direction(self, angle):
        # Convert an angle into a cardinal direction (e.g., N, NE, E)
        if -2.0 < angle <= 2.0:
            return 'E'
        elif 2.0 < angle <= 88.0:
            return 'NE'
        elif 88.0 < angle <= 92.0:
            return 'N'
        elif 92.0 < angle <= 178.0:
            return 'NW'
        elif -88.0 < angle <= -2.0:
            return 'SE'
        elif -92.0 < angle <= -88.0:
            return 'S'
        elif -178.0 < angle <= -92.0:
            return 'SW'
        else:
            return 'W'

    def extract_step3(self, root):
        # Initialize data structures
        data = {}
        pole_types = {}
        
        # Find all construction staking reports
        for report in root.findall('.//construction_staking_report'):
            sequence = report.find('structure_number').text or ''
            framing = report.find('structure_name').text or ''
            latitude = report.find('latitude').text or ''
            longitude = report.find('longitude').text or ''
            x_easting = report.find('x_easting').text or ''
            y_northing = report.find('y_northing').text or ''
            stake_description = report.find('stake_description').text or ''

            # Clean and standardize sequence number (remove "SEQ" and spaces)
            original_sequence = sequence
            sequence = sequence.replace('SEQ', '').strip()
            if '&' in sequence:
                sequence = sequence.split('&')[0].strip()

            # Track pole types only for P1 points
            if "P1" in stake_description:
                pole_type = report.find('pole_property_label').text or ''
                pole_types[sequence] = pole_type

            # Clean and standardize the framing name
            framing = self.clean_framing_name(framing, original_sequence)

            # Initialize sequence data if not exists
            if sequence not in data:
                data[sequence] = []

            # Store all points
            data[sequence].append({
                'framing': framing,
                'latitude': latitude,
                'longitude': longitude,
                'x_easting': x_easting,
                'y_northing': y_northing,
                'stake_description': stake_description
            })

        # Process anchor data
        anchor_data = []
        guy_types = ["P2", "PG", "SE", "NG", "CM", "FG"]

        for sequence, points in data.items():
            # Find P1 point first
            p1_point = None
            for point in points:
                if "P1" in point['stake_description']:
                    p1_point = point
                    break

            # Only process if P1 exists
            if p1_point:
                x_origin = float(p1_point['x_easting'])
                y_origin = float(p1_point['y_northing'])
                stake_description_set = set()

                # Process all guy points
                for point in points:
                    for guy_type in guy_types:
                        if guy_type in point['stake_description']:
                            # Calculate distances and angles
                            x_next = float(point['x_easting'])
                            y_next = float(point['y_northing'])
                            lead_length = math.sqrt((x_next - x_origin) ** 2 + 
                                                  (y_next - y_origin) ** 2)
                            theta = math.degrees(math.atan2(y_next - y_origin, 
                                                          x_next - x_origin))
                            direction = self.get_cardinal_direction(theta)
                            
                            # Handle multiple descriptions (split by comma)
                            descriptions = point['stake_description'].split(',')
                            for description in descriptions:
                                desc = description.strip()
                                if desc not in stake_description_set:
                                    stake_description_set.add(desc)
                                    anchor_data.append({
                                        'sequence': sequence,
                                        'type': f"P1 to {desc}",
                                        'latitude': point['latitude'],
                                        'longitude': point['longitude'],
                                        'framing': point['framing'],
                                        'anchor_direction': direction,
                                        'lead_length': lead_length
                                    })

                # Add standalone P1 if no anchors found
                if not stake_description_set:
                    anchor_data.append({
                        'sequence': sequence,
                        'type': "P1",
                        'latitude': p1_point['latitude'],
                        'longitude': p1_point['longitude'],
                        'framing': p1_point['framing'],
                        'anchor_direction': '',
                        'lead_length': 0.0
                    })

        # Sort and save anchor data
        anchor_data.sort(key=lambda x: int(re.findall(r'\d+', x['sequence'])[0]))
        
        # Write anchor data file
        with open("XMLextractConstrucStakingReport_framing_type_direction_length.txt", "w") as file:
            # Calculate maximum lengths for formatting
            max_lengths = {
                'sequence': max(len(item['sequence']) for item in anchor_data),
                'type': max(len(item['type']) for item in anchor_data),
                'latitude': max(len(str(item['latitude'])) for item in anchor_data),
                'longitude': max(len(str(item['longitude'])) for item in anchor_data),
                'framing': max(len(item['framing']) for item in anchor_data),
                'anchor_direction': max(len(item['anchor_direction']) for item in anchor_data),
                'lead_length': max(len(f"{item['lead_length']:.2f}") for item in anchor_data)
            }

            # Write headers
            headers = [
                ("Sequence", max_lengths['sequence']),
                ("Type", max_lengths['type']),
                ("Latitude", max_lengths['latitude']),
                ("Longitude", max_lengths['longitude']),
                ("Framing", max_lengths['framing']),
                ("Anchor Direction", max_lengths['anchor_direction']),
                ("Lead Length", max_lengths['lead_length'])
            ]

            # Write header row
            header_row = " | ".join(f"{header[0]:<{header[1]}}" for header in headers)
            file.write(header_row + "\n")
            file.write("-" * len(header_row) + "\n")

            # Write data rows
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

        # Write pole type data
        with open("XMLextractPoleType.txt", "w") as file:
            file.write("Sequence\tPole Type\n")
            for sequence, pole_type in sorted(pole_types.items(), 
                                            key=lambda x: int(re.findall(r'\d+', x[0])[0])):
                file.write(f"{sequence}\t{pole_type}\n")


    def clean_framing_name(self, framing, sequence):
        """Clean framing name to handle special cases"""
        # Remove sequence number and 'SEQ' from start
        framing = framing.replace(f"SEQ {sequence}", "").strip()
        
        # Handle cases with '&' in sequence
        if '&' in framing:
            parts = framing.split('&', 1)
            if len(parts) > 1:
                framing = parts[1].strip()
        
        # Remove .POL from end
        framing = framing.replace(".POL", "").strip()
        
        # Remove ST patterns from end (like ST50.01, ST45.03)
        framing = re.sub(r'\s+ST\d+\.\d+(?:\s+ST\d+\.\d+)*$', '', framing)
        
        return framing.strip()

    def extract_step6(self, root):
        step6_file = "XMLextractGuyUsage_seq_elementType_usage.txt"
        with open(step6_file, "w") as file:
            file.write("Step 6: Structure Usage Tags\n")
            file.write("Sequence # | Element Label      | Element Type    | Maximum Usage\n")
            file.write("-" * 70 + "\n")

            # Collect filtered data first
            output_data = []
            for report in root.findall(".//summary_of_maximum_usages_by_load_case_for_structure_range"):
                try:
                    seq_no = report.find("str_no").text or "N/A"
                    element_label = report.find("element_label").text or "N/A"
                    element_type = report.find("element_type").text or "N/A"
                    max_usage = report.find("maximum_usage").text or "N/A"

                    # Only include Guy and Cable elements
                    if element_type in ["Guy", "Cable"]:
                        output_data.append({
                            'sequence': seq_no,
                            'element_label': element_label,
                            'element_type': element_type,
                            'max_usage': max_usage
                        })

                except Exception as e:
                    print(f"Skipping an entry due to error: {e}")

            # Sort by sequence number
            output_data.sort(key=lambda x: int(re.findall(r'\d+', x['sequence'])[0]))

            # Calculate maximum lengths for formatting
            max_lengths = {
                'sequence': max(len(item['sequence']) for item in output_data),
                'element_label': max(len(item['element_label']) for item in output_data),
                'element_type': max(len(item['element_type']) for item in output_data),
                'max_usage': max(len(item['max_usage']) for item in output_data)
            }

            # Write filtered and formatted data
            for item in output_data:
                file.write(f"{item['sequence']:<{max_lengths['sequence']}} | "
                          f"{item['element_label']:<{max_lengths['element_label']}} | "
                          f"{item['element_type']:<{max_lengths['element_type']}} | "
                          f"{item['max_usage']:<{max_lengths['max_usage']}}\n")

    def extract_step7(self, root):
        step7_file = "XMLextractMAX_sequence_MaxForce.txt"
        with open(step7_file, "w") as file:
            file.write("Step 7: Joint Support Tags\n")
            file.write("Sequence | Max Force\n")
            file.write("-" * 30 + "\n")

            for report in root.findall(".//summary_of_joint_support_reactions_for_all_load_cases_for_structure_range"):
                try:
                    seq_no = report.find("str_no").text or "N/A"
                    shear_force = float(report.find("shear_force").text or 0)
                    bending_moment = float(report.find("bending_moment").text or 0)
                    max_force = max(shear_force, bending_moment)

                    # Write the sequence and max force into the file
                    file.write(f"{seq_no:<10} | {max_force:<8.2f}\n")

                except Exception as e:
                    print(f"Skipping an entry due to error: {e}")

    def extract_step4_span_guy(self, root):
        step4_file = "XMLextractStringingChartNeutralSpan_section_seq_totalSpanLength_circuitType.txt"
        with open(step4_file, "w") as file:
            # Write headers
            file.write("Section # | Sequence #s       | Total Span Length | Circuit Type\n")
            file.write("-" * 74 + "\n")

            # Collect all data first
            span_data = []
            for section in root.findall(".//section_sagging_data"):
                circuit_type = section.find("circuit").text or ""
                section_num = section.find("sec_no").text or ""
                start_seq = section.find("from_str").text or ""
                end_seq = section.find("to_str").text or ""
                ruling_span = float(section.find("ruling_span").text or 0)
                
                # Only include Neutral or Span Guy circuit types
                if "Neutral" in circuit_type or "Span Guy" in circuit_type:
                    # Clean up sequence format
                    start_clean = start_seq.replace('SEQ ', '').strip()
                    end_clean = end_seq.replace('SEQ ', '').strip()
                    
                    span_data.append({
                        'section_num': section_num,
                        'start_seq': start_clean,
                        'end_seq': end_clean,
                        'total_span': ruling_span,
                        'circuit_type': circuit_type.strip()
                    })

            # Sort by section number
            span_data.sort(key=lambda x: int(x['section_num']))

            # Write data rows
            for item in span_data:
                # Format sequence range
                seq_range = f"{item['start_seq']} - {item['end_seq']}"
                
                file.write(f"{item['section_num']:<3} | "
                          f"{seq_range:<15} | "
                          f"{item['total_span']:<8.2f} | "
                          f"{item['circuit_type']:<24}\n")
                
    
    def extract_loading_grades(self, root):
        grade_file = "XMLextractGrade_sequence_label_type_grade.txt"
        with open(grade_file, "w") as file:
            # Write headers
            file.write("Sequence | Label | Type      | Grade\n")
            file.write("-" * 40 + "\n")
            
            for report in root.findall(".//summary_of_maximum_usages_by_load_case_for_structure_range"):
                try:
                    seq_no = report.find("str_no").text or "N/A"
                    element_label = report.find("element_label").text or "N/A"
                    element_type = report.find("element_type").text or "N/A"
                    load_case = report.find("load_case").text or ""
                    
                    # Extract grade from load_case (e.g., "GRADE B")
                    grade_match = re.search(r"GRADE ([A-Z])", load_case)
                    grade = grade_match.group(1) if grade_match else "N/A"

                    # Skip rows where the grade is N/A
                    if grade == "N/A":
                        continue

                    # Write the extracted information to the file
                    file.write(f"{seq_no:<8} | {element_label:<5} | {element_type:<10} | {grade}\n")

                except Exception as e:
                    print(f"Skipping an entry due to error: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = XMLTagExtractorApp(root)
    root.mainloop()