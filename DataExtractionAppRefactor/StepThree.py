from StepBase import StepBase
import tkinter as tk
from tkinter import ttk, messagebox
import xml.etree.ElementTree as ET
import math

class StepThree(StepBase):
    def __init__(self):
        super().__init__()
        self.step = 3
        self.data = {}
        self.pole_types = {}

    def setup_widgets(self, parent_frame):
        # Create step label
        self.step_label = ttk.Label(parent_frame, 
                                  text="Step 3: Upload your Construction Staking Report", 
                                  font=("Arial", 14))
        self.step_label.pack(pady=10)

        # Create file upload button
        self.create_upload_widgets(parent_frame, 
                                 "Upload Construction Staking Report",
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
            self.data = {}
            self.pole_types = {}

            # Extract data from XML
            for report in root.findall('.//construction_staking_report'):
                sequence = report.find('structure_number').text or ''
                framing = report.find('structure_name').text or ''
                latitude = report.find('latitude').text or ''
                longitude = report.find('longitude').text or ''
                x_easting = report.find('x_easting').text or ''
                y_northing = report.find('y_northing').text or ''
                stake_description = report.find('stake_description').text or ''

                # Store pole type if it is "P1"
                if "P1" in stake_description:
                    pole_type = report.find('pole_property_label').text or ''
                    self.pole_types[sequence] = pole_type

                # Process framing description
                framing_parts = framing.split(" ", 2)
                if len(framing_parts) > 2:
                    framing = framing_parts[-1]
                    framing = " ".join(framing.split()[:-1])

                if sequence not in self.data:
                    self.data[sequence] = []

                self.data[sequence].append({
                    'framing': framing,
                    'latitude': latitude,
                    'longitude': longitude,
                    'x_easting': x_easting,
                    'y_northing': y_northing,
                    'stake_description': stake_description
                })
            
            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process XML file: {e}")
            return False

    def save_data(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please upload a file first!")
            return

        try:
            # Process anchor data
            anchor_data = []
            guy_types = ["P2", "PG", "SE", "NG", "CM", "FG"]

            for sequence, points in self.data.items():
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
                                
                                # Calculate lead length and direction
                                lead_length = math.sqrt(
                                    (x_next - x_origin) ** 2 + 
                                    (y_next - y_origin) ** 2
                                )
                                theta = math.degrees(math.atan2(
                                    y_next - y_origin, 
                                    x_next - x_origin
                                ))
                                direction = self._get_cardinal_direction(theta)
                                
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

            # Save anchor data to file
            self._save_anchor_data(anchor_data)
            
            # Save pole type data
            self._save_pole_type_data()

            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {e}")

    def _save_anchor_data(self, anchor_data):
        with open("extractConstrucStakingReport_framing_type_direction_length.txt", "w") as file:
            # Calculate max lengths for formatting
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

            # Write headers
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

    def _save_pole_type_data(self):
        with open("extractPoleType.txt", "w") as file:
            file.write("Sequence\tPole Type\n")
            for sequence, pole_type in sorted(self.pole_types.items()):
                file.write(f"{sequence}\t{pole_type}\n")

    def _get_cardinal_direction(self, angle):
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