# StepThreeSixXML.py

from StepBase import StepBase
import tkinter as tk
from tkinter import ttk, messagebox
import xml.etree.ElementTree as ET
import math

class StepThreeSixXML(StepBase):
    def __init__(self):
        super().__init__()
        self.step = 3  # or 3/6 combined
        # Data structures for the "construction staking" portion
        self.construction_data = {}
        self.pole_types = {}
        # Data structure for the "structure usage" portion
        self.usage_data = []

    def setup_widgets(self, parent_frame):
        """Sets up the GUI elements for uploading one big XML file."""
        self.step_label = ttk.Label(parent_frame,
            text="Steps 3 & 6 Combined: Upload One Large XML",
            font=("Arial", 14))
        self.step_label.pack(pady=10)

        # Single upload button for your large XML
        self.create_upload_widgets(
            parent_frame,
            "Upload Combined (Construction & Usage) XML",
            [("XML files", "*.xml")]
        )

        self.process_btn = ttk.Button(
            parent_frame,
            text="Parse Data (3 & 6) and Move to Next Step",
            command=self.save_data
        )
        self.process_btn.pack(pady=10)

        self.skip_btn = ttk.Button(
            parent_frame,
            text="Skip This Step",
            command=self.next_step
        )
        self.skip_btn.pack(pady=10)

    def process_file(self):
        """Parses the large XML (both Step 3 & Step 6 data)."""
        if not self.file_path:
            messagebox.showerror("Error", "Please upload the combined XML first!")
            return False

        try:
            tree = ET.parse(self.file_path)
            root = tree.getroot()

            ##################################
            # 1) Parse the Construction data
            ##################################
            # (Similar to Step Three code)
            for report in root.findall('.//construction_staking_report'):
                sequence = report.find('structure_number').text or ''
                framing = report.find('structure_name').text or ''
                latitude = report.find('latitude').text or ''
                longitude = report.find('longitude').text or ''
                x_easting = report.find('x_easting').text or ''
                y_northing = report.find('y_northing').text or ''
                stake_description = report.find('stake_description').text or ''

                # If "P1" in stake_description => store the 'pole_property_label'
                if "P1" in stake_description:
                    pole_type = report.find('pole_property_label').text or ''
                    self.pole_types[sequence] = pole_type

                # Simplify the “framing” if needed
                framing_parts = framing.split(" ", 2)
                if len(framing_parts) > 2:
                    framing = framing_parts[-1]
                    framing = " ".join(framing.split()[:-1])

                if sequence not in self.construction_data:
                    self.construction_data[sequence] = []

                self.construction_data[sequence].append({
                    'framing': framing,
                    'latitude': latitude,
                    'longitude': longitude,
                    'x_easting': x_easting,
                    'y_northing': y_northing,
                    'stake_description': stake_description
                })

            ##################################
            # 2) Parse the Usage data
            ##################################
            # (Similar to Step Six code)
            for summary in root.findall('.//summary_of_maximum_element_usages_for_structure_range'):
                seq_no = summary.find('str_no').text
                element_label = summary.find('element_label').text
                element_type = summary.find('element_type').text
                max_usage = summary.find('maximum_usage').text

                # If it’s Guy or Cable (like in your step6)
                if element_type in ("Guy", "Cable"):
                    self.usage_data.append((seq_no, element_label, element_type, max_usage))

            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process the combined XML: {e}")
            return False

    def save_data(self):
        """Writes out two files:
           1) extractConstrucStakingReport_framing_type_direction_length.txt
           2) extractGuyUsage_seq_elementType_usage.txt
        """
        if not self.process_file():
            return

        try:
            # 1) Save construction data to text file (like Step3)
            self._save_construction_staking()

            # 2) Save usage data to text file (like Step6)
            self._save_structure_usage()

            messagebox.showinfo("Success", "Data from combined Steps 3 & 6 saved successfully!")
            # Then move to next step or do whatever
            self.next_step()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {e}")

    def _save_construction_staking(self):
        """Mimics your Step 3 anchor logic.  This includes anchor_data, etc."""
        anchor_data = []
        guy_types = ["P2", "PG", "SE", "NG", "CM", "FG"]

        for sequence, points in self.construction_data.items():
            # Look for “P1”
            p1_point = None
            for point in points:
                if "P1" in point['stake_description']:
                    p1_point = point
                    break

            if p1_point:
                x_origin = float(p1_point['x_easting'])
                y_origin = float(p1_point['y_northing'])
                stake_desc_set = set()
                for point in points:
                    for gtype in guy_types:
                        if gtype in point['stake_description']:
                            x_next = float(point['x_easting'])
                            y_next = float(point['y_northing'])
                            lead_length = math.sqrt((x_next - x_origin)**2 + (y_next - y_origin)**2)
                            angle = math.degrees(math.atan2(y_next - y_origin, x_next - x_origin))
                            direction = self._get_cardinal_direction(angle)

                            descs = point['stake_description'].split(',')
                            for desc in descs:
                                desc_clean = desc.strip()
                                if desc_clean not in stake_desc_set:
                                    stake_desc_set.add(desc_clean)
                                    anchor_data.append({
                                        'sequence': sequence,
                                        'type': f"P1 to {desc_clean}",
                                        'latitude': point['latitude'],
                                        'longitude': point['longitude'],
                                        'framing': point['framing'],
                                        'anchor_direction': direction,
                                        'lead_length': lead_length
                                    })

        anchor_data.sort(key=lambda x: x['sequence'])

        # Write to file
        with open("extractConstrucStakingReport_framing_type_direction_length.txt", "w") as f:
            # Compute column widths, etc. (same as Step 3)
            # ...
            pass

        # Also save pole types
        with open("extractPoleType.txt", "w") as pt_file:
            pt_file.write("Sequence\tPole Type\n")
            for seq, ptype in sorted(self.pole_types.items()):
                pt_file.write(f"{seq}\t{ptype}\n")

    def _save_structure_usage(self):
        """Mimics your Step 6 usage logic."""
        if not self.usage_data:
            return

        with open("extractGuyUsage_seq_elementType_usage.txt", "w") as f:
            # Calculate max lengths
            max_lengths = {
                'sequence': max(len(row[0]) for row in self.usage_data),
                'element_label': max(len(row[1]) for row in self.usage_data),
                'element_type': max(len(row[2]) for row in self.usage_data),
                'max_usage': max(len(row[3]) for row in self.usage_data)
            }
            headers = [
                ("Sequence #", max_lengths['sequence']),
                ("Element Label", max_lengths['element_label']),
                ("Element Type", max_lengths['element_type']),
                ("Maximum Usage", max_lengths['max_usage'])
            ]
            # Write header row
            header_row = " | ".join(f"{h[0]:<{h[1]}}" for h in headers)
            f.write(header_row + "\n")
            f.write("-" * len(header_row) + "\n")
            # Write data rows
            for (seq_no, label, etype, usage) in self.usage_data:
                row = [
                    f"{seq_no:<{max_lengths['sequence']}}",
                    f"{label:<{max_lengths['element_label']}}",
                    f"{etype:<{max_lengths['element_type']}}",
                    f"{usage:<{max_lengths['max_usage']}}"
                ]
                f.write(" | ".join(row) + "\n")

    def _get_cardinal_direction(self, angle):
        """Same cardinal-direction logic from your Step 3 code."""
        # ...
        return "E"  # placeholder




# for additional tags: #

# Suppose you want to also parse <span_and_wire_summary_for_structure_range>
    # for wire_summary in root.findall('.//span_and_wire_summary_for_structure_range'):
        # do something
        # store it in a data structure, or write to file, etc.