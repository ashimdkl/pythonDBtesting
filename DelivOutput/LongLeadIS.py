import re
from openpyxl import load_workbook
from tkinter import filedialog, messagebox
import os

class LongLeadGenerator:
    def __init__(self, workbook, wo_number, county, city_place):
        self.source_workbook = workbook
        self.wo_number = wo_number
        self.county = county
        self.city_place = city_place
        self.template_path = os.path.join(os.path.dirname(__file__), '..', 'templates', 'longLeadTemplate.xlsx')

    def safe_set_cell_value(self, sheet, cell_ref, value):
        """Safely set cell value handling merged cells"""
        # Get all merged cell ranges
        merged_ranges = sheet.merged_cells.ranges
        target_cell = sheet[cell_ref]
        
        # Check if cell is in a merged range
        for merged_range in merged_ranges:
            if target_cell.coordinate in merged_range:
                # Use the top-left cell of the merged range
                top_left = merged_range.start_cell
                top_left.value = value
                return
        
        # If not in merged range, set directly
        target_cell.value = value

    def process_poles(self, data_sheet):
        """Process pole data from column N"""
        pole_counts = {}
        
        for row in data_sheet.iter_rows(min_row=2, values_only=True):
            pole_type = row[13]  # Column N
            if pole_type and isinstance(pole_type, str):
                if pole_type not in pole_counts:
                    pole_counts[pole_type] = 1
                else:
                    pole_counts[pole_type] += 1
                    
        return pole_counts

    def process_conductors(self, data_sheet):
        """Process conductor data with 1.01 safety factor"""
        conductor_data = {}
        
        for row in data_sheet.iter_rows(min_row=2, values_only=True):
            if row[5] and isinstance(row[5], str):  # Column F
                if "ACSR" in row[5]:
                    conductor_type = row[5]
                    length = float(row[2]) if row[2] else 0  # Column C
                    if conductor_type not in conductor_data:
                        conductor_data[conductor_type] = 0
                    conductor_data[conductor_type] += length * 1.01  # Apply safety factor
                    
        return conductor_data

    def process_guys_and_anchors(self, data_sheet):
        """Calculate guy wire needs and anchor counts"""
        guy_count = 0
        anchor_numbers = set()
        
        for row in data_sheet.iter_rows(min_row=2, values_only=True):
            if row[7] and isinstance(row[7], str):  # Column H
                guy_info = row[7].upper()
                if any(x in guy_info for x in ['PG', 'FG', 'SPAN', 'NEUTRAL']) and 'CM' not in guy_info:
                    guy_count += 1
                matches = re.findall(r'PG(\d+)', guy_info)
                anchor_numbers.update(matches)
        
        total_guy_length = guy_count * 50  # 50 ft per guy
        num_spools = (total_guy_length / 200) + 1  # Safety factor
        
        return int(num_spools), len(anchor_numbers)

    def generate_sheet(self):
        """Generate the sheet using the template"""
        try:
            # Load template
            template_wb = load_workbook(self.template_path)
            sheet = template_wb.active
            
            # Fill in the basic information
            self.safe_set_cell_value(sheet, 'C12', self.wo_number)
            self.safe_set_cell_value(sheet, 'E12', "5R234 Plumtree Ln Pt 2")

            # Process source data
            data_sheet = self.source_workbook.active
            current_row = 16  # Start after headers
            
            # Process poles
            pole_counts = self.process_poles(data_sheet)
            for pole_type, count in pole_counts.items():
                sheet.cell(row=current_row, column=1, value=current_row - 15)  # Item number
                sheet.cell(row=current_row, column=2, value='')  # Stock Item #
                sheet.cell(row=current_row, column=3, value=count)
                sheet.cell(row=current_row, column=4, value='EA')
                sheet.cell(row=current_row, column=5, value=pole_type)
                current_row += 1

            # Process conductors
            conductor_data = self.process_conductors(data_sheet)
            for conductor_type, length in conductor_data.items():
                sheet.cell(row=current_row, column=1, value=current_row - 15)
                sheet.cell(row=current_row, column=2, value='')
                sheet.cell(row=current_row, column=3, value=round(length))
                sheet.cell(row=current_row, column=4, value='FT')
                sheet.cell(row=current_row, column=5, value=conductor_type)
                current_row += 1

            # Process guys and anchors
            num_spools, num_anchors = self.process_guys_and_anchors(data_sheet)
            if num_spools > 0:
                sheet.cell(row=current_row, column=1, value=current_row - 15)
                sheet.cell(row=current_row, column=2, value='6155402')
                sheet.cell(row=current_row, column=3, value=num_spools)
                sheet.cell(row=current_row, column=4, value='EA')
                sheet.cell(row=current_row, column=5, value="250' coil of 7/16\" UG guy wire")
                sheet.cell(row=current_row, column=10, value=250.54)
                sheet.cell(row=current_row, column=11, value=f'=C{current_row}*J{current_row}')
                current_row += 1

            # Save the workbook
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f"Long_Lead_Items_{self.wo_number}.xlsx"
            )
            if file_path:
                template_wb.save(file_path)
                messagebox.showinfo("Success", f"File saved successfully as {file_path}")
                return file_path

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            raise