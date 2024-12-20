from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side
from tkinter import filedialog, messagebox

class LocateSheetGenerator:
    def __init__(self, workbook, wo_number, county, city_place):
        """Initialize the generator with workbook and location information"""
        self.workbook = workbook
        self.wo_number = wo_number
        self.county = county
        self.city_place = city_place
        
        # Default values for locate sheet
        self.defaults = {
            'cc_number': 11151,
            'equipment_used': 'Auger',
            'work_being_done_for': '(Utility) Pacific Power',
            'type_of_work': 'Pole Replacement',
            'directional_drilling': 'No',
            'using_equipment': 'Y',
            'within_overhead_line': 'Y',
            'location_of_work': "Locate a 30' radius around pole",
            'comments': 'Pole is marked in white paint',
            'township': '35S',
            'range': '06W',
            'section': '26',
            'quarter_section': 'SE'
        }

    def setup_header_cell(self, cell, value):
        """Apply formatting to header cells"""
        cell.value = value
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='left', vertical='center')
        cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

    def setup_data_cell(self, cell, value):
        """Apply formatting to data cells"""
        cell.value = value
        cell.alignment = Alignment(horizontal='left', vertical='center')
        cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

    def adjust_column_widths(self, worksheet):
        """Auto-adjust column widths based on content"""
        for col in worksheet.columns:
            max_length = 0
            for cell in col:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            worksheet.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

    def get_row_data(self, sequence, facility_id, latitude_longitude=""):
        """Generate a row of data for the locate sheet"""
        return [
            self.defaults['cc_number'],          # CC#
            self.wo_number,                      # WO#
            '',                                  # Pulling Section
            facility_id,                         # FP#
            sequence,                            # Seq#
            self.defaults['using_equipment'],    # Equipment extends 15' above ground?
            self.defaults['equipment_used'],     # Equipment Used
            self.defaults['within_overhead_line'], # Working within 10' of power line?
            self.defaults['directional_drilling'], # Directional Drilling
            self.defaults['type_of_work'],       # Type of work
            self.defaults['work_being_done_for'], # Work being done for
            self.county,                         # County
            self.city_place,                     # City/Place
            latitude_longitude,                  # Latitude, Longitude
            self.defaults['location_of_work'],   # Location of work
            self.defaults['comments'],           # Comments
            self.defaults['township'],           # Township
            self.defaults['range'],              # Range
            self.defaults['section'],            # Section
            self.defaults['quarter_section']     # Quarter Section
        ]

    def generate_sheet(self):
        """Generate the locate sheet"""
        try:
            # Get the first worksheet from the workbook
            data_report_sheet = self.workbook.worksheets[0]
            
            # Create new workbook and get active sheet
            output_workbook = Workbook()
            output_sheet = output_workbook.active
            output_sheet.title = "Locate Sheet"

            # Define headers
            headers = [
                'CC#', 
                'WO#', 
                'Pulling Section', 
                'FP#', 
                'Seq#',
                'Will be using equipment that extends 15\' above ground?',
                'Equipment Used',
                'Will you be working within 10\' of an overhead power line?',
                'Directional Drilling',
                'Type of work',
                'Work being done for',
                'County',
                'City/Place',
                'Latitude, Longitude',
                'Location of work',
                'Comments',
                'Township',
                'Range',
                'Section',
                'Quarter Section'
            ]
            
            # Setup headers
            for col_num, header in enumerate(headers, 1):
                cell = output_sheet.cell(row=1, column=col_num)
                self.setup_header_cell(cell, header)

            # Process data rows
            row_num = 2
            for row in data_report_sheet.iter_rows(min_row=2, values_only=True):
                sequence = row[0]
                if not sequence:  # Skip empty rows
                    continue

                # Extract data from row
                facility_id = row[1]
                latitude_longitude = row[13] if len(row) > 13 else ""

                # Get row data
                row_data = self.get_row_data(sequence, facility_id, latitude_longitude)
                
                # Write row data
                for col_num, value in enumerate(row_data, 1):
                    cell = output_sheet.cell(row=row_num, column=col_num)
                    self.setup_data_cell(cell, value)
                
                row_num += 1

            # Adjust column widths
            self.adjust_column_widths(output_sheet)

            # Save the workbook
            self.save_workbook(output_workbook, "Locate_Sheet.xlsx")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while generating the locate sheet: {str(e)}")
            raise

    def save_workbook(self, workbook, default_filename):
        """Save the workbook to a user-specified location"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_filename
        )
        if file_path:
            try:
                workbook.save(file_path)
                messagebox.showinfo("Success", f"File saved successfully as {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {str(e)}")
                raise