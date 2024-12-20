import re
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side
from tkinter import filedialog, messagebox

class NewFramingGenerator:
    def __init__(self, workbook):
        self.workbook = workbook
        self.setup_standards()

    def setup_standards(self):
        """Initialize the standards dictionary with all framing types"""
        self.standards = {
            'EH101': 'EH101 - 1PH TAN',
            'EH106': 'EH106 - 1PH ANGLE',
            'EH111': 'EH111 - 1PH DDE',
            'EH121': 'EH121 - 1PH CORNER',
            'EH131': 'EH131 - 1PH DE',
            'EH221': 'EH221 - 1PH TAP',
            'EH226': 'EH226 - 1PH TAP FROM 2PH LINE',
            'EH231': 'EH231 - 1PH TAP FOR FUSING',
            'EH341': 'EH341 - 1PH HORIZIONTAL DE',
            'EH421': 'EH421 - 1PH TAP WITH CROSSARM',
            'EJ300': 'EJ300 - 3PH TAN/ANGLE/DDE/DE/CORNER',
            'EJ909': 'EJ909 - 3PH TAN UNDERBUILD',
            'TF200': 'TF200 - TRANSMISSION TANGENT'
        }

    def get_standard_code(self, standard):
        """Extract standard code from the full standard string"""
        if not standard:
            return ''
        match = re.search(r'(EH|EJ|TF)\d{3}', str(standard))
        return match.group() if match else ''

    def get_new_framing(self, primary, secondary):
        """Determine new framing based on primary and secondary standards"""
        if not primary:
            return ''
        if 'TF200' in str(primary):
            return 'TF200/EJ909'
        return primary

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

    def generate_sheet(self):
        """Generate the new framing sheet"""
        try:
            # Get the first worksheet from the workbook
            data_report_sheet = self.workbook.worksheets[0]
            
            # Create new workbook and get active sheet
            output_workbook = Workbook()
            output_sheet = output_workbook.active
            output_sheet.title = "New Framing Sheet"

            # Define headers
            headers = [
                'Sequence',
                'Facility ID',
                'New Framing',
                'Transmission Framing',
                'Primary Framing Standard',
                'Secondary Framing Standard'
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
                primary_standard = row[9] if len(row) > 9 else None
                secondary_standard = row[10] if len(row) > 10 else None

                # Handle transmission framing
                transmission_framing = ''
                if primary_standard and 'TF200' in str(primary_standard):
                    transmission_framing = 'TF200 - TRANSMISSION TANGENT'
                    primary_standard = 'EJ909 - 3PH TAN UNDERBUILD'

                # Get and format standards
                primary_code = self.get_standard_code(primary_standard) if primary_standard else ''
                secondary_code = self.get_standard_code(secondary_standard) if secondary_standard else ''
                primary_formatted = self.standards.get(primary_code, primary_standard)
                secondary_formatted = self.standards.get(secondary_code, secondary_standard)

                # Get new framing
                new_framing = self.get_new_framing(primary_standard, secondary_standard)

                # Prepare row data
                row_data = [
                    sequence,
                    facility_id,
                    new_framing,
                    transmission_framing,
                    primary_formatted,
                    secondary_formatted
                ]
                
                # Write row data
                for col_num, value in enumerate(row_data, 1):
                    cell = output_sheet.cell(row=row_num, column=col_num)
                    self.setup_data_cell(cell, value)
                
                row_num += 1

            # Adjust column widths
            self.adjust_column_widths(output_sheet)

            # Save the workbook
            self.save_workbook(output_workbook, "New_Framing_Sheet.xlsx")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while generating the framing sheet: {str(e)}")
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