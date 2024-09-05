import openpyxl

# Load the Excel template file
def load_template(template_file):
    """Load the Excel template file"""
    wb = openpyxl.load_workbook(template_file)
    return wb

# Fill out the template with data
def fill_template(wb, data):
    """Fill out the template with data"""
    # Select the first sheet
    sheet = wb.active
    
    # Iterate over the data and fill out the template
    for key, value in data.items():
        # Find the cell with the matching key
        cell = sheet.cell(row=1, column=1)
        while cell.value != key:
            cell = cell.offset(row=0, column=1)
        
        # Fill out the cell with the value
        cell.offset(row=1, column=0).value = value
    
    return wb

# Save the filled-out template to a new file
def save_template(wb, output_file):
    """Save the filled-out template to a new file"""
    wb.save(output_file)

# Example usage
template_file = "template.xlsx"
output_file = "filled_template.xlsx"
data = {
    "Name": "John Doe",
    "Age": 30,
    "City": "New York"
}

wb = load_template(template_file)
wb = fill_template(wb, data)
save_template(wb, output_file)
