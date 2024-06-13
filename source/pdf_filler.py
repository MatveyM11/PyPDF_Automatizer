import io  # Import the input/output library to handle in-memory data streams.
import os  # Import the library to interact with the operating system, such as listing files in a directory.
import time  # Import the library to handle time-related operations, such as pauses in script execution.
import pdfrw  # Import pdfrw to read and manipulate PDF files.
from reportlab.pdfgen import canvas  # Import the canvas class from reportlab to generate PDF content.
import openpyxl  # Import openpyxl to read and write Excel files.

def run(template_path, output_filename, path_to_watch):
    processed_files = []  # List to keep track of already processed files.

    # Infinite loop constantly watching the specified directory for new Excel files.
    while True:
        for filename in os.listdir(path_to_watch):
            if filename.endswith('.xlsx') and filename not in processed_files:
                filepath = os.path.join(path_to_watch, filename)
                print(f"Detected new Excel file: {filepath}")
                process_file(filepath, template_path, output_filename)
                processed_files.append(filename)  # Add the file to the list of processed files.
        time.sleep(5)  # Wait 5 seconds before checking the directory again.

def process_file(filepath, template_path, output_filename):
    # Read data from the Excel file.
    data_from_excel = read_data_from_excel(filepath)
    # Generate PDF canvas data using the read data.
    canvas_data = get_overlay_canvas(data_from_excel)
    # Merge the canvas data with the PDF template.
    form = merge(canvas_data, template_path=template_path)
    # Save the generated PDF with the specified filename.
    save(form, filename=output_filename)
    print(f"Processed and saved PDF: {output_filename}")

def read_data_from_excel(filepath):
    # Open the Excel file.
    wb = openpyxl.load_workbook(filepath)
    sheet = wb.active  # Select the active sheet (usually the first sheet).
    # Read the values from the specified cells and store them in a dictionary.
    data = {
        'first_name': sheet['B2'].value or '',   # First name
        'last_name': sheet['B3'].value or '',    # Last name
        'company': sheet['B4'].value or '',      # Company
        'country': sheet['B5'].value or '',      # Country
        'address1': sheet['B6'].value or '',     # Address 1
        'address2': sheet['B7'].value or '',     # Address 2
        'city': sheet['B8'].value or '',         # City
        'province': sheet['B9'].value or '',     # Province
        'postal_code': sheet['B10'].value or '', # Postal code
        'phone': sheet['B11'].value or '',       # Phone
        'email': sheet['B12'].value or '',       # Email
        'renewal': sheet['B13'].value or '',     # Renewal
        'renewal_type': sheet['B14'].value or '',# Renewal type
        'dob': sheet['B15'].value or ''          # Date of birth
    }
    return data

def get_overlay_canvas(data) -> io.BytesIO:
    data_io = io.BytesIO()  # Create an in-memory data stream.
    pdf = canvas.Canvas(data_io)  # Create a new PDF canvas.

    # Page 1
    pdf.drawString(x=315, y=720, text=data['first_name']) # Write the first name.
    pdf.drawString(x=33, y=720, text=data['last_name'])   # Write the last name.
    pdf.drawString(x=455, y=720, text=data['dob'])        # Write the date of birth.
    pdf.drawString(x=33, y=694, text=f"{data['address1']}, {data['address2']}") # Write the address.
    pdf.drawString(x=405, y=694, text=data['city'])       # Write the city.
    pdf.drawString(x=33, y=670, text=data['province'])    # Write the province.
    pdf.drawString(x=143, y=670, text=data['phone'])      # Write the phone.
    pdf.drawString(x=343, y=670, text=data['email'])      # Write the email.

    pdf.showPage()  # Move to the next page.

    # Page 2
    pdf.showPage()  # Move to the next page.

    # Page 3
    pdf.drawString(x=160, y=403, text=f"{data['first_name']} {data['last_name']}") # Write the first and last name.

    pdf.save()  # Save the PDF canvas.
    data_io.seek(0)  # Reset the pointer of the data stream to the beginning.
    return data_io

def merge(overlay_canvas: io.BytesIO, template_path: str) -> io.BytesIO:
    # Read the PDF template.
    template_pdf = pdfrw.PdfReader(template_path)
    # Read the generated PDF canvas data.
    overlay_pdf = pdfrw.PdfReader(overlay_canvas)
    # Merge the template pages with the canvas pages.
    for page, data in zip(template_pdf.pages, overlay_pdf.pages):
        overlay = pdfrw.PageMerge().add(data)[0]
        pdfrw.PageMerge(page).add(overlay).render()
    form = io.BytesIO()  # Create a new in-memory data stream for the PDF form.
    pdfrw.PdfWriter().write(form, template_pdf)  # Write the merged PDF to the data stream.
    form.seek(0)  # Reset the pointer of the data stream to the beginning.
    return form

def save(form: io.BytesIO, filename: str):
    # Save the data stream of the PDF form to a file.
    with open(filename, 'wb') as f:
        f.write(form.read())

if __name__ == '__main__':
    path_to_watch = '.'  # Directory to watch.
    template_path = './Solicitud de carnet GVA.pdf'  # Path to the PDF template.
    output_filename = 'merged.pdf'  # Filename of the generated PDF.

    print(f"Watching for new Excel files in: {path_to_watch}")
    run(template_path, output_filename, path_to_watch)  # Execute the main function.
