import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
import os
import subprocess


def get_numbers_from_file(file_path, column_name):
    """
    Read numbers from a CSV or Excel file.
    Args:
        file_path (str): Path to CSV or Excel file.
        column_name (str): Column name or index containing numbers.
    Returns:
        list: List of floats extracted from the column.
    """
    print(f"Checking if file exists: {os.path.exists(file_path)}")
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            raise ValueError("Unsupported file format. Use .csv or .xlsx")
        
        # Handle column_name as index if numeric
        if isinstance(column_name, int) or column_name.isdigit():
            column_name = int(column_name)
            if column_name >= len(df.columns):
                raise ValueError(f"Column index {column_name} out of range")
        elif column_name not in df.columns:
            raise ValueError(f"Column '{column_name}' not found in file")
        
        # Extract numbers, convert to float, drop NaN, ignore non-numeric
        numbers = df[column_name].dropna().apply(
            lambda x: float(x) if str(x).replace(".", "").replace("-", "").isdigit() else None
        ).dropna().tolist()
        
        if not numbers:
            print("Warning: No valid numbers found in the specified column")
        return numbers
    
    except Exception as e:
        print(f"Error reading file: {e}")
        return []

def create_pdf_with_numbers(numbers, output_path, template_path=None, coordinates=None):
    """
    Create a new PDF or overlay numbers on an existing PDF at specified coordinates.
    Args:
        numbers (list): List of numbers to place in the PDF.
        output_path (str): Path to save the output PDF.
        template_path (str): Path to existing PDF template (optional).
        coordinates (list): List of (x, y) tuples for text placement (optional).
    """
    try:
        c = canvas.Canvas(output_path, pagesize=letter)
        
        if template_path and os.path.exists(template_path):
            # Overlay on existing PDF
            reader = PdfReader(template_path)
            c.setFont("Helvetica", 12)  # Use built-in font
            
            for page_num in range(len(reader.pages)):
                if coordinates and page_num == 0:  # Place numbers on first page
                    for i, (x, y) in enumerate(coordinates):
                        if i < len(numbers):
                            c.drawString(x, y, f"{numbers[i]:.2f}")
                c.showPage()
        else:
            # Create new PDF
            c.setFont("Helvetica", 12)
            for i, num in enumerate(numbers):
                c.drawString(100, 750 - i * 20, f"Number {i + 1}: {num:.2f}")
            c.showPage()
        
        c.save()
        print(f"PDF created successfully: {output_path}")
    
    except Exception as e:
        print(f"Error creating PDF: {e}")

def fill_pdf_form(template_path, output_path, numbers, field_mappings):
    """
    Fill form fields in a PDF with numbers.
    Args:
        template_path (str): Path to PDF template with form fields.
        output_path (str): Path to save the output PDF.
        numbers (list): List of numbers to fill in fields.
        field_mappings (list): List of form field names.
    """
    try:
        reader = PdfReader(template_path)
        writer = PdfWriter()
        
        for page in reader.pages:
            writer.add_page(page)
        
        # Update form fields if they exist
        if reader.get_fields():
            for i, field_name in enumerate(field_mappings):
                if i < len(numbers) and field_name in reader.get_fields():
                    writer.update_page_form_field_values(
                        writer.pages[0], {field_name: f"{numbers[i]:.2f}"}
                    )
        
        with open(output_path, "wb") as f:
            writer.write(f)
        print(f"PDF with filled form fields saved: {output_path}")
    
    except Exception as e:
        print(f"Error filling PDF form: {e}")

def main():
    # Configuration - EDIT THESE VALUES FOR YOUR PROJECT
    
    print(f"Current working directory: {os.getcwd()}")
    file_path = "/home/ajc/JROTC/Testing.xlsx"
    print(f"File path being checked: {file_path}")
    print(f"Checking if file exists: {os.path.exists(file_path)}")

    import pandas as pd
    df = pd.read_excel(file_path)
    print(df.columns)
    
    column_name = 1  # Column name or index (e.g., "B", 1)
    output_pdf = "output.pdf"  # Output PDF file name
    template_pdf = "blank.pdf"  # Path to PDF template (set to None if not used)
    coordinates = [(100, 700), (100, 650), (100, 600)]  # (x, y) coordinates for text overlay
    form_fields = ["field1", "field2", "field3"]  # Form field names (if PDF has fields)
    
    # Get numbers from file
    numbers = get_numbers_from_file(file_path, column_name)
    if not numbers:
        print("No numbers to process. Exiting.")
        return
    print(f"Extracted numbers: {numbers}")
    
    
    # Create or overlay PDF
    if template_pdf and os.path.exists(template_pdf):
        reader = PdfReader(template_pdf)
        if reader.get_fields():
            # PDF has form fields, fill them
            fill_pdf_form(template_pdf, output_pdf, numbers, form_fields)
        else:
            # No form fields, overlay text
            create_pdf_with_numbers(numbers, output_pdf, template_pdf, coordinates)
    else:
        # No template, create new PDF
        create_pdf_with_numbers(numbers, output_pdf)

    try:
        subprocess.run(["evince", output_pdf], check=True)
        print(f"Opened {output_pdf} with Evince")
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("Evince failed or not installed. Trying xdg-open...")
        try:
            subprocess.run(["xdg-open", output_pdf], check=True)
            print(f"Opened {output_pdf} with default PDF viewer")
        except subprocess.CalledProcessError as e:
            print(f"Error opening PDF: {e}")
        except FileNotFoundError:
            print("Error: xdg-open not found. Please install xdg-utils or evince.")
script = input("Input script")
if script == "1":
    main()

