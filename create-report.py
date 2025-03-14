import os
import re
import comtypes.client
import PyPDF2

def ppt_to_pdf(input_file, output_file):
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"The file {input_file} does not exist.")
    
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    try:
        deck = powerpoint.Presentations.Open(input_file)
        deck.SaveAs(output_file, 32)  # FormatType = 32 for ppt to pdf
        deck.Close()
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        powerpoint.Quit()

def split_pdf(input_file, page_number, output_dir):
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"The file {input_file} does not exist.")
    
    pdf_reader = PyPDF2.PdfReader(input_file)
    pdf_writer1 = PyPDF2.PdfWriter()
    pdf_writer2 = PyPDF2.PdfWriter()

    for page in range(page_number):
        pdf_writer1.add_page(pdf_reader.pages[page])

    for page in range(page_number, len(pdf_reader.pages)):
        pdf_writer2.add_page(pdf_reader.pages[page])

    part1_path = os.path.join(output_dir, "part1.pdf")
    part2_path = os.path.join(output_dir, "part2.pdf")

    with open(part1_path, "wb") as output_pdf1:
        pdf_writer1.write(output_pdf1)

    with open(part2_path, "wb") as output_pdf2:
        pdf_writer2.write(output_pdf2)

    return part1_path, part2_path

def merge_pdfs(pdfs, output_file):
    pdf_writer = PyPDF2.PdfWriter()

    for pdf in pdfs:
        if not os.path.exists(pdf):
            raise FileNotFoundError(f"The file {pdf} does not exist.")
        
        pdf_reader = PyPDF2.PdfReader(pdf)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            pdf_writer.add_page(page)

    with open(output_file, "wb") as output_pdf:
        pdf_writer.write(output_pdf)

def clean_up(files):
    for file in files:
        if os.path.exists(file):
            os.remove(file)

def find_pdf_to_merge(pptx_dir):
    pattern1 = re.compile(r'.*ESF.*Fragebogen.*\.pdf')
    pattern2 = re.compile(r'.*Fragebogen.*ESF.*\.pdf')
    pattern3 = re.compile(r'.*ESF.*\.pdf')
    pattern4 = re.compile(r'.*EFS.*Fragebogen.*\.pdf')
    pattern5 = re.compile(r'.*Fragebogen.*EFS.*\.pdf')
    pattern6 = re.compile(r'.*EFS.*\.pdf')
    pattern7 = re.compile(r'.*Fragebogen.*\.pdf')
    
    for file in os.listdir(pptx_dir):
        if pattern1.match(file) or pattern2.match(file):
            return os.path.join(pptx_dir, file)
    
    for file in os.listdir(pptx_dir):
        if pattern3.match(file):
            return os.path.join(pptx_dir, file)
    
    for file in os.listdir(pptx_dir):
        if pattern4.match(file):
            return os.path.join(pptx_dir, file)
    
    for file in os.listdir(pptx_dir):
        if pattern5.match(file):
            return os.path.join(pptx_dir, file)
    
    for file in os.listdir(pptx_dir):
        if pattern6.match(file):
            return os.path.join(pptx_dir, file)
    
    for file in os.listdir(pptx_dir):
        if pattern7.match(file):
            return os.path.join(pptx_dir, file)
    
    return None

try:
    # Ask for file location
    pptx_dir = input("Please enter the full path to your PowerPoint file: ")
    pptx_dir = os.path.abspath(pptx_dir)

    if not os.path.isdir(pptx_dir):
        raise NotADirectoryError(f"The directory {pptx_dir} does not exist.")

    # Search for the pptx-file
    pptx_files = [f for f in os.listdir(pptx_dir) if f.endswith('.pptx')]

    if not pptx_files:
        raise FileNotFoundError("No pptx files found in the specified directory.")

    # Get first pptx-file name
    pptx_file = pptx_files[0]
    pptx_file_path = os.path.join(pptx_dir, pptx_file)

    # Create output-PDF file based on the input pptx-filename
    base_name = os.path.splitext(pptx_file)[0]
    output_pdf = os.path.join(pptx_dir, f"{base_name}_output.pdf")

    # Convert PPT to PDF
    ppt_to_pdf(pptx_file_path, output_pdf)

    # Calculate the split page number
    pdf_reader = PyPDF2.PdfReader(output_pdf)
    total_pages = len(pdf_reader.pages)
    split_page = total_pages - 3

    # Split the PDF at the calculated position
    part1_pdf, part2_pdf = split_pdf(output_pdf, split_page, pptx_dir)

    # Find the PDF to merge with
    file_to_merge_pdf = find_pdf_to_merge(pptx_dir)
    if file_to_merge_pdf is None:
        file_name_for_merging = input("Please enter the file name which should be merged with: ")
        file_to_merge_pdf = os.path.join(pptx_dir, file_name_for_merging)

    final_output_pdf = os.path.join(pptx_dir, f"{base_name}.pdf")

    merge_pdfs([part1_pdf, file_to_merge_pdf, part2_pdf], final_output_pdf)

    # Remove temp files
    clean_up([output_pdf, part1_pdf, part2_pdf])

    print(f"Report was created: {final_output_pdf}")
except Exception as e:
    print(f"An error occurred during processing: {e}")
