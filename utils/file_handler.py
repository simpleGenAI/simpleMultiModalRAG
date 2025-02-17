import os
import subprocess
import fitz  # PyMuPDF

def ppt_to_pdf_unoconv(input_ppt):
    subprocess.run(["unoconv", "-f", "pdf", input_ppt], check=True)

def pdf_to_images_fitz(pdf_path, output_folder):
    """
    Convert PDF pages to images using PyMuPDF.
    
    :param pdf_path: Path to the input PDF file.
    :param output_folder: Directory to save images.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder, exist_ok=True)
    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap()
        image_path = os.path.join(output_folder, f"page_{page_num + 1:03d}.png")
        pix.save(image_path)
        print(f"Saved: {image_path}")

def ppt_to_images(ppt_path):
    """
    Convert a PowerPoint file to images.
    The process is as follows:
      1. Convert the PPT to a PDF using unoconv.
      2. Convert the resulting PDF to images using PyMuPDF.
    The output images are stored in a directory named after the original PPT filename (without extension).

    :param ppt_path: Path to the PowerPoint (.ppt/.pptx) file.
    """
    # Extract the base filename (without extension) and directory.
    ppt_dir = os.path.dirname(ppt_path)
    base_name = os.path.splitext(os.path.basename(ppt_path))[0]
    
    # Construct the PDF path. Assumes unoconv creates the PDF in the same directory.
    pdf_path = os.path.join(ppt_dir, f"{base_name}.pdf")
    
    # Convert PPT to PDF.
    ppt_to_pdf_unoconv(ppt_path)
    
    # Define output folder for images (named after the original PPT file).
    output_folder = os.path.join(ppt_dir, base_name)
    
    # Convert the resulting PDF to images.
    pdf_to_images_fitz(pdf_path, output_folder)
