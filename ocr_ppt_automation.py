import os
import shutil
import comtypes.client
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import pytesseract
from cairosvg import svg2png

# Converts each slide of a PowerPoint file into individual PNG images.
def ppt_to_images(input_ppt, output_dir):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    presentation = powerpoint.Presentations.Open(input_ppt)
    presentation.Export(output_dir, "PNG")
    presentation.Close()
    powerpoint.Quit()

# Converts SVG images to PNG format for OCR compatibility.
def convert_svg_images(image_dir):
    for file in os.listdir(image_dir):
        if file.lower().endswith('.svg'):
            svg_path = os.path.join(image_dir, file)
            png_path = os.path.join(image_dir, file[:-4] + '.png')
            svg2png(url=svg_path, write_to=png_path)
            os.remove(svg_path)

# Performs OCR (Optical Character Recognition) on each image file and extracts text.
def ocr_images(image_dir):
    extracted_texts = []
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    image_files = sorted([file for file in os.listdir(image_dir) if file.lower().endswith('.png')])
    for img in image_files:
        img_path = os.path.join(image_dir, img)
        text = pytesseract.image_to_string(Image.open(img_path))
        extracted_texts.append(text.strip())
    return extracted_texts

# Creates a new PowerPoint file with each OCR-extracted text as a slide.
def create_text_ppt(texts, output_ppt):
    prs = Presentation()
    for text in texts:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        frame = textbox.text_frame
        frame.text = text
    prs.save(output_ppt)

# Main pipeline for OCR automation.
def process_ppt_file(input_ppt):
    base_name = os.path.splitext(input_ppt)[0]
    temp_dir = base_name + "_temp_images"
    output_ppt = base_name + "_ocr_output.pptx"

    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

    os.makedirs(temp_dir, exist_ok=True)

    ppt_to_images(input_ppt, temp_dir)
    convert_svg_images(temp_dir)
    texts = ocr_images(temp_dir)
    create_text_ppt(texts, output_ppt)

    shutil.rmtree(temp_dir)

# Trawls directories recursively and processes all PPTX files.
def process_all_ppts(root_dir):
    for subdir, dirs, files in os.walk(root_dir):
        for file in files:
            if file.lower().endswith('.pptx'):
                input_ppt = os.path.join(subdir, file)
                process_ppt_file(input_ppt)

# Entry point of the script.
if __name__ == "__main__":
    current_dir = os.getcwd()
    process_all_ppts(current_dir)
