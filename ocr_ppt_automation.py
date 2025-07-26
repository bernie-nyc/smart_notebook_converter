import os
import shutil
import comtypes.client
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import pytesseract
from cairosvg import svg2png

# Converts PPTX slides to images
def ppt_to_images(input_ppt, output_dir):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    presentation = powerpoint.Presentations.Open(input_ppt)
    presentation.Export(output_dir, "PNG")
    presentation.Close()
    powerpoint.Quit()

# Convert SVG to PNG if needed
def convert_svg_images(image_dir):
    for file in os.listdir(image_dir):
        if file.lower().endswith('.svg'):
            svg_path = os.path.join(image_dir, file)
            png_path = os.path.join(image_dir, file[:-4] + '.png')
            svg2png(url=svg_path, write_to=png_path)
            os.remove(svg_path)

# OCR extraction from images
def ocr_images(image_dir):
    texts = []
    image_files = sorted([
        file for file in os.listdir(image_dir) 
        if file.lower().endswith('.png')
    ])
    for img in image_files:
        img_path = os.path.join(image_dir, img)
        text = pytesseract.image_to_string(Image.open(img_path))
        texts.append(text.strip())
    return texts

# Create a new PPTX with extracted texts
def create_text_ppt(texts, output_ppt):
    prs = Presentation()
    for text in texts:
        slide = prs.slides.add_slide(prs.slide_layouts[5])  # Blank layout
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        frame = textbox.text_frame
        frame.text = text
    prs.save(output_ppt)

# Main test function
def test_ocr_pipeline(input_ppt, temp_dir, output_ppt):
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir, exist_ok=True)

    ppt_to_images(input_ppt, temp_dir)
    convert_svg_images(temp_dir)
    texts = ocr_images(temp_dir)
    create_text_ppt(texts, output_ppt)
    shutil.rmtree(temp_dir)

if __name__ == "__main__":
    input_ppt = "test_presentation.pptx"
    temp_dir = "temp_images"
    output_ppt = "ocr_output.pptx"
    test_ocr_pipeline(input_ppt, temp_dir, output_ppt)
