import os
import shutil
import comtypes.client
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image
import pytesseract
from pytesseract import Output
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

# Performs OCR (Optical Character Recognition) on each image file and extracts text along with layout positions.
def ocr_images_with_layout(image_path):
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    img = Image.open(image_path)
    data = pytesseract.image_to_data(img, output_type=Output.DICT)

    blocks = []
    current_block = []

    n_boxes = len(data['level'])
    for i in range(n_boxes):
        if int(data['conf'][i]) > 60:  # Confidence threshold
            (x, y, w, h, text) = (data['left'][i], data['top'][i], data['width'][i], data['height'][i], data['text'][i])
            if text.strip():
                current_block.append((x, y, w, h, text))
    
    blocks = cluster_text_blocks(current_block)
    return blocks

# Cluster text into separate blocks based on vertical spacing.
def cluster_text_blocks(text_elements, spacing_threshold=20):
    sorted_elements = sorted(text_elements, key=lambda el: el[1])
    clusters = []
    current_cluster = []
    last_y = None

    for element in sorted_elements:
        _, y, _, h, _ = element
        if last_y is not None and (y - (last_y + h)) > spacing_threshold:
            if current_cluster:
                clusters.append(current_cluster)
                current_cluster = []
        current_cluster.append(element)
        last_y = y

    if current_cluster:
        clusters.append(current_cluster)

    return clusters

# Creates a new PowerPoint slide with OCR-extracted text placed based on original positions.
def create_layout_slide(prs, clusters):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    for cluster in clusters:
        texts = ' '.join([el[4] for el in cluster])
        x, y, _, _, _ = cluster[0]
        textbox = slide.shapes.add_textbox(Inches(x / 96), Inches(y / 96), Inches(4), Inches(1))
        frame = textbox.text_frame
        p = frame.add_paragraph()
        p.text = texts
        p.font.size = Pt(12)

# Main pipeline for OCR automation with layout preservation.
def process_ppt_file(input_ppt):
    base_name = os.path.splitext(input_ppt)[0]
    temp_dir = base_name + "_temp_images"
    output_ppt = base_name + "_ocr_output.pptx"

    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

    os.makedirs(temp_dir, exist_ok=True)

    ppt_to_images(input_ppt, temp_dir)
    convert_svg_images(temp_dir)

    prs = Presentation()

    image_files = sorted([file for file in os.listdir(temp_dir) if file.lower().endswith('.png')])
    for img_file in image_files:
        img_path = os.path.join(temp_dir, img_file)
        clusters = ocr_images_with_layout(img_path)
        create_layout_slide(prs, clusters)

    prs.save(output_ppt)
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
