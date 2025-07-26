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
    # Initialize PowerPoint COM object.
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    # Open the specified PowerPoint file.
    presentation = powerpoint.Presentations.Open(input_ppt)

    # Export slides to PNG images into the specified directory.
    presentation.Export(output_dir, "PNG")

    # Close PowerPoint application.
    presentation.Close()
    powerpoint.Quit()

# Converts SVG images to PNG format for OCR compatibility.
def convert_svg_images(image_dir):
    # Iterate over files in the specified directory.
    for file in os.listdir(image_dir):
        if file.lower().endswith('.svg'):
            svg_path = os.path.join(image_dir, file)
            png_path = os.path.join(image_dir, file[:-4] + '.png')

            # Perform SVG to PNG conversion.
            svg2png(url=svg_path, write_to=png_path)

            # Remove the original SVG file after conversion.
            os.remove(svg_path)

# Performs OCR (Optical Character Recognition) on each image file and extracts text.
def ocr_images(image_dir):
    extracted_texts = []

    # List and sort PNG images to maintain slide order.
    image_files = sorted([
        file for file in os.listdir(image_dir) 
        if file.lower().endswith('.png')
    ])

    # Perform OCR on each image and store the extracted text.
    for img in image_files:
        img_path = os.path.join(image_dir, img)

        # Use pytesseract to extract text from image.
        text = pytesseract.image_to_string(Image.open(img_path))
        extracted_texts.append(text.strip())

    return extracted_texts

# Creates a new PowerPoint file with each OCR-extracted text as a slide.
def create_text_ppt(texts, output_ppt):
    prs = Presentation()

    # Create slides with OCR-extracted text.
    for text in texts:
        # Add a new blank slide.
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        # Define textbox dimensions and position.
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))

        # Add extracted text to textbox.
        frame = textbox.text_frame
        frame.text = text

    # Save the new PowerPoint file.
    prs.save(output_ppt)

# Main pipeline for OCR automation testing.
def test_ocr_pipeline(input_ppt, temp_dir, output_ppt):
    # Clean existing temporary directory if present.
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

    # Create a fresh temporary directory for intermediate files.
    os.makedirs(temp_dir, exist_ok=True)

    # Step 1: Export slides from input PPT to images.
    ppt_to_images(input_ppt, temp_dir)

    # Step 2: Convert SVG images to PNG format (if applicable).
    convert_svg_images(temp_dir)

    # Step 3: Perform OCR to extract text from images.
    texts = ocr_images(temp_dir)

    # Step 4: Generate new PPT with OCR-extracted texts.
    create_text_ppt(texts, output_ppt)

    # Clean up temporary images directory after processing.
    shutil.rmtree(temp_dir)

# Entry point of the script for direct execution.
if __name__ == "__main__":
    input_ppt = "test_presentation.pptx"    # Input PowerPoint filename.
    temp_dir = "temp_images"                # Directory for temporary image files.
    output_ppt = "ocr_output.pptx"          # Output PowerPoint filename.

    # Execute the OCR pipeline test.
    test_ocr_pipeline(input_ppt, temp_dir, output_ppt)
