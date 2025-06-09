import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import re
from PIL import Image
import io
import cv2
import numpy as np
from pdf2image import convert_from_path
import tempfile
import subprocess

def parse_word_document(doc_path):
    """Parse the Word document and extract questions with their arrangements and directions"""
    doc = Document(doc_path)
    
    content_blocks = []
    current_direction = None
    pending_arrangement = None  # Store arrangement for next question
    
    # Track if we're inside an arrangement block
    in_arrangement = False
    
    for para in doc.paragraphs:
        text = para.text.strip()
        
        if not text:
            continue
            
        # Check if it's a direction
        if "Directions for questions" in text or "DIRECTIONS:" in text:
            # Start new direction
            current_direction = text
            
        # Check if it's an arrangement line
        elif "arrangement is as follows" in text.lower() or "final arrangement" in text.lower():
            in_arrangement = True
            pending_arrangement = text + '\n'
            
        # Check if it's the actual arrangement (underlined text, special format, or circular diagram)
        elif in_arrangement:
            # Check for various arrangement patterns
            if (text.startswith('[') or  # Underlined arrangement
                '>' in text[:5] or  # Comparison arrangement
                re.match(r'^[A-Z]\s+[A-Z]', text) or  # Letter arrangement
                re.match(r'^[A-Z]\s+\>', text) or  # Ranking arrangement
                '_' in text or  # Underlined format
                text.startswith('**') or  # Bold text in markdown format
                re.match(r'^[A-Z]+\s+[A-Z]+', text) or  # Multiple letter arrangement
                any(c in text for c in ['→', '↑', '↓', '←'])):  # Directional arrows
                
                # Clean up markdown bold formatting if present
                clean_text = text.replace('**', '').strip()
                if pending_arrangement:
                    pending_arrangement += clean_text + '\n'
                else:
                    pending_arrangement = clean_text + '\n'
                    
                # Check if this is likely the end of arrangement
                if not text.endswith(','):
                    in_arrangement = False
                    
        # Check if it's a question number
        elif re.match(r'^\*?\*?\d+\.', text):
            # Start new question with pending arrangement
            content_blocks.append({
                'type': 'question',
                'direction': current_direction,
                'arrangement': pending_arrangement,
                'content': text,
                'options': []
            })
            pending_arrangement = None  # Reset for next question
            
        # If we have a current question, check if this is an option
        elif content_blocks and content_blocks[-1]['type'] == 'question':
            # Check if this line is an option (starts with number in parentheses)
            if re.match(r'^\(\d+\)', text) or re.match(r'^\\\(\d+\\\)', text):
                content_blocks[-1]['options'].append(text)
            # Otherwise, append to question text
            elif not (text.startswith('-----') or 
                    text == 'Person' or text == 'Game' or text == 'Color' or
                    text == 'City' or text == 'Car' or text == 'Country' or
                    text == 'Fruit' or text.startswith('===') or 
                    '|' in text and len(text.split('|')) > 2):
                content_blocks[-1]['content'] += '\n' + text
    
    return content_blocks

def extract_images_from_docx(doc_path, output_dir="extracted_images"):
    """Extract images and diagrams from the Word document"""
    os.makedirs(output_dir, exist_ok=True)
    extracted_images = []
    
    try:
        # Convert DOCX to PDF for better image extraction
        with tempfile.TemporaryDirectory() as tmpdir:
            # Use LibreOffice or similar to convert
            pdf_path = os.path.join(tmpdir, "temp.pdf")
            subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, doc_path], 
                         capture_output=True)
            
            if os.path.exists(pdf_path):
                # Convert PDF pages to images
                images = convert_from_path(pdf_path, dpi=300)
                
                for i, img in enumerate(images):
                    # Convert PIL Image to numpy array for OpenCV
                    img_np = np.array(img)
                    
                    # Convert RGB to BGR for OpenCV
                    img_cv = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
                    
                    # Extract diagrams/charts
                    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
                    blur = cv2.GaussianBlur(gray, (5, 5), 0)
                    _, thresh = cv2.threshold(blur, 200, 255, cv2.THRESH_BINARY_INV)
                    
                    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                    
                    for j, cnt in enumerate(contours):
                        x, y, w, h = cv2.boundingRect(cnt)
                        area = w * h
                        
                        # Filter for significant shapes (diagrams)
                        if area > 10000 and w > 100 and h > 100:
                            # Determine padding based on shape
                            peri = cv2.arcLength(cnt, True)
                            approx = cv2.approxPolyDP(cnt, 0.04 * peri, True)
                            aspect_ratio = float(w) / h
                            
                            if len(approx) >= 6 and 0.8 < aspect_ratio < 1.2:
                                # Likely circular diagram
                                pad = 50
                            else:
                                # Rectangular or other shape
                                pad = 20
                            
                            # Crop with padding
                            x1 = max(x - pad, 0)
                            y1 = max(y - pad, 0)
                            x2 = min(x + w + pad, img_cv.shape[1])
                            y2 = min(y + h + pad, img_cv.shape[0])
                            
                            cropped = img_cv[y1:y2, x1:x2]
                            
                            # Save the extracted image
                            img_path = os.path.join(output_dir, f"diagram_{i+1}_{j+1}.png")
                            cv2.imwrite(img_path, cropped)
                            
                            # Store metadata about the image
                            extracted_images.append({
                                'path': img_path,
                                'page': i + 1,
                                'position': y,  # Vertical position for mapping to questions
                                'type': 'circular' if len(approx) >= 6 and 0.8 < aspect_ratio < 1.2 else 'rectangular'
                            })
    
    except Exception as e:
        print(f"Warning: Could not extract images using PDF conversion: {e}")
        # Fallback: try to extract embedded images directly from DOCX
        doc = Document(doc_path)
        for i, rel in enumerate(doc.part.rels.values()):
            if "image" in rel.reltype:
                img_data = rel.target_part.blob
                img_path = os.path.join(output_dir, f"embedded_{i+1}.png")
                with open(img_path, 'wb') as f:
                    f.write(img_data)
                extracted_images.append({
                    'path': img_path,
                    'page': 0,
                    'position': 0,
                    'type': 'embedded'
                })
    
    return extracted_images

def create_arrangement_slide(prs, arrangement_text, direction_text=None, num=0):
    """Create a slide with just the arrangement text"""
    slide_layout = prs.slide_layouts[5]  # Blank slide
    slide = prs.slides.add_slide(slide_layout)
    
    # Set black background
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 0, 0)
    
    # Calculate positioning
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    
    text_left = slide_width * 0.4
    text_width = slide_width * 0.58
    text_top = slide_height * 0.35  # Center vertically
    text_height = slide_height * 0.3
    
    # Add text box
    textbox = slide.shapes.add_textbox(
        left=text_left,
        top=text_top,
        width=text_width,
        height=text_height
    )
    
    text_frame = textbox.text_frame
    text_frame.clear()
    text_frame.word_wrap = True

    for shape in slide.shapes:
        if shape == slide.shapes.title:
            sp = shape
            slide.shapes._spTree.remove(sp._element)
            break

    # Add direction if exists
    if direction_text and num == 0:
        p = text_frame.add_paragraph()
        p.text = direction_text
        p.font.name = 'Arial'
        p.font.size = Pt(25)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.font.bold = True
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(12)
    
    # Add arrangement text
    p = text_frame.add_paragraph()
    p.text = arrangement_text.strip()
    p.font.name = 'Arial'
    p.font.size = Pt(25)
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.font.bold = True
    p.alignment = PP_ALIGN.LEFT
    
    return slide

def create_question_slide(prs, question_data, diagram_path=None):
    """Create a slide with the MCQ question"""
    slide_layout = prs.slide_layouts[5]  # Blank slide
    slide = prs.slides.add_slide(slide_layout)
    
    # Set black background
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 0, 0)
    
    # Calculate positioning
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    
    text_left = slide_width * 0.4
    text_width = slide_width * 0.58
    text_top = Inches(1)
    text_height = slide_height - Inches(1.5)
    
    # If there's a diagram, adjust layout
    if diagram_path and os.path.exists(diagram_path):
        # Add the diagram in the left 40% area
        try:
            left = slide_width * 0.05
            top = slide_height * 0.25
            width = slide_width * 0.3
            
            slide.shapes.add_picture(diagram_path, left, top, width=width)
        except Exception as e:
            print(f"Warning: Could not add diagram: {e}")
    
    # Add text box
    textbox = slide.shapes.add_textbox(
        left=text_left,
        top=text_top,
        width=text_width,
        height=text_height
    )
    
    text_frame = textbox.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    
    # Add question number and text
    p = text_frame.add_paragraph()
    p.text = question_data['content']
    p.font.name = 'Arial'
    p.font.size = Pt(25)
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.LEFT
    p.space_after = Pt(8)
    
    for shape in slide.shapes:
        if shape == slide.shapes.title:
            sp = shape
            slide.shapes._spTree.remove(sp._element)
            break
    
    # Add options
    for option in question_data['options']:
        p = text_frame.add_paragraph()
        p.text = option
        p.font.name = 'Arial'
        p.font.size = Pt(25)
        p.font.color.rgb = RGBColor(255, 255, 0)  # Yellow for options
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(4)

    return slide


def convert_word_to_ppt(word_path, ppt_path):
    """Main function to convert Word document to PowerPoint"""
    # Parse Word document
    print("Parsing Word document...")
    questions = parse_word_document(word_path)
    
    # Extract images from document
    print("Extracting diagrams and images...")
    extracted_images = extract_images_from_docx(word_path)
    
    # Create PowerPoint presentation
    print("Creating PowerPoint presentation...")
    prs = Presentation()
    
    # Set slide size to 16:9
    prs.slide_width = Inches(16)
    prs.slide_height = Inches(9)
    
    # Track the current direction for all slides
    current_direction = None
    
    # Create slides for each question
    slide_count = 0
    for i, question in enumerate(questions):
        # Update direction if it changes
        if question['direction'] and question['direction'] != current_direction:
            current_direction = question['direction']
        
        # If there's an arrangement, create a separate slide for it
        if question.get('arrangement'):
            print(f"Creating arrangement slide {slide_count + 1}...")
            create_arrangement_slide(prs, question['arrangement'], current_direction, i)
            slide_count += 1
        
        # Create the question slide
        print(f"Creating question slide {slide_count + 1}...")
        
        # Try to find a relevant diagram for this question
        diagram_path = None
        # Simple heuristic: if question mentions "circular" or specific question numbers
        # that typically have diagrams (like questions 8-14 for circular arrangements)
        question_num_match = re.search(r'^(\d+)\.', question['content'])
        if question_num_match:
            q_num = int(question_num_match.group(1))
            # Questions 8-14 typically have circular diagrams
            if 8 <= q_num <= 14:
                # Find circular diagram if available
                for img in extracted_images:
                    if img['type'] == 'circular':
                        diagram_path = img['path']
                        break
        
        create_question_slide(prs, question, diagram_path)
        slide_count += 1
    
    # Save presentation
    print(f"Saving presentation to {ppt_path}...")
    prs.save(ppt_path)
    print(f"Conversion complete! Created {slide_count} slides.")
    
    # Cleanup extracted images if needed
    if extracted_images and os.path.exists("extracted_images"):
        import shutil
        shutil.rmtree("extracted_images")

# Main execution
if __name__ == "__main__":
    # Input and output file paths
    input_file = "mcq2.docx"
    output_file = os.path.join("output2.pptx")
    
    # Check if input file exists
    convert_word_to_ppt(input_file, output_file)