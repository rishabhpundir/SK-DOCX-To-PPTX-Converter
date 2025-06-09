import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import re

def parse_word_document(doc_path):
    """Parse the Word document and extract questions with their directions"""
    doc = Document(doc_path)
    
    content_blocks = []
    current_direction = None
    current_question = None
    
    for para in doc.paragraphs:
        text = para.text.strip()
        
        if not text:
            continue
            
        # Check if it's a direction
        if "Directions for questions" in text or "DIRECTIONS:" in text:
            # Save previous question if exists
            if current_question:
                content_blocks.append({
                    'type': 'question',
                    'direction': current_direction,
                    'content': current_question
                })
                current_question = None
            
            # Start new direction
            current_direction = text
            # Get the direction details from next paragraph
            
        # Check if it's a question number
        elif re.match(r'^\d+\.', text):
            # Save previous question if exists
            if current_question:
                content_blocks.append({
                    'type': 'question',
                    'direction': current_direction,
                    'content': current_question
                })
            
            # Start new question
            current_question = text
            
        # If we're in a question, append the options
        elif current_question:
            current_question += '\n' + text
    
    # Don't forget the last question
    if current_question:
        content_blocks.append({
            'type': 'question',
            'direction': current_direction,
            'content': current_question
        })
    
    return content_blocks

def create_slide(prs, question_data):
    """Create a slide with the MCQ question"""
    slide_layout = prs.slide_layouts[5]  # Blank slide
    slide = prs.slides.add_slide(slide_layout)
    
    # Set black background
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 0, 0)
    
    # Calculate positioning (60% of slide width, starting from 40%)
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    
    text_left = slide_width * 0.4  # Start at 40% from left
    text_width = slide_width * 0.58  # Use 58% width (leaving 2% margin)
    text_top = Inches(1.5)  # Start 1 inch from top
    text_height = slide_height - Inches(1.5)  # Leave some bottom margin
    
    # Add text box
    textbox = slide.shapes.add_textbox(
        left=text_left,
        top=text_top,
        width=text_width,
        height=text_height
    )
    
    text_frame = textbox.text_frame
    text_frame.clear()  # Clear default text
    text_frame.word_wrap = True
    
    # Add direction if exists
    if question_data['direction']:
        p = text_frame.add_paragraph()
        p.text = question_data['direction']
        p.font.name = 'Arial'
        p.font.size = Pt(25)
        p.font.color.rgb = RGBColor(255, 255, 255)  # White
        p.font.bold = True
        p.alignment = PP_ALIGN.LEFT
        
        # Add empty line after direction
        p = text_frame.add_paragraph()
        p.text = ""
    
    # Process question content
    lines = question_data['content'].split('\n')
    
    for i, line in enumerate(lines):
        if i == 0:  # First paragraph (no need to add new)
            if question_data['direction']:
                p = text_frame.add_paragraph()
            else:
                p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        
        p.text = line
        p.font.name = 'Arial'
        p.font.size = Pt(25)
        
        # Check if this is a question line (starts with number) or contains "?" 
        if re.match(r'^\d+\.', line) or '?' in line:
            p.font.color.rgb = RGBColor(255, 255, 255)  # White for questions
        else:
            # This is likely an option line
            p.font.color.rgb = RGBColor(255, 255, 0)  # Yellow for options
        
        p.alignment = PP_ALIGN.LEFT
        
        # Add some spacing between lines
        p.space_after = Pt(6)
    return slide

def convert_word_to_ppt(word_path, ppt_path):
    """Main function to convert Word document to PowerPoint"""
    # Parse Word document
    print("Parsing Word document...")
    questions = parse_word_document(word_path)
    
    # Create PowerPoint presentation
    print("Creating PowerPoint presentation...")
    prs = Presentation()
    
    # Set slide size to 16:9
    prs.slide_width = Inches(16)
    prs.slide_height = Inches(9)
    
    # Create slides for each question
    for i, question in enumerate(questions):
        print(f"Creating slide {i+1}...")
        create_slide(prs, question)
    
    # Save presentation
    print(f"Saving presentation to {ppt_path}...")
    prs.save(ppt_path)
    print("Conversion complete!")

# Main execution
if __name__ == "__main__":
    # Input and output file paths
    input_file = "mcq3.docx"
    output_file = "output3.pptx"
    
    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found!")
    else:
        convert_word_to_ppt(input_file, output_file)
        print(f"Successfully created {output_file}")