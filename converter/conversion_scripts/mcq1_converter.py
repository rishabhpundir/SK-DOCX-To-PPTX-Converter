#!/usr/bin/env python3
"""
MCQ Word to PowerPoint Converter
Converts Word documents containing MCQs directly to formatted PowerPoint presentations
"""

import re
import os
import zipfile
import argparse
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor


class MCQConverter:
    def __init__(self):
        # MCQ parsing patterns
        self.mcq_pattern = r'\*\*(\d+)\.\*\*\s*(.*?)(?=\*\*\d+\.\*\*|\Z)'
        self.option_pattern = r'\\?\((\d+)\)\s*([^\\(]+?)(?=\\?\(\d+\)|$)'
        
        # Color definitions
        self.bg_color = RGBColor(0, 0, 0)  # Black background
        self.question_color = RGBColor(255, 255, 255)  # White for questions
        self.option_color = RGBColor(255, 215, 0)  # Gold/Yellow for options
        
        # Slide dimensions (16:9)
        self.slide_width = Inches(13.33)
        self.slide_height = Inches(7.5)
        
        # Layout settings
        self.left_margin = Inches(5.33)  # 40% of slide width
        self.content_width = Inches(7.5)  # 60% of slide width
        self.top_margin = Inches(0.5)
        self.content_height = Inches(6.5)
    
    def extract_images_from_docx(self, docx_path):
        """Extract all images from the Word document"""
        images = {}
        try:
            with zipfile.ZipFile(docx_path, 'r') as doc_zip:
                image_files = [f for f in doc_zip.namelist() if f.startswith('word/media/')]
                for img_file in image_files:
                    img_name = os.path.basename(img_file)
                    img_data = doc_zip.read(img_file)
                    images[img_name] = img_data
        except Exception as e:
            print(f"Error extracting images: {e}")
        return images
    
    def parse_mcq_text(self, text):
        """Parse MCQ text to extract question and options"""
        text = re.sub(r'\s+', ' ', text.strip())
        
        question_match = re.match(r'^(\d+)\.\s*(.*)', text)
        if not question_match:
            return None
            
        question_num = question_match.group(1)
        remaining_text = question_match.group(2)
        
        option_start = re.search(r'\(\d+\)', remaining_text)
        if not option_start:
            return None
            
        question_text = remaining_text[:option_start.start()].strip()
        options_text = remaining_text[option_start.start():].strip()
        
        options = []
        option_matches = re.findall(r'\((\d+)\)\s*([^(]+?)(?=\(\d+\)|$)', options_text)
        
        for opt_num, opt_text in option_matches:
            options.append(f"({opt_num}) {opt_text.strip()}")
            
        return {
            'number': int(question_num),
            'question': question_text,
            'options': options
        }
    
    def extract_mcqs_from_document(self, doc_path):
        """Extract all MCQs from the Word document"""
        doc = Document(doc_path)
        mcqs = []
        current_text = ""
        
        # Extract all text content
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                current_text += paragraph.text + "\n"
        
        # Also extract text from tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        current_text += cell.text + "\n"
        
        # Look for question patterns like **1.**, **2.**, etc.
        question_pattern = r'\*\*(\d+)\.\*\*'
        questions = re.split(question_pattern, current_text)
        
        # Process each question section
        for i in range(1, len(questions), 2):
            if i + 1 < len(questions):
                question_num = questions[i]
                question_content = questions[i + 1]
                
                mcq = self.parse_question_content(question_num, question_content)
                if mcq:
                    mcqs.append(mcq)
        
        # If no questions found with ** pattern, try alternative parsing
        if not mcqs:
            mcqs = self.alternative_parsing(current_text)
            
        return mcqs
    
    def parse_question_content(self, question_num, content):
        """Parse individual question content"""
        lines = content.strip().split('\n')
        question_text = ""
        options = []
        
        # Find where options start
        option_start_idx = -1
        for i, line in enumerate(lines):
            if re.match(r'\\?\(\d+\)', line.strip()):
                option_start_idx = i
                break
        
        # Extract question text
        if option_start_idx > 0:
            question_text = ' '.join(lines[:option_start_idx]).strip()
        elif option_start_idx == 0:
            question_text = "Question text missing"
        else:
            question_text = ' '.join(lines).strip()
        
        # Extract options
        if option_start_idx >= 0:
            for line in lines[option_start_idx:]:
                line = line.strip()
                if re.match(r'\\?\(\d+\)', line):
                    option_text = re.sub(r'\\?\(\d+\)\s*', '', line).strip()
                    option_match = re.match(r'\\?\((\d+)\)', line)
                    if option_match:
                        option_num = option_match.group(1)
                        options.append(f"({option_num}) {option_text}")
        
        if question_text and options:
            return {
                'number': int(question_num),
                'question': question_text,
                'options': options
            }
        
        return None
    
    def alternative_parsing(self, text):
        """Alternative parsing method for different text formats"""
        mcqs = []
        
        # Look for numbered questions followed by options
        pattern = r'(?<=\s)(\d{1,2})\.\s+(.*?)(?=\s\d{1,2}\.\s+|\Z)'
        matches = re.findall(pattern, text, re.DOTALL)
        
        for match in matches:
            question_num = match[0]
            content = match[1].strip()
            
            # Find options in the content
            option_matches = re.findall(r'\((\d+)\)\s*([^(]+?)(?=\(\d+\)|$)', content)
            
            if option_matches:
                # Extract question text (everything before first option)
                first_option_pos = content.find(f"({option_matches[0][0]})")
                question_text = content[:first_option_pos].strip()
                
                options = []
                for opt_num, opt_text in option_matches:
                    options.append(f"({opt_num}) {opt_text.strip()}")
                
                mcqs.append({
                    'number': int(question_num),
                    'question': question_text,
                    'options': options
                })
        
        return mcqs
    
    def set_slide_background(self, slide):
        """Set slide background to black"""
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = self.bg_color
    
    def add_question_directive(self, slide):
        """Add the directive text at the top of the slide"""
        directive_box = slide.shapes.add_textbox(
            self.left_margin,
            Inches(0.1),
            self.content_width,
            Inches(0.8)
        )
        
        text_frame = directive_box.text_frame
        p = text_frame.paragraphs[0]
        p.text = "DIRECTIONS: Select the correct alternative from the given choices."
        p.alignment = PP_ALIGN.LEFT
        p.font.name = 'Arial'
        p.font.size = Pt(16)
        p.font.color.rgb = self.question_color
        p.font.bold = True
    
    def create_formatted_slide(self, prs, mcq, is_first_slide=False):
        """Create a single formatted slide with MCQ content"""
        # Create new slide
        slide_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)
        
        # Set background
        self.set_slide_background(slide)
        
        # Add directive to first slide
        if is_first_slide:
            self.add_question_directive(slide)
            top_margin = Inches(1.0)
        else:
            top_margin = self.top_margin
        
        # Create formatted text box
        text_box = slide.shapes.add_textbox(
            self.left_margin,
            top_margin,
            self.content_width,
            self.slide_height - top_margin - Inches(0.2)
        )
        
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        text_frame.margin_left = Inches(0.1)
        text_frame.margin_right = Inches(0.1)
        text_frame.margin_top = Inches(0.1)
        text_frame.margin_bottom = Inches(0.1)
        
        # Add question text
        p = text_frame.paragraphs[0]
        p.text = f"{mcq['number']}. {mcq['question']}"
        p.alignment = PP_ALIGN.LEFT
        p.font.name = 'Arial'
        p.font.size = Pt(18)
        p.font.color.rgb = self.question_color
        p.font.bold = False
        p.space_after = Pt(12)
        
        # Add options
        for option in mcq['options']:
            p = text_frame.add_paragraph()
            p.text = option
            p.alignment = PP_ALIGN.LEFT
            p.font.name = 'Arial'
            p.font.size = Pt(18)
            p.font.color.rgb = self.option_color
            p.font.bold = False
            p.space_after = Pt(6)
    
    def convert_document(self, input_docx, output_pptx):
        """Main conversion function - converts Word to formatted PowerPoint"""
        print(f"Converting: {input_docx} to {output_pptx}")
        
        # Extract images
        images = self.extract_images_from_docx(input_docx)
        print(f"Found {len(images)} images")
        
        # Extract MCQs
        mcqs = self.extract_mcqs_from_document(input_docx)
        print(f"Found {len(mcqs)} MCQs")
        
        if not mcqs:
            print("No MCQs found in the document!")
            return False
        
        # Create formatted presentation
        prs = Presentation()
        prs.slide_width = self.slide_width
        prs.slide_height = self.slide_height
        
        # Create slides for each MCQ
        for i, mcq in enumerate(mcqs):
            self.create_formatted_slide(prs, mcq, is_first_slide=(i == 0))
        
        # Save presentation
        prs.save(output_pptx)
        print(f"Formatted presentation saved: {output_pptx}")
        
        return True


def convert_word_to_ppt(input_file, output_file):
    """
    Programmatic interface for converting Word to PowerPoint.
    Can be called from Django views or other Python code.
    
    Args:
        input_file (str): Path to input Word document
        output_file (str): Path for output PowerPoint file
        
    Returns:
        bool: True if conversion successful, False otherwise
    """
    converter = MCQConverter()
    return converter.convert_document(input_file, output_file)


def main():
    """Command-line interface"""
    parser = argparse.ArgumentParser(
        description='Convert Word document with MCQs to formatted PowerPoint presentation'
    )
    parser.add_argument(
        'input',
        help='Input Word document (.docx) file path'
    )
    parser.add_argument(
        'output',
        help='Output PowerPoint (.pptx) file path'
    )
    parser.add_argument(
        '--verbose',
        '-v',
        action='store_true',
        help='Enable verbose output'
    )
    
    args = parser.parse_args()
    
    # Validate input file exists
    if not os.path.exists(args.input):
        print(f"Error: Input file not found: {args.input}")
        return 1
    
    # Validate input file extension
    if not args.input.lower().endswith('.docx'):
        print("Error: Input file must be a .docx file")
        return 1
    
    # Ensure output has .pptx extension
    if not args.output.lower().endswith('.pptx'):
        args.output += '.pptx'
    
    # Perform conversion
    try:
        success = convert_word_to_ppt(args.input, args.output)
        if success:
            if args.verbose:
                print("Conversion completed successfully!")
            return 0
        else:
            print("Conversion failed: No MCQs found")
            return 1
    except Exception as e:
        print(f"Error during conversion: {e}")
        return 1


if __name__ == "__main__":
    exit(main())