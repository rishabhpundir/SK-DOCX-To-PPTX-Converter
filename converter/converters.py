"""
Converter utilities that integrate the 4 conversion scripts
Place your conversion scripts in converter/conversion_scripts/ directory
"""

import os
import sys
import time
import shutil
import traceback
from django.conf import settings
from django.core.files.base import ContentFile

# Add conversion scripts directory to Python path
SCRIPTS_DIR = os.path.join(os.path.dirname(__file__), 'conversion_scripts')
sys.path.append(SCRIPTS_DIR)


class ConversionError(Exception):
    """Custom exception for conversion errors"""
    pass


class ConverterManager:
    """Manages the conversion process for different template types"""
    
    def __init__(self):
        self.converters = {
            'passage': self.convert_passage,
            'mcq1': self.convert_mcq1,
            'mcq2': self.convert_mcq2,
            'mcq3': self.convert_mcq3,
        }
    
    def convert(self, job):
        """Main conversion method that processes the job"""
        start_time = time.time()
        
        try:
            # Update status to processing
            job.status = 'processing'
            job.save()
            
            # Get the converter function based on template type
            converter_func = self.converters.get(job.template_type)
            if not converter_func:
                raise ConversionError(f"Unknown template type: {job.template_type}")
            
            # Perform conversion
            output_path = converter_func(job)
            
            # Save the output file
            with open(output_path, 'rb') as f:
                output_filename = os.path.basename(output_path)
                job.output_file.save(output_filename, ContentFile(f.read()))
            
            # Clean up temporary file
            if os.path.exists(output_path):
                os.remove(output_path)
            
            # Update job status
            job.status = 'completed'
            job.processing_time = time.time() - start_time
            job.save()
            
            return True
            
        except Exception as e:
            # Handle errors
            print(f"Conversion error: \n{traceback.format_exc()}")
            job.status = 'failed'
            job.error_message = str(e)
            job.processing_time = time.time() - start_time
            job.save()
            
            return False
        finally:
            # Ensure any temporary files are cleaned up
            temp_dir = os.path.join(settings.MEDIA_ROOT, 'extraction')
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
    
    def convert_passage(self, job):
        """Convert using passage converter"""
        try:
            from converter.conversion_scripts.passage_converter import WordToPowerPointConverter
            
            # Get input file path
            input_path = job.input_file.path
            
            # Generate output path
            output_filename = f"passage_{job.id}.pptx"
            output_path = os.path.join(settings.MEDIA_ROOT, 'temp', output_filename)
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Create converter and convert
            converter = WordToPowerPointConverter()
            success = converter.convert(input_path, output_path)
            
            if not success:
                raise ConversionError("Passage conversion failed")
            
            return output_path
            
        except ImportError:
            raise ConversionError("Passage converter script not found. Please ensure passage_converter.py is in converter/conversion_scripts/")
    
    def convert_mcq1(self, job):
        """Convert using MCQ Type 1 converter"""
        try:
            from converter.conversion_scripts.mcq1_converter import convert_word_to_ppt
            
            # Get input file path
            input_path = job.input_file.path
            
            # Generate output path
            output_filename = f"mcq1_{job.id}.pptx"
            output_path = os.path.join(settings.MEDIA_ROOT, 'temp', output_filename)
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Convert
            success = convert_word_to_ppt(input_path, output_path)
            
            if not success:
                raise ConversionError("MCQ Type 1 conversion failed")
            
            return output_path
            
        except ImportError:
            raise ConversionError("MCQ Type 1 converter script not found. Please ensure mcq1_converter.py is in converter/conversion_scripts/")
    
    def convert_mcq2(self, job):
        """Convert using MCQ Type 2 converter"""
        try:
            from converter.conversion_scripts.mcq2_converter import convert_word_to_ppt
            
            # Get input file path
            input_path = job.input_file.path
            
            # Generate output path
            output_filename = f"mcq2_{job.id}.pptx"
            output_path = os.path.join(settings.MEDIA_ROOT, 'temp', output_filename)
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Convert
            convert_word_to_ppt(input_path, output_path)
            
            # Check if file was created
            if not os.path.exists(output_path):
                raise ConversionError("MCQ Type 2 conversion failed - output file not created")
            
            return output_path
            
        except ImportError:
            raise ConversionError("MCQ Type 2 converter script not found. Please ensure mcq2_converter.py is in converter/conversion_scripts/")
        except Exception as e:
            raise ConversionError(f"MCQ Type 2 conversion error: {str(e)}")
    
    def convert_mcq3(self, job):
        """Convert using MCQ Type 3 converter"""
        try:
            from converter.conversion_scripts.mcq3_converter import convert_word_to_ppt
            
            # Get input file path
            input_path = job.input_file.path
            
            # Generate output path
            output_filename = f"mcq3_{job.id}.pptx"
            output_path = os.path.join(settings.MEDIA_ROOT, 'temp', output_filename)
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Convert
            convert_word_to_ppt(input_path, output_path)
            
            # Check if file was created
            if not os.path.exists(output_path):
                raise ConversionError("MCQ Type 3 conversion failed - output file not created")
            
            return output_path
            
        except ImportError:
            raise ConversionError("MCQ Type 3 converter script not found. Please ensure mcq3_converter.py is in converter/conversion_scripts/")
        except Exception as e:
            raise ConversionError(f"MCQ Type 3 conversion error: {str(e)}")


# Create a singleton instance
converter_manager = ConverterManager()