from django.test import TestCase, Client
from django.urls import reverse
from django.core.files.uploadedfile import SimpleUploadedFile
from .models import ConversionJob
import os


class ConverterViewsTestCase(TestCase):
    """Basic tests for converter views"""
    
    def setUp(self):
        self.client = Client()
        
    def test_home_page_loads(self):
        """Test that home page loads successfully"""
        response = self.client.get(reverse('converter:home'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Word to PPTX Converter')
        
    def test_form_validation(self):
        """Test form validation for file upload"""
        # Test with no file
        response = self.client.post(reverse('converter:home'), {
            'template_type': 'mcq1'
        })
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'This field is required')
        
    def test_invalid_file_type(self):
        """Test upload of invalid file type"""
        # Create a text file
        test_file = SimpleUploadedFile(
            "test.txt", 
            b"This is a test file",
            content_type="text/plain"
        )
        
        response = self.client.post(reverse('converter:home'), {
            'input_file': test_file,
            'template_type': 'mcq1'
        })
        
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Only .docx files are allowed')
        
    def test_model_creation(self):
        """Test ConversionJob model creation"""
        # Create a mock docx file
        test_file = SimpleUploadedFile(
            "test.docx",
            b"Mock docx content",
            content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        job = ConversionJob.objects.create(
            input_file=test_file,
            template_type='mcq1',
            status='pending'
        )
        
        self.assertEqual(job.status, 'pending')
        self.assertEqual(job.template_type, 'mcq1')
        self.assertTrue(job.input_file)
        
        # Clean up
        job.delete_files()
        job.delete()


class ConverterUtilsTestCase(TestCase):
    """Tests for converter utilities"""
    
    def test_converter_manager_import(self):
        """Test that converter manager can be imported"""
        try:
            from .converters import converter_manager
            self.assertIsNotNone(converter_manager)
        except ImportError:
            self.fail("Could not import converter_manager")
            
    def test_template_types(self):
        """Test that all template types are configured"""
        from .converters import converter_manager
        
        expected_types = ['passage', 'mcq1', 'mcq2', 'mcq3']
        for template_type in expected_types:
            self.assertIn(template_type, converter_manager.converters)