from django.db import models
from django.core.validators import FileExtensionValidator
import os
from uuid import uuid4


def upload_to_input(instance, filename):
    """Generate unique path for input files"""
    fname, ext = filename.rsplit(".", 1)
    uhex = uuid4().hex[:8]
    filename = f"{fname}_{uhex}.{ext}"
    return os.path.join('inputs', filename)


def upload_to_output(instance, filename):
    """Generate unique path for output files"""
    fname, ext = filename.rsplit(".", 1)
    uhex = uuid4().hex[:8]
    filename = f"{fname}_{uhex}.{ext}"
    return os.path.join('outputs', filename)


class ConversionJob(models.Model):
    """Model to track conversion jobs"""
    
    TEMPLATE_CHOICES = [
        ('passage', 'CLAT'),
        ('mcq1', 'Blank - QWH'),
        ('mcq2', 'Blank - LWH'),
        ('mcq3', 'YouTube'),
    ]
    
    STATUS_CHOICES = [
        ('pending', 'Pending'),
        ('processing', 'Processing'),
        ('completed', 'Completed'),
        ('failed', 'Failed'),
    ]
    
    # File fields
    input_file = models.FileField(
        upload_to=upload_to_input,
        validators=[FileExtensionValidator(allowed_extensions=['docx'])],
        help_text="Upload Word document (.docx)"
    )
    output_file = models.FileField(
        upload_to=upload_to_output,
        blank=True,
        null=True,
        help_text="Generated PowerPoint file"
    )
    
    # Metadata fields
    template_type = models.CharField(
        max_length=20,
        choices=TEMPLATE_CHOICES,
        help_text="Type of conversion template"
    )
    status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default='pending'
    )
    
    # Tracking fields
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    processing_time = models.FloatField(
        null=True,
        blank=True,
        help_text="Processing time in seconds"
    )
    error_message = models.TextField(
        blank=True,
        help_text="Error message if conversion failed"
    )
    
    # User tracking (optional - can be used if you add authentication later)
    user_ip = models.GenericIPAddressField(
        null=True,
        blank=True,
        help_text="IP address of the user"
    )
    user_agent = models.CharField(
        max_length=255,
        blank=True,
        help_text="User agent string"
    )
    
    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Conversion Job'
        verbose_name_plural = 'Conversion Jobs'
    
    def __str__(self):
        return f"{self.template_type} - {self.created_at.strftime('%Y-%m-%d %H:%M')} - {self.status}"
    
    def get_input_filename(self):
        """Get the original filename of the input file"""
        return os.path.basename(self.input_file.name)
    
    def get_output_filename(self):
        """Get a user-friendly output filename"""
        if self.output_file:
            input_name = self.get_input_filename()
            base_name = os.path.splitext(input_name)[0]
            return f"{base_name}_converted.pptx"
        return None
    
    def delete_files(self):
        """Delete associated files when job is deleted"""
        if self.input_file:
            self.input_file.delete(save=False)
        if self.output_file:
            self.output_file.delete(save=False)