from django import forms
from .models import ConversionJob


class ConversionForm(forms.ModelForm):
    """Form for uploading Word documents and selecting conversion template"""
    
    class Meta:
        model = ConversionJob
        fields = ['input_file', 'template_type']
        widgets = {
            'input_file': forms.FileInput(attrs={
                'class': 'form-control',
                'accept': '.docx',
                'required': True,
            }),
            'template_type': forms.Select(attrs={
                'class': 'form-select',
                'required': True,
            }),
        }
        labels = {
            'input_file': 'Select Word Document',
            'template_type': 'Select Conversion Template',
        }
        help_texts = {
            'input_file': 'Only .docx files are supported',
            'template_type': 'Choose the appropriate template based on your document content',
        }
    
    def clean_input_file(self):
        """Additional validation for input file"""
        file = self.cleaned_data.get('input_file')
        if file:
            # Check file size (limit to 50MB)
            if file.size > 50 * 1024 * 1024:
                raise forms.ValidationError('File size must be under 50MB.')
            
            # Verify it's actually a DOCX file
            if not file.name.lower().endswith('.docx'):
                raise forms.ValidationError('Only .docx files are allowed.')
        
        return file