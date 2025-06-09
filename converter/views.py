from django.shortcuts import render, redirect, get_object_or_404
from django.http import HttpResponse, FileResponse, Http404
from django.contrib import messages
from django.urls import reverse
from django.views import View
from django.views.generic import FormView, DetailView
import os
from .models import ConversionJob
from .forms import ConversionForm
from .converters import converter_manager


class HomeView(FormView):
    """Home page with upload form"""
    template_name = 'converter/home.html'
    form_class = ConversionForm
    
    def get_client_info(self, request):
        """Extract client information from request"""
        x_forwarded_for = request.META.get('HTTP_X_FORWARDED_FOR')
        if x_forwarded_for:
            ip = x_forwarded_for.split(',')[0]
        else:
            ip = request.META.get('REMOTE_ADDR')
        
        user_agent = request.META.get('HTTP_USER_AGENT', '')
        
        return ip, user_agent
    
    def form_valid(self, form):
        """Handle valid form submission"""
        # Get client info
        ip, user_agent = self.get_client_info(self.request)
        
        # Create conversion job
        job = form.save(commit=False)
        job.user_ip = ip
        job.user_agent = user_agent[:255]  # Limit to field max length
        job.save()
        
        # Process the conversion (synchronously for now)
        # In production, you might want to use Celery for async processing
        success = converter_manager.convert(job)
        
        if success:
            messages.success(self.request, 'Conversion completed successfully!')
            return redirect('converter:download', pk=job.pk)
        else:
            messages.error(self.request, f'Conversion failed: {job.error_message}')
            return redirect('converter:home')
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['title'] = 'Word to PPTX Converter'
        
        # Add template descriptions
        context['template_descriptions'] = {
            'passage': 'For documents containing passages and reading comprehension questions',
            'mcq1': 'For standard MCQ format with numbered questions and options',
            'mcq2': 'For MCQs with arrangements, directions, and circular diagrams',
            'mcq3': 'For simple MCQ format with basic question-answer structure',
        }
        
        return context


class DownloadView(DetailView):
    """Download page for completed conversions"""
    model = ConversionJob
    template_name = 'converter/download.html'
    context_object_name = 'job'
    
    def get_object(self):
        """Get the conversion job and verify it's completed"""
        job = super().get_object()
        if job.status != 'completed':
            raise Http404("Conversion not completed")
        return job
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['title'] = 'Download Your Converted File'
        return context


class DownloadFileView(View):
    """Serve the converted file for download"""
    
    def get(self, request, pk):
        job = get_object_or_404(ConversionJob, pk=pk, status='completed')
        
        if not job.output_file:
            raise Http404("Output file not found")
        
        # Get the file path
        file_path = job.output_file.path
        
        if not os.path.exists(file_path):
            raise Http404("File not found on server")
        
        # Create response
        response = FileResponse(
            open(file_path, 'rb'),
            content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
        # Set filename for download
        filename = job.get_output_filename()
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        
        return response


class StatusView(DetailView):
    """Check conversion status (useful for AJAX polling if needed)"""
    model = ConversionJob
    
    def get(self, request, pk):
        job = self.get_object()
        
        # Return JSON response
        import json
        data = {
            'status': job.status,
            'error_message': job.error_message,
            'processing_time': job.processing_time,
        }
        
        if job.status == 'completed':
            data['download_url'] = reverse('converter:download', kwargs={'pk': job.pk})
        
        return HttpResponse(
            json.dumps(data),
            content_type='application/json'
        )