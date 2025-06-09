from django.contrib import admin
from django.utils.html import format_html
from .models import ConversionJob


@admin.register(ConversionJob)
class ConversionJobAdmin(admin.ModelAdmin):
    """Admin interface for ConversionJob model"""
    
    list_display = [
        'id', 
        'get_input_filename',
        'template_type', 
        'status_badge', 
        'processing_time_display',
        'created_at',
        'download_link'
    ]
    
    list_filter = [
        'status',
        'template_type',
        'created_at',
    ]
    
    search_fields = [
        'user_ip',
        'error_message',
    ]
    
    readonly_fields = [
        'input_file',
        'output_file',
        'template_type',
        'status',
        'created_at',
        'updated_at',
        'processing_time',
        'error_message',
        'user_ip',
        'user_agent',
        'file_preview',
    ]
    
    ordering = ['-created_at']
    
    date_hierarchy = 'created_at'
    
    def status_badge(self, obj):
        """Display status as a colored badge"""
        colors = {
            'pending': 'warning',
            'processing': 'info',
            'completed': 'success',
            'failed': 'danger',
        }
        color = colors.get(obj.status, 'secondary')
        return format_html(
            '<span class="badge bg-{}">{}</span>',
            color,
            obj.get_status_display()
        )
    status_badge.short_description = 'Status'
    
    def processing_time_display(self, obj):
        """Display processing time in readable format"""
        if obj.processing_time:
            return f"{obj.processing_time:.2f}s"
        return "-"
    processing_time_display.short_description = 'Processing Time'
    
    def download_link(self, obj):
        """Provide download link for completed conversions"""
        if obj.status == 'completed' and obj.output_file:
            return format_html(
                '<a href="{}" class="btn btn-sm btn-primary">Download</a>',
                obj.output_file.url
            )
        return "-"
    download_link.short_description = 'Download'
    
    def file_preview(self, obj):
        """Show file information"""
        html = f"<strong>Input:</strong> {obj.get_input_filename()}<br>"
        if obj.output_file:
            html += f"<strong>Output:</strong> {obj.get_output_filename()}"
        return format_html(html)
    file_preview.short_description = 'Files'
    
    def has_add_permission(self, request):
        """Disable manual creation of conversion jobs"""
        return False
    
    def has_change_permission(self, request, obj=None):
        """Disable editing of conversion jobs"""
        return False
    
    class Media:
        css = {
            'all': ('https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css',)
        }