from django.apps import AppConfig


class ConverterConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'converter'
    verbose_name = 'Word to PPTX Converter'
    
    def ready(self):
        """Initialize app when Django starts"""
        # You can add signal handlers or other initialization code here
        pass