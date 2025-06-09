# Create these directories first:
# converter/management/
# converter/management/commands/
# Then add this file: converter/management/commands/cleanup_old_conversions.py

from django.core.management.base import BaseCommand
from django.utils import timezone
from django.conf import settings
from datetime import timedelta
from converter.models import ConversionJob


class Command(BaseCommand):
    """Management command to clean up old conversion files"""
    
    help = 'Delete old conversion jobs and their associated files'
    
    def add_arguments(self, parser):
        parser.add_argument(
            '--days',
            type=int,
            default=getattr(settings, 'CONVERSION_FILE_RETENTION_DAYS', 1),
            help='Number of days to retain files (default: 1)'
        )
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Show what would be deleted without actually deleting'
        )
    
    def handle(self, *args, **options):
        days = options['days']
        dry_run = options['dry_run']
        
        # Calculate cutoff date
        cutoff_date = timezone.now() - timedelta(days=days)
        
        # Find old jobs
        old_jobs = ConversionJob.objects.filter(created_at__lt=cutoff_date)
        
        count = old_jobs.count()
        
        if count == 0:
            self.stdout.write(self.style.SUCCESS('No old conversion jobs found.'))
            return
        
        if dry_run:
            self.stdout.write(
                self.style.WARNING(
                    f'DRY RUN: Would delete {count} conversion job(s) older than {days} day(s).'
                )
            )
            for job in old_jobs[:10]:  # Show first 10
                self.stdout.write(f'  - Job {job.id}: {job.created_at}')
            if count > 10:
                self.stdout.write(f'  ... and {count - 10} more')
        else:
            # Delete files and jobs
            deleted_count = 0
            for job in old_jobs:
                try:
                    # Delete associated files
                    job.delete_files()
                    # Delete the job
                    job.delete()
                    deleted_count += 1
                except Exception as e:
                    self.stdout.write(
                        self.style.ERROR(f'Error deleting job {job.id}: {e}')
                    )
            
            self.stdout.write(
                self.style.SUCCESS(
                    f'Successfully deleted {deleted_count} old conversion job(s).'
                )
            )

# Add __init__.py files to:
# converter/management/__init__.py (empty file)
# converter/management/commands/__init__.py (empty file)