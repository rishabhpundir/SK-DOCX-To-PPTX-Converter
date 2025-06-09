# Word to PPTX Converter - Django Setup Instructions

## Project Structure

```
core/                           # Django project directory
├── core/
│   ├── __init__.py
│   ├── settings.py            # Update with provided settings
│   ├── urls.py                # Replace with provided urls.py
│   ├── asgi.py
│   └── wsgi.py
├── converter/                  # Django app directory
│   ├── __init__.py
│   ├── admin.py               # Replace with provided admin.py
│   ├── apps.py
│   ├── forms.py               # Create and add provided forms.py
│   ├── models.py              # Replace with provided models.py
│   ├── views.py               # Replace with provided views.py
│   ├── urls.py                # Create and add provided urls.py
│   ├── converters.py          # Create and add provided converters.py
│   ├── conversion_scripts/    # Create this directory for conversion scripts
│   │   ├── __init__.py        # Create empty file
│   │   ├── passage_converter.py
│   │   ├── mcq1_converter.py
│   │   ├── mcq2_converter.py
│   │   └── mcq3_converter.py
│   ├── management/
│   │   ├── __init__.py        # Create empty file
│   │   └── commands/
│   │       ├── __init__.py    # Create empty file
│   │       └── cleanup_old_conversions.py
│   ├── templatetags/
│   │   ├── __init__.py        # Create empty file
│   │   └── converter_tags.py  # Add provided template tags
│   └── migrations/
├── templates/                  # Create this directory
│   ├── base.html              # Add provided base.html
│   └── converter/             # Create this directory
│       ├── home.html          # Add provided home.html
│       └── download.html      # Add provided download.html
├── media/                     # Will be created automatically
│   ├── inputs/
│   ├── outputs/
│   └── temp/
├── static/                    # Create if needed
└── manage.py
```

## Step-by-Step Setup

### 1. Install Required Dependencies

```bash
pip install django
pip install python-docx
pip install python-pptx
pip install pillow
pip install opencv-python  # For mcq2_converter.py
pip install pdf2image      # For mcq3_converter.py
pip install numpy          # For mcq3_converter.py
```

### 2. Create the Directory Structure

```bash
# From your project root (where manage.py is located)
mkdir -p converter/conversion_scripts
mkdir -p converter/management/commands
mkdir -p converter/templatetags
mkdir -p templates/converter
touch converter/conversion_scripts/__init__.py
touch converter/management/__init__.py
touch converter/management/commands/__init__.py
touch converter/templatetags/__init__.py
```

### 3. Add the Provided Files

1. Copy all the provided Python files to their respective locations
2. Copy the 4 conversion scripts to `converter/conversion_scripts/`
3. Copy the HTML templates to `templates/` and `templates/converter/`

### 4. Update Settings

Add the provided settings configurations to your `core/settings.py` file.

### 5. Run Migrations

```bash
python manage.py makemigrations
python manage.py migrate
```

### 6. Create Superuser (Optional)

```bash
python manage.py createsuperuser
```

### 7. Collect Static Files (for production)

```bash
python manage.py collectstatic
```

### 8. Run the Development Server

```bash
python manage.py runserver
```

## Usage

1. Navigate to http://localhost:8000/
2. Upload a Word document (.docx)
3. Select the appropriate conversion template:
   - **Passage**: For documents with passages and comprehension questions
   - **MCQ Type 1**: For standard MCQ format
   - **MCQ Type 2**: For MCQs with arrangements and diagrams
   - **MCQ Type 3**: For simple MCQ format
4. Click "Convert to PowerPoint"
5. Download the converted file

## Additional Features

### Admin Interface

Access the admin interface at http://localhost:8000/admin/ to:
- View all conversion jobs
- Monitor conversion status
- Download converted files
- View error messages for failed conversions

### Cleanup Old Files

Run the cleanup command to delete old conversion files:

```bash
# Delete files older than 1 day (default)
python manage.py cleanup_old_conversions

# Delete files older than 7 days
python manage.py cleanup_old_conversions --days 7

# Dry run (show what would be deleted)
python manage.py cleanup_old_conversions --dry-run
```

### Setting Up Periodic Cleanup (Optional)

For Linux/Mac, add to crontab:
```bash
# Run cleanup daily at 2 AM
0 2 * * * cd /path/to/project && python manage.py cleanup_old_conversions
```

For Windows, use Task Scheduler to run the command periodically.

## Troubleshooting

### Common Issues

1. **Import Error for conversion scripts**
   - Ensure all 4 conversion scripts are in `converter/conversion_scripts/`
   - Check that `__init__.py` exists in the directory

2. **File upload fails**
   - Check that media directories have write permissions
   - Ensure file size is under 50MB
   - Verify the file is a valid .docx file

3. **Conversion fails**
   - Check the error message in the admin interface
   - Ensure all required libraries are installed
   - Verify the Word document matches the expected format

4. **Static files not loading**
   - Run `python manage.py collectstatic`
   - Check STATIC_URL and STATIC_ROOT settings

## Production Deployment

For production deployment:

1. Set `DEBUG = False` in settings.py
2. Configure a proper database (PostgreSQL recommended)
3. Use a web server (Nginx) with Gunicorn/uWSGI
4. Set up proper media file serving
5. Configure ALLOWED_HOSTS
6. Use environment variables for sensitive settings
7. Set up SSL/HTTPS
8. Configure proper logging
9. Set up automated backups
10. Monitor disk space for media files

## Security Considerations

1. The app validates file types and sizes
2. Files are stored with UUID names to prevent conflicts
3. Consider adding user authentication for production
4. Implement rate limiting to prevent abuse
5. Regular cleanup of old files is recommended
6. Monitor disk usage and set alerts

## Customization

You can customize:
- File size limits in `forms.py`
- Retention period in settings
- UI colors and styling in templates
- Add more conversion templates
- Add user authentication
- Implement async processing with Celery