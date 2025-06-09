# Quick Start Guide - Word to PPTX Converter

## 🚀 Quick Setup (5 minutes)

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Create Required Directories
```bash
# Run from project root (where manage.py is)
mkdir -p converter/conversion_scripts converter/templatetags converter/management/commands templates/converter
touch converter/conversion_scripts/__init__.py converter/templatetags/__init__.py converter/management/__init__.py converter/management/commands/__init__.py
```

### 3. Copy Your Conversion Scripts
Place these files in `converter/conversion_scripts/`:
- `passage_converter.py`
- `mcq1_converter.py`
- `mcq2_converter.py`
- `mcq3_converter.py`

### 4. Run Migrations
```bash
python manage.py makemigrations
python manage.py migrate
```

### 5. Run the Server
```bash
python manage.py runserver
```

### 6. Open Your Browser
Navigate to: http://localhost:8000/

## 📝 File Placement Guide

```
Your Django Project/
├── manage.py
├── requirements.txt (create from provided)
├── core/
│   ├── settings.py (update with provided settings)
│   └── urls.py (replace with provided)
├── converter/
│   ├── __init__.py (add provided)
│   ├── admin.py (replace with provided)
│   ├── apps.py (add provided)
│   ├── forms.py (create new, add provided)
│   ├── models.py (replace with provided)
│   ├── views.py (replace with provided)
│   ├── urls.py (create new, add provided)
│   ├── converters.py (create new, add provided)
│   ├── tests.py (replace with provided)
│   ├── conversion_scripts/
│   │   ├── __init__.py
│   │   └── [your 4 conversion scripts here]
│   ├── templatetags/
│   │   ├── __init__.py
│   │   └── converter_tags.py (add provided)
│   └── management/commands/
│       └── cleanup_old_conversions.py (add provided)
└── templates/
    ├── base.html (add provided)
    ├── error.html (add provided)
    └── converter/
        ├── home.html (add provided)
        └── download.html (add provided)
```

## ⚡ Usage

1. **Upload**: Select your Word document
2. **Choose Template**: Pick the appropriate conversion type
3. **Convert**: Click the convert button
4. **Download**: Get your PowerPoint file

## 🛠️ Troubleshooting

### "Module not found" Error
- Ensure all 4 conversion scripts are in `converter/conversion_scripts/`
- Check that all `__init__.py` files exist

### "Template does not exist" Error
- Make sure templates are in the correct directories
- Update `TEMPLATES` setting in `settings.py`

### Conversion Fails
- Check file is a valid .docx
- Ensure file size is under 50MB
- Check admin panel for error details

## 🎯 Next Steps

1. **Create Admin User** (optional):
   ```bash
   python manage.py createsuperuser
   ```

2. **Set Up Cleanup** (optional):
   ```bash
   python manage.py cleanup_old_conversions --dry-run
   ```

3. **Customize**:
   - Edit templates for different styling
   - Modify file size limits in `forms.py`
   - Add more conversion templates

## 💡 Tips

- The admin interface is at `/admin/`
- Old files are automatically cleaned up after 1 day
- Check the console for conversion progress
- All uploads are validated for security

Need more help? Check the full `SETUP_INSTRUCTIONS.md` for detailed information.