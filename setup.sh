#!/bin/bash
# setup.sh - Quick setup script for Word to PPTX Converter
# Run with: bash setup.sh

echo "🚀 Setting up Word to PPTX Converter..."

# Create required directories
echo "📁 Creating directories..."
mkdir -p converter/conversion_scripts
mkdir -p converter/templatetags
mkdir -p converter/management/commands
mkdir -p templates/converter
mkdir -p media/inputs
mkdir -p media/outputs
mkdir -p media/temp
mkdir -p static

# Create __init__.py files
echo "📄 Creating __init__.py files..."
touch converter/conversion_scripts/__init__.py
touch converter/templatetags/__init__.py
touch converter/management/__init__.py
touch converter/management/commands/__init__.py

# Install dependencies
echo "📦 Installing dependencies..."
pip install -r requirements.txt

# Check if conversion scripts exist
echo "🔍 Checking for conversion scripts..."
if [ ! -f "converter/conversion_scripts/passage_converter.py" ]; then
    echo "⚠️  Warning: Conversion scripts not found in converter/conversion_scripts/"
    echo "   Please copy your 4 conversion scripts to this directory:"
    echo "   - passage_converter.py"
    echo "   - mcq1_converter.py"
    echo "   - mcq2_converter.py"
    echo "   - mcq3_converter.py"
fi

# Run migrations
echo "🔧 Running migrations..."
python manage.py makemigrations
python manage.py migrate

# Create superuser prompt
echo ""
read -p "👤 Would you like to create a superuser account? (y/n) " -n 1 -r
echo ""
if [[ $REPLY =~ ^[Yy]$ ]]; then
    python manage.py createsuperuser
fi

# Success message
echo ""
echo "✅ Setup complete!"
echo ""
echo "📌 Next steps:"
echo "1. Make sure your 4 conversion scripts are in converter/conversion_scripts/"
echo "2. Run the server with: python manage.py runserver"
echo "3. Open your browser to: http://localhost:8000/"
echo ""
echo "Happy converting! 🎉"

# For Windows users, create setup.bat:
# @echo off
# echo Setting up Word to PPTX Converter...
# 
# echo Creating directories...
# mkdir converter\conversion_scripts 2>nul
# mkdir converter\templatetags 2>nul
# mkdir converter\management\commands 2>nul
# mkdir templates\converter 2>nul
# mkdir media\inputs 2>nul
# mkdir media\outputs 2>nul
# mkdir media\temp 2>nul
# mkdir static 2>nul
# 
# echo Creating __init__.py files...
# type nul > converter\conversion_scripts\__init__.py
# type nul > converter\templatetags\__init__.py
# type nul > converter\management\__init__.py
# type nul > converter\management\commands\__init__.py
# 
# echo Installing dependencies...
# pip install -r requirements.txt
# 
# echo Running migrations...
# python manage.py makemigrations
# python manage.py migrate
# 
# echo Setup complete!
# echo.
# echo Next steps:
# echo 1. Copy your 4 conversion scripts to converter\conversion_scripts\
# echo 2. Run: python manage.py runserver
# echo 3. Open: http://localhost:8000/
# pause