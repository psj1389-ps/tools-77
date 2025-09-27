#!/bin/bash

# Render build script for Python 3.12 compatibility with resolvelib fix
set -e

echo "Starting build process..."

# Force upgrade pip to 25.2 to fix resolvelib issues
echo "Upgrading pip to 25.2..."
python -m pip install --upgrade --no-cache-dir pip==25.2

# Clear pip cache to avoid conflicts
echo "Clearing pip cache..."
python -m pip cache purge

# Install essential build tools first with legacy resolver
echo "Installing build tools..."
python -m pip install --use-deprecated=legacy-resolver --no-cache-dir setuptools>=75.1.0 wheel>=0.42.0 build>=1.2.1

# Install setuptools-scm and other build dependencies
echo "Installing build dependencies..."
python -m pip install --use-deprecated=legacy-resolver --no-cache-dir setuptools-scm>=8.0.0 cython>=3.0.0 packaging>=24.0 distlib>=0.3.8

# Install numpy first (critical dependency)
echo "Installing numpy..."
python -m pip install --use-deprecated=legacy-resolver --no-cache-dir numpy==1.26.4

# Install core web framework
echo "Installing web framework..."
python -m pip install --use-deprecated=legacy-resolver --no-cache-dir Flask==2.3.3 Werkzeug==2.3.7 gunicorn==21.2.0 python-dotenv==1.0.0

# Install document processing libraries
echo "Installing document processing libraries..."
python -m pip install --use-deprecated=legacy-resolver --no-cache-dir python-docx==0.8.11 reportlab==4.0.4 "Pillow>=10.2.0"

# Install PDF processing libraries
echo "Installing PDF processing libraries..."
python -m pip install --use-deprecated=legacy-resolver --no-cache-dir pdfservices-sdk==4.2.0 pdf2image==1.16.3 "pdfplumber>=0.10.3" "python-pptx>=0.6.21" "PyPDF2>=3.0.1" "pdf2docx>=0.5.6"

# Install image processing and OCR (potential conflict sources)
echo "Installing image processing and OCR..."
python -m pip install --use-deprecated=legacy-resolver --no-cache-dir opencv-python-headless==4.9.0.80 pytesseract==0.3.10

# Install utilities
echo "Installing utilities..."
python -m pip install --use-deprecated=legacy-resolver --no-cache-dir "qrcode>=7.4.2" "requests>=2.31.0"

echo "Build completed successfully!"