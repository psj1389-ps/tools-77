#!/bin/bash

# Render build script for Python 3.12 compatibility with modern resolver
set -e

echo "Starting build process..."

# Upgrade pip to latest version with modern resolver
echo "Upgrading pip to 25.2..."
python -m pip install --upgrade --no-cache-dir pip==25.2

# Clear pip cache to avoid conflicts
echo "Clearing pip cache..."
python -m pip cache purge

# Set pip configuration for better dependency resolution
echo "Configuring pip for optimal dependency resolution..."
export PIP_RESOLVER="2020-resolver"
export PIP_USE_FEATURE="2020-resolver"
export PIP_CONFIG_FILE="./pip.conf"

# Install essential build tools first with modern resolver
echo "Installing build tools..."
python -m pip install --no-cache-dir --disable-pip-version-check setuptools>=75.1.0 wheel>=0.42.0 build>=1.2.1

# Install setuptools-scm and other build dependencies
echo "Installing build dependencies..."
python -m pip install --no-cache-dir --disable-pip-version-check setuptools-scm>=8.0.0 cython>=3.0.0 packaging>=24.0 distlib>=0.3.8

# Install numpy first (critical dependency)
echo "Installing numpy..."
python -m pip install --no-cache-dir --disable-pip-version-check numpy==1.26.4

# Install core web framework
echo "Installing web framework..."
python -m pip install --no-cache-dir --disable-pip-version-check Flask==2.3.3 Werkzeug==2.3.7 gunicorn==21.2.0 python-dotenv==1.0.0

# Install document processing libraries
echo "Installing document processing libraries..."
python -m pip install --no-cache-dir --disable-pip-version-check python-docx==0.8.11 reportlab==4.0.4 "Pillow>=10.2.0,<11.0.0"

# Install PDF processing libraries (one by one to avoid conflicts)
echo "Installing PDF processing libraries..."
python -m pip install --no-cache-dir --disable-pip-version-check pdfservices-sdk==4.2.0
python -m pip install --no-cache-dir --disable-pip-version-check pdf2image==1.16.3
python -m pip install --no-cache-dir --disable-pip-version-check "pdfplumber>=0.10.3"
python -m pip install --no-cache-dir --disable-pip-version-check "python-pptx>=0.6.21"
python -m pip install --no-cache-dir --disable-pip-version-check "PyMuPDF>=1.23.0,<1.25.0"

# Install image processing and OCR (potential conflict sources)
echo "Installing image processing and OCR..."
python -m pip install --no-cache-dir --disable-pip-version-check opencv-python-headless==4.9.0.80 pytesseract==0.3.10

# Install utilities
echo "Installing utilities..."
python -m pip install --no-cache-dir --disable-pip-version-check "qrcode>=7.4.2" "requests>=2.31.0"

# Verify installation
echo "Verifying critical packages..."
python -c "import numpy, cv2, PIL, flask; print('All critical packages imported successfully')"

echo "Build completed successfully!"