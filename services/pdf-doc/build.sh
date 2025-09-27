#!/bin/bash

# Render build script for Python 3.13 compatibility
set -e

echo "Starting build process..."

# Update pip first
python -m pip install --upgrade pip>=24.2

# Install essential build tools first
echo "Installing build tools..."
python -m pip install --no-cache-dir setuptools>=75.1.0 wheel>=0.42.0 build>=1.2.1

# Install setuptools-scm and other build dependencies
echo "Installing build dependencies..."
python -m pip install --no-cache-dir setuptools-scm>=8.0.0 cython>=3.0.0 packaging>=24.0

# Install requirements
echo "Installing application dependencies..."
python -m pip install --no-cache-dir -r requirements.txt

echo "Build completed successfully!"