"""
PDF2Docx Converter Setup
"""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="pdf2docx-converter",
    version="1.0.0",
    author="Chen1Mmm",
    description="A GUI tool to convert PDF files to Word documents with OCR support",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/Chen1Mmm/pdf2docx-converter",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
    install_requires=[
        "pdf2docx>=0.5.1",
        "python-docx>=0.8.11",
        "pytesseract>=0.3.10",
        "pdf2image>=1.16.3",
        "PyQt6>=6.7.0",
        "Pillow>=10.0.0",
    ],
    entry_points={
        "console_scripts": [
            "pdf2docx-converter=pdf_converter:main",
        ],
    },
)