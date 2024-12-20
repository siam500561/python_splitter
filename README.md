# Splitter

A modern desktop application for splitting and merging PDF files, and extracting slides from PowerPoint presentations.

## Features

- 🔍 **PDF Page Extraction**: Extract specific pages from PDF files using simple range syntax
- 📑 **PDF Merging**: Combine two PDF files into a single document
- 📊 **PowerPoint Slide Extraction**: Extract specific slides from PPTX files
- 💫 **Modern UI**: Clean and intuitive interface with dark mode support
- 🎯 **Smart Output**: Automatically opens the output location after processing

## Installation

### For Users

1. Download the latest release from the [Releases](https://github.com/siam500561/python_splitter/releases) page
2. Extract the `Splitter.zip` file
3. Run `Splitter.exe` from the extracted folder

**Note**: Keep all files in the Splitter folder together - they are required for the application to work.

### For Developers

```bash
# Clone the repository
git clone https://github.com/siam500561/python_splitter.git

# Navigate to the directory
cd python_splitter

# Create a virtual environment
python -m venv venv

# Activate the virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py
```

## Usage

### PDF Page Extraction

1. Click "Extract PDF" button
2. Select your PDF file
3. Enter the page range (e.g., "1,2" or "1-3,5-7")
4. The output file will be created in the same directory as the input file

### PDF Merging

1. Click "Merge PDFs" button
2. Select your first PDF file
3. Select your second PDF file
4. The merged file will be created in the same directory as the first input file

### PowerPoint Slide Extraction

1. Click "Extract Slides" button
2. Select your PPTX file
3. Enter the slide range (e.g., "1,2" or "1-3,5-7")
4. The output file will be created in the same directory as the input file

## Building from Source

To create an executable:

```bash
# Install PyInstaller
pip install pyinstaller

# Build the executable
python setup.py
```

The executable will be created in the `dist/Splitter` directory.

## Requirements

- Windows 10 or later
- For development:
  - Python 3.8 or later
  - Dependencies listed in requirements.txt

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built with [PyQt6](https://www.riverbankcomputing.com/software/pyqt/)
- PDF processing with [PyPDF2](https://pypdf2.readthedocs.io/)
- PowerPoint handling with [python-pptx](https://python-pptx.readthedocs.io/)

## Support

If you encounter any issues or have suggestions, please [open an issue](https://github.com/siam500561/python_splitter/issues).
