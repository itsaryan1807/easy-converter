import sys
import os
import argparse
import subprocess
import platform
from pathlib import Path

def is_windows():
    """Check if running on Windows"""
    return platform.system().lower() == 'windows'

def is_linux():
    """Check if running on Linux"""
    return platform.system().lower() == 'linux'

# Windows Method: Using docx2pdf (requires Microsoft Word)
def convert_with_docx2pdf(input_path, output_path=None):
    """Convert using docx2pdf on Windows"""
    try:
        from docx2pdf import convert
        
        # Generate output path if not provided
        if output_path is None:
            input_file = Path(input_path)
            output_path = input_file.with_suffix('.pdf')
        
        # Ensure output directory exists
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"Converting '{input_path}' to '{output_path}' using docx2pdf (Windows)...")
        
        # Perform conversion
        convert(input_path, output_path)
        
        if os.path.exists(output_path):
            print(f"Success! PDF saved as: {output_path}")
            return True
        else:
            print("Error: Conversion failed - output file not created.")
            return False
            
    except ImportError:
        print("Error: docx2pdf library not found. Please install it using:")
        print("pip install docx2pdf")
        print("Note: docx2pdf requires Microsoft Word to be installed on Windows.")
        return False
    except Exception as e:
        print(f"Error during conversion: {str(e)}")
        return False

# Linux Method 1: Using LibreOffice (recommended for Linux)
def convert_with_libreoffice(input_path, output_path=None):
    """Convert using LibreOffice headless mode on Linux"""
    try:
        # Check if LibreOffice is installed
        result = subprocess.run(['libreoffice', '--version'], 
                              capture_output=True, text=True)
        if result.returncode != 0:
            print("Error: LibreOffice is not installed. Install it with:")
            print("sudo apt-get install libreoffice")
            return False
        
        # Generate output path if not provided
        if output_path is None:
            input_file = Path(input_path)
            output_path = input_file.with_suffix('.pdf')
        
        output_dir = Path(output_path).parent
        
        print(f"Converting '{input_path}' to '{output_path}' using LibreOffice (Linux)...")
        
        # Convert using LibreOffice headless
        cmd = [
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', str(output_dir),
            input_path
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            # LibreOffice creates output with same name as input but .pdf extension
            input_name = Path(input_path).stem
            generated_pdf = output_dir / f"{input_name}.pdf"
            
            # Rename to desired output name if different
            if generated_pdf != Path(output_path):
                generated_pdf.rename(output_path)
            
            print(f"Success! PDF saved as: {output_path}")
            return True
        else:
            print(f"Error: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"Error during LibreOffice conversion: {str(e)}")
        return False

# Linux Method 2: Using python-docx + reportlab (fallback for Linux)
def convert_with_python_libs(input_path, output_path=None):
    """Convert using python-docx and reportlab (basic conversion for Linux)"""
    try:
        from docx import Document
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet
        
        # Generate output path if not provided
        if output_path is None:
            input_file = Path(input_path)
            output_path = input_file.with_suffix('.pdf')
        
        print(f"Converting '{input_path}' to '{output_path}' using Python libraries (Linux fallback)...")
        
        # Read DOCX
        doc = Document(input_path)
        
        # Create PDF
        pdf_doc = SimpleDocTemplate(str(output_path), pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        # Extract text from DOCX and add to PDF
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                p = Paragraph(paragraph.text, styles['Normal'])
                story.append(p)
                story.append(Spacer(1, 12))
        
        pdf_doc.build(story)
        
        print(f"Success! PDF saved as: {output_path}")
        print("Note: This method provides basic conversion. For better formatting, install LibreOffice.")
        return True
        
    except ImportError as e:
        missing_lib = str(e).split("'")[1] if "'" in str(e) else "required library"
        print(f"Error: Required library '{missing_lib}' not found.")
        print("Install required libraries with:")
        print("pip install python-docx reportlab")
        return False
    except Exception as e:
        print(f"Error during Python library conversion: {str(e)}")
        return False

def convert_word_to_pdf(input_path, output_path=None, method='auto'):
    """Main conversion function with platform detection"""
    
    # Validate input file
    if not os.path.exists(input_path):
        print(f"Error: Input file '{input_path}' not found.")
        return False
    
    if not input_path.lower().endswith('.docx'):
        print("Error: Input file must be a .docx file.")
        return False
    
    # Ensure output directory exists
    if output_path:
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)
    
    # Detect platform and choose method
    current_os = platform.system().lower()
    print(f"Detected OS: {current_os}")
    
    if method == 'auto':
        if is_windows():
            print("Using Windows method (docx2pdf)...")
            return convert_with_docx2pdf(input_path, output_path)
        elif is_linux():
            print("Using Linux method (LibreOffice)...")
            success = convert_with_libreoffice(input_path, output_path)
            if not success:
                print("LibreOffice failed, trying Python libraries fallback...")
                return convert_with_python_libs(input_path, output_path)
            return success
        else:
            print(f"Error: Unsupported operating system '{current_os}'")
            print("This script supports only Windows and Linux.")
            return False
    
    # Manual method selection
    elif method == 'windows':
        if not is_windows():
            print("Warning: Using Windows method on non-Windows system may not work.")
        return convert_with_docx2pdf(input_path, output_path)
    elif method == 'libreoffice':
        return convert_with_libreoffice(input_path, output_path)
    elif method == 'python':
        return convert_with_python_libs(input_path, output_path)
    else:
        print(f"Error: Unknown method '{method}'")
        return False

def main():
    parser = argparse.ArgumentParser(
        description="Cross-platform Word to PDF converter (Windows & Linux)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s document.docx
  %(prog)s document.docx -o output.pdf
  %(prog)s document.docx --method windows
  %(prog)s document.docx --method libreoffice -o output.pdf

Methods:
  auto         - Auto-detect OS and use appropriate method (default)
  windows      - Use docx2pdf (requires Microsoft Word on Windows)
  libreoffice  - Use LibreOffice (Linux, requires LibreOffice installed)
  python       - Use python-docx + reportlab (Linux fallback, basic conversion)

Platform Support:
  Windows: Uses docx2pdf with Microsoft Word
  Linux:   Uses LibreOffice (primary) or Python libraries (fallback)
        """
    )
    
    parser.add_argument(
        'input_file',
        help='Path to the input Word document (.docx)'
    )
    
    parser.add_argument(
        '-o', '--output',
        dest='output_file',
        help='Path for the output PDF file (optional)'
    )
    
    parser.add_argument(
        '--method',
        choices=['auto', 'windows', 'libreoffice', 'python'],
        default='auto',
        help='Conversion method to use (default: auto)'
    )
    
    args = parser.parse_args()
    
    # Convert the document
    success = convert_word_to_pdf(args.input_file, args.output_file, args.method)
    
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()