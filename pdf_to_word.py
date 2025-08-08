import sys
import os
import argparse
import platform
from pathlib import Path

def is_windows():
    return platform.system().lower() == 'windows'

def is_linux():
    return platform.system().lower() == 'linux'

def convert_pdf_to_word(input_path, output_path=None):
    """Convert PDF to Word using pdf2docx"""
    try:
        from pdf2docx import Converter

        input_path = Path(input_path)
        if output_path is None:
            output_path = input_path.with_suffix('.docx')
        else:
            output_path = Path(output_path)

        print(f"Converting '{input_path}' to '{output_path}' using pdf2docx...")

        # Create converter
        cv = Converter(str(input_path))
        cv.convert(str(output_path), start=0, end=None)
        cv.close()

        if output_path.exists():
            print(f"✅ Success! Word file saved as: {output_path}")
            return True
        else:
            print("❌ Error: Conversion failed - output file not created.")
            return False

    except ImportError:
        print("❌ Error: Required library 'pdf2docx' not found.")
        print("Install it using: pip install pdf2docx")
        return False
    except Exception as e:
        print(f"❌ Error during conversion: {str(e)}")
        return False

def main():
    parser = argparse.ArgumentParser(
        description="Cross-platform PDF to Word converter (Windows & Linux)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python pdf_to_word.py input.pdf
  python pdf_to_word.py input.pdf -o output.docx
        """
    )

    parser.add_argument(
        'input_file',
        help='Path to the input PDF file'
    )

    parser.add_argument(
        '-o', '--output',
        dest='output_file',
        help='Path for the output Word file (optional)'
    )

    args = parser.parse_args()

    # Check file existence
    if not os.path.exists(args.input_file):
        print(f"❌ Error: Input file '{args.input_file}' not found.")
        sys.exit(1)

    # Run conversion
    success = convert_pdf_to_word(args.input_file, args.output_file)

    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()
