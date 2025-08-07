

import sys
import os
import argparse
from pathlib import Path

try:
    from docx2pdf import convert
except ImportError:
    print("Error: docx2pdf library not found. Please install it using:")
    print("pip install docx2pdf")
    sys.exit(1)


def convert_word_to_pdf(input_path, output_path=None):
    
    try:
        # Validate input file
        if not os.path.exists(input_path):
            print(f"Error: Input file '{input_path}' not found.")
            return False
        
        if not input_path.lower().endswith('.docx'):
            print("Error: Input file must be a .docx file.")
            return False
        
        # Generate output path if not provided
        if output_path is None:
            input_file = Path(input_path)
            output_path = input_file.with_suffix('.pdf')
        
        # Ensure output directory exists
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"Converting '{input_path}' to '{output_path}'...")
        
        # Perform conversion
        convert(input_path, output_path)
        
        if os.path.exists(output_path):
            print(f"Success! PDF saved as: {output_path}")
            return True
        else:
            print("Error: Conversion failed - output file not created.")
            return False
            
    except Exception as e:
        print(f"Error during conversion: {str(e)}")
        return False


def main():
    
    parser = argparse.ArgumentParser(
        description="Convert Word documents (.docx) to PDF format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s document.docx
  %(prog)s document.docx -o output.pdf
  %(prog)s path/to/document.docx -o path/to/output.pdf
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
    
    args = parser.parse_args()
    
    # Convert the document
    success = convert_word_to_pdf(args.input_file, args.output_file)
    
    if not success:
        sys.exit(1)


if __name__ == "__main__":
    main()