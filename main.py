# main.py
import argparse
import os
import sys
from word_to_pdf import convert_word_to_pdf
from pdf_to_word import convert_pdf_to_word

def parse_args():
    parser = argparse.ArgumentParser(description="Easy Converter: Convert Word <-> PDF files easily.")
    parser.add_argument('--mode', choices=['word2pdf', 'pdf2word'], required=True, help="Choose the conversion mode.")
    parser.add_argument('--input', help='Path to the input file.')
    parser.add_argument('--output', help='Path to the output file (optional).')
    parser.add_argument('--input-folder', help='Path to the folder containing files for batch processing.')
    parser.add_argument('--output-folder', help='Path to the folder for saving converted files.')
    parser.add_argument('--batch', action='store_true', help='Enable batch processing mode.')
    return parser.parse_args()

def process_single(mode, inp, out):
    if not os.path.isfile(inp):
        print(f"Error: Input file '{inp}' does not exist.")
        sys.exit(1)

    if not out:
        base_name = os.path.splitext(os.path.basename(inp))[0]
        if mode == 'word2pdf':
            out = base_name + '.pdf'
        else:
            out = base_name + '.docx'
        print(f"No output file specified. Using default output: {out}")

    print(f"Processing file: {inp}")
    if mode == 'word2pdf':
        convert_word_to_pdf(inp, out)
    else:
        convert_pdf_to_word(inp, out)
    print(f"File converted successfully! Output saved to: {out}")

def process_batch(mode, input_folder, output_folder):
    if not os.path.isdir(input_folder):
        print(f"Error: Input folder '{input_folder}' does not exist.")
        sys.exit(1)

    os.makedirs(output_folder, exist_ok=True)
    print(f"Processing all files in {input_folder}...")

    for fname in os.listdir(input_folder):
        src = os.path.join(input_folder, fname)

        if mode == 'word2pdf' and fname.lower().endswith(('.docx', '.doc')):
            out_name = os.path.splitext(fname)[0] + '.pdf'
        elif mode == 'pdf2word' and fname.lower().endswith('.pdf'):
            out_name = os.path.splitext(fname)[0] + '.docx'
        else:
            continue  # Skip unsupported files

        out = os.path.join(output_folder, out_name)
        print(f"Converting {src} -> {out}")

        if mode == 'word2pdf':
            convert_word_to_pdf(src, out)
        else:
            convert_pdf_to_word(src, out)

    print("Batch processing completed.")

if __name__ == '__main__':
    args = parse_args()

    if args.batch:
        if not args.input_folder or not args.output_folder:
            print("Error: --input-folder and --output-folder are required in batch mode.")
            sys.exit(1)
        process_batch(args.mode, args.input_folder, args.output_folder)
    else:
        if not args.input:
            print("Error: --input is required in single file mode.")
            sys.exit(1)
        process_single(args.mode, args.input, args.output)
