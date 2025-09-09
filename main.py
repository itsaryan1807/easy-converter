# main.py (skeleton)
import argparse
import os
from word_to_pdf import convert_word_to_pdf
from pdf_to_word import convert_pdf_to_word

def parse_args():
    p = argparse.ArgumentParser(description="Easy Converter: word<->pdf")
    p.add_argument('--mode', choices=['word2pdf','pdf2word'], required=True)
    p.add_argument('--input', help='path to input file')
    p.add_argument('--output', help='path to output file (optional)')
    p.add_argument('--input-folder', help='path to folder for batch processing')
    p.add_argument('--output-folder', help='output folder for batch processing')
    p.add_argument('--batch', action='store_true', help='process all files in folder')
    return p.parse_args()

def process_single(mode, inp, out):
    if mode == 'word2pdf':
        convert_word_to_pdf(inp, out)
    else:
        convert_pdf_to_word(inp, out)

def process_batch(mode, input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    for fname in os.listdir(input_folder):
        src = os.path.join(input_folder, fname)
        if mode == 'word2pdf' and fname.lower().endswith(('.docx','.doc')):
            out = os.path.join(output_folder, os.path.splitext(fname)[0] + '.pdf')
            convert_word_to_pdf(src, out)
        if mode == 'pdf2word' and fname.lower().endswith('.pdf'):
            out = os.path.join(output_folder, os.path.splitext(fname)[0] + '.docx')
            convert_pdf_to_word(src, out)

if __name__ == '__main__':
    args = parse_args()
    if args.batch:
        process_batch(args.mode, args.input_folder, args.output_folder)
    else:
        process_single(args.mode, args.input, args.output)
