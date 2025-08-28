#!/usr/bin/env python3

import os
import sys
from pathlib import Path
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import inch
import argparse


class ImageToPdfConverter:
    def __init__(self, page_size='A4', margin=0.5):
        self.page_size = self._get_page_size(page_size)
        self.margin = margin * inch
        
    def _get_page_size(self, page_size):
        if page_size.upper() == 'A4':
            return A4
        elif page_size.upper() == 'LETTER':
            return letter
        else:
            try:
                width, height = map(float, page_size.split('x'))
                return (width * inch, height * inch)
            except:
                print(f"Invalid page size: {page_size}. Using A4.")
                return A4
    
    def convert_single_image(self, image_path, output_path=None):
        try:
            with Image.open(image_path) as img:
                if img.mode in ('RGBA', 'LA', 'P'):
                    img = img.convert('RGB')
                
                img_width, img_height = img.size
                page_width, page_height = self.page_size
                
                scale_x = (page_width - 2 * self.margin) / img_width
                scale_y = (page_height - 2 * self.margin) / img_height
                scale = min(scale_x, scale_y)
                
                final_width = img_width * scale
                final_height = img_height * scale
                x = (page_width - final_width) / 2
                y = (page_height - final_height) / 2
                
                if output_path is None:
                    output_path = str(Path(image_path).with_suffix('.pdf'))
                
                c = canvas.Canvas(output_path, pagesize=self.page_size)
                c.drawImage(image_path, x, y, final_width, final_height)
                c.save()
                
                print(f"✓ Converted: {image_path} → {output_path}")
                return output_path
                
        except Exception as e:
            print(f"✗ Error converting {image_path}: {str(e)}")
            return None
    
    def convert_multiple_images(self, image_paths, output_path, layout='vertical'):
        try:
            c = canvas.Canvas(output_path, pagesize=self.page_size)
            page_width, page_height = self.page_size
            
            if layout == 'vertical':
                self._create_vertical_layout(c, image_paths, page_width, page_height)
            elif layout == 'horizontal':
                self._create_horizontal_layout(c, image_paths, page_width, page_height)
            elif layout == 'grid':
                self._create_grid_layout(c, image_paths, page_width, page_height)
            else:
                print(f"Unknown layout: {layout}. Using vertical layout.")
                self._create_vertical_layout(c, image_paths, page_width, page_height)
            
            c.save()
            print(f"✓ Created multi-image PDF: {output_path}")
            return output_path
            
        except Exception as e:
            print(f"✗ Error creating multi-image PDF: {str(e)}")
            return None
    
    def _create_vertical_layout(self, canvas_obj, image_paths, page_width, page_height):
        for i, img_path in enumerate(image_paths):
            if i > 0:
                canvas_obj.showPage()
            
            try:
                with Image.open(img_path) as img:
                    if img.mode in ('RGBA', 'LA', 'P'):
                        img = img.convert('RGB')
                    
                    img_width, img_height = img.size
                    scale_x = (page_width - 2 * self.margin) / img_width
                    scale_y = (page_height - 2 * self.margin) / img_height
                    scale = min(scale_x, scale_y)
                    
                    final_width = img_width * scale
                    final_height = img_height * scale
                    x = (page_width - final_width) / 2
                    y = (page_height - final_height) / 2
                    
                    canvas_obj.drawImage(img_path, x, y, final_width, final_height)
                    
            except Exception as e:
                print(f"✗ Error processing {img_path}: {str(e)}")
                continue
    
    def _create_horizontal_layout(self, canvas_obj, image_paths, page_width, page_height):
        self._create_vertical_layout(canvas_obj, image_paths, page_width, page_height)
    
    def _create_grid_layout(self, canvas_obj, image_paths, page_width, page_height):
        self._create_vertical_layout(canvas_obj, image_paths, page_width, page_height)


def main():
    parser = argparse.ArgumentParser(
        description='Convert images to PDF format',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python image_to_pdf.py image.jpg
  python image_to_pdf.py image.jpg -o output.pdf
  python image_to_pdf.py *.jpg -m output.pdf
  python image_to_pdf.py *.jpg -m output.pdf -l grid
  python image_to_pdf.py *.jpg -m output.pdf -s letter
        """
    )
    
    parser.add_argument('images', nargs='+', help='Input image file(s)')
    parser.add_argument('-o', '--output', help='Output PDF file path')
    parser.add_argument('-m', '--multiple', help='Output PDF for multiple images')
    parser.add_argument('-s', '--size', default='A4', 
                       help='Page size (A4, letter, or WxH in inches)')
    parser.add_argument('-l', '--layout', default='vertical',
                       choices=['vertical', 'horizontal', 'grid'],
                       help='Layout for multiple images')
    parser.add_argument('--margin', type=float, default=0.5,
                       help='Margin in inches (default: 0.5)')
    
    args = parser.parse_args()
    
    valid_images = []
    for img_path in args.images:
        if os.path.isfile(img_path):
            try:
                with Image.open(img_path) as img:
                    valid_images.append(img_path)
            except:
                print(f"✗ Skipping {img_path}: Not a valid image file")
        else:
            print(f"✗ Skipping {img_path}: File not found")
    
    if not valid_images:
        print("No valid image files found!")
        return
    
    converter = ImageToPdfConverter(page_size=args.size, margin=args.margin)
    
    if len(valid_images) == 1 and not args.multiple:
        converter.convert_single_image(valid_images[0], args.output)
    else:
        if not args.multiple:
            args.multiple = 'output.pdf'
        converter.convert_multiple_images(valid_images, args.multiple, args.layout)


if __name__ == '__main__':
    main()
