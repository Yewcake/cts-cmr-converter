#!/usr/bin/env python3
"""
Batch PDF to CMR Converter
Process multiple packing list PDFs in a directory
"""

import os
import sys
import glob
from datetime import datetime

# Import from the main script
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pdf_to_cmr import PackingListExtractor, CMRExcelPopulator


def process_directory(input_dir: str, output_dir: str, template_path: str):
    """Process all PDFs in a directory"""
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Find all PDF files
    pdf_files = glob.glob(os.path.join(input_dir, "*.pdf"))
    pdf_files += glob.glob(os.path.join(input_dir, "Packing_List_*.pdf"))
    pdf_files = list(set(pdf_files))  # Remove duplicates
    
    if not pdf_files:
        print(f"No PDF files found in {input_dir}")
        return
    
    print(f"Found {len(pdf_files)} PDF file(s) to process\n")
    
    successful = 0
    failed = 0
    
    for pdf_path in sorted(pdf_files):
        try:
            print(f"Processing: {os.path.basename(pdf_path)}")
            
            # Extract data
            extractor = PackingListExtractor(pdf_path)
            data = extractor.extract()
            
            # Generate output filename
            packing_list_no = data.get('our_ref') or data.get('packing_list_number')
            if not packing_list_no:
                packing_list_no = os.path.splitext(os.path.basename(pdf_path))[0]
            
            timestamp = datetime.now().strftime('%Y%m%d')
            output_filename = f"CMR_{packing_list_no}_{timestamp}.xlsx"
            output_path = os.path.join(output_dir, output_filename)
            
            # Populate template
            populator = CMRExcelPopulator(template_path)
            populator.populate(data, output_path)
            
            print(f"  ✓ Success: {output_filename}\n")
            successful += 1
            
        except Exception as e:
            print(f"  ✗ Error: {e}\n")
            failed += 1
    
    # Summary
    print("="*60)
    print(f"Processing complete!")
    print(f"  Successful: {successful}")
    print(f"  Failed: {failed}")
    print(f"  Output directory: {output_dir}")
    print("="*60)


def main():
    """Main execution"""
    
    if len(sys.argv) < 2:
        print("Batch PDF to CMR Converter")
        print("\nUsage:")
        print("  python batch_convert.py <input_directory> [output_directory] [template_path]")
        print("\nExamples:")
        print("  python batch_convert.py ./packing_lists")
        print("  python batch_convert.py ./packing_lists ./output_cmr")
        print("  python batch_convert.py ./packing_lists ./output ./template.xlsx")
        sys.exit(1)
    
    input_dir = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "./cmr_output"
    template_path = sys.argv[3] if len(sys.argv) > 3 else "CTS_NL_CMR_Template.xlsx"
    
    if not os.path.isdir(input_dir):
        print(f"Error: Input directory does not exist: {input_dir}")
        sys.exit(1)
    
    print("Batch PDF to CMR Converter")
    print(f"Input directory: {input_dir}")
    print(f"Output directory: {output_dir}")
    print(f"Template: {template_path}")
    print("="*60 + "\n")
    
    process_directory(input_dir, output_dir, template_path)


if __name__ == "__main__":
    main()
