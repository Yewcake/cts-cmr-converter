# CTS Packing List to CMR Excel Converter

A Python-based solution to extract data from CTS packing list PDFs and automatically populate CMR (Convention Relative au Contrat de Transport International de Marchandises par Route) Excel documents.

## 🎯 Purpose

This tool solves the problem of Excel freezing when using JetReports to fetch data from Navision. Instead of querying Navision directly, it extracts information from PDF packing lists and populates the CMR template, providing a faster and more reliable workflow.

## ✨ Features

- **PDF Data Extraction**: Automatically extracts all relevant information from CTS packing list PDFs
- **Excel Population**: Fills CMR Excel template with extracted data
- **Multiple Interfaces**:
  - Command-line interface for automation
  - PowerShell script for Windows users
  - Graphical User Interface (GUI) for easy use
  - Batch processing for multiple PDFs
- **No Navision Required**: Works directly with PDF files, avoiding database connection issues

## 📋 Requirements

### System Requirements
- Python 3.8 or higher
- Windows, macOS, or Linux

### Python Dependencies
```
pdfplumber
openpyxl
tkinter (usually included with Python)
```

## 🚀 Installation

### Step 1: Install Python
Download and install Python from [python.org](https://www.python.org/downloads/)

**Important for Windows users**: During installation, check "Add Python to PATH"

### Step 2: Install Dependencies
Open terminal/command prompt and run:

```bash
pip install pdfplumber openpyxl
```

### Step 3: Download the Scripts
Place all the following files in the same directory:
- `pdf_to_cmr.py` - Main conversion script
- `pdf_to_cmr_gui.py` - GUI application
- `batch_convert.py` - Batch processor
- `convert_pdf_to_cmr.ps1` - PowerShell script (Windows)

## 📖 Usage

### Method 1: Graphical User Interface (Easiest)

**For Windows users:**
1. Double-click `pdf_to_cmr_gui.py`
2. Click "Browse..." to select your packing list PDF
3. (Optional) Select your CMR template
4. Choose output directory
5. Click "Convert to CMR"

**For batch processing:**
- Click "Batch Process Folder" and select a folder containing multiple PDFs

### Method 2: Command Line

**Basic usage with packing list number:**
```bash
python pdf_to_cmr.py 5523
```

This assumes your PDF is named `Packing_List_5523.pdf`

**Usage with full file path:**
```bash
python pdf_to_cmr.py /path/to/Packing_List_5523.pdf
```

### Method 3: PowerShell (Windows)

```powershell
.\convert_pdf_to_cmr.ps1 -Input 5523
```

Or with a file path:
```powershell
.\convert_pdf_to_cmr.ps1 -Input "C:\Documents\Packing_List_5523.pdf"
```

### Method 4: Batch Processing

Process all PDFs in a directory:

```bash
python batch_convert.py ./packing_lists ./output
```

Parameters:
- First argument: Input directory containing PDFs
- Second argument: Output directory for CMR files (optional)

## 📁 File Organization

### Recommended Folder Structure

```
project/
├── pdf_to_cmr.py              # Main script
├── pdf_to_cmr_gui.py           # GUI application
├── batch_convert.py            # Batch processor
├── convert_pdf_to_cmr.ps1     # PowerShell script
├── CTS_NL_CMR_Template.xlsx   # Your CMR template
├── packing_lists/              # Input PDFs
│   ├── Packing_List_5523.pdf
│   ├── Packing_List_5524.pdf
│   └── ...
└── cmr_output/                 # Generated CMR files
    ├── CMR_5523_20250722.xlsx
    ├── CMR_5524_20250722.xlsx
    └── ...
```

## 🔧 Configuration

### Custom Template Path

If your CMR template is in a different location, specify it:

**Command line:**
```bash
python pdf_to_cmr.py 5523 --template /path/to/template.xlsx
```

**GUI:**
- Use the "Browse..." button next to "CMR Template"

### Custom Output Directory

**Command line:**
```bash
python pdf_to_cmr.py 5523 --output /path/to/output
```

**GUI:**
- Use the "Browse..." button next to "Output Directory"

## 📊 What Data is Extracted?

The tool extracts the following information from packing list PDFs:

### Header Information
- Packing list number
- Date
- Your reference
- Our reference (5-digit number)

### Consignee Details
- Company name
- Address lines
- Telephone
- Fax

### Shipping Information
- Delivery terms (e.g., EXW Barendrecht)
- Case information (dimensions, volume)
- Measurements
- Gross weight
- Country of origin

### Item List
For each item:
- Description
- Article number
- HS Code
- Quantity

## 🐛 Troubleshooting

### "PDF file not found"
- Ensure your PDF follows the naming convention: `Packing_List_XXXXX.pdf`
- Or provide the full path to the PDF file

### "Module not found" errors
```bash
pip install pdfplumber openpyxl
```

### "Excel file is corrupted"
- Ensure you're using a valid .xlsx template
- Try creating a new template from Excel

### GUI won't open
- Ensure tkinter is installed: `python -m tkinter`
- On Linux: `sudo apt-get install python3-tk`

### PDF extraction errors
- Verify the PDF is not corrupted
- Ensure the PDF follows the CTS packing list format
- Check if PDF is password-protected

## 💡 Tips for Best Results

1. **Consistent Naming**: Name your PDFs consistently (e.g., `Packing_List_5523.pdf`)
2. **Template Setup**: Configure your CMR template once with proper cell references
3. **Batch Processing**: Use batch mode for multiple packing lists to save time
4. **Backup**: Keep original PDFs in case you need to re-process them

## 🔄 Workflow Integration

### Current Workflow (with JetReports - problematic):
1. Open Excel template
2. Enter 5-digit packing list number
3. JetReports queries Navision
4. Excel often freezes
5. Manual restart required

### New Workflow (with this tool):
1. Export/save packing list as PDF (if not already available)
2. Run converter:
   - GUI: Click and select PDF
   - Command: `python pdf_to_cmr.py 5523`
3. CMR Excel file is generated instantly
4. No freezing, no Navision connection needed

## 🔐 Data Privacy

- All processing happens locally on your computer
- No data is sent to external servers
- PDF and Excel files remain on your system

## 📝 Customization

### Modifying Cell Mappings

Edit `pdf_to_cmr.py` to customize where data is placed in the Excel template:

```python
def _populate_header(self, data: Dict):
    cells = {
        'G3': data.get('packing_list_number', ''),
        'G4': data.get('date', ''),
        # Add or modify cell references here
    }
```

### Adding Custom Fields

To extract additional fields from PDFs:

1. Add extraction method in `PackingListExtractor` class
2. Update `_populate_*` methods in `CMRExcelPopulator` class

## 🤝 Support

For issues or questions:
1. Check the troubleshooting section above
2. Verify your Python and package versions
3. Ensure PDFs follow the expected format

## 📜 License

This tool is provided as-is for internal use at CTS Netherlands B.V.

## 🔄 Version History

### Version 1.0 (2025-07-22)
- Initial release
- PDF extraction functionality
- Excel population
- GUI and batch processing
- PowerShell script for Windows

---

**Note**: This tool is designed specifically for CTS packing list PDFs. If your PDF format changes, the extraction patterns may need to be updated.
