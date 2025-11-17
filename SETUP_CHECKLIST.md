# Setup Checklist - CTS PDF to CMR Converter

## ✅ Pre-Installation Checklist

### Before you begin:

- [ ] **Python Installed?**
  - Version 3.8 or higher required
  - Test: Open command prompt and type `python --version`
  - Download from: https://www.python.org/downloads/
  - ⚠️ **IMPORTANT**: Check "Add Python to PATH" during installation!

- [ ] **Have Sample PDFs Ready?**
  - At least one CTS packing list PDF for testing
  - Example: `Packing_List_5523.pdf`

- [ ] **Have CMR Template? (Optional)**
  - Your existing CMR Excel template
  - Or the tool will create a basic one

---

## 📦 Installation Steps

### Step 1: Extract Files
- [ ] Download/extract all files to a folder
- [ ] Example: `C:\CTS_Converter\`

### Step 2: Install Dependencies

**Windows:**
- [ ] Double-click `install_windows.bat`
- [ ] Wait for "Installation Complete!" message

**Mac/Linux:**
- [ ] Open terminal in the folder
- [ ] Run: `pip install pdfplumber openpyxl`

### Step 3: Verify Installation
- [ ] Open command prompt in the converter folder
- [ ] Run: `python pdf_to_cmr.py --help`
- [ ] If no errors, you're ready! ✅

---

## 🧪 Testing (First Use)

### Test 1: Run the Example Demo
```bash
python example_demo.py
```
- [ ] Demo runs successfully
- [ ] Creates `CTS_NL_CMR_Template.xlsx`
- [ ] Creates `CMR_EXAMPLE_*.xlsx`

### Test 2: Convert Your First Real PDF

**Option A: Using GUI (Recommended for first time)**
```bash
python pdf_to_cmr_gui.py
```
- [ ] GUI window opens
- [ ] Click "Browse..." and select a packing list PDF
- [ ] Click "Convert to CMR"
- [ ] CMR file created successfully
- [ ] Open and verify the output

**Option B: Using Command Line**
```bash
python pdf_to_cmr.py 5523
```
(Replace 5523 with your actual packing list number)

- [ ] Command runs without errors
- [ ] CMR Excel file created
- [ ] Open and verify the output

---

## 🎯 Configuration Checklist

### Folder Structure Setup
- [ ] Create input folder: `./packing_lists/`
- [ ] Create output folder: `./cmr_output/`
- [ ] Place your CMR template in the main folder

### Verify Data Extraction
Open the generated CMR file and check:
- [ ] Packing list number is correct
- [ ] Date is extracted properly
- [ ] Consignee information is complete
- [ ] All items are listed
- [ ] Quantities are accurate
- [ ] HS codes are present

### If Data is Incorrect:
- [ ] Check that PDF follows CTS format
- [ ] Verify PDF is not corrupted
- [ ] Check PDF is not password-protected

---

## 🚀 Production Setup Checklist

### For Daily Use:

- [ ] **Create Desktop Shortcuts**
  - Right-click `pdf_to_cmr_gui.py` → Send to → Desktop

- [ ] **Set Default Folders**
  - Edit `pdf_to_cmr_gui.py` if needed
  - Change default output directory

- [ ] **Establish File Naming Convention**
  - Ensure PDFs are named: `Packing_List_XXXXX.pdf`
  - Or be prepared to use full paths

- [ ] **Test Batch Processing**
  ```bash
  python batch_convert.py ./packing_lists ./cmr_output
  ```
  - [ ] Multiple PDFs converted successfully

---

## 👥 Team Rollout Checklist

### For Each Team Member:

- [ ] **Install on their computer**
  - Run `install_windows.bat` on each machine
  
- [ ] **Provide training**
  - Show GUI usage (5 minutes)
  - Demonstrate command line (optional)
  - Show where outputs are saved

- [ ] **Share documentation**
  - Give them `QUICKSTART.md`
  - Point to `README.md` for details

- [ ] **Test with their PDFs**
  - Each person converts one PDF
  - Verify output looks correct

### Network Setup (Optional):

- [ ] **Set up shared folders**
  - Central input folder for PDFs
  - Central output folder for CMR files
  
- [ ] **Update scripts with network paths**
  - Edit default folders in scripts
  
- [ ] **Test multi-user access**
  - Ensure multiple people can use simultaneously

---

## 🔧 Customization Checklist

### Template Customization:

- [ ] **Review default cell mappings**
  - Open `pdf_to_cmr.py`
  - Find `_populate_header()` method
  - Verify cells match your template

- [ ] **Adjust if needed**
  - Change cell references (e.g., 'G3' to 'H3')
  - Add new fields if required

- [ ] **Test custom template**
  - Run conversion with your template
  - Verify all data appears correctly

### PDF Format Changes:

If your packing list format changes:
- [ ] Update extraction patterns in `pdf_to_cmr.py`
- [ ] Test with new format PDFs
- [ ] Document any changes made

---

## ✅ Go-Live Checklist

### Final Steps Before Full Deployment:

- [ ] **Tested with 5+ different packing lists**
  - All formats working correctly
  
- [ ] **Compared outputs with manual CMR documents**
  - Data accuracy verified
  
- [ ] **Team trained and comfortable**
  - Everyone can use the tool
  
- [ ] **Backup plan established**
  - Keep manual process available initially
  
- [ ] **Performance tested**
  - Processing speed acceptable
  - No freezing or crashes

### First Week Monitoring:

- [ ] **Track usage**
  - How many PDFs converted?
  - Any errors reported?
  
- [ ] **Collect feedback**
  - What works well?
  - What needs improvement?
  
- [ ] **Document issues**
  - Keep log of any problems
  - Note solutions

---

## 📊 Success Metrics

After one week, evaluate:

- [ ] **Time savings**
  - Compare old vs new process time
  - Calculate time saved per document

- [ ] **Error rate**
  - How many required corrections?
  - Are errors decreasing?

- [ ] **User satisfaction**
  - Are users happy with the tool?
  - Is it easier than before?

### Expected Results:
- ✅ 90%+ time reduction per document
- ✅ No more Excel freezing issues
- ✅ No Navision connection problems
- ✅ Immediate CMR document availability

---

## 🆘 Support Checklist

### If Problems Occur:

- [ ] **Check the troubleshooting guide** (README.md)
- [ ] **Verify Python and packages are installed**
- [ ] **Test with the example demo**
- [ ] **Check PDF file is readable**
- [ ] **Try a different PDF**

### Common Issues Quick Reference:

| Issue | Solution |
|-------|----------|
| "Python not found" | Reinstall Python, add to PATH |
| "Module not found" | Run: `pip install pdfplumber openpyxl` |
| "PDF not found" | Check filename or use full path |
| GUI won't open | Use command line instead |
| Wrong data extracted | Verify PDF format matches expected |

---

## 🎓 Training Materials Prepared

- [ ] **Quick reference card** (print QUICKSTART.md)
- [ ] **Video demo recorded** (optional)
- [ ] **FAQ document created** (optional)
- [ ] **Contact person assigned** (for questions)

---

## ✨ You're Ready!

Once all checkboxes above are complete, you're ready for full deployment!

**Quick Status Check:**

Essential items (Must have ✅):
- [ ] Python installed
- [ ] Dependencies installed  
- [ ] Test conversion successful
- [ ] Output verified correct

Recommended items (Should have):
- [ ] GUI working
- [ ] Batch processing tested
- [ ] Team trained
- [ ] Documentation distributed

Optional items (Nice to have):
- [ ] Desktop shortcuts created
- [ ] Network folders configured
- [ ] Custom templates set up
- [ ] Automation scripts created

**Minimum to start:** Just the essential items! ✅

---

**Date Completed:** _______________

**Completed By:** _______________

**Notes:**
_________________________________________________________________
_________________________________________________________________
_________________________________________________________________

