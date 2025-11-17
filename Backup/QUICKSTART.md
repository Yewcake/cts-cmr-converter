# Quick Start Guide - CTS PDF to CMR Converter

## 🚀 Get Started in 3 Minutes

### Step 1: Install (One-time setup)

**Windows Users:**
1. Double-click `install_windows.bat`
2. Wait for installation to complete

**Mac/Linux Users:**
```bash
pip install pdfplumber openpyxl
```

### Step 2: Run the Converter

#### Option A: Use the GUI (Recommended for most users)

**Windows:**
- Double-click `pdf_to_cmr_gui.py`

**Mac/Linux:**
```bash
python3 pdf_to_cmr_gui.py
```

Then:
1. Click "Browse..." next to "Select Packing List PDF"
2. Select your PDF (e.g., `Packing_List_5523.pdf`)
3. Click "Convert to CMR"
4. Done! Your CMR Excel file is ready

#### Option B: Use Command Line (Fast for single files)

**If your PDF is named `Packing_List_5523.pdf`:**
```bash
python pdf_to_cmr.py 5523
```

**If you want to specify the full path:**
```bash
python pdf_to_cmr.py "C:\Documents\Packing_List_5523.pdf"
```

#### Option C: Batch Process Multiple Files

**GUI Method:**
1. Open `pdf_to_cmr_gui.py`
2. Click "Batch Process Folder"
3. Select folder containing your PDFs
4. All PDFs will be converted automatically

**Command Line Method:**
```bash
python batch_convert.py ./packing_lists ./output
```

### Step 3: Find Your CMR Files

By default, CMR files are saved to:
- `./cmr_output/` folder
- Named as: `CMR_5523_20250722.xlsx`

## 📋 Daily Usage Example

### Scenario: You have 3 packing lists to process

**Method 1 - Quick (Command Line):**
```bash
python pdf_to_cmr.py 5523
python pdf_to_cmr.py 5524
python pdf_to_cmr.py 5525
```

**Method 2 - Easiest (GUI Batch):**
1. Put all PDFs in one folder (e.g., `today_packing_lists/`)
2. Run GUI: `pdf_to_cmr_gui.py`
3. Click "Batch Process Folder"
4. Select the folder
5. All 3 CMR files created in seconds!

## 🎯 Real-World Workflow

### Old Way (with JetReports):
1. Open Excel template ⏱️ 10 seconds
2. Enter packing list number ⏱️ 5 seconds
3. Wait for Navision query ⏱️ 30-120 seconds
4. Excel freezes 😢 ⏱️ Restart needed (2-5 minutes)
5. **Total: 3-6 minutes per document** (when it works)

### New Way (with this tool):
1. Run command: `python pdf_to_cmr.py 5523` ⏱️ 2 seconds
2. CMR file ready! ✅
3. **Total: 2 seconds per document**

**Time saved per document: ~3-5 minutes**
**For 20 documents/day: Save 60-100 minutes daily! 🎉**

## 💡 Pro Tips

### Tip 1: Create a Desktop Shortcut
**Windows:**
1. Right-click `pdf_to_cmr_gui.py`
2. Send to → Desktop (create shortcut)
3. Double-click the shortcut anytime!

### Tip 2: Drag and Drop (Windows)
Create a batch file `convert_pdf.bat`:
```batch
@echo off
python pdf_to_cmr.py %1
pause
```
Now drag any PDF onto `convert_pdf.bat` to convert it!

### Tip 3: Set Default Folders
Edit the GUI script to change default folders:
```python
self.output_dir = StringVar(value="C:/CTS/CMR_Output")
```

## ⚙️ File Naming Convention

Your PDFs should be named:
```
Packing_List_5523.pdf
Packing_List_5524.pdf
Packing_List_5525.pdf
```

If they have different names, use the full file path:
```bash
python pdf_to_cmr.py "C:\Downloads\invoice_5523.pdf"
```

## 🆘 Common Issues & Solutions

### Issue: "Python is not recognized"
**Solution:** Reinstall Python and check "Add Python to PATH"

### Issue: "Module 'pdfplumber' not found"
**Solution:** Run: `pip install pdfplumber openpyxl`

### Issue: "PDF not found"
**Solution:** 
- Check PDF filename matches `Packing_List_XXXXX.pdf`
- Or use full file path

### Issue: GUI won't open
**Solution:**
- Try: `python -m tkinter` (test if tkinter works)
- Use command line instead: `python pdf_to_cmr.py 5523`

## 📞 Need Help?

1. Check `README.md` for detailed documentation
2. Review the troubleshooting section
3. Verify your PDF matches the CTS format

## 🎓 Training for Team

### For New Users:
1. Show them the GUI
2. Demonstrate: Browse → Select PDF → Convert
3. Show them where output files are saved

### For Power Users:
1. Teach command line usage
2. Show batch processing
3. Explain folder structure

### For IT/Admins:
1. Review `pdf_to_cmr.py` source code
2. Customize cell mappings if needed
3. Set up network folder paths

---

## 🏁 You're Ready!

Start converting your packing lists to CMR documents in seconds. No more Excel freezing, no more Navision delays! 🚀

**Most Common Commands:**
```bash
# Single file by number
python pdf_to_cmr.py 5523

# Single file by path
python pdf_to_cmr.py "path/to/file.pdf"

# Batch process
python batch_convert.py ./input_folder ./output_folder

# GUI (easiest)
python pdf_to_cmr_gui.py
```

Happy converting! 😊
