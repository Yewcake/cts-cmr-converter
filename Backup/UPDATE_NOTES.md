# UPDATE SUMMARY - CMR Template Cell Mapping Fix

## What Was Updated

Based on your actual CMR template image, I've updated the scripts to map data to the **correct cells**.

---

## Cell Mapping (Exact Locations)

### Header Information
| Data | Cell | Example |
|------|------|---------|
| Packing List Number | E2 | 76988-1 |
| Date | E3 | 22-07-2025 |
| Your Reference | E4 | 01/NIPO/19465 |
| Our Reference | E5 | 5523 |

### Consignee/Delivery Address
Starting at **Row 10**, Column A:
- Row 10: Company Name
- Row 11: Address Line 1
- Row 12: Address Line 2
- Row 13: TEL / FAX info

### Shipping Information (Incoterms, Packages, Weight)
| Data | Cell | Example |
|------|------|---------|
| Delivery Terms | B15 | EXW Barendrecht (NL) |
| Case Info (Packages) | B16 | Case 1 of 1 / 2.63 m3 |
| Measurements | B17 | 160 x 160 x 103 cm |
| Gross Weight | B18 | 287 KG |

### Items Table
Starting at **Row 22**:
| Column | Data | Example |
|--------|------|---------|
| A | Description | CTS60 or CTS70 IFR wiper seal... |
| D | Article Number | 60WS350XPE |
| E | HS Code | 40169990 |
| F | Quantity | 4 Rolls (31.2m each) |

---

## Files Updated

1. **pdf_to_cmr.py** - Main converter script
   - Updated `_populate_header()` to use E2-E5
   - Updated `_populate_consignee()` to start at A10
   - Updated `_populate_shipping_info()` to use B15-B18
   - Updated `_populate_items()` to use columns A, D, E, F starting at row 22

2. **pdf_to_cmr_gui.py** - GUI application
   - Uses the updated main script
   - No changes needed (imports from main script)

3. **batch_convert.py** - Batch processor
   - Uses the updated main script
   - No changes needed (imports from main script)

---

## What Gets Extracted from PDF

### From Header:
✅ Packing list number (e.g., "15880-1")
✅ Date (e.g., "22-07-2025")
✅ Your reference (e.g., "01/NIPO/19465")
✅ Our reference / 5-digit number (e.g., "5523")

### From Consignee Section:
✅ Company name (e.g., "Dohat Al Khaleej LLC")
✅ Address line 1 (e.g., "PO BOX 503, PC 133 AL KHUWAIR")
✅ Address line 2 (e.g., "SULTANATE OF OMAN")
✅ Telephone (e.g., "+968 24052867")
✅ Fax (e.g., "+968 24054165")

### From Shipping Info:
✅ Delivery terms / Incoterms (e.g., "EXW Barendrecht (NL)")
✅ Case info / Number of packages (e.g., "Case 1 of 1 / 2.63 m3")
✅ Measurements / Dimensions (e.g., "160 x 160 x 103 cm")
✅ Gross weight / Total KG (e.g., "287 KG")
✅ Country of origin (e.g., "NL")

### From Items Table:
For each item:
✅ Description (full text, e.g., "CTS60 or CTS70 IFR wiper seal XPE...")
✅ Article number (e.g., "60WS350XPE")
✅ HS Code (e.g., "40169990")
✅ Quantity (e.g., "4 Rolls (31.2m each)")

---

## How to Use the Updated Scripts

### Option 1: Command Line
```bash
python pdf_to_cmr.py Packing_List_5523.pdf
```

### Option 2: GUI (Easiest)
```bash
python pdf_to_cmr_gui.py
```
Then:
1. Click "Browse" to select your PDF
2. Click "Convert to CMR"
3. Done!

### Option 3: Batch Process
```bash
python batch_convert.py ./packing_lists ./output
```

---

## Data Format Notes

### Per Line Items Format
Each item appears on **one line** in the CMR:
- Row 22: First item
- Row 23: Second item
- Row 24: Third item
- And so on...

### Text Wrapping
Long descriptions in column A will wrap automatically if needed.

### Package Info Format
The script preserves the exact format from the PDF:
- "Case 1 of 1 / 2.63 m3" ✅
- "160 x 160 x 103 cm" ✅
- "287 KG" ✅

---

## Testing Checklist

After using the updated scripts, verify:

- [ ] Packing list number appears in cell E2
- [ ] Date appears in cell E3
- [ ] References appear in cells E4-E5
- [ ] Consignee info starts at row 10, column A
- [ ] Delivery terms in cell B15
- [ ] Package info in cell B16
- [ ] Measurements in cell B17
- [ ] Weight in cell B18
- [ ] Items table starts at row 22
- [ ] All item data is in correct columns (A, D, E, F)
- [ ] Print area fits on A4 CMR paper

---

## Print Area Note

The template is designed to print on **A4 CMR paper**. The cell mapping ensures:
- Header fits in the top section
- Consignee in the designated area
- Shipping info in the left column
- Items table in the main body
- All data within A4 boundaries

---

## Need to Adjust Cell Mappings?

If you need to fine-tune the exact cells, edit `pdf_to_cmr.py`:

```python
def _populate_header(self, data: Dict):
    # Change these cell references as needed:
    if data.get('packing_list_number'):
        self.ws['E2'] = data['packing_list_number']  # ← Change E2
    # ... etc
```

---

## Common Issues & Solutions

### Issue: Data not in right cells
**Solution**: Check your template matches the structure in the image you provided

### Issue: Items not showing
**Solution**: Verify PDF table structure is consistent with CTS format

### Issue: Template not loading
**Solution**: Script will create a basic template automatically

---

**Status: ✅ All scripts updated with correct cell mappings!**

Ready to use immediately. The exact cell positions now match your actual CMR template.

