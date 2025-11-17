# Quick fix - adds font sizing to pdf_to_cmr.py

with open('pdf_to_cmr.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Find the save line and add font code before it
old_code = '''        self.wb.save(output_path)
        print(f"✓ CMR saved: {output_path}")'''

new_code = '''        # Apply font size 13 to all cells
        print("  [FINAL] Applying font size 13...")
        default_font = Font(name='Arial', size=13)
        for row in self.ws.iter_rows():
            for cell in row:
                if cell.value:
                    cell.font = default_font
        
        self.wb.save(output_path)
        print(f"✓ CMR saved: {output_path}")'''

content = content.replace(old_code, new_code)

with open('pdf_to_cmr.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("✓ Font fix applied!")