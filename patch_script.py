import re

with open(r'e:\Fproject\build_synopsis.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Replace em-dashes
code = code.replace('—', '-')

# 2. Fix heading sizes (Level 1 = 14, Level 2 = 12)
# Original: set_font(r, size=12, bold=True, underline=True)
code = re.sub(
    r"if level == 1:\s+set_font\(r, size=12, bold=True, underline=True\)",
    "if level == 1:\n        set_font(r, size=14, bold=True)",
    code
)

# 3. Remove Figures 2.1, 3.1, 4.1 and their captions
# Fig 2.1 (image4.png)
code = re.sub(r"doc\.add_picture\('ppt_images/image4\.png'.*?italic=True\)", "", code, flags=re.DOTALL)
# Fig 3.1 (image6.png)
code = re.sub(r"doc\.add_picture\('ppt_images/image6\.png'.*?italic=True\)", "", code, flags=re.DOTALL)
# Fig 4.1 (image5.png)
code = re.sub(r"doc\.add_picture\('ppt_images/image5\.png'.*?italic=True\)", "", code, flags=re.DOTALL)

# 4. Update numbering for remaining figures and tables
code = code.replace('Figure 1.1:', 'Figure 1:')
code = code.replace('Figure 6.1:', 'Figure 2:')
code = code.replace('Figure 6.1 below', 'Figure 2 below')
code = code.replace('Table 5.1:', 'Table 1:')
code = code.replace('Table 7.1:', 'Table 2:')

# 5. Change output filename
code = code.replace('ELEVATE_Synopsis_Final_NoAI.docx', 'ELEVATE_Synopsis_Final_Clean.docx')

with open(r'e:\Fproject\build_synopsis_clean.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("Created build_synopsis_clean.py")
