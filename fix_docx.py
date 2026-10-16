from docx import Document
from docx.shared import Pt
import re

def fix_formatting(filepath, outpath):
    doc = Document(filepath)
    fig_counter = 1
    table_counter = 1
    
    for p in doc.paragraphs:
        text = p.text.strip()
        
        # Fix Headings
        if re.match(r'^\d+\.\s+[A-Za-z]', text):
            # Main heading (e.g., "1. Introduction")
            for run in p.runs:
                run.font.size = Pt(14)
                run.font.bold = True
        elif re.match(r'^\d+\.\d+\s+[A-Za-z]', text):
            # Subheading (e.g., "5.1 Existing Learning")
            for run in p.runs:
                run.font.size = Pt(12)
                run.font.bold = True
                
        # Fix Figure Captions
        if text.startswith('Figure'):
            # Replace "Figure X.X:" with "Figure N:"
            new_text = re.sub(r'^Figure\s+\d+\.\d+:', f'Figure {fig_counter}:', text)
            p.text = new_text
            for run in p.runs:
                run.font.size = Pt(10)
                if 'Figure' in run.text:
                    run.font.bold = True
                else:
                    run.font.italic = True
            fig_counter += 1
            
        # Fix Table Captions
        if text.startswith('Table'):
            new_text = re.sub(r'^Table\s+\d+\.\d+:', f'Table {table_counter}:', text)
            p.text = new_text
            for run in p.runs:
                run.font.size = Pt(10)
                if 'Table' in run.text:
                    run.font.bold = True
                else:
                    run.font.italic = True
            table_counter += 1
            
        # Fix inline reference to Figure 6.1
        if 'Figure 6.1' in text:
            # We know Figure 6.1 became Figure 2
            p.text = text.replace('Figure 6.1', 'Figure 2')
            for run in p.runs:
                run.font.size = Pt(12)
                
    doc.save(outpath)
    print("Formatted successfully:", outpath)

if __name__ == '__main__':
    fix_formatting(r'e:\Fproject\ELEVATE_Synopsis_Final_NoAI.docx', r'e:\Fproject\ELEVATE_Synopsis_Final_Formatted.docx')
