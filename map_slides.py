import sys, io, zipfile, re, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

pptx_path = r'e:\Fproject\Final_1 (1).pptx'
with zipfile.ZipFile(pptx_path, 'r') as z:
    all_files = z.namelist()
    slide_files = sorted([f for f in all_files if f.startswith('ppt/slides/slide') and 'rels' not in f])
    for sf in slide_files:
        slide_name = sf.split('/')[-1]
        rels_path = 'ppt/slides/_rels/' + slide_name + '.rels'
        if rels_path in all_files:
            with z.open(rels_path) as f:
                rels = f.read().decode('utf-8')
                imgs = re.findall(r'Target="../media/([^"]+)"', rels)
                print(slide_name + ': ' + str(imgs))
        else:
            print(slide_name + ': (no rels file)')
