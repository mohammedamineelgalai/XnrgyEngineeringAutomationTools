# Debug WPA1302-0101
import fitz
import re
from collections import defaultdict

TAG_PATTERN = re.compile(r'([A-Z]{2,4}\d{3,4}[-_]\d{3,4})', re.IGNORECASE)

pdf_path = r'C:\Vault\Engineering\Projects\10381\REF13\M02\6-Shop Drawing PDF\Production\BatchPrint\02-Machines-.pdf'
doc = fitz.open(pdf_path)
page = doc[16]  # Page 17

words = page.get_text('words')
lines = defaultdict(list)
for w in words:
    y_key = round(w[1] / 5) * 5
    lines[y_key].append({'text': w[4], 'x': w[0]})

# Trouver la ligne avec WPA1302-0101
target_y = None
for y_key, line_words in lines.items():
    for w in line_words:
        if 'WPA1302-0101' in w['text']:
            target_y = y_key
            break

if target_y:
    print(f'Ligne Y={target_y}:')
    line_words = sorted(lines[target_y], key=lambda w: w['x'])
    
    tags = []
    numbers = []
    for w in line_words:
        m = TAG_PATTERN.match(w['text'])
        if m:
            tags.append({'tag': m.group(1).upper().replace('_','-'), 'x': w['x'], 'text': w['text']})
            print(f'  TAG: {w["text"]} @ x={w["x"]:.0f}')
        if w['text'].isdigit() and len(w['text']) <= 3:
            numbers.append({'qty': int(w['text']), 'x': w['x']})
            print(f'  NUM: {w["text"]} @ x={w["x"]:.0f}')
    
    print()
    print('Tags trouves:', [t['tag'] for t in tags])
    print('Numbers trouves:', [n['qty'] for n in numbers])
    
    # Debug l'algo
    for t in tags:
        if t['tag'] == 'WPA1302-0101':
            print(f'\nDebug pour {t["tag"]} @ x={t["x"]}:')
            for n in numbers:
                dist = n['x'] - t['x']
                print(f'  Number {n["qty"]} @ x={n["x"]}: dist={dist:.0f}, droite={n["x"] > t["x"]}, <150={abs(dist) < 150}')

doc.close()
