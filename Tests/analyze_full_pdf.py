# -*- coding: utf-8 -*-
# Analyse complete du PDF 02-Machines-.pdf

import fitz
import re
from collections import defaultdict

TAG_PATTERN = re.compile(r'([A-Z]{2,4}\d{3,4}[-_]\d{3,4})', re.IGNORECASE)

pdf_path = r'C:\Vault\Engineering\Projects\10381\REF13\M02\6-Shop Drawing PDF\Production\BatchPrint\02-Machines-.pdf'

doc = fitz.open(pdf_path)
print(f'PDF: 02-Machines-.pdf ({len(doc)} pages)')
print('=' * 70)

# Extraire TOUT avec l'algorithme de proximite
tag_qty_extracted = {}
tag_occurrences = defaultdict(list)

for page_num, page in enumerate(doc):
    words = page.get_text('words')
    lines = defaultdict(list)
    for w in words:
        y_key = round(w[1] / 5) * 5
        lines[y_key].append({'text': w[4], 'x': w[0]})
    
    for y_key in sorted(lines.keys()):
        line_words = sorted(lines[y_key], key=lambda w: w['x'])
        tags = []
        numbers = []
        for w in line_words:
            m = TAG_PATTERN.match(w['text'])
            if m:
                tags.append({'tag': m.group(1).upper().replace('_','-'), 'x': w['x']})
            if w['text'].isdigit() and len(w['text']) <= 3:
                numbers.append({'qty': int(w['text']), 'x': w['x']})
        
        for t in tags:
            tag = t['tag']
            best_qty = None
            best_dist = 999999
            
            # Droite d'abord
            for n in numbers:
                if n['x'] > t['x']:
                    dist = n['x'] - t['x']
                    if dist < best_dist and dist < 150:
                        best_dist = dist
                        best_qty = n['qty']
            # Sinon gauche
            if best_qty is None:
                for n in numbers:
                    if n['x'] < t['x']:
                        dist = t['x'] - n['x']
                        if dist < best_dist and dist < 150:
                            best_dist = dist
                            best_qty = n['qty']
            
            if best_qty:
                tag_occurrences[tag].append({'page': page_num+1, 'qty': best_qty})
                if tag not in tag_qty_extracted:
                    tag_qty_extracted[tag] = best_qty

doc.close()

print(f'Tags uniques extraits: {len(tag_qty_extracted)}')
print()

# Afficher les tags avec leurs quantites
print('TOUS LES TAGS EXTRAITS:')
print('-' * 70)
print(f'{"Tag":<20} {"Qty":<8} Occurrences')
print('-' * 70)

for tag in sorted(tag_qty_extracted.keys()):
    qty = tag_qty_extracted[tag]
    occs = tag_occurrences[tag]
    occ_str = ', '.join([f'p{o["page"]}:{o["qty"]}' for o in occs[:5]])
    if len(occs) > 5:
        occ_str += f'... (+{len(occs)-5})'
    print(f'{tag:<20} {qty:<8} {occ_str}')

# Verifier coherence
print()
print('=' * 70)
print('VERIFICATION COHERENCE:')
inconsistent = []
for tag, occs in tag_occurrences.items():
    qtys = set(o['qty'] for o in occs)
    if len(qtys) > 1:
        inconsistent.append((tag, qtys, occs))

if inconsistent:
    print(f'[!] {len(inconsistent)} tags avec quantites DIFFERENTES selon la page:')
    for tag, qtys, occs in inconsistent[:10]:
        print(f'  {tag}: qtys={qtys}')
        for o in occs[:4]:
            print(f'    Page {o["page"]}: qty={o["qty"]}')
else:
    print('[+] Tous les tags ont une quantite coherente sur toutes les pages')

# Comparer avec CSV
print()
print('=' * 70)
print('COMPARAISON AVEC CSV:')

csv_path = r'C:\Vault\Engineering\Projects\10381\REF13\M02\5_Exportation\Sheet_Metal_Nesting\Punch\10381-13-M02.csv'
csv_tags = {}
with open(csv_path, 'r', encoding='utf-8-sig') as f:
    for line in f:
        parts = re.split(r'[;,]', line)
        if len(parts) >= 2:
            try:
                qty = int(parts[0].strip())
                m = TAG_PATTERN.search(parts[1].strip())
                if m:
                    csv_tags[m.group(1).upper().replace('_','-')] = qty
            except:
                pass

print(f'CSV: {len(csv_tags)} tags')
print(f'PDF: {len(tag_qty_extracted)} tags')

correct = sum(1 for t, q in tag_qty_extracted.items() if csv_tags.get(t) == q)
wrong = [(t, csv_tags[t], tag_qty_extracted[t]) for t in tag_qty_extracted if t in csv_tags and csv_tags[t] != tag_qty_extracted[t]]
missing_in_pdf = [t for t in csv_tags if t not in tag_qty_extracted]
extra_in_pdf = [t for t in tag_qty_extracted if t not in csv_tags]

print(f'Correct: {correct}/{len(csv_tags)} ({100*correct/len(csv_tags):.1f}%)')
print(f'Wrong Qty: {len(wrong)}')
print(f'Dans CSV mais pas PDF: {len(missing_in_pdf)}')
print(f'Dans PDF mais pas CSV: {len(extra_in_pdf)}')

if wrong:
    print()
    print('Erreurs de quantite:')
    for t, csv_q, pdf_q in wrong:
        print(f'  {t}: CSV={csv_q}, PDF={pdf_q}')

if extra_in_pdf:
    print()
    print(f'Tags EXTRA dans PDF (pas dans CSV): {len(extra_in_pdf)}')
    for t in extra_in_pdf:
        print(f'  {t}: qty={tag_qty_extracted[t]}')
