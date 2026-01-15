# -*- coding: utf-8 -*-
# PDF TABLE EXTRACTION BENCHMARK v3

import re
import os
from collections import defaultdict

PDF_PATH = r"C:\Vault\Engineering\Projects\10381\REF13\M02\6-Shop Drawing PDF\Production\BatchPrint\02-Machines.pdf"
CSV_PATH = r"C:\Vault\Engineering\Projects\10381\REF13\M02\5_Exportation\Sheet_Metal_Nesting\Punch\10381-13-M02.csv"

TAG_PATTERN = re.compile(r'([A-Z]{2,4}\d{3,4}[-_]\d{3,4})', re.IGNORECASE)

QTY_HEADERS = ['qty', 'quantity', 'quantite', 'qte', 'nb', 'nombre', 'count', 'pcs', 'pieces', 'units', 'amt']
TAG_HEADERS = ['item', 'tag', 'part', 'piece', 'file', 'filename', 'fichier', 'name', 'ref', 'drawing', 'dxf']

def load_csv_reference():
    tag_qty = {}
    if not os.path.exists(CSV_PATH):
        print(f"[-] CSV non trouve: {CSV_PATH}")
        return tag_qty
    with open(CSV_PATH, 'r', encoding='utf-8-sig') as f:
        for line in f:
            parts = re.split(r'[;,]', line)
            if len(parts) >= 2:
                try:
                    qty = int(parts[0].strip())
                except:
                    continue
                match = TAG_PATTERN.search(parts[1].strip())
                if match:
                    tag = match.group(1).upper().replace("_", "-")
                    tag_qty[tag] = qty
    return tag_qty

def evaluate(method, extracted, reference):
    total = len(reference)
    correct = sum(1 for t, q in extracted.items() if reference.get(t) == q)
    wrong = sum(1 for t, q in extracted.items() if t in reference and reference[t] != q)
    missing = [t for t in reference if t not in extracted]
    acc = correct / total * 100 if total > 0 else 0
    return {
        'method': method, 'correct': correct, 'total': total, 
        'wrong': wrong, 'missing': len(missing), 'accuracy': round(acc, 1),
        'missing_tags': missing[:10],
        'wrong_details': [(t, reference[t], extracted[t]) for t in extracted if t in reference and reference[t] != extracted[t]][:10]
    }

def test_auto_structure():
    import pdfplumber
    tag_qty = {}
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                lines = defaultdict(list)
                for w in words:
                    lines[round(w['top'], 0)].append(w)
                
                header_y = None
                qty_col_x = None
                for y in sorted(lines.keys())[:15]:
                    for w in lines[y]:
                        if any(h in w['text'].lower() for h in QTY_HEADERS):
                            header_y = y
                            qty_col_x = w['x0']
                            break
                    if header_y:
                        break
                
                for y in sorted(lines.keys()):
                    if header_y and y <= header_y:
                        continue
                    line_words = sorted(lines[y], key=lambda w: w['x0'])
                    tag = None
                    qty = None
                    for w in line_words:
                        if TAG_PATTERN.match(w['text']):
                            tag = TAG_PATTERN.match(w['text']).group(1).upper().replace("_", "-")
                        if w['text'].isdigit():
                            if qty_col_x and abs(w['x0'] - qty_col_x) < 30:
                                qty = int(w['text'])
                            elif qty is None:
                                qty = int(w['text'])
                    if tag and qty is not None and tag not in tag_qty:
                        tag_qty[tag] = qty
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    return tag_qty

def test_multi_strategy():
    import pdfplumber
    tag_qty = {}
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                lines = defaultdict(list)
                for w in words:
                    lines[round(w['top'], 0)].append(w)
                
                for y in sorted(lines.keys()):
                    line_words = sorted(lines[y], key=lambda w: w['x0'])
                    line_text = " ".join(w['text'] for w in line_words)
                    if any(h in line_text.lower() for h in QTY_HEADERS + TAG_HEADERS):
                        continue
                    tag_match = TAG_PATTERN.search(line_text)
                    if not tag_match:
                        continue
                    tag = tag_match.group(1).upper().replace("_", "-")
                    tag_idx = None
                    for i, w in enumerate(line_words):
                        if TAG_PATTERN.match(w['text']):
                            tag_idx = i
                            break
                    qty = None
                    if tag_idx and tag_idx > 0:
                        prev = line_words[tag_idx - 1]['text']
                        if prev.isdigit():
                            qty = int(prev)
                    if qty is None and tag_idx is not None and tag_idx < len(line_words) - 1:
                        next_w = line_words[tag_idx + 1]['text']
                        if next_w.isdigit():
                            qty = int(next_w)
                    if qty is None:
                        for w in line_words:
                            if w['text'].isdigit():
                                qty = int(w['text'])
                                break
                    if tag and qty is not None and tag not in tag_qty:
                        tag_qty[tag] = qty
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    return tag_qty

def test_smart_tables():
    import pdfplumber
    tag_qty = {}
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables():
                    if not table or len(table) < 2:
                        continue
                    header = table[0] if table else []
                    tag_col = qty_col = None
                    for i, cell in enumerate(header):
                        if cell:
                            cl = str(cell).lower()
                            if any(h in cl for h in TAG_HEADERS) and tag_col is None:
                                tag_col = i
                            if any(h in cl for h in QTY_HEADERS) and qty_col is None:
                                qty_col = i
                    if tag_col is None:
                        for row in table[1:5]:
                            for i, cell in enumerate(row):
                                if cell and TAG_PATTERN.match(str(cell)):
                                    tag_col = i
                                    qty_col = i - 1 if i > 0 else i + 1
                                    break
                            if tag_col is not None:
                                break
                    for row in table[1:]:
                        if not row:
                            continue
                        tag = qty = None
                        if tag_col is not None and tag_col < len(row) and row[tag_col]:
                            m = TAG_PATTERN.search(str(row[tag_col]))
                            if m:
                                tag = m.group(1).upper().replace("_", "-")
                        if not tag:
                            for cell in row:
                                if cell:
                                    m = TAG_PATTERN.search(str(cell))
                                    if m:
                                        tag = m.group(1).upper().replace("_", "-")
                                        break
                        if qty_col is not None and qty_col < len(row) and row[qty_col]:
                            if str(row[qty_col]).strip().isdigit():
                                qty = int(row[qty_col])
                        if qty is None:
                            for cell in row:
                                if cell and str(cell).strip().isdigit():
                                    qty = int(cell)
                                    break
                        if tag and qty is not None and tag not in tag_qty:
                            tag_qty[tag] = qty
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    return tag_qty

def test_pymupdf_grid():
    import fitz
    tag_qty = {}
    try:
        doc = fitz.open(PDF_PATH)
        for page in doc:
            words = page.get_text("words")
            lines = defaultdict(list)
            for w in words:
                lines[round(w[1], 0)].append(w)
            all_x = [round(w[0], -1) for w in words]
            x_counts = defaultdict(int)
            for x in all_x:
                x_counts[x] += 1
            columns = sorted([x for x, c in x_counts.items() if c >= 3])
            def get_col(wx):
                for i, cx in enumerate(columns):
                    if abs(wx - cx) < 15:
                        return i
                return None
            qty_col = header_y = None
            for y in sorted(lines.keys())[:15]:
                for w in lines[y]:
                    if any(h in w[4].lower() for h in QTY_HEADERS):
                        qty_col = get_col(w[0])
                        header_y = y
                        break
                if header_y:
                    break
            for y in sorted(lines.keys()):
                if header_y and y <= header_y:
                    continue
                line_words = sorted(lines[y], key=lambda w: w[0])
                tag = qty = None
                for w in line_words:
                    word = w[4]
                    col_idx = get_col(w[0])
                    if TAG_PATTERN.match(word):
                        tag = TAG_PATTERN.match(word).group(1).upper().replace("_", "-")
                    if word.isdigit():
                        if qty_col is not None and col_idx == qty_col:
                            qty = int(word)
                        elif qty is None:
                            qty = int(word)
                if tag and qty is not None and tag not in tag_qty:
                    tag_qty[tag] = qty
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    return tag_qty

def test_scoring():
    import pdfplumber
    tag_qty = {}
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                lines = defaultdict(list)
                for w in words:
                    lines[round(w['top'], 0)].append(w)
                for y in sorted(lines.keys()):
                    line_words = sorted(lines[y], key=lambda w: w['x0'])
                    tags_found = []
                    numbers_found = []
                    for i, w in enumerate(line_words):
                        if TAG_PATTERN.match(w['text']):
                            tags_found.append((i, w['text'], w['x0']))
                        if w['text'].isdigit():
                            numbers_found.append((i, int(w['text']), w['x0']))
                    if not tags_found or not numbers_found:
                        continue
                    for tag_idx, tag_text, tag_x in tags_found:
                        tag = TAG_PATTERN.match(tag_text).group(1).upper().replace("_", "-")
                        if tag in tag_qty:
                            continue
                        best_qty = None
                        best_score = -999
                        for num_idx, num_val, num_x in numbers_found:
                            score = 0
                            col_dist = abs(num_idx - tag_idx)
                            score -= col_dist * 10
                            if num_idx == tag_idx - 1:
                                score += 50
                            if num_idx == tag_idx + 1:
                                score += 30
                            if col_dist > 3:
                                score -= 100
                            if score > best_score:
                                best_score = score
                                best_qty = num_val
                        if best_qty is not None:
                            tag_qty[tag] = best_qty
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    return tag_qty

def test_qty_tag_pattern():
    import fitz
    tag_qty = {}
    pattern = re.compile(r'(\d+)\s+([A-Z]{2,4}\d{3,4}[-_]\d{3,4})', re.IGNORECASE)
    try:
        doc = fitz.open(PDF_PATH)
        for page in doc:
            for match in pattern.findall(page.get_text()):
                qty = int(match[0])
                tag = match[1].upper().replace("_", "-")
                if tag not in tag_qty:
                    tag_qty[tag] = qty
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    return tag_qty

def test_tag_qty_pattern():
    """Pattern: Tag suivi de Qty (format XNRGY standard)"""
    import fitz
    tag_qty = {}
    # Pattern plus flexible: Tag puis Qty avec possibles colonnes entre
    pattern = re.compile(r'([A-Z]{2,4}\d{3,4}[-_]\d{3,4})\s+(\d+)', re.IGNORECASE)
    try:
        doc = fitz.open(PDF_PATH)
        for page in doc:
            for match in pattern.findall(page.get_text()):
                tag = match[0].upper().replace("_", "-")
                qty = int(match[1])
                if tag not in tag_qty:
                    tag_qty[tag] = qty
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    return tag_qty

def test_column_based():
    """Extraction basee sur colonnes: Tag en col 1, Qty en col 2"""
    import pdfplumber
    tag_qty = {}
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                lines = defaultdict(list)
                for w in words:
                    lines[round(w['top'], 0)].append(w)
                
                for y in sorted(lines.keys()):
                    line_words = sorted(lines[y], key=lambda w: w['x0'])
                    if len(line_words) < 2:
                        continue
                    
                    # Chercher Tag dans les premiers mots
                    for i, w in enumerate(line_words):
                        if TAG_PATTERN.match(w['text']):
                            tag = TAG_PATTERN.match(w['text']).group(1).upper().replace("_", "-")
                            # Qty est le mot suivant
                            if i + 1 < len(line_words):
                                next_text = line_words[i + 1]['text']
                                if next_text.isdigit():
                                    if tag not in tag_qty:
                                        tag_qty[tag] = int(next_text)
                            break
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    return tag_qty

def main():
    print("=" * 80)
    print("PDF TABLE EXTRACTION BENCHMARK v3 - DETECTION INTELLIGENTE")
    print("=" * 80)
    print()
    
    print("[>] Chargement reference CSV...")
    reference = load_csv_reference()
    print(f"[+] Reference: {len(reference)} paires (Tag, Qty)")
    print()
    
    if not os.path.exists(PDF_PATH):
        print(f"[-] PDF non trouve: {PDF_PATH}")
        return
    
    results = []
    
    print("[>] Methode 1: Auto Structure")
    r = test_auto_structure()
    e = evaluate("Auto Structure", r, reference)
    print(f"    => {e['correct']}/{e['total']} ({e['accuracy']}%)")
    results.append(e)
    
    print("[>] Methode 2: Multi-Strategy")
    r = test_multi_strategy()
    e = evaluate("Multi-Strategy", r, reference)
    print(f"    => {e['correct']}/{e['total']} ({e['accuracy']}%)")
    results.append(e)
    
    print("[>] Methode 3: Smart Tables")
    r = test_smart_tables()
    e = evaluate("Smart Tables", r, reference)
    print(f"    => {e['correct']}/{e['total']} ({e['accuracy']}%)")
    results.append(e)
    
    print("[>] Methode 4: PyMuPDF Grid")
    r = test_pymupdf_grid()
    e = evaluate("PyMuPDF Grid", r, reference)
    print(f"    => {e['correct']}/{e['total']} ({e['accuracy']}%)")
    results.append(e)
    
    print("[>] Methode 5: Scoring")
    r = test_scoring()
    e = evaluate("Scoring", r, reference)
    print(f"    => {e['correct']}/{e['total']} ({e['accuracy']}%)")
    results.append(e)
    
    print("[>] Methode 6: Pattern Qty+Tag")
    r = test_qty_tag_pattern()
    e = evaluate("Qty+Tag", r, reference)
    print(f"    => {e['correct']}/{e['total']} ({e['accuracy']}%)")
    results.append(e)
    
    print("[>] Methode 7: Pattern Tag+Qty (XNRGY Standard)")
    r = test_tag_qty_pattern()
    e = evaluate("Tag+Qty", r, reference)
    print(f"    => {e['correct']}/{e['total']} ({e['accuracy']}%)")
    results.append(e)
    
    print("[>] Methode 8: Column Based (Tag col1, Qty col2)")
    r = test_column_based()
    e = evaluate("Column Based", r, reference)
    print(f"    => {e['correct']}/{e['total']} ({e['accuracy']}%)")
    results.append(e)
    
    print()
    print("=" * 80)
    print("RESUME COMPARATIF")
    print("=" * 80)
    print()
    
    sorted_results = sorted(results, key=lambda x: x['accuracy'], reverse=True)
    print(f"{'Rang':<5} {'Methode':<20} {'Correct':<10} {'Accuracy':<10} {'Wrong':<8} {'Missing':<8}")
    print("-" * 70)
    
    for i, r in enumerate(sorted_results, 1):
        marker = " ***" if i == 1 else ""
        print(f"{i:<5} {r['method']:<20} {r['correct']:<10} {r['accuracy']:>6}% {r['wrong']:<8} {r['missing']:<8}{marker}")
    
    print()
    best = sorted_results[0]
    print(f"[+] MEILLEURE: {best['method']} ({best['accuracy']}%)")
    
    if best['wrong_details']:
        print(f"\n    Erreurs Qty:")
        for tag, exp, got in best['wrong_details'][:5]:
            print(f"      {tag}: attendu={exp}, obtenu={got}")
    
    if best['missing_tags']:
        print(f"\n    Manquants ({best['missing']}):")
        for tag in best['missing_tags'][:5]:
            print(f"      - {tag}")
    
    print()
    print("=" * 80)

if __name__ == "__main__":
    main()
