# -*- coding: utf-8 -*-
"""
BENCHMARK MOTEURS EXTRACTION TABLEAUX PDF
Objectif: Trouver le meilleur moteur pour extraire Tag + Qty
peu importe le format/style du tableau
"""

import re
import os
from collections import defaultdict

PDF_PATH = r"C:\Vault\Engineering\Projects\10381\REF13\M02\6-Shop Drawing PDF\Production\BatchPrint\02-Machines.pdf"
CSV_PATH = r"C:\Vault\Engineering\Projects\10381\REF13\M02\5_Exportation\Sheet_Metal_Nesting\Punch\10381-13-M02.csv"

TAG_PATTERN = re.compile(r'([A-Z]{2,4}\d{3,4}[-_]\d{3,4})', re.IGNORECASE)

def load_csv_reference():
    """Charge reference depuis CSV (format: Qty,Filename.dxf,...)"""
    tag_qty = {}
    if not os.path.exists(CSV_PATH):
        return tag_qty
    with open(CSV_PATH, 'r', encoding='utf-8-sig') as f:
        for line in f:
            parts = re.split(r'[;,]', line)
            if len(parts) >= 2:
                try:
                    qty = int(parts[0].strip())
                    match = TAG_PATTERN.search(parts[1].strip())
                    if match:
                        tag = match.group(1).upper().replace("_", "-")
                        tag_qty[tag] = qty
                except:
                    continue
    return tag_qty

def evaluate(name, extracted, reference):
    """Evalue resultats"""
    total = len(reference)
    correct = sum(1 for t, q in extracted.items() if reference.get(t) == q)
    wrong = sum(1 for t, q in extracted.items() if t in reference and reference[t] != q)
    missing = [t for t in reference if t not in extracted]
    acc = correct / total * 100 if total > 0 else 0
    return {
        'name': name, 'correct': correct, 'total': total,
        'wrong': wrong, 'missing': len(missing), 'accuracy': round(acc, 1),
        'missing_list': missing[:10],
        'wrong_list': [(t, reference[t], extracted[t]) for t in extracted 
                       if t in reference and reference[t] != extracted[t]][:10]
    }

# =============================================================================
# MOTEUR 1: CAMELOT (specialise tableaux)
# =============================================================================
def test_camelot():
    """Camelot - Detection automatique de tableaux"""
    tag_qty = {}
    try:
        import camelot
        tables = camelot.read_pdf(PDF_PATH, pages='all', flavor='stream')
        
        for table in tables:
            df = table.df
            if df.empty:
                continue
            
            # Detecter colonnes Tag et Qty
            tag_col = qty_col = None
            for col in df.columns:
                first_vals = df[col].head(5).str.lower().tolist()
                if any('tag' in str(v) or 'item' in str(v) for v in first_vals):
                    tag_col = col
                if any('qty' in str(v) or 'qte' in str(v) for v in first_vals):
                    qty_col = col
            
            # Si pas trouve, chercher par contenu
            if tag_col is None:
                for col in df.columns:
                    for val in df[col]:
                        if TAG_PATTERN.match(str(val)):
                            tag_col = col
                            break
                    if tag_col:
                        break
            
            if tag_col is None:
                continue
                
            # Qty = colonne suivante si pas detectee
            cols = list(df.columns)
            if qty_col is None and tag_col in cols:
                idx = cols.index(tag_col)
                if idx + 1 < len(cols):
                    qty_col = cols[idx + 1]
            
            # Extraire
            for _, row in df.iterrows():
                tag_val = str(row.get(tag_col, ''))
                qty_val = str(row.get(qty_col, '')) if qty_col else ''
                
                match = TAG_PATTERN.search(tag_val)
                if match:
                    tag = match.group(1).upper().replace("_", "-")
                    if qty_val.strip().isdigit():
                        qty = int(qty_val)
                        if tag not in tag_qty:
                            tag_qty[tag] = qty
                            
    except Exception as e:
        print(f"    [-] Erreur Camelot: {e}")
    return tag_qty

# =============================================================================
# MOTEUR 2: TABULA
# =============================================================================
def test_tabula():
    """Tabula - Extraction tableaux via Java"""
    tag_qty = {}
    try:
        import tabula
        tables = tabula.read_pdf(PDF_PATH, pages='all', multiple_tables=True, silent=True)
        
        for df in tables:
            if df.empty:
                continue
            
            # Detecter colonnes
            tag_col = qty_col = None
            for col in df.columns:
                col_lower = str(col).lower()
                if 'tag' in col_lower or 'item' in col_lower:
                    tag_col = col
                if 'qty' in col_lower or 'qte' in col_lower:
                    qty_col = col
            
            # Fallback: chercher par contenu
            if tag_col is None:
                for col in df.columns:
                    for val in df[col].dropna():
                        if TAG_PATTERN.match(str(val)):
                            tag_col = col
                            break
                    if tag_col:
                        break
            
            if tag_col is None:
                continue
            
            # Qty = colonne suivante
            cols = list(df.columns)
            if qty_col is None and tag_col in cols:
                idx = cols.index(tag_col)
                if idx + 1 < len(cols):
                    qty_col = cols[idx + 1]
            
            for _, row in df.iterrows():
                tag_val = str(row.get(tag_col, ''))
                qty_val = str(row.get(qty_col, '')) if qty_col else ''
                
                match = TAG_PATTERN.search(tag_val)
                if match:
                    tag = match.group(1).upper().replace("_", "-")
                    try:
                        qty = int(float(qty_val))
                        if tag not in tag_qty:
                            tag_qty[tag] = qty
                    except:
                        pass
                        
    except Exception as e:
        print(f"    [-] Erreur Tabula: {e}")
    return tag_qty

# =============================================================================
# MOTEUR 3: PDFPLUMBER avec detection structure
# =============================================================================
def test_pdfplumber_tables():
    """pdfplumber - Tables avec analyse structure"""
    tag_qty = {}
    try:
        import pdfplumber
        
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    
                    # Analyser header (premiere ligne)
                    header = [str(c).lower() if c else '' for c in table[0]]
                    
                    tag_col = qty_col = None
                    for i, h in enumerate(header):
                        if 'tag' in h or 'item' in h or 'part' in h:
                            tag_col = i
                        if 'qty' in h or 'qte' in h or 'quantit' in h:
                            qty_col = i
                    
                    # Fallback: detecter par contenu
                    if tag_col is None:
                        for row in table[1:5]:
                            for i, cell in enumerate(row):
                                if cell and TAG_PATTERN.match(str(cell)):
                                    tag_col = i
                                    break
                            if tag_col is not None:
                                break
                    
                    if tag_col is None:
                        continue
                    
                    # Qty = colonne suivante ou precedente
                    if qty_col is None:
                        if tag_col + 1 < len(header):
                            qty_col = tag_col + 1
                        elif tag_col > 0:
                            qty_col = tag_col - 1
                    
                    # Extraire donnees
                    for row in table[1:]:
                        if not row or len(row) <= tag_col:
                            continue
                        
                        tag_val = str(row[tag_col]) if row[tag_col] else ''
                        qty_val = str(row[qty_col]) if qty_col and len(row) > qty_col and row[qty_col] else ''
                        
                        match = TAG_PATTERN.search(tag_val)
                        if match:
                            tag = match.group(1).upper().replace("_", "-")
                            if qty_val.strip().isdigit():
                                qty = int(qty_val)
                                if tag not in tag_qty:
                                    tag_qty[tag] = qty
                                    
    except Exception as e:
        print(f"    [-] Erreur pdfplumber: {e}")
    return tag_qty

# =============================================================================
# MOTEUR 4: PYMUPDF avec reconstruction grille
# =============================================================================
def test_pymupdf_structure():
    """PyMuPDF - Reconstruction structure tableau"""
    tag_qty = {}
    try:
        import fitz
        
        doc = fitz.open(PDF_PATH)
        for page in doc:
            words = page.get_text("words")
            if not words:
                continue
            
            # Grouper par lignes (tolerance Y = 5 pixels)
            lines = defaultdict(list)
            for w in words:
                y_key = round(w[1] / 5) * 5
                lines[y_key].append({
                    'text': w[4],
                    'x0': w[0],
                    'x1': w[2],
                    'y': w[1]
                })
            
            # Pour chaque ligne, trouver Tag et Qty le plus proche
            for y_key in sorted(lines.keys()):
                line_words = sorted(lines[y_key], key=lambda w: w['x0'])
                
                # Trouver tous les tags et tous les nombres sur cette ligne
                tags_on_line = []
                numbers_on_line = []
                
                for w in line_words:
                    match = TAG_PATTERN.match(w['text'])
                    if match:
                        tags_on_line.append({
                            'tag': match.group(1).upper().replace("_", "-"),
                            'x': w['x0']
                        })
                    if w['text'].isdigit() and len(w['text']) <= 3:
                        numbers_on_line.append({
                            'qty': int(w['text']),
                            'x': w['x0']
                        })
                
                # Pour chaque tag, trouver le nombre le plus proche (gauche OU droite)
                for tag_info in tags_on_line:
                    tag = tag_info['tag']
                    tag_x = tag_info['x']
                    
                    if tag in tag_qty:
                        continue
                    
                    # Chercher nombre le plus proche du tag (priorite: droite puis gauche)
                    best_qty = None
                    best_dist = 999999
                    
                    # D'abord chercher a DROITE (format Tag | Qty)
                    for num_info in numbers_on_line:
                        num_x = num_info['x']
                        if num_x > tag_x:  # A droite
                            dist = num_x - tag_x
                            if dist < best_dist and dist < 150:
                                best_dist = dist
                                best_qty = num_info['qty']
                    
                    # Si rien a droite, chercher a gauche (format Qty | Tag)
                    if best_qty is None:
                        for num_info in numbers_on_line:
                            num_x = num_info['x']
                            if num_x < tag_x:  # A gauche
                                dist = tag_x - num_x
                                if dist < best_dist and dist < 150:
                                    best_dist = dist
                                    best_qty = num_info['qty']
                    
                    if best_qty is not None:
                        tag_qty[tag] = best_qty
        
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur PyMuPDF: {e}")
    return tag_qty

# =============================================================================
# MOTEUR 5: HYBRIDE (meilleur de chaque)
# =============================================================================
def test_hybrid():
    """Hybride - Combine plusieurs moteurs avec vote"""
    results = {}
    
    # Collecter de tous les moteurs
    for name, func in [('camelot', test_camelot), 
                       ('tabula', test_tabula),
                       ('pdfplumber', test_pdfplumber_tables),
                       ('pymupdf', test_pymupdf_structure)]:
        try:
            data = func()
            for tag, qty in data.items():
                if tag not in results:
                    results[tag] = {}
                results[tag][name] = qty
        except:
            pass
    
    # Vote majoritaire
    tag_qty = {}
    for tag, votes in results.items():
        if not votes:
            continue
        # Prendre la quantite la plus frequente
        qty_counts = defaultdict(int)
        for qty in votes.values():
            qty_counts[qty] += 1
        best_qty = max(qty_counts.keys(), key=lambda q: qty_counts[q])
        tag_qty[tag] = best_qty
    
    return tag_qty

# =============================================================================
# MAIN
# =============================================================================
def main():
    print("=" * 80)
    print("BENCHMARK MOTEURS EXTRACTION TABLEAUX PDF")
    print("=" * 80)
    print()
    print(f"PDF: {os.path.basename(PDF_PATH)}")
    print(f"CSV: {os.path.basename(CSV_PATH)}")
    print()
    
    print("[>] Chargement reference CSV...")
    reference = load_csv_reference()
    print(f"[+] Reference: {len(reference)} paires (Tag, Qty)")
    print()
    
    if not os.path.exists(PDF_PATH):
        print(f"[-] PDF non trouve!")
        return
    
    results = []
    
    print("[>] Test CAMELOT (specialise tableaux)...")
    r = test_camelot()
    e = evaluate("Camelot", r, reference)
    print(f"    => {e['correct']}/{e['total']} ({e['accuracy']}%)")
    results.append(e)
    
    print("[>] Test TABULA (Java-based)...")
    r = test_tabula()
    e = evaluate("Tabula", r, reference)
    print(f"    => {e['correct']}/{e['total']} ({e['accuracy']}%)")
    results.append(e)
    
    print("[>] Test PDFPLUMBER Tables...")
    r = test_pdfplumber_tables()
    e = evaluate("pdfplumber", r, reference)
    print(f"    => {e['correct']}/{e['total']} ({e['accuracy']}%)")
    results.append(e)
    
    print("[>] Test PYMUPDF Structure...")
    r = test_pymupdf_structure()
    e = evaluate("PyMuPDF", r, reference)
    print(f"    => {e['correct']}/{e['total']} ({e['accuracy']}%)")
    results.append(e)
    
    print("[>] Test HYBRIDE (vote)...")
    r = test_hybrid()
    e = evaluate("Hybride", r, reference)
    print(f"    => {e['correct']}/{e['total']} ({e['accuracy']}%)")
    results.append(e)
    
    print()
    print("=" * 80)
    print("RESULTATS")
    print("=" * 80)
    print()
    
    sorted_results = sorted(results, key=lambda x: x['accuracy'], reverse=True)
    print(f"{'Rang':<5} {'Moteur':<15} {'Correct':<10} {'Accuracy':<10} {'Wrong':<8} {'Missing':<8}")
    print("-" * 65)
    
    for i, r in enumerate(sorted_results, 1):
        marker = " ***" if i == 1 else ""
        print(f"{i:<5} {r['name']:<15} {r['correct']:<10} {r['accuracy']:>6}% {r['wrong']:<8} {r['missing']:<8}{marker}")
    
    # Details du meilleur
    best = sorted_results[0]
    print()
    print(f"[+] MEILLEUR: {best['name']} ({best['accuracy']}%)")
    
    if best['missing_list']:
        print(f"\n    Tags manquants ({best['missing']}):")
        for tag in best['missing_list'][:5]:
            print(f"      - {tag}")
    
    if best['wrong_list']:
        print(f"\n    Erreurs Qty:")
        for tag, exp, got in best['wrong_list'][:5]:
            print(f"      {tag}: attendu={exp}, obtenu={got}")
    
    print()
    print("=" * 80)

if __name__ == "__main__":
    main()
