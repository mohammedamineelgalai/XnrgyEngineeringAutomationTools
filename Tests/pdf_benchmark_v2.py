"""
=============================================================================
PDF TABLE EXTRACTION BENCHMARK v2 - TAG + QUANTITÉ
=============================================================================
Le VRAI défi: Extraire correctement les PAIRES (Tag, Quantité) depuis les tableaux
Référence: CSV contient Tag + Quantité
=============================================================================
"""

import re
import os
from collections import defaultdict
from typing import Dict, Set, List, Tuple

# Fichiers de reference
PDF_PATH = r"C:\Vault\Engineering\Projects\10381\REF13\M02\6-Shop Drawing PDF\Production\BatchPrint\02-Machines.pdf"
CSV_PATH = r"C:\Vault\Engineering\Projects\10381\REF13\M02\5_Exportation\Sheet_Metal_Nesting\Punch\10381-13-M02.csv"

# Pattern pour les tags XNRGY
TAG_PATTERN = re.compile(r'([A-Z]{2,4}\d{3,4}[-_]\d{3,4})', re.IGNORECASE)

def load_csv_reference() -> Dict[str, int]:
    """Charge les paires (Tag, Quantité) depuis le CSV"""
    tag_qty = {}
    
    if not os.path.exists(CSV_PATH):
        print(f"[-] ERREUR: CSV non trouve: {CSV_PATH}")
        return tag_qty
    
    with open(CSV_PATH, 'r', encoding='utf-8-sig') as f:
        lines = f.readlines()
        
    for line in lines:
        parts = re.split(r'[;,]', line)
        # Format CSV: Qty, Filename.dxf, Material, ...
        if len(parts) >= 2:
            try:
                qty = int(parts[0].strip())
            except:
                continue
            
            filename = parts[1].strip()
            match = TAG_PATTERN.search(filename)
            if match:
                tag = match.group(1).upper().replace("_", "-")
                tag_qty[tag] = qty
    
    return tag_qty

def evaluate_tag_qty_result(method_name: str, extracted: Dict[str, int], reference: Dict[str, int]) -> Dict:
    """Évalue les résultats d'extraction TAG + QUANTITÉ"""
    
    total_ref = len(reference)
    total_extracted = len(extracted)
    
    # Tags corrects avec bonne quantité
    correct_tag_qty = 0
    correct_tag_wrong_qty = 0
    missing_tags = []
    wrong_qty_details = []
    
    for tag, ref_qty in reference.items():
        if tag in extracted:
            if extracted[tag] == ref_qty:
                correct_tag_qty += 1
            else:
                correct_tag_wrong_qty += 1
                wrong_qty_details.append({
                    'tag': tag,
                    'expected': ref_qty,
                    'got': extracted[tag]
                })
        else:
            missing_tags.append(tag)
    
    # Tags en trop (pas dans la référence)
    extra_tags = [t for t in extracted if t not in reference]
    
    # Métriques
    tag_recall = (len(extracted) - len(extra_tags)) / total_ref * 100 if total_ref > 0 else 0
    qty_accuracy = correct_tag_qty / total_ref * 100 if total_ref > 0 else 0
    
    return {
        'method_name': method_name,
        'total_reference': total_ref,
        'total_extracted': total_extracted,
        'correct_tag_qty': correct_tag_qty,
        'correct_tag_wrong_qty': correct_tag_wrong_qty,
        'missing_tags': missing_tags,
        'extra_tags': extra_tags,
        'wrong_qty_details': wrong_qty_details,
        'tag_recall': round(tag_recall, 1),
        'qty_accuracy': round(qty_accuracy, 1)
    }

def print_result(result: Dict):
    """Affiche le résultat"""
    print(f"    Tags trouvés: {result['total_extracted']}/{result['total_reference']}")
    print(f"    Tag+Qty corrects: {result['correct_tag_qty']} ({result['qty_accuracy']}%)")
    print(f"    Tag OK, Qty FAUX: {result['correct_tag_wrong_qty']}")
    print(f"    Tags manquants: {len(result['missing_tags'])}")
    print(f"    Tags en trop: {len(result['extra_tags'])}")
    print()

# =============================================================================
# MÉTHODE 1: PyMuPDF - Extraction par lignes (position Y)
# =============================================================================
def test_pymupdf_lines() -> Dict[str, int]:
    """PyMuPDF - Reconstruction des lignes par position Y puis extraction Tag+Qty"""
    import fitz
    
    tag_qty = {}
    
    try:
        doc = fitz.open(PDF_PATH)
        
        for page in doc:
            # Récupérer tous les mots avec leurs positions
            words = page.get_text("words")
            # Format: (x0, y0, x1, y1, "word", block_no, line_no, word_no)
            
            # Grouper par ligne (position Y)
            lines = defaultdict(list)
            for w in words:
                y = round(w[1], 0)  # y0 arrondi
                lines[y].append(w)
            
            # Pour chaque ligne, chercher Tag + Quantité
            for y in sorted(lines.keys()):
                line_words = sorted(lines[y], key=lambda x: x[0])  # Trier par X
                line_text = " ".join(w[4] for w in line_words)
                
                # Chercher un tag
                tag_match = TAG_PATTERN.search(line_text)
                if tag_match:
                    tag = tag_match.group(1).upper().replace("_", "-")
                    
                    # Chercher la quantité (premier nombre dans la ligne)
                    # Stratégie: Le premier nombre AVANT ou APRÈS le tag
                    words_text = [w[4] for w in line_words]
                    
                    qty = None
                    for word in words_text:
                        if word.isdigit():
                            qty = int(word)
                            break
                    
                    if qty is not None and tag not in tag_qty:
                        tag_qty[tag] = qty
        
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tag_qty

# =============================================================================
# MÉTHODE 2: PyMuPDF - Dict structure avec spans
# =============================================================================
def test_pymupdf_dict_spans() -> Dict[str, int]:
    """PyMuPDF - Utilise la structure dict pour reconstruire les lignes"""
    import fitz
    
    tag_qty = {}
    
    try:
        doc = fitz.open(PDF_PATH)
        
        for page in doc:
            text_dict = page.get_text("dict")
            
            for block in text_dict.get("blocks", []):
                for line in block.get("lines", []):
                    # Reconstruire le texte de la ligne
                    line_text = ""
                    for span in line.get("spans", []):
                        line_text += span.get("text", "") + " "
                    
                    line_text = line_text.strip()
                    
                    # Chercher Tag + Qty
                    tag_match = TAG_PATTERN.search(line_text)
                    if tag_match:
                        tag = tag_match.group(1).upper().replace("_", "-")
                        
                        # Chercher nombre
                        numbers = re.findall(r'\b(\d+)\b', line_text)
                        if numbers:
                            qty = int(numbers[0])
                            if tag not in tag_qty:
                                tag_qty[tag] = qty
        
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tag_qty

# =============================================================================
# MÉTHODE 3: pdfplumber - Tables natives
# =============================================================================
def test_pdfplumber_tables() -> Dict[str, int]:
    """pdfplumber - Extraction des tableaux natifs"""
    import pdfplumber
    
    tag_qty = {}
    
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                
                for table in tables:
                    for row in table:
                        if not row:
                            continue
                        
                        # Chercher un tag dans la ligne
                        row_text = " ".join(str(cell) if cell else "" for cell in row)
                        tag_match = TAG_PATTERN.search(row_text)
                        
                        if tag_match:
                            tag = tag_match.group(1).upper().replace("_", "-")
                            
                            # Chercher la quantité dans les cellules
                            for cell in row:
                                if cell and str(cell).strip().isdigit():
                                    qty = int(cell)
                                    if tag not in tag_qty:
                                        tag_qty[tag] = qty
                                    break
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tag_qty

# =============================================================================
# MÉTHODE 4: pdfplumber - Tables avec settings text
# =============================================================================
def test_pdfplumber_tables_text_strategy() -> Dict[str, int]:
    """pdfplumber - Tables avec stratégie text"""
    import pdfplumber
    
    tag_qty = {}
    
    table_settings = {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "snap_tolerance": 5,
        "join_tolerance": 5,
    }
    
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables(table_settings)
                
                for table in tables:
                    for row in table:
                        if not row:
                            continue
                        
                        row_text = " ".join(str(cell) if cell else "" for cell in row)
                        tag_match = TAG_PATTERN.search(row_text)
                        
                        if tag_match:
                            tag = tag_match.group(1).upper().replace("_", "-")
                            
                            for cell in row:
                                if cell and str(cell).strip().isdigit():
                                    qty = int(cell)
                                    if tag not in tag_qty:
                                        tag_qty[tag] = qty
                                    break
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tag_qty

# =============================================================================
# MÉTHODE 5: pdfplumber - Lignes par position Y
# =============================================================================
def test_pdfplumber_y_clustering() -> Dict[str, int]:
    """pdfplumber - Clustering par position Y"""
    import pdfplumber
    
    tag_qty = {}
    
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                
                # Grouper par Y
                lines = defaultdict(list)
                for word in words:
                    y = round(word['top'], 0)
                    lines[y].append(word)
                
                for y in sorted(lines.keys()):
                    line_words = sorted(lines[y], key=lambda w: w['x0'])
                    line_text = " ".join(w['text'] for w in line_words)
                    
                    tag_match = TAG_PATTERN.search(line_text)
                    if tag_match:
                        tag = tag_match.group(1).upper().replace("_", "-")
                        
                        # Premier nombre
                        for w in line_words:
                            if w['text'].isdigit():
                                qty = int(w['text'])
                                if tag not in tag_qty:
                                    tag_qty[tag] = qty
                                break
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tag_qty

# =============================================================================
# MÉTHODE 6: pdfplumber - Détection structure colonnes
# =============================================================================
def test_pdfplumber_column_detection() -> Dict[str, int]:
    """pdfplumber - Détection des colonnes puis extraction"""
    import pdfplumber
    
    tag_qty = {}
    
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                
                # Trouver les colonnes (positions X fréquentes)
                x_positions = [round(w['x0'], -1) for w in words]
                x_counts = defaultdict(int)
                for x in x_positions:
                    x_counts[x] += 1
                
                # Colonnes = positions X avec beaucoup de mots
                columns = sorted([x for x, count in x_counts.items() if count > 5])
                
                # Grouper les mots par ligne
                lines = defaultdict(list)
                for word in words:
                    y = round(word['top'], 0)
                    lines[y].append(word)
                
                # Pour chaque ligne
                for y in sorted(lines.keys()):
                    line_words = sorted(lines[y], key=lambda w: w['x0'])
                    
                    # Trouver le tag et sa colonne
                    tag_word = None
                    tag_col_idx = None
                    
                    for w in line_words:
                        if TAG_PATTERN.match(w['text']):
                            tag_word = w
                            # Trouver l'index de colonne
                            for i, col_x in enumerate(columns):
                                if abs(w['x0'] - col_x) < 20:
                                    tag_col_idx = i
                                    break
                            break
                    
                    if tag_word:
                        tag = TAG_PATTERN.match(tag_word['text']).group(1).upper().replace("_", "-")
                        
                        # Chercher le nombre dans les autres colonnes
                        for w in line_words:
                            if w['text'].isdigit():
                                qty = int(w['text'])
                                if tag not in tag_qty:
                                    tag_qty[tag] = qty
                                break
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tag_qty

# =============================================================================
# MÉTHODE 7: PyMuPDF - Analyse structure de tableau
# =============================================================================
def test_pymupdf_table_structure() -> Dict[str, int]:
    """PyMuPDF - Analyse de la structure de tableau avec détection de header"""
    import fitz
    
    tag_qty = {}
    
    try:
        doc = fitz.open(PDF_PATH)
        
        for page in doc:
            words = page.get_text("words")
            
            # Identifier la colonne "Qty" ou nombre entête
            qty_column_x = None
            tag_column_x = None
            
            # Grouper par lignes
            lines = defaultdict(list)
            for w in words:
                y = round(w[1], 0)
                lines[y].append(w)
            
            # Chercher l'entête (ligne avec "QTY", "QUANTITY", "Item" etc.)
            for y in sorted(lines.keys())[:10]:  # Premières lignes
                line_words = lines[y]
                for w in line_words:
                    word_lower = w[4].lower()
                    if word_lower in ['qty', 'quantity', 'qte', 'quantite']:
                        qty_column_x = w[0]  # x0
                    if word_lower in ['item', 'tag', 'part', 'piece']:
                        tag_column_x = w[0]
            
            # Extraire les données
            for y in sorted(lines.keys()):
                line_words = sorted(lines[y], key=lambda x: x[0])
                
                tag = None
                qty = None
                
                for w in line_words:
                    word = w[4]
                    x = w[0]
                    
                    # Tag
                    if TAG_PATTERN.match(word):
                        tag = TAG_PATTERN.match(word).group(1).upper().replace("_", "-")
                    
                    # Quantité - soit dans la colonne Qty, soit premier nombre
                    if word.isdigit():
                        if qty_column_x and abs(x - qty_column_x) < 30:
                            qty = int(word)
                        elif qty is None:
                            qty = int(word)
                
                if tag and qty is not None and tag not in tag_qty:
                    tag_qty[tag] = qty
        
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tag_qty

# =============================================================================
# MÉTHODE 8: Pattern "Tag suivi de Qty"
# =============================================================================
def test_tag_qty_pattern() -> Dict[str, int]:
    """Extraction directe avec pattern: Tag suivi de nombre"""
    import fitz
    
    tag_qty = {}
    
    # Pattern: Tag puis espace(s) puis nombre
    pattern = re.compile(r'([A-Z]{2,4}\d{3,4}[-_]\d{3,4})\s+(\d+)', re.IGNORECASE)
    
    try:
        doc = fitz.open(PDF_PATH)
        
        for page in doc:
            text = page.get_text()
            
            matches = pattern.findall(text)
            for match in matches:
                tag = match[0].upper().replace("_", "-")
                qty = int(match[1])
                if tag not in tag_qty:
                    tag_qty[tag] = qty
        
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tag_qty

# =============================================================================
# MÉTHODE 9: Pattern "Qty suivi de Tag"
# =============================================================================
def test_qty_tag_pattern() -> Dict[str, int]:
    """Extraction directe avec pattern: Nombre suivi de Tag"""
    import fitz
    
    tag_qty = {}
    
    # Pattern: Nombre puis espace(s) puis Tag
    pattern = re.compile(r'(\d+)\s+([A-Z]{2,4}\d{3,4}[-_]\d{3,4})', re.IGNORECASE)
    
    try:
        doc = fitz.open(PDF_PATH)
        
        for page in doc:
            text = page.get_text()
            
            matches = pattern.findall(text)
            for match in matches:
                qty = int(match[0])
                tag = match[1].upper().replace("_", "-")
                if tag not in tag_qty:
                    tag_qty[tag] = qty
        
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tag_qty

# =============================================================================
# MÉTHODE 10: Combinée - Meilleur des deux patterns
# =============================================================================
def test_combined_patterns() -> Dict[str, int]:
    """Combine les deux patterns Tag+Qty et Qty+Tag"""
    import fitz
    
    tag_qty = {}
    
    pattern1 = re.compile(r'([A-Z]{2,4}\d{3,4}[-_]\d{3,4})\s+(\d+)', re.IGNORECASE)
    pattern2 = re.compile(r'(\d+)\s+([A-Z]{2,4}\d{3,4}[-_]\d{3,4})', re.IGNORECASE)
    
    try:
        doc = fitz.open(PDF_PATH)
        
        for page in doc:
            text = page.get_text()
            
            # Pattern 1: Tag puis Qty
            for match in pattern1.findall(text):
                tag = match[0].upper().replace("_", "-")
                qty = int(match[1])
                if tag not in tag_qty:
                    tag_qty[tag] = qty
            
            # Pattern 2: Qty puis Tag
            for match in pattern2.findall(text):
                qty = int(match[0])
                tag = match[1].upper().replace("_", "-")
                if tag not in tag_qty:
                    tag_qty[tag] = qty
        
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tag_qty

# =============================================================================
# MAIN
# =============================================================================
def main():
    print("=" * 80)
    print("PDF TABLE EXTRACTION BENCHMARK v2 - TAG + QUANTITÉ")
    print("Le VRAI défi: Extraire correctement les PAIRES (Tag, Quantité)")
    print("=" * 80)
    print()
    
    # 1. Charger la référence CSV
    print("[>] Chargement de la référence CSV (Tag + Quantité)...")
    reference = load_csv_reference()
    print(f"[+] Référence CSV: {len(reference)} paires (Tag, Quantité)")
    
    # Afficher quelques exemples
    print("\n    Exemples de la référence:")
    for i, (tag, qty) in enumerate(list(reference.items())[:5]):
        print(f"      {tag}: {qty}")
    print()
    
    if not os.path.exists(PDF_PATH):
        print(f"[-] PDF non trouvé: {PDF_PATH}")
        return
    
    # 2. Exécuter les tests
    results = []
    
    print("[>] Test: Méthode 1 - PyMuPDF Lignes par Y")
    tags1 = test_pymupdf_lines()
    result1 = evaluate_tag_qty_result("PyMuPDF - Lignes Y", tags1, reference)
    print_result(result1)
    results.append(result1)
    
    print("[>] Test: Méthode 2 - PyMuPDF Dict Spans")
    tags2 = test_pymupdf_dict_spans()
    result2 = evaluate_tag_qty_result("PyMuPDF - Dict Spans", tags2, reference)
    print_result(result2)
    results.append(result2)
    
    print("[>] Test: Méthode 3 - pdfplumber Tables")
    tags3 = test_pdfplumber_tables()
    result3 = evaluate_tag_qty_result("pdfplumber - Tables", tags3, reference)
    print_result(result3)
    results.append(result3)
    
    print("[>] Test: Méthode 4 - pdfplumber Tables Text")
    tags4 = test_pdfplumber_tables_text_strategy()
    result4 = evaluate_tag_qty_result("pdfplumber - Tables Text", tags4, reference)
    print_result(result4)
    results.append(result4)
    
    print("[>] Test: Méthode 5 - pdfplumber Y Clustering")
    tags5 = test_pdfplumber_y_clustering()
    result5 = evaluate_tag_qty_result("pdfplumber - Y Cluster", tags5, reference)
    print_result(result5)
    results.append(result5)
    
    print("[>] Test: Méthode 6 - pdfplumber Column Detection")
    tags6 = test_pdfplumber_column_detection()
    result6 = evaluate_tag_qty_result("pdfplumber - Columns", tags6, reference)
    print_result(result6)
    results.append(result6)
    
    print("[>] Test: Méthode 7 - PyMuPDF Table Structure")
    tags7 = test_pymupdf_table_structure()
    result7 = evaluate_tag_qty_result("PyMuPDF - Table Struct", tags7, reference)
    print_result(result7)
    results.append(result7)
    
    print("[>] Test: Méthode 8 - Pattern Tag+Qty")
    tags8 = test_tag_qty_pattern()
    result8 = evaluate_tag_qty_result("Pattern Tag+Qty", tags8, reference)
    print_result(result8)
    results.append(result8)
    
    print("[>] Test: Méthode 9 - Pattern Qty+Tag")
    tags9 = test_qty_tag_pattern()
    result9 = evaluate_tag_qty_result("Pattern Qty+Tag", tags9, reference)
    print_result(result9)
    results.append(result9)
    
    print("[>] Test: Méthode 10 - Patterns Combinés")
    tags10 = test_combined_patterns()
    result10 = evaluate_tag_qty_result("Patterns Combinés", tags10, reference)
    print_result(result10)
    results.append(result10)
    
    # 3. Résumé comparatif
    print("=" * 80)
    print("RÉSUMÉ COMPARATIF - EXTRACTION TAG + QUANTITÉ")
    print("=" * 80)
    print()
    print(f"Référence: {len(reference)} paires (Tag, Quantité)")
    print()
    
    # Trier par accuracy Qty
    sorted_results = sorted(results, key=lambda x: x['qty_accuracy'], reverse=True)
    
    print(f"{'Rang':<5} {'Méthode':<25} {'Tag+Qty OK':<12} {'Accuracy':<10} {'Tag OK Qty FAUX':<15} {'Manquants':<10}")
    print("-" * 90)
    
    for i, result in enumerate(sorted_results, 1):
        marker = "***" if i == 1 else "   "
        print(f"{i:<5} {result['method_name']:<25} {result['correct_tag_qty']:<12} "
              f"{result['qty_accuracy']:>6}% {result['correct_tag_wrong_qty']:<15} {len(result['missing_tags']):<10} {marker}")
    
    print()
    best = sorted_results[0]
    print(f"[+] MEILLEURE MÉTHODE: {best['method_name']}")
    print(f"    Tag+Qty corrects: {best['correct_tag_qty']}/{best['total_reference']} ({best['qty_accuracy']}%)")
    print(f"    Tag OK mais Qty FAUX: {best['correct_tag_wrong_qty']}")
    print()
    
    # 4. Analyse des erreurs de quantité
    if best['wrong_qty_details']:
        print("=" * 80)
        print(f"ERREURS DE QUANTITÉ ({len(best['wrong_qty_details'])} tags)")
        print("=" * 80)
        print()
        for detail in best['wrong_qty_details'][:20]:  # Max 20
            print(f"  {detail['tag']}: Attendu={detail['expected']}, Obtenu={detail['got']}")
        if len(best['wrong_qty_details']) > 20:
            print(f"  ... et {len(best['wrong_qty_details']) - 20} autres erreurs")
    
    # 5. Tags manquants
    if best['missing_tags']:
        print()
        print("=" * 80)
        print(f"TAGS MANQUANTS ({len(best['missing_tags'])})")
        print("=" * 80)
        for tag in best['missing_tags'][:20]:
            print(f"  - {tag}")
    
    print()
    print("=" * 80)
    print("FIN DU BENCHMARK")
    print("=" * 80)

if __name__ == "__main__":
    main()
