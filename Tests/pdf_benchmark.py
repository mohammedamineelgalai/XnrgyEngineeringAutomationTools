"""
=============================================================================
PDF TABLE EXTRACTION BENCHMARK - XNRGY DXF Verifier
=============================================================================
Objectif: Tester plusieurs moteurs/methodes d'extraction de tableaux PDF
Reference: CSV = verite terrain
PDF: 02-Machines.pdf
=============================================================================
"""

import re
import os
from collections import defaultdict
from typing import Dict, Set, List, Tuple

# Fichiers de reference
PDF_PATH = r"C:\Vault\Engineering\Projects\10381\REF13\M02\6-Shop Drawing PDF\Production\BatchPrint\02-Machines.pdf"
CSV_PATH = r"C:\Vault\Engineering\Projects\10381\REF13\M02\5_Exportation\Sheet_Metal_Nesting\Punch\10381-13-M02.csv"

# Pattern pour les tags XNRGY (2-4 lettres + 3-4 chiffres + tiret/underscore + 3-4 chiffres)
TAG_PATTERN = re.compile(r'([A-Z]{2,4}\d{3,4}[-_]\d{3,4})', re.IGNORECASE)

def load_csv_reference() -> Set[str]:
    """Charge les tags de reference depuis le CSV"""
    tags = set()
    
    if not os.path.exists(CSV_PATH):
        print(f"[-] ERREUR: CSV non trouve: {CSV_PATH}")
        return tags
    
    with open(CSV_PATH, 'r', encoding='utf-8-sig') as f:
        lines = f.readlines()  # Pas de header dans ce CSV
        
    for line in lines:
        parts = re.split(r'[;,]', line)
        # Format: Qty, Filename.dxf, Material, ...
        # Le tag est dans le nom de fichier (colonne 2)
        if len(parts) >= 2:
            filename = parts[1].strip()
            # Extraire le tag du nom de fichier (sans .dxf)
            match = TAG_PATTERN.search(filename)
            if match:
                tag = match.group(1).upper().replace("_", "-")
                tags.add(tag)
    
    return tags

def evaluate_result(method_name: str, extracted_tags: Set[str], reference_tags: Set[str]) -> Dict:
    """Evalue les resultats d'une methode"""
    
    # True Positives (tags corrects)
    correct_tags = extracted_tags & reference_tags
    true_positives = len(correct_tags)
    
    # False Negatives (tags manquants)
    missing_tags = reference_tags - extracted_tags
    false_negatives = len(missing_tags)
    
    # False Positives (tags en trop)
    extra_tags = extracted_tags - reference_tags
    false_positives = len(extra_tags)
    
    total_extracted = len(extracted_tags)
    total_reference = len(reference_tags)
    
    # Metriques
    precision = (true_positives / total_extracted * 100) if total_extracted > 0 else 0
    recall = (true_positives / total_reference * 100) if total_reference > 0 else 0
    f1 = (2 * precision * recall / (precision + recall)) if (precision + recall) > 0 else 0
    
    return {
        'method_name': method_name,
        'total_extracted': total_extracted,
        'true_positives': true_positives,
        'false_negatives': false_negatives,
        'false_positives': false_positives,
        'precision': round(precision, 1),
        'recall': round(recall, 1),
        'f1_score': round(f1, 1),
        'missing_tags': sorted(missing_tags),
        'extra_tags': sorted(extra_tags)
    }

def print_result(result: Dict):
    """Affiche le resultat d'une methode"""
    print(f"    Extraits: {result['total_extracted']} | Corrects: {result['true_positives']} | "
          f"Manquants: {result['false_negatives']} | En trop: {result['false_positives']}")
    print(f"    Precision: {result['precision']}% | Recall: {result['recall']}% | F1: {result['f1_score']}%")
    print()

# =============================================================================
# METHODE 1: PyMuPDF (fitz) - Texte brut
# =============================================================================
def test_pymupdf_raw_text() -> Set[str]:
    """Extraction avec PyMuPDF - texte brut simple"""
    import fitz
    
    tags = set()
    try:
        doc = fitz.open(PDF_PATH)
        for page in doc:
            text = page.get_text()
            matches = TAG_PATTERN.findall(text)
            for match in matches:
                tag = match.upper().replace("_", "-")
                tags.add(tag)
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tags

# =============================================================================
# METHODE 2: PyMuPDF - Extraction par blocs
# =============================================================================
def test_pymupdf_blocks() -> Set[str]:
    """Extraction avec PyMuPDF - blocs de texte"""
    import fitz
    
    tags = set()
    try:
        doc = fitz.open(PDF_PATH)
        for page in doc:
            blocks = page.get_text("blocks")
            for block in blocks:
                if len(block) >= 5:  # Bloc de texte
                    text = block[4]
                    matches = TAG_PATTERN.findall(text)
                    for match in matches:
                        tag = match.upper().replace("_", "-")
                        tags.add(tag)
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tags

# =============================================================================
# METHODE 3: PyMuPDF - Extraction par mots
# =============================================================================
def test_pymupdf_words() -> Set[str]:
    """Extraction avec PyMuPDF - mots individuels"""
    import fitz
    
    tags = set()
    try:
        doc = fitz.open(PDF_PATH)
        for page in doc:
            words = page.get_text("words")
            for word_info in words:
                word = word_info[4]  # Le texte est a l'index 4
                if TAG_PATTERN.match(word):
                    tag = TAG_PATTERN.match(word).group(1).upper().replace("_", "-")
                    tags.add(tag)
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tags

# =============================================================================
# METHODE 4: PyMuPDF - Dictionnaire structure
# =============================================================================
def test_pymupdf_dict() -> Set[str]:
    """Extraction avec PyMuPDF - dictionnaire structure"""
    import fitz
    
    tags = set()
    try:
        doc = fitz.open(PDF_PATH)
        for page in doc:
            text_dict = page.get_text("dict")
            for block in text_dict.get("blocks", []):
                for line in block.get("lines", []):
                    line_text = ""
                    for span in line.get("spans", []):
                        line_text += span.get("text", "")
                    
                    matches = TAG_PATTERN.findall(line_text)
                    for match in matches:
                        tag = match.upper().replace("_", "-")
                        tags.add(tag)
        doc.close()
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tags

# =============================================================================
# METHODE 5: pdfplumber - Texte brut
# =============================================================================
def test_pdfplumber_text() -> Set[str]:
    """Extraction avec pdfplumber - texte brut"""
    import pdfplumber
    
    tags = set()
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                matches = TAG_PATTERN.findall(text)
                for match in matches:
                    tag = match.upper().replace("_", "-")
                    tags.add(tag)
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tags

# =============================================================================
# METHODE 6: pdfplumber - Extraction de tableaux
# =============================================================================
def test_pdfplumber_tables() -> Set[str]:
    """Extraction avec pdfplumber - detection de tableaux"""
    import pdfplumber
    
    tags = set()
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell:
                                matches = TAG_PATTERN.findall(str(cell))
                                for match in matches:
                                    tag = match.upper().replace("_", "-")
                                    tags.add(tag)
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tags

# =============================================================================
# METHODE 7: pdfplumber - Mots individuels
# =============================================================================
def test_pdfplumber_words() -> Set[str]:
    """Extraction avec pdfplumber - mots individuels"""
    import pdfplumber
    
    tags = set()
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                for word_info in words:
                    word = word_info.get('text', '')
                    if TAG_PATTERN.match(word):
                        tag = TAG_PATTERN.match(word).group(1).upper().replace("_", "-")
                        tags.add(tag)
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tags

# =============================================================================
# METHODE 8: pdfplumber - Lignes reconstruites par Y
# =============================================================================
def test_pdfplumber_lines_by_y() -> Set[str]:
    """Extraction avec pdfplumber - reconstruction des lignes par position Y"""
    import pdfplumber
    
    tags = set()
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                chars = page.chars
                if not chars:
                    continue
                
                # Grouper par position Y (arrondie)
                lines_dict = defaultdict(list)
                for char in chars:
                    y = round(char['top'], 0)
                    lines_dict[y].append(char)
                
                # Reconstruire les lignes
                for y in sorted(lines_dict.keys()):
                    chars_in_line = sorted(lines_dict[y], key=lambda c: c['x0'])
                    line_text = ''.join(c['text'] for c in chars_in_line)
                    
                    matches = TAG_PATTERN.findall(line_text)
                    for match in matches:
                        tag = match.upper().replace("_", "-")
                        tags.add(tag)
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tags

# =============================================================================
# METHODE 9: pdfplumber - Tables avec settings optimises
# =============================================================================
def test_pdfplumber_tables_optimized() -> Set[str]:
    """Extraction avec pdfplumber - tableaux avec parametres optimises"""
    import pdfplumber
    
    tags = set()
    try:
        table_settings = {
            "vertical_strategy": "text",
            "horizontal_strategy": "text",
            "snap_tolerance": 5,
            "join_tolerance": 5,
            "edge_min_length": 10,
            "min_words_vertical": 2,
            "min_words_horizontal": 2,
        }
        
        with pdfplumber.open(PDF_PATH) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables(table_settings)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell:
                                matches = TAG_PATTERN.findall(str(cell))
                                for match in matches:
                                    tag = match.upper().replace("_", "-")
                                    tags.add(tag)
                
                # Aussi extraire le texte hors tableaux
                text = page.extract_text() or ""
                matches = TAG_PATTERN.findall(text)
                for match in matches:
                    tag = match.upper().replace("_", "-")
                    tags.add(tag)
    except Exception as e:
        print(f"    [-] Erreur: {e}")
    
    return tags

# =============================================================================
# METHODE 10: Combinee (toutes les methodes)
# =============================================================================
def test_combined(*tag_sets) -> Set[str]:
    """Combine toutes les extractions"""
    combined = set()
    for tags in tag_sets:
        combined.update(tags)
    return combined

# =============================================================================
# MAIN
# =============================================================================
def main():
    print("=" * 80)
    print("PDF TABLE EXTRACTION BENCHMARK - XNRGY DXF Verifier")
    print("=" * 80)
    print()
    
    # 1. Charger la reference CSV
    print("[>] Chargement de la reference CSV...")
    reference_tags = load_csv_reference()
    print(f"[+] Reference CSV: {len(reference_tags)} tags uniques")
    print()
    
    if not os.path.exists(PDF_PATH):
        print(f"[-] PDF non trouve: {PDF_PATH}")
        return
    
    print(f"[>] PDF: {PDF_PATH}")
    print()
    
    # 2. Executer les tests
    results = []
    
    # PyMuPDF tests
    print("[>] Test: Methode 1 - PyMuPDF Texte brut")
    tags1 = test_pymupdf_raw_text()
    result1 = evaluate_result("PyMuPDF - Texte brut", tags1, reference_tags)
    print_result(result1)
    results.append(result1)
    
    print("[>] Test: Methode 2 - PyMuPDF Blocs")
    tags2 = test_pymupdf_blocks()
    result2 = evaluate_result("PyMuPDF - Blocs", tags2, reference_tags)
    print_result(result2)
    results.append(result2)
    
    print("[>] Test: Methode 3 - PyMuPDF Mots")
    tags3 = test_pymupdf_words()
    result3 = evaluate_result("PyMuPDF - Mots", tags3, reference_tags)
    print_result(result3)
    results.append(result3)
    
    print("[>] Test: Methode 4 - PyMuPDF Dict structure")
    tags4 = test_pymupdf_dict()
    result4 = evaluate_result("PyMuPDF - Dict structure", tags4, reference_tags)
    print_result(result4)
    results.append(result4)
    
    # pdfplumber tests
    print("[>] Test: Methode 5 - pdfplumber Texte")
    tags5 = test_pdfplumber_text()
    result5 = evaluate_result("pdfplumber - Texte", tags5, reference_tags)
    print_result(result5)
    results.append(result5)
    
    print("[>] Test: Methode 6 - pdfplumber Tables")
    tags6 = test_pdfplumber_tables()
    result6 = evaluate_result("pdfplumber - Tables", tags6, reference_tags)
    print_result(result6)
    results.append(result6)
    
    print("[>] Test: Methode 7 - pdfplumber Mots")
    tags7 = test_pdfplumber_words()
    result7 = evaluate_result("pdfplumber - Mots", tags7, reference_tags)
    print_result(result7)
    results.append(result7)
    
    print("[>] Test: Methode 8 - pdfplumber Lignes Y")
    tags8 = test_pdfplumber_lines_by_y()
    result8 = evaluate_result("pdfplumber - Lignes Y", tags8, reference_tags)
    print_result(result8)
    results.append(result8)
    
    print("[>] Test: Methode 9 - pdfplumber Tables optimise")
    tags9 = test_pdfplumber_tables_optimized()
    result9 = evaluate_result("pdfplumber - Tables opt.", tags9, reference_tags)
    print_result(result9)
    results.append(result9)
    
    print("[>] Test: Methode 10 - Combinee")
    tags_combined = test_combined(tags1, tags2, tags3, tags4, tags5, tags6, tags7, tags8, tags9)
    result10 = evaluate_result("Combinee (toutes)", tags_combined, reference_tags)
    print_result(result10)
    results.append(result10)
    
    # 3. Resume comparatif
    print("=" * 80)
    print("RESUME COMPARATIF")
    print("=" * 80)
    print()
    print(f"Reference CSV: {len(reference_tags)} tags uniques")
    print()
    
    # Trier par F1 score
    sorted_results = sorted(results, key=lambda x: x['f1_score'], reverse=True)
    
    print(f"{'Rang':<5} {'Methode':<30} {'Recall':<10} {'Precision':<10} {'F1':<10} {'Manquants':<10}")
    print("-" * 85)
    
    for i, result in enumerate(sorted_results, 1):
        print(f"{i:<5} {result['method_name']:<30} {result['recall']:>6}% {result['precision']:>9}% "
              f"{result['f1_score']:>6}% {result['false_negatives']:>8}")
    
    print()
    best = sorted_results[0]
    print(f"[+] MEILLEURE METHODE: {best['method_name']}")
    print(f"    Recall: {best['recall']}% ({best['true_positives']}/{len(reference_tags)} tags trouves)")
    print(f"    Precision: {best['precision']}%")
    print(f"    Tags manquants: {best['false_negatives']}")
    print()
    
    # 4. Analyse des tags manquants
    if best['missing_tags']:
        print("=" * 80)
        print(f"TAGS MANQUANTS ({len(best['missing_tags'])})")
        print("=" * 80)
        print()
        for tag in best['missing_tags']:
            print(f"  - {tag}")
        print()
        
        # Verifier si les tags manquants existent dans le PDF
        print("Verification dans le PDF brut (PyMuPDF):")
        import fitz
        doc = fitz.open(PDF_PATH)
        full_text = ""
        for page in doc:
            full_text += page.get_text()
        doc.close()
        
        found_in_raw = 0
        for tag in best['missing_tags']:
            if tag.lower() in full_text.lower():
                found_in_raw += 1
                print(f"  [+] '{tag}' existe dans le PDF brut")
            else:
                print(f"  [-] '{tag}' N'EXISTE PAS dans le PDF brut")
        
        print()
        print(f"Resultat: {found_in_raw}/{len(best['missing_tags'])} tags manquants sont dans le PDF brut")
        vraiment_absents = len(best['missing_tags']) - found_in_raw
        print(f"=> {vraiment_absents} tags sont VRAIMENT ABSENTS du PDF")
    
    print()
    print("=" * 80)
    print("FIN DU BENCHMARK")
    print("=" * 80)

if __name__ == "__main__":
    main()
