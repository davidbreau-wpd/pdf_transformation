#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Script pour extraire des données de rapports d'inspection PDF de Vestas
et les sauvegarder dans un format CSV standardisé conforme au modèle de référence.

Ce script peut traiter un fichier individuel ou tous les fichiers PDF d'un répertoire.
"""

import os
import re
import csv
import argparse
import pdfplumber
from pathlib import Path
from datetime import datetime


def extract_metadata(pdf):
    """Extrait les métadonnées du PDF (informations d'entête)."""
    metadata = {}
    
    # Extraire les premières pages pour les métadonnées
    text = ""
    for i in range(min(2, len(pdf.pages))):
        text += pdf.pages[i].extract_text() + "\n"
    
    # Extraire les informations principales
    turbine_number_match = re.search(r'Turbine No\./Id:?\s*(\d+)', text)
    if turbine_number_match:
        metadata['turbine_number'] = turbine_number_match.group(1)
    
    service_order_match = re.search(r'Service Order:?\s*(\d+)', text)
    if service_order_match:
        metadata['service_order'] = service_order_match.group(1)
    
    # Motif de l'appel
    reason_match = re.search(r'Check ICPE Electrical (V\d+)', text)
    if reason_match:
        turbine_type = reason_match.group(1)
        metadata['reason_for_call_out'] = f"Check ICPE Electrical {turbine_type}"
    
    # Chercher le libellé spécifique pour le formulaire (format complet incluant 3MW/etc.)
    form_title_match = re.search(r'Check ICPE Electrical V\d+-\w+', text)
    if form_title_match:
        metadata['form_title'] = form_title_match.group(0)
    elif 'reason_for_call_out' in metadata:
        # Par défaut, utiliser la raison de l'appel
        metadata['form_title'] = metadata['reason_for_call_out']
    
    return metadata


def extract_inspection_table(pdf):
    """Extrait la table d'inspection du PDF."""
    inspection_items = []
    # Dictionnaire pour suivre les éléments déjà vus
    items_seen = {}
    
    # Expression régulière pour capturer les numéros d'éléments d'inspection (par ex. "1.01", "10.02", etc.)
    item_number_pattern = r'^(\d+\.\d+)\s+(.+?)(?:\s*(OK|NOT OK|NOK|N/A|Not Applicable)\s*)?$'
    
    # Expression régulière pour détecter des lignes qui pourraient contenir des éléments d'inspection
    inspection_line_pattern = r'^\d+\.\d+'
    
    current_item_number = None
    current_description = None
    current_comment = None
    
    for page_num, page in enumerate(pdf.pages):
        text = page.extract_text()
        lines = text.split('\n')
        
        for line in lines:
            # Ignorer les lignes vides
            if not line.strip():
                continue
            
            # Vérifier si nous sommes dans une page de rapport d'inspection
            if "Service Inspection Form" in line or "0. DMS:" in line:
                continue
            
            # Si nous sommes en train de traiter un élément en cours et que cette ligne ne commence pas par un numéro
            if current_item_number and not re.match(inspection_line_pattern, line):
                # Si la ligne contient un statut connu, on l'ajoute à la description et on récupère le statut
                status_match = re.search(r'(OK|NOT OK|NOK|N/A|Not Applicable)$', line)
                if status_match:
                    status = status_match.group(1).upper()
                    additional_text = line[:status_match.start()].strip()
                    
                    # Mettre à jour la description ou le commentaire selon ce qui est approprié
                    if current_comment:
                        current_comment += " " + additional_text
                    else:
                        current_description += " " + additional_text
                else:
                    # Sinon, on ajoute simplement cette ligne à la description en cours
                    if current_comment:
                        current_comment += " " + line
                    else:
                        current_description += " " + line
            else:
                # Vérifier si cette ligne contient un numéro d'élément d'inspection
                item_number_match = re.match(item_number_pattern, line)
                
                if item_number_match:
                    # Si nous avons un élément en cours, on le sauvegarde avant d'en commencer un nouveau
                    if current_item_number and current_item_number not in items_seen:
                        items_seen[current_item_number] = True
                        inspection_items.append({
                            'item_number': current_item_number,
                            'description': current_description,
                            'comment': current_comment,
                            'status': status if 'status' in locals() else ""
                        })
                    
                    # Extraire les informations du nouvel élément
                    item_number = item_number_match.group(1).strip()
                    description = item_number_match.group(2).strip()
                    status = item_number_match.group(3).upper() if item_number_match.group(3) else ""
                    
                    # Vérifier si le statut "OK" est concaténé à la fin de la description
                    if not status:
                        status_patterns = ["OK", "NOT OK", "NOK", "N/A", "Not Applicable"]
                        for pattern in status_patterns:
                            if description.endswith(pattern):
                                description = description[:-len(pattern)].strip()
                                status = pattern.upper()
                                break
                    
                    # Vérifier si la description contient un commentaire après ":"
                    comment = None
                    if ":" in description:
                        parts = description.split(":", 1)
                        description = parts[0].strip()
                        comment_text = parts[1].strip() if len(parts) > 1 else None
                        
                        # Vérifier si le commentaire contient un statut à la fin
                        if comment_text:
                            comment = comment_text
                            for pattern in ["OK", "NOT OK", "NOK", "N/A", "Not Applicable"]:
                                if comment.endswith(pattern) and not status:
                                    comment = comment[:-len(pattern)].strip()
                                    status = pattern.upper()
                                    break
                    
                    # Si la description continue sur plusieurs lignes, on sauvegarde le numéro courant
                    current_item_number = item_number
                    current_description = description
                    current_comment = comment
                    
                    # Si cet élément existe déjà, on met à jour sa description
                    if item_number in items_seen:
                        for item in inspection_items:
                            if item['item_number'] == item_number:
                                item['description'] += " " + description
                                if comment and 'comment' in item:
                                    if item['comment']:
                                        item['comment'] += " " + comment
                                    else:
                                        item['comment'] = comment
                    # Si nous avons un statut pour cet élément, on le sauvegarde immédiatement
                    elif status:
                        items_seen[item_number] = True
                        inspection_items.append({
                            'item_number': item_number,
                            'description': description,
                            'comment': comment,
                            'status': status
                        })
    
    # Traiter les tables extraites avec pdfplumber pour trouver des éléments supplémentaires
    for page_num, page in enumerate(pdf.pages):
        tables = page.extract_tables()
        for table in tables:
            if not table:
                continue
            
            for row in table:
                # Ignorer les lignes vides ou avec moins de 2 cellules
                if not row or len(row) < 2 or not row[0]:
                    continue
                
                # Essayer de trouver un numéro d'élément d'inspection dans la première cellule
                first_cell = row[0].strip() if row[0] else ""
                if re.match(r'^\d+\.\d+$', first_cell):
                    item_number = first_cell
                    description = row[1] if len(row) > 1 else ""
                    
                    # Vérifier si la description contient un commentaire après ":"
                    comment = None
                    if description and ":" in description:
                        parts = description.split(":", 1)
                        description = parts[0].strip()
                        comment_text = parts[1].strip() if len(parts) > 1 else None
                        
                        # Vérifier si le commentaire contient un statut à la fin
                        if comment_text:
                            comment = comment_text
                            for pattern in ["OK", "NOT OK", "NOK", "N/A", "Not Applicable"]:
                                if comment.endswith(pattern):
                                    comment = comment[:-len(pattern)].strip()
                                    status = pattern.upper()
                                    break
                    
                    # Vérifier si le statut est concaténé à la fin de la description
                    status = ""
                    status_patterns = ["OK", "NOT OK", "NOK", "N/A", "Not Applicable"]
                    for pattern in status_patterns:
                        if description.endswith(pattern):
                            description = description[:-len(pattern)].strip()
                            status = pattern.upper()
                            break
                    
                    # Si aucun statut n'a été trouvé, chercher dans les cellules suivantes
                    if not status:
                        for i in range(2, len(row)):
                            if row[i]:
                                status_match = re.search(r'(OK|NOT OK|NOK|N/A|Not Applicable)', row[i])
                                if status_match:
                                    status = status_match.group(1).upper()
                                    break
                    
                    # Vérifier si cet élément existe déjà
                    if item_number not in items_seen:
                        items_seen[item_number] = True
                        inspection_items.append({
                            'item_number': item_number,
                            'description': description,
                            'comment': comment,
                            'status': status
                        })
    
    # Si nous avons encore un élément en cours, on le sauvegarde
    if current_item_number and current_item_number not in items_seen:
        items_seen[current_item_number] = True
        inspection_items.append({
            'item_number': current_item_number,
            'description': current_description,
            'comment': current_comment,
            'status': status if 'status' in locals() else ""
        })
    
    return inspection_items


def calculate_compliance_ratio(inspection_items):
    """Calcule le ratio de conformité (OK / total) en pourcentage."""
    if not inspection_items:
        return 0, 0, 0
    
    # Compter uniquement les éléments qui ont un statut (ignorer les sections)
    items_with_status = [item for item in inspection_items if item['status']]
    total_items = len(items_with_status)
    ok_items = sum(1 for item in items_with_status if item['status'] == 'OK')
    
    ratio = (ok_items / total_items) * 100 if total_items > 0 else 0
    
    return ok_items, total_items, ratio


def save_to_csv(metadata, inspection_items, compliance_data, output_path):
    """Sauvegarde les données extraites dans un fichier CSV."""
    with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        
        # Écrire l'en-tête
        writer.writerow(["0", "1", "2", "3", "4"])
        writer.writerow(["", "", "", "", ""])
        
        # Écrire les métadonnées
        writer.writerow(["Service Inspection Form", "", "", "", ""])
        
        for key, value in metadata.items():
            if key == "":
                continue
            writer.writerow([key, value, "", "", ""])
        
        # Écrire les éléments d'inspection regroupés par section
        current_section = None
        
        for item in inspection_items:
            item_number = item['item_number']
            section_number = item_number.split('.')[0]
            
            # Vérifier si nous commençons une nouvelle section
            if section_number != current_section:
                current_section = section_number
                
                # Trouver le titre de la section
                section_title = ""
                for i, item2 in enumerate(inspection_items):
                    if item2['item_number'].startswith(section_number + ".") and item2['description'] and not item2['description'][0].isdigit():
                        section_title = item2['description']
                        break
                
                # Écrire l'en-tête de section
                writer.writerow([section_number, section_title, "", "", ""])
            
            # Écrire l'élément d'inspection
            description = item['description']
            status = item['status']
            comment = item.get('comment', '')
            
            # Tous les cas utilisent maintenant le même format avec description, commentaire et statut dans des colonnes séparées
            writer.writerow([item['item_number'], description, comment, status, ''])


def process_pdf(pdf_path, output_path=None, debug=False):
    """Traite un fichier PDF et extrait les données."""
    print(f"Traitement du fichier: {pdf_path}")
    
    # Déterminer le chemin de sortie si non spécifié
    if not output_path:
        pdf_filename = os.path.basename(pdf_path)
        pdf_name = os.path.splitext(pdf_filename)[0]
        
        # Utiliser un chemin relatif pour le dossier output
        script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        output_dir = os.path.join(script_dir, "output")
        
        # Créer le répertoire s'il n'existe pas
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        output_path = os.path.join(output_dir, f"{pdf_name}_output.csv")
    
    # Ouvrir le PDF
    with pdfplumber.open(pdf_path) as pdf:
        # Extraire les métadonnées
        metadata = extract_metadata(pdf)
        
        # Mode debug: afficher le texte brut
        if debug:
            print("====== TEXTE BRUT DU PDF ======")
            for i, page in enumerate(pdf.pages):
                print(f"--- PAGE {i+1} ---")
                print(page.extract_text())
                print()
            print("===============================")
        
        # Extraire la table d'inspection
        inspection_items = extract_inspection_table(pdf)
        
        # Calculer le ratio de conformité
        compliance_data = calculate_compliance_ratio(inspection_items)
        
        # Sauvegarder les données dans un fichier CSV
        save_to_csv(metadata, inspection_items, compliance_data, output_path)
    
    print(f"Données extraites et sauvegardées dans: {output_path}")
    print(f"Conformité: {compliance_data[0]} OK sur {compliance_data[1]} éléments ({compliance_data[2]:.2f}%)")
    
    return metadata, inspection_items, compliance_data


def process_directory(directory_path, output_directory=None, debug=False):
    """Traite tous les fichiers PDF dans le répertoire spécifié."""
    pdf_files = [f for f in os.listdir(directory_path) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"Aucun fichier PDF trouvé dans {directory_path}")
        return
    
    print(f"Traitement de {len(pdf_files)} fichiers PDF...")
    
    # Si aucun répertoire de sortie n'est spécifié, utiliser le dossier output du projet
    if not output_directory:
        script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        output_directory = os.path.join(script_dir, "output")
    
    # Créer le répertoire de sortie s'il n'existe pas
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    
    # Traiter chaque fichier PDF
    results = []
    for pdf_file in pdf_files:
        pdf_path = os.path.join(directory_path, pdf_file)
        
        # Déterminer le chemin de sortie
        pdf_name = os.path.splitext(pdf_file)[0]
        output_path = os.path.join(output_directory, f"{pdf_name}_output.csv")
        
        # Traiter le fichier
        metadata, inspection_items, compliance_data = process_pdf(pdf_path, output_path, debug)
        
        # Stocker les résultats
        results.append({
            'pdf_file': pdf_file,
            'output_file': os.path.basename(output_path) if output_path else None,
            'ok_items': compliance_data[0],
            'total_items': compliance_data[1],
            'compliance_ratio': compliance_data[2]
        })
    
    # Créer un rapport de traitement
    report_path = os.path.join(output_directory, f"processing_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    with open(report_path, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['pdf_file', 'output_file', 'ok_items', 'total_items', 'compliance_ratio']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        for result in results:
            writer.writerow(result)
    
    print(f"Rapport de traitement sauvegardé dans: {report_path}")
    
    return results


def main():
    parser = argparse.ArgumentParser(description='Extraire des données de rapports d\'inspection PDF de Vestas et les sauvegarder au format CSV standardisé.')
    parser.add_argument('input_path', help='Chemin vers le fichier PDF ou le répertoire contenant les fichiers PDF à traiter')
    parser.add_argument('--output', '-o', help='Chemin vers le fichier CSV de sortie ou le répertoire de sortie')
    parser.add_argument('--debug', '-d', action='store_true', help='Mode debug: afficher le texte brut extrait du PDF')
    
    args = parser.parse_args()
    
    # Vérifier si c'est un fichier ou un répertoire
    if os.path.isfile(args.input_path):
        process_pdf(args.input_path, args.output, args.debug)
    elif os.path.isdir(args.input_path):
        process_directory(args.input_path, args.output, args.debug)
    else:
        print(f"Le chemin spécifié n'existe pas: {args.input_path}")


if __name__ == "__main__":
    main()
