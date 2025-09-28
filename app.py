import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
import PyPDF2
import pdfplumber
from typing import Dict, List, Optional, Tuple
import logging

# Configuration de la page
st.set_page_config(
    page_title="PDF Extractor Select T.T",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Configuration du logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PDFExtractor:
    """Classe pour extraire les données des auto-factures Select T.T"""
    
    def __init__(self):
        self.extracted_data = []
    
    def normalize_amount(self, amount_str: str) -> Optional[float]:
        """Normalise les montants avec différents formats"""
        try:
            if not amount_str:
                return None
            
            # Nettoyer la chaîne
            amount_str = str(amount_str).strip()
            
            # Gérer les formats français et anglais
            if ',' in amount_str and '.' in amount_str:
                # Format: 12,942.38 (virgule = milliers, point = décimales)
                amount_str = amount_str.replace(',', '')
            elif ',' in amount_str:
                # Format: 1059,61 (virgule = décimales)
                amount_str = amount_str.replace(',', '.')
            
            # Supprimer les espaces et caractères non numériques sauf le point
            amount_str = re.sub(r'[^\d\.]', '', amount_str)
            
            result = float(amount_str)
            return result if not pd.isna(result) else None
            
        except (ValueError, TypeError):
            logger.warning(f"Impossible de normaliser le montant: {amount_str}")
            return None
    
    def extract_with_regex(self, text: str, pattern: str, group: int = 1) -> Optional[str]:
        """Extrait une valeur avec regex de manière sécurisée"""
        try:
            match = re.search(pattern, text, re.IGNORECASE)
            return match.group(group).strip() if match and match.group(group) else None
        except (AttributeError, IndexError):
            return None
    
    def extract_invoice_data(self, pdf_file) -> Dict:
        """Extrait les données d'une facture PDF"""
        try:
            # Lire le PDF avec pdfplumber (plus robuste)
            with pdfplumber.open(pdf_file) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() or ""
            
            # Si pdfplumber échoue, essayer PyPDF2
            if not text.strip():
                pdf_file.seek(0)
                reader = PyPDF2.PdfReader(pdf_file)
                text = ""
                for page in reader.pages:
                    text += page.extract_text() or ""
            
            logger.info(f"Texte extrait (premiers 200 chars): {text[:200]}...")
            
            # Extraction des données principales
            data = {
                'nom_fichier': pdf_file.name,
                'numero_facture': self.extract_invoice_number(text),
                'date_facture': self.extract_invoice_date(text),
                'numero_commande': self.extract_purchase_order(text),
                'date_echeance': self.extract_due_date(text),
                'destinataire': self.extract_recipient(text),
                'emetteur': 'Select T.T',
                'batch_id': self.extract_batch_id(text),
                'assignment_id': self.extract_assignment_id(text),
                'total_net': self.extract_total_net(text),
                'total_tva': self.extract_total_vat(text),
                'total_brut': self.extract_total_gross(text),
                'taux_tva': '20%',
                'devise': 'EUR',
                'lignes_detail': self.extract_line_items(text),
                'rubriques_analyse': None  # Sera calculé après
            }
            
            # Analyser les rubriques si on a des lignes de détail
            if data['lignes_detail']:
                data['rubriques_analyse'] = self.analyze_rubriques(data['lignes_detail'])
            
            return data
            
        except Exception as e:
            logger.error(f"Erreur lors de l'extraction de {pdf_file.name}: {e}")
            return {
                'nom_fichier': pdf_file.name,
                'erreur': str(e),
                'numero_facture': None,
                'date_facture': None,
                'numero_commande': None,
                'date_echeance': None,
                'destinataire': None,
                'emetteur': 'Select T.T',
                'batch_id': None,
                'assignment_id': None,
                'total_net': None,
                'total_tva': None,
                'total_brut': None,
                'taux_tva': '20%',
                'devise': 'EUR',
                'lignes_detail': []
            }
    
    def extract_invoice_number(self, text: str) -> Optional[str]:
        """Extrait le numéro de facture"""
        patterns = [
            r'Invoice ID/Number[^0-9A-Z]*([0-9]+S[0-9]+)',
            r'(\d{4}S\d{4})',
            r'Invoice.*?(\d{4}S\d{3,4})'
        ]
        
        for pattern in patterns:
            result = self.extract_with_regex(text, pattern)
            if result:
                return result
        return None
    
    def extract_invoice_date(self, text: str) -> Optional[str]:
        """Extrait la date de facture"""
        patterns = [
            r'Invoice Date[^0-9]*(\d{4}/\d{2}/\d{2})',
            r'(\d{4}/\d{2}/\d{2})'
        ]
        
        for pattern in patterns:
            result = self.extract_with_regex(text, pattern)
            if result:
                return result
        return None
    
    def extract_purchase_order(self, text: str) -> Optional[str]:
        """Extrait le numéro de commande"""
        patterns = [
            r'Purchase Order[^0-9]*(\d{10})',
            r'Bon de commande[^0-9]*(\d{10})',
            r'(\d{10})'
        ]
        
        for pattern in patterns:
            result = self.extract_with_regex(text, pattern)
            if result and len(result) == 10:
                return result
        return None
    
    def extract_due_date(self, text: str) -> Optional[str]:
        """Extrait la date d'échéance"""
        patterns = [
            r'Payment Terms[^0-9]*(\d{4}/\d{2}/\d{2})',
            r'Modalités de Paiement[^0-9]*(\d{4}/\d{2}/\d{2})'
        ]
        
        for pattern in patterns:
            result = self.extract_with_regex(text, pattern)
            if result:
                return result
        return None
    
    def extract_recipient(self, text: str) -> str:
        """Extrait le destinataire"""
        if 'Mars Information Services' in text:
            return 'Mars Information Services'
        elif 'Mars Petcare Food France' in text:
            return 'Mars Petcare Food France SAS'
        elif 'Mars' in text:
            return 'Mars (à préciser)'
        return 'Non déterminé'
    
    def extract_batch_id(self, text: str) -> Optional[str]:
        """Extrait le Batch ID"""
        result = self.extract_with_regex(text, r'(\d{4})_\d{5}_')
        return result
    
    def extract_assignment_id(self, text: str) -> Optional[str]:
        """Extrait l'Assignment ID"""
        result = self.extract_with_regex(text, r'\d{4}_(\d{5})_')
        return result
    
    def extract_total_net(self, text: str) -> Optional[float]:
        """Extrait le total net"""
        patterns = [
            r'Invoice Total.*?EUR.*?([\d,\.]+)\s+[\d,\.]+\s+[\d,\.]+',
            r'Invoice Total.*?EUR.*?([\d,\.]+)',
            r'Net Amount.*?([\d,\.]+)',
            r'Montant Net.*?([\d,\.]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return self.normalize_amount(match.group(1))
        return None
    
    def extract_total_vat(self, text: str) -> Optional[float]:
        """Extrait le total TVA"""
        patterns = [
            r'Invoice Total.*?EUR.*?[\d,\.]+\s+([\d,\.]+)\s+[\d,\.]+',
            r'VAT Amount.*?([\d,\.]+)',
            r'Montant TVA.*?([\d,\.]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return self.normalize_amount(match.group(1))
        return None
    
    def extract_total_gross(self, text: str) -> Optional[float]:
        """Extrait le total brut"""
        patterns = [
            r'Invoice Total.*?EUR.*?[\d,\.]+\s+[\d,\.]+\s+([\d,\.]+)',
            r'Gross Amount.*?([\d,\.]+)',
            r'Montant brut.*?([\d,\.]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return self.normalize_amount(match.group(1))
        return None
    
    def extract_line_items(self, text: str) -> List[Dict]:
        """Extrait les lignes de détail avec analyse des rubriques"""
        lines = []
        
        # Pattern pour les lignes de facture (amélioré pour capturer plus d'infos)
        line_patterns = [
            # Pattern principal pour lignes détaillées
            r'(\d{4}_\d{5}_[^0-9]*?)\s+(\d{4}/\d{2}/\d{2})\s+(\w+)\s+([\d,\.]+)\s+(\d+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)',
            # Pattern alternatif pour autres formats
            r'(\d{4}_\d{5}_[^\s]+)\s+(\d{4}/\d{2}/\d{2})\s+(\w+)\s+([\d,\.]+)\s+(\d+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)'
        ]
        
        for pattern in line_patterns:
            matches = re.finditer(pattern, text, re.MULTILINE)
            
            for match in matches:
                try:
                    description = match.group(1).strip()
                    
                    # Extraire les informations de la description
                    batch_id, assignment_id, type_prestation = self.parse_description(description)
                    
                    line_data = {
                        'description': description,
                        'batch_id': batch_id,
                        'assignment_id': assignment_id,
                        'type_prestation': type_prestation,
                        'code_rubrique': self.extract_rubrique_code(description, text),
                        'date_periode': match.group(2),
                        'unite': match.group(3),
                        'prix_unitaire': self.normalize_amount(match.group(4)),
                        'quantite': int(match.group(5)),
                        'montant_net': self.normalize_amount(match.group(6)),
                        'montant_tva': self.normalize_amount(match.group(7)),
                        'montant_brut': self.normalize_amount(match.group(8))
                    }
                    lines.append(line_data)
                except (ValueError, IndexError) as e:
                    logger.warning(f"Erreur lors de l'extraction d'une ligne: {e}")
                    continue
        
        return lines
    
    def parse_description(self, description: str) -> Tuple[str, str, str]:
        """Parse la description pour extraire Batch ID, Assignment ID et type de prestation"""
        try:
            # Pattern: 4973_65744_Temporary employees - Expense
            parts = description.split('_')
            if len(parts) >= 3:
                batch_id = parts[0]
                assignment_id = parts[1]
                type_part = '_'.join(parts[2:])
                
                # Identifier le type de prestation
                if 'Expense' in type_part:
                    type_prestation = 'Expense'
                elif 'Timesheet' in type_part:
                    type_prestation = 'Timesheet'
                else:
                    type_prestation = 'Autre'
                
                return batch_id, assignment_id, type_prestation
            else:
                return None, None, 'Non déterminé'
        except Exception:
            return None, None, 'Non déterminé'
    
    def extract_rubrique_code(self, description: str, full_text: str) -> Optional[str]:
        """Extrait le code rubrique (ex: OT125) depuis la description ou le texte complet"""
        # Chercher dans la description d'abord
        rubrique_patterns = [
            r'([A-Z]{2}\d{3})',  # Pattern OT125
            r'Code rubrique[^A-Z]*([A-Z]{2}\d{3})',
            r'rubrique[^A-Z]*([A-Z]{2}\d{3})'
        ]
        
        for pattern in rubrique_patterns:
            match = re.search(pattern, description, re.IGNORECASE)
            if match:
                return match.group(1)
        
        # Si pas trouvé dans description, chercher dans le texte complet
        for pattern in rubrique_patterns:
            match = re.search(pattern, full_text, re.IGNORECASE)
            if match:
                return match.group(1)
        
        return None
    
    def analyze_rubriques(self, lines: List[Dict]) -> List[Dict]:
        """Analyse et regroupe les données par rubrique"""
        rubriques_data = {}
        
        for line in lines:
            # Créer une clé unique pour chaque rubrique/type
            rubrique_key = f"{line.get('code_rubrique', 'SANS_CODE')}_{line.get('type_prestation', 'SANS_TYPE')}"
            
            if rubrique_key not in rubriques_data:
                rubriques_data[rubrique_key] = {
                    'code_rubrique': line.get('code_rubrique', 'Non déterminé'),
                    'type_prestation': line.get('type_prestation', 'Non déterminé'),
                    'batch_id': line.get('batch_id'),
                    'assignment_id': line.get('assignment_id'),
                    'nb_lignes': 0,
                    'total_quantite': 0,
                    'total_net': 0,
                    'total_tva': 0,
                    'total_brut': 0,
                    'unites': set(),
                    'periodes': set()
                }
            
            # Agrégation des données
            rubrique = rubriques_data[rubrique_key]
            rubrique['nb_lignes'] += 1
            rubrique['total_quantite'] += line.get('quantite', 0)
            rubrique['total_net'] += line.get('montant_net', 0) or 0
            rubrique['total_tva'] += line.get('montant_tva', 0) or 0
            rubrique['total_brut'] += line.get('montant_brut', 0) or 0
            
            if line.get('unite'):
                rubrique['unites'].add(line['unite'])
            if line.get('date_periode'):
                rubrique['periodes'].add(line['date_periode'])
        
        # Convertir les sets en strings pour l'export
        for rubrique in rubriques_data.values():
            rubrique['unites'] = ', '.join(sorted(rubrique['unites'])) if rubrique['unites'] else ''
            rubrique['periodes'] = ', '.join(sorted(rubrique['periodes'])) if rubrique['periodes'] else ''
        
        return list(rubriques_data.values())
    
    def process_files(self, uploaded_files) -> List[Dict]:
        """Traite tous les fichiers uploadés"""
        self.extracted_data = []
        
        for uploaded_file in uploaded_files:
            st.write(f"📄 Traitement de: {uploaded_file.name}")
            
            # Créer un objet BytesIO pour le fichier
            file_content = io.BytesIO(uploaded_file.read())
            file_content.name = uploaded_file.name
            
            # Extraire les données
            data = self.extract_invoice_data(file_content)
            self.extracted_data.append(data)
            
            # Réinitialiser le pointeur du fichier
            uploaded_file.seek(0)
        
        return self.extracted_data
    
    def create_excel_report(self) -> io.BytesIO:
        """Crée un rapport Excel avec les données extraites incluant l'analyse par rubriques"""
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Feuille 1: Résumé des factures
            summary_data = []
            for data in self.extracted_data:
                summary_data.append({
                    'Nom_Fichier': data['nom_fichier'],
                    'Numero_Facture': data['numero_facture'],
                    'Date_Facture': data['date_facture'],
                    'Numero_Commande': data['numero_commande'],
                    'Date_Echeance': data['date_echeance'],
                    'Destinataire': data['destinataire'],
                    'Batch_ID': data['batch_id'],
                    'Assignment_ID': data['assignment_id'],
                    'Total_Net_EUR': data['total_net'],
                    'Total_TVA_EUR': data['total_tva'],
                    'Total_Brut_EUR': data['total_brut'],
                    'Nb_Lignes_Detail': len(data['lignes_detail']) if data['lignes_detail'] else 0,
                    'Nb_Rubriques': len(data['rubriques_analyse']) if data['rubriques_analyse'] else 0,
                    'Devise': data['devise'],
                    'Erreur': data.get('erreur', '')
                })
            
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='Résumé_Factures', index=False)
            
            # Feuille 2: Détail des lignes (enrichi avec rubriques)
            detail_data = []
            for data in self.extracted_data:
                if data['lignes_detail']:
                    for line in data['lignes_detail']:
                        detail_data.append({
                            'Nom_Fichier': data['nom_fichier'],
                            'Numero_Facture': data['numero_facture'],
                            'Numero_Commande': data['numero_commande'],
                            'Batch_ID': line.get('batch_id', ''),
                            'Assignment_ID': line.get('assignment_id', ''),
                            'Code_Rubrique': line.get('code_rubrique', ''),
                            'Type_Prestation': line.get('type_prestation', ''),
                            'Description': line['description'],
                            'Date_Periode': line['date_periode'],
                            'Unite': line['unite'],
                            'Prix_Unitaire': line['prix_unitaire'],
                            'Quantite': line['quantite'],
                            'Montant_Net': line['montant_net'],
                            'Montant_TVA': line['montant_tva'],
                            'Montant_Brut': line['montant_brut']
                        })
                else:
                    # Ajouter une ligne même si pas de détail
                    detail_data.append({
                        'Nom_Fichier': data['nom_fichier'],
                        'Numero_Facture': data['numero_facture'],
                        'Numero_Commande': data['numero_commande'],
                        'Batch_ID': data['batch_id'],
                        'Assignment_ID': data['assignment_id'],
                        'Code_Rubrique': '',
                        'Type_Prestation': 'Total facture',
                        'Description': 'Total facture',
                        'Date_Periode': data['date_facture'],
                        'Unite': 'Global',
                        'Prix_Unitaire': None,
                        'Quantite': 1,
                        'Montant_Net': data['total_net'],
                        'Montant_TVA': data['total_tva'],
                        'Montant_Brut': data['total_brut']
                    })
            
            df_detail = pd.DataFrame(detail_data)
            df_detail.to_excel(writer, sheet_name='Detail_Lignes', index=False)
            
            # Feuille 3: Données pour analyse
            analysis_data = []
            for data in self.extracted_data:
                analysis_data.append({
                    'Numero_Facture': data['numero_facture'],
                    'Numero_Commande': data['numero_commande'],
                    'Date_Facture': data['date_facture'],
                    'Semaine_Finissant_Le': None,  # À remplir manuellement ou extraire
                    'Emetteur': data['emetteur'],
                    'Destinataire': data['destinataire'],
                    'Batch_ID': data['batch_id'],
                    'Assignment_ID': data['assignment_id'],
                    'Total_Net': data['total_net'],
                    'Total_TVA': data['total_tva'],
                    'Total_Brut': data['total_brut'],
                    'Nb_Lignes': len(data['lignes_detail']) if data['lignes_detail'] else 1,
                    'Nb_Rubriques': len(data['rubriques_analyse']) if data['rubriques_analyse'] else 0,
                    'Statut_Extraction': 'Succès' if data['numero_facture'] else 'Partiel'
                })
            
            df_analysis = pd.DataFrame(analysis_data)
            df_analysis.to_excel(writer, sheet_name='Donnees_Analyse', index=False)
            
            # NOUVELLE Feuille 4: Analyse par rubriques
            rubriques_data = []
            for data in self.extracted_data:
                if data['rubriques_analyse']:
                    for rubrique in data['rubriques_analyse']:
                        rubriques_data.append({
                            'Nom_Fichier': data['nom_fichier'],
                            'Numero_Facture': data['numero_facture'],
                            'Numero_Commande': data['numero_commande'],
                            'Code_Rubrique': rubrique['code_rubrique'],
                            'Type_Prestation': rubrique['type_prestation'],
                            'Batch_ID': rubrique['batch_id'],
                            'Assignment_ID': rubrique['assignment_id'],
                            'Nb_Lignes': rubrique['nb_lignes'],
                            'Total_Quantite': rubrique['total_quantite'],
                            'Unites': rubrique['unites'],
                            'Periodes': rubrique['periodes'],
                            'Total_Net_EUR': rubrique['total_net'],
                            'Total_TVA_EUR': rubrique['total_tva'],
                            'Total_Brut_EUR': rubrique['total_brut'],
                            'Pourcentage_Facture': round((rubrique['total_net'] / data['total_net']) * 100, 2) if data['total_net'] and data['total_net'] > 0 else 0
                        })
            
            if rubriques_data:
                df_rubriques = pd.DataFrame(rubriques_data)
                df_rubriques.to_excel(writer, sheet_name='Analyse_Rubriques', index=False)
            
            # NOUVELLE Feuille 5: Synthèse par type de prestation
            synthese_data = {}
            for data in self.extracted_data:
                if data['rubriques_analyse']:
                    for rubrique in data['rubriques_analyse']:
                        type_key = rubrique['type_prestation']
                        if type_key not in synthese_data:
                            synthese_data[type_key] = {
                                'Type_Prestation': type_key,
                                'Nb_Factures': 0,
                                'Nb_Lignes_Total': 0,
                                'Total_Net_EUR': 0,
                                'Total_TVA_EUR': 0,
                                'Total_Brut_EUR': 0,
                                'Codes_Rubriques': set(),
                                'Factures': set()
                            }
                        
                        synthese = synthese_data[type_key]
                        synthese['Nb_Lignes_Total'] += rubrique['nb_lignes']
                        synthese['Total_Net_EUR'] += rubrique['total_net']
                        synthese['Total_TVA_EUR'] += rubrique['total_tva']
                        synthese['Total_Brut_EUR'] += rubrique['total_brut']
                        synthese['Codes_Rubriques'].add(rubrique['code_rubrique'])
                        synthese['Factures'].add(data['numero_facture'])
            
            # Convertir pour export
            synthese_export = []
            for synthese in synthese_data.values():
                synthese_export.append({
                    'Type_Prestation': synthese['Type_Prestation'],
                    'Nb_Factures': len(synthese['Factures']),
                    'Nb_Lignes_Total': synthese['Nb_Lignes_Total'],
                    'Codes_Rubriques': ', '.join(sorted(synthese['Codes_Rubriques'])),
                    'Total_Net_EUR': synthese['Total_Net_EUR'],
                    'Total_TVA_EUR': synthese['Total_TVA_EUR'],
                    'Total_Brut_EUR': synthese['Total_Brut_EUR']
                })
            
            if synthese_export:
                df_synthese = pd.DataFrame(synthese_export)
                df_synthese.to_excel(writer, sheet_name='Synthese_Prestations', index=False)
        
        output.seek(0)
        return output


def main():
    st.title("📄 PDF Extractor Select T.T")
    st.markdown("### Application d'extraction automatique des auto-factures Select T.T")
    
    # Sidebar
    st.sidebar.header("📋 Instructions")
    st.sidebar.markdown("""
    1. **Uploadez** vos fichiers PDF
    2. **Lancez** l'extraction
    3. **Vérifiez** les résultats
    4. **Téléchargez** le fichier Excel
    """)
    
    st.sidebar.header("📊 Données extraites")
    st.sidebar.markdown("""
    - N° Auto-facture
    - N° Commande  
    - Dates (facture, échéance)
    - Destinataire
    - Batch ID & Assignment ID
    - Montants (Net, TVA, Brut)
    - Détail des lignes
    """)
    
    # Upload des fichiers
    st.header("📤 Upload des fichiers PDF")
    uploaded_files = st.file_uploader(
        "Sélectionnez vos fichiers PDF d'auto-factures",
        type=['pdf'],
        accept_multiple_files=True,
        help="Vous pouvez sélectionner plusieurs fichiers PDF à la fois"
    )
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} fichier(s) sélectionné(s)")
        
        # Afficher la liste des fichiers
        with st.expander("📁 Fichiers sélectionnés", expanded=True):
            for i, file in enumerate(uploaded_files, 1):
                st.write(f"{i}. {file.name} ({file.size / 1024:.1f} KB)")
        
        # Bouton d'extraction
        if st.button("🚀 Lancer l'extraction", type="primary"):
            with st.spinner("Extraction en cours..."):
                extractor = PDFExtractor()
                
                # Barre de progression
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                try:
                    extracted_data = extractor.process_files(uploaded_files)
                    progress_bar.progress(1.0)
                    status_text.text("✅ Extraction terminée!")
                    
                    # Affichage des résultats
                    st.header("📊 Résultats de l'extraction")
                    
                    # Statistiques enrichies
                    col1, col2, col3, col4, col5 = st.columns(5)
                    
                    with col1:
                        st.metric("📄 Fichiers traités", len(extracted_data))
                    
                    with col2:
                        success_count = sum(1 for d in extracted_data if d['numero_facture'])
                        st.metric("✅ Extractions réussies", success_count)
                    
                    with col3:
                        total_net = sum(d['total_net'] or 0 for d in extracted_data)
                        st.metric("💰 Total Net (EUR)", f"{total_net:,.2f}")
                    
                    with col4:
                        total_gross = sum(d['total_brut'] or 0 for d in extracted_data)
                        st.metric("💰 Total Brut (EUR)", f"{total_gross:,.2f}")
                    
                    with col5:
                        total_lignes = sum(len(d['lignes_detail']) if d['lignes_detail'] else 0 for d in extracted_data)
                        st.metric("📋 Lignes extraites", total_lignes)
                    
                    # Nouvelle section : Analyse par rubriques
                    st.subheader("🏷️ Analyse par rubriques")
                    
                    # Regrouper toutes les rubriques
                    all_rubriques = []
                    for data in extracted_data:
                        if data['rubriques_analyse']:
                            for rubrique in data['rubriques_analyse']:
                                rubrique['numero_facture'] = data['numero_facture']
                                all_rubriques.append(rubrique)
                    
                    if all_rubriques:
                        # Créer un tableau synthétique par type de prestation
                        types_prestations = {}
                        for rubrique in all_rubriques:
                            type_key = rubrique['type_prestation']
                            if type_key not in types_prestations:
                                types_prestations[type_key] = {
                                    'nb_factures': 0,
                                    'nb_lignes': 0,
                                    'total_net': 0,
                                    'total_brut': 0,
                                    'factures': set()
                                }
                            
                            types_prestations[type_key]['nb_lignes'] += rubrique['nb_lignes']
                            types_prestations[type_key]['total_net'] += rubrique['total_net']
                            types_prestations[type_key]['total_brut'] += rubrique['total_brut']
                            types_prestations[type_key]['factures'].add(rubrique['numero_facture'])
                        
                        # Affichage des métriques par type
                        st.markdown("**Répartition par type de prestation:**")
                        cols = st.columns(len(types_prestations))
                        
                        for i, (type_name, stats) in enumerate(types_prestations.items()):
                            with cols[i]:
                                st.metric(
                                    f"🔧 {type_name}",
                                    f"{stats['total_net']:,.2f} €",
                                    f"{stats['nb_lignes']} lignes"
                                )
                        
                        # Tableau détaillé des rubriques
                        rubriques_display = []
                        for rubrique in all_rubriques:
                            rubriques_display.append({
                                'Facture': rubrique['numero_facture'],
                                'Code Rubrique': rubrique['code_rubrique'] or '❌',
                                'Type': rubrique['type_prestation'],
                                'Lignes': rubrique['nb_lignes'],
                                'Quantité': rubrique['total_quantite'],
                                'Net (€)': f"{rubrique['total_net']:,.2f}",
                                'Brut (€)': f"{rubrique['total_brut']:,.2f}"
                            })
                        
                        if rubriques_display:
                            st.dataframe(pd.DataFrame(rubriques_display), use_container_width=True)
                    else:
                        st.info("ℹ️ Aucune rubrique détaillée trouvée dans les PDFs")
                    
                    # Tableau de résultats principal
                    st.subheader("📋 Détail des extractions")
                    
                    with col2:
                        success_count = sum(1 for d in extracted_data if d['numero_facture'])
                        st.metric("✅ Extractions réussies", success_count)
                    
                    with col3:
                        total_net = sum(d['total_net'] or 0 for d in extracted_data)
                        st.metric("💰 Total Net (EUR)", f"{total_net:,.2f}")
                    
                    with col4:
                        total_gross = sum(d['total_brut'] or 0 for d in extracted_data)
                        st.metric("💰 Total Brut (EUR)", f"{total_gross:,.2f}")
                    
                    # Tableau de résultats
                    st.subheader("📋 Détail des extractions")
                    
                    display_data = []
                    for data in extracted_data:
                        display_data.append({
                            'Fichier': data['nom_fichier'],
                            'N° Facture': data['numero_facture'] or '❌',
                            'N° Commande': data['numero_commande'] or '❌',
                            'Date': data['date_facture'] or '❌',
                            'Destinataire': data['destinataire'],
                            'Total Net': f"{data['total_net']:,.2f} €" if data['total_net'] else '❌',
                            'Total Brut': f"{data['total_brut']:,.2f} €" if data['total_brut'] else '❌',
                            'Statut': '✅' if data['numero_facture'] else '⚠️'
                        })
                    
                    df_display = pd.DataFrame(display_data)
                    st.dataframe(df_display, use_container_width=True)
                    
                    # Génération du fichier Excel
                    st.header("💾 Export Excel")
                    
                    with st.spinner("Génération du fichier Excel..."):
                        excel_file = extractor.create_excel_report()
                    
                    # Bouton de téléchargement
                    filename = f"Extraction_PDF_SelectTT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    
                    st.download_button(
                        label="📊 Télécharger le fichier Excel",
                        data=excel_file.getvalue(),
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                    
                    st.success("🎉 Extraction terminée ! Vous pouvez télécharger le fichier Excel.")
                    
                    # Informations sur le fichier Excel
                    with st.expander("ℹ️ Contenu du fichier Excel"):
                        st.markdown("""
                        **Le fichier Excel contient 3 feuilles :**
                        
                        1. **Résumé_Factures** : Vue d'ensemble de toutes les factures
                        2. **Detail_Lignes** : Détail ligne par ligne de toutes les factures  
                        3. **Donnees_Analyse** : Format optimisé pour l'analyse et le rapprochement
                        """)
                
                except Exception as e:
                    st.error(f"❌ Erreur lors de l'extraction: {e}")
                    logger.error(f"Erreur: {e}")
    
    else:
        st.info("👆 Commencez par uploader vos fichiers PDF d'auto-factures")
    
    # Footer
    st.markdown("---")
    st.markdown("**PDF Extractor Select T.T** - Version 1.0 | Développé pour l'extraction automatique des auto-factures")


if __name__ == "__main__":
    main()
