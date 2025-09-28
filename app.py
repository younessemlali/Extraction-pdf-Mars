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
                'lignes_detail': self.extract_line_items(text)
            }
            
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
        """Extrait les lignes de détail"""
        lines = []
        
        # Pattern pour les lignes de facture
        line_pattern = r'(\d{4}_\d{5}_[^0-9]*?)\s+(\d{4}/\d{2}/\d{2})\s+(\w+)\s+([\d,\.]+)\s+(\d+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)'
        
        matches = re.finditer(line_pattern, text)
        
        for match in matches:
            try:
                line_data = {
                    'description': match.group(1).strip(),
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
        """Crée un rapport Excel avec les données extraites"""
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
                    'Devise': data['devise'],
                    'Erreur': data.get('erreur', '')
                })
            
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='Résumé_Factures', index=False)
            
            # Feuille 2: Détail des lignes
            detail_data = []
            for data in self.extracted_data:
                if data['lignes_detail']:
                    for line in data['lignes_detail']:
                        detail_data.append({
                            'Nom_Fichier': data['nom_fichier'],
                            'Numero_Facture': data['numero_facture'],
                            'Numero_Commande': data['numero_commande'],
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
                    'Statut_Extraction': 'Succès' if data['numero_facture'] else 'Partiel'
                })
            
            df_analysis = pd.DataFrame(analysis_data)
            df_analysis.to_excel(writer, sheet_name='Donnees_Analyse', index=False)
        
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
                    
                    # Statistiques
                    col1, col2, col3, col4 = st.columns(4)
                    
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
