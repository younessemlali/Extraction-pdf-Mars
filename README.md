# 📄 PDF Extractor Select T.T

Application Streamlit pour l'extraction automatique des données des auto-factures Select T.T.

## 🎯 Fonctionnalités

- ✅ **Extraction automatique** de toutes les données des PDFs
- ✅ **Upload multiple** de fichiers PDF
- ✅ **Normalisation des montants** (gestion formats français/anglais)
- ✅ **Export Excel structuré** en 3 feuilles
- ✅ **Interface intuitive** avec Streamlit
- ✅ **Gestion d'erreurs robuste**

## 📊 Données extraites

### Informations principales
- N° Auto-facture (ex: 4968S0001)
- N° Commande (ex: 5600025054)
- Dates (facture, échéance)
- Destinataire (Mars Information Services, Mars Petcare, etc.)
- Batch ID & Assignment ID
- Montants (Net, TVA, Brut) avec normalisation automatique

### Détail des lignes
- Description des prestations
- Dates de période
- Unités (Each, Hours)
- Prix unitaires et quantités
- Montants détaillés par ligne

## 🚀 Installation et utilisation

### 1. Cloner le repository
```bash
git clone https://github.com/votre-username/pdf-extractor-selecttt.git
cd pdf-extractor-selecttt
```

### 2. Installer les dépendances
```bash
pip install -r requirements.txt
```

### 3. Lancer l'application
```bash
streamlit run app.py
```

### 4. Utilisation
1. Uploadez vos fichiers PDF d'auto-factures
2. Cliquez sur "Lancer l'extraction"
3. Vérifiez les résultats dans l'interface
4. Téléchargez le fichier Excel généré

## 📁 Structure du projet

```
pdf-extractor-selecttt/
│
├── app.py                 # Application Streamlit principale
├── requirements.txt       # Dépendances Python
├── README.md             # Documentation
└── .streamlit/
    └── config.toml       # Configuration Streamlit
```

## 📊 Format de sortie Excel

Le fichier Excel généré contient 3 feuilles :

### 1. Résumé_Factures
Vue d'ensemble de toutes les factures avec les informations principales.

### 2. Detail_Lignes  
Détail ligne par ligne de toutes les factures pour analyse fine.

### 3. Donnees_Analyse
Format optimisé pour l'analyse et le rapprochement avec d'autres systèmes.

## 🔧 Déploiement sur Streamlit Cloud

### 1. Fork ce repository sur GitHub

### 2. Connecter à Streamlit Cloud
- Aller sur [share.streamlit.io](https://share.streamlit.io)
- Connecter votre compte GitHub
- Déployer l'application

### 3. L'application sera accessible via une URL publique

## 🛠️ Technologies utilisées

- **Streamlit** : Interface web interactive
- **pandas** : Manipulation des données
- **PyPDF2** : Extraction de texte PDF (méthode 1)
- **pdfplumber** : Extraction de texte PDF (méthode 2, plus robuste)
- **openpyxl** : Génération de fichiers Excel

## 📈 Fonctionnalités avancées

### Normalisation des montants
L'application gère automatiquement différents formats :
- Format français : `12,942.38` (virgule = milliers)
- Format simple : `235.62` (point = décimales)
- Format français : `1059,61` (virgule = décimales)

### Gestion d'erreurs
- Extraction partielle si certaines données manquent
- Logs détaillés pour le débogage
- Interface utilisateur informative

### Performance
- Traitement par lot de multiples PDFs
- Barre de progression en temps réel
- Optimisation mémoire pour gros volumes

## 📞 Support

Pour toute question ou problème, créer une issue sur GitHub.

## 📝 License

Ce projet est sous licence MIT.

---

**PDF Extractor Select T.T** - Version 1.0 | Extraction automatique des auto-factures
