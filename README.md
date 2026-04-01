# InstalPro

InstalPro est une application Streamlit de pilotage terrain pour :
- importer le tableau **ETAT FTTH RTC**
- filtrer les lignes par jour et par produit
- affecter chaque ligne à un **vrai agent**
- envoyer la ligne complète à l’agent par **WhatsApp**
- saisir le **retour d’intervention**
- envoyer ce retour par **email Outlook** configuré par l’administrateur

---

## Fonctionnalités principales

### 1. Import du tableau ETAT
L’application permet d’importer :
- `ETAT FTTH RTC RTCL.xlsx`
- `MOTIF TOTAL (1).xlsx`

Si aucun fichier n’est importé, l’application peut utiliser les fichiers présents dans le repository.

### 2. Filtres de travail
L’application applique :
- un **filtre journalier** basé sur la **colonne A**
- un **filtre produit** basé sur la colonne **s.produit**

Produits pris en charge :
- `FTTH`
- `FTTHDFO`
- `RTC`
- `RTCDTL`

### 3. Codes intervention
La colonne **État** du fichier source peut contenir :
- `NA`
- `RM`
- `TR`
- `TL`

Ces valeurs représentent des **codes d’intervention**, pas des noms d’agents :
- `NA` : nouvelle installation
- `RM` : remise en service
- `TR` : transfert
- `TL` : transfert local

### 4. Affectation à un vrai agent
Pour chaque ligne du tableau :
- l’administrateur choisit un **vrai agent**
- clique sur **Affecter et enregistrer l’agent choisi**
- puis peut envoyer la ligne complète à l’agent par **WhatsApp**

Le message WhatsApp contient la ligne complète :
- commande
- adresse
- contact
- secteur
- état
- produit
- autres champs disponibles

### 5. Retour d’intervention terrain
Après intervention, les informations fournies par l’agent peuvent être saisies dans l’application.

#### Pour `RTC` et `RTCDTL`
Champs disponibles :
- `SR`
- `TT`
- `PC`
- `Port`
- `Rosasse`
- `MSAN.port`
- `Câble` (`1/6` ou `5/9`)

#### Pour `FTTH` et `FTTHDFO`
Champs disponibles :
- `Numéro de validation`
- `MSAN.slot.port.sn`
- `Combien de mètre FTTH`
- `Autre consommable`

### 6. Envoi Outlook du retour terrain
Après la saisie du retour :
- l’application enregistre les données
- puis envoie le retour par **email Outlook**
- l’adresse Outlook est configurée par l’administrateur

### 7. Dashboard et rapports
L’application affiche :
- le nombre d’affectations
- le nombre d’envois WhatsApp
- le nombre de retours terrain
- le nombre d’emails envoyés

---

## Structure du projet

```text
app.py
README.md
requirements.txt
runtime.txt
ETAT FTTH RTC RTCL.xlsx
MOTIF TOTAL (1).xlsx
