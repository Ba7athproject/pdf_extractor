# Ba7ath PDF Extract OCR

Un outil moderne d'extraction de texte et de tableaux à partir de fichiers PDF, doté d'une interface graphique sombre élégante (PyQt5) et d'une accélération GPU (via OpenCV CUDA).

## 🚀 Fonctionnalités

*   **Extraction de texte multilingue** : Utilise `pytesseract` pour l'OCR avec support du français, anglais et arabe (`fra+eng+ara`).
*   **Accélération GPU** : Pré-traitement des images (conversion en niveaux de gris et seuillage binaire) via OpenCV CUDA pour de meilleures performances (si disponible).
*   **Extraction de tableaux** : Détecte et extrait les données tabulaires (séparées par `|`) de l'OCR.
*   **Traitement par lots** : Possibilité de spécifier des plages de pages précises (ex: `1,3-4,8`) et de définir la taille des lots pour ne pas surcharger la mémoire.
*   **Exportation** :
    *   Export du texte brut au format **DOCX**.
    *   Export des tableaux détectés au format **XLSX**.
*   **Interface Moderne** : Interface PyQt5 "Darkly Modern" avec barre de progression et recherche intégrée.
*   **Support Windows** : Ajoute dynamiquement les chemins CUDA nécessaires sous Windows pour s'assurer que les DLL sont bien chargées.

## 📋 Prérequis

Pour utiliser cette application, plusieurs dépendances systèmes et Python sont requises.

### Dépendances Système

1.  **Tesseract OCR** : Doit être installé et configuré (avec les packs de langues FRA, ENG, ARA).
2.  **Poppler** : Requis par `pdf2image` pour convertir les pages PDF en images. (Sous Windows, téléchargez les binaires et ajoutez le dossier `bin` à votre PATH ou spécifiez-le dans le code).
3.  **CUDA Toolkit** (Optionnel mais recommandé) : Pour utiliser l'accélération OpenCV GPU.

### Dépendances Python

Les bibliothèques suivantes sont nécessaires (peuvent être installées dans l'environnement virtuel `venv37` inclus) :

```bash
pip install PyQt5 pdf2image pytesseract clipboard python-docx openpyxl Pillow numpy
```

*(Note : Pour le support GPU, une version compilée d'OpenCV avec CUDA est nécessaire, par exemple le fichier `opencv_contrib_python_rolling-4.13.0.20250812-cp37-abi3-win_amd64.whl` présent dans le projet).*

## 🛠️ Installation et Utilisation

1.  **Cloner le dépôt / Ouvrir le dossier** :
    Placez-vous dans `C:\Ba7ath_scripts\pdf_extract`.
2.  **Activer l'environnement virtuel** :
    ```bash
    # Sous Windows
    venv37\Scripts\activate
    ```
3.  **Vérifier le support GPU** (Optionnel) :
    Vous pouvez tester si OpenCV détecte bien votre GPU en exécutant :
    ```bash
    python test_GPU.py
    ```
    Cela devrait afficher la version d'OpenCV et le nombre de périphériques CUDA détectés.
4.  **Lancer l'application** :
    ```bash
    python pdf_extract.py
    ```

## 🖥️ Comment l'utiliser ?

1.  Cliquez sur **📄 Sélectionner PDF** pour charger votre document.
2.  Indiquez les **Pages à extraire** (laissez vide pour tout traiter).
3.  Ajustez la **Taille de lot** selon la puissance de votre machine (par défaut : 5).
4.  Cliquez sur **🚦 Lancer OCR**.
5.  Une fois terminé, utilisez la recherche intégrée ou exportez les résultats via **💾 Exporter en DOCX** ou **📊 Exporter tableaux en XLSX**.
