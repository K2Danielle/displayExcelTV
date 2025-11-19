
# 📊 Excel TV Display

> Application web pour afficher des plannings Excel sur écran TV avec navigation par semaine et rafraîchissement automatique

[![Python](https://img.shields.io/badge/Python-3.11+-blue.svg)](https://www.python.org/downloads/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.109.0-green.svg)](https://fastapi.tiangolo.com/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

![Excel TV Display](screenshot.png)

## 🎯 Fonctionnalités

- ✅ **Conversion Excel vers HTML** avec préservation des styles (couleurs, fusion de cellules)
- ✅ **Navigation par semaine** avec boutons Précédent/Suivant
- ✅ **Zoom ajustable** (50% à 200%) pour adaptation aux différentes tailles d'écran
- ✅ **Sélection automatique** de la semaine en cours
- ✅ **Rafraîchissement automatique** toutes les 15 minutes
- ✅ **Détection des modifications** du fichier Excel en temps réel (WebSocket)
- ✅ **Upload de fichiers** via interface web ou URL
- ✅ **Limitation intelligente** : colonnes A à M, lignes 1 à 24
- ✅ **Formatage des dates** en français (lundi 17 novembre 2025)
- ✅ **Interface responsive** adaptée aux écrans TV et tablettes

## 📸 Captures d'écran

### Interface principale
![Planning](docs/planning-view.png)

### Page d'upload
![Upload](docs/upload-view.png)

## 🚀 Installation Rapide

### Prérequis
- Python 3.11 ou supérieur
- Windows 10/11, macOS, ou Linux

### Installation

```bash
# 1. Cloner le projet
git clone https://github.com/votre-username/excel-tv-display.git
cd excel-tv-display

# 2. Créer un environnement virtuel (optionnel mais recommandé)
python -m venv venv
source venv/bin/activate  # Sur Windows: venv\Scripts\activate

# 3. Installer les dépendances
pip install -r requirements.txt

# 4. Créer le dossier uploads
mkdir uploads

# 5. Lancer l'application
python main.py
```

L'application sera accessible sur : **http://localhost:8001**

## 📦 Dépendances

```
fastapi==0.109.0
uvicorn[standard]==0.27.0
python-multipart==0.0.6
aiofiles==23.2.1
openpyxl==3.1.2
watchdog==3.0.0
```

## 🎮 Utilisation

### 1. Démarrer le serveur

**Windows :**
```cmd
start.bat
```

**Linux/Mac :**
```bash
python main.py
```

### 2. Charger un fichier Excel

**Option A : Via l'interface web**
1. Ouvrir http://localhost:8001
2. Cliquer sur "Charger un fichier"
3. Sélectionner votre fichier .xlsx

**Option B : Via URL**
1. Entrer l'URL du fichier Excel
2. Cliquer sur "Charger depuis URL"

**Option C : Copie directe**
```bash
cp votre-planning.xlsx uploads/
```
Le fichier sera détecté automatiquement.

### 3. Navigation

- **Dropdown** : Sélectionner une semaine spécifique
- **◄ Précédent** : Semaine précédente
- **Suivant ►** : Semaine suivante
- **Zoom** : Boutons - / + / Reset pour ajuster la taille

## 📐 Structure du Projet

```
excel-tv-display/
│
├── main.py                 # Application principale FastAPI
├── requirements.txt        # Dépendances Python
├── start.bat              # Script de lancement Windows
├── README.md              # Ce fichier
│
├── uploads/               # Dossier des fichiers Excel
│   └── planning.xlsx      # Exemple de fichier
│
└── docs/                  # Documentation et captures d'écran
    ├── installation.md
    └── configuration.md
```

## ⚙️ Configuration

### Modifier le port

Dans `main.py`, ligne finale :
```python
uvicorn.run(app, host="0.0.0.0", port=8001)  # Changer 8001
```

Ou via ligne de commande :
```bash
python -m uvicorn main:app --host 0.0.0.0 --port 8002
```

### Modifier l'intervalle de rafraîchissement

Dans `main.py` :
```python
REFRESH_INTERVAL = 15 * 60  # 15 minutes en secondes
```

### Formats de fichiers supportés

```python
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".csv"}
```

## 🌐 Accès Réseau

### Trouver l'IP du serveur

**Windows :**
```cmd
ipconfig
```

**Linux/Mac :**
```bash
ifconfig
```

Chercher "Adresse IPv4" (exemple : 192.168.1.100)

### Accéder depuis un autre appareil

1. Connecter l'appareil au même réseau Wi-Fi
2. Ouvrir un navigateur
3. Aller sur : `http://192.168.1.100:8001`

### Configuration du pare-feu (Windows)

```cmd
netsh advfirewall firewall add rule name="Excel TV Display" dir=in action=allow protocol=TCP localport=8001
```

## 🔧 API Endpoints

### Endpoints disponibles

| Endpoint | Méthode | Description |
|----------|---------|-------------|
| `/` | GET | Page principale avec planning |
| `/upload` | POST | Upload d'un fichier Excel |
| `/sheets` | GET | Liste des feuilles disponibles |
| `/sheet/{name}` | GET | Contenu d'une feuille spécifique |
| `/files` | GET | Liste des fichiers uploadés |
| `/file-info` | GET | Informations sur le fichier actuel |
| `/status` | GET | Statut du serveur |
| `/ws` | WebSocket | Connexion temps réel |

### Exemple d'utilisation de l'API

```python
import requests

# Récupérer les feuilles disponibles
response = requests.get('http://localhost:8001/sheets')
sheets = response.json()['sheets']

# Récupérer une feuille spécifique
response = requests.get('http://localhost:8001/sheet/47')
html_content = response.json()['html']
```

## 🎨 Personnalisation des Styles

### Modifier les couleurs de l'interface

Dans la fonction `root()` de `main.py`, section `<style>` :

```css
.header { background: #0a0a0a; }  /* En-tête */
.controls { background: #2d2d2d; }  /* Barre de contrôles */
body { background: #1a1a1a; }  /* Arrière-plan */
```

### Modifier le style de la colonne A

Dans `main.py`, fonction `sheet_to_html()` :

```python
style += "; background-color: #4472C4; color: #FFFFFF"  # Bleu par défaut
```

## 📱 Mode TV/Plein Écran

### Activation automatique du mode plein écran

1. Ouvrir le navigateur sur la TV
2. Aller sur : `http://[IP]:8001`
3. Appuyer sur **F11** pour le mode plein écran
4. Le planning s'affiche en grand format

### Empêcher la mise en veille

**Windows :**
- Paramètres → Système → Alimentation
- "Mettre en veille" : Jamais

**Linux :**
```bash
sudo systemctl mask sleep.target suspend.target
```

## 🚀 Déploiement en Production

### Lancement automatique au démarrage (Windows)

**Option 1 : Dossier de démarrage**
```cmd
Win + R → shell:startup
```
Créer un raccourci vers `start.bat`

**Option 2 : Service Windows avec NSSM**
```cmd
nssm install ExcelTV "C:\ExcelTV\venv\Scripts\python.exe" "C:\ExcelTV\main.py"
nssm start ExcelTV
```

### Déploiement avec Docker (optionnel)

```dockerfile
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8001

CMD ["python", "main.py"]
```

```bash
docker build -t excel-tv-display .
docker run -d -p 8001:8001 -v $(pwd)/uploads:/app/uploads excel-tv-display
```

## 🐛 Dépannage

### Problème : "Python n'est pas reconnu"

**Solution :**
Réinstaller Python en cochant "Add Python to PATH"

### Problème : "Port 8001 déjà utilisé"

**Solution :**
```bash
# Trouver le processus
netstat -ano | findstr :8001

# Tuer le processus (Windows)
taskkill /PID <PID> /F

# Ou utiliser un autre port
python -m uvicorn main:app --port 8002
```

### Problème : "Module not found"

**Solution :**
```bash
pip install -r requirements.txt --force-reinstall
```

### Problème : "Les dates ne s'affichent pas"

**Solution :**
1. Mettre de vraies dates dans Excel (pas de formules)
2. Ou convertir les formules en valeurs :
   - Sélectionner → Copier → Collage spécial → Valeurs

### Problème : "Pas d'accès depuis un autre PC"

**Solutions :**
1. Vérifier le pare-feu
2. Vérifier que les appareils sont sur le même réseau
3. Tester avec l'IP locale : `http://192.168.x.x:8001`

## 📊 Format Excel Recommandé

### Structure du fichier Excel

```
Ligne 1 : En-tête (Ressources, Équipes...)
Ligne 2-4 : Sous-en-têtes
Ligne 5 : Lundi [date]
Ligne 8 : Mardi [date]
Ligne 13 : Mercredi [date]
Ligne 17 : Jeudi [date]
Ligne 20 : Vendredi [date]
Ligne 23 : Samedi [date]

Colonnes : A (Dates) → M (max)
```

### Format des dates

- **Colonne A** : Dates au format `lundi 17 novembre 2025`
- **Type** : Cellules avec vraies dates Excel (pas de formules)
- **Limite** : Lundi à Samedi (6 jours)


## 👨‍💻 Auteur
K2Danielle


## 🙏 Remerciements

- [FastAPI](https://fastapi.tiangolo.com/) - Framework web moderne
- [openpyxl](https://openpyxl.readthedocs.io/) - Manipulation de fichiers Excel
- [Uvicorn](https://www.uvicorn.org/) - Serveur ASGI performant
- [Watchdog](https://pythonhosted.org/watchdog/) - Surveillance des fichiers

## ⭐ Support

Si ce projet vous a été utile, n'hésitez pas à lui donner une étoile ⭐ !

---

**Made with ❤️ for easy planning display**
