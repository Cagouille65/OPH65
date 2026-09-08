# OPH65 — Version spéciale Streamlit Community Cloud

Cette archive est prête à être déposée dans un dépôt GitHub puis déployée sur Streamlit Community Cloud.

## 1. Fichiers à déposer sur GitHub

Décompressez l'archive puis déposez **le contenu du dossier** à la racine du dépôt GitHub. Les fichiers indispensables sont :

- `app.py`
- `requirements.txt`
- `logo.jpg`
- `oph65_logements.json`
- `.streamlit/config.toml`
- `.gitignore`

Le fichier `.streamlit/secrets.toml.example` est seulement un modèle. Il ne contient aucun vrai secret.

## 2. Déployer dans Streamlit Community Cloud

1. Créez un dépôt GitHub, de préférence **Private**.
2. Dans le dépôt : `Add file` > `Upload files` et chargez le contenu décompressé.
3. Ouvrez Streamlit Community Cloud et créez une nouvelle application.
4. Sélectionnez votre dépôt GitHub.
5. Main file path : `app.py`.
6. Avant ou juste après le déploiement, ouvrez `Settings` > `Secrets`.
7. Copiez le modèle ci-dessous et remplacez les valeurs nécessaires.

```toml
[app]
cloud_mode = true
public_base_url = "https://VOTRE-APP.streamlit.app"

[smtp]
server = "smtp.gmail.com"
port = 465
use_ssl = true
auth_required = true
use_starttls = false
email = "VOTRE_ADRESSE@EXEMPLE.FR"
password = "VOTRE_MOT_DE_PASSE_APPLICATION"
default_recipients = ""
```

**Ne mettez jamais le mot de passe SMTP dans GitHub.** Il doit uniquement être saisi dans `Settings > Secrets` de Streamlit.

## 3. Utilisation du classeur Excel dans le Cloud

Un serveur Streamlit Cloud ne peut pas lire un chemin Windows du type `C:\\...`. Utilisez l'onglet **Import / Export** de l'application pour charger le classeur Excel/XLSM depuis votre ordinateur.

Le référentiel de logements `oph65_logements.json` est fourni dans le dépôt et peut donc être utilisé dès le démarrage.

## 4. Point important sur la persistance

Streamlit Community Cloud peut redémarrer l'application et son disque local n'est pas conçu comme une base de données permanente. Les modifications faites uniquement dans des fichiers locaux (`oph65_config.json`, référentiel logements modifié depuis l'interface, etc.) peuvent donc être perdues lors d'un redéploiement ou d'un redémarrage.

Pour une utilisation durable avec les futurs QR codes et la validation des BC par les entreprises, il faudra connecter l'application à une base persistante externe (par exemple Supabase/PostgreSQL). La présente version est prête pour le déploiement et les tests Web, mais le stockage permanent des validations BC devra être ajouté dans la version « Gestion des BC ».

## 5. Sécurité

- Gardez le dépôt GitHub privé si le référentiel logements contient des données internes.
- Le mot de passe SMTP doit rester dans Streamlit Secrets.
- Lorsque la page QR publique sera créée, elle devra utiliser un jeton sécurisé et ne jamais donner accès au tableau de bord complet.
