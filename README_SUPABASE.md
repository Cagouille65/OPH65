# OPH65 V2.14 Cloud — Supabase + Excel synchronisé

Cette version utilise **Supabase comme base persistante** et conserve en parallèle un **classeur Excel/XLSM synchronisé**.

## 1. Créer le projet Supabase

1. Créez un projet sur Supabase.
2. Ouvrez **SQL Editor**.
3. Créez une nouvelle requête, copiez tout le contenu de `SUPABASE_SETUP.sql`, puis exécutez-la.
4. Dans le dialogue **Connect** / la page des clés API, relevez :
   - l'URL du projet ;
   - une **Secret key** serveur (`sb_secret_...`).

> Ne placez jamais cette clé dans GitHub. Elle donne un accès privilégié à la base.

## 2. Mettre à jour les Secrets Streamlit

Dans Streamlit Community Cloud : **App > Settings > Secrets**, copiez le modèle de `.streamlit/secrets.toml.example` et remplacez les valeurs.

Points importants :

- `app.public_base_url` = l'URL exacte de votre application Streamlit, sans slash final ;
- `app.access_password` = mot de passe protégeant l'interface interne OPH65 ;
- `app.rts_email` = adresse qui recevra automatiquement les validations de fin de prestation ;
- `supabase.url` et `supabase.secret_key` = paramètres Supabase ;
- la section `[smtp]` reste votre configuration d'envoi de mails.

Enregistrez les Secrets puis redémarrez l'application.

## 3. Premier chargement du fichier Excel

Dans OPH65 : **Import / Export > Importer le fichier Excel / XLSM de suivi**.

Au premier import, l'application :

1. lit le classeur ;
2. enregistre le suivi dans Supabase ;
3. conserve une copie compressée du classeur XLSM dans Supabase ;
4. utilise le référentiel des logements fourni avec l'application et l'initialise dans Supabase si la table est vide.

Après cela, le suivi est automatiquement rechargé depuis Supabase lorsque Streamlit redémarre.

## 4. Synchronisation Excel

Quand un logement est modifié dans **Ajouter / Modifier** :

- les données sont enregistrées dans Supabase ;
- le classeur XLSM stocké dans Supabase est mis à jour ;
- les macros VBA sont conservées pour un fichier `.xlsm` ;
- le fichier à jour peut être récupéré dans **Import / Export > Télécharger le classeur Excel synchronisé**.

La validation d'un BC par QR code renseigne également la date réelle de fin dans Supabase **et dans le classeur Excel synchronisé**.

> Une application Web ne peut pas modifier directement le fichier qui se trouve sur votre PC. Il faut télécharger la dernière version synchronisée depuis l'interface lorsque vous souhaitez la récupérer localement.

## 5. Gestion des BC et QR codes

L'onglet **Gestion des BC** liste tous les bons de commande **En cours** ou **En retard**, classés par N° LOT.

Pour un BC :

1. sélectionnez le BC ;
2. cliquez sur **Préparer le mail et le QR code** ;
3. vérifiez/modifiez le destinataire, l'objet et le corps du message ;
4. envoyez le mail.

Objet par défaut : `BC [N° du bon de commande]`.

Le message reprend le texte demandé et le QR code est intégré dans le corps du mail.

Lorsque l'entreprise scanne le QR code, elle accède uniquement à une page publique minimale de validation du BC. Elle ne peut pas accéder au tableau de bord. Une validation :

- marque le BC comme terminé à la date du jour ;
- met à jour Supabase ;
- met à jour le classeur Excel synchronisé ;
- envoie automatiquement une information au Responsable Technique de Secteur configuré dans les Secrets / Paramètres.

## 6. Sécurité

L'URL Streamlit peut rester publique pour que les QR codes fonctionnent, mais l'interface de gestion est protégée par `app.access_password`.

Les secrets SMTP et Supabase doivent rester exclusivement dans **Streamlit Secrets**. Ne téléversez jamais un vrai fichier `secrets.toml` sur GitHub.
