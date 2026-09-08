# Modifications V2.14

- Ajout de Supabase comme base de données persistante.
- Persistance du suivi des logements, du référentiel logements, des paramètres, des validations BC et de l'historique des mails.
- Conservation et synchronisation du classeur XLSM dans Supabase.
- Téléchargement du classeur XLSM synchronisé depuis Import / Export.
- Ajout de l'onglet Gestion des BC : BC en cours / en retard classés par N° LOT.
- Génération d'un mail BC avec objet `BC [N° BC]`, corps demandé et QR code intégré.
- Page publique de validation QR sans accès au tableau de bord.
- Validation QR = date réelle de fin mise à jour dans Supabase et dans le classeur Excel synchronisé.
- Notification automatique au Responsable Technique de Secteur.
- Protection par mot de passe de l'interface interne Streamlit Cloud.
- Compatibilité avec les nouvelles Secret keys Supabase (`sb_secret_...`) et la clé legacy `service_role`.
