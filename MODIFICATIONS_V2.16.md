# OPH65 V2.16

- Zone de sélection de l'utilisateur dans la barre latérale : fond gris clair forcé et texte foncé pour rester lisible quel que soit le thème Streamlit.
- Validation QR BC renforcée : recherche sur N° LOT + corps d'état + N° BC.
- La validation renseigne automatiquement la date réelle de fin de la prestation dans `suivi_logements` (Supabase).
- Contrôle de lecture après écriture Supabase avant confirmation de la validation.
- Mise à jour automatique de la copie XLSM synchronisée après validation QR.
- Rechargement du suivi depuis Supabase à chaque rerun de l'interface Cloud afin de refléter les validations faites depuis le téléphone d'une entreprise.
