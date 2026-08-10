# Achats — Générateur de contrats fournisseurs (solution SPFx)

Application interne pour centraliser les fichiers de conditions commerciales,
les templates de contrats Word, et générer des contrats remplis
automatiquement à partir d'un code sous-segment.

L'application est livrée comme une **solution SharePoint Framework (SPFx)**,
hébergée nativement dans SharePoint (voir `spfx/`) : un unique web part avec
une navigation interne à 3 onglets, correspondant aux 3 modules de la
spécification. Aucun backend séparé n'est nécessaire — le traitement des
fichiers Excel/Word s'exécute entièrement côté client, et le stockage repose
sur des listes SharePoint (métadonnées) et une bibliothèque de documents
(fichiers bruts). L'authentification est déjà gérée par SharePoint/Azure AD.

👉 Voir `spfx/README.md` pour le détail de l'architecture, le modèle de
stockage SharePoint et les instructions de build/déploiement.

## Modules

1. **Conditions commerciales** — dépôt et historisation des fichiers Excel de
   conditions commerciales fournisseurs.
2. **Templates de contrats** — dépôt et historisation des templates Word
   (`.docx`, variables au format `{{Variable}}`), avec mapping
   variable ↔ colonne du fichier de conditions actif.
3. **Générer un contrat** — sélection template + fichier de conditions +
   code sous-segment → formulaire pré-rempli et éditable → génération du
   `.docx` final.
