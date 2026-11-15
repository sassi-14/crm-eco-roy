# CRM Eco-Roy

Appli interne (PWA) pour gérer tout le cycle d'un dossier PAC : lead → éligibilité → OPX/MPR → commande → pose → photos et satisfaction → encaissements.
Elle marche sur téléphone et ordi, même sans réseau, et se synchronise toute seule.

## Comment c'est construit

| Brique | Où | Rôle |
|---|---|---|
| Appli (`index.html`, `sw.js`, `manifest.webmanifest`) | GitHub Pages (gratuit) | Ce que toi et les poseurs utilisez |
| Serveur (`apps-script/Code.gs`) | Ton compte Google (Apps Script) | Synchro, stockage, bots |
| Base de données | Google Sheet « CRM Eco-Roy — base » | 1 ligne = 1 dossier (lisible dans Sheets) |
| Photos, signatures | Dossier Drive « CRM Eco-Roy — Photos chantiers » | 1 sous-dossier par client |

Aucune donnée client n'est sur GitHub : le code y est public, les données restent dans ton Google.

## Installation (une seule fois, ~15 min)

1. **Serveur** : ouvre https://script.google.com → Nouveau projet → nomme-le « CRM Eco-Roy ».
   Colle le contenu de `apps-script/Code.gs`. Dans Paramètres du projet, coche « Afficher le fichier manifeste » et colle `apps-script/appsscript.json`.
2. Choisis la fonction `setup` → **Exécuter** → accepte les autorisations Google.
   Ça crée la base, le dossier photos, tes codes d'accès et les 3 bots.
3. **Déployer → Nouveau déploiement → Application Web**
   - Exécuter en tant que : **Moi**
   - Qui a accès : **Tout le monde**
   Copie l'URL qui finit par `/exec`.
4. Choisis la fonction `migrationExistant` → **Exécuter** : tes Sheets actuels (call Maroc, Suivi clients, compte PAC 1€, igor install) sont importés. Chaque dossier importé porte le badge « à vérifier ».
5. **Appli** : ouvre l'adresse GitHub Pages, colle l'URL `/exec`, puis ton **code gérant** (onglet « Config » de la base).
   Sur iPhone : Partager → « Sur l'écran d'accueil ». Sur Android/ordi : « Installer l'application ».

### Poseurs
Réglages → « Copier le lien d'accès poseur » → envoie le lien + son PIN (onglet Config, lignes `POSEUR_Igor`, `POSEUR_Kevin`).
Pour ajouter un poseur : ajoute une ligne `POSEUR_Prenom` avec un PIN à 6 chiffres.
Un poseur voit uniquement ses chantiers (pose programmée, puis 30 jours après la pose) : adresse, téléphone, matériel, infos techniques, photos avant/après, notes, fiche satisfaction signée. **Jamais** les prix, marges, revenus ou pièces fiscales.

## Les bots

| Bot | Quand | Ce qu'il fait |
|---|---|---|
| `botImportLeads` | toutes les 15 min | Lit les Sheets du call Maroc, crée les nouveaux leads (dédoublonnage nom + CP / téléphone), t'envoie une alerte mail |
| `botRecapMatin` | 7h45 | Mail récap : leads à rappeler, poses J-3 (matériel commandé ?), pièces manquantes, OPX sans retour > 10 j, chantiers sans photos/satisfaction, MPR à déposer, soldes > 30 j, marges faibles |
| `botRelancesClients` | 10h (désactivé par défaut) | Relance mail les clients à qui il manque l'avis d'impôt ou la CNI (tous les 3 j, 4 fois max). Active-le avec `RELANCE_CLIENTS = oui` dans Config |
| Formulaire du site | en direct | `action: "lead"` : crée le lead, calcule la couleur MPR et un score /100, t'alerte (🔥 si chaud) |

Les réglages (marge mini, e-mail des alertes, signature des mails) sont dans l'onglet **Config** de la base.

## Mettre à jour l'appli
Modifie les fichiers sur GitHub : l'appli se met à jour d'elle-même à la prochaine ouverture avec réseau.
Pour le serveur : colle le nouveau `Code.gs`, puis Déployer → Gérer les déploiements → modifier → Nouvelle version (l'URL ne change pas).

## Barèmes
Plafonds MaPrimeRénov' 2026 (Anah, janvier 2026) et aide PAC air/eau par geste (bleu 5 000 €, jaune 4 000 €, violet 3 000 €) dans `index.html` (`PLAFONDS`, `AIDE_MPR_PAC`) et `Code.gs` (`PLAFONDS`). À mettre à jour quand l'Anah publie les nouveaux barèmes.
