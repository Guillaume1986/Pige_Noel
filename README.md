# 🎄 Pige de Noël

Site familial de pige de Noël : chacun entre son code, inscrit trois idées de cadeaux, et la pige se fait
automatiquement quand tout le monde a répondu. Chacun reçoit alors un courriel avec sa personne pigée.

- **Site** (ce dépôt) : GitHub Pages — `index.html` + `config.js`
- **Moteur** : Google Apps Script + Google Sheets (`apps-script/Code.gs`) — fait la pige et envoie les courriels

Règles : on ne se pige pas soi-même, les conjoints ne se pigent pas entre eux,
et personne ne pige la même personne que l’an dernier.

---

## 1. Mettre le site sur GitHub (5 min)
1. Crée un dépôt (ex. `pige-noel`) et téléverse le contenu de ce dossier.
2. **Settings > Pages** : *Source* = `Deploy from a branch`, branche `main`, dossier `/ (root)`.
3. Après une minute, le site est à `https://TON-COMPTE.github.io/pige-noel/`.
   Tant que `config.js` est vide, il fonctionne en **mode démo** (codes `DEMO1` et `DEMO2`).

## 2. Installer le moteur (10 min)
1. Crée une feuille Google Sheets, puis **Extensions > Apps Script**.
2. Remplace le contenu de `Code.gs` par `apps-script/Code.gs`.
3. En haut du script, remplis `CONFIG` : titre, budget, date de l’échange et **`urlSite`** (l’adresse GitHub Pages).
4. **Déployer > Nouveau déploiement > Application Web** :
   *Exécuter en tant que* **Moi**, *Qui a accès* **Tout le monde**. Copie l’URL (se termine par `/exec`).
5. Sur GitHub, colle cette URL dans `config.js` (`API_URL`) et enregistre.

## 3. Préparer la famille
Dans la feuille, recharge la page : un menu **🎁 Pige de Noël** apparaît.
1. **Préparer les onglets**, puis remplis *Nom*, *Courriel* et *Conjoint(e)* dans **Participants**
   (le nom du conjoint doit être identique à la colonne *Nom*; un seul côté du couple suffit).
2. **Générer les codes**.
3. Première année : inscris la pige de l’an dernier dans **Historique** (Année, Donneur, Receveur).
4. **Vérifier qu’une pige est possible**, puis **Envoyer les invitations** : chacun reçoit un lien personnel.

## Pendant et après
- La guirlande s’allume à mesure que les gens répondent. La page indique aussi qui n’a pas encore répondu.
- Au dernier inscrit, la pige part toute seule et les courriels sont envoyés.
- Tu reçois une confirmation **sans les jumelages**, car la pige est gardée hors de la feuille.
- Quelqu’un tarde? **Forcer la pige maintenant**. Un courriel s’est perdu? **Renvoyer les courriels**.
- L’an prochain : **Préparer une nouvelle année** (archive la pige, efface les idées, garde les codes).

## Bon à savoir
- Après toute modification de `Code.gs` : **Déployer > Gérer les déploiements > Modifier > Nouvelle version**.
  L’URL ne change pas.
- Le dépôt peut être public : il ne contient ni noms ni codes. Les données sont dans ta feuille Google.
  La page demande aux moteurs de recherche de ne pas l’indexer.
- Les idées de chacun sont visibles dans la feuille. Seule la pige est cachée.
