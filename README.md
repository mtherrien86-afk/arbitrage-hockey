# Arbitrage 2025-2026

Le tableau de ton fichier Excel, en site web :

- hébergé gratuitement sur **GitHub Pages**;
- **enregistré automatiquement** dans un **Google Sheet** (environ 1 seconde après chaque modification);
- **rechargé automatiquement** (toutes les 30 secondes, et dès que tu reviens sur l’onglet), donc ton téléphone et ton ordinateur restent synchronisés;
- **tarifs et arénas configurables** dans *Réglages* : les prix, les coûts de déplacement et les noms se modifient sans toucher au code.

Les calculs sont ceux de l’Excel : Prix (liste de prix) + Déplacement (coût de l’aréna si « Déplacement à payer » est coché) + Supplément = Total.

## Installation (environ 15 minutes)

### 1. Créer le Google Sheet
1. Va sur <https://sheets.google.com> et crée un **classeur vide** (nomme-le « Arbitrage 2025-2026 »).
2. *Fichier > Paramètres* : mets le fuseau horaire à **Montréal** (ou Toronto).

### 2. Ajouter le script
1. Dans le Sheet : *Extensions > Apps Script*.
2. Efface tout, puis colle le contenu de `apps-script/Code.gs`.
3. En haut du fichier, remplace `CHANGE-MOI` par **un mot de passe long** de ton choix. Note-le, tu le taperas une fois par appareil.
4. Clique sur l’icône de disquette (Enregistrer).

### 3. Déployer le script
1. *Déployer > Nouveau déploiement*, roue dentée > **Application Web**.
2. *Exécuter en tant que* : **Moi**. *Qui a accès* : **Tout le monde**.
3. *Déployer*, puis *Autoriser l’accès*. Google affiche « Google n’a pas validé cette application » : c’est normal, c’est ton propre script. Clique sur *Paramètres avancés > Accéder à … (non sécurisé) > Autoriser*.
4. Copie l’**URL de l’application Web** (elle se termine par `/exec`).

### 4. Publier le site sur GitHub
1. Crée un dépôt GitHub, puis ajoute-y `index.html`, `style.css`, `app.js` (et `.gitignore`, `README.md`).
2. *Settings > Pages* : *Source* = **Deploy from a branch**, branche **main**, dossier **/ (root)**.
3. Après une minute, le site est à `https://TON-NOM.github.io/NOM-DU-DEPOT/`.

### 5. Relier le site au Google Sheet
1. Ouvre le site : les *Réglages* s’ouvrent sur l’onglet **Connexion**.
2. Colle l’URL `/exec` et ton mot de passe, puis **Enregistrer et tester**.
3. *Réglages > Données > Importer un fichier (.json)* : choisis `donnees-depart.json` (tes 70 parties, 44 tarifs et 4 arénas de l’Excel).
   Vérifie que le total affiche **2 289,83 $**.
4. Ouvre le Google Sheet : les trois onglets (`Feuille 1`, `Liste de prix`, `Liste des arénas`) sont remplis, avec les mêmes formules que l’Excel.

### 6. Autres appareils
*Réglages > Connexion > Copier le lien pour un autre appareil*, puis ouvre ce lien sur ton téléphone. Il configure l’appareil et efface le mot de passe de l’adresse.
Tu peux aussi l’ajouter à l’écran d’accueil depuis le navigateur.

## Bon à savoir

- **Si deux appareils modifient en même temps**, ou si tu modifies le Sheet à la main pendant que le site est ouvert, le site ne remplace rien en silence : une bannière te demande quelle version garder.
- **Hors ligne**, tes modifications sont gardées sur l’appareil et envoyées dès que le réseau revient.
- **Modifier un tarif change aussi les anciennes parties** qui l’utilisent (comme les `RECHERCHEV` de l’Excel). Renommer une aréna ou un tarif met à jour les parties concernées.
- **Dans le Google Sheet**, tu peux corriger des valeurs à la main, mais ne renomme pas les onglets et ne déplace pas les colonnes. Les colonnes *Prix*, *Déplacement* et *Total* sont des formules, réécrites à chaque enregistrement du site.
- **Si tu modifies `Code.gs`** : *Déployer > Gérer les déploiements >* crayon *> Version : Nouvelle version*. L’URL ne change pas.

## Sécurité

- Le dépôt GitHub ne contient **que du code**, aucune donnée. Ne mets jamais `donnees-depart.json` ni une copie exportée dans le dépôt (le `.gitignore` t’aide).
- L’URL du script est publique en théorie, mais elle refuse toute requête sans ton mot de passe. Garde le mot de passe et le « lien pour un autre appareil » pour toi.
- Il n’y a pas de comptes d’utilisateurs : qui a l’URL et le mot de passe peut lire et modifier les données.
