# pronote-heures-perdues

Outil **non officiel** pour les parents d'élèves : il relève dans Pronote (espace Parents) les heures de cours
non assurées depuis la rentrée, produit une analyse Excel et prépare la déclaration sur l'outil de la FCPE,
[Ouyapacours](https://ouyapacours.fcpe.asso.fr/).

> Ce projet n'est ni édité ni approuvé par Index Éducation (éditeur de Pronote), la FCPE ou l'Éducation nationale.
> Les noms cités appartiennent à leurs propriétaires.

## Ce que fait l'outil

- Ouvre Edge ou Chrome. **Vous** vous connectez vous-même (identifiants ou ENT), puis il parcourt les semaines.
- Prend une capture de l'emploi du temps de chaque semaine.
- Classe les cours : professeur absent, annulé, grève, déplacé ou modifié (rattrapage), remplacé, maintenu…
- Produit un classeur Excel :

| Onglet | Contenu |
|---|---|
| Synthèse | heures prévues, perdues, rattrapées, solde et pourcentages |
| Saisie FCPE | les lignes à saisir dans Ouyapacours (matière, date, durée) |
| Justificatifs FCPE | la liste cours par cours, avec les captures |
| Par matière | totaux et graphiques |
| Par semaine | les mêmes totaux, semaine par semaine |
| Détail | un cours par ligne, catégorie modifiable |
| Captures | une miniature par semaine |

- Prépare des captures justificatives sans le bandeau portant les noms, avec les cours déclarés entourés en rouge.

## Installation

1. Téléchargez `pronote_suivi-0.9.zip` depuis la page
   [Releases](https://github.com/IGNF-Xavier/pronote-heures-perdues/releases) et extrayez-le.
2. Ouvrez **`LISEZMOI.html`** : le mode d'emploi pas à pas, écrit pour une personne qui ne connaît pas
   l'informatique (ouvrir l'invite de commandes, installer Python, régler, lancer, lire le résultat).

En bref, sous Windows avec Python 3.10 à 3.14 et Edge ou Chrome, dans le dossier `pronote_suivi` :

```bat
python -m pip install -r requirements.txt
```

```bat
python pronote_suivi.py
```

ou double-clic sur `lancer.bat`. L'outil demande l'adresse de Pronote (la première fois), la date de rentrée
(par défaut le 31/08/2026) et la dernière semaine à analyser (par défaut la semaine en cours).

## Réglages

Tout se trouve dans [`parametres.ini`](parametres.ini) : adresse de Pronote, dates, dossier de sortie, matières
exclues des totaux (par défaut « Vie de classe »), champs et motifs de la déclaration FCPE (par défaut
« Prof absent »).

Pour refaire le classeur à partir des semaines déjà capturées, sans se reconnecter :

```bat
python pronote_suivi.py --etape analyse
```

## Confidentialité

- L'outil ne lit ni n'enregistre votre mot de passe.
- Les données restent sur votre ordinateur.
- Les dossiers de sortie contiennent le nom de l'enfant, ses enseignants, les captures et, dans
  `profil_navigateur`, la **session de connexion**. Ne les partagez pas et ne les publiez pas.
- Si vous contribuez au dépôt, ne commitez jamais votre `parametres.ini` rempli.

## État et limites (version 0.9)

- Préversion, mise au point sur un seul établissement (collège, espace Parents).
- L'outil lit la page de Pronote : une autre version ou un autre affichage peut demander des ajustements.
- Les statuts sont reconnus par mots-clés. Une grève n'est identifiée que si Pronote l'écrit sur le cours.
- Vérifiez les chiffres avec les captures avant toute déclaration.
- Respectez les conditions d'utilisation de votre établissement et de Pronote. L'outil ne fait que consulter,
  avec votre propre compte, ce que vous pouvez déjà voir.

Les retours et corrections sont bienvenus dans les
[issues](https://github.com/IGNF-Xavier/pronote-heures-perdues/issues), **sans capture ni donnée personnelle**.

## Licence

[GNU AGPL v3](LICENSE).
