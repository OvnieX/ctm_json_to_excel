# Control-M JSON to Excel

Outil de génération automatique de fichiers Excel à partir d'exports JSON Control-M.

Le projet contient :

- `ctm_json_to_excel.py` : script Python qui analyse un export JSON Control-M et génère trois classeurs Excel complémentaires ;
- `ctm_json_to_excel.yml` : pipeline Azure DevOps permettant d'exécuter le script de manière industrialisée à partir des JSON présents dans le repository Control-M.

> **Gestion des branches**
>
> - **GitHub** : les sources de référence de ce projet sont publiées sur la branche `main`.
> - **Azure DevOps / CONTROLM** : le pipeline opérationnel doit être exécuté uniquement depuis la branche `integration`.
>
> Ces deux règles sont volontaires : `main` sur GitHub sert de version de référence du code, tandis que l'exécution dans le repository Azure DevOps CONTROLM est protégée et réservée à `integration`.

---

## 1. Objectif

Les exports Control-M au format JSON contiennent de nombreuses informations techniques, mais ils sont peu pratiques à exploiter directement pour de la documentation ou pour construire des tableaux.

Ce projet automatise cette transformation afin de :

- reconstruire la hiérarchie réelle `Folder -> SubFolder -> Job` ;
- prendre en charge plusieurs structures d'exports JSON Control-M ;
- extraire les informations utiles à la documentation ;
- conserver les variables et les Events ;
- restituer les informations de planification et de calendrier ;
- récupérer la criticité depuis la variable Control-M `SEVERITE` ;
- restituer les paramètres des jobs cycliques ;
- générer automatiquement trois vues Excel adaptées à des usages différents ;
- permettre une utilisation centralisée via Azure DevOps sans exécution locale du script.

Le script ne déduit pas la hiérarchie à partir du nom des jobs. La hiérarchie utilisée est celle réellement décrite dans le JSON Control-M.

---

## 2. Fichiers du projet

```text
.
├── README.md
├── ctm_json_to_excel.py
└── ctm_json_to_excel.yml
```

### `ctm_json_to_excel.py`

Script principal Python.

Il charge le JSON, normalise les objets Control-M détectés, puis génère les trois classeurs Excel.

### `ctm_json_to_excel.yml`

Pipeline Azure DevOps utilisé dans le repository **CONTROLM**.

Il permet de :

1. contrôler que l'exécution est faite sur la branche `integration` ;
2. sélectionner la provenance du JSON ;
3. préparer Python 3.11 ;
4. installer `openpyxl` ;
5. rechercher le ou les fichiers JSON demandés ;
6. exécuter le script Python ;
7. publier les fichiers Excel dans un artifact Azure DevOps nommé `ControlM_Excel`.

---

## 3. Formats JSON Control-M pris en charge

Les exports Control-M rencontrés ne possèdent pas toujours la même structure.

Le parser prend notamment en charge deux formes principales.

### 3.1 Objets nommés directement par les clés JSON

Exemple :

```json
{
  "FOLDER_A": {
    "Type": "Folder",
    "SUB_A": {
      "Type": "SubFolder",
      "JOB_A": {
        "Type": "Job:Command"
      }
    }
  }
}
```

Dans ce cas, les noms `FOLDER_A`, `SUB_A` et `JOB_A` sont portés par les clés JSON.

### 3.2 Jobs stockés dans une liste `Jobs`

Exemple :

```json
{
  "FOLDER_A": {
    "Type": "Folder",
    "Jobs": [
      {
        "Type": "Job:Command",
        "Name": "JOB_A"
      }
    ]
  }
}
```

Dans ce cas, le nom du job est lu dans la propriété `Name`.

Le parser construit ensuite un modèle commun afin que la génération Excel soit indépendante de la structure initiale du JSON.

---

## 4. Gestion des occurrences multiples

Un même job peut apparaître plusieurs fois sous le même parent avec le même nom.

Le script ne fusionne pas ces occurrences afin de ne pas perdre d'information.

Chaque occurrence obtient un chemin technique unique avec un suffixe du type :

```text
__OCC_2
__OCC_3
```

Le nom Control-M affiché dans les fichiers Excel reste inchangé.

---

## 5. Informations extraites

Selon les informations disponibles dans le JSON, le script peut notamment restituer :

- Folder ;
- SubFolder ;
- Job ;
- type Control-M exact ;
- parent direct ;
- Folder parent ;
- SubFolder parent ;
- chemin hiérarchique complet ;
- Description ;
- Application ;
- SubApplication ;
- Control-M Server ;
- Host ;
- RunAs ;
- FileName ;
- FilePath ;
- Command ;
- Script ;
- ConnectionProfile ;
- DocumentationUrl ;
- Variables ;
- planification ;
- calendriers ;
- Events attendus ;
- Events ajoutés ;
- Events supprimés ;
- Events ajoutés conditionnellement ;
- criticité ;
- informations de cyclicité.

### Criticité

La colonne **Criticité** de la synthèse correspond strictement à la variable Control-M :

```text
Variables -> SEVERITE
```

`Priority`, `Severity`, `Criticality` ou d'autres propriétés ne sont pas utilisées comme substitut.

### Jobs cycliques

Le script ne se contente pas d'afficher `Oui` ou `Non`.

Lorsqu'un cycle est défini, il tente de restituer les paramètres utiles, par exemple :

```text
Toutes les 2 minutes | depuis le début du run | illimité
```

ou :

```text
Heures: 0900, 1100, 1710 | Tolérance: 20 min
```

Les formats pris en charge incluent notamment :

- `Rerun` ;
- `RerunIntervals` ;
- `RerunSpecificTimes` ;
- plusieurs anciens champs de cyclicité conservés pour compatibilité.

Pour un objet non cyclique, la cellule est laissée vide.

---

## 6. Les trois fichiers Excel générés

Pour un fichier source :

```text
export-controlm.json
```

le script produit :

```text
export-controlm_hierarchie.xlsx
export-controlm_ligne.xlsx
export-controlm_synthese.xlsx
```

### 6.1 `*_hierarchie.xlsx`

Il s'agit de la vue la plus détaillée.

Le classeur contient plusieurs feuilles :

| Feuille | Contenu |
|---|---|
| `Hierarchie` | Folder, SubFolder et Job dans l'ordre du JSON avec leurs principales propriétés |
| `Variables` | Variables Control-M par objet |
| `Events` | WaitFor, Add, Delete et Add conditionnels |
| `Planning` | Détail des propriétés du bloc `When` |
| `Proprietes` | Autres propriétés techniques de l'objet |

Cette vue est adaptée à l'analyse détaillée d'une chaîne.

### 6.2 `*_ligne.xlsx`

Vue technique à plat.

Elle contient une feuille `Tous_les_objets` avec une ligne par objet Control-M.

Les propriétés imbriquées sont transformées en colonnes et les variables deviennent par exemple :

```text
Variable.UG
Variable.SEVERITE
Variable.TASK
```

Les Events possèdent également leurs propres colonnes.

Cette vue est particulièrement utile pour :

- filtrer ;
- trier ;
- rechercher une valeur ;
- comparer plusieurs objets ;
- exploiter les données de manière technique.

### 6.3 `*_synthese.xlsx`

Vue volontairement plus courte et plus lisible, destinée notamment à la documentation.

Elle contient les principales colonnes :

```text
Niveau
Nom
Parent Folder
SubFolder
Type
Cyclique
Planification
Calendrier
Criticité
Run As
```

Les lignes sont également différenciées visuellement selon qu'elles représentent un Folder, un SubFolder ou un Job.

---

## 7. Prérequis pour une exécution locale

### Python

Python 3 est nécessaire.

Le pipeline Azure DevOps utilise actuellement :

```text
Python 3.11
```

### Dépendance Python

Le script utilise `openpyxl` pour produire les fichiers `.xlsx`.

Installation :

```bash
python -m pip install openpyxl
```

---

## 8. Utilisation locale

### Génération standard

```bash
python ctm_json_to_excel.py export.json
```

Les trois fichiers sont créés à côté du JSON source.

### Définir les noms de sortie

```bash
python ctm_json_to_excel.py export.json \
  --output export_hierarchie.xlsx \
  --flat-output export_ligne.xlsx \
  --summary-output export_synthese.xlsx
```

Sous PowerShell, la commande peut être écrite sur une seule ligne :

```powershell
python .\ctm_json_to_excel.py .\export.json --output .\export_hierarchie.xlsx --flat-output .\export_ligne.xlsx --summary-output .\export_synthese.xlsx
```

---

## 9. Utilisation dans Azure DevOps

Dans Azure DevOps, l'utilisation prévue est :

```text
CONTROLM
└── Pipelines
    └── EXCEL
        └── CTM_JSON_TO_EXCEL
```

Le pipeline est volontairement manuel :

```yaml
trigger: none
pr: none
```

L'utilisateur choisit la source JSON et le fichier à analyser depuis l'écran **Run pipeline**.

### Important : branche Azure DevOps

Le pipeline doit être lancé uniquement depuis :

```text
integration
```

Le YAML contient une protection explicite : si `Build.SourceBranch` n'est pas `refs/heads/integration`, l'exécution est arrêtée.

Il ne faut donc pas utiliser la branche `main` du repository Azure DevOps pour exécuter ce pipeline.

---

## 10. Paramètre `Source JSON Control-M`

Le pipeline propose les valeurs suivantes :

| Choix | Répertoire Azure DevOps |
|---|---|
| `DEV BUILD` | `build/dev` |
| `PPRD BUILD` | `build/pep` |
| `QAL BUILD` | `build/qaf` |
| `PROD BUILD` | `build/prd` |
| `DEV SMARTFOLDER` | `src/smartfolder/dev` |
| `PPRD SMARTFOLDER` | `src/smartfolder/pep` |
| `QAL SMARTFOLDER` | `src/smartfolder/qaf` |
| `PROD SMARTFOLDER` | `src/smartfolder/prd` |

### BUILD

Les répertoires `BUILD` contiennent des exports JSON Control-M organisés par environnement.

Ils sont adaptés lorsqu'on souhaite travailler sur un export ciblé ou sur un périmètre précis, sans traiter l'ensemble des SmartFolders disponibles.

Il suffit de sélectionner la source `BUILD` correspondant à l'environnement concerné puis d'indiquer le nom du fichier JSON ou un masque de fichiers.

### SMARTFOLDER

Les répertoires `SMARTFOLDER` contiennent les SmartFolders extraits automatiquement depuis Control-M et régulièrement mis à jour.

Ils sont particulièrement adaptés aux périmètres importants, lorsque les chaînes complètes deviennent trop volumineuses.

Dans ce cas, il est préférable de travailler SmartFolder par SmartFolder afin de conserver des tableaux et des représentations plus lisibles.

---

## 11. Paramètre `Fichier(s) JSON à analyser`

Le pipeline accepte un nom exact ou un masque.

### Un fichier précis

```text
export-controlm.json
```

### Plusieurs fichiers d'une même famille

```text
export-*.json
```

### Tous les JSON du répertoire sélectionné

```text
*.json
```

La recherche reste limitée au répertoire déterminé par `Source JSON Control-M`.

Le pipeline rejette également les chemins contenant `..` ou les chemins absolus afin d'éviter de sortir du répertoire sélectionné.

---

## 12. Emplacement du script dans Azure DevOps

Le YAML fourni utilise actuellement :

```yaml
pythonScript: 'cipipeline/ctm_to_excel/ctm_json_to_excel.py'
```

Le repository Azure DevOps doit donc contenir le script à cet emplacement, sauf modification volontaire de la variable `pythonScript`.

Exemple :

```text
CONTROLM/
├── build/
│   ├── dev/
│   ├── pep/
│   ├── qaf/
│   └── prd/
├── src/
│   └── smartfolder/
│       ├── dev/
│       ├── pep/
│       ├── qaf/
│       └── prd/
└── cipipeline/
    └── ctm_to_excel/
        ├── ctm_json_to_excel.py
        └── ctm_json_to_excel.yml
```

---

## 13. Déroulement du pipeline Azure DevOps

Le pipeline réalise les étapes suivantes :

1. vérification de la branche `integration` ;
2. checkout propre du repository ;
3. préparation de Python 3.11 ;
4. installation de `openpyxl` ;
5. résolution du répertoire source à partir du choix BUILD/SMARTFOLDER ;
6. validation du script, du répertoire et du masque JSON ;
7. recherche du ou des JSON à traiter ;
8. génération des trois fichiers Excel pour chaque JSON ;
9. publication des résultats dans Azure DevOps.

Les fichiers générés sont regroupés dans l'artifact :

```text
ControlM_Excel
```

---

## 14. Stratégie de branches : GitHub vs Azure DevOps

Il est important de distinguer les deux usages.

### GitHub

La branche de référence est :

```text
main
```

Elle contient la version publiée du script, du YAML et de cette documentation.

Exemple :

```text
GitHub
└── main
    ├── README.md
    ├── ctm_json_to_excel.py
    └── ctm_json_to_excel.yml
```

### Azure DevOps

La branche utilisée pour l'exécution est :

```text
integration
```

Exemple :

```text
Azure DevOps - CONTROLM
└── integration
    └── cipipeline/
        └── ctm_to_excel/
            ├── ctm_json_to_excel.py
            └── ctm_json_to_excel.yml
```

La présence des fichiers sur `main` dans GitHub ne signifie donc pas que le pipeline Azure DevOps doit être exécuté sur `main`.

**Règle à retenir :**

```text
GitHub       -> main        = code de référence / publication
Azure DevOps -> integration = exécution du pipeline CONTROLM
```

Lorsqu'une nouvelle version est validée sur GitHub, il faut reporter la version correspondante dans la branche `integration` du repository Azure DevOps avant de l'utiliser avec le pipeline.

---

# GitHub

Le dépôt GitHub est publié sur :

```text
main
```

Pour une modification classique :

```bash
git status
git add .
git commit -m "Description de la modification"
git push
```

La branche `integration` mentionnée dans le YAML concerne l'exécution du pipeline **Azure DevOps** et ne modifie pas le fonctionnement du dépôt GitHub.

---

## 18. Maintenance et remontée de bugs

Les exports Control-M pouvant varier selon les types de jobs et les versions, de nouveaux cas particuliers peuvent apparaître.

En cas d'anomalie, fournir si possible :

- le type de source (`BUILD` ou `SMARTFOLDER`) ;
- l'environnement (`DEV`, `PPRD`, `QAL`, `PROD`) ;
- le nom du JSON concerné ;
- le type de job concerné ;
- la valeur présente dans le JSON ;
- la valeur obtenue dans Excel ;
- la valeur attendue.

Cela facilite la reproduction et la correction du problème sans modifier le comportement des cas déjà supportés.

---

## 19. Résumé rapide

```text
Entrée
  Export JSON Control-M
        |
        v
ctm_json_to_excel.py
        |
        +--> *_hierarchie.xlsx
        +--> *_ligne.xlsx
        +--> *_synthese.xlsx

GitHub       : branche main
Azure DevOps : branche integration
Pipeline     : CONTROLM / EXCEL / CTM_JSON_TO_EXCEL
Artifact     : ControlM_Excel
```

