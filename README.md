# Control-M JSON to Excel

Outil de génération automatique de fichiers Excel à partir d'exports JSON Control-M.

Le projet contient :

- `ctm_json_to_excel.py` : script Python qui analyse un export JSON Control-M et génère trois classeurs Excel complémentaires et deux fichiers de bilan ;
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
- permettre une utilisation centralisée via Azure DevOps sans exécution locale du script ;
- traiter plusieurs JSON successivement, avec une source commune ;
- consolider les classeurs si demandé, sans fusionner les JSON ;
- signaler les définitions incomplètes et produire un bilan de génération.

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

Il charge un seul JSON par appel, normalise les objets Control-M détectés, puis génère les trois classeurs Excel et les bilans JSON/TXT. La fonction de consolidation est appelée séparément par la pipeline sur les classeurs déjà produits.

### `ctm_json_to_excel.yml`

Pipeline Azure DevOps utilisé dans le repository **CONTROLM**.

Il permet de :

1. contrôler que l'exécution est faite sur la branche `integration` ;
2. sélectionner une source commune et un ou plusieurs JSON ;
3. préparer Python 3.11 et installer `openpyxl` ;
4. valider les noms, les masques et l'existence de tous les fichiers ;
5. dédoublonner la sélection et nettoyer une seule fois le dossier de sortie ;
6. convertir les JSON successivement, avec arrêt au premier échec ;
7. consolider les classeurs si `mergeExcel` est activé ;
8. produire le bilan global et publier l'artifact `ControlM_Excel` en cas de succès.

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
        "Type": "Job:Command",
        "Command": "echo test"
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
        "Name": "JOB_A",
        "Command": "echo test"
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

La première occurrence conserve son chemin de base. Les occurrences suivantes reçoivent un suffixe technique du type :

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

Pour un objet non cyclique, la cellule est laissée vide. `RerunLimit` seul ne déclenche pas la détection d’un cycle.

### Planification et Events

La synthèse lit le premier bloc dictionnaire parmi `When`, `Scheduling` et
`Schedule`. La feuille détaillée `Planning` lit uniquement `When`.
`USE PARENT` est affiché lorsque le calendrier est hérité sans calendrier
explicite ; aucune valeur de planning ou de calendrier n'est calculée depuis
le parent. Les calendriers exclus sont préfixés par `EXCL:`.

Les ajouts conditionnels d'Events sont extraits pour les jobs à partir des
dictionnaires d'actions imbriqués. Les listes d'actions ne sont pas parcourues
par cet extracteur dédié.

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

Cette vue est adaptée à l'analyse détaillée d'une chaîne. Les colonnes `Command`, `Script`, `FileName` et `FilePath` restent séparées dans `Hierarchie`. La colonne commune « Commande / Script » appartient uniquement à la synthèse.

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

La feuille `Synthese` contient les colonnes suivantes :

| Colonne | Intitulé |
|---|---|
| A | Niveau |
| B | Nom |
| C | Parent Folder |
| D | SubFolder |
| E | Type |
| F | Commande / Script |
| G | Cyclique |
| H | Planification |
| I | Calendrier |
| J | Criticité |
| K | Host |
| L | Run As |
| M | Contrôle définition |

Les lignes distinguent visuellement Folder, SubFolder et Job. Les feuilles
possèdent des filtres, des en-têtes figés et des cellules avec retour à la ligne.

#### Commande / Script

La feuille `Synthese` possède la colonne **Commande / Script** en **F**, juste
après `Type`.

- `Job:Command` : uniquement `Command`, même si un ancien `FileName` est présent.
- `Job:Script` : uniquement le chemin `FilePath` / `FileName`, sans commande résiduelle.
- `Job:EmbeddedScript` : contenu `Script` et nom de fichier éventuel. C'est le seul
  cas permettant d'afficher le contenu et le fichier ensemble.
- Pour les autres types, une seule valeur est retenue : commande, sinon contenu
  du script, sinon fichier.
- Pour `Job:Command` et `Job:Script`, si le champ attendu est absent ou vide,
  la cellule reste vide. Pour `Job:EmbeddedScript`, un fichier éventuel reste
  affiché même si le contenu manque ; le contrôle de définition signale alors
  le script manquant.
- Un script sous forme de liste est affiché ligne par ligne.

Dans la synthèse consolidée, la colonne reste immédiatement après `Type`
(en G, car `Source JSON` est ajouté en première colonne).

Les colonnes existantes `Command`, `Script`, `FileName` et `FilePath` restent
inchangées. Aucun script externe n'est ouvert et les variables Control-M ne sont
pas résolues. Un chemin de script ne fournit donc pas le contenu du fichier.
La nouvelle colonne est conservée dans la synthèse consolidée.
La vue hiérarchique conserve ses colonnes d’origine, sans colonne commune supplémentaire.

#### Host et contrôles de définition

`Host` est en colonne K, juste avant `Run As` (L). Il reprend uniquement la
valeur présente sur l'objet, sans rechercher de valeur héritée. La colonne
`Contrôle définition` (M) reste vide pour un objet sans anomalie détectée.
Elle signale, sur fond orangé et en texte rouge :

| Type | Définition incomplète détectée |
|---|---|
| `Job:Command` | `Command` absent, vide ou composé d'espaces |
| `Job:Script` | `FileName` absent, vide ou composé d'espaces |
| `Job:EmbeddedScript` | `Script` absent, vide ou sans ligne non blanche |

Ces contrôles sont non bloquants. Les mêmes avertissements sont inscrits dans
les logs et les bilans, avec le chemin technique et le type du job.
Un ancien fichier résiduel ne remplace pas une commande manquante ; un nom de
fichier ne remplace pas le contenu d'un script embarqué.
Les autres types, notamment les jobs de plugins et `Job:Dummy`, ne sont pas
soumis à ces trois règles. Ce contrôle ciblé ne remplace pas une validation
complète par Control-M. Aucun contrôle d'existence du script sur le host n'est effectué.

#### Lecture des commandes et scripts longs

La colonne F utilise une police monospace Consolas 10 et conserve les retours
à la ligne. La hauteur de chaque ligne est estimée d'après son contenu et les
largeurs des colonnes, entre 18 et 300 points (environ 20 lignes visibles).
Ce plafond évite qu'un script occupe toute la feuille. Au-delà, utiliser la
barre de formule développée ou augmenter manuellement la hauteur dans Excel.
Le texte n'est pas raccourci par cette mise en forme ; la limite Excel de
32 767 caractères par cellule reste applicable. Les mêmes règles s'appliquent
à la synthèse consolidée, ainsi que la coloration des avertissements.


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

Les trois classeurs et les bilans `<export>_generation-info.json` et
`<export>_generation-info.txt` sont créés à côté du JSON source.
Le Python accepte toujours un seul JSON par appel ; les listes sont gérées
par la pipeline.

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

Les dossiers de destination doivent exister. Par défaut, les bilans sont écrits
dans le dossier de la synthèse. L'option `--report-output chemin/bilan.json`
permet de choisir leur emplacement ; `chemin/bilan.txt` est écrit à côté.
Ce chemin doit être distinct du JSON source et des classeurs.
Les JSON UTF-8 avec ou sans BOM sont acceptés.

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

L'utilisateur choisit la source commune, la sélection `jsonFile` et l'option
`mergeExcel` depuis l'écran **Run pipeline**. Les valeurs par défaut sont
`DEV BUILD`, `*.json` et `false` respectivement.

### Important : branche Azure DevOps

Le pipeline doit être lancé uniquement depuis :

```text
integration
```

Le YAML contient une protection explicite : si `Build.SourceBranch` n'est pas `refs/heads/integration`, l'exécution est arrêtée.

Il ne faut donc pas utiliser la branche `main` du repository Azure DevOps pour exécuter ce pipeline.

---

## 10. Paramètre `Source JSON Control-M`

Le paramètre YAML `sourceJsonControlM`, affiché « Source JSON Control-M »,
propose les valeurs suivantes :

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

## 11. Paramètre `jsonFile` : sélection multi-JSON

Dans **Run pipeline**, `jsonFile` accepte un nom, un masque ou une liste séparée
par des virgules :

```text
EXPORT_A.json, EXPORT_B.json, EXPORT_C.json
```

Les usages existants restent valides : `EXPORT_A.json`, `EXPORT_*.json`, `*.json`.
Une liste peut également mélanger noms et masques :

```text
EXPORT_B.json, EXPORT_*.json
```

Règles :

1. Retirer les espaces autour de chaque entrée, sans toucher aux espaces internes.
2. Valider chaque nom/masque et l'existence de tous les fichiers avant toute conversion.
3. Refuser les entrées vides, chemins absolus, séparateurs `/` ou `\`, `:`,
   `..`, caractères de contrôle, liens symboliques et fichiers non JSON.
4. Chercher uniquement dans le répertoire sélectionné, sans récursion.
5. Développer chaque masque dans l'ordre alphabétique des noms ; un nom exact
   existant est prioritaire sur son interprétation comme masque.
6. Dédoublonner les noms de fichiers en conservant leur première apparition.
7. Nettoyer le dossier de sortie une seule fois, après validation de toute la sélection.
8. Appeler successivement le convertisseur, avec un seul JSON par appel.
9. Afficher `JSON 1/3`, `JSON 2/3`, etc., après dédoublonnage.
10. S'arrêter au premier échec, sans traiter les fichiers suivants.

Une virgule sépare les entrées et ne peut donc pas faire partie d'un nom.
Un masque sans correspondance fait échouer la validation. Les paramètres sont
transmis par variables d'environnement et les appels Python par liste d'arguments,
sans interprétation des noms de fichiers comme commandes shell.
La validation préalable porte sur les noms et fichiers ; le contenu JSON est
analysé lors de chaque conversion. Si celle-ci échoue, les sorties déjà produites
peuvent rester dans le dossier de travail, mais l'artifact n'est pas publié.

Tous les JSON utilisent la même source et les mêmes paramètres. Ils ne sont pas
fusionnés et aucune dépendance entre fichiers distincts n’est reconstruite.

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
3. préparation de Python 3.11 et installation d'`openpyxl` ;
4. résolution de la source et validation de toute la sélection ;
5. dédoublonnage, nettoyage unique et création des sous-dossiers par nom complet ;
6. conversions successives avec affichage `JSON 1/3`, `JSON 2/3`, etc. ;
7. consolidation optionnelle après le succès de tous les exports ;
8. écriture et affichage du bilan global ;
9. publication des classeurs et bilans si le run réussit.

La pipeline utilise Bash sur `ubuntu-latest`. Une erreur de sélection, de
conversion ou de consolidation empêche la publication. Les avertissements de
définition sont non bloquants et figurent dans les résultats publiés.

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

## 15. Sorties et consolidation optionnelle

Le dossier `$(Build.ArtifactStagingDirectory)/controlm-excel` contient un
sous-dossier par nom complet, extension `.json` comprise :

```text
controlm-excel/
├── generation-info.txt
├── EXPORT_A.json/
│   ├── EXPORT_A_hierarchie.xlsx
│   ├── EXPORT_A_ligne.xlsx
│   ├── EXPORT_A_synthese.xlsx
│   ├── EXPORT_A_generation-info.json
│   └── EXPORT_A_generation-info.txt
├── EXPORT_B.json/
│   ├── EXPORT_B_hierarchie.xlsx
│   ├── EXPORT_B_ligne.xlsx
│   ├── EXPORT_B_synthese.xlsx
│   ├── EXPORT_B_generation-info.json
│   └── EXPORT_B_generation-info.txt
└── CONSOLIDE/                       # seulement si mergeExcel = true
    ├── consolide_hierarchie.xlsx
    ├── consolide_ligne.xlsx
    └── consolide_synthese.xlsx
```

Le paramètre booléen **`mergeExcel`**, affiché « Consolider les Excel de tous les
JSON », est **désactivé par défaut**. Activé, il produit trois classeurs
supplémentaires après le succès de toutes les conversions. Les classeurs
individuels sont toujours conservés, même avec un seul export sélectionné.

La consolidation concatène les lignes des feuilles de même nom. Elle conserve
les cinq feuilles du classeur hiérarchique et la feuille de chacune des deux
autres vues. Chaque feuille consolidée commence par **Source JSON**.
Les objets de même nom provenant de tickets distincts restent des lignes
séparées. Utiliser `Source JSON` avec `Chemin` / `Chemin complet` pour les distinguer.

Les colonnes sont alignées par intitulé et réunies dans leur ordre de première
apparition. Une propriété absente d'un export donne une cellule vide, sans
perdre les propriétés propres à un autre export. L'ordre des exports et des
lignes est conservé. Les filtres, couleurs et largeurs sont repris ; les volets
sont figés en B2 pour garder la source visible.

Aucun JSON n'est fusionné ; aucun lien entre exports n'est reconstruit.
La consolidation lit uniquement les XLSX produits par le run courant.
Elle échoue explicitement si un classeur est absent, si ses feuilles/en-têtes
sont incompatibles ou si les limites Excel de lignes/colonnes sont dépassées.
L'intitulé `Source JSON` est réservé dans les classeurs à consolider.

---

## 16. Bilans de génération

Chaque conversion écrit deux fichiers, même si une erreur survient pendant
la lecture du JSON ou la génération (à condition que le dossier du bilan soit
accessible en écriture) :

- `<export>_generation-info.json` : bilan structuré exploité par la pipeline.
- `<export>_generation-info.txt` : bilan lisible pour l'exploitation.

Ils contiennent le nom du JSON, le statut `OK`, `OK_AVEC_AVERTISSEMENTS` ou
`ECHEC`, les dates de début/fin en UTC, la durée, les nombres de Folders,
SubFolders et Jobs, les définitions incomplètes et les fichiers générés.
En cas d'échec, l'erreur et les sorties déjà produites sont indiquées.
Les compteurs portent sur les occurrences analysées, pas sur les seuls noms uniques.

Par défaut, les bilans sont écrits dans le dossier de la synthèse.
Pour personnaliser le chemin, utiliser `--report-output chemin/bilan.json` ;
le fichier `chemin/bilan.txt` est créé à côté. Le dossier doit exister.
Ce chemin doit être distinct de l'entrée JSON et des trois classeurs.

La pipeline écrit également **`generation-info.txt`** à la racine de l'artifact.
Ce bilan global contient la source choisie, la sélection, les conversions
réussies, les compteurs totaux, les avertissements par export et les chemins
des classeurs individuels et consolidés.

Si une conversion échoue, il distingue les JSON réussis, celui en échec et ceux
non traités. Une erreur de consolidation est également signalée. Le bilan est
affiché dans les logs et reste sur l'agent ; la publication de `ControlM_Excel`
reste conditionnée au succès du run. Les avertissements de définition ne font
pas échouer le run et sont donc publiés avec les classeurs.
Une erreur de sélection avant le nettoyage est signalée dans les logs ; aucun
nouveau bilan n'est écrit dans le dossier de sortie à ce stade.

---

## 17. Maintenance et remontée de bugs

Les exports Control-M pouvant varier selon les types de jobs et les versions, de nouveaux cas particuliers peuvent apparaître.

En cas d'anomalie, fournir si possible :

- le type de source (`BUILD` ou `SMARTFOLDER`) ;
- l'environnement (`DEV`, `PPRD`, `QAL`, `PROD`) ;
- la sélection `jsonFile` et la valeur de `mergeExcel` ;
- le nom du JSON concerné et les bilans de génération ;
- le type de job concerné ;
- la valeur présente dans le JSON ;
- la valeur obtenue dans Excel ;
- la valeur attendue.

Cela facilite la reproduction et la correction du problème sans modifier le comportement des cas déjà supportés.

---

## 18. Résumé rapide

| Élément | Fonctionnement |
|---|---|
| Entrée locale | Un seul JSON par appel Python |
| Entrée Azure DevOps | Nom, masque ou liste séparée par des virgules |
| Traitement | Successif, dédoublonné, arrêt au premier échec |
| Sorties par JSON | Trois XLSX et deux bilans JSON/TXT |
| Dossier par export | Nom complet avec extension `.json` |
| Consolidation | `mergeExcel`, désactivé par défaut ; trois XLSX supplémentaires |
| Synthèse | Commande / Script en F, Host en K, Run As en L, contrôle en M |
| Bilan global | `generation-info.txt` |
| GitHub | `main` |
| Azure DevOps | `integration` |
| Pipeline | `CONTROLM / EXCEL / CTM_JSON_TO_EXCEL` |
| Artifact | `ControlM_Excel` |
