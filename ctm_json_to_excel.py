#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Convertit un JSON Control-M en trois classeurs Excel et deux bilans.

Usage : python ctm_json_to_excel.py export.json
Dépendance externe : openpyxl. Options et formats détaillés dans README.md.

Les objets nommés par clés JSON et les listes Jobs alimentent le même modèle
Folder/SubFolder/Job. La hiérarchie vient du JSON ; les occurrences sont gardées.
Les vues produites sont hiérarchique, à plat et synthétique.

La synthèse conserve Commande / Script après Type, Host, Run As et les contrôles
non bloquants de définition. La criticité provient uniquement de SEVERITE.
Les paramètres cycliques sont détaillés et aucun planning hérité n'est calculé.

La pipeline gère la sélection multi-JSON et appelle ce convertisseur une fois
par export. merge_excel_exports concatène ensuite les classeurs, si demandé,
avec une colonne Source JSON ; elle ne reconstruit pas de liens entre exports.
Les bilans JSON/TXT décrivent les compteurs, avertissements, sorties et erreurs.
"""

import argparse
import json
import math
import ntpath
import re
import time
from collections import Counter, defaultdict
from copy import copy
from datetime import datetime, timezone
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit(
        "Le module 'openpyxl' n'est pas installé.\n"
        "Installe-le avec : python -m pip install openpyxl"
    )


# ---------------------------------------------------------------------------
# 1. CONSTANTES DE PARSING
# ---------------------------------------------------------------------------
# Folder et SubFolder ont des Type exacts. Les jobs sont reconnus plus loin
# avec startswith("Job") car Control-M utilise de nombreux sous-types :
# Job:Command, Job:Script, Job:EmbeddedScript, Job:VMware By Broadcom, etc.
OBJECT_TYPES = {"Folder", "SubFolder"}
SKIP_RECURSIVE_KEYS = {
    "eventsToAdd", "eventsToWaitFor", "eventsToDelete",
}


# ---------------------------------------------------------------------------
# 2. FONCTIONS GENERIQUES D'EXTRACTION
# ---------------------------------------------------------------------------

def is_controlm_object(node):
    """Reconnaît les conteneurs et les types de jobs pris en charge."""
    if not isinstance(node, dict):
        return False
    obj_type = str(node.get("Type", ""))
    return obj_type in OBJECT_TYPES or obj_type.startswith("Job")


def flatten_value(value):
    """Convertit listes/dicts en texte lisible pour Excel."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (str, int, float)):
        return str(value)
    if isinstance(value, list):
        simple = []
        for item in value:
            if isinstance(item, (str, int, float, bool)) or item is None:
                simple.append(flatten_value(item))
            else:
                simple.append(json.dumps(item, ensure_ascii=False, separators=(",", ":")))
        return " | ".join(simple)
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def extract_variables(node):
    """Accepte une liste de dictionnaires ou un dictionnaire de variables."""
    out = []
    vars_block = node.get("Variables")
    if isinstance(vars_block, list):
        for item in vars_block:
            if isinstance(item, dict):
                for k, v in item.items():
                    out.append((str(k), flatten_value(v)))
    elif isinstance(vars_block, dict):
        for k, v in vars_block.items():
            out.append((str(k), flatten_value(v)))
    return out


def extract_events(block):
    """Extrait les noms des Events du bloc, dans leur ordre d’origine."""
    events = []
    if isinstance(block, dict):
        items = block.get("Events", [])
        if isinstance(items, list):
            for item in items:
                if isinstance(item, dict) and item.get("Event") is not None:
                    events.append(str(item.get("Event")))
    return events


def collect_conditional_event_adds(node, current_status=None, result=None):
    """Recherche les Event:Add et leur statut dans les dictionnaires d’actions.

    Les listes d’actions et les objets Control-M enfants ne sont pas parcourus.
    """
    if result is None:
        result = []
    if not isinstance(node, dict):
        return result

    local_status = current_status
    if node.get("Type") == "If:CompletionStatus":
        local_status = node.get("CompletionStatus", "")

    if node.get("Type") == "Event:Add" and node.get("Event"):
        result.append((str(node.get("Event")), str(local_status or "")))

    for value in node.values():
        if not isinstance(value, dict):
            continue
        if is_controlm_object(value):
            continue
        collect_conditional_event_adds(value, local_status, result)
    return result


def planning_pairs(when):
    """Aplatit récursivement When en paires clé/valeur."""
    rows = []
    if not isinstance(when, dict):
        return rows

    def rec(value, prefix=""):
        if isinstance(value, dict):
            for k, v in value.items():
                key = f"{prefix}.{k}" if prefix else str(k)
                rec(v, key)
        elif isinstance(value, list):
            rows.append((prefix, " | ".join(flatten_value(x) for x in value)))
        else:
            rows.append((prefix, flatten_value(value)))

    rec(when)
    return rows


def planning_summary(when):
    pairs = planning_pairs(when)
    return "; ".join(f"{k}={v}" for k, v in pairs if v != "")


# ---------------------------------------------------------------------------
# 3. PARSEUR CONTROL-M : TRANSFORMER LES DEUX FORMATS JSON EN UN MODELE UNIQUE
# ---------------------------------------------------------------------------

class ControlMParser:
    """Normalise les objets et conserve leur parent, ordre et occurrence."""

    def __init__(self):
        self.objects = []
        self._path_counts = defaultdict(int)

    def add_object(self, node, name, parent_path, depth, occurrence=1, occurrence_total=1):
        base_path = f"{parent_path}/{name}" if parent_path else name
        self._path_counts[base_path] += 1
        if self._path_counts[base_path] > 1:
            path = f"{base_path}__OCC_{self._path_counts[base_path]}"
        else:
            path = base_path

        item = {
            "path": path,
            "parent_path": parent_path or "",
            "depth": depth,
            "name": name,
            "type": str(node.get("Type", "")),
            "node": node,
            "occurrence": occurrence,
            "occurrence_total": occurrence_total,
        }
        self.objects.append(item)
        return path

    def walk(self, node, name=None, parent_path=None, depth=0):
        """
        Parcours récursif générique du JSON.

        Trois cas sont traités :
        - une LISTE : typiquement le tableau Jobs ;
        - un DICTIONNAIRE représentant directement un objet Control-M ;
        - un DICTIONNAIRE conteneur qui ne possède pas lui-même de Type.

        parent_path mémorise le parent hiérarchique réel et depth le niveau
        d'indentation qui sera ensuite utilisé dans Excel.
        """
        # CAS A : liste JSON. On repère d'abord les objets Control-M afin de
        # compter les doublons de (Type, Name) avant de les enregistrer.
        if isinstance(node, list):
            object_items = []
            for i, item in enumerate(node):
                if not is_controlm_object(item):
                    continue
                item_name = (
                    item.get("Name") or item.get("FolderName") or item.get("SubFolderName")
                    or f"{name or 'ITEM'}_{i + 1}"
                )
                object_items.append((i, item, str(item_name)))

            counts = Counter((str(item.get("Type", "")), item_name) for _, item, item_name in object_items)
            seen = defaultdict(int)

            for i, item in enumerate(node):
                if is_controlm_object(item):
                    item_name = (
                        item.get("Name") or item.get("FolderName") or item.get("SubFolderName")
                        or f"{name or 'ITEM'}_{i + 1}"
                    )
                    item_name = str(item_name)
                    key = (str(item.get("Type", "")), item_name)
                    seen[key] += 1
                    self.walk_object(
                        item, item_name, parent_path, depth,
                        occurrence=seen[key], occurrence_total=counts[key]
                    )
                elif isinstance(item, (dict, list)):
                    self.walk(item, name, parent_path, depth)
            return

        # Une valeur scalaire n'a rien à apporter à la hiérarchie.
        if not isinstance(node, dict):
            return

        # CAS B : le dictionnaire EST lui-même un Folder/SubFolder/Job.
        if is_controlm_object(node):
            object_name = name or node.get("Name") or node.get("FolderName") or node.get("SubFolderName")
            if not object_name:
                object_name = node.get("FileName") or "OBJECT"
            self.walk_object(node, str(object_name), parent_path, depth)
            return

        # CAS C : conteneur sans Type, fréquent à la racine du fichier.
        # On descend dans chaque clé et la clé sert de nom potentiel à l'objet.
        for key, value in node.items():
            if isinstance(value, dict):
                self.walk(value, str(key), parent_path, depth)
            elif isinstance(value, list):
                self.walk(value, str(key), parent_path, depth)

    def walk_object(self, node, name, parent_path, depth, occurrence=1, occurrence_total=1):
        """Enregistre l'objet courant puis cherche uniquement ses enfants Control-M."""
        # Le chemin enregistré devient le parent des enfants directs.
        current_path = self.add_object(
            node, name, parent_path, depth,
            occurrence=occurrence, occurrence_total=occurrence_total
        )

        child_depth = depth + 1
        # Variables/When/Events restent des propriétés, pas des objets enfants.
        for key, value in node.items():
            if key in SKIP_RECURSIVE_KEYS:
                continue
            if key in {"Variables", "When", "DocumentationUrl", "Arguments"}:
                continue

            if isinstance(value, dict):
                if is_controlm_object(value):
                    self.walk(value, str(key), current_path, child_depth)
                else:
                    # Les blocs d'actions/notifications ne sont pas des enfants hiérarchiques.
                    continue
            elif isinstance(value, list):
                if key == "Jobs":
                    self.walk(value, key, current_path, child_depth)
                else:
                    # Certaines variantes peuvent contenir des objets Control-M dans une autre liste.
                    if any(is_controlm_object(x) for x in value if isinstance(x, dict)):
                        self.walk(value, key, current_path, child_depth)


# ---------------------------------------------------------------------------
# 4. OUTILS DE HIERARCHIE ET DE PRESENTATION
# ---------------------------------------------------------------------------

def obj_category(obj_type):
    if obj_type == "Folder":
        return "Folder"
    if obj_type == "SubFolder":
        return "SubFolder"
    if obj_type.startswith("Job"):
        return "Job"
    return obj_type


def direct_parent_name(path):
    if not path or "/" not in path:
        return ""
    parent = path.rsplit("/", 1)[0]
    name = parent.rsplit("/", 1)[-1]
    return re.sub(r"__OCC_\d+$", "", name)


def nearest_ancestor_name(objects_by_path, parent_path, wanted_type):
    path = parent_path
    while path:
        obj = objects_by_path.get(path)
        if not obj:
            break
        if obj["type"] == wanted_type:
            return obj["name"]
        path = obj["parent_path"]
    return ""


def simple_extra_properties(node):
    excluded = {
        "Type", "Name", "FolderName", "SubFolderName", "Variables", "When",
        "eventsToAdd", "eventsToWaitFor", "eventsToDelete", "Jobs",
    }
    rows = []
    for key, value in node.items():
        if key.startswith("_ctm_") or key in excluded:
            continue
        if isinstance(value, dict):
            # Les objets Control-M enfants ne sont pas des propriétés du parent.
            if is_controlm_object(value):
                continue
            # Les actions/notifications sont conservées dans Propriétés sous forme JSON compacte.
            rows.append((str(key), flatten_value(value)))
        elif isinstance(value, list):
            if any(is_controlm_object(x) for x in value if isinstance(x, dict)):
                continue
            rows.append((str(key), flatten_value(value)))
        else:
            rows.append((str(key), flatten_value(value)))
    return rows


def style_sheet(ws, freeze="A2", filter_row=1):
    """Applique les en-têtes, filtres, volets et largeurs communes aux vues."""
    ws.freeze_panes = freeze
    ws.auto_filter.ref = ws.dimensions

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")

    for cell in ws[filter_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # Largeurs raisonnables, sans auto-fit extrême sur scripts/JSON.
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            text = "" if cell.value is None else str(cell.value)
            widths[cell.column] = min(max(widths.get(cell.column, 0), len(text) + 2), 55)
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = max(10, width)


def add_hierarchy_colors(ws, type_col=3):
    """Colore les données par catégorie ou type exact, sans toucher à l’en-tête."""
    fills = {
        "Folder": PatternFill("solid", fgColor="D9EAF7"),
        "SubFolder": PatternFill("solid", fgColor="E2F0D9"),
        "Job": PatternFill("solid", fgColor="FFF2CC"),
    }
    for r in range(2, ws.max_row + 1):
        category = obj_category(str(ws.cell(r, type_col).value or ""))
        fill = fills.get(category)
        if fill:
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = fill


# ---------------------------------------------------------------------------
# 5. EXCEL N°1 : VUE HIERARCHIQUE DETAILLEE
# ---------------------------------------------------------------------------

def build_excel(parser, output_path, source_name):
    """Crée la vue hiérarchique et ses quatre feuilles de détail."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Hierarchie"

    objects_by_path = {o["path"]: o for o in parser.objects}

    headers = [
        "Niveau", "Nom hiérarchique", "Catégorie", "Type Control-M",
        "Parent direct", "Parent Folder", "SubFolder", "Chemin complet",
        "Occurrence", "Description", "Application", "SubApplication",
        "ControlmServer", "Host", "RunAs", "FileName", "FilePath",
        "Command", "Script", "ConnectionProfile", "Planification",
        "Variables", "Events attendus", "Events ajoutés", "Events supprimés",
        "DocumentationUrl",
    ]
    ws.append(headers)

    for obj in parser.objects:
        node = obj["node"]
        category = obj_category(obj["type"])
        parent_folder = nearest_ancestor_name(objects_by_path, obj["parent_path"], "Folder")
        subfolder = nearest_ancestor_name(objects_by_path, obj["parent_path"], "SubFolder")
        if category == "Folder":
            parent_folder = ""
            subfolder = ""
        elif category == "SubFolder":
            subfolder = obj["name"]

        vars_text = "; ".join(f"{k}={v}" for k, v in extract_variables(node))
        docs = node.get("DocumentationUrl", "")
        if isinstance(docs, dict):
            docs = docs.get("Url", flatten_value(docs))

        occurrence = ""
        if obj["occurrence_total"] > 1:
            occurrence = f"{obj['occurrence']}/{obj['occurrence_total']}"

        row = [
            obj["depth"],
            ("    " * obj["depth"]) + obj["name"],
            category,
            obj["type"],
            direct_parent_name(obj["path"]),
            parent_folder,
            subfolder,
            obj["path"],
            occurrence,
            flatten_value(node.get("Description")),
            flatten_value(node.get("Application")),
            flatten_value(node.get("SubApplication")),
            flatten_value(node.get("ControlmServer")),
            flatten_value(node.get("Host")),
            flatten_value(node.get("RunAs")),
            flatten_value(node.get("FileName")),
            flatten_value(node.get("FilePath")),
            flatten_value(node.get("Command")),
            flatten_value(node.get("Script")),
            flatten_value(node.get("ConnectionProfile")),
            planning_summary(node.get("When")),
            vars_text,
            " | ".join(extract_events(node.get("eventsToWaitFor"))),
            " | ".join(extract_events(node.get("eventsToAdd"))),
            " | ".join(extract_events(node.get("eventsToDelete"))),
            flatten_value(docs),
        ]
        ws.append(row)

    style_sheet(ws)
    add_hierarchy_colors(ws)
    ws.column_dimensions["B"].width = 48
    ws.column_dimensions["H"].width = 60
    ws.column_dimensions["J"].width = 45
    ws.column_dimensions["R"].width = 50
    ws.column_dimensions["S"].width = 60
    ws.column_dimensions["U"].width = 55
    ws.column_dimensions["V"].width = 50

    # Variables
    vws = wb.create_sheet("Variables")
    vws.append(["Objet", "Catégorie", "Chemin", "Variable", "Valeur"])
    for obj in parser.objects:
        for key, value in extract_variables(obj["node"]):
            vws.append([obj["name"], obj_category(obj["type"]), obj["path"], key, value])
    style_sheet(vws)
    vws.column_dimensions["C"].width = 60
    vws.column_dimensions["E"].width = 55

    # Events
    ews = wb.create_sheet("Events")
    ews.append(["Objet", "Catégorie", "Chemin", "Nature", "Statut/Condition", "Event"])
    for obj in parser.objects:
        node = obj["node"]
        for event in extract_events(node.get("eventsToWaitFor")):
            ews.append([obj["name"], obj_category(obj["type"]), obj["path"], "WaitFor", "", event])
        for event in extract_events(node.get("eventsToAdd")):
            ews.append([obj["name"], obj_category(obj["type"]), obj["path"], "Add", "OK/direct", event])
        for event in extract_events(node.get("eventsToDelete")):
            ews.append([obj["name"], obj_category(obj["type"]), obj["path"], "Delete", "", event])
        if obj["type"].startswith("Job"):
            for event, status in collect_conditional_event_adds(node):
                ews.append([obj["name"], "Job", obj["path"], "Add conditionnel", status, event])
    style_sheet(ews)
    ews.column_dimensions["C"].width = 60
    ews.column_dimensions["F"].width = 80

    # Planning
    pws = wb.create_sheet("Planning")
    pws.append(["Objet", "Catégorie", "Chemin", "Propriété When", "Valeur"])
    for obj in parser.objects:
        pairs = planning_pairs(obj["node"].get("When"))
        if not pairs:
            continue
        for key, value in pairs:
            pws.append([obj["name"], obj_category(obj["type"]), obj["path"], key, value])
    style_sheet(pws)
    pws.column_dimensions["C"].width = 60
    pws.column_dimensions["D"].width = 38
    pws.column_dimensions["E"].width = 55

    # Propriétés supplémentaires
    prws = wb.create_sheet("Proprietes")
    prws.append(["Objet", "Catégorie", "Chemin", "Propriété", "Valeur"])
    for obj in parser.objects:
        for key, value in simple_extra_properties(obj["node"]):
            prws.append([obj["name"], obj_category(obj["type"]), obj["path"], key, value])
    style_sheet(prws)
    prws.column_dimensions["C"].width = 60
    prws.column_dimensions["D"].width = 35
    prws.column_dimensions["E"].width = 80

    # Métadonnée source discrète
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.center.text = f"Control-M - {source_name}"
    ws.oddFooter.right.text = "Page &P / &N"

    wb.save(output_path)


# ---------------------------------------------------------------------------
# 6. EXCEL N°2 : VUE A PLAT, UNE LIGNE PAR OBJET
# ---------------------------------------------------------------------------

def flatten_for_columns(value, prefix=""):
    """Aplatit récursivement un bloc JSON en colonnes clé -> valeur."""
    out = {}

    if isinstance(value, dict):
        for key, child in value.items():
            key = str(key)
            full_key = f"{prefix}.{key}" if prefix else key
            if is_controlm_object(child):
                continue
            if isinstance(child, list) and any(
                is_controlm_object(x) for x in child if isinstance(x, dict)
            ):
                continue
            out.update(flatten_for_columns(child, full_key))
        return out

    if isinstance(value, list):
        # Les listes simples sont gardées dans une cellule. Les listes de dictionnaires
        # sont converties en texte JSON compact, sauf Variables qui est géré à part.
        out[prefix] = flatten_value(value)
        return out

    out[prefix] = flatten_value(value)
    return out


def flat_object_record(obj, objects_by_path):
    """Construit la ligne 'à plat' d'un objet Control-M."""
    node = obj["node"]
    category = obj_category(obj["type"])

    record = {
        "Niveau": obj["depth"],
        "Catégorie": category,
        "Nom": obj["name"],
        "Type Control-M": obj["type"],
        "Parent direct": direct_parent_name(obj["path"]),
        "Parent Folder": nearest_ancestor_name(objects_by_path, obj["parent_path"], "Folder"),
        "SubFolder": nearest_ancestor_name(objects_by_path, obj["parent_path"], "SubFolder"),
        "Chemin complet": obj["path"],
        "Occurrence": (
            f"{obj['occurrence']}/{obj['occurrence_total']}"
            if obj["occurrence_total"] > 1 else ""
        ),
    }

    if category == "Folder":
        record["Parent Folder"] = ""
        record["SubFolder"] = ""
    elif category == "SubFolder":
        record["SubFolder"] = obj["name"]

    # Propriétés générales et imbriquées. On exclut les objets enfants :
    # ils auront leur propre ligne dans le classeur.
    excluded = {"Name", "FolderName", "SubFolderName", "Variables", "Jobs"}
    for key, value in node.items():
        if key in excluded:
            continue
        if isinstance(value, dict) and is_controlm_object(value):
            continue
        if isinstance(value, list) and any(
            is_controlm_object(x) for x in value if isinstance(x, dict)
        ):
            continue
        if key in {"eventsToAdd", "eventsToWaitFor", "eventsToDelete"}:
            continue
        record.update(flatten_for_columns(value, str(key)))

    # Une colonne par variable : Variable.UG, Variable.SEVERITE, etc.
    # Si une variable apparaît plusieurs fois, on conserve toutes les valeurs.
    variables = defaultdict(list)
    for key, value in extract_variables(node):
        variables[key].append(value)
    for key, values in variables.items():
        record[f"Variable.{key}"] = " | ".join(values)

    # Events principaux dans des colonnes lisibles.
    record["Events.WaitFor"] = " | ".join(extract_events(node.get("eventsToWaitFor")))
    record["Events.Add"] = " | ".join(extract_events(node.get("eventsToAdd")))
    record["Events.Delete"] = " | ".join(extract_events(node.get("eventsToDelete")))

    conditional = collect_conditional_event_adds(node) if obj["type"].startswith("Job") else []
    if conditional:
        record["Events.AddConditionnel"] = " | ".join(
            f"{event} [{status}]" if status else event for event, status in conditional
        )
    else:
        record["Events.AddConditionnel"] = ""

    return record


def build_flat_excel(parser, output_path, source_name):
    """Crée le second Excel : une feuille, une ligne par objet."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Tous_les_objets"

    objects_by_path = {o["path"]: o for o in parser.objects}
    records = [flat_object_record(obj, objects_by_path) for obj in parser.objects]

    # Colonnes essentielles en premier, puis toutes les propriétés rencontrées.
    fixed_headers = [
        "Niveau", "Catégorie", "Nom", "Type Control-M",
        "Parent direct", "Parent Folder", "SubFolder", "Chemin complet", "Occurrence",
    ]
    preferred = [
        "Description", "Application", "SubApplication", "ControlmServer",
        "Host", "RunAs", "FileName", "FilePath", "Command", "Script",
        "ConnectionProfile", "DocumentationUrl.Url",
        "When.FromTime", "When.ToTime", "When.WeekDays", "When.MonthDays",
        "When.Months", "When.DaysRelation", "When.RuleBasedCalendars.Included",
        "Events.WaitFor", "Events.Add", "Events.Delete", "Events.AddConditionnel",
    ]

    all_keys = set()
    for record in records:
        all_keys.update(record.keys())

    headers = []
    for key in fixed_headers + preferred:
        if key in all_keys and key not in headers:
            headers.append(key)

    # Variables regroupées ensemble avant les propriétés restantes.
    variable_headers = sorted(k for k in all_keys if k.startswith("Variable."))
    headers.extend(k for k in variable_headers if k not in headers)

    remaining = sorted(k for k in all_keys if k not in headers)
    headers.extend(remaining)

    ws.append(headers)
    for record in records:
        ws.append([record.get(header, "") for header in headers])

    style_sheet(ws)

    # Coloration de la ligne en fonction de la catégorie.
    category_col = headers.index("Catégorie") + 1
    add_hierarchy_colors(ws, type_col=category_col)

    # Largeurs ciblées pour les colonnes les plus lourdes.
    for col_idx, header in enumerate(headers, start=1):
        letter = get_column_letter(col_idx)
        if header in {"Nom", "Chemin complet", "Description"}:
            ws.column_dimensions[letter].width = 45 if header != "Chemin complet" else 65
        elif header in {"Command", "Script"}:
            ws.column_dimensions[letter].width = 70
        elif header.startswith("Events."):
            ws.column_dimensions[letter].width = 55
        elif header.startswith("Variable."):
            ws.column_dimensions[letter].width = 28

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.center.text = f"Control-M - vue à plat - {source_name}"
    ws.oddFooter.right.text = "Page &P / &N"

    wb.save(output_path)


# ---------------------------------------------------------------------------
# 7. EXTRACTION DES INFORMATIONS PRINCIPALES POUR LA SYNTHESE
# ---------------------------------------------------------------------------
# Les exports peuvent légèrement varier selon les versions / types de job.
# Ces fonctions centralisent les commandes/scripts, le planning, les calendriers,
# la criticité, les cycles et les contrôles de définition.

def extract_command_script(node):
    """Affiche uniquement la définition pertinente pour le type du job.

    Command : commande seule. Script : fichier seul. EmbeddedScript : contenu
    et nom de fichier éventuel. Pour ces trois types, les champs non pertinents
    sont ignorés. Pour les autres types : commande, sinon contenu, sinon fichier.
    Aucun fichier externe n'est lu et aucune variable Control-M n'est résolue.
    """
    command = flatten_value(node.get("Command"))
    value = node.get("Script")
    script = "\n".join(flatten_value(x) for x in value) if isinstance(value, list) else flatten_value(value)
    filename = flatten_value(node.get("FileName"))
    directory = flatten_value(node.get("FilePath"))
    if filename.strip() and directory and not ntpath.isabs(filename):
        separator = "\\" if "\\" in directory else "/"
        filename = directory.rstrip("/\\") + separator + filename

    job_type = str(node.get("Type", ""))
    if job_type == "Job:Command":
        return command if command.strip() else ""
    if job_type == "Job:Script":
        return filename if filename.strip() else ""
    if job_type == "Job:EmbeddedScript":
        parts = [(label, text) for label, text in (("Script", script), ("Fichier", filename)) if text.strip()]
        if len(parts) == 1:
            return parts[0][1]
        return "\n\n".join(f"{label}: {text}" for label, text in parts)
    # Compatibilité pour les autres types : une seule information, par priorité.
    return next((text for text in (command, script, filename) if text.strip()), "")


def get_when_block(node):
    """Lit le premier bloc dictionnaire parmi When, Scheduling et Schedule."""
    for key in ("When", "Scheduling", "Schedule"):
        value = node.get(key)
        if isinstance(value, dict):
            return value
    return {}


def _clean_list(value, keep_use_parent=False):
    """Nettoie les valeurs techniques NONE et, selon le cas, USE PARENT."""
    if value is None:
        return []
    values = value if isinstance(value, list) else [value]
    result = []
    for item in values:
        text = flatten_value(item).strip()
        if not text or text.upper() == "NONE":
            continue
        if text.upper() == "USE PARENT" and not keep_use_parent:
            continue
        result.append(text)
    return result


def extract_calendar(node):
    """Extrait le ou les calendriers Control-M dans une cellule lisible."""
    when = get_when_block(node)
    if not when:
        return ""

    calendars = []
    use_parent = False

    # Forme la plus courante : When.RuleBasedCalendars.Included / Excluded.
    rbc = when.get("RuleBasedCalendars")
    if isinstance(rbc, dict):
        included = _clean_list(rbc.get("Included"), keep_use_parent=True)
        excluded = _clean_list(rbc.get("Excluded"), keep_use_parent=True)
        for value in included:
            if value.upper() == "USE PARENT":
                use_parent = True
            else:
                calendars.append(value)
        for value in excluded:
            if value.upper() == "USE PARENT":
                use_parent = True
            else:
                calendars.append(f"EXCL:{value}")

    # Quelques variantes d'exports utilisent des clés plus directes.
    for key in ("Calendar", "Calendars", "RuleBasedCalendar", "RuleBasedCalendarName"):
        for value in _clean_list(when.get(key), keep_use_parent=True):
            if value.upper() == "USE PARENT":
                use_parent = True
            else:
                calendars.append(value)

    # Déduplication en conservant l'ordre.
    calendars = list(dict.fromkeys(calendars))
    if calendars:
        return " | ".join(calendars)
    if use_parent:
        return "USE PARENT"
    return ""


def extract_planification_main(node):
    """
    Produit une planification courte en séparant volontairement les calendriers.

    Exemples :
      1700
      1700-2200 ; MON,TUE
      MON,TUE ; Jours mois:1,15
      USE PARENT
    """
    when = get_when_block(node)
    if not when:
        return ""

    parts = []

    from_time = when.get("FromTime") or when.get("StartTime")
    to_time = when.get("ToTime") or when.get("EndTime")
    if from_time and to_time:
        parts.append(f"{from_time}-{to_time}")
    elif from_time:
        parts.append(str(from_time))

    week_days = _clean_list(when.get("WeekDays") or when.get("Days"))
    if week_days:
        parts.append(",".join(week_days))

    month_days = _clean_list(when.get("MonthDays"))
    if month_days:
        parts.append("Jours mois:" + ",".join(month_days))

    months = _clean_list(when.get("Months"))
    if months:
        parts.append("Mois:" + ",".join(months))

    relation = flatten_value(when.get("DaysRelation")).strip()
    if relation and relation.upper() not in {"OR", "NONE"}:
        parts.append("Relation:" + relation)

    # Si aucune information de planning propre n'est présente mais que le calendrier
    # est explicitement hérité, on rend l'information visible.
    if not parts and extract_calendar(node) == "USE PARENT":
        return "USE PARENT"

    return " ; ".join(parts)


def extract_criticality(node):
    """
    Retourne STRICTEMENT la variable Control-M ``SEVERITE``.

    Règle métier demandée :
      Criticité = Variables -> SEVERITE

    On ne fait volontairement aucun rapprochement avec Priority, Severity,
    Criticality ou d'autres alias afin que la colonne Criticité corresponde
    exactement à la donnée utilisée dans les chaînes Control-M.
    """
    values = []
    for key, value in extract_variables(node):
        if str(key).strip().upper() == "SEVERITE" and value not in values:
            values.append(value)
    return " | ".join(values)


def _human_unit(value):
    """Traduit les unités Control-M les plus courantes pour la synthèse."""
    text = flatten_value(value).strip()
    mapping = {
        "minute": "minute", "minutes": "minutes",
        "hour": "heure", "hours": "heures",
        "day": "jour", "days": "jours",
    }
    return mapping.get(text.lower(), text)


def _human_from(value):
    """Traduit le point de référence d'un cycle : Start / End / Target."""
    text = flatten_value(value).strip()
    mapping = {
        "start": "depuis le début du run",
        "end": "depuis la fin du run",
        "target": "par rapport à l'heure cible",
    }
    return mapping.get(text.lower(), text)


def extract_cyclic(node):
    """
    Restitue les PARAMETRES cycliques utiles au lieu d'un simple Oui / Non.

    Formats Automation API actuels pris en charge :
      "Rerun": {
          "Every": "2", "Units": "Minutes", "From": "Start", "Times": "0"
      }

      "RerunIntervals": {
          "Intervals": ["30m", "2h"], "From": "End"
      }

      "RerunSpecificTimes": {
          "At": ["0900", "1100", "1710"], "Tolerance": "20"
      }

    Les anciens exports avec les champs Cyclic / CyclicType / Interval... sont
    également conservés pour compatibilité.

    Important : ``RerunLimit`` n'est PAS considéré comme cyclique. Dans
    Control-M, il limite les reruns d'un objet non cyclique.

    Valeur retournée :
      - objet non cyclique -> cellule vide ;
      - objet cyclique -> détails lisibles, par exemple :
            Toutes les 2 minutes | depuis le début du run | illimité
        ou  Heures: 0900, 1100, 1710 | Tolérance: 20 min
    """
    details = []

    # ------------------------------------------------------------------
    # 1) Format Automation API : Rerun
    # ------------------------------------------------------------------
    rerun = node.get("Rerun")
    if isinstance(rerun, dict):
        every = flatten_value(rerun.get("Every")).strip()
        units = _human_unit(rerun.get("Units") or "Minutes")
        from_value = _human_from(rerun.get("From") or "Start")
        times = flatten_value(rerun.get("Times")).strip()

        if every:
            details.append(f"Toutes les {every} {units}")
        if from_value:
            details.append(from_value)
        if times:
            details.append("illimité" if times == "0" else f"{times} cycles")

        return " | ".join(details) if details else "Cyclique"

    # ------------------------------------------------------------------
    # 2) Format Automation API : séquence d'intervalles
    # ------------------------------------------------------------------
    rerun_intervals = node.get("RerunIntervals")
    if isinstance(rerun_intervals, dict):
        intervals = rerun_intervals.get("Intervals")
        interval_values = _clean_list(intervals, keep_use_parent=True)
        if interval_values:
            details.append("Intervalles: " + ", ".join(interval_values))

        from_value = _human_from(rerun_intervals.get("From") or "Start")
        if from_value:
            details.append(from_value)

        return " | ".join(details) if details else "Cyclique"

    # ------------------------------------------------------------------
    # 3) Format Automation API : heures précises
    # ------------------------------------------------------------------
    rerun_times = node.get("RerunSpecificTimes")
    if isinstance(rerun_times, dict):
        at_values = _clean_list(rerun_times.get("At"), keep_use_parent=True)
        if at_values:
            details.append("Heures: " + ", ".join(at_values))

        tolerance = flatten_value(rerun_times.get("Tolerance")).strip()
        if tolerance:
            details.append(f"Tolérance: {tolerance} min")

        return " | ".join(details) if details else "Cyclique"

    # ------------------------------------------------------------------
    # 4) Compatibilité avec les exports plus anciens / plats
    # ------------------------------------------------------------------
    cyclic_flag = None
    for key in ("Cyclic", "IsCyclic", "CyclicJob"):
        if key not in node:
            continue
        raw = node.get(key)
        if isinstance(raw, bool):
            cyclic_flag = raw
        else:
            text = str(raw).strip().lower()
            if text in {"true", "yes", "y", "1", "oui", "o"}:
                cyclic_flag = True
            elif text in {"false", "no", "n", "0", "non"}:
                cyclic_flag = False
            elif text:
                cyclic_flag = True

    # Ces propriétés décrivent réellement le cycle. RerunLimit est volontairement
    # absent de la liste : ce n'est pas un indicateur de job cyclique.
    legacy_keys = (
        "CyclicType", "CyclicTolerance",
        "CyclicIntervalSequence", "CyclicTimesSequence",
        "RerunInterval", "RerunIntervalUnits", "RerunFrom",
        "RepeatInterval", "RepeatIntervalUnits",
        "Interval", "IntervalUnits", "Every",
    )
    labels = {
        "CyclicType": "Type",
        "CyclicTolerance": "Tolérance",
        "CyclicIntervalSequence": "Intervalles",
        "CyclicTimesSequence": "Heures",
        "RerunInterval": "Intervalle",
        "RerunIntervalUnits": "Unité",
        "RerunFrom": "Depuis",
        "RepeatInterval": "Intervalle",
        "RepeatIntervalUnits": "Unité",
        "Interval": "Intervalle",
        "IntervalUnits": "Unité",
        "Every": "Toutes les",
    }

    for key in legacy_keys:
        value = node.get(key)
        if value in (None, "", [], {}):
            continue
        cyclic_flag = True if cyclic_flag is None else cyclic_flag
        details.append(f"{labels.get(key, key)}: {flatten_value(value)}")

    if cyclic_flag is True:
        return " | ".join(details) if details else "Cyclique"

    # Pas de cycle défini : on préfère une cellule vide plutôt qu'un simple "Non".
    return ""


def definition_warning(node):
    """Contrôle ciblé, non bloquant, des trois types de jobs OS connus."""
    fields = {
        "Job:Command": ("Command", "Commande absente ou vide"),
        "Job:Script": ("FileName", "Nom de fichier de script absent ou vide"),
        "Job:EmbeddedScript": ("Script", "Contenu du script embarqué absent ou vide"),
    }
    field, message = fields.get(str(node.get("Type", "")), (None, ""))
    if field is None:
        return ""
    value = node.get(field)
    if isinstance(value, list):
        present = any(flatten_value(item).strip() for item in value)
    else:
        present = bool(flatten_value(value).strip())
    return "" if present else message


# ---------------------------------------------------------------------------
# 8. EXCEL N°3 : VUE SYNTHETIQUE
# ---------------------------------------------------------------------------

def format_summary_details(ws):
    """Hauteurs estimées selon les largeurs, plafonnées à 300 points (20 lignes).

    Cette mise en forme ne raccourcit pas la valeur. Au-delà de la hauteur
    visible, utiliser la barre de formule. La limite Excel reste 32 767 caractères.
    """
    headers = {cell.value: cell.column for cell in ws[1]}
    command_column = headers["Commande / Script"]
    warning_column = headers["Contrôle définition"]
    for row in ws.iter_rows(min_row=2):
        row[command_column - 1].font = Font(name="Consolas", size=10)
        lines = 1
        for cell in row:
            width = max(1, ws.column_dimensions[cell.column_letter].width - 3)
            text = str(cell.value or "").expandtabs(4)
            wrapped = sum(max(1, math.ceil(len(line) / width)) for line in text.split("\n"))
            lines = max(lines, wrapped)
        ws.row_dimensions[row[0].row].height = min(300, max(18, lines * 15 + 3))
        warning = row[warning_column - 1]
        if warning.value:
            warning.fill = PatternFill("solid", fgColor="FCE4D6")
            warning.font = Font(color="9C0006", bold=True)


def build_summary_excel(parser, output_path, source_name):
    """Crée le troisième Excel synthétique : une ligne par objet, colonnes principales."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthese"

    objects_by_path = {o["path"]: o for o in parser.objects}
    headers = [
        "Niveau", "Nom", "Parent Folder", "SubFolder",
        "Type", "Commande / Script", "Cyclique", "Planification",
        "Calendrier", "Criticité", "Host", "Run As", "Contrôle définition",
    ]
    ws.append(headers)

    for obj in parser.objects:
        node = obj["node"]
        category = obj_category(obj["type"])
        parent_folder = nearest_ancestor_name(objects_by_path, obj["parent_path"], "Folder")
        subfolder = nearest_ancestor_name(objects_by_path, obj["parent_path"], "SubFolder")

        if category == "Folder":
            parent_folder = ""
            subfolder = ""
        elif category == "SubFolder":
            subfolder = obj["name"]

        run_as = node.get("RunAs")
        if run_as in (None, ""):
            run_as = node.get("User", "")

        ws.append([
            obj["depth"],
            ("    " * obj["depth"]) + obj["name"],
            parent_folder,
            subfolder,
            obj["type"],
            extract_command_script(node),
            extract_cyclic(node),
            extract_planification_main(node),
            extract_calendar(node),
            extract_criticality(node),
            flatten_value(node.get("Host")),
            flatten_value(run_as),
            definition_warning(node),
        ])
        # Une commande commençant par '=' reste du texte, jamais une formule.
        ws.cell(ws.max_row, 6).data_type = "s"

    style_sheet(ws)

    # Type Control-M exact en E ; la coloration commence à la ligne 2.
    add_hierarchy_colors(ws, type_col=5)

    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 36
    ws.column_dimensions["D"].width = 36
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 80
    ws.column_dimensions["G"].width = 55
    ws.column_dimensions["H"].width = 45
    ws.column_dimensions["I"].width = 45
    ws.column_dimensions["J"].width = 16
    ws.column_dimensions["K"].width = 32
    ws.column_dimensions["L"].width = 28
    ws.column_dimensions["M"].width = 48
    format_summary_details(ws)

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.center.text = f"Control-M - synthèse - {source_name}"
    ws.oddFooter.right.text = "Page &P / &N"
    wb.save(output_path)

# ---------------------------------------------------------------------------
# 9. CONSOLIDATION OPTIONNELLE DES CLASSEURS
# ---------------------------------------------------------------------------

def merge_excel_exports(exports, output_directory):
    """Concatène les XLSX déjà générés, sans relire ni fusionner les JSON.

    exports : couples ordonnés (nom JSON complet, dossier des XLSX).
    Colonnes alignées par intitulé ; toutes les lignes sont conservées.
    La pipeline appelle cette fonction après les conversions unitaires.
    """
    exports = list(exports)
    if not exports:
        raise ValueError("Aucun export à consolider.")
    output_directory = Path(output_directory)
    families = {
        "hierarchie": ["Hierarchie", "Variables", "Events", "Planning", "Proprietes"],
        "ligne": ["Tous_les_objets"],
        "synthese": ["Synthese"],
    }
    for source, directory in exports:
        for family in families:
            path = Path(directory) / f"{Path(source).stem}_{family}.xlsx"
            if not path.is_file():
                raise FileNotFoundError(f"Classeur manquant : {path}")
    output_directory.mkdir(parents=True, exist_ok=True)
    results = []
    for family, sheet_names in families.items():
        merged = Workbook()
        merged.remove(merged.active)
        headers = {name: ["Source JSON"] for name in sheet_names}
        widths = {name: {} for name in sheet_names}
        # Union ordonnée des en-têtes, notamment des propriétés/variables à plat.
        for source, directory in exports:
            path = Path(directory) / f"{Path(source).stem}_{family}.xlsx"
            workbook = load_workbook(path)
            try:
                if workbook.sheetnames != sheet_names:
                    raise ValueError(f"Feuilles inattendues : {path}")
                for name in sheet_names:
                    ws = workbook[name]
                    source_headers = [cell.value for cell in ws[1]]
                    if (any(not isinstance(h, str) or not h for h in source_headers)
                            or len(set(source_headers)) != len(source_headers)
                            or "Source JSON" in source_headers):
                        raise ValueError(f"En-têtes invalides ou réservés : {path} / {name}")
                    for column, header in enumerate(source_headers, 1):
                        if header not in headers[name]:
                            headers[name].append(header)
                        width = ws.column_dimensions[get_column_letter(column)].width or 13
                        widths[name][header] = max(widths[name].get(header, 0), width)
            finally:
                workbook.close()
        for name in sheet_names:
            if len(headers[name]) > 16384:
                raise ValueError(f"Limite Excel de colonnes dépassée : {name}")
            merged.create_sheet(name).append(headers[name])
        # Concaténation dans l'ordre saisi, sans dédoublonner les objets.
        for source, directory in exports:
            path = Path(directory) / f"{Path(source).stem}_{family}.xlsx"
            workbook = load_workbook(path)
            try:
                for name in sheet_names:
                    src, dst = workbook[name], merged[name]
                    if dst.max_row + src.max_row - 1 > 1048576:
                        raise ValueError(f"Limite Excel de lignes dépassée : {name}")
                    columns = {header: i for i, header in enumerate(headers[name], 1)}
                    positions = [columns[cell.value] for cell in src[1]]
                    for row in src.iter_rows(min_row=2):
                        target_row = dst.max_row + 1
                        provenance = dst.cell(target_row, 1, source)
                        provenance.data_type = "s"
                        provenance.fill = copy(row[0].fill)
                        for cell, column in zip(row, positions):
                            target = dst.cell(target_row, column, cell.value)
                            target.data_type = cell.data_type
                            target.font = copy(cell.font)
                            target.fill = copy(cell.fill)
                            target.border = copy(cell.border)
                            target.number_format = cell.number_format
                            target.alignment = copy(cell.alignment)
                        dst.row_dimensions[target_row].height = src.row_dimensions[row[0].row].height
            finally:
                workbook.close()
        for name in sheet_names:
            ws = merged[name]
            style_sheet(ws, freeze="B2")
            ws.column_dimensions["A"].width = 36
            for column, header in enumerate(headers[name][1:], 2):
                ws.column_dimensions[get_column_letter(column)].width = widths[name][header]
            if name == "Synthese" and "Contrôle définition" in headers[name]:
                format_summary_details(ws)
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.oddHeader.center.text = "Control-M - consolidation"
            ws.oddFooter.right.text = "Page &P / &N"
        destination = output_directory / f"consolide_{family}.xlsx"
        merged.save(destination)
        merged.close()
        results.append(destination)
    return results


# ---------------------------------------------------------------------------
# 10. INTERFACE LIGNE DE COMMANDE ET BILANS
# ---------------------------------------------------------------------------

def parse_args():
    """Lit les options de conversion d’un seul JSON."""
    p = argparse.ArgumentParser(
        description=(
            "Convertit un export JSON Control-M en trois fichiers Excel : "
            "hiérarchique, à plat et synthétique."
        )
    )
    p.add_argument("input_json", help="Fichier JSON Control-M")
    p.add_argument(
        "-o", "--output",
        help="Nom du fichier Excel hiérarchique (défaut : <json>_hierarchie.xlsx)"
    )
    p.add_argument(
        "--flat-output",
        help="Nom du fichier Excel à plat (défaut : <json>_ligne.xlsx)"
    )
    p.add_argument(
        "--summary-output",
        help="Nom du fichier Excel synthétique (défaut : <json>_synthese.xlsx)"
    )
    p.add_argument(
        "--report-output",
        help=(
            "Bilan JSON (défaut : à côté de la synthèse, "
            "<json>_generation-info.json). Un bilan .txt est aussi écrit."
        ),
    )
    return p.parse_args()


def write_generation_report(report, path):
    """Écrit les bilans JSON/TXT dans un dossier existant et accessible.

    Le bilan peut décrire un échec de conversion ; une erreur d’écriture reste
    bloquante et empêche de garantir sa disponibilité.
    """
    path = Path(path)
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    lines = [
        "Bilan de génération Control-M", f"Source : {report['source_json']}",
        f"Statut : {report['status']}", f"Début (UTC) : {report['started_at']}",
        f"Fin (UTC) : {report['finished_at']}", f"Durée : {report['duration_seconds']} s",
        "Objets : " + ", ".join(f"{key}={value}" for key, value in report['counts'].items()),
        f"Définitions incomplètes : {len(report['warnings'])}",
    ]
    lines.extend(f"AVERTISSEMENT : {w['path']} ({w['type']}) : {w['message']}" for w in report['warnings'])
    lines.extend(f"Fichier généré : {file}" for file in report['outputs'])
    if report.get("error"):
        lines.append("ERREUR : " + report['error'])
    path.with_suffix(".txt").write_text("\n".join(lines) + "\n", encoding="utf-8")


# ---------------------------------------------------------------------------
# 11. CONVERSION UNITAIRE ET POINT D’ENTRÉE
# ---------------------------------------------------------------------------

def main():
    """Convertit un JSON et produit son bilan, puis propage tout échec."""
    args = parse_args()
    input_path = Path(args.input_json)
    output_path = (
        Path(args.output)
        if args.output
        else input_path.with_name(input_path.stem + "_hierarchie.xlsx")
    )
    flat_output_path = (
        Path(args.flat_output)
        if args.flat_output
        else input_path.with_name(input_path.stem + "_ligne.xlsx")
    )
    summary_output_path = (
        Path(args.summary_output)
        if args.summary_output
        else input_path.with_name(input_path.stem + "_synthese.xlsx")
    )
    report_path = (
        Path(args.report_output)
        if args.report_output
        else summary_output_path.parent / (input_path.stem + "_generation-info.json")
    )
    if report_path.suffix.lower() != ".json":
        raise SystemExit("--report-output doit désigner un fichier .json.")
    # Ne jamais écraser une entrée ou un classeur avec le bilan.
    protected = {p.resolve() for p in (input_path, output_path, flat_output_path, summary_output_path)}
    if report_path.resolve() in protected or report_path.with_suffix(".txt").resolve() in protected:
        raise SystemExit("Le chemin du bilan doit être distinct des entrées et des classeurs.")
    started = time.perf_counter()
    report = {
        "source_json": input_path.name,
        "started_at": datetime.now(timezone.utc).isoformat(),
        "status": "ECHEC", "counts": {"Folder": 0, "SubFolder": 0, "Job": 0, "Total": 0},
        "warnings": [], "outputs": [],
    }
    error = None
    try:
        with input_path.open("r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
        parser = ControlMParser()
        parser.walk(data)
        if not parser.objects:
            raise ValueError("Aucun Folder/SubFolder/Job Control-M détecté dans le JSON.")
        for obj in parser.objects:
            report["counts"][obj_category(obj["type"])] += 1
            warning = definition_warning(obj["node"])
            if warning:
                report["warnings"].append({"path": obj["path"], "type": obj["type"], "message": warning})
                print(f"AVERTISSEMENT : {obj['path']} : {warning}", flush=True)
        report["counts"]["Total"] = len(parser.objects)
        for builder, path in (
            (build_excel, output_path),
            (build_flat_excel, flat_output_path),
            (build_summary_excel, summary_output_path),
        ):
            builder(parser, path, input_path.name)
            report["outputs"].append(str(path))
            print(f"Fichier généré : {path}")
        report["status"] = "OK_AVEC_AVERTISSEMENTS" if report["warnings"] else "OK"
    except Exception as exc:
        # Conserver l’erreur dans le bilan avant de terminer avec un code non nul.
        error = exc
        report["error"] = f"{type(exc).__name__}: {exc}"
    finally:
        report["finished_at"] = datetime.now(timezone.utc).isoformat()
        report["duration_seconds"] = round(time.perf_counter() - started, 3)
        write_generation_report(report, report_path)
    print(f"Bilan : {report_path.with_suffix('.txt')}")
    if error is not None:
        raise SystemExit(report["error"])
    print(
        f"{report['status']} : {report['counts']['Total']} objets exportés, "
        f"{len(report['warnings'])} avertissement(s)."
    )


if __name__ == "__main__":
    main()
