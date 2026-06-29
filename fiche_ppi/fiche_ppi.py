"""
fiche_ppi.py
Génère une fiche PPI consolidée à partir de deux fichiers Excel (oral / écrit).

Usage:
    # Mode simple
    python -m ppi_analyser.fiche_ppi <file_oral.xlsx> <file_ecrit.xlsx>
    python -m ppi_analyser.fiche_ppi <file_oral.xlsx> <file_ecrit.xlsx> -o fiche.xlsx

    # Mode batch
    python -m ppi_analyser.fiche_ppi --batch <dossier/>
"""

# ── Imports ────────────────────────────────────────────────────────────────────

import argparse
import os
import re
import sys
from collections import Counter
from urllib.parse import quote

import numpy as np
import pandas as pd
import requests
from bs4 import BeautifulSoup

from .format_excel import format_ppi_bold


# ── CLI ────────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Génère une fiche PPI consolidée à partir de deux fichiers Excel."
    )
    parser.add_argument(
        "file_oral",  nargs="?", help="Fichier Excel oral  (.xlsx)"
    )
    parser.add_argument(
        "file_ecrit", nargs="?", help="Fichier Excel écrit (.xlsx)"
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="Fichier de sortie (mode simple ; défaut : <stem>_fiche.xlsx)."
    )
    parser.add_argument(
        "--batch", metavar="DOSSIER",
        help="Traite toutes les paires *_Or.xlsx / *_Ph.xlsx du dossier."
    )
    return parser.parse_args()


# ── Utilitaires partagés (réutilisés par gui.py) ──────────────────────────────

def derive_output(input_path: str) -> str:
    """
    /data/à bientôt_Or.xlsx  →  /data/à bientôt_fiche.xlsx
    /data/à bientôt_Ph.xlsx  →  /data/à bientôt_fiche.xlsx
    """
    dirpath  = os.path.dirname(input_path)
    basename = os.path.splitext(os.path.basename(input_path))[0]
    stem     = re.sub(r'[_\s]*(Or|Ph)$', '', basename, flags=re.IGNORECASE).strip()
    filename = f"{stem}_fiche.xlsx"
    return os.path.join(dirpath, filename) if dirpath else filename


def find_pairs(folder: str) -> list[tuple[str, str]]:
    """
    Cherche tous les *_Or.xlsx dans le dossier et tente de trouver
    la contrepartie *_Ph.xlsx (insensible à la casse).
    Retourne une liste de tuples (oral_path, ecrit_path).
    """
    pairs = []
    for fname in sorted(os.listdir(folder)):
        if re.search(r'_Or\.xlsx$', fname, re.IGNORECASE):
            stem = re.sub(r'[_\s]*Or\.xlsx$', '', fname, flags=re.IGNORECASE).strip()
            ph_name = None
            for candidate in os.listdir(folder):
                if re.search(r'[_\s]*Ph\.xlsx$', candidate, re.IGNORECASE):
                    cand_stem = re.sub(r'[_\s]*Ph\.xlsx$', '', candidate,
                                       flags=re.IGNORECASE).strip()
                    if cand_stem.lower() == stem.lower():
                        ph_name = candidate
                        break
            if ph_name:
                pairs.append((
                    os.path.join(folder, fname),
                    os.path.join(folder, ph_name),
                ))
    return pairs


def generate_one(oral: str, ecrit: str) -> tuple[str, list[str]]:
    df_oral     = stringify(pd.read_excel(oral))
    df_ecrit    = stringify(pd.read_excel(ecrit))

    warnings = check_integrity(df_oral, df_ecrit)
    for w in warnings:
        print(f"[warn] {w}", file=sys.stderr)

    df_combined = pd.concat([df_oral, df_ecrit], ignore_index=True)
    df_fiche    = build_fiche(df_oral, df_ecrit, df_combined)
    output      = derive_output(oral)
    format_ppi_bold(df_fiche, output)
    return output, warnings

def check_integrity(df_oral: pd.DataFrame, df_ecrit: pd.DataFrame) -> list[str]:
    warnings = []

    for label, df in [("Oral", df_oral), ("Écrit", df_ecrit)]:
        # Empty rows
        empty_rows = df.index[df.apply(lambda r: all(v == "" for v in r), axis=1)].tolist()
        if empty_rows:
            warnings.append(f"[{label}] {len(empty_rows)} ligne(s) vide(s) : {[i+2 for i in empty_rows]}")

        # Required columns
        for col in ["Forme PPI", "Lemme"]:
            if col in df.columns:
                missing = df.index[df[col] == ""].tolist()
                if missing:
                    warnings.append(f"[{label}] Colonne '{col}' vide sur {len(missing)} ligne(s) : {[i+2 for i in missing]}")

        # Portée vs Position count
        if "Portée" in df.columns and "Position" in df.columns:
            n_portee   = df["Portée"].apply(lambda x: x != "").sum()
            n_position = df["Position"].apply(lambda x: x != "").sum()
            if n_portee != n_position:
                warnings.append(f"[{label}] Portée ({n_portee}) ≠ Position ({n_position}) — ligne(s) incomplète(s)")

    return warnings
# ── Wiktionary ─────────────────────────────────────────────────────────────────

def get_wiktionary_pronunciation(expression: str, lang: str = "fr") -> list[str]:
    """
    Récupère les URLs audio MP3 pour une expression depuis Wiktionary.
    Retourne une liste d'URLs ou une liste vide en cas d'échec.
    """
    url = f"https://{lang}.wiktionary.org/wiki/{quote(expression)}"
    headers = {"User-Agent": "Mozilla/5.0 (pronunciation-fetcher/1.0)"}

    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")
        audio_urls = []

        for source in soup.find_all("source"):
            src = source.get("src", "")
            if src and src.startswith("//") and src.endswith(".mp3"):
                audio_urls.append("https:" + src)

        if not audio_urls:
            print(f"[warn] Aucun audio trouvé pour '{expression}'.", file=sys.stderr)

        return list(set(audio_urls))

    except requests.exceptions.ConnectionError:
        print("[warn] Impossible de joindre Wiktionary.", file=sys.stderr)
    except requests.exceptions.Timeout:
        print(f"[warn] Timeout pour '{expression}'.", file=sys.stderr)
    except requests.exceptions.HTTPError as e:
        print(f"[warn] HTTP {e.response.status_code} pour '{expression}'.", file=sys.stderr)
    except Exception as e:
        print(f"[warn] Erreur inattendue : {e}", file=sys.stderr)

    return []


# ── Nettoyage des données ──────────────────────────────────────────────────────

def stringify(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remplace les NaN / None / chaînes vides par "" et convertit toutes
    les colonnes en str. Ne modifie pas le DataFrame en place — retourne
    une copie.
    """
    df = df.copy()
    for col in df.columns:
        df[col] = df[col].apply(
            lambda x: ""
            if (x is None or (isinstance(x, float) and np.isnan(x)))
            else x
        )
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].apply(lambda x: "" if not x.split() else x)
    return df


def clean_list(series) -> list[str]:
    """Retourne les valeurs non-vides d'une Series ou d'une liste."""
    return [str(v) for v in series if isinstance(v, str) and v.strip()]


def join_unique(series) -> str:
    seen, values = set(), []
    for v in sorted(series.dropna().astype(str).str.strip().str.lower()):
        if v and v != "nan" and v not in seen:
            seen.add(v)
            values.append(v)
    return ", ".join(values)


def clean_modifieurs(modifieurs: list[str]) -> list[str]:
    return [re.sub(r"[.,»]", "", m).strip().lower() for m in modifieurs]


# ── Variantes formelles ────────────────────────────────────────────────────────

def remove_modifier(modifier_tokens: list[str], forme: str) -> str:
    """Retire les tokens de modifieur présents dans la forme."""
    modifier_lower = {t.lower() for t in modifier_tokens}
    tokens = forme.split()
    filtered = [t for t in tokens if t.lower() not in modifier_lower]
    result = " ".join(filtered).replace(" -", "-")
    return result.strip()


def compute_variantes_formelles(df: pd.DataFrame) -> list[str]:
    modifieurs = [m for m in df["Modifieurs"] if isinstance(m, str) and m.strip()]
    modifier_tokens = " ".join(modifieurs).lower().split()

    nodes = df["node"].str.replace(" -", "-", regex=False)
    variantes = nodes.apply(lambda x: remove_modifier(modifier_tokens, x))
    variantes = clean_modifieurs(variantes.tolist())
    return list(set(variantes))


# ── Cooccurrents ───────────────────────────────────────────────────────────────

def count_cooccurrents(df: pd.DataFrame) -> dict:
    raw = ", ".join(clean_list(df["Cooccurrents"]))

    # Normalize newlines
    raw = raw.replace("\n", " ")

    # Ensure ante:/post: always start a new item
    raw = re.sub(r"\s*(post\s*:)", r"; \1", raw)
    raw = re.sub(r"\s*(ante\s*:)", r"; \1", raw)

    # Split into segments by semicolon
    segments = [s.strip() for s in raw.replace(",", ";").split(";") if s.strip()]

    cooc_a, cooc_p = [], []
    for seg in segments:
        seg_lower = seg.lower()
        if re.match(r"ante\s*:", seg_lower):
            items = re.sub(r"^ante\s*:\s*", "", seg_lower).split(",")
            cooc_a.extend(i.strip() for i in items if i.strip())
        elif re.match(r"post\s*:", seg_lower):
            items = re.sub(r"^post\s*:\s*", "", seg_lower).split(",")
            cooc_p.extend(i.strip() for i in items if i.strip())
        elif "(a)" in seg_lower:
            cooc_a.append(re.sub(r"\(a\)", "", seg_lower).strip())
        elif "(p)" in seg_lower:
            cooc_p.append(re.sub(r"\(p\)", "", seg_lower).strip())

    all_coocs = cooc_a + cooc_p

    total = len(cooc_a) + len(cooc_p)
    pct_a = round(len(cooc_a) / total * 100, 2) if total > 0 else 0.0
    pct_p = round(len(cooc_p) / total * 100, 2) if total > 0 else 0.0

    def fmt(items):
        return ", ".join(f"{cooc} ({n})" for cooc, n in Counter(items).most_common())

    return {
        "all":       fmt(all_coocs),
        "anteposés": fmt(cooc_a),
        "postposés": fmt(cooc_p),
        "pct_a":     pct_a,
        "pct_p":     pct_p,
    }


def format_cooc_stats(stats: dict, label: str) -> str:
    return (
        f"<bold>{label}</bold>:\n"
        f"\t- <bold>Cooccurrents antéposés</bold>  ({stats['pct_a']}%) : {stats['anteposés']}\n"
        f"\t- <bold>Cooccurrents postposés</bold>  ({stats['pct_p']}%) : {stats['postposés']}\n"
        f"\t- <bold>Total</bold> : {stats['all']}\n"
    )


# ── Structure interactionnelle ─────────────────────────────────────────────────

def get_interaction_stats(df: pd.DataFrame) -> str:
    """Comptages sur Déclenchement, Portée, Position."""

    def fmt_counts(series):
        counts = Counter(clean_list(series.str.lower().str.strip()))
        return ", ".join(f"{val} ({n})" for val, n in counts.most_common())

    decl = fmt_counts(df["Déclenchement"])
    port = fmt_counts(df["Portée"])
    pos  = fmt_counts(df["Position"])

    return (
        f"\t- <bold>Déclenchement</bold> : {decl}\n"
        f"\t- <bold>Portée</bold>        : {port}\n"
        f"\t- <bold>Position</bold>      : {pos}\n"
    )


# ── Construction de la fiche ───────────────────────────────────────────────────

FICHE_COLS = [
    "Fe_1a PPI",
    "Fe_1b Acception",
    "Fe_1c Variantes formelles",
    "Fe_1e Prononciation",
    "Fe_2a Statut syntaxique phrase",
    "Fe_2b Type phrase",
    "Fe_2c Modalité de phrase",
    "Fe_2d Structure syntaxique globale",
    "Fe_2e Expansion éventuelle",
    "Fe_2f Construction syntaxique détaillée",
    "Fe_2g Alternances syntaxiques",
    "Fe_3a Fonction globale",
    "Fe_3b Fonctions spécifiques",
    "Fe_3c Codes Fonction globale",
    "Fe_3d Codes Fonctions spécifiques",
    "Fe_3e Fonctionnement pragma-sémantique",
    "Fe_3f Structure interactionnelle",
    "Fe_3g Contexte spécifique",
    "Fe_3h Modalité écrite et orale",
    "Fe_4a Cooccurrents privilégiés communs à la PPI",
    "Fe_4b Modifieurs de la PPI",
    "Fe_4c Renvois synonymiques",
    "Fe_5a Marques d'usage de la PPI",
    "Fe_6a Définitions et sources",
    "Fe_7a Gestes/comportements associés",
    "Fe_9a Remarques",
    "Fe_9b Références",
    "Fe_10a Noms des rédacteurs",
    "Fe_10b Date de mise à jour",
]


def build_fiche(
    df_oral:     pd.DataFrame,
    df_ecrit:    pd.DataFrame,
    df_combined: pd.DataFrame,
) -> pd.DataFrame:

    # Variantes formelles
    var_oral  = compute_variantes_formelles(df_oral)
    var_ecrit = compute_variantes_formelles(df_ecrit)
    variantes_str = (
        f"\n- <bold>Oral</bold>  : {', '.join(var_oral)}"
        f"\n- <bold>Écrit</bold> : {', '.join(var_ecrit)}\n"
    )

    # Prononciation
    lemme = df_oral["Lemme"].iloc[0]
    prononciation = "\n".join(get_wiktionary_pronunciation(lemme))

    # Cooccurrents
    cooc_oral     = count_cooccurrents(df_oral)
    cooc_ecrit    = count_cooccurrents(df_ecrit)
    cooc_combined = count_cooccurrents(df_combined)

    cooc_str = (
        format_cooc_stats(cooc_oral,      "Oral")
        + format_cooc_stats(cooc_ecrit,   "Écrit")
        + format_cooc_stats(cooc_combined, "Les deux modes combiné")
    )

    # Structure interactionnelle
    stats_oral     = get_interaction_stats(df_oral)
    stats_ecrit    = get_interaction_stats(df_ecrit)
    stats_combined = get_interaction_stats(df_combined)

    stats_str = (
        "<bold>Oral</bold> :\n"      + stats_oral
        + "<bold>Écrit</bold> :\n"   + stats_ecrit
        + "<bold>Les deux modes combinés</bold> :\n" + stats_combined
    )

    # Remplissage
    df_fiche = pd.DataFrame(columns=FICHE_COLS, index=[0])

    df_fiche["Fe_1a PPI"]                                       = df_combined["Forme PPI"].values[0]
    df_fiche["Fe_1b Acception"]                                 = join_unique(df_combined["Acception"])
    df_fiche["Fe_1c Variantes formelles"]                       = variantes_str
    df_fiche["Fe_1e Prononciation"]                             = prononciation
    df_fiche["Fe_2a Statut syntaxique phrase"]                  = join_unique(df_combined["Type de phrase"])
    df_fiche["Fe_2c Modalité de phrase"]                        = join_unique(df_combined["Modalité d'énonciation"])
    df_fiche["Fe_2e Expansion éventuelle"]                      = join_unique(df_combined["Expansion"])
    df_fiche["Fe_3a Fonction globale"]                          = join_unique(df_combined["Fonction globale"])
    df_fiche["Fe_3b Fonctions spécifiques"]                     = join_unique(df_combined["Fonctions spécifiques"])
    df_fiche["Fe_3c Codes Fonction globale"]                    = join_unique(df_combined["Fonction globale"])
    df_fiche["Fe_3d Codes Fonctions spécifiques"]               = join_unique(df_combined["Fonctions spécifiques"])
    df_fiche["Fe_3f Structure interactionnelle"]                = stats_str
    df_fiche["Fe_3g Contexte spécifique"]                       = join_unique(df_combined["milieu"])
    df_fiche["Fe_3h Modalité écrite et orale"]                  = join_unique(df_combined["secteur"])
    df_fiche["Fe_4a Cooccurrents privilégiés communs à la PPI"] = cooc_str
    df_fiche["Fe_4b Modifieurs de la PPI"]                      = join_unique(df_combined["Modifieurs"])
    df_fiche["Fe_9a Remarques"]                                 = join_unique(df_combined["Remarques"])

    # Pivot
    df_fiche = df_fiche.T.reset_index()
    df_fiche.rename(columns={"index": "Propriétés", 0: "Valeurs"}, inplace=True)

    return df_fiche


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()

    # ── Mode batch ─────────────────────────────────────────────────────────────
    if args.batch:
        pairs = find_pairs(args.batch)
        if not pairs:
            print("[erreur] Aucune paire *_Or.xlsx / *_Ph.xlsx trouvée.", file=sys.stderr)
            sys.exit(1)
        print(f"[batch] {len(pairs)} paire(s) détectée(s).")
        ok, errors = 0, []
        for oral, ecrit in pairs:
            try:
                out, warnings = generate_one(oral, ecrit)
                ok += 1
                print(f"  [ok] {os.path.basename(out)}")
            except Exception as e:
                errors.append(os.path.basename(oral))
                print(f"  [err] {os.path.basename(oral)} : {e}", file=sys.stderr)
        print(f"\n[batch] {ok}/{len(pairs)} fiche(s) générée(s).")
        sys.exit(1 if errors else 0)

    # ── Mode simple ────────────────────────────────────────────────────────────
    if not args.file_oral or not args.file_ecrit:
        print(
            "[erreur] Fournir file_oral et file_ecrit, ou --batch DOSSIER.",
            file=sys.stderr,
        )
        sys.exit(1)

    df_oral     = stringify(pd.read_excel(args.file_oral))
    df_ecrit    = stringify(pd.read_excel(args.file_ecrit))
    df_combined = pd.concat([df_oral, df_ecrit], ignore_index=True)
    df_fiche    = build_fiche(df_oral, df_ecrit, df_combined)
    output      = args.output or derive_output(args.file_oral)
    format_ppi_bold(df_fiche, output)
    print(f"[ok] Fiche exportée → {output}")


if __name__ == "__main__":
    main()