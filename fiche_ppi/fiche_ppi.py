"""
fiche_ppi.py
Génère une fiche PPI consolidée à partir de deux fichiers Excel (oral / écrit).

Usage:
    python fiche_ppi.py <file_oral.xlsx> <file_ecrit.xlsx>
"""

# ── Imports ────────────────────────────────────────────────────────────────────

import argparse
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
    parser.add_argument("file_oral",  help="Fichier Excel oral  (.xlsx)")
    parser.add_argument("file_ecrit", help="Fichier Excel écrit (.xlsx)")
    parser.add_argument(
        "-o", "--output",
        default="fiche.xlsx",
        help="Fichier de sortie (défaut: fiche.xlsx)"
    )
    return parser.parse_args()


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
        # Remplace les chaînes vides ou purement blanches par ""
        df[col] = df[col].apply(lambda x: "" if not x.split() else x)
    return df


def clean_list(series) -> list[str]:
    """Retourne les valeurs non-vides d'une Series ou d'une liste."""
    return [str(v) for v in series if isinstance(v, str) and v.strip()]


def join_unique(series) -> str:
    """Joint les valeurs uniques non-vides triées d'une Series."""
    values = sorted(
        set(series.dropna().astype(str).str.strip()) - {"", "nan"}
    )
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
    """
    Parse la colonne Cooccurrents et calcule les stats
    antéposés (a) / postposés (p).
    """
    raw = ", ".join(clean_list(df["Cooccurrents"]))
    items = [c.strip() for c in raw.replace(";", ",").split(",") if c.strip()]

    cooc_a = [c.lower() for c in items if "(a)" in c]
    cooc_p = [c.lower() for c in items if "(p)" in c]
    all_coocs = [
        re.sub(r"\((a|p)\)", "", c).strip().lower()
        for c in items
    ]

    total = len(cooc_a) + len(cooc_p)
    pct_a = round(len(cooc_a) / total * 100, 2) if total > 0 else 0.0
    pct_p = round(len(cooc_p) / total * 100, 2) if total > 0 else 0.0

    def fmt(counter_items, tag):
        pattern = f"({tag})"
        return ", ".join(
            f"{cooc.replace(pattern, '').strip()} ({n})"
            for cooc, n in counter_items
        )

    return {
        "all":       ", ".join(f"{c} ({n})" for c, n in Counter(all_coocs).most_common()),
        "anteposés": fmt(Counter(cooc_a).most_common(), "a"),
        "postposés": fmt(Counter(cooc_p).most_common(), "p"),
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
    df_oral:   pd.DataFrame,
    df_ecrit:  pd.DataFrame,
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
        format_cooc_stats(cooc_oral,     "Oral")
        + format_cooc_stats(cooc_ecrit,  "Écrit")
        + format_cooc_stats(cooc_combined, "Combiné")
    )

    # Structure interactionnelle
    stats_oral     = get_interaction_stats(df_oral)
    stats_ecrit    = get_interaction_stats(df_ecrit)
    stats_combined = get_interaction_stats(df_combined)

    stats_str = (
        "<bold>Oral</bold> :\n"     + stats_oral
        + "<bold>Écrit</bold> :\n"  + stats_ecrit
        + "<bold>Combiné</bold> :\n" + stats_combined
    )

    # Remplissage
    df_fiche = pd.DataFrame(columns=FICHE_COLS, index=[0])

    df_fiche["Fe_1a PPI"]                                   = df_combined["Forme PPI"].values[0]
    df_fiche["Fe_1b Acception"]                             = join_unique(df_combined["Acception"])
    df_fiche["Fe_1c Variantes formelles"]                   = variantes_str
    df_fiche["Fe_1e Prononciation"]                         = prononciation
    df_fiche["Fe_2a Statut syntaxique phrase"]              = join_unique(df_combined["Type de phrase"])
    df_fiche["Fe_2c Modalité de phrase"]                    = join_unique(df_combined["Modalité d'énonciation"])
    df_fiche["Fe_2e Expansion éventuelle"]                  = join_unique(df_combined["Expansion"])
    df_fiche["Fe_3a Fonction globale"]                      = join_unique(df_combined["Fonction globale"])
    df_fiche["Fe_3b Fonctions spécifiques"]                 = join_unique(df_combined["Fonctions spécifiques"])
    df_fiche["Fe_3c Codes Fonction globale"]                = join_unique(df_combined["Fonction globale"])
    df_fiche["Fe_3d Codes Fonctions spécifiques"]           = join_unique(df_combined["Fonctions spécifiques"])
    df_fiche["Fe_3f Structure interactionnelle"]            = stats_str
    df_fiche["Fe_3g Contexte spécifique"]                   = join_unique(df_combined["milieu"])
    df_fiche["Fe_3h Modalité écrite et orale"]              = join_unique(df_combined["secteur"])
    df_fiche["Fe_4a Cooccurrents privilégiés communs à la PPI"] = cooc_str
    df_fiche["Fe_4b Modifieurs de la PPI"]                  = join_unique(df_combined["Modifieurs"])
    df_fiche["Fe_9a Remarques"]                             = join_unique(df_combined["Remarques"])

    # Pivot
    df_fiche = df_fiche.T.reset_index()
    df_fiche.rename(columns={"index": "Propriétés", 0: "Valeurs"}, inplace=True)

    return df_fiche


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()

    df_oral  = stringify(pd.read_excel(args.file_oral))
    df_ecrit = stringify(pd.read_excel(args.file_ecrit))
    df_combined = pd.concat([df_oral, df_ecrit], ignore_index=True)

    df_fiche = build_fiche(df_oral, df_ecrit, df_combined)

    print(df_fiche.to_string(index=False))
    format_ppi_bold(df_fiche, args.output)
    print(f"\n[ok] Fiche exportée → {args.output}")


if __name__ == "__main__":
    main()
