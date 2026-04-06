"""
Application Streamlit - Gestion des rattrapages étudiants (v2 avec compensations UE)
Lancer avec : streamlit run rattrapages_app_v2.py
Dépendances : pip install streamlit pandas openpyxl
"""

import io
import re
import base64
import itertools
import streamlit as st
import pandas as pd

# ─── CONFIG PAGE ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Rattrapages – Tableau de bord",
    page_icon="🎓",
    layout="wide",
)

# ─── STYLES CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #f0f4ff 0%, #faf0ff 100%);
}
.hero {
    background: linear-gradient(120deg, #4f46e5, #7c3aed);
    border-radius: 16px;
    padding: 2rem 2.5rem;
    margin-bottom: 1.8rem;
    color: white;
    box-shadow: 0 8px 32px rgba(79,70,229,0.25);
}
.hero h1 { margin: 0; font-size: 2rem; font-weight: 800; letter-spacing: -0.5px; }
.hero p  { margin: 0.3rem 0 0; opacity: 0.85; font-size: 1rem; }

.badge {
    display: inline-block;
    font-weight: 700;
    font-size: 0.78rem;
    padding: 2px 9px;
    border-radius: 20px;
    margin: 1px;
}
.badge-A { background:#d1fae5; color:#065f46; border:1px solid #6ee7b7; }
.badge-B { background:#dbeafe; color:#1e3a8a; border:1px solid #93c5fd; }
.badge-C { background:#fef3c7; color:#78350f; border:1px solid #fcd34d; }
.badge-D { background:#fee2e2; color:#7f1d1d; border:1px solid #fca5a5; }
.badge-ABS { background:#ffedd5; color:#7c2d12; border:1px solid #fb923c; }
.badge-VAL { background:#d1fae5; color:#065f46; border:1px solid #6ee7b7; font-size:0.72rem; }
.badge-NVAL { background:#fee2e2; color:#7f1d1d; border:1px solid #fca5a5; font-size:0.72rem; }
.badge-COMP { background:#e0e7ff; color:#3730a3; border:1px solid #a5b4fc; font-size:0.72rem; }

.legend-row {
    display:flex; gap:10px; align-items:center; flex-wrap:wrap;
    margin-bottom: 0.5rem;
}
.legend-item { display:flex; align-items:center; gap:6px; font-size:0.85rem; font-weight:600; }
.dot { width:14px; height:14px; border-radius:50%; display:inline-block; }
.dot-A {background:#10b981;} .dot-B {background:#3b82f6;}
.dot-C {background:#f59e0b;} .dot-D {background:#ef4444;} .dot-ABS {background:#f97316;}

[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

.section-title {
    font-size: 1.1rem; font-weight: 700;
    color: #4f46e5; margin-bottom: 0.4rem;
    display: flex; align-items: center; gap: 8px;
}

.ue-card {
    border-radius: 10px;
    padding: 12px 16px;
    margin-bottom: 8px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTES MENTIONS ─────────────────────────────────────────────────────────
GRADE_VALUES = {"A": 5, "B": 4, "C": 2, "D": 1}

# ─── HELPERS ────────────────────────────────────────────────────────────────────

def badge(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    v = str(val).strip()
    css = f"badge badge-{v}" if v in ("A","B","C","D","ABS") else "badge"
    return f'<span class="{css}">{v}</span>'


def split_name(personne: str):
    if "," in personne:
        parts  = personne.split(",", 1)
        nom    = parts[0].strip().title()
        prenom = parts[1].strip().title()
    else:
        tokens = personne.strip().split()
        nom    = tokens[0].title() if tokens else personne
        prenom = " ".join(tokens[1:]).title() if len(tokens) > 1 else ""
    return prenom, nom


def short_eval_name(col: str) -> str:
    return re.sub(r"^Eval\s*-\s*", "", col).strip()


def generate_email(prenom: str, nom: str, matieres: list, tutoyer: bool) -> str:
    if tutoyer:
        intro  = f"Bonjour {prenom},"
        corps1 = "Nous t'informons que tu es concerné(e) par des rattrapages dans les matières suivantes :"
        corps2 = "Nous t'invitons donc à te présenter aux sessions de rattrapage qui te seront communiquées prochainement."
        sign   = "N'hésite pas à nous contacter si tu as des questions.\n\nBien cordialement,\nL'équipe pédagogique"
    else:
        intro  = f"Bonjour {prenom} {nom},"
        corps1 = "Nous vous informons que vous êtes concerné(e) par des rattrapages dans les matières suivantes :"
        corps2 = "Nous vous invitons donc à vous présenter aux sessions de rattrapage qui vous seront communiquées prochainement."
        sign   = "N'hésitez pas à nous contacter si vous avez des questions.\n\nBien cordialement,\nL'équipe pédagogique"

    liste = "\n".join(f"  • {m}" for m in matieres)
    return f"{intro}\n\n{corps1}\n\n{liste}\n\n{corps2}\n\n{sign}"


def df_to_excel_bytes(
    df: pd.DataFrame,
    student_ue_results: dict = None,
    eval_display_cols: list = None,
    use_compensation: bool = False,
) -> bytes:
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:

        # ── Feuille 1 : tableau des mentions ──────────────────────────────────
        df.to_excel(writer, index=False, sheet_name="Mentions")
        ws = writer.sheets["Mentions"]

        header_fill = PatternFill("solid", fgColor="4F46E5")
        thin   = Side(style="thin", color="D1D5DB")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        fill_map = {
            "A":   PatternFill("solid", fgColor="D1FAE5"),
            "B":   PatternFill("solid", fgColor="DBEAFE"),
            "C":   PatternFill("solid", fgColor="FEF3C7"),
            "D":   PatternFill("solid", fgColor="FEE2E2"),
            "ABS": PatternFill("solid", fgColor="FFEDD5"),
        }
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if cell.value in fill_map:
                    cell.fill = fill_map[cell.value]
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 18
        for i in range(3, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(i)].width = 28
        ws.row_dimensions[1].height = 45

        # ── Feuille 2 : rattrapages après compensation ─────────────────────────
        if use_compensation and student_ue_results and eval_display_cols:
            rows_export = []
            for _, row in df.iterrows():
                student_key = f"{row['Prénom']} {row['Nom']}"
                ue_res = student_ue_results.get(student_key, {})

                def mat_is_compensated(col_short):
                    for result in ue_res.values():
                        if result["validated"] and result["compensation"]:
                            for e in result["elements"]:
                                if col_short.lower() in e["element"].lower() or e["element"].lower() in col_short.lower():
                                    return True
                    return False

                matieres_rattrapage = []
                matieres_compensees = []
                for col in eval_display_cols:
                    grade = str(row.get(col, "")).strip()
                    if grade in ("C", "D", "ABS"):
                        if mat_is_compensated(col):
                            matieres_compensees.append(col)
                        else:
                            matieres_rattrapage.append(col)

                rows_export.append({
                    "Prénom": row["Prénom"],
                    "Nom":    row["Nom"],
                    "Matières en rattrapage": ", ".join(matieres_rattrapage) if matieres_rattrapage else "—",
                    "Nb rattrapages": len(matieres_rattrapage),
                    "Matières compensées (dispensées)": ", ".join(matieres_compensees) if matieres_compensees else "—",
                    "Nb compensées": len(matieres_compensees),
                })

            df_rattrap = pd.DataFrame(rows_export)
            # Trier : ceux avec rattrapages d'abord, puis par nb décroissant
            df_rattrap = df_rattrap.sort_values(["Nb rattrapages", "Nom"], ascending=[False, True])
            df_rattrap.to_excel(writer, index=False, sheet_name="Rattrapages")

            ws2 = writer.sheets["Rattrapages"]
            purple_fill = PatternFill("solid", fgColor="4F46E5")
            orange_fill = PatternFill("solid", fgColor="FEE2E2")
            green_fill  = PatternFill("solid", fgColor="D1FAE5")
            comp_fill   = PatternFill("solid", fgColor="E0E7FF")

            for cell in ws2[1]:
                cell.font = Font(bold=True, color="FFFFFF", size=10)
                cell.fill = purple_fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws2.row_dimensions[1].height = 40

            for row_cells in ws2.iter_rows(min_row=2):
                nb_rattrap_val = row_cells[3].value  # colonne "Nb rattrapages"
                nb_comp_val    = row_cells[5].value  # colonne "Nb compensées"
                for cell in row_cells:
                    cell.border = border
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                # Surligner la ligne selon statut
                if nb_rattrap_val and nb_rattrap_val > 0:
                    for cell in row_cells:
                        cell.fill = orange_fill
                elif nb_comp_val and nb_comp_val > 0:
                    for cell in row_cells:
                        cell.fill = comp_fill
                else:
                    for cell in row_cells:
                        cell.fill = green_fill
                # Centrer les chiffres
                row_cells[3].alignment = Alignment(horizontal="center", vertical="center")
                row_cells[5].alignment = Alignment(horizontal="center", vertical="center")

            ws2.column_dimensions["A"].width = 16
            ws2.column_dimensions["B"].width = 16
            ws2.column_dimensions["C"].width = 45
            ws2.column_dimensions["D"].width = 14
            ws2.column_dimensions["E"].width = 45
            ws2.column_dimensions["F"].width = 14

            # ── Feuille 3 : UE par étudiant ────────────────────────────────────
            rows_ue = []
            for _, row in df.iterrows():
                student_key = f"{row['Prénom']} {row['Nom']}"
                ue_res = student_ue_results.get(student_key, {})
                for ue_name, result in ue_res.items():
                    if result["weighted_avg"] is None:
                        continue
                    rows_ue.append({
                        "Prénom": row["Prénom"],
                        "Nom":    row["Nom"],
                        "UE": ue_name,
                        "Mention UE": result["mention"] or "—",
                        "Moy. pondérée": round(result["weighted_avg"], 2) if result["weighted_avg"] else None,
                        "Statut": (
                            "Validée par compensation" if result["validated"] and result["compensation"]
                            else "Validée" if result["validated"]
                            else "Non validée"
                        ),
                    })
            if rows_ue:
                df_ue = pd.DataFrame(rows_ue)
                df_ue.to_excel(writer, index=False, sheet_name="Résultats UE")
                ws3 = writer.sheets["Résultats UE"]
                for cell in ws3[1]:
                    cell.font = Font(bold=True, color="FFFFFF", size=10)
                    cell.fill = purple_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws3.row_dimensions[1].height = 35

                statut_fills = {
                    "Validée par compensation": PatternFill("solid", fgColor="E0E7FF"),
                    "Validée":                  PatternFill("solid", fgColor="D1FAE5"),
                    "Non validée":              PatternFill("solid", fgColor="FEE2E2"),
                }
                for row_cells in ws3.iter_rows(min_row=2):
                    statut_val = row_cells[5].value if len(row_cells) > 5 else None
                    row_fill = statut_fills.get(statut_val, PatternFill())
                    for cell in row_cells:
                        cell.border = border
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                        cell.fill = row_fill
                    if len(row_cells) > 3:
                        row_cells[3].alignment = Alignment(horizontal="center", vertical="center")
                    if len(row_cells) > 4:
                        row_cells[4].alignment = Alignment(horizontal="center", vertical="center")
                ws3.column_dimensions["A"].width = 16
                ws3.column_dimensions["B"].width = 16
                ws3.column_dimensions["C"].width = 40
                ws3.column_dimensions["D"].width = 12
                ws3.column_dimensions["E"].width = 14
                ws3.column_dimensions["F"].width = 24

    buf.seek(0)
    return buf.read()


# ─── LOGIQUE DE COMPENSATION UE ─────────────────────────────────────────────────

def load_ue_structure(rn_file, semestre: int) -> dict:
    """
    Charge la structure UE → éléments évaluables + coefficients pour un semestre donné.
    Retourne un dict: {nom_ue: [{element, coeff}, ...]}
    """
    df = pd.read_excel(rn_file)
    df_sem = df[df["Semestre Unite Enseignement"] == semestre].dropna(subset=["Libelle Element Evaluable"])
    df_sem = df_sem[df_sem["Coefficient Element Evaluable"].notna()]
    df_sem = df_sem[df_sem["Coefficient Element Evaluable"] > 0]

    ue_structure = {}
    for _, row in df_sem.iterrows():
        ue = row["Libelle Unite Enseignement"]
        elem = row["Libelle Element Evaluable"]
        coeff = float(row["Coefficient Element Evaluable"])
        if ue not in ue_structure:
            ue_structure[ue] = []
        ue_structure[ue].append({"element": elem, "coeff": coeff})
    return ue_structure


def compute_ue_result(grades_row: dict, ue_elements: list) -> dict:
    """
    Calcule le résultat d'une UE pour un étudiant.
    grades_row: dict {element_name: mention (A/B/C/D/ABS/NaN)}
    ue_elements: list of {element, coeff}
    
    Retourne: {
        'mention': A/B/C/D ou None si données manquantes,
        'weighted_avg': float,
        'elements': [{element, coeff, mention, value}],
        'validated': bool,
        'compensation': bool,  # True si validée par compensation
        'missing': bool        # True si données manquantes
    }
    """
    elements_data = []
    total_coeff = 0
    weighted_sum = 0
    missing = False

    for elem_info in ue_elements:
        elem_name = elem_info["element"]
        coeff = elem_info["coeff"]
        
        # Trouver la mention correspondante (matching partiel du nom)
        mention = None
        for key, val in grades_row.items():
            if elem_name.lower().strip() in key.lower().strip() or key.lower().strip() in elem_name.lower().strip():
                if pd.notna(val) and str(val).strip() != "":
                    mention = str(val).strip()
                break
        
        # Si pas trouvé, chercher par mots clés
        if mention is None:
            for key, val in grades_row.items():
                key_words = set(re.split(r'[\s\-:,]+', elem_name.lower()))
                col_words = set(re.split(r'[\s\-:,]+', key.lower()))
                overlap = key_words & col_words - {'', 'de', 'la', 'le', 'les', 'du', 'et', 'en', 'un', 'une'}
                if len(overlap) >= 3:
                    if pd.notna(val) and str(val).strip() != "":
                        mention = str(val).strip()
                    break

        if mention is None or mention == "ABS":
            # Absent ou manquant
            elements_data.append({
                "element": elem_name, "coeff": coeff,
                "mention": mention or "—", "value": None
            })
            missing = True
        elif mention in GRADE_VALUES:
            val = GRADE_VALUES[mention]
            weighted_sum += val * coeff
            total_coeff += coeff
            elements_data.append({
                "element": elem_name, "coeff": coeff,
                "mention": mention, "value": val
            })
        else:
            elements_data.append({
                "element": elem_name, "coeff": coeff,
                "mention": mention, "value": None
            })
            missing = True

    if total_coeff == 0:
        return {
            "mention": None, "weighted_avg": None,
            "elements": elements_data, "validated": False,
            "compensation": False, "missing": True
        }

    weighted_avg = weighted_sum / total_coeff

    # Règles de validation avec compensation
    # A si moy > 4.6, B si > 3.6, C si > 1.6 (non validé), D si <= 1.6 (non validé)
    if weighted_avg > 4.6:
        mention_ue = "A"
        validated = True
        compensation = False
    elif weighted_avg > 3.6:
        mention_ue = "B"
        validated = True
        compensation = False
    elif weighted_avg > 1.6:
        mention_ue = "C"
        validated = False
        compensation = False
    else:
        mention_ue = "D"
        validated = False
        compensation = False

    # Compensation : si une matière individuelle est C mais la moy UE > 3.6 → validée B
    # (Déjà géré par la logique ci-dessus)
    # On marque "compensation" si l'UE est validée alors qu'il y a au moins un C ou D individuel
    has_cd = any(
        e["mention"] in ("C", "D")
        for e in elements_data if e["value"] is not None
    )
    if validated and has_cd:
        compensation = True

    return {
        "mention": mention_ue,
        "weighted_avg": round(weighted_avg, 3),
        "elements": elements_data,
        "validated": validated,
        "compensation": compensation,
        "missing": missing
    }


def build_grade_lookup(row: pd.Series, eval_cols: list) -> dict:
    """Construit un dict {nom_court_element: mention} pour une ligne étudiant."""
    result = {}
    for col in eval_cols:
        short = short_eval_name(col)
        val = row.get(col, None)
        result[short] = val if pd.notna(val) else None
    return result


# ─── INTERFACE ──────────────────────────────────────────────────────────────────

st.markdown("""
<div class="hero">
  <h1>🎓 Gestion des rattrapages</h1>
  <p>Importez votre fichier de notes et le référentiel, filtrez par mention et calculez les compensations UE.</p>
</div>
""", unsafe_allow_html=True)

# ─── CHARGEMENT FICHIERS ─────────────────────────────────────────────────────────
col_up1, col_up2 = st.columns(2)

with col_up1:
    uploaded_notes = st.file_uploader(
        "📂 Fichier de notes (.xlsx)",
        type=["xlsx"],
        help="Fichier contenant 'Personne' et des colonnes 'Eval - ...'",
        key="notes_file"
    )

with col_up2:
    uploaded_rn = st.file_uploader(
        "📋 Référentiel cahier des charges (.xlsx)",
        type=["xlsx"],
        help="Fichier RN avec la structure UE/éléments évaluables et coefficients",
        key="rn_file"
    )

if uploaded_notes is None:
    st.info("⬆️ Veuillez importer le fichier de notes pour commencer.")
    st.stop()

try:
    raw_df = pd.read_excel(uploaded_notes)
except Exception as e:
    st.error(f"Impossible de lire le fichier de notes : {e}")
    st.stop()

if "Personne" not in raw_df.columns:
    st.error("Colonne 'Personne' introuvable. Vérifiez votre fichier.")
    st.stop()

all_eval_cols = [c for c in raw_df.columns if str(c).startswith("Eval")]
if not all_eval_cols:
    st.error("Aucune colonne commençant par 'Eval' trouvée dans le fichier.")
    st.stop()

# ─── SÉLECTION SEMESTRE ──────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">📅 Semestre</div>', unsafe_allow_html=True)

semestre = st.selectbox(
    "Semestre concerné :",
    options=[5, 6, 7, 8],
    index=2,
    help="Sélectionnez le semestre pour charger la bonne structure UE depuis le référentiel."
)

# ─── OPTIONS DE COLONNES ─────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">⚙️ Matières à inclure</div>', unsafe_allow_html=True)

GLOBAL_EXAM_KW  = "Préparation à la certification (Global exam)"
has_global_exam = any(GLOBAL_EXAM_KW.lower() in c.lower() for c in all_eval_cols)

opt_col1, opt_col2 = st.columns([1, 2])

with opt_col1:
    exclude_global = False
    if has_global_exam:
        exclude_global = st.toggle(
            "🚫 Exclure « Global exam »",
            value=False,
            help=f"Exclut toutes les colonnes contenant : « {GLOBAL_EXAM_KW} »",
        )

with opt_col2:
    short_names   = {c: short_eval_name(c) for c in all_eval_cols}
    short_to_full = {v: k for k, v in short_names.items()}
    all_short     = list(short_names.values())

    default_excluded = (
        [short_names[c] for c in all_eval_cols if GLOBAL_EXAM_KW.lower() in c.lower()]
        if exclude_global else []
    )

    excluded_short = st.multiselect(
        "Matières à exclure manuellement :",
        options=all_short,
        default=default_excluded,
        help="Ces matières n'apparaissent ni dans le tableau ni dans les mails.",
    )

excluded_full = {short_to_full[s] for s in excluded_short}
if exclude_global:
    excluded_full |= {c for c in all_eval_cols if GLOBAL_EXAM_KW.lower() in c.lower()}

eval_cols = [c for c in all_eval_cols if c not in excluded_full]

if not eval_cols:
    st.error("Toutes les matières sont exclues — veuillez en conserver au moins une.")
    st.stop()

# ─── NETTOYAGE ───────────────────────────────────────────────────────────────────
working = raw_df[["Personne"] + eval_cols].copy()
working = working[working[eval_cols].notna().any(axis=1)].reset_index(drop=True)

working.insert(0, "Prénom", working["Personne"].apply(lambda x: split_name(x)[0]))
working.insert(1, "Nom",    working["Personne"].apply(lambda x: split_name(x)[1]))
working = working.drop(columns=["Personne"])

display_cols      = {c: short_eval_name(c) for c in eval_cols}
display_df        = working.rename(columns=display_cols)
eval_display_cols = list(display_cols.values())

# ─── OPTION ABSENT ──────────────────────────────────────────────────────────────
absent_as_rattrapage = st.toggle(
    "🚨 Cellule vide = absent → convoqué aux rattrapages",
    value=False,
    help=(
        "Si activé : pour toute matière ayant au moins un résultat saisi, "
        "une cellule vide est considérée comme une absence et entraîne une convocation. "
        "Les colonnes entièrement vides (résultats non encore saisis) ne sont pas affectées."
    ),
)

if absent_as_rattrapage:
    active_eval_cols = [c for c in eval_display_cols if display_df[c].notna().any()]
    for col in active_eval_cols:
        display_df[col] = display_df[col].apply(
            lambda v: "ABS" if (pd.isna(v) or str(v).strip() == "") else v
        )

nb_excluded = len(excluded_full)
st.success(
    f"✅ {len(working)} étudiant(s) chargé(s) — "
    f"{len(eval_cols)} matière(s) active(s)"
    + (f" ({nb_excluded} exclue(s))" if nb_excluded else "") + "."
)

# ─── COMPENSATIONS UE ────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">⚖️ Compensations au sein des UE</div>', unsafe_allow_html=True)

use_compensation = st.toggle(
    "✨ Activer le calcul des compensations UE",
    value=True if uploaded_rn is not None else False,
    help=(
        "Calcule la moyenne pondérée par UE et applique les règles de compensation. "
        "Nécessite le fichier référentiel (cahier des charges)."
    ),
)

ue_structure = {}
if use_compensation:
    if uploaded_rn is None:
        st.warning("⚠️ Importez le fichier référentiel (cahier des charges) pour activer les compensations.")
        use_compensation = False
    else:
        try:
            ue_structure = load_ue_structure(uploaded_rn, semestre)
            st.success(f"📚 {len(ue_structure)} UE chargée(s) pour le semestre {semestre}.")
        except Exception as e:
            st.error(f"Erreur lors du chargement du référentiel : {e}")
            use_compensation = False

# ─── CALCUL COMPENSATIONS PAR ÉTUDIANT ──────────────────────────────────────────
student_ue_results = {}  # {student_name: {ue_name: result_dict}}

if use_compensation and ue_structure:
    for _, row in display_df.iterrows():
        student_key = f"{row['Prénom']} {row['Nom']}"
        grades_lookup = {}
        for col in eval_display_cols:
            val = row.get(col, None)
            grades_lookup[col] = val if (pd.notna(val) if val is not None else False) else None

        ue_results = {}
        for ue_name, ue_elems in ue_structure.items():
            result = compute_ue_result(grades_lookup, ue_elems)
            ue_results[ue_name] = result
        student_ue_results[student_key] = ue_results

# ─── FILTRES MENTIONS ────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">🔍 Filtres</div>', unsafe_allow_html=True)

fcol1, fcol2 = st.columns([1, 1])

with fcol1:
    st.markdown("""
    <div class="legend-row">
      <div class="legend-item"><span class="dot dot-A"></span> A – Admis</div>
      <div class="legend-item"><span class="dot dot-B"></span> B – Bien</div>
      <div class="legend-item"><span class="dot dot-C"></span> C – Ajourné léger</div>
      <div class="legend-item"><span class="dot dot-D"></span> D – Ajourné</div>
      <div class="legend-item"><span class="dot dot-ABS"></span> ABS – Absent</div>
    </div>
    """, unsafe_allow_html=True)

    selected_grades = st.multiselect(
        "Étudiants ayant au moins une de ces mentions :",
        options=["A", "B", "C", "D"],
        default=["C", "D"],
    )

with fcol2:
    group_filter = st.radio(
        "Groupe de mentions :",
        options=[
            "Tous (pas de filtre groupe)",
            "A ou B uniquement (admis / bien)",
            "C ou D uniquement (à rattraper)",
        ],
        index=2,
    )


def filter_df(df: pd.DataFrame, grades: list, group: str) -> pd.DataFrame:
    cols = eval_display_cols
    rattrapage_vals = ["C", "D", "ABS"]
    if grades:
        grades_with_abs = grades + (["ABS"] if any(g in grades for g in ["C", "D"]) else [])
        df = df[df[cols].isin(grades_with_abs).any(axis=1)]
    if group == "A ou B uniquement (admis / bien)":
        df = df[df[cols].isin(["A", "B"]).any(axis=1) & ~df[cols].isin(rattrapage_vals).any(axis=1)]
    elif group == "C ou D uniquement (à rattraper)":
        df = df[df[cols].isin(rattrapage_vals).any(axis=1)]
    return df


filtered_df = filter_df(display_df.copy(), selected_grades, group_filter)
st.markdown(f"**{len(filtered_df)}** étudiant(s) correspondent aux critères.")

# ─── TABLEAU ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">📋 Tableau des résultats</div>', unsafe_allow_html=True)

if not filtered_df.empty:
    headers = "<tr>" + "".join(
        f'<th style="background:#4f46e5;color:white;padding:8px 10px;'
        f'font-size:0.78rem;text-align:center;white-space:nowrap;">{h}</th>'
        for h in filtered_df.columns
    ) + "</tr>"

    html_rows = ""
    for i, (_, row) in enumerate(filtered_df.iterrows()):
        cells = ""
        for col in filtered_df.columns:
            val = row[col]
            if col in ("Prénom", "Nom"):
                cells += (
                    f'<td style="padding:6px 10px;font-weight:600;'
                    f'white-space:nowrap;font-size:0.85rem;">{val}</td>'
                )
            else:
                cells += f'<td style="text-align:center;padding:4px 8px;">{badge(val)}</td>'
        bg = "#f8fafc" if i % 2 == 0 else "white"
        html_rows += f'<tr style="background:{bg}">{cells}</tr>'

    st.markdown(f"""
    <div style="overflow-x:auto;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,0.08);">
    <table style="width:100%;border-collapse:collapse;font-family:sans-serif;">
      <thead>{headers}</thead><tbody>{html_rows}</tbody>
    </table></div>""", unsafe_allow_html=True)
else:
    st.warning("Aucun étudiant ne correspond aux filtres sélectionnés.")

# ─── SECTION COMPENSATIONS UE ─────────────────────────────────────────────────────
if use_compensation and ue_structure and not filtered_df.empty:
    st.markdown("---")
    st.markdown('<div class="section-title">⚖️ Résultats par UE avec compensations</div>', unsafe_allow_html=True)

    st.markdown("""
    <p style="font-size:0.85rem;color:#6b7280;margin-bottom:1rem;">
    Règles de compensation : <strong>A si moy > 4,6</strong> · <strong>B si moy > 3,6</strong> · 
    <strong>C (non validée) si moy > 1,6</strong> · <strong>D (non validée) si moy ≤ 1,6</strong><br>
    <em>Valeurs des mentions : A=5, B=4, C=2, D=1</em>
    </p>
    """, unsafe_allow_html=True)

    for _, row in filtered_df.iterrows():
        student_key = f"{row['Prénom']} {row['Nom']}"
        ue_results = student_ue_results.get(student_key, {})
        if not ue_results:
            continue

        # Vérifier si l'étudiant a des UE non validées
        non_validated = {ue: r for ue, r in ue_results.items() if not r["validated"] and not r["missing"]}
        compensated   = {ue: r for ue, r in ue_results.items() if r["validated"] and r["compensation"]}

        with st.expander(
            f"{'🔴' if non_validated else '🟢'} {student_key} — "
            f"{len(non_validated)} UE non validée(s) · {len(compensated)} compensation(s)"
        ):
            for ue_name, result in ue_results.items():
                if result["missing"] and result["weighted_avg"] is None:
                    continue

                # Couleur de la carte
                if result["missing"]:
                    bg, border = "#f8fafc", "#e5e7eb"
                elif result["validated"] and result["compensation"]:
                    bg, border = "#e0e7ff", "#a5b4fc"  # violet = compensé
                elif result["validated"]:
                    bg, border = "#d1fae5", "#6ee7b7"  # vert = validé
                else:
                    bg, border = "#fee2e2", "#fca5a5"  # rouge = non validé

                mention_badge = ""
                if result["mention"]:
                    css = f"badge badge-{result['mention']}"
                    mention_badge = f'<span class="{css}">{result["mention"]}</span>'

                status_badge = ""
                if result["missing"]:
                    status_badge = '<span class="badge" style="background:#f3f4f6;color:#6b7280;border:1px solid #d1d5db;">Données manquantes</span>'
                elif result["validated"] and result["compensation"]:
                    status_badge = '<span class="badge badge-COMP">✓ Validée par compensation</span>'
                elif result["validated"]:
                    status_badge = '<span class="badge badge-VAL">✓ Validée</span>'
                else:
                    status_badge = '<span class="badge badge-NVAL">✗ Non validée</span>'

                avg_str = f"{result['weighted_avg']:.2f}" if result['weighted_avg'] is not None else "—"

                # Détail des éléments
                elems_html = ""
                for e in result["elements"]:
                    e_mention = e["mention"]
                    e_badge = f'<span class="badge badge-{e_mention}">{e_mention}</span>' if e_mention in ("A","B","C","D","ABS") else f'<span style="color:#9ca3af;font-size:0.8rem;">{e_mention}</span>'
                    elems_html += (
                        f'<div style="display:flex;justify-content:space-between;'
                        f'align-items:center;padding:3px 0;border-bottom:1px solid #f3f4f6;">'
                        f'<span style="font-size:0.78rem;color:#374151;flex:1;">{e["element"]}</span>'
                        f'<span style="font-size:0.75rem;color:#6b7280;margin:0 8px;">coeff {int(e["coeff"])}</span>'
                        f'{e_badge}</div>'
                    )

                st.markdown(f"""
                <div style="background:{bg};border:1.5px solid {border};border-radius:10px;
                            padding:12px 16px;margin-bottom:8px;">
                  <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;flex-wrap:wrap;">
                    <span style="font-weight:700;font-size:0.88rem;flex:1;">{ue_name}</span>
                    {mention_badge}
                    <span style="font-size:0.82rem;color:#6b7280;">moy. {avg_str}</span>
                    {status_badge}
                  </div>
                  <div style="border-top:1px solid {border};padding-top:8px;">
                    {elems_html}
                  </div>
                </div>""", unsafe_allow_html=True)

# ─── EXPORT EXCEL ────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">📥 Export Excel</div>', unsafe_allow_html=True)

if not filtered_df.empty:
    st.download_button(
        label="⬇️ Télécharger le tableau filtré (.xlsx)",
        data=df_to_excel_bytes(
            filtered_df,
            student_ue_results=student_ue_results,
            eval_display_cols=eval_display_cols,
            use_compensation=use_compensation,
        ),
        file_name="rattrapages_filtrés.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ─── RÉCAP MATIÈRES CONCERNÉES ──────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">📊 Récapitulatif des matières en rattrapage</div>', unsafe_allow_html=True)


def is_compensated(student_name, element_short):
    """Retourne True si cet élément est compensé pour cet étudiant."""
    ue_res = student_ue_results.get(student_name, {})
    for result in ue_res.values():
        if result["validated"] and result["compensation"]:
            for e in result["elements"]:
                if element_short.lower() in e["element"].lower() or e["element"].lower() in element_short.lower():
                    return True
    return False


recap_rows = []
if not filtered_df.empty:
    for col in eval_display_cols:  # eval_display_cols tient déjà compte des exclusions
        eleves_c = [
            f"{row['Prénom']} {row['Nom']}"
            for _, row in filtered_df.iterrows()
            if str(row.get(col, "")).strip() == "C"
        ]
        eleves_d = [
            f"{row['Prénom']} {row['Nom']}"
            for _, row in filtered_df.iterrows()
            if str(row.get(col, "")).strip() == "D"
        ]
        eleves_abs = [
            f"{row['Prénom']} {row['Nom']}"
            for _, row in filtered_df.iterrows()
            if str(row.get(col, "")).strip() == "ABS"
        ]
        # Appliquer la compensation si activée
        if use_compensation and student_ue_results:
            eleves_c_comp   = [e for e in eleves_c   if is_compensated(e, col)]
            eleves_d_comp   = [e for e in eleves_d   if is_compensated(e, col)]
            eleves_c_nocomp = [e for e in eleves_c   if not is_compensated(e, col)]
            eleves_d_nocomp = [e for e in eleves_d   if not is_compensated(e, col)]
        else:
            eleves_c_comp = eleves_d_comp = []
            eleves_c_nocomp = eleves_c
            eleves_d_nocomp = eleves_d

        total = len(eleves_c) + len(eleves_d) + len(eleves_abs)
        if total > 0:
            recap_rows.append({
                "Matière": col,
                "eleves_c": eleves_c_nocomp, "nb_c": len(eleves_c_nocomp),
                "eleves_d": eleves_d_nocomp, "nb_d": len(eleves_d_nocomp),
                "eleves_abs": eleves_abs, "nb_abs": len(eleves_abs),
                "eleves_c_comp": eleves_c_comp, "eleves_d_comp": eleves_d_comp,
                "total": total,
                "total_rattrapage": len(eleves_c_nocomp) + len(eleves_d_nocomp) + len(eleves_abs),
            })

if not recap_rows:
    st.info("Aucune matière avec C ou D pour les étudiants sélectionnés.")
else:
    recap_rows.sort(key=lambda x: -x["total_rattrapage"])

    total_cd = sum(r["total_rattrapage"] for r in recap_rows)
    nb_mat   = len([r for r in recap_rows if r["total_rattrapage"] > 0])
    nb_comp_total = sum(len(r["eleves_c_comp"]) + len(r["eleves_d_comp"]) for r in recap_rows)

    st.markdown(
        f"<p style='font-size:0.85rem;color:#6b7280;margin-bottom:0.8rem;'>"
        f"<strong>{nb_mat}</strong> matière(s) concernée(s) · "
        f"<strong>{total_cd}</strong> situation(s) de rattrapage "
        f"(<span style='color:#f59e0b;font-weight:700;'>C : {sum(r['nb_c'] for r in recap_rows)}</span> · "
        f"<span style='color:#ef4444;font-weight:700;'>D : {sum(r['nb_d'] for r in recap_rows)}</span> · "
        f"<span style='color:#f97316;font-weight:700;'>ABS : {sum(r.get('nb_abs',0) for r in recap_rows)}</span>)"
        + (f" · <span style='color:#6366f1;font-weight:700;'>⚖️ {nb_comp_total} compensé(s)</span>" if nb_comp_total > 0 else "")
        + "</p>",
        unsafe_allow_html=True,
    )

    for i, r in enumerate(recap_rows):
        bg = "#f8fafc" if i % 2 == 0 else "white"

        def eleve_pill(name, mention, compensated=False):
            if compensated:
                color = ("#e0e7ff", "#3730a3", "#a5b4fc")
                label = f"{name} ⚖️"
            else:
                color = {
                    "C":   ("#fef3c7", "#78350f", "#fcd34d"),
                    "D":   ("#fee2e2", "#7f1d1d", "#fca5a5"),
                    "ABS": ("#ffedd5", "#7c2d12", "#fb923c"),
                }[mention]
                label = name
            return (
                f'<span style="display:inline-block;background:{color[0]};color:{color[1]};'
                f'border:1px solid {color[2]};border-radius:20px;padding:1px 10px;'
                f'font-size:0.76rem;font-weight:600;margin:2px;">{label}</span>'
            )

        pills_c = "".join(eleve_pill(e, "C") for e in r["eleves_c"])
        pills_d = "".join(eleve_pill(e, "D") for e in r["eleves_d"])
        pills_abs = "".join(eleve_pill(e, "ABS") for e in r.get("eleves_abs", []))
        pills_c_comp = "".join(eleve_pill(e, "C", compensated=True) for e in r.get("eleves_c_comp", []))
        pills_d_comp = "".join(eleve_pill(e, "D", compensated=True) for e in r.get("eleves_d_comp", []))

        badge_c   = f'<span class="badge badge-C">{r["nb_c"]}</span>' if r["nb_c"] else ""
        badge_d   = f'<span class="badge badge-D">{r["nb_d"]}</span>' if r["nb_d"] else ""
        badge_abs = f'<span class="badge badge-ABS">{r["nb_abs"]}</span>' if r.get("nb_abs") else ""
        nb_comp = len(r.get("eleves_c_comp", [])) + len(r.get("eleves_d_comp", []))
        badge_comp = f'<span class="badge badge-COMP">⚖️ {nb_comp} compensé(s)</span>' if nb_comp else ""

        bar_width = int(r["total_rattrapage"] / max(recap_rows[0]["total_rattrapage"], 1) * 100)
        bar_c   = int(r["nb_c"] / max(r["total_rattrapage"], 1) * bar_width)
        bar_d   = int(r["nb_d"] / max(r["total_rattrapage"], 1) * bar_width)
        bar_abs = bar_width - bar_c - bar_d
        bar = (
            f'<div style="display:flex;height:6px;border-radius:4px;overflow:hidden;'
            f'width:{bar_width}%;min-width:4px;margin-top:5px;">'
            f'<div style="flex:{bar_c} 0 0;background:#f59e0b;"></div>'
            f'<div style="flex:{bar_d} 0 0;background:#ef4444;"></div>'
            f'<div style="flex:{bar_abs} 0 0;background:#f97316;"></div>'
            f'</div>'
        )

        all_pills = pills_c + pills_d + pills_abs + pills_c_comp + pills_d_comp

        st.markdown(f"""
        <div style="background:{bg};border-radius:10px;padding:10px 16px;
                    margin-bottom:6px;box-shadow:0 1px 4px rgba(0,0,0,0.05);">
          <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
            <div style="flex:1;min-width:200px;">
              <span style="font-weight:700;font-size:0.88rem;">{r["Matière"]}</span>
              {bar}
            </div>
            <div style="display:flex;gap:5px;align-items:center;flex-wrap:wrap;">
              {badge_c}{badge_d}{badge_abs}{badge_comp}
              <span style="font-size:0.78rem;color:#6b7280;margin-left:2px;">/ {r["total_rattrapage"]}</span>
            </div>
          </div>
          <div style="margin-top:7px;line-height:2;">{all_pills}</div>
        </div>""", unsafe_allow_html=True)

# ─── MAILS ───────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">✉️ Mails de convocation aux rattrapages</div>', unsafe_allow_html=True)

if filtered_df.empty:
    st.info("Aucun étudiant sélectionné — ajustez les filtres.")
else:
    tutoyer    = st.toggle("👋 Tutoyer les étudiants (sinon vouvoiement)", value=False)
    toggle_key = "tu" if tutoyer else "vous"

    students_with_rattrapage = []
    for _, row in filtered_df.iterrows():
        student_key = f"{row['Prénom']} {row['Nom']}"
        matieres = []
        for c in eval_display_cols:  # déjà filtré selon les exclusions
            grade = str(row.get(c, "")).strip()
            if grade not in ("C", "D", "ABS"):
                continue
            if use_compensation and student_ue_results and is_compensated(student_key, c):
                continue
            matieres.append(c)
        students_with_rattrapage.append((row["Prénom"], row["Nom"], matieres))

    students_with_rattrapage = [(p, n, m) for p, n, m in students_with_rattrapage if m]

    if not students_with_rattrapage:
        st.info("✅ Aucun étudiant avec C ou D non compensé dans les matières actives.")
    else:
        for prenom, nom, matieres in students_with_rattrapage:
            with st.expander(f"📧 {prenom} {nom}  —  {len(matieres)} matière(s)"):
                badges_html = " ".join(
                    f'<span style="background:#fee2e2;color:#7f1d1d;border:1px solid #fca5a5;'
                    f'border-radius:6px;padding:2px 8px;font-size:0.78rem;font-weight:600;">{m}</span>'
                    for m in matieres
                )
                st.markdown(f"**Matières concernées :** {badges_html}", unsafe_allow_html=True)
                st.markdown("")

                mail_text = generate_email(prenom, nom, matieres, tutoyer)

                edited_mail = st.text_area(
                    "Contenu du mail (modifiable) :",
                    value=mail_text,
                    height=280,
                    key=f"mail_{toggle_key}_{prenom}_{nom}",
                )
                dl_col, copy_col = st.columns([2, 1])
                with dl_col:
                    st.download_button(
                        label="⬇️ Télécharger (.txt)",
                        data=edited_mail.encode("utf-8"),
                        file_name=f"mail_{prenom}_{nom}.txt".replace(" ", "_"),
                        mime="text/plain",
                        key=f"dl_{toggle_key}_{prenom}_{nom}",
                    )
                with copy_col:
                    b64 = base64.b64encode(edited_mail.encode("utf-8")).decode()
                    copy_id = f"copy_{toggle_key}_{prenom}_{nom}".replace(" ", "_").replace("-", "_")
                    st.markdown(f"""
                    <button id="{copy_id}"
                      onclick="
                        var txt = atob('{b64}');
                        navigator.clipboard.writeText(txt).then(function() {{
                          var btn = document.getElementById('{copy_id}');
                          btn.innerText = '✅ Copié !';
                          btn.style.background = '#d1fae5';
                          btn.style.color = '#065f46';
                          setTimeout(function() {{
                            btn.innerText = '📋 Copier le mail';
                            btn.style.background = '#ede9fe';
                            btn.style.color = '#4f46e5';
                          }}, 2000);
                        }});
                      "
                      style="width:100%;padding:8px 12px;border:1px solid #c4b5fd;
                             border-radius:8px;background:#ede9fe;color:#4f46e5;
                             font-weight:600;font-size:0.85rem;cursor:pointer;
                             transition:all 0.2s;">
                      📋 Copier le mail
                    </button>""", unsafe_allow_html=True)

# ─── RÉCAP CONVOCATIONS – MAIL CLASSE ───────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">📣 Récap convocations — mail à la classe</div>', unsafe_allow_html=True)

if filtered_df.empty or not recap_rows:
    st.info("Aucune donnée à afficher — ajustez les filtres.")
else:
    lines = ["Bonjour à tous,",
             "",
             "Voici le récapitulatif des rattrapages par matière :",
             ""]

    for r in recap_rows:
        if r["total_rattrapage"] == 0:
            continue
        all_students = r["eleves_c"] + r["eleves_d"] + r.get("eleves_abs", [])
        noms_liste   = ", ".join(all_students)
        comp_note = ""
        comp_students = r.get("eleves_c_comp", []) + r.get("eleves_d_comp", [])
        if comp_students:
            comp_note = f" [compensé(s) : {', '.join(comp_students)}]"
        lines.append(f"• {r['Matière']} : {noms_liste}{comp_note}")

    lines += [
        "",
        "Les étudiants concernés sont invités à se présenter aux sessions de rattrapage "
        "dont les modalités leur seront communiquées prochainement.",
        "",
        "Bien cordialement,",
        "L'équipe pédagogique",
    ]

    recap_classe_text = "\n".join(lines)

    edited_recap = st.text_area(
        "Contenu du mail (modifiable) :",
        value=recap_classe_text,
        height=320,
        key="recap_classe_textarea",
    )

    dl_col2, copy_col2 = st.columns([2, 1])
    with dl_col2:
        st.download_button(
            label="⬇️ Télécharger (.txt)",
            data=edited_recap.encode("utf-8"),
            file_name="recap_convocations_classe.txt",
            mime="text/plain",
            key="dl_recap_classe",
        )
    with copy_col2:
        b64_recap = base64.b64encode(edited_recap.encode("utf-8")).decode()
        st.markdown(f"""
        <button id="copy_recap_classe"
          onclick="
            var txt = atob('{b64_recap}');
            navigator.clipboard.writeText(txt).then(function() {{
              var btn = document.getElementById('copy_recap_classe');
              btn.innerText = '✅ Copié !';
              btn.style.background = '#d1fae5';
              btn.style.color = '#065f46';
              setTimeout(function() {{
                btn.innerText = '📋 Copier le mail';
                btn.style.background = '#ede9fe';
                btn.style.color = '#4f46e5';
              }}, 2000);
            }});
          "
          style="width:100%;padding:8px 12px;border:1px solid #c4b5fd;
                 border-radius:8px;background:#ede9fe;color:#4f46e5;
                 font-weight:600;font-size:0.85rem;cursor:pointer;
                 transition:all 0.2s;">
          📋 Copier le mail
        </button>""", unsafe_allow_html=True)

# ─── COMPATIBILITÉ CRÉNEAUX ─────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">🗓️ Rattrapages pouvant être organisés en parallèle</div>', unsafe_allow_html=True)
st.markdown(
    "<p style='font-size:0.85rem;color:#6b7280;margin-bottom:1rem;'>"
    "Deux matières sont <strong>compatibles</strong> (peuvent avoir lieu en même temps) "
    "si elles n'ont <strong>aucun élève en commun</strong> parmi les convoqués (C, D ou ABS) "
    "<strong>non compensés</strong>."
    "</p>", unsafe_allow_html=True
)

if not filtered_df.empty and recap_rows:
    mat_students = {}
    for r in recap_rows:
        eleves = set(r["eleves_c"] + r["eleves_d"] + r.get("eleves_abs", []))
        if eleves:
            mat_students[r["Matière"]] = eleves

    matieres_list = list(mat_students.keys())

    if len(matieres_list) < 2:
        st.info("Pas assez de matières avec rattrapages pour calculer des compatibilités.")
    else:
        def sont_compatibles(m1, m2):
            return mat_students[m1].isdisjoint(mat_students[m2])

        groupes = []
        reste = list(matieres_list)
        while reste:
            groupe = [reste[0]]
            for m in reste[1:]:
                if all(sont_compatibles(m, g) for g in groupe):
                    groupe.append(m)
            groupes.append(groupe)
            reste = [m for m in reste if m not in groupe]

        slot_colors = [
            ("#ede9fe", "#4f46e5", "#c4b5fd"),
            ("#d1fae5", "#065f46", "#6ee7b7"),
            ("#dbeafe", "#1e3a8a", "#93c5fd"),
            ("#fef3c7", "#78350f", "#fcd34d"),
            ("#fee2e2", "#7f1d1d", "#fca5a5"),
            ("#f3e8ff", "#581c87", "#d8b4fe"),
            ("#ffedd5", "#7c2d12", "#fb923c"),
        ]

        st.markdown(
            f"<p style='font-size:0.9rem;'><strong>{len(groupes)} créneau(x) minimum</strong> "
            f"nécessaire(s) pour organiser tous les rattrapages sans conflit.</p>",
            unsafe_allow_html=True
        )

        for i, groupe in enumerate(groupes):
            bg, fg, border = slot_colors[i % len(slot_colors)]

            all_eleves_creneau = set()
            for m in groupe:
                all_eleves_creneau |= mat_students[m]

            lignes_parts = []
            for m in groupe:
                eleves_m = sorted(mat_students[m])
                pills = "".join(
                    '<span style="display:inline-block;background:white;color:' + fg
                    + ';border:1px solid ' + border
                    + ';border-radius:20px;padding:1px 9px;font-size:0.74rem;font-weight:600;margin:2px;">'
                    + e + '</span>'
                    for e in eleves_m
                )
                lignes_parts.append(
                    '<div style="margin-bottom:8px;">'
                    + '<span style="font-weight:700;font-size:0.85rem;">' + m + '</span>'
                    + '<span style="font-size:0.78rem;color:' + fg + ';opacity:0.8;margin-left:6px;">'
                    + '(' + str(len(eleves_m)) + ' élève(s))</span>'
                    + '<div style="margin-top:4px;">' + pills + '</div>'
                    + '</div>'
                )
            lignes_html = "".join(lignes_parts)

            card = (
                '<div style="background:' + bg + ';border:1.5px solid ' + border
                + ';border-radius:12px;padding:14px 18px;margin-bottom:10px;box-shadow:0 2px 8px rgba(0,0,0,0.06);">'
                + '<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">'
                + '<span style="background:' + fg + ';color:white;border-radius:8px;padding:3px 12px;font-weight:800;font-size:0.9rem;">'
                + 'Créneau ' + str(i + 1) + '</span>'
                + '<span style="font-size:0.82rem;color:' + fg + ';font-weight:600;">'
                + str(len(groupe)) + ' matière(s) · ' + str(len(all_eleves_creneau)) + ' élève(s) au total'
                + '</span></div>'
                + lignes_html
                + '</div>'
            )
            st.markdown(card, unsafe_allow_html=True)

        with st.expander("🔍 Voir la matrice de compatibilité complète"):
            # Légende numérotée
            legend_parts = []
            for idx_m, m in enumerate(matieres_list):
                legend_parts.append(
                    f'<div style="font-size:0.78rem;padding:2px 0;color:#374151;">'
                    f'<span style="display:inline-block;width:30px;font-weight:800;color:#4f46e5;">M{idx_m+1}</span>'
                    f'{m}</div>'
                )
            st.markdown(
                '<div style="background:#f8fafc;border-radius:8px;padding:10px 14px;margin-bottom:12px;">'
                '<div style="font-weight:700;font-size:0.82rem;color:#4f46e5;margin-bottom:6px;">Légende des matières</div>'
                + "".join(legend_parts) + "</div>",
                unsafe_allow_html=True
            )

            # En-têtes numérotés M1, M2...
            th = "".join(
                f'<th style="background:#4f46e5;color:white;padding:7px 4px;'
                f'font-size:0.78rem;text-align:center;min-width:36px;font-weight:800;">'
                f'M{idx_m+1}</th>'
                for idx_m in range(len(matieres_list))
            )
            header_row = (
                f'<tr><th style="background:#4f46e5;color:white;padding:7px 12px;'
                f'font-size:0.78rem;text-align:left;white-space:nowrap;min-width:200px;">Matière</th>{th}</tr>'
            )

            rows_html = ""
            for i_m, m1 in enumerate(matieres_list):
                row_label = (
                    f'<td style="font-size:0.78rem;padding:5px 12px;white-space:nowrap;'
                    f'background:#f1f5f9;border-right:2px solid #e2e8f0;">'
                    f'<span style="color:#4f46e5;font-weight:800;margin-right:6px;">M{i_m+1}</span>'
                    f'{m1}</td>'
                )
                cells = row_label
                for j_m, m2 in enumerate(matieres_list):
                    if i_m == j_m:
                        cells += '<td style="background:#e2e8f0;text-align:center;font-size:0.82rem;">—</td>'
                    elif sont_compatibles(m1, m2):
                        cells += '<td style="background:#d1fae5;color:#065f46;text-align:center;font-weight:800;font-size:0.88rem;">✓</td>'
                    else:
                        nb_communs = len(mat_students[m1] & mat_students[m2])
                        cells += (
                            f'<td style="background:#fee2e2;color:#7f1d1d;text-align:center;'
                            f'font-size:0.8rem;font-weight:700;">{nb_communs}</td>'
                        )
                bg_row = "#f8fafc" if i_m % 2 == 0 else "white"
                rows_html += f'<tr style="background:{bg_row}">{cells}</tr>'

            st.markdown(
                '<p style="font-size:0.8rem;color:#6b7280;margin-bottom:6px;">'
                '✓ = compatible (0 élève en commun) &nbsp;·&nbsp; chiffre = nb d\'élèves en conflit</p>'
                '<div style="overflow-x:auto;border-radius:10px;box-shadow:0 1px 6px rgba(0,0,0,0.08);">'
                '<table style="border-collapse:collapse;font-family:sans-serif;">'
                f'<thead>{header_row}</thead><tbody>{rows_html}</tbody>'
                '</table></div>',
                unsafe_allow_html=True
            )
else:
    st.info("Aucune donnée disponible — importez un fichier et appliquez les filtres.")
