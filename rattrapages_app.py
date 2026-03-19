"""
Application Streamlit - Gestion des rattrapages étudiants
Lancer avec : streamlit run rattrapages_app.py
Dépendances : pip install streamlit pandas openpyxl
"""

import io
import re
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
/* ── Fond général ── */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #f0f4ff 0%, #faf0ff 100%);
}

/* ── Header hero ── */
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

/* ── Badges A B C D ── */
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

/* ── Légende filtres ── */
.legend-row {
    display:flex; gap:10px; align-items:center; flex-wrap:wrap;
    margin-bottom: 0.5rem;
}
.legend-item {
    display:flex; align-items:center; gap:6px;
    font-size:0.85rem; font-weight:600;
}
.dot {
    width:14px; height:14px; border-radius:50%;
    display:inline-block;
}
.dot-A {background:#10b981;} .dot-B {background:#3b82f6;}
.dot-C {background:#f59e0b;} .dot-D {background:#ef4444;}

/* ── Carte étudiant ── */
.student-card {
    background: white;
    border-radius: 12px;
    padding: 1rem 1.3rem;
    margin-bottom: 0.7rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
    border-left: 4px solid #4f46e5;
}

/* ── Zone mail ── */
.mail-box {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    font-family: 'Courier New', monospace;
    font-size: 0.88rem;
    white-space: pre-wrap;
    line-height: 1.6;
    box-shadow: inset 0 1px 4px rgba(0,0,0,0.04);
}

/* ── Tableau ── */
[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

/* ── Sections ── */
.section-title {
    font-size: 1.1rem; font-weight: 700;
    color: #4f46e5; margin-bottom: 0.4rem;
    display: flex; align-items: center; gap: 8px;
}
</style>
""", unsafe_allow_html=True)

# ─── HELPERS ────────────────────────────────────────────────────────────────────

GRADE_COLORS = {
    "A": ("#10b981", "#d1fae5"),
    "B": ("#3b82f6", "#dbeafe"),
    "C": ("#f59e0b", "#fef3c7"),
    "D": ("#ef4444", "#fee2e2"),
}

def badge(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    v = str(val).strip()
    css = f"badge badge-{v}" if v in "ABCD" else "badge"
    return f'<span class="{css}">{v}</span>'


def split_name(personne: str):
    """'NOM, Prénom' → (prénom, nom)"""
    if "," in personne:
        parts = personne.split(",", 1)
        nom    = parts[0].strip().title()
        prenom = parts[1].strip().title()
    else:
        tokens = personne.strip().split()
        nom    = tokens[0].title() if tokens else personne
        prenom = " ".join(tokens[1:]).title() if len(tokens) > 1 else ""
    return prenom, nom


def short_eval_name(col: str) -> str:
    """Retire le préfixe 'Eval - ' et garde le reste."""
    return re.sub(r"^Eval\s*-\s*", "", col).strip()


def generate_email(prenom: str, nom: str, matieres: list[str], tutoyer: bool) -> str:
    if tutoyer:
        intro     = f"Bonjour {prenom},"
        corps1    = (
            f"Nous t'informons que tu es concerné(e) par des rattrapages "
            f"dans les matières suivantes :"
        )
        corps2    = (
            "Nous t'invitons donc à te présenter aux sessions de rattrapage "
            "qui te seront communiquées prochainement."
        )
        signature = (
            "N'hésite pas à nous contacter si tu as des questions.\n\n"
            "Bien cordialement,\nL'équipe pédagogique"
        )
    else:
        intro     = f"Bonjour {prenom} {nom},"
        corps1    = (
            f"Nous vous informons que vous êtes concerné(e) par des rattrapages "
            f"dans les matières suivantes :"
        )
        corps2    = (
            "Nous vous invitons donc à vous présenter aux sessions de rattrapage "
            "qui vous seront communiquées prochainement."
        )
        signature = (
            "N'hésitez pas à nous contacter si vous avez des questions.\n\n"
            "Bien cordialement,\nL'équipe pédagogique"
        )

    liste = "\n".join(f"  • {m}" for m in matieres)
    return f"{intro}\n\n{corps1}\n\n{liste}\n\n{corps2}\n\n{signature}"


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Résultats")
        ws = writer.sheets["Résultats"]
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        # En-têtes
        header_fill = PatternFill("solid", fgColor="4F46E5")
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Couleurs A/B/C/D
        fill_map = {
            "A": PatternFill("solid", fgColor="D1FAE5"),
            "B": PatternFill("solid", fgColor="DBEAFE"),
            "C": PatternFill("solid", fgColor="FEF3C7"),
            "D": PatternFill("solid", fgColor="FEE2E2"),
        }
        thin = Side(style="thin", color="D1D5DB")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if cell.value in fill_map:
                    cell.fill = fill_map[cell.value]

        # Largeurs
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 22
        for i in range(3, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(i)].width = 30
        ws.row_dimensions[1].height = 50
    buf.seek(0)
    return buf.read()


# ─── CHARGEMENT DU FICHIER ───────────────────────────────────────────────────────

st.markdown("""
<div class="hero">
  <h1>🎓 Gestion des rattrapages</h1>
  <p>Importez votre fichier de notes, filtrez par mention et générez les mails de convocation.</p>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader(
    "📂 Importer un fichier Excel (.xlsx)",
    type=["xlsx"],
    help="Le fichier doit contenir une colonne 'Personne' et des colonnes commençant par 'Eval'.",
)

if uploaded is None:
    st.info("⬆️ Veuillez importer un fichier Excel pour commencer.")
    st.stop()

# ── Lecture ──────────────────────────────────────────────────────────────────────
try:
    raw_df = pd.read_excel(uploaded)
except Exception as e:
    st.error(f"Impossible de lire le fichier : {e}")
    st.stop()

# ── Détection colonnes ────────────────────────────────────────────────────────────
if "Personne" not in raw_df.columns:
    st.error("Colonne 'Personne' introuvable. Vérifiez votre fichier.")
    st.stop()

eval_cols = [c for c in raw_df.columns if str(c).startswith("Eval")]
if not eval_cols:
    st.error("Aucune colonne commençant par 'Eval' trouvée dans le fichier.")
    st.stop()

# ── Nettoyage : garder Personne + Eval, lignes non vides ─────────────────────────
working = raw_df[["Personne"] + eval_cols].copy()
working = working[working[eval_cols].notna().any(axis=1)].reset_index(drop=True)

# Séparer prénom / nom
working.insert(0, "Prénom", working["Personne"].apply(lambda x: split_name(x)[0]))
working.insert(1, "Nom",    working["Personne"].apply(lambda x: split_name(x)[1]))
working = working.drop(columns=["Personne"])

# Renommer colonnes Eval → nom court pour l'affichage
display_cols = {c: short_eval_name(c) for c in eval_cols}
display_df   = working.rename(columns=display_cols)

st.success(f"✅ {len(working)} étudiant(s) chargé(s) — {len(eval_cols)} matière(s) détectée(s).")

# ─── FILTRES ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">🔍 Filtres</div>', unsafe_allow_html=True)

col1, col2 = st.columns([1, 1])

with col1:
    st.markdown("""
    <div class="legend-row">
      <div class="legend-item"><span class="dot dot-A"></span> A – Admis</div>
      <div class="legend-item"><span class="dot dot-B"></span> B – Bien</div>
      <div class="legend-item"><span class="dot dot-C"></span> C – Ajourné léger</div>
      <div class="legend-item"><span class="dot dot-D"></span> D – Ajourné</div>
    </div>
    """, unsafe_allow_html=True)

    selected_grades = st.multiselect(
        "Filtrer les étudiants ayant au moins une mention :",
        options=["A", "B", "C", "D"],
        default=["C", "D"],
    )

with col2:
    group_filter = st.radio(
        "Groupe de mentions :",
        options=[
            "Tous (pas de filtre groupe)",
            "A ou B uniquement (admis / bien)",
            "C ou D uniquement (à rattraper)",
        ],
        index=2,
    )

# ── Application des filtres ───────────────────────────────────────────────────────
def filter_df(df: pd.DataFrame, grades: list, group: str) -> pd.DataFrame:
    eval_display = [display_cols[c] for c in eval_cols]

    # Filtre mentions
    if grades:
        mask = df[eval_display].isin(grades).any(axis=1)
        df = df[mask]

    # Filtre groupe
    if group == "A ou B uniquement (admis / bien)":
        mask2 = df[eval_display].isin(["A", "B"]).any(axis=1) & \
               ~df[eval_display].isin(["C", "D"]).any(axis=1)
        df = df[mask2]
    elif group == "C ou D uniquement (à rattraper)":
        mask2 = df[eval_display].isin(["C", "D"]).any(axis=1)
        df = df[mask2]

    return df

filtered_df = filter_df(display_df.copy(), selected_grades, group_filter)

st.markdown(f"**{len(filtered_df)}** étudiant(s) correspondent aux critères.")

# ─── TABLEAU RÉSULTAT ────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">📋 Tableau des résultats</div>', unsafe_allow_html=True)

# Affichage HTML colorisé
if not filtered_df.empty:
    html_rows = ""
    eval_display_cols = [display_cols[c] for c in eval_cols]

    # En-tête
    headers = "<tr>" + "".join(
        f'<th style="background:#4f46e5;color:white;padding:8px 10px;'
        f'font-size:0.78rem;text-align:center;white-space:nowrap;">{h}</th>'
        for h in filtered_df.columns
    ) + "</tr>"

    for _, row in filtered_df.iterrows():
        cells = ""
        for col in filtered_df.columns:
            val = row[col]
            if col in ("Prénom", "Nom"):
                cells += (
                    f'<td style="padding:6px 10px;font-weight:600;'
                    f'white-space:nowrap;font-size:0.85rem;">{val}</td>'
                )
            else:
                b = badge(val)
                cells += f'<td style="text-align:center;padding:4px 8px;">{b}</td>'
        alt = 'background:#f8fafc' if _ % 2 == 0 else 'background:white'
        html_rows += f'<tr style="{alt}">{cells}</tr>'

    table_html = f"""
    <div style="overflow-x:auto;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,0.08);">
    <table style="width:100%;border-collapse:collapse;font-family:sans-serif;">
      <thead>{headers}</thead>
      <tbody>{html_rows}</tbody>
    </table>
    </div>"""
    st.markdown(table_html, unsafe_allow_html=True)
else:
    st.warning("Aucun étudiant ne correspond aux filtres sélectionnés.")

# ─── EXPORT EXCEL ────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">📥 Export Excel</div>', unsafe_allow_html=True)

if not filtered_df.empty:
    excel_bytes = df_to_excel_bytes(filtered_df)
    st.download_button(
        label="⬇️ Télécharger le tableau filtré (.xlsx)",
        data=excel_bytes,
        file_name="rattrapages_filtrés.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ─── GÉNÉRATION DES MAILS ────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">✉️ Mails de convocation aux rattrapages</div>', unsafe_allow_html=True)

if filtered_df.empty:
    st.info("Aucun étudiant sélectionné — ajustez les filtres.")
else:
    tutoyer = st.toggle("Tutoyer les étudiants (sinon vouvoiement)", value=False)

    # On considère que les matières de rattrapage = celles avec C ou D
    eval_display_cols = [display_cols[c] for c in eval_cols]

    for _, row in filtered_df.iterrows():
        prenom = row["Prénom"]
        nom    = row["Nom"]

        # Matières C ou D
        matieres_rattrapage = [
            col for col in eval_display_cols
            if str(row.get(col, "")).strip() in ("C", "D")
        ]

        if not matieres_rattrapage:
            continue  # pas de rattrapage pour cet étudiant

        with st.expander(f"📧 {prenom} {nom}  —  {len(matieres_rattrapage)} matière(s)"):
            mail_text = generate_email(prenom, nom, matieres_rattrapage, tutoyer)

            # Affichage coloré des matières
            badges_html = " ".join(
                f'<span style="background:#fee2e2;color:#7f1d1d;border:1px solid #fca5a5;'
                f'border-radius:6px;padding:2px 8px;font-size:0.78rem;font-weight:600;">{m}</span>'
                for m in matieres_rattrapage
            )
            st.markdown(f"**Matières concernées :** {badges_html}", unsafe_allow_html=True)
            st.markdown("")

            edited_mail = st.text_area(
                "Contenu du mail (modifiable) :",
                value=mail_text,
                height=280,
                key=f"mail_{prenom}_{nom}",
            )
            st.download_button(
                label="⬇️ Télécharger ce mail (.txt)",
                data=edited_mail.encode("utf-8"),
                file_name=f"mail_{prenom}_{nom}.txt".replace(" ", "_"),
                mime="text/plain",
                key=f"dl_{prenom}_{nom}",
            )
