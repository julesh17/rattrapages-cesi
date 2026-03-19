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

.legend-row {
    display:flex; gap:10px; align-items:center; flex-wrap:wrap;
    margin-bottom: 0.5rem;
}
.legend-item { display:flex; align-items:center; gap:6px; font-size:0.85rem; font-weight:600; }
.dot { width:14px; height:14px; border-radius:50%; display:inline-block; }
.dot-A {background:#10b981;} .dot-B {background:#3b82f6;}
.dot-C {background:#f59e0b;} .dot-D {background:#ef4444;}

[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

.section-title {
    font-size: 1.1rem; font-weight: 700;
    color: #4f46e5; margin-bottom: 0.4rem;
    display: flex; align-items: center; gap: 8px;
}
</style>
""", unsafe_allow_html=True)

# ─── HELPERS ────────────────────────────────────────────────────────────────────

def badge(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    v = str(val).strip()
    css = f"badge badge-{v}" if v in "ABCD" else "badge"
    return f'<span class="{css}">{v}</span>'


def split_name(personne: str):
    """'NOM, Prénom' -> (prénom, nom)"""
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


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Résultats")
        ws = writer.sheets["Résultats"]
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        header_fill = PatternFill("solid", fgColor="4F46E5")
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        fill_map = {
            "A": PatternFill("solid", fgColor="D1FAE5"),
            "B": PatternFill("solid", fgColor="DBEAFE"),
            "C": PatternFill("solid", fgColor="FEF3C7"),
            "D": PatternFill("solid", fgColor="FEE2E2"),
        }
        thin   = Side(style="thin", color="D1D5DB")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if cell.value in fill_map:
                    cell.fill = fill_map[cell.value]

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

try:
    raw_df = pd.read_excel(uploaded)
except Exception as e:
    st.error(f"Impossible de lire le fichier : {e}")
    st.stop()

if "Personne" not in raw_df.columns:
    st.error("Colonne 'Personne' introuvable. Vérifiez votre fichier.")
    st.stop()

all_eval_cols = [c for c in raw_df.columns if str(c).startswith("Eval")]
if not all_eval_cols:
    st.error("Aucune colonne commençant par 'Eval' trouvée dans le fichier.")
    st.stop()

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

    # Pré-exclure Global exam si le toggle est activé
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

# Fusion des deux exclusions
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

nb_excluded = len(excluded_full)
st.success(
    f"✅ {len(working)} étudiant(s) chargé(s) — "
    f"{len(eval_cols)} matière(s) active(s)"
    + (f" ({nb_excluded} exclue(s))" if nb_excluded else "") + "."
)

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
    if grades:
        df = df[df[cols].isin(grades).any(axis=1)]
    if group == "A ou B uniquement (admis / bien)":
        df = df[df[cols].isin(["A", "B"]).any(axis=1) & ~df[cols].isin(["C", "D"]).any(axis=1)]
    elif group == "C ou D uniquement (à rattraper)":
        df = df[df[cols].isin(["C", "D"]).any(axis=1)]
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

# ─── EXPORT EXCEL ────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">📥 Export Excel</div>', unsafe_allow_html=True)

if not filtered_df.empty:
    st.download_button(
        label="⬇️ Télécharger le tableau filtré (.xlsx)",
        data=df_to_excel_bytes(filtered_df),
        file_name="rattrapages_filtrés.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ─── RÉCAP MATIÈRES CONCERNÉES ──────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">📊 Récapitulatif des matières en rattrapage</div>', unsafe_allow_html=True)

if not filtered_df.empty:
    # Pour chaque matière : compter les C et D séparément
    recap_rows = []
    for col in eval_display_cols:
        nb_c = (filtered_df[col] == "C").sum()
        nb_d = (filtered_df[col] == "D").sum()
        total = nb_c + nb_d
        if total > 0:
            recap_rows.append({"Matière": col, "nb_c": nb_c, "nb_d": nb_d, "total": total})

    if not recap_rows:
        st.info("Aucune matière avec C ou D pour les étudiants sélectionnés.")
    else:
        recap_rows.sort(key=lambda x: -x["total"])

        # Tableau HTML récap
        recap_headers = "<tr>" + "".join(
            f'<th style="background:#4f46e5;color:white;padding:8px 14px;font-size:0.82rem;text-align:{a};">{h}</th>'
            for h, a in [("Matière", "left"), ("C", "center"), ("D", "center"), ("Total", "center")]
        ) + "</tr>"

        recap_html_rows = ""
        for i, r in enumerate(recap_rows):
            bg = "#f8fafc" if i % 2 == 0 else "white"
            bar_width = int(r["total"] / recap_rows[0]["total"] * 100)
            bar_c = int(r["nb_c"] / r["total"] * bar_width) if r["total"] else 0
            bar_d = bar_width - bar_c

            bar = (
                f'<div style="display:flex;height:8px;border-radius:4px;overflow:hidden;'
                f'width:{bar_width}%;min-width:4px;margin-top:4px;">'
                f'<div style="width:{bar_c}px;flex:{bar_c} 0 0;background:#f59e0b;"></div>'
                f'<div style="width:{bar_d}px;flex:{bar_d} 0 0;background:#ef4444;"></div>'
                f'</div>'
            )
            badge_c = f'<span class="badge badge-C">{r["nb_c"]}</span>' if r["nb_c"] else '<span style="color:#ccc">—</span>'
            badge_d = f'<span class="badge badge-D">{r["nb_d"]}</span>' if r["nb_d"] else '<span style="color:#ccc">—</span>'
            badge_t = f'<strong style="font-size:0.95rem;">{r["total"]}</strong>'

            recap_html_rows += f"""
            <tr style="background:{bg}">
              <td style="padding:8px 14px;font-size:0.85rem;font-weight:600;">
                {r["Matière"]}{bar}
              </td>
              <td style="text-align:center;padding:6px 10px;">{badge_c}</td>
              <td style="text-align:center;padding:6px 10px;">{badge_d}</td>
              <td style="text-align:center;padding:6px 10px;">{badge_t}</td>
            </tr>"""

        st.markdown(f"""
        <div style="overflow-x:auto;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,0.08);max-width:820px;">
        <table style="width:100%;border-collapse:collapse;font-family:sans-serif;">
          <thead>{recap_headers}</thead><tbody>{recap_html_rows}</tbody>
        </table></div>""", unsafe_allow_html=True)

        # Petite ligne de synthèse
        total_cd  = sum(r["total"] for r in recap_rows)
        nb_mat    = len(recap_rows)
        st.markdown(
            f"<p style='margin-top:0.7rem;font-size:0.85rem;color:#6b7280;'>"
            f"<strong>{nb_mat}</strong> matière(s) concernée(s) · "
            f"<strong>{total_cd}</strong> situation(s) de rattrapage au total "
            f"(<span style='color:#f59e0b;font-weight:700;'>C : {sum(r['nb_c'] for r in recap_rows)}</span> · "
            f"<span style='color:#ef4444;font-weight:700;'>D : {sum(r['nb_d'] for r in recap_rows)}</span>)"
            f"</p>",
            unsafe_allow_html=True,
        )

# ─── MAILS ───────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">✉️ Mails de convocation aux rattrapages</div>', unsafe_allow_html=True)

if filtered_df.empty:
    st.info("Aucun étudiant sélectionné — ajustez les filtres.")
else:
    # Toggle HORS des expanders : un changement régénère toutes les clés des text_area
    tutoyer    = st.toggle("👋 Tutoyer les étudiants (sinon vouvoiement)", value=False)
    toggle_key = "tu" if tutoyer else "vous"

    students_with_rattrapage = [
        (row["Prénom"], row["Nom"],
         [c for c in eval_display_cols if str(row.get(c, "")).strip() in ("C", "D")])
        for _, row in filtered_df.iterrows()
    ]
    students_with_rattrapage = [(p, n, m) for p, n, m in students_with_rattrapage if m]

    if not students_with_rattrapage:
        st.info("Aucun étudiant avec C ou D dans les matières actives.")
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
                    # La clé intègre toggle_key → Streamlit recharge value à chaque changement
                    key=f"mail_{toggle_key}_{prenom}_{nom}",
                )
                st.download_button(
                    label="⬇️ Télécharger ce mail (.txt)",
                    data=edited_mail.encode("utf-8"),
                    file_name=f"mail_{prenom}_{nom}.txt".replace(" ", "_"),
                    mime="text/plain",
                    key=f"dl_{toggle_key}_{prenom}_{nom}",
                )
