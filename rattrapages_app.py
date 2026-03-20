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
.badge-ABS { background:#ffedd5; color:#7c2d12; border:1px solid #fb923c; }

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
</style>
""", unsafe_allow_html=True)

# ─── HELPERS ────────────────────────────────────────────────────────────────────

def badge(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    v = str(val).strip()
    css = f"badge badge-{v}" if v in ("A","B","C","D","ABS") else "badge"
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

# Appliquer la logique ABS uniquement si le toggle est activé
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
    # ABS is always treated as requiring rattrapage (same as C/D)
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
    recap_rows = []
    for col in eval_display_cols:
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
        total = len(eleves_c) + len(eleves_d) + len(eleves_abs)
        if total > 0:
            recap_rows.append({
                "Matière": col,
                "eleves_c": eleves_c, "nb_c": len(eleves_c),
                "eleves_d": eleves_d, "nb_d": len(eleves_d),
                "eleves_abs": eleves_abs, "nb_abs": len(eleves_abs),
                "total": total,
            })

    if not recap_rows:
        st.info("Aucune matière avec C ou D pour les étudiants sélectionnés.")
    else:
        recap_rows.sort(key=lambda x: -x["total"])

        # Ligne de synthèse
        total_cd = sum(r["total"] for r in recap_rows)
        nb_mat   = len(recap_rows)
        st.markdown(
            f"<p style='font-size:0.85rem;color:#6b7280;margin-bottom:0.8rem;'>"
            f"<strong>{nb_mat}</strong> matière(s) concernée(s) · "
            f"<strong>{total_cd}</strong> situation(s) de rattrapage "
            f"(<span style='color:#f59e0b;font-weight:700;'>C : {sum(r['nb_c'] for r in recap_rows)}</span> · "
            f"<span style='color:#ef4444;font-weight:700;'>D : {sum(r['nb_d'] for r in recap_rows)}</span> · "
            f"<span style='color:#f97316;font-weight:700;'>ABS : {sum(r.get('nb_abs',0) for r in recap_rows)}</span>)"
            f"</p>",
            unsafe_allow_html=True,
        )

        # Carte par matière avec élèves
        for i, r in enumerate(recap_rows):
            bg = "#f8fafc" if i % 2 == 0 else "white"

            def eleve_pill(name, mention):
                color = {
                    "C":   ("#fef3c7", "#78350f", "#fcd34d"),
                    "D":   ("#fee2e2", "#7f1d1d", "#fca5a5"),
                    "ABS": ("#ffedd5", "#7c2d12", "#fb923c"),
                }[mention]
                return (
                    f'<span style="display:inline-block;background:{color[0]};color:{color[1]};'
                    f'border:1px solid {color[2]};border-radius:20px;padding:1px 10px;'
                    f'font-size:0.76rem;font-weight:600;margin:2px;">{name}</span>'
                )

            pills_c = "".join(eleve_pill(e, "C") for e in r["eleves_c"])
            pills_d = "".join(eleve_pill(e, "D") for e in r["eleves_d"])
            pills_abs = "".join(eleve_pill(e, "ABS") for e in r.get("eleves_abs", []))

            badge_c   = f'<span class="badge badge-C">{r["nb_c"]}</span>' if r["nb_c"] else ""
            badge_d   = f'<span class="badge badge-D">{r["nb_d"]}</span>' if r["nb_d"] else ""
            badge_abs = f'<span class="badge badge-ABS">{r["nb_abs"]}</span>' if r.get("nb_abs") else ""

            bar_width = int(r["total"] / recap_rows[0]["total"] * 100)
            bar_c   = int(r["nb_c"]   / r["total"] * bar_width) if r["total"] else 0
            bar_d   = int(r["nb_d"]   / r["total"] * bar_width) if r["total"] else 0
            bar_abs = bar_width - bar_c - bar_d
            bar = (
                f'<div style="display:flex;height:6px;border-radius:4px;overflow:hidden;'
                f'width:{bar_width}%;min-width:4px;margin-top:5px;">'
                f'<div style="flex:{bar_c} 0 0;background:#f59e0b;"></div>'
                f'<div style="flex:{bar_d} 0 0;background:#ef4444;"></div>'
                f'<div style="flex:{bar_abs} 0 0;background:#f97316;"></div>'
                f'</div>'
            )

            st.markdown(f"""
            <div style="background:{bg};border-radius:10px;padding:10px 16px;
                        margin-bottom:6px;box-shadow:0 1px 4px rgba(0,0,0,0.05);">
              <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
                <div style="flex:1;min-width:200px;">
                  <span style="font-weight:700;font-size:0.88rem;">{r["Matière"]}</span>
                  {bar}
                </div>
                <div style="display:flex;gap:5px;align-items:center;">
                  {badge_c}{badge_d}{badge_abs}
                  <span style="font-size:0.78rem;color:#6b7280;margin-left:2px;">/ {r["total"]}</span>
                </div>
              </div>
              <div style="margin-top:7px;line-height:2;">{pills_c}{pills_d}{pills_abs}</div>
            </div>""", unsafe_allow_html=True)

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
         [c for c in eval_display_cols if str(row.get(c, "")).strip() in ("C", "D", "ABS")])
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
                    # Bouton copier via JS — encode le texte en base64 pour éviter les problèmes de quotes
                    import base64
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
    # Construire le texte du récap
    lines = ["Bonjour à tous,",
             "",
             "Voici le récapitulatif des rattrapages par matière :",
             ""]

    for r in recap_rows:
        all_students = r["eleves_c"] + r["eleves_d"] + r.get("eleves_abs", [])
        noms_liste   = ", ".join(all_students)
        lines.append(f"• {r['Matière']} : {noms_liste}")

    lines += [
        "",
        "Les étudiants concernés sont invités à se présenter aux sessions de rattrapage "
        "dont les modalités leur seront communiquées prochainement.",
        "",
        "Bien cordialement,",
        "L'équipe pédagogique",
    ]

    recap_classe_text = "\n".join(lines)

    # Affichage + édition
    edited_recap = st.text_area(
        "Contenu du mail (modifiable) :",
        value=recap_classe_text,
        height=320,
        key="recap_classe_textarea",
    )

    import base64
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
    "si elles n'ont <strong>aucun élève en commun</strong> parmi les convoqués (C, D ou ABS)."
    "</p>", unsafe_allow_html=True
)

if not filtered_df.empty and recap_rows:
    import itertools

    # Construire le dictionnaire matière → ensemble d'élèves convoqués
    rattrapage_vals = {"C", "D", "ABS"}
    mat_students = {}
    for r in recap_rows:
        eleves = set(r["eleves_c"] + r["eleves_d"] + r.get("eleves_abs", []))
        if eleves:
            mat_students[r["Matière"]] = eleves

    matieres_list = list(mat_students.keys())

    if len(matieres_list) < 2:
        st.info("Pas assez de matières avec rattrapages pour calculer des compatibilités.")
    else:
        # Construire les groupes compatibles par coloration gloutonne (graph coloring greedy)
        # Deux matières incompatibles = elles partagent au moins un élève
        def sont_compatibles(m1, m2):
            return mat_students[m1].isdisjoint(mat_students[m2])

        # Groupes : liste de listes de matières pouvant toutes se tenir en même temps
        groupes = []
        reste = list(matieres_list)
        while reste:
            groupe = [reste[0]]
            for m in reste[1:]:
                if all(sont_compatibles(m, g) for g in groupe):
                    groupe.append(m)
            groupes.append(groupe)
            reste = [m for m in reste if m not in groupe]

        # Couleurs de créneaux
        slot_colors = [
            ("#ede9fe", "#4f46e5", "#c4b5fd"),  # violet
            ("#d1fae5", "#065f46", "#6ee7b7"),  # vert
            ("#dbeafe", "#1e3a8a", "#93c5fd"),  # bleu
            ("#fef3c7", "#78350f", "#fcd34d"),  # jaune
            ("#fee2e2", "#7f1d1d", "#fca5a5"),  # rouge
            ("#f3e8ff", "#581c87", "#d8b4fe"),  # mauve
            ("#ffedd5", "#7c2d12", "#fb923c"),  # orange
        ]

        st.markdown(
            f"<p style='font-size:0.9rem;'><strong>{len(groupes)} créneau(x) minimum</strong> "
            f"nécessaire(s) pour organiser tous les rattrapages sans conflit.</p>",
            unsafe_allow_html=True
        )

        for i, groupe in enumerate(groupes):
            bg, fg, border = slot_colors[i % len(slot_colors)]

            # Élèves totaux du créneau (union)
            all_eleves_creneau = set()
            for m in groupe:
                all_eleves_creneau |= mat_students[m]

            # Lignes matière + élèves
            lignes_html = ""
            for m in groupe:
                eleves_m = sorted(mat_students[m])
                pills = "".join(
                    f'<span style="display:inline-block;background:white;color:{fg};'
                    f'border:1px solid {border};border-radius:20px;padding:1px 9px;'
                    f'font-size:0.74rem;font-weight:600;margin:2px;">{e}</span>'
                    for e in eleves_m
                )
                lignes_html += f"""
                <div style="margin-bottom:8px;">
                  <span style="font-weight:700;font-size:0.85rem;">{m}</span>
                  <span style="font-size:0.78rem;color:{fg};opacity:0.8;margin-left:6px;">
                    ({len(eleves_m)} élève(s))
                  </span>
                  <div style="margin-top:4px;">{pills}</div>
                </div>"""

            nb_mat_creneau   = len(groupe)
            nb_elev_creneau  = len(all_eleves_creneau)

            st.markdown(f"""
            <div style="background:{bg};border:1.5px solid {border};border-radius:12px;
                        padding:14px 18px;margin-bottom:10px;
                        box-shadow:0 2px 8px rgba(0,0,0,0.06);">
              <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">
                <span style="background:{fg};color:white;border-radius:8px;
                             padding:3px 12px;font-weight:800;font-size:0.9rem;">
                  Créneau {i+1}
                </span>
                <span style="font-size:0.82rem;color:{fg};font-weight:600;">
                  {nb_mat_creneau} matière(s) · {nb_elev_creneau} élève(s) au total
                </span>
              </div>
              {lignes_html}
            </div>""", unsafe_allow_html=True)

        # Matrice de compatibilité
        with st.expander("🔍 Voir la matrice de compatibilité complète"):
            n = len(matieres_list)
            # En-tête
            th = "".join(
                f'<th style="background:#4f46e5;color:white;padding:5px 8px;'
                f'font-size:0.7rem;text-align:center;white-space:nowrap;'
                f'writing-mode:vertical-rl;transform:rotate(180deg);max-width:28px;">'
                f'{m[:25]}</th>'
                for m in matieres_list
            )
            header_row = f'<tr><th style="background:#4f46e5;"></th>{th}</tr>'

            rows_html = ""
            for i_m, m1 in enumerate(matieres_list):
                cells = f'<td style="font-weight:700;font-size:0.75rem;padding:4px 8px;white-space:nowrap;background:#f8fafc;">{m1[:30]}</td>'
                for j_m, m2 in enumerate(matieres_list):
                    if i_m == j_m:
                        cells += '<td style="background:#e2e8f0;text-align:center;">—</td>'
                    elif sont_compatibles(m1, m2):
                        cells += '<td style="background:#d1fae5;color:#065f46;text-align:center;font-weight:700;font-size:0.8rem;">✓</td>'
                    else:
                        nb_communs = len(mat_students[m1] & mat_students[m2])
                        cells += f'<td style="background:#fee2e2;color:#7f1d1d;text-align:center;font-size:0.75rem;font-weight:600;">{nb_communs}</td>'
                bg_row = "#f8fafc" if i_m % 2 == 0 else "white"
                rows_html += f'<tr style="background:{bg_row}">{cells}</tr>'

            st.markdown(f"""
            <p style="font-size:0.8rem;color:#6b7280;margin-bottom:0.5rem;">
              ✓ = compatible (0 élève en commun) · chiffre = nb d'élèves en conflit
            </p>
            <div style="overflow-x:auto;border-radius:10px;box-shadow:0 1px 6px rgba(0,0,0,0.08);">
            <table style="border-collapse:collapse;font-family:sans-serif;">
              <thead>{header_row}</thead><tbody>{rows_html}</tbody>
            </table></div>""", unsafe_allow_html=True)
else:
    st.info("Aucune donnée disponible — importez un fichier et appliquez les filtres.")
