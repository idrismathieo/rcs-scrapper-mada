"""Application web Streamlit pour le scraper RCS Madagascar.

Importe la logique de scraping depuis ../rcs_scraper.py (aucune duplication).
Lancer en local : streamlit run streamlit_app.py
"""

import sys
import time
import traceback
from datetime import datetime
from io import BytesIO
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import requests  # noqa: E402
import streamlit as st  # noqa: E402

from rcs_scraper import (  # noqa: E402
    FORMES_LABELS,
    GREFFE_MAP,
    TYPE_LABELS,
    build_workbook,
    extract_entreprises,
    format_duration,
    make_session,
)


st.set_page_config(
    page_title="RCS Madagascar",
    page_icon=None,
    layout="centered",
    initial_sidebar_state="collapsed",
    menu_items={
        "Get help": None,
        "Report a bug": None,
        "About": "Extracteur d'entreprises du RCS Madagascar.",
    },
)


CSS = """
<style>
    /* === Reset & layout === */
    .block-container {
        padding-top: 3rem;
        padding-bottom: 1rem;
        max-width: 820px;
    }

    #MainMenu, header[data-testid="stHeader"], footer[data-testid="stFooter"] {
        visibility: hidden;
        height: 0;
    }

    html, body, [class*="st-"] {
        font-family: -apple-system, BlinkMacSystemFont, "Inter", "Segoe UI", Roboto, sans-serif;
    }

    /* === Header === */
    .hero {
        margin-bottom: 2.5rem;
    }
    .hero-title {
        font-size: 2.4rem;
        font-weight: 800;
        color: #0F172A;
        letter-spacing: -1px;
        line-height: 1.15;
        margin: 0;
    }
    .hero-subtitle {
        font-size: 1rem;
        color: #475569;
        margin-top: 0.6rem;
        line-height: 1.5;
    }

    /* === Section title === */
    .section-title {
        font-size: 0.78rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        color: #64748B;
        margin: 2.4rem 0 1rem 0;
        padding-bottom: 0.6rem;
        border-bottom: 1px solid #E2E8F0;
    }

    /* === Form card === */
    div[data-testid="stForm"] {
        background: #FFFFFF;
        padding: 1.8rem !important;
        border-radius: 14px;
        border: 1px solid #E5E7EB;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04), 0 4px 12px rgba(15, 23, 42, 0.04);
    }

    /* === Form inputs === */
    .stSelectbox label, .stNumberInput label {
        font-weight: 600 !important;
        color: #1E293B !important;
        font-size: 0.9rem !important;
        margin-bottom: 0.3rem;
    }
    .stSelectbox div[data-baseweb="select"] > div,
    .stNumberInput input {
        border-radius: 8px !important;
        border: 1px solid #E2E8F0 !important;
        font-size: 0.95rem !important;
    }
    .stSelectbox div[data-baseweb="select"] > div:hover,
    .stNumberInput input:hover {
        border-color: #CBD5E1 !important;
    }

    /* === Primary button === */
    .stFormSubmitButton > button,
    .stButton > button[kind="primary"] {
        background: linear-gradient(180deg, #4F46E5 0%, #4338CA 100%) !important;
        color: white !important;
        font-weight: 600 !important;
        padding: 0.75rem 1.5rem !important;
        border-radius: 10px !important;
        border: none !important;
        font-size: 0.98rem !important;
        width: 100%;
        box-shadow: 0 1px 2px rgba(67, 56, 202, 0.2), 0 2px 8px rgba(67, 56, 202, 0.15);
        transition: all 0.15s ease;
    }
    .stFormSubmitButton > button:hover,
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(180deg, #6366F1 0%, #4F46E5 100%) !important;
        box-shadow: 0 2px 4px rgba(67, 56, 202, 0.25), 0 4px 16px rgba(67, 56, 202, 0.2);
        transform: translateY(-1px);
    }
    .stFormSubmitButton > button:active,
    .stButton > button[kind="primary"]:active {
        transform: translateY(0);
    }

    /* === Download button === */
    .stDownloadButton > button {
        background: linear-gradient(180deg, #059669 0%, #047857 100%) !important;
        color: white !important;
        font-weight: 600 !important;
        padding: 0.85rem 1.5rem !important;
        border-radius: 10px !important;
        border: none !important;
        font-size: 1rem !important;
        width: 100%;
        box-shadow: 0 1px 2px rgba(5, 150, 105, 0.2), 0 2px 8px rgba(5, 150, 105, 0.15);
        transition: all 0.15s ease;
    }
    .stDownloadButton > button:hover {
        background: linear-gradient(180deg, #10B981 0%, #059669 100%) !important;
        box-shadow: 0 2px 4px rgba(5, 150, 105, 0.25), 0 4px 16px rgba(5, 150, 105, 0.2);
        transform: translateY(-1px);
    }

    /* === Progress card === */
    .progress-card {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 14px;
        padding: 1.6rem;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }
    .progress-header {
        display: flex;
        justify-content: space-between;
        align-items: baseline;
        margin-bottom: 1rem;
    }
    .progress-phase {
        font-size: 0.95rem;
        font-weight: 600;
        color: #0F172A;
    }
    .progress-detail {
        font-size: 0.85rem;
        color: #64748B;
        margin-top: 4px;
    }
    .progress-pct {
        font-size: 1.5rem;
        font-weight: 700;
        color: #4F46E5;
        font-variant-numeric: tabular-nums;
    }

    .stProgress > div > div > div {
        background: linear-gradient(90deg, #6366F1 0%, #4F46E5 100%) !important;
        border-radius: 999px;
    }
    .stProgress > div > div {
        background-color: #EEF2FF !important;
        border-radius: 999px;
        height: 8px !important;
    }

    /* === Stepper === */
    .stepper {
        display: flex;
        gap: 8px;
        margin: 1.2rem 0 1.5rem 0;
    }
    .step {
        flex: 1;
        padding: 10px 14px;
        border-radius: 10px;
        background: #F1F5F9;
        border: 1px solid transparent;
        font-size: 0.82rem;
        color: #94A3B8;
        font-weight: 500;
        display: flex;
        align-items: center;
        gap: 8px;
        transition: all 0.2s;
    }
    .step .step-num {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 22px;
        height: 22px;
        border-radius: 50%;
        background: #E2E8F0;
        color: #64748B;
        font-size: 0.75rem;
        font-weight: 700;
    }
    .step.active {
        background: #EEF2FF;
        border-color: #C7D2FE;
        color: #4338CA;
    }
    .step.active .step-num {
        background: #4F46E5;
        color: white;
    }
    .step.done {
        background: #ECFDF5;
        border-color: #A7F3D0;
        color: #047857;
    }
    .step.done .step-num {
        background: #059669;
        color: white;
    }

    /* === Result card === */
    .result-card {
        background: linear-gradient(135deg, #ECFDF5 0%, #F0FDFA 100%);
        border: 1px solid #A7F3D0;
        border-radius: 14px;
        padding: 1.6rem;
        margin-bottom: 1rem;
    }
    .result-header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 0.8rem;
    }
    .result-check {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 32px;
        height: 32px;
        border-radius: 50%;
        background: #059669;
        color: white;
        font-weight: bold;
        font-size: 1rem;
    }
    .result-title {
        font-size: 1.1rem;
        font-weight: 700;
        color: #064E3B;
        margin: 0;
    }
    .result-meta {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 1rem;
        margin: 1.2rem 0;
    }
    .meta-item {
        background: white;
        padding: 0.9rem;
        border-radius: 10px;
        border: 1px solid #D1FAE5;
    }
    .meta-label {
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        color: #047857;
        font-weight: 600;
    }
    .meta-value {
        font-size: 1.4rem;
        font-weight: 700;
        color: #064E3B;
        font-variant-numeric: tabular-nums;
        margin-top: 4px;
    }
    .meta-value-sub {
        font-size: 0.8rem;
        color: #047857;
        margin-top: 2px;
    }

    /* === Error card === */
    .error-card {
        background: #FEF2F2;
        border: 1px solid #FECACA;
        border-radius: 14px;
        padding: 1.5rem;
        margin-bottom: 1rem;
    }
    .error-header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 0.6rem;
    }
    .error-icon {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 32px;
        height: 32px;
        border-radius: 50%;
        background: #DC2626;
        color: white;
        font-weight: bold;
    }
    .error-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #7F1D1D;
        margin: 0;
    }
    .error-message {
        color: #991B1B;
        font-size: 0.95rem;
        line-height: 1.5;
        margin: 0.6rem 0;
    }
    .error-hint {
        background: white;
        border-left: 3px solid #DC2626;
        padding: 0.7rem 1rem;
        margin-top: 0.8rem;
        font-size: 0.88rem;
        color: #7F1D1D;
        border-radius: 4px;
    }

    /* === Empty state hint === */
    .empty-hint {
        text-align: center;
        color: #94A3B8;
        font-size: 0.88rem;
        padding: 2rem 1rem;
    }

    /* === Footer === */
    .app-footer {
        text-align: center;
        color: #94A3B8;
        font-size: 12px;
        padding: 32px 0 12px 0;
        margin-top: 80px;
        border-top: 1px solid #E2E8F0;
        letter-spacing: 0.3px;
    }
    .app-footer .footer-divider {
        margin: 0 8px;
        color: #CBD5E1;
    }
</style>
"""

st.markdown(CSS, unsafe_allow_html=True)


# === Header ===
st.markdown(
    """
<div class="hero">
    <h1 class="hero-title">Extracteur RCS Madagascar</h1>
    <p class="hero-subtitle">
        Récupérez en un clic la liste complète des entreprises immatriculées au
        Registre du Commerce et des Sociétés, exportée au format Excel prêt à l'emploi.
    </p>
</div>
""",
    unsafe_allow_html=True,
)


# === Session state ===
if "result" not in st.session_state:
    st.session_state.result = None
if "error" not in st.session_state:
    st.session_state.error = None
if "running" not in st.session_state:
    st.session_state.running = False


# === Form ===
st.markdown('<div class="section-title">Critères de recherche</div>', unsafe_allow_html=True)

greffes_sorted = sorted(GREFFE_MAP.keys())
type_keys = list(TYPE_LABELS.keys())
type_labels = list(TYPE_LABELS.values())

with st.form("scraper_form", clear_on_submit=False, border=False):
    current_year = datetime.now().year
    annee_options = list(range(current_year + 1, 1999, -1))

    col1, col2 = st.columns(2)
    with col1:
        greffe = st.selectbox(
            "Ville (greffe)",
            options=greffes_sorted,
            index=greffes_sorted.index("Nosy Be"),
            help="La ville d'immatriculation au RCS",
        )
    with col2:
        annee = st.selectbox(
            "Année",
            options=annee_options,
            index=annee_options.index(current_year),
            help="Année d'immatriculation",
        )

    col3, col4 = st.columns(2)
    with col3:
        forme = st.selectbox(
            "Forme juridique",
            options=list(FORMES_LABELS.keys()),
            format_func=lambda code: f"{code} — {FORMES_LABELS[code]}",
            help="Codes officiels du RCS Madagascar",
        )
    with col4:
        type_label = st.selectbox(
            "Type d'assujetti",
            options=type_labels,
            index=type_keys.index("B"),
            help="Personne physique, morale, GIE, etc.",
        )
    type_code = type_keys[type_labels.index(type_label)]

    st.write("")
    launch = st.form_submit_button("Lancer l'extraction", type="primary", use_container_width=True)


# === Helpers ===
def classify_error(exc: Exception):
    """Retourne (titre, message_user, conseil) en fonction du type d'erreur."""
    if isinstance(exc, requests.exceptions.ConnectionError):
        return (
            "Connexion impossible",
            "Impossible de joindre le serveur RCS Madagascar.",
            "Vérifiez votre connexion Internet, puis réessayez. Si le problème persiste, le site rcsmada.mg est peut-être temporairement hors ligne.",
        )
    if isinstance(exc, requests.exceptions.Timeout):
        return (
            "Temps de réponse dépassé",
            "Le serveur RCS met trop de temps à répondre.",
            "Le site est probablement saturé. Réessayez dans quelques minutes.",
        )
    if isinstance(exc, requests.exceptions.HTTPError):
        return (
            "Erreur du serveur RCS",
            f"Le serveur a renvoyé une erreur : {exc.response.status_code if exc.response else 'inconnue'}.",
            "Réessayez dans quelques instants. Si l'erreur persiste, contactez le support technique du RCS.",
        )
    msg = str(exc)
    if "HTTP 403" in msg or "HTTP 429" in msg or "Trop de" in msg:
        return (
            "Trop de requêtes",
            "Le serveur RCS limite temporairement votre accès.",
            "Attendez 5 à 10 minutes avant de relancer une nouvelle extraction. Évitez de lancer plusieurs extractions à la suite.",
        )
    if "HTTP 5" in msg:
        return (
            "Erreur côté serveur RCS",
            "Le serveur RCS rencontre un problème interne.",
            "Réessayez dans quelques minutes. Ce n'est pas dû à votre côté.",
        )
    return (
        "Erreur inattendue",
        msg or "Une erreur inattendue s'est produite pendant l'extraction.",
        "Vérifiez les critères saisis. Si l'erreur persiste, signalez-la avec les détails techniques ci-dessous.",
    )


def render_error(exc: Exception):
    title, message, hint = classify_error(exc)
    st.markdown(
        f"""
<div class="error-card">
    <div class="error-header">
        <div class="error-icon">!</div>
        <h3 class="error-title">{title}</h3>
    </div>
    <p class="error-message">{message}</p>
    <div class="error-hint">{hint}</div>
</div>
""",
        unsafe_allow_html=True,
    )
    with st.expander("Détails techniques"):
        st.code(f"{type(exc).__name__}: {exc}\n\n{traceback.format_exc()}", language="text")


def render_stepper(active: int):
    """active: 1, 2 ou 3 (la phase en cours). Les précédentes sont 'done'."""
    steps = [
        ("Recherche", "Liste des entreprises"),
        ("Détails", "Informations par entreprise"),
        ("Export", "Génération Excel"),
    ]
    html = '<div class="stepper">'
    for i, (label, _) in enumerate(steps, 1):
        if i < active:
            cls = "step done"
        elif i == active:
            cls = "step active"
        else:
            cls = "step"
        html += f'<div class="{cls}"><span class="step-num">{i}</span><span>{label}</span></div>'
    html += "</div>"
    return html


# === Run scraping ===
if launch:
    st.session_state.result = None
    st.session_state.error = None
    st.session_state.running = True

    st.markdown('<div class="section-title">Progression en cours</div>', unsafe_allow_html=True)

    stepper_placeholder = st.empty()
    stepper_placeholder.markdown(render_stepper(1), unsafe_allow_html=True)

    progress_container = st.container()
    with progress_container:
        st.markdown('<div class="progress-card">', unsafe_allow_html=True)
        header_placeholder = st.empty()
        progress_bar = st.progress(0.0)
        meta_placeholder = st.empty()
        st.markdown("</div>", unsafe_allow_html=True)

    start = time.time()
    state = {"phase": 1, "total_known": False}

    def progress_cb(event, **kwargs):
        if event == "status":
            msg = kwargs.get("message", "")
            header_placeholder.markdown(
                f'<div class="progress-header">'
                f'<div><div class="progress-phase">{msg}</div></div>'
                f'<div class="progress-pct">—</div></div>',
                unsafe_allow_html=True,
            )
        elif event == "progress":
            current = kwargs.get("current", 0)
            total = kwargs.get("total", 0)
            msg = kwargs.get("message", "")
            eta = kwargs.get("eta", 0)

            if total > 0 and not state["total_known"]:
                state["total_known"] = True
                state["phase"] = 2
                stepper_placeholder.markdown(render_stepper(2), unsafe_allow_html=True)

            if total > 0:
                pct = min(max(current / total, 0.0), 1.0)
                progress_bar.progress(pct)
                detail = msg if msg else f"{current}/{total} entreprises"
                if len(detail) > 60:
                    detail = detail[:57] + "..."
                header_placeholder.markdown(
                    f'<div class="progress-header">'
                    f'<div>'
                    f'<div class="progress-phase">Récupération des détails</div>'
                    f'<div class="progress-detail">{detail}</div>'
                    f'</div>'
                    f'<div class="progress-pct">{int(pct * 100)}%</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
                eta_str = f"~{format_duration(eta)} restant" if eta > 0 else "Bientôt terminé"
                meta_placeholder.markdown(
                    f'<div style="display:flex; justify-content:space-between; '
                    f'font-size:0.85rem; color:#64748B; margin-top:0.6rem;">'
                    f'<span><strong style="color:#0F172A;">{current}</strong> sur <strong style="color:#0F172A;">{total}</strong></span>'
                    f'<span>{eta_str}</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

    try:
        session = make_session()
        rows = extract_entreprises(
            session, type_code, GREFFE_MAP[greffe], str(annee), forme,
            sleep_s=1.0, progress_cb=progress_cb, quiet=True,
        )

        stepper_placeholder.markdown(render_stepper(3), unsafe_allow_html=True)
        header_placeholder.markdown(
            '<div class="progress-header">'
            '<div><div class="progress-phase">Génération du fichier Excel</div></div>'
            '<div class="progress-pct">100%</div></div>',
            unsafe_allow_html=True,
        )
        progress_bar.progress(1.0)

        wb = build_workbook(rows, {
            "greffe": greffe, "annee": str(annee),
            "forme": forme, "type": type_code,
        })
        buffer = BytesIO()
        wb.save(buffer)
        excel_bytes = buffer.getvalue()
        file_size_kb = len(excel_bytes) / 1024

        elapsed = time.time() - start

        progress_container.empty()
        stepper_placeholder.empty()

        st.session_state.result = {
            "excel": excel_bytes,
            "filename": f"rcs_{greffe.replace(' ', '_')}_{annee}_{forme}.xlsx",
            "count": len(rows),
            "elapsed": elapsed,
            "size_kb": file_size_kb,
            "greffe": greffe,
            "annee": annee,
            "forme": forme,
            "type_label": type_label,
        }
        st.session_state.running = False
    except Exception as exc:
        progress_container.empty()
        stepper_placeholder.empty()
        st.session_state.error = exc
        st.session_state.running = False


# === Result display ===
if st.session_state.error is not None:
    st.markdown('<div class="section-title">Une erreur s\'est produite</div>', unsafe_allow_html=True)
    render_error(st.session_state.error)
    if st.button("Réessayer", type="primary", use_container_width=True):
        st.session_state.error = None
        st.rerun()


if st.session_state.result:
    r = st.session_state.result
    st.markdown('<div class="section-title">Extraction terminée</div>', unsafe_allow_html=True)
    st.markdown(
        f"""
<div class="result-card">
    <div class="result-header">
        <div class="result-check">✓</div>
        <h3 class="result-title">{r['count']} entreprise(s) extraite(s) avec succès</h3>
    </div>
    <div class="result-meta">
        <div class="meta-item">
            <div class="meta-label">Greffe</div>
            <div class="meta-value" style="font-size:1.05rem;">{r['greffe']}</div>
            <div class="meta-value-sub">{r['annee']} · {r['forme']}</div>
        </div>
        <div class="meta-item">
            <div class="meta-label">Durée</div>
            <div class="meta-value">{format_duration(r['elapsed'])}</div>
            <div class="meta-value-sub">Temps d'extraction</div>
        </div>
        <div class="meta-item">
            <div class="meta-label">Fichier</div>
            <div class="meta-value">{r['size_kb']:.0f} <span style="font-size:0.9rem; font-weight:500;">Ko</span></div>
            <div class="meta-value-sub">Format .xlsx</div>
        </div>
    </div>
</div>
""",
        unsafe_allow_html=True,
    )
    st.download_button(
        label="Télécharger le fichier Excel",
        data=r["excel"],
        file_name=r["filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption(f"Nom du fichier : `{r['filename']}`")


# === Footer ===
st.markdown(
    """
<div class="app-footer">
    © 2026 Tous droits réservés
    <span class="footer-divider">·</span>
    Idris Mathieo
</div>
""",
    unsafe_allow_html=True,
)
