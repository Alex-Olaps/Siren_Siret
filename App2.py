import re
from pathlib import Path

import pandas as pd
import streamlit as st

from insee_sirene import get_sirets_from_siren
from export_excel import export_sirets_xlsx


# =========================
# Page config + Style
# =========================
st.set_page_config(page_title="INSEE Sirene — SIREN → SIRET", layout="wide")

st.markdown(
    """
<style>
.block-container { padding-top: 2rem; padding-bottom: 2rem; max-width: 1200px; }
.app-title {
  font-size: 2rem;
  font-weight: 800;
  margin: 0;
  line-height: 1.25;
  padding-top: 0.50rem;
}

.app-subtitle { color: #6B7280; margin-top: .25rem; margin-bottom: 0.25rem; }
.badge {
  display: inline-block; padding: .15rem .6rem; border-radius: 999px;
  background: #EEF2FF; color: #3730A3; font-size: .8rem; font-weight: 700;
}
div[data-testid="stVerticalBlockBorderWrapper"]{
  border-radius: 14px;
  border: 1px solid rgba(0,0,0,0.06);
  background: rgba(255,255,255,0.7);
  backdrop-filter: blur(6px);
  padding: 1rem 1.25rem;
}
.small-muted { color: #6B7280; font-size: .9rem; }
div.stButton > button { border-radius: 12px; padding: .6rem 1rem; }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown('<p class="app-title">INSEE Sirene — SIREN → SIRET</p>', unsafe_allow_html=True)

st.write("")


# =========================
# Session state
# =========================
if "stop" not in st.session_state:
    st.session_state.stop = False

if "sirens_from_file" not in st.session_state:
    st.session_state.sirens_from_file = []

if "df_result" not in st.session_state:
    st.session_state.df_result = None


# =========================
# Helpers (SIREN parsing + upload)
# =========================
def extract_sirens_from_text(text: str) -> list[str]:
    """
    Extrait toutes les occurrences de 9 chiffres (SIREN) d'un texte.
    Accepte '481 986 446' => '481986446'
    """
    if not text:
        return []
    compact = text.replace(" ", "").replace("\t", "")
    found = re.findall(r"\b\d{9}\b", compact)

    seen = set()
    out = []
    for s in found:
        if s not in seen:
            seen.add(s)
            out.append(s)
    return out


def load_df_from_upload(uploaded_file) -> pd.DataFrame:
    suffix = Path(uploaded_file.name).suffix.lower()

    if suffix in [".xlsx", ".xls"]:
        return pd.read_excel(uploaded_file, engine="openpyxl")

    if suffix == ".csv":
        # lecture tolérante (auto-sep) + fallback ';' fréquent FR
        try:
            return pd.read_csv(uploaded_file, sep=None, engine="python")
        except Exception:
            return pd.read_csv(uploaded_file, sep=";")

    raise ValueError("Format non supporté. Merci d'importer un .csv ou .xlsx/.xls")


def extract_sirens_from_df(df: pd.DataFrame, column: str | None = None) -> list[str]:
    if df is None or df.empty:
        return []

    if column and column in df.columns:
        text = "\n".join(df[column].astype(str).tolist())
        return extract_sirens_from_text(text)

    # scan toutes colonnes
    all_text = []
    for col in df.columns:
        all_text.extend(df[col].astype(str).tolist())
    return extract_sirens_from_text("\n".join(all_text))


def merge_unique(a: list[str], b: list[str]) -> list[str]:
    seen = set()
    out = []
    for s in (a + b):
        if s not in seen:
            seen.add(s)
            out.append(s)
    return out


# =========================
# Secrets
# =========================
INSEE_API_KEY = st.secrets.get("INSEE_API_KEY", "")


# =========================
# Layout
# =========================
left, right = st.columns([1.45, 1], gap="large")

with left:

    with st.container(border=True):
        st.subheader("📥 Import & saisie")
    
        uploaded = st.file_uploader("Importer un fichier CSV / Excel", type=["csv", "xlsx", "xls"])

        df_upload = None
        selected_col = None

        if uploaded is not None:
            try:
                df_upload = load_df_from_upload(uploaded)
                st.success(f"Fichier chargé : {uploaded.name} — {len(df_upload)} lignes • {len(df_upload.columns)} colonnes")
                st.dataframe(df_upload.head(30), use_container_width=True)

                cols = ["(scanner toutes les colonnes)"] + list(df_upload.columns)
                choice = st.selectbox("Colonne contenant les SIREN (optionnel)", cols, index=0)
                selected_col = None if choice == "(scanner toutes les colonnes)" else choice

                sirens_file = extract_sirens_from_df(df_upload, selected_col)
                st.session_state.sirens_from_file = sirens_file
                st.info(f"SIREN détectés dans le fichier : {len(sirens_file)}")

            except Exception as e:
                st.session_state.sirens_from_file = []
                st.error(f"Impossible de lire le fichier : {e}")

        st.write("")
        st.caption("Tu peux aussi coller des SIREN directement (1 par ligne, ou séparés par espaces/virgules).")
        sirens_text = st.text_area(
            "SIREN",
            placeholder="481 986 446\n410408959\n410 409 460",
            height=140,
            label_visibility="collapsed",
        )

        sirens_from_text = extract_sirens_from_text(sirens_text)
        sirens_list = merge_unique(st.session_state.sirens_from_file, sirens_from_text)

        st.markdown(f'<p class="small-muted">SIREN total détectés : <b>{len(sirens_list)}</b></p>', unsafe_allow_html=True)

        if sirens_list:
            with st.expander("Voir la liste des SIREN détectés"):
                st.code("\n".join(sirens_list))

        st.markdown("</div>", unsafe_allow_html=True)

with right:

    with st.container(border=True):
        st.subheader("⚙️ Paramètres")

        if not INSEE_API_KEY:
            st.error("INSEE_API_KEY manquant. Ajoute-le dans Streamlit Cloud → Settings → Secrets.")
        only_active = st.checkbox("Uniquement établissements actifs", value=True)
        page_size = st.number_input("Taille de page", min_value=20, max_value=200000, value=500, step=50)

        st.divider()
        st.subheader("🚀 Actions")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⛔ Stop", use_container_width=True):
                st.session_state.stop = True
        with c2:
            if st.button("🔄 Reset", use_container_width=True):
                st.session_state.stop = False

        btn_run = st.button(
            "🔍 Récupérer les SIRET",
            type="primary",
            use_container_width=True,
            disabled=(not sirens_list or not INSEE_API_KEY),
        )

        st.markdown("</div>", unsafe_allow_html=True)


# =========================
# Run batch
# =========================
if btn_run:
    if st.session_state.stop:
        st.warning("Arrêt demandé. Clique sur 'Reset' puis relance.")
        st.stop()

    status_text = st.empty()
    overall = st.progress(0)
    status_text.write("Prêt")

    all_rows = []
    total = len(sirens_list)

    try:
        for i, s in enumerate(sirens_list, start=1):
            if st.session_state.stop:
                raise RuntimeError("STOP_REQUESTED")

            status_text.info(f"SIREN {i}/{total} : {s}")
            rows, _ = get_sirets_from_siren(
                siren=s,
                api_key=INSEE_API_KEY,
                only_active=only_active,
                page_size=int(page_size),
                should_stop=lambda: st.session_state.stop,
            )

            all_rows.extend(rows)
            overall.progress(i / total)

        df = pd.DataFrame(all_rows)

        # Dédoublonnage global par SIRET
        if "SIRET" in df.columns:
            df = df.drop_duplicates(subset=["SIRET"])

        # Tri lisible si possible
        if "SIREN" in df.columns and "SIRET" in df.columns:
            df = df.sort_values(["SIREN", "SIRET"])

        st.session_state.df_result = df
        status_text.success("Terminé ✅")
        overall.progress(1.0)

    except RuntimeError as e:
        if str(e) == "STOP_REQUESTED":
            status_text.info(label="Arrêt demandé : traitement interrompu.", state="error")
            st.stop()
        status_text.info(label="Erreur", state="error")
        st.exception(e)
    except Exception as e:
        status_text.info(label="Erreur", state="error")
        st.exception(e)


# =========================
# Display results (persist after rerun)
# =========================
df = st.session_state.df_result

if df is not None:
    st.write("")
    st.subheader("📊 Résultats")

    col1, col2, col3, col4 = st.columns(4)

    nb_siren = df["SIREN"].nunique() if "SIREN" in df.columns else "—"
    nb_siret = df["SIRET"].nunique() if "SIRET" in df.columns else len(df)
    nb_actifs = int((df["État administratif"] == "Actif").sum()) if "État administratif" in df.columns else "—"
    nb_sieges = int(df["Siège"].sum()) if "Siège" in df.columns else "—"

    col1.metric("SIREN", nb_siren)
    col2.metric("SIRET", nb_siret)
    col3.metric("Actifs", nb_actifs)
    col4.metric("Sièges", nb_sieges)

    tabs = st.tabs(["📄 Données", "⬇️ Export"])
    with tabs[0]:
        st.dataframe(df, use_container_width=True, height=560)

    with tabs[1]:
        filename = "sirets.xlsx" if (isinstance(nb_siren, int) and nb_siren == 1) else "sirets_batch.xlsx"
        xlsx_bytes = export_sirets_xlsx(df)
        st.download_button(
            "Télécharger XLSX (avec résumé)",
            data=xlsx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.caption("Le fichier contient 2 onglets : **SIRET** (table filtrable) et **Résumé** (global + par SIREN).")

