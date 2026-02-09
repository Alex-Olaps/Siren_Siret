import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from insee_sirene import get_sirets_from_siren
from export_excel import export_sirets_xlsx



# ----------------------------
# SIREN parsing
# ----------------------------
def normalize_siren_text(text: str) -> str:
    return (text or "").replace(" ", "").replace("\t", "").strip()


def extract_sirens_from_text(text: str) -> list[str]:
    """
    Extrait toutes les occurrences de 9 chiffres (SIREN) d'un texte.
    Accepte '481 986 446' => '481986446'
    """
    compact = normalize_siren_text(text)
    found = re.findall(r"\b\d{9}\b", compact)

    # Dédoublonnage en conservant l'ordre
    seen = set()
    out = []
    for s in found:
        if s not in seen:
            seen.add(s)
            out.append(s)
    return out


def extract_sirens_from_df(df: pd.DataFrame, column: str | None = None) -> list[str]:
    """
    Extrait des SIREN depuis un DataFrame.
    - Si column est fourni : on prend uniquement cette colonne
    - Sinon : on scanne toutes les colonnes (robuste)
    """
    if df is None or df.empty:
        return []

    if column and column in df.columns:
        series = df[column].astype(str)
        text = "\n".join(series.tolist())
        return extract_sirens_from_text(text)

    # Scan toutes colonnes (robuste)
    all_text = []
    for col in df.columns:
        all_text.extend(df[col].astype(str).tolist())
    return extract_sirens_from_text("\n".join(all_text))


# ----------------------------
# Upload parsing (CSV/XLSX)
# ----------------------------
def load_df_from_upload(uploaded_file) -> pd.DataFrame:
    suffix = Path(uploaded_file.name).suffix.lower()

    if suffix in [".xlsx", ".xls"]:
        # Excel
        return pd.read_excel(uploaded_file, engine="openpyxl")

    if suffix == ".csv":
        # CSV : on essaie d’être tolérant (séparateurs fréquents)
        # -> si tu sais que c'est toujours ';', tu peux mettre sep=';'
        try:
            return pd.read_csv(uploaded_file, sep=None, engine="python")
        except Exception:
            # fallback séparateur ; (fréquent en France)
            return pd.read_csv(uploaded_file, sep=";")

    raise ValueError("Format non supporté. Merci d'importer un .csv ou .xlsx/.xls")


# ----------------------------
# Streamlit App
# ----------------------------
st.set_page_config(page_title="SIREN → SIRET (INSEE)", layout="wide")
st.title("INSEE Sirene — SIREN → liste complète de SIRET (batch)")

INSEE_API_KEY = st.secrets.get("INSEE_API_KEY", "")

if "stop" not in st.session_state:
    st.session_state.stop = False

# On conserve les SIREN issus du fichier dans la session
if "sirens_from_file" not in st.session_state:
    st.session_state.sirens_from_file = []

with st.sidebar:
    st.header("Configuration")
    if not INSEE_API_KEY:
        st.warning("Ajoute INSEE_API_KEY dans .streamlit/secrets.toml")

    only_active = st.checkbox("Uniquement établissements actifs", value=True)
    page_size = st.number_input("Taille de page", min_value=20, max_value=200000, value=500, step=50)

    colA, colB = st.columns(2)
    with colA:
        if st.button("Stop"):
            st.session_state.stop = True
    with colB:
        if st.button("Reset Stop"):
            st.session_state.stop = False


st.markdown("## Importer un fichier (CSV / Excel)")
uploaded = st.file_uploader("Choisis un fichier .csv ou .xlsx", type=["csv", "xlsx", "xls"])

selected_col = None
df_upload = None

if uploaded is not None:
    try:
        df_upload = load_df_from_upload(uploaded)
        st.success(f"Fichier chargé : {uploaded.name} — {len(df_upload)} lignes, {len(df_upload.columns)} colonnes")
        st.dataframe(df_upload.head(50), use_container_width=True)

        # Si tu veux une colonne dédiée
        cols = ["(scanner toutes les colonnes)"] + list(df_upload.columns)
        choice = st.selectbox("Colonne contenant les SIREN (optionnel)", cols, index=0)
        selected_col = None if choice == "(scanner toutes les colonnes)" else choice

        sirens_file = extract_sirens_from_df(df_upload, selected_col)
        st.session_state.sirens_from_file = sirens_file

        st.info(f"SIREN détectés dans le fichier : {len(sirens_file)}")

    except Exception as e:
        st.session_state.sirens_from_file = []
        st.error(f"Impossible de lire le fichier : {e}")


st.markdown("---")
st.markdown("## SIREN à traiter")

sirens_text = st.text_area(
    "Tu peux coller ici des SIREN en plus (1 par ligne, ou séparés par espaces/virgules).",
    placeholder="481 986 446\n552100554\n...",
    height=140,
)

sirens_from_text = extract_sirens_from_text(sirens_text)

# Fusion fichier + texte
sirens_list = []
seen = set()
for s in (st.session_state.sirens_from_file + sirens_from_text):
    if s not in seen:
        seen.add(s)
        sirens_list.append(s)

st.caption(f"SIREN total à traiter : {len(sirens_list)}")
if sirens_list:
    with st.expander("Voir la liste des SIREN détectés"):
        st.code("\n".join(sirens_list))

btn_run = st.button(
    "Récupérer les SIRET",
    type="primary",
    disabled=(not sirens_list or not INSEE_API_KEY),
)

if btn_run:
    if st.session_state.stop:
        st.warning("Arrêt demandé. Clique sur 'Reset Stop' puis relance.")
        st.stop()

    status = st.empty()
    overall = st.progress(0)

    all_rows = []
    total = len(sirens_list)

    try:
        for i, s in enumerate(sirens_list, start=1):
            if st.session_state.stop:
                raise RuntimeError("STOP_REQUESTED")

            status.info(f"SIREN {i}/{total} : {s}")

            rows, _ = get_sirets_from_siren(
                siren=s,
                api_key=INSEE_API_KEY,
                only_active=only_active,
                page_size=int(page_size),
                should_stop=lambda: st.session_state.stop,
            )

            # Trace du SIREN "demandé"
            # for r in rows:
            #     r["SIREN demandé"] = s

            all_rows.extend(rows)
            overall.progress(i / total)

        status.success("Terminé ✅")

        df = pd.DataFrame(all_rows)

        if "SIRET" in df.columns:
            df = df.drop_duplicates(subset=["SIRET"])

        if "SIREN" in df.columns and "SIRET" in df.columns:
            df = df.sort_values(["SIREN", "SIRET"])

        st.success(f"Lignes totales : {len(df)}")
        st.dataframe(df, use_container_width=True)

        filename = "sirets.xlsx" if total == 1 else "sirets_batch.xlsx"
        xlsx_bytes = export_sirets_xlsx(df)
        st.download_button(
            "Télécharger XLSX (avec résumé)",
            data=xlsx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except RuntimeError as e:
        if str(e) == "STOP_REQUESTED":
            status.warning("Arrêt demandé : traitement interrompu.")
            st.stop()
        st.exception(e)
    except Exception as e:
        st.exception(e)
