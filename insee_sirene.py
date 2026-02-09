import time
from typing import Dict, List, Tuple, Optional, Callable

import requests
from datetime import date


BASE_URL = "https://api.insee.fr/api-sirene/3.11"

# Champs "aplatis" attendus par le paramètre `champs`
CHAMPS_ETAB = ",".join([
    # Identifiants
    "siret", "siren", "etablissementSiege",

    # Unité légale (nom)
    "denominationUniteLegale", "nomUniteLegale", "prenom1UniteLegale",

    # Établissement (enseigne / dénomination usuelle / état admin)
    "enseigne1Etablissement", "enseigne2Etablissement", "enseigne3Etablissement",
    "denominationUsuelleEtablissement",
    "etatAdministratifEtablissement",

    # Adresse établissement
    "complementAdresseEtablissement",
    "numeroVoieEtablissement", "indiceRepetitionEtablissement",
    "typeVoieEtablissement", "libelleVoieEtablissement",
    "distributionSpecialeEtablissement",
    "codePostalEtablissement",
    "libelleCommuneEtablissement",
    "libelleCedexEtablissement",
])


def _headers(api_key: str) -> Dict[str, str]:
    return {
        "Accept": "application/json",
        "X-INSEE-Api-Key-Integration": api_key,
        "User-Agent": "streamlit-siren-to-siret/1.0",
    }


def _normalize_siren(raw: str) -> str:
    siren = "".join(ch for ch in (raw or "") if ch.isdigit())
    if len(siren) != 9:
        raise ValueError("Le SIREN doit contenir exactement 9 chiffres.")
    return siren


def _latest_period(periodes: list) -> dict:
    """Prend la période la plus pertinente : dateFin vide (courante) sinon dateDebut max."""
    if not periodes:
        return {}
    current = [p for p in periodes if not p.get("dateFin")]
    if current:
        return current[0]
    return max(periodes, key=lambda p: p.get("dateDebut", "0000-00-00"))


def _get_unite_legale_name(etab: dict) -> str:
    ul = (etab or {}).get("uniteLegale") or {}

    # 1) format "nested"
    denom = (ul.get("denominationUniteLegale") or "").strip()
    if denom:
        return denom
    nom = (ul.get("nomUniteLegale") or "").strip()
    prenom = (ul.get("prenom1UniteLegale") or "").strip()
    full = " ".join([prenom, nom]).strip()
    if full:
        return full

    # 2) format "aplati"
    denom = (etab.get("denominationUniteLegale") or "").strip()
    if denom:
        return denom
    nom = (etab.get("nomUniteLegale") or "").strip()
    prenom = (etab.get("prenom1UniteLegale") or "").strip()
    return " ".join([prenom, nom]).strip()


def _get_etablissement_label(etab: dict) -> str:
    # 1) nested (périodes)
    periodes = (etab or {}).get("periodesEtablissement") or []
    p0 = _latest_period(periodes) if periodes else {}

    enseignes = [
        (p0.get("enseigne1Etablissement") or "").strip(),
        (p0.get("enseigne2Etablissement") or "").strip(),
        (p0.get("enseigne3Etablissement") or "").strip(),
    ]
    enseignes = [e for e in enseignes if e]
    if enseignes:
        return " / ".join(enseignes)

    denom_usuelle = (p0.get("denominationUsuelleEtablissement") or "").strip()
    if denom_usuelle:
        return denom_usuelle

    # 2) aplati
    enseignes = [
        (etab.get("enseigne1Etablissement") or "").strip(),
        (etab.get("enseigne2Etablissement") or "").strip(),
        (etab.get("enseigne3Etablissement") or "").strip(),
    ]
    enseignes = [e for e in enseignes if e]
    if enseignes:
        return " / ".join(enseignes)

    denom_usuelle = (etab.get("denominationUsuelleEtablissement") or "").strip()
    return denom_usuelle or ""


def _get_etat_admin(etab: dict) -> str:
    # 1) nested (périodes)
    periodes = (etab or {}).get("periodesEtablissement") or []
    if periodes:
        p0 = _latest_period(periodes)
        v = (p0.get("etatAdministratifEtablissement") or "").strip()
        if v:
            return v

    # 2) aplati
    return (etab.get("etatAdministratifEtablissement") or "").strip()


def _format_adresse(etab: dict) -> Dict[str, str]:
    adr = (etab or {}).get("adresseEtablissement") or {}

    def pick(k: str) -> str:
        return (adr.get(k) or etab.get(k) or "").strip()

    voie = " ".join(filter(None, [
        pick("numeroVoieEtablissement"),
        pick("indiceRepetitionEtablissement"),
        pick("typeVoieEtablissement"),
        pick("libelleVoieEtablissement"),
    ])).strip()

    complement = pick("complementAdresseEtablissement")
    dist = pick("distributionSpecialeEtablissement")
    cp = pick("codePostalEtablissement")
    commune = pick("libelleCommuneEtablissement")
    cedex = pick("libelleCedexEtablissement")

    parts = [p for p in [complement, voie, dist] if p]
    adresse = ", ".join(parts)
    ville = cedex or commune

    return {"Adresse": adresse, "Code postal": cp, "Ville": ville}


def get_sirets_from_siren(
    siren: str,
    api_key: str,
    only_active: bool = True,
    as_of_date: str | None = None,
    page_size: int = 500,
    max_pages: int = 500,
    max_429_retries: int = 15,
    base_sleep_s: float = 0.2,
    timeout_s: int = 30,
    should_stop: Optional[Callable[[], bool]] = None,
    on_page: Optional[Callable[[int, int], None]] = None,
) -> Tuple[List[dict], List[dict]]:
    
    siren = _normalize_siren(siren)

    if as_of_date is None:
        as_of_date = date.today().isoformat()  # "YYYY-MM-DD"


    q = f'siren:"{siren}"'
    if only_active:
        q += " AND periode(etatAdministratifEtablissement:A)"

    url = f"{BASE_URL}/siret"
    curseur = "*"

    all_etabs: List[dict] = []
    rows: List[dict] = []

    retry_429 = 0
    status_map = {"A": "Actif", "F": "Fermé"}

    # Optionnel mais utile : réutilise la connexion TCP
    session = requests.Session()

    for page in range(max_pages):
        if should_stop and should_stop():
            raise RuntimeError("STOP_REQUESTED")

        time.sleep(base_sleep_s)

        params = {
            "q": q,
            "date": as_of_date,
            "nombre": page_size,
            "curseur": curseur,
            "champs": CHAMPS_ETAB,
        }

        try:
            r = session.get(url, headers=_headers(api_key), params=params, timeout=timeout_s)
        except requests.RequestException as e:
            raise RuntimeError(f"Erreur réseau INSEE: {e}") from e

        if r.status_code == 429:
            retry_429 += 1
            if retry_429 > max_429_retries:
                raise RuntimeError("Trop de 429 (rate limit). Réessaie plus tard ou ralentis les appels.")
            time.sleep(1.0 * retry_429)
            continue

        retry_429 = 0

        if r.status_code == 401:
            raise RuntimeError(
                "401 Unauthorized : clé INSEE invalide/non autorisée. "
                "Vérifie la souscription à l'API Sirene (plan Public) et le header X-INSEE-Api-Key-Integration."
            )

        if r.status_code == 400:
            raise RuntimeError(f"400. URL={r.url}\nRéponse={r.text[:400]}")

        r.raise_for_status()
        data = r.json()

        etabs = data.get("etablissements", []) or []
        header = data.get("header", {}) or {}

        all_etabs.extend(etabs)

        for e in etabs:
            siret_val = e.get("siret", "")
            if not siret_val:
                continue

            etat_code = _get_etat_admin(e)
            if only_active and etat_code and etat_code != "A":
                continue
            adr = _format_adresse(e)

            nom_ul = _get_unite_legale_name(e)
            nom_etab = _get_etablissement_label(e) or ""

            rows.append({
                "SIRET": siret_val,
                "SIREN": e.get("siren", ""),
                "Nom unité légale": nom_ul,
                "Nom établissement": nom_etab,
                "Siège": bool(e.get("etablissementSiege")),
                "État administratif": status_map.get(etat_code, etat_code),
                **adr,
            })

        if on_page:
            on_page(page + 1, len(rows))

        next_cursor = header.get("curseurSuivant")

        if not next_cursor:
            break
        if next_cursor == curseur:
            break

        curseur = next_cursor

    else:
        raise RuntimeError(f"Arrêt sécurité : max_pages={max_pages} atteint (SIREN très volumineux ?).")

    # Dédoublonnage par SIRET
    dedup = {row["SIRET"]: row for row in rows}
    rows = list(dedup.values())
    rows.sort(key=lambda x: x.get("SIRET", ""))

    return rows, all_etabs
