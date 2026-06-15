# -*- coding: utf-8 -*-
"""
=============================================================================
APPLICATION STREAMLIT - COWORKINGS IDF
=============================================================================
Auteur      : [Votre nom]
Date        : 2025

Ce script charge le fichier Excel généré par le scraping PyQuery
(coworking_paris_complet.xlsx) et affiche les données sur une carte Folium.

Placer coworking_paris_complet.xlsx dans le même dossier que app.py.

Pour lancer :
  pip install streamlit folium streamlit-folium pandas openpyxl
  streamlit run app.py
=============================================================================
"""

# ============================================================
# IMPORTATION DES LIBRAIRIES
# ============================================================
import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import numpy as np
import os

# ============================================================
# CONFIGURATION DE LA PAGE
# ============================================================
st.set_page_config(
    page_title="Coworkings IDF",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CSS
# ============================================================
st.markdown("""
<style>
    .titre-app {
        text-align: center; font-size: 2.2rem; font-weight: 800;
        color: #2C3E50; padding: 0.6rem 0 0.2rem;
    }
    .sous-titre {
        text-align: center; color: #7f8c8d;
        font-size: 1rem; margin-bottom: 1.5rem;
    }
    .metric-box {
        background: white; border-radius: 10px;
        padding: 0.9rem 1.2rem;
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
        border-left: 4px solid #2980b9; margin-bottom: 0.8rem;
    }
    .metric-box h3 { font-size: 1.7rem; font-weight: 700; color: #2C3E50; margin: 0; }
    .metric-box p  { font-size: 0.8rem; color: #7f8c8d; margin: 0; }
</style>
""", unsafe_allow_html=True)


# ============================================================
# COORDONNÉES GPS PAR VILLE (pour la carte Folium)
# ============================================================
COORDS_VILLES = {
    "Paris":      (48.8566,  2.3522),
    "Boulogne":   (48.8352,  2.2409),
    "Nanterre":   (48.8924,  2.2070),
    "Vincennes":  (48.8479,  2.4395),
    "Montreuil":  (48.8641,  2.4439),
    "Créteil":    (48.7771,  2.4572),
    "Versailles": (48.8014,  2.1301),
    "Évry":       (48.6258,  2.4282),
    "Cergy":      (49.0359,  2.0799),
    "Puteaux":    (48.8835,  2.2395),
    "Levallois":  (48.8951,  2.2874),
    "Issy":       (48.8236,  2.2721),
    "Neuilly":    (48.8846,  2.2691),
    "Courbevoie": (48.8978,  2.2531),
    "Saint-Denis":(48.9362,  2.3574),
    "Massy":      (48.7257,  2.2558),
}


# ============================================================
# CHARGEMENT DES DONNÉES
# @st.cache_data : mis en cache — vu en cours module 4
# ============================================================
@st.cache_data
def charger_donnees():
    """
    Charge le fichier Excel issu du scraping PyQuery.
    Ajoute les colonnes LAT/LON si absentes (via le dictionnaire COORDS_VILLES).

    Returns:
        pd.DataFrame ou None si le fichier est introuvable
    """
    fichiers = [
        "coworking_enrichi.xlsx",        # fichier enrichi avec LAT/LON (priorité)
        "coworking_paris_complet.xlsx",
        "coworking_idf.xlsx",
        "coworkings_nettoyes.xlsx",
        "coworkings_nettoyes.csv",
    ]
    for f in fichiers:
        if not os.path.exists(f):
            continue
        try:
            df = pd.read_csv(f, encoding="utf-8") if f.endswith(".csv") \
                 else pd.read_excel(f, engine="openpyxl")

            # ── Nettoyage : convertir toutes les colonnes texte en str ──
            # Évite les erreurs PyArrow / Arrow lors de l'affichage
            for col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].astype(str).replace("nan", "")

            # ── Code Postal : "75002.0" → "75002" (Pandas lit en float depuis Excel)
            if "Code Postal" in df.columns:
                df["Code Postal"] = df["Code Postal"].apply(
                    lambda x: str(x).split(".")[0].strip()
                    if str(x) not in ["nan", "", "None"] else ""
                )

            # ── Ajout LAT/LON si absentes du fichier ──────────────
            if "LAT" not in df.columns or df["LAT"].eq("").all():
                df["LAT"] = df["Ville"].map(
                    lambda v: COORDS_VILLES.get(str(v).strip(),
                              COORDS_VILLES["Paris"])[0]
                )
                df["LON"] = df["Ville"].map(
                    lambda v: COORDS_VILLES.get(str(v).strip(),
                              COORDS_VILLES["Paris"])[1]
                )
            else:
                df["LAT"] = pd.to_numeric(df["LAT"], errors="coerce").fillna(48.8566)
                df["LON"] = pd.to_numeric(df["LON"], errors="coerce").fillna(2.3522)

            return df

        except Exception as e:
            st.sidebar.error(f"Erreur lecture {f} : {e}")
            continue

    return None


# ============================================================
# EN-TÊTE
# ============================================================
st.markdown('<div class="titre-app">🏢 Coworkings Île-de-France</div>',
            unsafe_allow_html=True)
st.markdown(
    '<div class="sous-titre">Données issues du scraping PyQuery · leportagesalarial.com</div>',
    unsafe_allow_html=True
)

# ── Chargement ───────────────────────────────────────────────
df = charger_donnees()

if df is None:
    st.error("""
    ❌ Aucun fichier de données trouvé.

    Placez votre fichier **coworking_paris_complet.xlsx** dans le même dossier que **app.py** :
    ```
    venv/
    ├── app.py
    └── coworking_paris_complet.xlsx   ← ici
    ```
    Puis rechargez la page.
    """)
    st.stop()   # on arrête l'exécution ici si pas de données

# ============================================================
# SIDEBAR — FILTRES
# Widgets vus en cours : text_input, multiselect, radio, selectbox
# ============================================================
with st.sidebar:
    st.markdown("## 🔍 Filtres")
    st.success(f"✅ {len(df)} espaces chargés")
    st.markdown("---")

    # ── Recherche par nom — st.text_input vu en cours ────────
    recherche = st.text_input(
        "🔎 Rechercher un espace",
        placeholder="ex : Kwerk, Morning..."
    )

    # ── Filtre ville — st.multiselect vu en cours ─────────────
    # Par défaut : uniquement Paris (comme demandé)
    villes_dispo = sorted(df["Ville"].replace("", "Non renseigné")
                           .dropna().unique().tolist())
    # On met Paris en premier dans les valeurs par défaut
    defaut_villes = villes_dispo
    villes_sel = st.multiselect(
        "🌆 Ville",
        options=villes_dispo,
        default=defaut_villes,   # ← Paris par défaut uniquement
        help="Sélectionnez une ou plusieurs villes"
    )

    # ── Filtre transport — st.radio vu en cours ───────────────
    filtre_transport = st.radio(
        "🚇 Transport",
        options=["Tous", "Métro", "RER", "Tram", "Bus"],
        horizontal=True
    )

    st.markdown("---")

    # ── Style de carte — st.selectbox ─────────────────────────
    style_carte = st.selectbox(
        "🗺️ Fond de carte",
        ["OpenStreetMap", "CartoDB positron", "CartoDB dark_matter"]
    )


# ============================================================
# FILTRAGE DES DONNÉES
# Masque booléen Pandas — vu en cours module 4
# ============================================================
def nettoyer(valeur):
    """Retourne une chaîne propre, ou '' si NaN/None."""
    return "" if (valeur is None or str(valeur).strip() in ["nan", "None", ""]) \
           else str(valeur).strip()

# Masque de base : toutes les lignes sont True au départ
masque = pd.Series([True] * len(df), index=df.index)

# Filtre ville
if villes_sel:
    masque &= df["Ville"].isin(villes_sel)

# Filtre recherche par nom
if recherche.strip():
    masque &= df["Nom"].str.contains(recherche.strip(), case=False, na=False)

# Filtre transport — cherche le mot dans la colonne Accès
if filtre_transport != "Tous":
    if "Accès" in df.columns:
        masque &= df["Accès"].str.contains(
            filtre_transport, case=False, na=False
        )

df_filtre = df[masque].copy()

# Construction colonne GEOCODE — même logique que le prof (code Doctolib)
# df["geocode"] = df[['LAT', 'LON']].values.tolist()
geocode = []
for lat, lon in zip(df_filtre["LAT"], df_filtre["LON"]):
    geocode.append([float(lat), float(lon)])
df_filtre["GEOCODE"] = pd.Series(geocode, index=df_filtre.index)


# ============================================================
# MÉTRIQUES GLOBALES — st.columns vu en cours
# ============================================================
c1, c2, c3, c4 = st.columns(4)

with c1:
    st.markdown(
        f'<div class="metric-box"><p>Espaces affichés</p><h3>{len(df_filtre)}</h3></div>',
        unsafe_allow_html=True)
with c2:
    nb_villes = df_filtre["Ville"].nunique()
    st.markdown(
        f'<div class="metric-box" style="border-left-color:#27ae60">'
        f'<p>Villes</p><h3>{nb_villes}</h3></div>',
        unsafe_allow_html=True)
with c3:
    if "Téléphone" in df_filtre.columns:
        nb_tel = int((df_filtre["Téléphone"] != "").sum())
    else:
        nb_tel = 0
    st.markdown(
        f'<div class="metric-box" style="border-left-color:#f39c12">'
        f'<p>Avec téléphone</p><h3>{nb_tel}</h3></div>',
        unsafe_allow_html=True)
with c4:
    if "Site Web" in df_filtre.columns:
        nb_site = int(df_filtre["Site Web"].str.startswith("http").sum())
    else:
        nb_site = 0
    st.markdown(
        f'<div class="metric-box" style="border-left-color:#9b59b6">'
        f'<p>Avec site web</p><h3>{nb_site}</h3></div>',
        unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ============================================================
# ONGLETS PRINCIPAUX
# ============================================================
onglet_carte, onglet_liste, onglet_stats, onglet_brut = st.tabs([
    "🗺️  Carte Folium",
    "📋  Liste",
    "📊  Statistiques",
    "🔬  Données brutes"
])


# ────────────────────────────────────────────────────────────
# ONGLET 1 — CARTE FOLIUM
# Structure identique au code Doctolib du prof :
#   for ..., geo in zip(..., df["GEOCODE"]):
#       if geo != None and type(geo) == list:
#           folium.Marker(geo, popup=p, tooltip=tooltip).add_to(m)
# ────────────────────────────────────────────────────────────
with onglet_carte:
    st.subheader("🗺️ Carte des espaces de coworking")

    if len(df_filtre) == 0:
        st.warning("⚠️ Aucun espace trouvé avec ces filtres. Essayez d'élargir la recherche.")
    else:
        # Centrage automatique sur les données filtrées
        centre_lat = df_filtre["LAT"].mean()
        centre_lon = df_filtre["LON"].mean()

        # Création de la carte — vu en cours module 6
        m = folium.Map(
            location=[centre_lat, centre_lon],
            zoom_start=12,
            tiles=style_carte
        )

        # ── Boucle for avec zip — même structure que le prof ──
        for nom, adresse, cp, ville, tel, site, url, acces, description, geo in zip(
            df_filtre["Nom"],
            df_filtre.get("Adresse",     pd.Series([""] * len(df_filtre), index=df_filtre.index)),
            df_filtre.get("Code Postal", pd.Series([""] * len(df_filtre), index=df_filtre.index)),
            df_filtre["Ville"],
            df_filtre.get("Téléphone",   pd.Series([""] * len(df_filtre), index=df_filtre.index)),
            df_filtre.get("Site Web",    pd.Series([""] * len(df_filtre), index=df_filtre.index)),
            df_filtre.get("URL",         pd.Series([""] * len(df_filtre), index=df_filtre.index)),
            df_filtre.get("Accès",       pd.Series([""] * len(df_filtre), index=df_filtre.index)),
            df_filtre.get("Description", pd.Series([""] * len(df_filtre), index=df_filtre.index)),
            df_filtre["GEOCODE"]
        ):
            # Nettoyage des valeurs
            nom_v   = nettoyer(nom)
            adr_v   = nettoyer(adresse)
            cp_v    = nettoyer(cp)
            ville_v = nettoyer(ville)
            tel_v   = nettoyer(tel)
            site_v  = nettoyer(site)
            url_v   = nettoyer(url)
            acces_v = nettoyer(acces)
            desc_v  = nettoyer(description)
            if len(desc_v) > 150:
                desc_v = desc_v[:150] + "..."

            # Vérification géocode — même condition que le prof
            if geo is not None and type(geo) == list:

              tooltip = f"{nom_v} | {adr_v} {cp_v} {ville_v} | 📞 {tel_v} | 🚇 {acces_v}"

                # ── POPUP : infos + liens cliquables au clic ──────
              lien_site   = f'<a href="{site_v}" target="_blank">🌐 Visiter le site</a>' \
                              if site_v.startswith("http") else ""
              lien_source = f'<a href="{url_v}" target="_blank" style="font-size:0.8em;color:#999;">📄 Fiche source</a>' \
                              if url_v.startswith("http") else ""

              p = f"""
                <div style="font-family:Arial,sans-serif;min-width:230px;max-width:290px;">
                    <h4 style="margin:0 0 5px;color:#2C3E50;">{nom_v}</h4>
                    <p style="margin:2px 0;font-size:0.82rem;color:#7f8c8d;">
                        📍 {adr_v}<br>{cp_v} {ville_v}
                    </p>
                    <hr style="margin:6px 0;border-color:#ecf0f1;">
                    {"<p style='margin:3px 0;'>📞 " + tel_v + "</p>" if tel_v else ""}
                    {"<p style='margin:3px 0;'>🚇 " + acces_v + "</p>" if acces_v else ""}
                    {"<p style='margin:4px 0;font-size:0.8rem;color:#555;font-style:italic;'>" + desc_v + "</p>" if desc_v else ""}
                    <hr style="margin:6px 0;border-color:#ecf0f1;">
                    {lien_site}<br>{lien_source}
                </div>
                """

                # Ajout du marqueur — exactement comme le prof
                folium.Marker(
                    geo,
                    popup=folium.Popup(p, max_width=300),
                    tooltip=folium.Tooltip(tooltip, sticky=True)
                ).add_to(m)

        # Affichage de la carte — st_folium (version moderne de folium_static)
        st_folium(m, height=520, returned_objects=[])
        st.caption(f"📍 {len(df_filtre)} espace(s) affiché(s) · Survolez un marqueur pour les détails")


# ────────────────────────────────────────────────────────────
# ONGLET 2 — LISTE
# ────────────────────────────────────────────────────────────
with onglet_liste:
    st.subheader("📋 Liste des espaces de coworking")

    if len(df_filtre) == 0:
        st.warning("Aucun résultat.")
    else:
        col_tri, col_ordre = st.columns([2, 1])
        with col_tri:
            critere = st.selectbox("Trier par", ["Nom", "Ville", "Code Postal"])
        with col_ordre:
            asc = st.radio("Ordre", ["↑ A→Z", "↓ Z→A"]) == "↑ A→Z"

        df_tri = df_filtre.sort_values(critere, ascending=asc)

        cols = st.columns(3)
        for i, (_, row) in enumerate(df_tri.iterrows()):
            with cols[i % 3]:
                tel   = nettoyer(row.get("Téléphone", "")) or "—"
                acces = nettoyer(row.get("Accès", ""))     or "—"
                site  = nettoyer(row.get("Site Web", ""))
                desc  = nettoyer(row.get("Description", ""))
                if len(desc) > 90: desc = desc[:90] + "..."
                lien  = f'<a href="{site}" target="_blank" style="font-size:0.77rem;color:#2980b9;">🌐 Site web</a>' \
                        if site.startswith("http") else ""

                st.markdown(f"""
                <div style="background:white;border-radius:10px;padding:0.9rem;
                            box-shadow:0 2px 6px rgba(0,0,0,0.07);
                            border-top:3px solid #2980b9;margin-bottom:0.9rem;">
                    <h4 style="margin:0 0 3px;font-size:0.92rem;color:#2C3E50;">
                        {nettoyer(row['Nom'])}
                    </h4>
                    <p style="margin:2px 0;font-size:0.77rem;color:#7f8c8d;">
                        📍 {nettoyer(row.get('Adresse',''))} · {nettoyer(row.get('Code Postal',''))} {nettoyer(row['Ville'])}
                    </p>
                    <p style="margin:3px 0;font-size:0.8rem;">📞 {tel}</p>
                    <p style="margin:2px 0;font-size:0.77rem;color:#555;">🚇 {acces}</p>
                    <p style="margin:4px 0;font-size:0.76rem;color:#666;font-style:italic;">{desc}</p>
                    {lien}
                </div>
                """, unsafe_allow_html=True)

        st.markdown("---")
        cols_show = [c for c in ["Nom","Adresse","Code Postal","Ville",
                                  "Téléphone","Accès","Site Web"]
                     if c in df_tri.columns]
        st.dataframe(df_tri[cols_show], hide_index=True)

        csv = df_tri.to_csv(index=False, encoding="utf-8-sig")
        st.download_button("⬇️ Télécharger (CSV)", data=csv,
                           file_name="coworkings_selection.csv", mime="text/csv")


# ────────────────────────────────────────────────────────────
# ONGLET 3 — STATISTIQUES — st.bar_chart vu en cours module 5
# ────────────────────────────────────────────────────────────
with onglet_stats:
    st.subheader("📊 Statistiques")

    if len(df_filtre) == 0:
        st.warning("Pas de données.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Espaces par ville**")
            st.bar_chart(df_filtre["Ville"].value_counts())

        with c2:
            st.markdown("**Espaces par département**")
            if "Code Postal" in df_filtre.columns:
                dept = df_filtre["Code Postal"].str[:2].value_counts()
                st.bar_chart(dept)

        c3, c4 = st.columns(2)
        with c3:
            st.markdown("**Avec site web vs sans**")
            if "Site Web" in df_filtre.columns:
                avec = int(df_filtre["Site Web"].str.startswith("http").sum())
                st.bar_chart(pd.Series({"Avec site": avec,
                                        "Sans site": len(df_filtre) - avec}))
        with c4:
            st.markdown("**Avec téléphone vs sans**")
            if "Téléphone" in df_filtre.columns:
                avec_t = int((df_filtre["Téléphone"] != "").sum())
                st.bar_chart(pd.Series({"Avec tél.": avec_t,
                                        "Sans tél.": len(df_filtre) - avec_t}))


# ────────────────────────────────────────────────────────────
# ONGLET 4 — DONNÉES BRUTES — st.dataframe vu en cours module 3
# ────────────────────────────────────────────────────────────
with onglet_brut:
    st.subheader("🔬 Données brutes")

    cols_afficher = [c for c in df_filtre.columns
                     if c not in ["GEOCODE", "LAT", "LON"]]
    st.dataframe(df_filtre[cols_afficher], hide_index=True)

    completude = (df_filtre[cols_afficher].replace("", np.nan)
                  .notna().mean() * 100).round(1).sort_values()
    st.markdown("**Complétude par colonne (%)**")
    st.bar_chart(completude)

    csv_brut = df_filtre.to_csv(index=False, encoding="utf-8-sig")
    st.download_button("⬇️ Exporter (CSV)", data=csv_brut,
                       file_name="coworkings_complet.csv", mime="text/csv")


# ============================================================
# PIED DE PAGE
# ============================================================
st.markdown("---")
st.markdown(
    "<p style='text-align:center;color:#bdc3c7;font-size:0.76rem;'>"
    "Application réalisée dans le cadre du cours Python · "
    "Scraping PyQuery sur leportagesalarial.com · "
    "Visualisation Folium + Streamlit"
    "</p>",
    unsafe_allow_html=True
)
