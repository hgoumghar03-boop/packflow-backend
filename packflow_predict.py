# -*- coding: utf-8 -*-
"""
======================================================================
 PackFlow Predict — Nature Growers (Agadir, Maroc)
 Application de pilotage de la station de conditionnement
 FICHIER UNIQUE : packflow_predict.py
======================================================================

UTILISATION DANS IDLE (Python 3.13) :
  1. Placez ce fichier dans  C:\\Users\\hanan\\Documents\\AI Logistique
     à côté de Nature_Growers_BDD.xlsx
  2. Choisissez le MODE juste en dessous, puis appuyez sur F5 :
        "installer"   → installe les bibliothèques (une seule fois)
        "valider"     → contrôle la base et affiche un résumé console
        "entrainer"   → entraîne et évalue l'IA (models/*.joblib + metriques.json)
        "application" → lance l'interface Streamlit dans le navigateur
        "api"         → lance l'API Flask en local (http://127.0.0.1:8000/api/sante)

DÉPLOIEMENT CLOUD (Render) :
  requirements.txt + commande de construction :
      pip install -r requirements.txt && python packflow_predict.py --entrainer
  et commande de démarrage :
      gunicorn packflow_predict:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 300
  Le port est fourni par l'hébergeur dans la variable d'environnement PORT (8000 en local).

ORGANISATION DU CODE (une classe = une responsabilité) :
  Config                → chemins, hypothèses, seuils, conversions
  ChargeurDonnees       → lecture du .xlsx, normalisation, compléments, export CSV
  ControleQualite       → anomalies, complétude, cohérence
  MoteurCapacite        → capacité des postes et des shifts, surcharge
  AnalyseDerive         → dérive de rendement et tests statistiques
  ConstructeurVariables → variables (features) pour le machine learning
  ModelesIA             → entraînement, évaluation, prédiction (predire_lot / predire_lots)
  ReferentielEmballages → mapping SKU + client → emballage, proposition, contrôle du stock
  GestionnairePrevisions→ espace prédiction : annonces futures, prédictions, traçabilité, statuts
  GestionnaireNotifications → alertes de la cloche 🔔 (critique, avertissement, attention, info)
  ContexteDonnees       → regroupe tous les résultats calculés
  ControlesRecette      → tests data (§21.1) et tests fonctionnels (§21.2) du cahier V2
  InterfaceUI           → CSS (thème clair), cartes KPI, badges, cloche, filtres
  PagesApplication      → les 10 écrans du cahier V2 (§15)
  Application           → orchestration Streamlit
  ValidationConsole     → bloc de validation (mode "valider")
  LanceurIDLE           → lance Streamlit depuis IDLE (mode "application")
  ServiceAPI / app      → API Flask (Cloud) : santé, référentiels, prédiction, annonces, capacité

La base Excel sert UNIQUEMENT à l'apprentissage et à l'alimentation de l'app :
aucune prédiction, alerte ni recommandation n'est réécrite dans l'historique.
Données 100 % synthétiques — démonstration et apprentissage.
"""

from __future__ import annotations   # doit rester la première instruction

# ======================================================================
#  PARAMÈTRE D'EXÉCUTION — À MODIFIER AVANT D'APPUYER SUR F5
# ======================================================================
MODE = "application"      # "installer" | "valider" | "entrainer" | "application" | "api"

VERSION = "3.3 — API Flask prête pour le Cloud (Render)"

# ======================================================================
#  IMPORTS
# ======================================================================

import importlib.util
import io
import json
import math
import os
import socket
import subprocess
import sys
import time
import unicodedata
import warnings
import webbrowser
from dataclasses import dataclass, field
from datetime import date, datetime, time as heure, timedelta
from pathlib import Path

warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning)

# Bibliothèques scientifiques : importées si disponibles (le mode "installer"
# doit pouvoir fonctionner même si elles sont absentes).
try:
    import numpy as np
    import pandas as pd
    BIBLIOS_DONNEES_OK = True
except ImportError:
    np = None
    pd = None
    BIBLIOS_DONNEES_OK = False

try:
    import streamlit as st
except ImportError:
    st = None

try:
    import plotly.express as px
    import plotly.graph_objects as go
except ImportError:
    px = None
    go = None


def _fichier_courant() -> Path:
    """Chemin de ce fichier (robuste dans IDLE comme dans Streamlit)."""
    try:
        return Path(__file__).resolve()
    except NameError:
        return Path.cwd() / "packflow_predict.py"


_BASE_PREVUE = Path(r"C:\Users\hanan\Documents\AI Logistique")


# ======================================================================
#  1. CONFIGURATION
# ======================================================================
class Config:
    """Toutes les constantes du projet (chemins, hypothèses chiffrées, seuils)."""

    # --- Chemins (pathlib : le dossier contient un espace) ----------------
    BASE_DIR = _BASE_PREVUE if _BASE_PREVUE.exists() else _fichier_courant().parent
    FICHIER_BDD = BASE_DIR / "Nature_Growers_BDD.xlsx"
    DATA_CLEAN_DIR = BASE_DIR / "data_clean"
    MODELS_DIR = BASE_DIR / "models"
    SORTIES_DIR = BASE_DIR / "sorties_moteur"        # prédictions / alertes (hors historique)
    FICHIER_ANOMALIES = DATA_CLEAN_DIR / "anomalies.csv"
    FICHIER_METRIQUES = MODELS_DIR / "metriques.json"
    FICHIER_SURCHARGES = BASE_DIR / "parametres_utilisateur.json"
    FICHIER_REQUIREMENTS = BASE_DIR / "requirements.txt"
    # Espace PRÉDICTION (séparé de l'historique d'apprentissage — cahier V2 §5)
    FICHIER_ANNONCES_FUTURES = SORTIES_DIR / "annonces_futures.csv"      # entrées futures
    FICHIER_PREDICTIONS = SORTIES_DIR / "predictions.csv"                # sorties du moteur
    FICHIER_HISTO_PREDICTIONS = SORTIES_DIR / "historique_predictions.csv"  # traçabilité
    FICHIER_NOTIFICATIONS = SORTIES_DIR / "notifications.csv"            # alertes de la cloche
    FICHIER_JOURNAL_ADMIN = SORTIES_DIR / "journal_admin.csv"            # actions d'administration
    FICHIER_RECEPTIONS = SORTIES_DIR / "receptions_emballages.csv"       # réapprovisionnements reçus
    HORIZON_PLAN_J = 60                                                  # horizon du plan de réapprovisionnement
    ALERTE_COMMANDE_J = 7                                                # alerte si commande à passer sous 7 jours
    FICHIER_CLASSEUR = DATA_CLEAN_DIR / "PackFlow_BDD_complete.xlsx"     # une feuille par table
    # Thème clair forcé : évite le texte blanc sur fond clair si le navigateur est en mode sombre
    THEME = {"base": "light", "primaryColor": "#2E7D32", "backgroundColor": "#FAF8F2",
             "secondaryBackgroundColor": "#FFFFFF", "textColor": "#1E2B22", "font": "sans serif"}
    ENCODAGE = "utf-8"

    # --- Tables --------------------------------------------------------
    FEUILLES_SOURCE = [
        "fermes", "produits_sku", "clients", "annonces_fermes", "planning_shifts",
        "stations", "stock_emballages", "parametres", "historique_production",
        "historique_temps_de_cycle", "hypotheses",
    ]
    TABLES_CAHIER = [
        "fermes", "produits_sku", "clients", "annonces_fermes", "planning_shifts",
        "stations", "machines_assemblage", "stock_emballages", "parametres",
        "historique_production", "historique_temps_de_cycle", "mapping_emballages",
    ]
    VOLUMES_ATTENDUS = {
        "fermes": 11, "produits_sku": 18, "clients": 4, "annonces_fermes": 2200,
        "planning_shifts": 3650, "stations": 5, "stock_emballages": 14600,
        "parametres": 10, "historique_production": 2000, "historique_temps_de_cycle": 2000,
    }

    # --- Période et shifts ---------------------------------------------
    DATE_DEBUT = date(2025, 10, 1)
    DATE_FIN = date(2026, 9, 30)
    SHIFTS = ["jour", "nuit"]
    DUREE_SHIFT_BRUTE_H = 8.0
    PAUSE_SHIFT_H = 0.5
    DUREE_SHIFT_UTILE_H = 7.5
    FACTEUR_NUIT = 0.95
    HEURE_DEBUT_SHIFT = {"jour": 6, "nuit": 22}

    # --- Stations / capacité ---------------------------------------------
    POSTES = ["réception", "tri", "soudage", "pesage", "palettisation"]
    POSTE_GOULOT = "soudage"
    CAPACITE_GOULOT_BDD_KG_H = 1300      # valeur de la table stations
    CAPACITE_GOULOT_CAHIER_KG_H = 1000   # valeur du cahier des charges / feuille Hypothèses
    # "cahier" : 1 000 kg/h → 7 500 kg/shift, cohérent avec la génération des données
    #            (≈ 13 % d'annonces en surcharge, cible 10–15 %)
    # "bdd"    : 1 300 kg/h (table stations) → ≈ 1,5 % de surcharge seulement
    SOURCE_CAPACITE_GOULOT = "cahier"
    NB_OPERATEURS_REFERENCE = {"réception": 5, "tri": 8, "soudage": 10, "pesage": 4, "palettisation": 6}
    # Contrainte emballages dans la capacité utile = min(poste limitant, emballages).
    # Désactivée par défaut : dans la BDD, les stocks sont des relevés indépendants des
    # annonces et feraient passer la surcharge à ≈ 22 % (cible 10–15 %). Activable dans Paramètres.
    CONTRAINTE_EMBALLAGE_CAPACITE = False
    SEUIL_TENSION = 0.85
    SEUIL_CHARGE_SHIFT = 1.0
    SEUIL_CHARGE_FORTE = 1.2
    CIBLE_TAUX_SURCHARGE = (0.10, 0.15)
    CIBLE_TAUX_ANOMALIES = 0.01

    # --- Cycle / transport ------------------------------------------------
    TEMPS_CHANGEMENT_LOT_MIN = 18
    VITESSE_CAMION_KM_H = 45
    ETAPES_CYCLE = [
        "heure_arrivee", "heure_debut_tri", "heure_fin_tri",
        "heure_debut_conditionnement", "heure_fin_conditionnement",
        "heure_fin_palettisation", "heure_expedition_reelle",
    ]
    LIBELLES_ETAPES = {
        "heure_arrivee": "Arrivée camion", "heure_debut_tri": "Début tri",
        "heure_fin_tri": "Fin tri", "heure_debut_conditionnement": "Début conditionnement",
        "heure_fin_conditionnement": "Fin conditionnement",
        "heure_fin_palettisation": "Fin palettisation", "heure_expedition_reelle": "Expédition",
    }
    # Durées par étape (colonne début, colonne fin)
    SEGMENTS_CYCLE = {
        "attente_reception_min": ("heure_arrivee", "heure_debut_tri"),
        "tri_min": ("heure_debut_tri", "heure_fin_tri"),
        "attente_conditionnement_min": ("heure_fin_tri", "heure_debut_conditionnement"),
        "conditionnement_min": ("heure_debut_conditionnement", "heure_fin_conditionnement"),
        "palettisation_min": ("heure_fin_conditionnement", "heure_fin_palettisation"),
        "expedition_min": ("heure_fin_palettisation", "heure_expedition_reelle"),
    }
    LIBELLES_SEGMENTS = {
        "attente_reception_min": "Attente avant tri", "tri_min": "Tri",
        "attente_conditionnement_min": "Attente avant conditionnement",
        "conditionnement_min": "Conditionnement", "palettisation_min": "Palettisation",
        "expedition_min": "Préparation expédition",
    }

    # --- Référentiels ----------------------------------------------------
    FAMILLES = ["Haricots verts", "Piments", "Mini poivrons"]
    DELAI_LIVRAISON_PAR_PAYS_H = {"Espagne": 24, "France": 36, "Allemagne": 48, "Royaume-Uni": 72}
    MAPPING_STATUTS = {"REALISEE": "REALISEE", "ANNULEE": "ANNULEE", "REPORTEE": "PLANIFIEE",
                       "PLANIFIEE": "PLANIFIEE", "EN_ROUTE": "EN_ROUTE"}
    ORIGINES_ANNONCE = ["ERP", "PLANIFICATEUR", "FERME", "IMPORT"]
    PROBA_ORIGINES = [0.45, 0.30, 0.15, 0.10]
    # Une annonce annulée ou reportée ne charge plus son shift d'origine
    STATUTS_HORS_CHARGE = ["ANNULEE", "REPORTEE"]
    PRIORITES = ["NORMALE", "HAUTE", "URGENTE"]
    ORIGINES_SAISIE = ["MANUEL", "ERP", "PLANIFICATEUR", "FERME", "IMPORT"]
    STATUTS_PREDICTION = ["CREEE", "RECALCULEE", "VALIDEE", "ARCHIVEE"]
    OPERATEURS_DEFAUT = {"jour": 10, "nuit": 9}       # médiane observée au soudage
    SEUIL_MARGE_EXPEDITION_MIN = 30                    # marge minimale avant l'heure limite client

    # --- Emballages --------------------------------------------------------
    CATEGORIE_ARTICLE = {
        "CARTON_EXPORT_2KG": "carton", "CARTON_EXPORT_5KG": "carton", "EMBALLAGE_SUR_MESURE": "carton",
        "ETIQUETTE_MERCADONA": "etiquette", "ETIQUETTE_LIDL": "etiquette",
        "ETIQUETTE_TESCO": "etiquette", "ETIQUETTE_CARREFOUR": "etiquette",
        "SACHET_IMPORTE_100G": "sachet_importe", "SACHET_IMPORTE_250G": "sachet_importe",
        "SACHET_IMPORTE_500G": "sachet_importe",
    }
    SEUILS_COUVERTURE_J = {"carton": 7, "etiquette": 21, "sachet_importe": 14}
    # Conversion stock → kg supportés (hypothèse documentée)
    ETIQUETTES_PAR_ROULEAU = 500          # 1 rouleau = 500 étiquettes
    KG_PAR_COLIS_ETIQUETE = 2.5           # 1 étiquette par colis de 2,5 kg en moyenne
    KG_PAR_UNITE = {
        "CARTON_EXPORT_2KG": 2.0, "CARTON_EXPORT_5KG": 5.0, "EMBALLAGE_SUR_MESURE": 3.0,
        "SACHET_IMPORTE_100G": 0.10, "SACHET_IMPORTE_250G": 0.25, "SACHET_IMPORTE_500G": 0.50,
    }
    FORMAT_VERS_ARTICLE = {
        "Vrac 2 kg": "CARTON_EXPORT_2KG", "Vrac 5 kg": "CARTON_EXPORT_5KG",
        "Sachet 50 g": "SACHET_IMPORTE_100G", "Sachet 100 g": "SACHET_IMPORTE_100G",
        "Sachet 200 g": "SACHET_IMPORTE_250G", "Sachet 250 g": "SACHET_IMPORTE_250G",
        "Sachet 500 g": "SACHET_IMPORTE_500G",
    }
    ETIQUETTE_PAR_CLIENT = {"CLI001": "ETIQUETTE_MERCADONA", "CLI002": "ETIQUETTE_LIDL",
                            "CLI003": "ETIQUETTE_TESCO", "CLI004": "ETIQUETTE_CARREFOUR"}

    # --- Dérive de rendement ---------------------------------------------
    FENETRE_DERIVE_J = 14
    DERIVE_BAISSE = (date(2026, 1, 5), date(2026, 2, 15))
    DERIVE_REDRESSEMENT = (date(2026, 2, 16), date(2026, 3, 31))
    SEUIL_P_VALUE = 0.05

    # --- Qualité des données ---------------------------------------------
    TAUX_DECHET_MAX = 0.15
    TOLERANCE_POIDS = 0.01
    TOLERANCE_RENDEMENT = 0.05     # rendement vs poids_emballe ÷ heures_travaillees
    SEUIL_Z_ROBUSTE = 3.5

    # --- Machine learning ------------------------------------------------
    RANDOM_STATE = 42
    ML_FIN_TRAIN = date(2026, 6, 30)       # train : 2025-10 → 2026-06
    ML_DEBUT_TEST = date(2026, 7, 1)       # test  : 2026-07 → 2026-09
    ML_N_SPLITS = 4

    # --- Interface -----------------------------------------------------------
    TITRE_APP = "PackFlow Predict"
    PIED_DE_PAGE = "Données synthétiques — démonstration et apprentissage"
    COULEURS = {
        "fond": "#FAF8F2", "carte": "#FFFFFF", "vert": "#2E7D32", "vert_clair": "#3FA34D",
        "panneau": "#0F2A1D", "texte": "#1E2B22", "texte2": "#6B7A70",
        "normal": "#2E7D32", "tension": "#E0A100", "critique": "#C62828", "neutre": "#8A8F8C",
    }
    PALETTE_GRAPHIQUES = ["#2E7D32", "#3FA34D", "#8BC34A", "#C5A572", "#5D8A73", "#A5D6A7"]
    PORT_STREAMLIT = 8501
    MOT_DE_PASSE_ADMIN = "packflow2026"     # remplaçable par st.secrets["admin_password"]

    # Paramètres modifiables depuis l'écran « Paramètres »
    PARAMETRES_MODIFIABLES = {
        "SOURCE_CAPACITE_GOULOT": "Source capacité goulot (cahier / bdd)",
        "CAPACITE_GOULOT_BDD_KG_H": "Capacité goulot BDD (kg/h)",
        "CAPACITE_GOULOT_CAHIER_KG_H": "Capacité goulot cahier (kg/h)",
        "CONTRAINTE_EMBALLAGE_CAPACITE": "Limiter la capacité utile par les emballages",
        "FACTEUR_NUIT": "Facteur de performance nuit",
        "SEUIL_TENSION": "Seuil de tension (ratio)",
        "SEUIL_CHARGE_SHIFT": "Seuil de surcharge (ratio)",
        "SEUIL_CHARGE_FORTE": "Seuil de forte surcharge (ratio)",
        "TEMPS_CHANGEMENT_LOT_MIN": "Temps de changement de lot (min)",
        "VITESSE_CAMION_KM_H": "Vitesse moyenne camion (km/h)",
        "SEUIL_MARGE_EXPEDITION_MIN": "Marge minimale avant l'heure limite client (min)",
        "ETIQUETTES_PAR_ROULEAU": "Étiquettes par rouleau",
        "KG_PAR_COLIS_ETIQUETE": "kg par colis étiqueté",
    }

    @classmethod
    def capacite_goulot_kg_h(cls) -> float:
        """Capacité nominale du goulot selon la source choisie."""
        if cls.SOURCE_CAPACITE_GOULOT == "bdd":
            return float(cls.CAPACITE_GOULOT_BDD_KG_H)
        return float(cls.CAPACITE_GOULOT_CAHIER_KG_H)

    @classmethod
    def facteur_shift(cls, shift: str) -> float:
        return float(cls.FACTEUR_NUIT) if shift == "nuit" else 1.0

    @classmethod
    def kg_par_unite(cls, article: str) -> float:
        if str(article).startswith("ETIQUETTE"):
            return float(cls.ETIQUETTES_PAR_ROULEAU) * float(cls.KG_PAR_COLIS_ETIQUETE)
        return float(cls.KG_PAR_UNITE.get(article, 0.0))

    @classmethod
    def kg_supportes(cls, article: str, quantite: float) -> float:
        """Convertit un stock d'emballage (cartons, sachets, rouleaux) en kg supportés."""
        try:
            q = float(quantite)
        except (TypeError, ValueError):
            return 0.0
        if math.isnan(q) or q <= 0:
            return 0.0
        return q * cls.kg_par_unite(article)

    @classmethod
    def seuil_couverture(cls, article: str) -> int:
        return cls.SEUILS_COUVERTURE_J.get(cls.CATEGORIE_ARTICLE.get(article, "carton"), 7)

    @classmethod
    def version_parametres(cls) -> str:
        """Empreinte courte des paramètres modifiables (traçabilité des prédictions)."""
        import hashlib
        valeurs = {k: getattr(cls, k) for k in sorted(cls.PARAMETRES_MODIFIABLES)}
        return "P-" + hashlib.sha1(json.dumps(valeurs, sort_keys=True, default=str).encode()).hexdigest()[:8]

    @classmethod
    def journaliser_admin(cls, action: str, detail: str = "") -> None:
        """Journal des actions d'administration (exigence de sécurité §20)."""
        try:
            cls.creer_dossiers()
            pd.DataFrame([{"horodatage": datetime.now().isoformat(timespec="seconds"), "action": action,
                           "detail": detail}]).to_csv(cls.FICHIER_JOURNAL_ADMIN, mode="a", index=False,
                                                     header=not cls.FICHIER_JOURNAL_ADMIN.exists(),
                                                     encoding=cls.ENCODAGE)
        except OSError:
            pass

    @classmethod
    def ecrire_theme(cls) -> None:
        """Écrit .streamlit/config.toml (thème clair) dans le dossier du projet."""
        dossier = cls.BASE_DIR / ".streamlit"
        dossier.mkdir(parents=True, exist_ok=True)
        lignes = ["[theme]"] + [f'{cle} = "{valeur}"' for cle, valeur in cls.THEME.items()]
        lignes += ["", "[browser]", "gatherUsageStats = false", ""]
        (dossier / "config.toml").write_text("\n".join(lignes), encoding=cls.ENCODAGE)

    @classmethod
    def creer_dossiers(cls) -> None:
        for dossier in (cls.DATA_CLEAN_DIR, cls.MODELS_DIR, cls.SORTIES_DIR):
            dossier.mkdir(parents=True, exist_ok=True)

    @classmethod
    def charger_surcharges(cls) -> dict:
        """Applique parametres_utilisateur.json (liste blanche uniquement)."""
        if not cls.FICHIER_SURCHARGES.exists():
            return {}
        try:
            donnees = json.loads(cls.FICHIER_SURCHARGES.read_text(encoding=cls.ENCODAGE))
        except (OSError, json.JSONDecodeError):
            return {}
        appliquees = {}
        for cle, valeur in donnees.items():
            if cle in cls.PARAMETRES_MODIFIABLES:
                setattr(cls, cle, valeur)
                appliquees[cle] = valeur
        return appliquees

    @classmethod
    def enregistrer_surcharges(cls, valeurs: dict) -> None:
        filtre = {k: v for k, v in valeurs.items() if k in cls.PARAMETRES_MODIFIABLES}
        cls.journaliser_admin("modification_parametres", json.dumps(filtre, ensure_ascii=False, default=str))
        cls.FICHIER_SURCHARGES.write_text(json.dumps(filtre, ensure_ascii=False, indent=2),
                                          encoding=cls.ENCODAGE)
        cls.charger_surcharges()

    @classmethod
    def reinitialiser_surcharges(cls) -> None:
        cls.journaliser_admin("reinitialisation_parametres")
        if cls.FICHIER_SURCHARGES.exists():
            cls.FICHIER_SURCHARGES.unlink()

    @classmethod
    def tableau_parametres(cls) -> "pd.DataFrame":
        lignes = [
            ("Période de la base d'apprentissage", f"{cls.DATE_DEBUT} → {cls.DATE_FIN}", "dates"),
            ("Shift utile", cls.DUREE_SHIFT_UTILE_H, "h"),
            ("Facteur nuit", cls.FACTEUR_NUIT, "ratio"),
            ("Poste goulot", cls.POSTE_GOULOT, "-"),
            ("Capacité goulot utilisée", cls.capacite_goulot_kg_h(), "kg/h"),
            ("Source capacité goulot", cls.SOURCE_CAPACITE_GOULOT, "-"),
            ("Capacité goulot / shift", cls.capacite_goulot_kg_h() * cls.DUREE_SHIFT_UTILE_H, "kg"),
            ("Contrainte emballages dans la capacité", "oui" if cls.CONTRAINTE_EMBALLAGE_CAPACITE else "non", "-"),
            ("Seuil de tension", cls.SEUIL_TENSION, "ratio"),
            ("Seuil de surcharge", cls.SEUIL_CHARGE_SHIFT, "ratio"),
            ("Seuil de forte surcharge", cls.SEUIL_CHARGE_FORTE, "ratio"),
            ("Changement de lot", cls.TEMPS_CHANGEMENT_LOT_MIN, "min"),
            ("Marge minimale d'expédition", cls.SEUIL_MARGE_EXPEDITION_MIN, "min"),
            ("Vitesse camion", cls.VITESSE_CAMION_KM_H, "km/h"),
            ("Couverture cartons", cls.SEUILS_COUVERTURE_J["carton"], "jours"),
            ("Couverture étiquettes", cls.SEUILS_COUVERTURE_J["etiquette"], "jours"),
            ("Couverture sachets importés", cls.SEUILS_COUVERTURE_J["sachet_importe"], "jours"),
            ("kg supportés par rouleau d'étiquettes", cls.kg_par_unite("ETIQUETTE_LIDL"), "kg"),
            ("Entraînement IA", f"{cls.DATE_DEBUT} → {cls.ML_FIN_TRAIN}", "dates"),
            ("Test IA", f"{cls.ML_DEBUT_TEST} → {cls.DATE_FIN}", "dates"),
        ]
        return pd.DataFrame(lignes, columns=["paramètre", "valeur", "unité"]).astype({"valeur": str})


Config.charger_surcharges()


# ======================================================================
#  2. CHARGEMENT ET NORMALISATION DES DONNÉES
# ======================================================================
class ErreurDonnees(Exception):
    """Erreur de données avec un message clair en français."""


class ChargeurDonnees:
    """Lit le .xlsx, normalise les types, complète ce qui manque, exporte les CSV."""

    COLONNES_DATE = {
        "annonces_fermes": ["date_annonce", "date_recolte_prevue"],
        "planning_shifts": ["date"],
        "stock_emballages": ["date_releve"],
        "historique_production": ["date"],
    }
    COLONNES_DATETIME = {
        "annonces_fermes": ["depart_camion_prevu", "arrivee_camion_prevue"],
        "historique_temps_de_cycle": Config.ETAPES_CYCLE,
    }
    # (table, colonne, table référencée, colonne référencée)
    RELATIONS = [
        ("annonces_fermes", "ferme_id", "fermes", "ferme_id"),
        ("annonces_fermes", "sku_id", "produits_sku", "sku_id"),
        ("annonces_fermes", "client_id", "clients", "client_id"),
        ("historique_production", "ferme_id", "fermes", "ferme_id"),
        ("historique_production", "sku_id", "produits_sku", "sku_id"),
        ("historique_production", "client_id", "clients", "client_id"),
        ("historique_production", "lot_id", "annonces_fermes", "lot_id"),
        ("historique_temps_de_cycle", "lot_id", "historique_production", "lot_id"),
        ("historique_production", "lot_id", "historique_temps_de_cycle", "lot_id"),
        ("stock_emballages", "client_id", "clients", "client_id"),
        ("planning_shifts", "poste", "stations", "poste"),
        ("mapping_emballages", "sku_id", "produits_sku", "sku_id"),
        ("mapping_emballages", "client_id", "clients", "client_id"),
        ("mapping_emballages", "article_emballage", "stock_emballages", "article"),
    ]
    CLES_PRIMAIRES = {
        "fermes": "ferme_id", "produits_sku": "sku_id", "clients": "client_id",
        "annonces_fermes": "id", "stations": "station_id", "machines_assemblage": "machine_id",
        "historique_production": "lot_id", "historique_temps_de_cycle": "lot_id",
        "mapping_emballages": "mapping_id",
    }

    def __init__(self, fichier: Path | None = None):
        self.fichier = Path(fichier) if fichier else Config.FICHIER_BDD
        self.ecarts: list[dict] = []          # écarts entre le fichier et le cahier des charges
        self.rapport_cles: pd.DataFrame | None = None

    # ------------------------------------------------------------------
    @staticmethod
    def _cle(nom: str) -> str:
        """Nom de feuille normalisé : minuscules, sans accents ni espaces."""
        texte = unicodedata.normalize("NFKD", str(nom))
        texte = "".join(c for c in texte if not unicodedata.combining(c))
        return texte.strip().lower().replace(" ", "_")

    def charger(self, exporter: bool = True, classeur: bool = False) -> dict:
        """Pipeline complet : lecture → normalisation → compléments → contrôles → CSV (+ classeur XLSX)."""
        tables = self.lire_excel()
        self.normaliser(tables)
        self.completer(tables)
        self.rapport_cles = self.verifier_cles(tables)
        if exporter:
            self.exporter_csv(tables)
        if classeur:
            self.exporter_classeur(tables)
        return tables

    def exporter_classeur(self, tables: dict) -> Path | None:
        """Livrable 04 : une feuille par table + hypothèses."""
        Config.creer_dossiers()
        try:
            with pd.ExcelWriter(Config.FICHIER_CLASSEUR, engine="openpyxl") as ecrivain:
                for nom in Config.TABLES_CAHIER + ["hypotheses"]:
                    df = tables.get(nom)
                    if df is not None and len(df.columns):
                        self.format_iso(df).to_excel(ecrivain, sheet_name=nom[:31], index=False)
            return Config.FICHIER_CLASSEUR
        except PermissionError:
            print(f"[Avertissement] {Config.FICHIER_CLASSEUR} est ouvert dans Excel : export ignoré.")
            return None

    def lire_excel(self) -> dict:
        if not self.fichier.exists():
            raise ErreurDonnees(
                "Base de données introuvable.\n"
                f"Chemin attendu : {self.fichier}\n"
                "Placez le fichier Nature_Growers_BDD.xlsx dans ce dossier puis relancez."
            )
        try:
            brut = pd.read_excel(self.fichier, sheet_name=None, engine="openpyxl")
        except PermissionError:
            raise ErreurDonnees(f"Le fichier est ouvert dans Excel ou verrouillé : {self.fichier}. "
                                "Fermez-le puis relancez.")
        except Exception as erreur:
            raise ErreurDonnees(f"Lecture impossible du fichier Excel {self.fichier} : {erreur}")

        index = {self._cle(nom): df for nom, df in brut.items()}
        tables, manquantes = {}, []
        for nom in Config.FEUILLES_SOURCE:
            if self._cle(nom) in index:
                tables[nom] = index[self._cle(nom)].copy()
            else:
                manquantes.append(nom)
        bloquantes = [n for n in manquantes if n not in ("hypotheses", "parametres")]
        if bloquantes:
            raise ErreurDonnees("Feuilles manquantes dans le classeur : " + ", ".join(bloquantes))
        for n in manquantes:
            tables[n] = pd.DataFrame()
        return tables

    def normaliser(self, tables: dict) -> None:
        for nom, df in tables.items():
            df.columns = [str(c).strip() for c in df.columns]
            for col in df.columns:
                if df[col].dtype == object or pd.api.types.is_string_dtype(df[col]):
                    df[col] = df[col].map(lambda v: v.strip() if isinstance(v, str) else v)
            for col in self.COLONNES_DATE.get(nom, []):
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors="coerce").dt.normalize()
            for col in self.COLONNES_DATETIME.get(nom, []):
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors="coerce")
        clients = tables["clients"]
        if "heure_limite_depart" in clients.columns:
            clients["heure_limite_depart"] = clients["heure_limite_depart"].map(self._format_heure)

    @staticmethod
    def _format_heure(valeur) -> str:
        if isinstance(valeur, (heure, datetime)):
            return valeur.strftime("%H:%M")
        texte = str(valeur)
        return texte[:5] if len(texte) >= 5 else texte

    def completer(self, tables: dict) -> None:
        """Ajoute tout ce que le cahier des charges prévoit mais que le fichier n'a pas."""
        rng = np.random.default_rng(Config.RANDOM_STATE)
        fermes, clients, annonces = tables["fermes"], tables["clients"], tables["annonces_fermes"]

        # 1. Feuille machines_assemblage absente → générée (seed fixe)
        if "machines_assemblage" not in tables or tables["machines_assemblage"].empty:
            tables["machines_assemblage"] = pd.DataFrame({
                "machine_id": ["MA01", "MA02", "MA03"],
                "type": ["fixe", "fixe", "reglable"],
                "format_gere": ["Carton 2–5 kg", "Sachet 50–500 g", "Carton 2–5 kg / Sachet 50–500 g"],
                "disponibilite_pct": np.round(rng.uniform(85, 98, 3), 1),
            })
            self._ecart("machines_assemblage", "Feuille absente du fichier",
                        "Créée : 2 machines fixes + 1 réglable (seed 42)")

        # 1 bis. Table mapping_emballages (cahier V2 §8 et §16.1)
        if "mapping_emballages" not in tables or tables["mapping_emballages"].empty:
            tables["mapping_emballages"] = ReferentielEmballages.generer_mapping(tables["produits_sku"], clients)
            self._ecart("mapping_emballages", "Table absente (mapping SKU → emballage non explicite)",
                        "Créée : défaut par format + règles client (sachets personnalisés Mercadona/Tesco, étiquettes)")

        # 2. Colonnes absentes
        if "capacite_hebdo_estimee_kg" not in fermes.columns:
            nb_semaines = (Config.DATE_FIN - Config.DATE_DEBUT).days / 7
            hebdo = annonces.groupby("ferme_id")["volume_annonce_kg"].sum() / nb_semaines
            observe = fermes["ferme_id"].map(hebdo)
            selon_distance = (20000 - 150 * fermes["distance_usine_km"].fillna(30)).clip(lower=2000)
            capacite = (observe * 1.3).fillna(selon_distance)
            fermes["capacite_hebdo_estimee_kg"] = (np.ceil(capacite / 100) * 100).astype(int)
            self._ecart("fermes.capacite_hebdo_estimee_kg", "Colonne absente",
                        "Estimée : volume hebdo observé × 1,3 (repli selon la distance)")
        if "delai_livraison_promis_h" not in clients.columns:
            clients["delai_livraison_promis_h"] = (
                clients["pays"].map(Config.DELAI_LIVRAISON_PAR_PAYS_H).fillna(48).astype(int))
            self._ecart("clients.delai_livraison_promis_h", "Colonne absente",
                        "Ajoutée : Espagne 24 h, France 36 h, Allemagne 48 h, Royaume-Uni 72 h")
        if "origine_annonce" not in annonces.columns:
            annonces["origine_annonce"] = rng.choice(Config.ORIGINES_ANNONCE, size=len(annonces),
                                                     p=Config.PROBA_ORIGINES)
            self._ecart("annonces_fermes.origine_annonce", "Colonne absente",
                        "Ajoutée : ERP / PLANIFICATEUR / FERME / IMPORT (seed 42)")

        # 3. Statuts : REPORTEE → PLANIFIEE dans une colonne dérivée
        annonces["statut_normalise"] = (annonces["statut_annonce"].map(Config.MAPPING_STATUTS)
                                        .fillna(annonces["statut_annonce"]))
        nb_rep = int((annonces["statut_annonce"] == "REPORTEE").sum())
        if nb_rep:
            self._ecart("annonces_fermes.statut_annonce",
                        f"{nb_rep} annonces au statut REPORTEE (absent du cahier)",
                        "Valeur conservée ; statut_normalise = PLANIFIEE")

        # 4. Référence d'annonce AN-AAAA-NNNNNN (id et lot_id conservés)
        if "annonce_ref" not in annonces.columns:
            annees = annonces["date_annonce"].dt.year.fillna(Config.DATE_FIN.year).astype(int)
            annonces["annonce_ref"] = [f"AN-{a}-{i:06d}" for i, a in enumerate(annees, start=1)]
            self._ecart("annonces_fermes.annonce_ref", "id identique à lot_id",
                        "Colonne annonce_ref ajoutée ; id et lot_id inchangés")

    def _ecart(self, element: str, constat: str, traitement: str) -> None:
        self.ecarts.append({"élément": element, "constat": constat, "traitement": traitement})

    def verifier_cles(self, tables: dict) -> pd.DataFrame:
        """Clés primaires uniques et clés étrangères sans orphelin."""
        lignes = []
        for table, cle in self.CLES_PRIMAIRES.items():
            df = tables.get(table, pd.DataFrame())
            if cle in df.columns:
                lignes.append({"contrôle": f"Unicité {table}.{cle}", "type": "clé primaire",
                               "nb_problèmes": int(df[cle].duplicated().sum())})
        for table, col, ref, ref_col in self.RELATIONS:
            orph = self.orphelins(tables, table, col, ref, ref_col)
            lignes.append({"contrôle": f"{table}.{col} → {ref}.{ref_col}", "type": "clé étrangère",
                           "nb_problèmes": int(len(orph))})
        return pd.DataFrame(lignes)

    @staticmethod
    def orphelins(tables, table, col, ref, ref_col) -> pd.DataFrame:
        """Lignes dont la clé étrangère (renseignée) n'existe pas dans la table référencée."""
        df, df_ref = tables.get(table), tables.get(ref)
        if df is None or df_ref is None or col not in df.columns or ref_col not in df_ref.columns:
            return pd.DataFrame()
        return df[df[col].notna() & ~df[col].isin(df_ref[ref_col])]

    @staticmethod
    def format_iso(df: pd.DataFrame) -> pd.DataFrame:
        """Dates en ISO 8601 : AAAA-MM-JJ ou AAAA-MM-JJTHH:MM:SS."""
        sortie = df.copy()
        for col in sortie.columns:
            if pd.api.types.is_datetime64_any_dtype(sortie[col]):
                valeurs = sortie[col].dropna()
                que_des_dates = bool((valeurs == valeurs.dt.normalize()).all())
                fmt = "%Y-%m-%d" if que_des_dates and not col.startswith("heure") else "%Y-%m-%dT%H:%M:%S"
                sortie[col] = sortie[col].dt.strftime(fmt)
        return sortie

    def exporter_csv(self, tables: dict) -> list[Path]:
        Config.creer_dossiers()
        fichiers = []
        for nom in Config.TABLES_CAHIER:
            chemin = Config.DATA_CLEAN_DIR / f"{nom}.csv"
            try:
                self.format_iso(tables[nom]).to_csv(chemin, index=False, sep=",",
                                                     encoding=Config.ENCODAGE)
                fichiers.append(chemin)
            except PermissionError:
                print(f"[Avertissement] {chemin} est ouvert ailleurs : export ignoré.")
        return fichiers


# ======================================================================
#  3. QUALITÉ DES DONNÉES
# ======================================================================
class ControleQualite:
    """Détecte les anomalies sans jamais modifier ni supprimer les clés."""

    COLONNES = ["table", "cle", "colonne", "type_anomalie", "valeur", "regle_violee"]
    TABLES_TRANSACTIONNELLES = ["annonces_fermes", "historique_production", "historique_temps_de_cycle"]
    COLONNES_NULLABLES = {("mapping_emballages", "client_id")}     # NULL = règle générique

    def __init__(self, tables: dict):
        self.tables = tables
        self._lignes: list[dict] = []
        self.anomalies = pd.DataFrame(columns=self.COLONNES)

    def executer(self, exporter: bool = True) -> pd.DataFrame:
        self._lignes = []
        self._unicite()
        self._cles_etrangeres()
        self._chronologie()
        self._poids_emballe()
        self._taux_dechet()
        self._negatifs()
        self._manquants()
        self._aberrantes()
        self.anomalies = pd.DataFrame(self._lignes, columns=self.COLONNES).drop_duplicates()
        self.anomalies["valeur"] = self.anomalies["valeur"].astype(str)
        if exporter:
            Config.creer_dossiers()
            try:
                self.anomalies.to_csv(Config.FICHIER_ANOMALIES, index=False, encoding=Config.ENCODAGE)
            except PermissionError:
                print(f"[Avertissement] {Config.FICHIER_ANOMALIES} est ouvert ailleurs.")
        return self.anomalies

    def _ajouter(self, table, cle, colonne, type_anomalie, valeur, regle):
        self._lignes.append({"table": table, "cle": str(cle), "colonne": colonne,
                             "type_anomalie": type_anomalie, "valeur": valeur, "regle_violee": regle})

    @staticmethod
    def _cle_ligne(df: pd.DataFrame, i) -> str:
        for c in ("lot_id", "id", "ferme_id", "sku_id", "client_id", "station_id"):
            if c in df.columns:
                return str(df.at[i, c])
        return f"ligne {i + 2}"     # numéro de ligne Excel (en-tête = ligne 1)

    def _unicite(self):
        for table, cle in ChargeurDonnees.CLES_PRIMAIRES.items():
            df = self.tables.get(table)
            if df is not None and cle in df.columns:
                for v in df.loc[df[cle].duplicated(), cle]:
                    self._ajouter(table, v, cle, "format", v, "clé primaire en double")

    def _cles_etrangeres(self):
        for table, col, ref, ref_col in ChargeurDonnees.RELATIONS:
            orph = ChargeurDonnees.orphelins(self.tables, table, col, ref, ref_col)
            for i, v in orph[col].items() if not orph.empty else []:
                type_a = "manquante" if col == "lot_id" else "format"
                self._ajouter(table, self._cle_ligne(self.tables[table], i), col, type_a, v,
                              f"clé étrangère absente de {ref}.{ref_col}")

    def _chronologie(self):
        cyc = self.tables["historique_temps_de_cycle"]
        etapes = Config.ETAPES_CYCLE
        for avant, apres in zip(etapes, etapes[1:]):
            masque = cyc[apres] < cyc[avant]
            for i in cyc.index[masque]:
                self._ajouter("historique_temps_de_cycle", cyc.at[i, "lot_id"], apres, "hors_borne",
                              cyc.at[i, apres], f"{apres} antérieure à {avant}")

    def _poids_emballe(self):
        p = self.tables["historique_production"]
        limite = p["poids_recu_kg"] * (1 - p["taux_dechet"]) * (1 + Config.TOLERANCE_POIDS)
        for i in p.index[p["poids_emballe_kg"] > limite]:
            self._ajouter("historique_production", p.at[i, "lot_id"], "poids_emballe_kg", "hors_borne",
                          p.at[i, "poids_emballe_kg"], "poids_emballe > poids_recu × (1 − taux_dechet) + 1 %")

    def _taux_dechet(self):
        p = self.tables["historique_production"]
        masque = (p["taux_dechet"] < 0) | (p["taux_dechet"] > Config.TAUX_DECHET_MAX)
        for i in p.index[masque]:
            self._ajouter("historique_production", p.at[i, "lot_id"], "taux_dechet", "hors_borne",
                          p.at[i, "taux_dechet"], "taux_dechet hors de [0 ; 0,15]")

    def _negatifs(self):
        for table in Config.TABLES_CAHIER:
            df = self.tables.get(table)
            if df is None or df.empty:
                continue
            for col in df.select_dtypes(include="number").columns:
                for i in df.index[df[col] < 0]:
                    self._ajouter(table, self._cle_ligne(df, i), col, "hors_borne", df.at[i, col],
                                  "valeur négative interdite")

    def _manquants(self):
        for table in Config.TABLES_CAHIER:
            df = self.tables.get(table)
            if df is None or df.empty:
                continue
            for col in df.columns:
                if (table, col) in self.COLONNES_NULLABLES:
                    continue
                for i in df.index[df[col].isna()]:
                    self._ajouter(table, self._cle_ligne(df, i), col, "manquante", "", "valeur obligatoire absente")

    @staticmethod
    def z_robuste(serie: pd.Series) -> pd.Series:
        mediane = serie.median()
        mad = (serie - mediane).abs().median()
        if not mad or pd.isna(mad):
            return pd.Series(0.0, index=serie.index)
        return 0.6745 * (serie - mediane) / mad

    def _aberrantes(self):
        p = self.tables["historique_production"]
        # Rendement incohérent avec poids_emballe ÷ heures_travaillees
        theorique = p["poids_emballe_kg"] / p["heures_travaillees"].replace(0, np.nan)
        ecart = (p["rendement_kg_h"] - theorique) / theorique
        for i in p.index[ecart.abs() > Config.TOLERANCE_RENDEMENT]:
            self._ajouter("historique_production", p.at[i, "lot_id"], "rendement_kg_h", "aberrante",
                          p.at[i, "rendement_kg_h"],
                          f"écart de {ecart[i]:+.0%} avec poids_emballe ÷ heures_travaillees")
        # z-score robuste (médiane / MAD) par SKU
        for col in ("rendement_kg_h", "arrets_min"):
            z = p.groupby("sku_id")[col].transform(self.z_robuste)
            for i in p.index[z.abs() > Config.SEUIL_Z_ROBUSTE]:
                self._ajouter("historique_production", p.at[i, "lot_id"], col, "aberrante",
                              p.at[i, col], f"z-score robuste {z[i]:+.1f} (seuil ±{Config.SEUIL_Z_ROBUSTE})")

    # ------------------------------------------------------------------
    def lots_aberrants(self) -> set:
        a = self.anomalies
        return set(a.loc[(a["table"] == "historique_production") & (a["type_anomalie"] == "aberrante"), "cle"])

    def completude(self) -> pd.DataFrame:
        lignes = []
        for table in Config.TABLES_CAHIER:
            df = self.tables.get(table, pd.DataFrame())
            cellules = int(df.size)
            manquantes = int(df.isna().sum().sum())
            lignes.append({"table": table, "nb_lignes": len(df), "nb_colonnes": df.shape[1],
                           "cellules_manquantes": manquantes,
                           "completude_pct": round(100 * (1 - manquantes / cellules), 2) if cellules else 0.0})
        return pd.DataFrame(lignes)

    def taux_anomalies(self) -> float:
        """Part des lots concernés par au moins une anomalie (cible ≈ 1 %)."""
        nb_lots = len(self.tables["historique_production"])
        lots = set(self.anomalies.loc[self.anomalies["table"].isin(self.TABLES_TRANSACTIONNELLES), "cle"])
        lots &= set(self.tables["historique_production"]["lot_id"])
        return len(lots) / nb_lots if nb_lots else 0.0

    def taux_cycles_chronologiques(self) -> float:
        cyc = self.tables["historique_temps_de_cycle"]
        problemes = self.anomalies.loc[
            (self.anomalies["table"] == "historique_temps_de_cycle")
            & self.anomalies["regle_violee"].str.contains("antérieure"), "cle"].nunique()
        return 1 - problemes / len(cyc) if len(cyc) else 0.0


# ======================================================================
#  4. MOTEUR DE CAPACITÉ
# ======================================================================
class MoteurCapacite:
    """Capacité des postes et des shifts, charge annoncée, surcharge. Calculée à la volée."""

    def __init__(self, tables: dict):
        self.tables = tables
        self.postes = self.capacite_postes()
        self.annonces = self.affecter_annonces()
        self.emballage = self.capacite_emballage()
        self.shifts = self.charge_shifts()

    def capacite_postes(self) -> pd.DataFrame:
        """Capacité opérationnelle = nominale × heures × (opérateurs − absences) / référence × facteur."""
        p = self.tables["planning_shifts"].merge(
            self.tables["stations"][["station_id", "poste", "capacite_nominale_kg_h"]], on="poste", how="left")
        nominale = p["capacite_nominale_kg_h"].astype(float)
        nominale = nominale.where(p["poste"] != Config.POSTE_GOULOT, Config.capacite_goulot_kg_h())
        reference = p["poste"].map(Config.NB_OPERATEURS_REFERENCE).fillna(p["nb_operateurs"]).astype(float)
        presents = (p["nb_operateurs"] - p["absences"]).clip(lower=0)
        facteur = p["shift"].map(lambda s: Config.facteur_shift(s))
        p["capacite_nominale_kg_h"] = nominale
        p["operateurs_presents"] = presents
        p["capacite_nominale_shift_kg"] = nominale * p["heures_disponibles"]
        p["capacite_operationnelle_kg"] = (nominale * p["heures_disponibles"] * presents
                                           / reference.replace(0, np.nan) * facteur).round(1)
        return p

    @staticmethod
    def shift_depuis_heure(horodatage: pd.Series) -> tuple[pd.Series, pd.Series]:
        """Arrivée 06:00–21:59 → shift jour ; sinon nuit (rattachée à la veille avant 06:00)."""
        h = horodatage.dt.hour
        shift = pd.Series(np.where((h >= 6) & (h < 22), "jour", "nuit"), index=horodatage.index)
        date_shift = (horodatage - pd.Timedelta(hours=6)).dt.normalize()
        return shift, date_shift

    def affecter_annonces(self) -> pd.DataFrame:
        a = self.tables["annonces_fermes"].merge(
            self.tables["produits_sku"][["sku_id", "famille", "format", "poids_unitaire_g"]], on="sku_id", how="left")
        a["shift_prevu"], a["date_shift"] = self.shift_depuis_heure(a["arrivee_camion_prevue"])
        a["article_emballage"] = a["format"].map(Config.FORMAT_VERS_ARTICLE)
        return a

    def capacite_emballage(self) -> pd.DataFrame:
        """kg supportés par le stock (tous clients) de chaque article, par jour."""
        s = self.tables["stock_emballages"].groupby(["date_releve", "article"], as_index=False)["stock_actuel"].sum()
        s["kg_supportes"] = [Config.kg_supportes(a, q) for a, q in zip(s["article"], s["stock_actuel"])]
        return s

    def charge_shifts(self) -> pd.DataFrame:
        postes = self.postes
        idx_min = postes.groupby(["date", "shift"])["capacite_operationnelle_kg"].idxmin()
        cap = postes.loc[idx_min, ["date", "shift", "poste", "capacite_operationnelle_kg"]].rename(
            columns={"poste": "poste_limitant", "capacite_operationnelle_kg": "capacite_poste_limitant_kg"})
        goulot = postes[postes["poste"] == Config.POSTE_GOULOT][
            ["date", "shift", "nb_operateurs", "absences", "operateurs_presents", "capacite_nominale_shift_kg"]]
        cap = cap.merge(goulot, on=["date", "shift"], how="left")

        actives = self.annonces_actives()
        charge = actives.groupby(["date_shift", "shift_prevu"]).agg(
            volume_annonce_kg=("volume_annonce_kg", "sum"), nb_annonces=("id", "count"),
            nb_fermes=("ferme_id", "nunique")).reset_index().rename(
            columns={"date_shift": "date", "shift_prevu": "shift"})

        # Capacité supportée par les emballages des articles nécessaires au shift
        besoins = actives[["date_shift", "shift_prevu", "article_emballage"]].drop_duplicates().rename(
            columns={"date_shift": "date", "shift_prevu": "shift", "article_emballage": "article"})
        besoins = besoins.merge(self.emballage.rename(columns={"date_releve": "date"}),
                                on=["date", "article"], how="left")
        cap_emb = besoins.groupby(["date", "shift"], as_index=False)["kg_supportes"].sum(min_count=1).rename(
            columns={"kg_supportes": "capacite_emballage_kg"})

        df = cap.merge(charge, on=["date", "shift"], how="left").merge(cap_emb, on=["date", "shift"], how="left")
        for col in ("volume_annonce_kg", "nb_annonces", "nb_fermes"):
            df[col] = df[col].fillna(0)
        emballage_limitant = df["capacite_emballage_kg"].fillna(np.inf) < df["capacite_poste_limitant_kg"]
        df["tension_emballage"] = emballage_limitant
        if Config.CONTRAINTE_EMBALLAGE_CAPACITE:
            df["capacite_utile_kg"] = np.fmin(df["capacite_poste_limitant_kg"],
                                              df["capacite_emballage_kg"].fillna(np.inf)).round(1)
            df["limite_par"] = np.where(emballage_limitant, "emballages", df["poste_limitant"])
        else:
            df["capacite_utile_kg"] = df["capacite_poste_limitant_kg"].round(1)
            df["limite_par"] = df["poste_limitant"]
        df["taux_utilisation"] = (df["volume_annonce_kg"] / df["capacite_utile_kg"].replace(0, np.nan)).fillna(0)
        df["statut_charge"] = df["taux_utilisation"].map(self.statut_charge)
        df["quantite_a_arbitrer_kg"] = (df["volume_annonce_kg"] - df["capacite_utile_kg"]).clip(lower=0).round(1)
        return df.sort_values(["date", "shift"]).reset_index(drop=True)

    @staticmethod
    def statut_charge(taux: float) -> str:
        if taux > Config.SEUIL_CHARGE_FORTE:
            return "forte_surcharge"
        if taux > Config.SEUIL_CHARGE_SHIFT:
            return "surcharge"
        if taux >= Config.SEUIL_TENSION:
            return "tension"
        return "normal"

    def annonces_actives(self) -> pd.DataFrame:
        return self.annonces[~self.annonces["statut_annonce"].isin(Config.STATUTS_HORS_CHARGE)]

    def taux_annonces_surcharge(self) -> float:
        """Part des annonces maintenues tombant dans un shift en surcharge (cible 10–15 %)."""
        actives = self.annonces_actives()
        m = actives.merge(self.shifts[["date", "shift", "taux_utilisation"]],
                          left_on=["date_shift", "shift_prevu"], right_on=["date", "shift"], how="left")
        return float((m["taux_utilisation"] > Config.SEUIL_CHARGE_SHIFT).mean()) if len(m) else 0.0

    def capacite_theorique(self, nb_operateurs: int, shift: str, absences: int = 0) -> float:
        """Capacité du goulot pour un shift hypothétique (utilisée par la prédiction)."""
        ref = Config.NB_OPERATEURS_REFERENCE[Config.POSTE_GOULOT]
        presents = max(nb_operateurs - absences, 0)
        return Config.capacite_goulot_kg_h() * Config.DUREE_SHIFT_UTILE_H * presents / ref * Config.facteur_shift(shift)

    def charge_poste_jour(self, jour: pd.Timestamp) -> pd.DataFrame:
        """Taux de charge par poste et shift pour une date donnée."""
        p = self.postes[self.postes["date"] == jour][["shift", "poste", "operateurs_presents",
                                                       "capacite_operationnelle_kg"]]
        v = self.shifts[self.shifts["date"] == jour][["shift", "volume_annonce_kg"]]
        p = p.merge(v, on="shift", how="left")
        p["taux_charge"] = p["volume_annonce_kg"] / p["capacite_operationnelle_kg"].replace(0, np.nan)
        return p


# ======================================================================
#  5. DÉRIVE DE RENDEMENT
# ======================================================================
class AnalyseDerive:
    """La dérive n'est jamais une colonne source : elle est calculée."""

    def __init__(self, tables: dict, lots_exclus: set | None = None):
        self.tables = tables
        self.lots_exclus = lots_exclus or set()
        self.lots = self._ratios()
        self.serie = self._serie_glissante()
        self.fenetres = self.statistiques_fenetres()
        self.tests = self.tests_statistiques()

    def _ratios(self) -> pd.DataFrame:
        """
        Rendement observé ÷ rendement de référence.
        Référence = modèle log-linéaire SKU × shift × nb_opérateurs à volume comparable
        (le rendement dépend fortement du volume reçu), ajusté sur les périodes normales
        (hors fenêtres de baisse et de redressement).
        """
        p = self.tables["historique_production"].merge(
            self.tables["produits_sku"][["sku_id", "format", "famille"]], on="sku_id", how="left")
        p = p[~p["lot_id"].isin(self.lots_exclus)].copy().reset_index(drop=True)
        X = pd.get_dummies(p[["sku_id", "shift"]].astype(str).assign(
            nb_operateurs=p["nb_operateurs"].astype(str)), dtype=float)
        X["log_volume"] = np.log(p["poids_recu_kg"].clip(lower=1))
        y = np.log(p["rendement_kg_h"].clip(lower=1))
        debut, fin = pd.Timestamp(Config.DERIVE_BAISSE[0]), pd.Timestamp(Config.DERIVE_REDRESSEMENT[1])
        normal = ((p["date"] < debut) | (p["date"] > fin)).values
        coef, *_ = np.linalg.lstsq(X.values[normal], y.values[normal], rcond=None)
        p["rendement_reference"] = np.exp(X.values @ coef)
        p["ratio"] = p["rendement_kg_h"] / p["rendement_reference"]
        p["derive_lot_pct"] = (p["ratio"] - 1) * 100
        return p

    def _serie_glissante(self) -> pd.DataFrame:
        jour = self.lots.groupby("date")["ratio"].median()
        index = pd.date_range(Config.DATE_DEBUT, Config.DATE_FIN, freq="D")
        jour = jour.reindex(index)
        glissante = jour.rolling(f"{Config.FENETRE_DERIVE_J}D", min_periods=3).median()
        return pd.DataFrame({"date": index, "derive_jour_pct": (jour.values - 1) * 100,
                             "derive_glissante_pct": (glissante.values - 1) * 100})

    def derive_au(self, jour) -> float:
        """Dernière dérive glissante connue STRICTEMENT avant la date (pas de fuite d'information)."""
        s = self.serie[self.serie["date"] < pd.Timestamp(jour)].dropna(subset=["derive_glissante_pct"])
        return float(s["derive_glissante_pct"].iloc[-1]) if len(s) else 0.0

    def _fenetre(self, debut, fin) -> pd.Series:
        d = self.lots["date"]
        return self.lots.loc[(d >= pd.Timestamp(debut)) & (d <= pd.Timestamp(fin)), "ratio"]

    def definitions_fenetres(self) -> dict:
        b0, b1 = Config.DERIVE_BAISSE
        r0, r1 = Config.DERIVE_REDRESSEMENT
        return {
            "Référence": (Config.DATE_DEBUT, b0 - timedelta(days=1)),
            "Baisse": (b0, b1),
            "Redressement": (r0, r1),
            "Après redressement": (r1 + timedelta(days=1), Config.DATE_FIN),
        }

    def statistiques_fenetres(self) -> pd.DataFrame:
        lignes = []
        definitions = self.definitions_fenetres()
        reference = self._fenetre(*definitions["Référence"]).mean()
        reference_med = self._fenetre(*definitions["Référence"]).median()
        for nom, (debut, fin) in definitions.items():
            s = self._fenetre(debut, fin)
            lignes.append({"fenêtre": nom, "début": debut, "fin": fin, "nb_lots": len(s),
                           "dérive_moyenne_pct": round((s.mean() / reference - 1) * 100, 2) if len(s) else np.nan,
                           "dérive_médiane_pct": round((s.median() / reference_med - 1) * 100, 2) if len(s) else np.nan})
        return pd.DataFrame(lignes)

    def tests_statistiques(self) -> pd.DataFrame:
        """Mann-Whitney unilatéral entre fenêtres (repli : test de Welch approché)."""
        f = self.definitions_fenetres()
        comparaisons = [("Référence", "Baisse", "greater"),
                        ("Baisse", "Redressement", "less"),
                        ("Baisse", "Après redressement", "less"),
                        ("Référence", "Après redressement", "two-sided")]
        lignes = []
        for a, b, alt in comparaisons:
            x, y = self._fenetre(*f[a]), self._fenetre(*f[b])
            p_value = self._mann_whitney(x, y, alt)
            lignes.append({"comparaison": f"{a} vs {b}", "hypothèse": {
                "greater": f"rendement {a} > {b}", "less": f"rendement {a} < {b}",
                "two-sided": "rendements différents"}[alt],
                "écart_médianes_pct": round((y.median() - x.median()) * 100, 2),
                "p_value": p_value, "significatif": bool(p_value < Config.SEUIL_P_VALUE)})
        return pd.DataFrame(lignes)

    @staticmethod
    def _mann_whitney(x, y, alternative) -> float:
        if len(x) < 3 or len(y) < 3:
            return 1.0
        try:
            from scipy.stats import mannwhitneyu
            return float(mannwhitneyu(x, y, alternative=alternative).pvalue)
        except ImportError:
            t = (x.mean() - y.mean()) / math.sqrt(x.var() / len(x) + y.var() / len(y))
            p_unilateral = 0.5 * math.erfc(t / math.sqrt(2))
            if alternative == "less":
                return 1 - p_unilateral
            if alternative == "greater":
                return p_unilateral
            return min(1.0, 2 * min(p_unilateral, 1 - p_unilateral))

    def derive_detectee(self) -> bool:
        stats = self.fenetres.set_index("fenêtre")
        baisse = stats.at["Baisse", "dérive_moyenne_pct"]
        test = self.tests.iloc[0]
        return bool(-12 <= baisse <= -3 and test["significatif"])

    def redressement_detecte(self) -> bool:
        """Redressement : rendement après la baisse significativement supérieur."""
        return bool(self.tests.iloc[1]["significatif"] or self.tests.iloc[2]["significatif"])


# ======================================================================
#  6. VARIABLES POUR LE MACHINE LEARNING
# ======================================================================
class ConstructeurVariables:
    """Construit les jeux de données d'apprentissage et la ligne de prédiction."""

    CAT_LOT = ["sku_id", "format", "famille", "shift", "ferme_id", "client_id", "priorite"]
    NUM_LOT = ["volume_kg", "cadence_standard_kg_h", "poids_unitaire_g", "nb_operateurs",
               "taux_dechet_moyen_famille", "mois", "jour_semaine", "derive_recente_pct",
               "couverture_emballage_j", "ratio_charge_capacite", "nb_lots_shift",
               "distance_usine_km", "heure_arrivee_h"]
    CAT_SHIFT = ["shift"]
    NUM_SHIFT = ["mois", "jour_semaine", "operateurs_presents", "absences", "volume_annonce_kg",
                 "nb_annonces", "nb_fermes", "capacite_poste_limitant_kg"]

    def __init__(self, tables: dict, capacite: MoteurCapacite, derive: AnalyseDerive, lots_aberrants: set):
        self.tables = tables
        self.capacite = capacite
        self.derive = derive
        self.lots_aberrants = lots_aberrants
        self.couverture = self._couverture_articles()
        self.consommation_client = self._consommation_par_client()
        prod = tables["historique_production"].merge(tables["produits_sku"], on="sku_id", how="left")
        train = prod[prod["date"] <= pd.Timestamp(Config.ML_FIN_TRAIN)]
        self.dechet_famille = train.groupby("famille")["taux_dechet"].mean().to_dict()
        # Décalage médian (h) entre le début du shift et l'arrivée du camion
        cyc = tables["historique_temps_de_cycle"]
        ps = prod[["lot_id", "shift", "date"]].merge(cyc[["lot_id", "heure_arrivee"]], on="lot_id")
        debut_shift = ps["date"] + pd.to_timedelta(ps["shift"].map(Config.HEURE_DEBUT_SHIFT), unit="h")
        decalage = (ps["heure_arrivee"] - debut_shift).dt.total_seconds() / 3600
        self.decalage_arrivee_h = decalage.groupby(ps["shift"]).median().to_dict()
        serie = derive.serie.set_index("date")["derive_glissante_pct"].ffill().shift(1)
        self.derive_decalee = serie.fillna(0.0)

    # --- Emballages -------------------------------------------------------
    def _couverture_articles(self) -> pd.DataFrame:
        """Couverture en jours par article (tous clients) = stock ÷ consommation moyenne."""
        s = self.tables["stock_emballages"].groupby(["date_releve", "article"], as_index=False)["stock_actuel"].sum()
        s = s.sort_values(["article", "date_releve"])
        baisse = (-s.groupby("article")["stock_actuel"].diff()).clip(lower=0)
        conso = baisse.groupby(s["article"]).mean().rename("consommation_moy_j")
        s = s.merge(conso, left_on="article", right_index=True, how="left")
        s["couverture_j"] = (s["stock_actuel"] / s["consommation_moy_j"].replace(0, np.nan)).clip(upper=365)
        return s

    def _consommation_par_client(self) -> pd.DataFrame:
        s = self.tables["stock_emballages"].sort_values(["article", "client_id", "date_releve"]).copy()
        s["consommation_j"] = (-s.groupby(["article", "client_id"])["stock_actuel"].diff()).clip(lower=0)
        return s

    def etat_stocks(self, stocks: pd.DataFrame | None = None, receptions: pd.DataFrame | None = None) -> pd.DataFrame:
        """Dernier relevé par article × client (+ réceptions saisies depuis) : couverture, seuil et statut."""
        cons = self.consommation_client
        s = cons if stocks is None else stocks
        if s.empty:
            return s
        dernier = s["date_releve"].max()
        conso = cons.groupby(["article", "client_id"])["consommation_j"].mean().rename("conso_moy_j")
        d = s[s["date_releve"] == dernier].merge(conso, left_on=["article", "client_id"], right_index=True, how="left")
        d["receptions_saisies"] = 0.0
        if receptions is not None and len(receptions):
            r = receptions[(receptions["date_reception"] > dernier)
                           & (receptions["date_reception"] <= pd.Timestamp(date.today()))]
            somme = r.groupby(["article", "client_id"])["quantite"].sum()
            d["receptions_saisies"] = [float(somme.get((a, c), 0.0)) for a, c in zip(d["article"], d["client_id"])]
            d["stock_actuel"] = d["stock_actuel"] + d["receptions_saisies"]
        d["categorie"] = d["article"].map(Config.CATEGORIE_ARTICLE)
        d["couverture_j"] = (d["stock_actuel"] / d["conso_moy_j"].replace(0, np.nan)).round(1)
        d["seuil_j"] = d["article"].map(Config.seuil_couverture)
        d["kg_supportes"] = [Config.kg_supportes(a, q) for a, q in zip(d["article"], d["stock_actuel"])]

        def statut(ligne):
            if pd.isna(ligne["couverture_j"]):
                return "neutre"
            if ligne["couverture_j"] < ligne["seuil_j"]:
                return "critique"
            if ligne["couverture_j"] < max(ligne["seuil_j"] * 1.5, ligne["delai_fournisseur_jours"]):
                return "tension"
            return "normal"
        d["statut"] = d.apply(statut, axis=1)
        d["priorite_affichage"] = (d["categorie"] != "sachet_importe").astype(int)
        return d.sort_values(["priorite_affichage", "couverture_j"]).reset_index(drop=True)

    def couverture_au(self, article: str, jour) -> float:
        c = self.couverture[(self.couverture["article"] == article) & (self.couverture["date_releve"] <= pd.Timestamp(jour))]
        if c.empty:
            c = self.couverture[self.couverture["article"] == article]
        return float(c["couverture_j"].iloc[-1]) if len(c) else np.nan

    # --- Jeux d'apprentissage ----------------------------------------------
    def jeu_lots(self) -> pd.DataFrame:
        """Une ligne par lot : variables + cibles (rendement, cycle, étapes)."""
        t = self.tables
        p = t["historique_production"].merge(t["produits_sku"], on="sku_id", how="left")
        p = p.merge(t["annonces_fermes"][["lot_id", "priorite", "volume_annonce_kg"]], on="lot_id", how="left")
        p = p.merge(t["fermes"][["ferme_id", "distance_usine_km"]], on="ferme_id", how="left")
        p = p.merge(t["historique_temps_de_cycle"], on="lot_id", how="left")
        p = p.merge(self.capacite.shifts[["date", "shift", "taux_utilisation"]], on=["date", "shift"], how="left")
        p["volume_kg"] = p["poids_recu_kg"]
        p["taux_dechet_moyen_famille"] = p["famille"].map(self.dechet_famille)
        p["mois"] = p["date"].dt.month
        p["jour_semaine"] = p["date"].dt.dayofweek
        p["derive_recente_pct"] = p["date"].map(self.derive_decalee).fillna(0.0)
        p["article"] = p["format"].map(Config.FORMAT_VERS_ARTICLE)
        cov = self.couverture.rename(columns={"date_releve": "date"})[["date", "article", "couverture_j"]]
        p = p.merge(cov, on=["date", "article"], how="left").rename(columns={"couverture_j": "couverture_emballage_j"})
        p["ratio_charge_capacite"] = p["taux_utilisation"]
        p["nb_lots_shift"] = p.groupby(["date", "shift"])["lot_id"].transform("count")
        p["heure_arrivee_h"] = p["heure_arrivee"].dt.hour + p["heure_arrivee"].dt.minute / 60
        # Cibles du modèle B
        p["duree_cycle_min"] = (p["heure_expedition_reelle"] - p["heure_arrivee"]).dt.total_seconds() / 60
        for nom, (debut, fin) in Config.SEGMENTS_CYCLE.items():
            p[nom] = (p[fin] - p[debut]).dt.total_seconds() / 60
        p["aberrant"] = p["lot_id"].isin(self.lots_aberrants)
        return p.sort_values("date").reset_index(drop=True)

    def jeu_shifts(self) -> pd.DataFrame:
        s = self.capacite.shifts.copy()
        s["mois"] = s["date"].dt.month
        s["jour_semaine"] = s["date"].dt.dayofweek
        s["surcharge"] = (s["taux_utilisation"] > Config.SEUIL_CHARGE_SHIFT).astype(int)
        return s

    # --- Ligne de prédiction ------------------------------------------------
    def heure_arrivee(self, jour, shift) -> pd.Timestamp:
        """Arrivée prévue = début du shift + décalage médian observé."""
        decalage = self.decalage_arrivee_h.get(shift, 1.0)
        return (pd.Timestamp(jour).normalize() + pd.Timedelta(hours=Config.HEURE_DEBUT_SHIFT[shift])
                + pd.Timedelta(hours=float(decalage))).floor("min")

    def ligne_lot(self, volume_kg, sku_id, client_id, ferme_id, jour, shift, nb_operateurs,
                  priorite="NORMALE", ratio_charge=None, nb_lots_shift=1) -> pd.DataFrame:
        sku = self.tables["produits_sku"].set_index("sku_id").loc[sku_id]
        ferme = self.tables["fermes"].set_index("ferme_id").loc[ferme_id]
        jour = pd.Timestamp(jour).normalize()
        article = Config.FORMAT_VERS_ARTICLE.get(sku["format"])
        derive = self.derive_decalee
        derive_val = float(derive[derive.index <= jour].iloc[-1]) if (derive.index <= jour).any() else 0.0
        return pd.DataFrame([{
            "sku_id": sku_id, "format": sku["format"], "famille": sku["famille"], "shift": shift,
            "ferme_id": ferme_id, "client_id": client_id, "priorite": priorite,
            "volume_kg": float(volume_kg), "cadence_standard_kg_h": float(sku["cadence_standard_kg_h"]),
            "poids_unitaire_g": float(sku["poids_unitaire_g"]), "nb_operateurs": int(nb_operateurs),
            "taux_dechet_moyen_famille": self.dechet_famille.get(sku["famille"], np.nan),
            "mois": jour.month, "jour_semaine": jour.dayofweek, "derive_recente_pct": derive_val,
            "couverture_emballage_j": self.couverture_au(article, jour),
            "ratio_charge_capacite": ratio_charge, "nb_lots_shift": nb_lots_shift,
            "distance_usine_km": float(ferme["distance_usine_km"]),
            "heure_arrivee_h": self.heure_arrivee(jour, shift).hour + self.heure_arrivee(jour, shift).minute / 60,
        }])


# ======================================================================
#  7. MODÈLES D'INTELLIGENCE ARTIFICIELLE
# ======================================================================
class ModelesIA:
    """Entraînement (validation temporelle), évaluation, sauvegarde et prédiction."""

    ETAPES = list(Config.SEGMENTS_CYCLE.keys())

    def __init__(self, variables: ConstructeurVariables):
        self.variables = variables
        self.modeles: dict = {}
        self.metriques: dict = {}

    # --- Outils ---------------------------------------------------------
    @staticmethod
    def _pipeline(estimateur, cat, num):
        from sklearn.compose import ColumnTransformer
        from sklearn.impute import SimpleImputer
        from sklearn.pipeline import Pipeline
        from sklearn.preprocessing import OneHotEncoder
        pre = ColumnTransformer([
            ("num", SimpleImputer(strategy="median"), num),
            ("cat", Pipeline([("imp", SimpleImputer(strategy="most_frequent")),
                              ("ohe", OneHotEncoder(handle_unknown="ignore", sparse_output=False))]), cat),
        ])
        return Pipeline([("pre", pre), ("modele", estimateur)])

    @staticmethod
    def _regresseurs():
        from sklearn.ensemble import GradientBoostingRegressor, RandomForestRegressor
        return {
            "GradientBoosting": GradientBoostingRegressor(n_estimators=300, learning_rate=0.05, max_depth=3,
                                                          subsample=0.9, random_state=Config.RANDOM_STATE),
            "RandomForest": RandomForestRegressor(n_estimators=300, min_samples_leaf=3, n_jobs=-1,
                                                  random_state=Config.RANDOM_STATE),
        }

    @staticmethod
    def scores(y_vrai, y_pred) -> dict:
        from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
        return {"MAE": round(float(mean_absolute_error(y_vrai, y_pred)), 3),
                "RMSE": round(float(math.sqrt(mean_squared_error(y_vrai, y_pred))), 3),
                "R2": round(float(r2_score(y_vrai, y_pred)), 4)}

    @staticmethod
    def importances(pipeline, cat, num) -> list:
        modele = pipeline.named_steps["modele"]
        if not hasattr(modele, "feature_importances_"):
            return []
        noms = pipeline.named_steps["pre"].get_feature_names_out()
        totaux = {c: 0.0 for c in num + cat}
        for nom, valeur in zip(noms, modele.feature_importances_):
            brut = nom.split("__", 1)[1]
            if nom.startswith("num__"):
                totaux[brut] += float(valeur)
            else:
                for c in sorted(cat, key=len, reverse=True):
                    if brut.startswith(c + "_"):
                        totaux[c] += float(valeur)
                        break
        return sorted(({"variable": k, "importance": round(v, 4)} for k, v in totaux.items()),
                      key=lambda d: d["importance"], reverse=True)

    @staticmethod
    def _decoupage(df: pd.DataFrame):
        """Validation temporelle : train 2025-10 → 2026-06, test 2026-07 → 2026-09."""
        train = df[df["date"] <= pd.Timestamp(Config.ML_FIN_TRAIN)]
        test = df[df["date"] >= pd.Timestamp(Config.ML_DEBUT_TEST)]
        return train, test

    def _validation_croisee(self, estimateur, X, y, cat, num) -> float:
        from sklearn.base import clone
        from sklearn.metrics import mean_absolute_error
        from sklearn.model_selection import TimeSeriesSplit
        maes = []
        for idx_train, idx_val in TimeSeriesSplit(n_splits=Config.ML_N_SPLITS).split(X):
            pipe = self._pipeline(clone(estimateur), cat, num)
            pipe.fit(X.iloc[idx_train], y.iloc[idx_train])
            maes.append(mean_absolute_error(y.iloc[idx_val], pipe.predict(X.iloc[idx_val])))
        return round(float(np.mean(maes)), 3)

    def _entrainer_regression(self, nom, df, cible, cat, num, journal) -> None:
        from sklearn.base import clone
        donnees = df.dropna(subset=[cible])
        train, test = self._decoupage(donnees)
        X_tr, y_tr, X_te, y_te = train[cat + num], train[cible], test[cat + num], test[cible]
        # Baseline : moyenne par SKU sur la période d'entraînement
        moyenne_sku = train.groupby("sku_id")[cible].mean()
        base = test["sku_id"].map(moyenne_sku).fillna(y_tr.mean())
        resultats = {"cible": cible, "n_train": len(train), "n_test": len(test),
                     "baseline_moyenne_sku": self.scores(y_te, base), "modeles": {}}
        meilleur, meilleur_mae, meilleur_pipe = None, np.inf, None
        for nom_algo, estimateur in self._regresseurs().items():
            pipe = self._pipeline(clone(estimateur), cat, num)
            pipe.fit(X_tr, y_tr)
            sc = self.scores(y_te, pipe.predict(X_te))
            sc["MAE_validation_croisee"] = self._validation_croisee(estimateur, X_tr, y_tr, cat, num)
            resultats["modeles"][nom_algo] = sc
            journal(f"   {nom_algo:<17} MAE={sc['MAE']:.2f}  RMSE={sc['RMSE']:.2f}  R²={sc['R2']:.3f}")
            if sc["MAE"] < meilleur_mae:
                meilleur, meilleur_mae, meilleur_pipe = nom_algo, sc["MAE"], pipe
        b = resultats["baseline_moyenne_sku"]
        journal(f"   {'Baseline SKU':<17} MAE={b['MAE']:.2f}  RMSE={b['RMSE']:.2f}  R²={b['R2']:.3f}")
        resultats["meilleur_modele"] = meilleur
        resultats["gain_vs_baseline_pct"] = round(100 * (1 - meilleur_mae / b["MAE"]), 1) if b["MAE"] else 0.0
        resultats["importances"] = self.importances(meilleur_pipe, cat, num)
        # Modèle final réentraîné sur toute la période (même algorithme)
        final = self._pipeline(clone(self._regresseurs()[meilleur]), cat, num)
        final.fit(donnees[cat + num], donnees[cible])
        self.modeles[nom] = {"pipeline": final, "cat": cat, "num": num, "algorithme": meilleur}
        self.metriques[nom] = resultats

    def _entrainer_etapes(self, df, cat, num, journal) -> None:
        from sklearn.ensemble import RandomForestRegressor
        donnees = df.dropna(subset=self.ETAPES)
        train, test = self._decoupage(donnees)
        pipe = self._pipeline(RandomForestRegressor(n_estimators=300, min_samples_leaf=3, n_jobs=-1,
                                                    random_state=Config.RANDOM_STATE), cat, num)
        pipe.fit(train[cat + num], train[self.ETAPES])
        pred = pd.DataFrame(pipe.predict(test[cat + num]), columns=self.ETAPES, index=test.index)
        detail = {e: self.scores(test[e], pred[e]) for e in self.ETAPES}
        for e, sc in detail.items():
            journal(f"   {Config.LIBELLES_SEGMENTS[e]:<31} MAE={sc['MAE']:.1f} min  R²={sc['R2']:.3f}")
        pipe.fit(donnees[cat + num], donnees[self.ETAPES])
        self.modeles["etapes"] = {"pipeline": pipe, "cat": cat, "num": num, "algorithme": "RandomForest multi-sorties"}
        self.metriques["etapes"] = {"cibles": self.ETAPES, "n_train": len(train), "n_test": len(test), "detail": detail}

    def _entrainer_surcharge(self, journal) -> None:
        """
        Classification surcharge oui/non par date + shift.
        La surcharge est saisonnière (pic d'hiver) : le test juillet–septembre n'en contient
        aucune. L'évaluation se fait donc par TimeSeriesSplit (prédictions hors échantillon
        sur toute l'année, toujours en respectant l'ordre chronologique).
        """
        from sklearn.base import clone
        from sklearn.ensemble import GradientBoostingClassifier
        from sklearn.metrics import accuracy_score, f1_score, precision_score, recall_score, roc_auc_score
        from sklearn.model_selection import TimeSeriesSplit
        v = self.variables
        df = v.jeu_shifts().sort_values(["date", "shift"]).reset_index(drop=True)
        cat, num = v.CAT_SHIFT, v.NUM_SHIFT
        X, y = df[cat + num], df["surcharge"]
        base = GradientBoostingClassifier(n_estimators=200, max_depth=3, learning_rate=0.05,
                                          random_state=Config.RANDOM_STATE)
        probas, verites = [], []
        for idx_tr, idx_te in TimeSeriesSplit(n_splits=Config.ML_N_SPLITS + 1).split(X):
            if y.iloc[idx_tr].nunique() < 2:
                continue
            pipe = self._pipeline(clone(base), cat, num)
            pipe.fit(X.iloc[idx_tr], y.iloc[idx_tr])
            probas.extend(pipe.predict_proba(X.iloc[idx_te])[:, 1])
            verites.extend(y.iloc[idx_te])
        probas, verites = np.array(probas), np.array(verites)
        pred = (probas >= 0.5).astype(int)
        sc = {"validation": "TimeSeriesSplit (prédictions hors échantillon)",
              "n_evalues": int(len(verites)),
              "accuracy": round(float(accuracy_score(verites, pred)), 4),
              "precision": round(float(precision_score(verites, pred, zero_division=0)), 4),
              "rappel": round(float(recall_score(verites, pred, zero_division=0)), 4),
              "F1": round(float(f1_score(verites, pred, zero_division=0)), 4),
              "ROC_AUC": round(float(roc_auc_score(verites, probas)), 4) if len(set(verites)) > 1 else None,
              "part_surcharge": round(float(np.mean(verites)), 4)}
        journal(f"   Accuracy={sc['accuracy']:.3f}  Précision={sc['precision']:.3f}  "
                f"Rappel={sc['rappel']:.3f}  F1={sc['F1']:.3f}  AUC={sc['ROC_AUC']}")
        final = self._pipeline(clone(base), cat, num)
        final.fit(X, y)
        self.modeles["surcharge"] = {"pipeline": final, "cat": cat, "num": num,
                                     "algorithme": "GradientBoostingClassifier"}
        self.metriques["surcharge"] = {"n_shifts": len(df), "scores": sc,
                                       "importances": self.importances(final, cat, num)}

    # --- API publique ------------------------------------------------------
    def entrainer(self, journal=print) -> dict:
        """Entraîne les modèles A, B (total + étapes) et C, puis les sauvegarde."""
        import joblib
        v = self.variables
        lots = v.jeu_lots()
        lots_sains = lots[~lots["aberrant"]]          # les lots aberrants sont exclus de l'apprentissage
        cat, num = v.CAT_LOT, v.NUM_LOT
        journal(f"Lots utilisés : {len(lots_sains)} (exclus car aberrants : {int(lots['aberrant'].sum())})")
        journal(f"Train : {Config.DATE_DEBUT} → {Config.ML_FIN_TRAIN} | Test : {Config.ML_DEBUT_TEST} → {Config.DATE_FIN}")
        journal("\nModèle A — rendement (kg/h)")
        self._entrainer_regression("rendement", lots_sains, "rendement_kg_h", cat, num, journal)
        journal("\nModèle B — durée totale du cycle (min)")
        self._entrainer_regression("cycle", lots_sains, "duree_cycle_min", cat, num, journal)
        journal("\nModèle B — durées par étape (min)")
        self._entrainer_etapes(lots_sains, cat, num, journal)
        journal("\nModèle C — risque de surcharge (par date et shift)")
        self._entrainer_surcharge(journal)

        Config.creer_dossiers()
        for nom, contenu in self.modeles.items():
            joblib.dump(contenu, Config.MODELS_DIR / f"modele_{nom}.joblib")
        self.metriques["_infos"] = {"date_entrainement": datetime.now().isoformat(timespec="seconds"),
                                    "random_state": Config.RANDOM_STATE,
                                    "train": f"{Config.DATE_DEBUT} → {Config.ML_FIN_TRAIN}",
                                    "test": f"{Config.ML_DEBUT_TEST} → {Config.DATE_FIN}",
                                    "capacite_goulot_kg_h": Config.capacite_goulot_kg_h()}
        Config.FICHIER_METRIQUES.write_text(json.dumps(self.metriques, ensure_ascii=False, indent=2, default=str),
                                            encoding=Config.ENCODAGE)
        journal(f"\nModèles sauvegardés dans : {Config.MODELS_DIR}")
        return self.metriques

    def charger(self) -> bool:
        """Charge les modèles .joblib existants. Retourne False s'il en manque."""
        import joblib
        noms = ["rendement", "cycle", "etapes", "surcharge"]
        try:
            for nom in noms:
                self.modeles[nom] = joblib.load(Config.MODELS_DIR / f"modele_{nom}.joblib")
            if Config.FICHIER_METRIQUES.exists():
                self.metriques = json.loads(Config.FICHIER_METRIQUES.read_text(encoding=Config.ENCODAGE))
            return True
        except Exception:
            self.modeles = {}
            return False

    @property
    def prets(self) -> bool:
        return all(n in self.modeles for n in ("rendement", "cycle", "etapes"))

    def _predire(self, nom, ligne):
        m = self.modeles[nom]
        return m["pipeline"].predict(ligne[m["cat"] + m["num"]])

    def predire_lots(self, entrees: pd.DataFrame, charge_externe: pd.DataFrame | None = None,
                     avec_reappro: bool = False) -> pd.DataFrame:
        """
        Prédit plusieurs lots d'un coup (annonces futures).
        Colonnes : volume_kg, sku_id, client_id, ferme_id et soit heure_arrivee (date + heure),
        soit date + shift. Optionnelles : nb_operateurs, priorite, absences.
        Les lots d'un même shift se cumulent pour la charge ; charge_externe ajoute d'autres
        annonces futures (colonnes date_shift, shift, volume_kg).
        Rien n'est écrit dans les tables historiques.
        """
        if not self.prets:
            raise RuntimeError("Modèles non entraînés : lancez d'abord le mode « entrainer ».")
        v, cap = self.variables, self.variables.capacite
        e = entrees.copy().reset_index(drop=True)
        if e.empty:
            return pd.DataFrame()
        for col, defaut in {"priorite": "NORMALE", "absences": 0, "nb_operateurs": 10, "shift": "jour"}.items():
            if col not in e.columns:
                e[col] = defaut
            e[col] = e[col].fillna(defaut)
        e["volume_kg"] = pd.to_numeric(e["volume_kg"], errors="coerce")
        if e["volume_kg"].isna().any() or (e["volume_kg"] <= 0).any():
            raise ValueError("Chaque volume doit être un nombre strictement positif (kg).")
        if not e["shift"].isin(Config.SHIFTS).all():
            raise ValueError("Le shift doit être « jour » ou « nuit ».")
        inconnus = set(e["sku_id"]) - set(v.tables["produits_sku"]["sku_id"])
        inconnus |= set(e["client_id"]) - set(v.tables["clients"]["client_id"])
        inconnus |= set(e["ferme_id"]) - set(v.tables["fermes"]["ferme_id"])
        if inconnus:
            raise ValueError("Identifiants inconnus : " + ", ".join(sorted(map(str, inconnus))))

        # Heure d'arrivée : saisie, sinon début du shift + décalage médian observé
        if "heure_arrivee" in e.columns:
            arrivee = pd.to_datetime(e["heure_arrivee"], errors="coerce")
        else:
            arrivee = pd.Series(pd.NaT, index=e.index, dtype="datetime64[ns]")
        manque = arrivee.isna()
        if manque.any():
            if "date" not in e.columns:
                raise ValueError("Indiquez la date ou l'heure d'arrivée de chaque lot.")
            arrivee = arrivee.astype("datetime64[ns]")
            arrivee.loc[manque] = [v.heure_arrivee(d, s) for d, s in zip(e.loc[manque, "date"], e.loc[manque, "shift"])]
        e["heure_arrivee"] = arrivee
        e["shift"], e["date_shift"] = MoteurCapacite.shift_depuis_heure(arrivee)

        # Charge du shift = historique + lots prédits ensemble + charge externe
        groupe = ["date_shift", "shift"]
        vol_lot = e.groupby(groupe)["volume_kg"].transform("sum").to_numpy(dtype=float)
        nb_lot = e.groupby(groupe)["volume_kg"].transform("count").to_numpy(dtype=float)
        vol_ext = np.zeros(len(e))
        nb_ext = np.zeros(len(e))
        if charge_externe is not None and len(charge_externe):
            ext = (charge_externe.assign(date_shift=pd.to_datetime(charge_externe["date_shift"]))
                   .groupby(groupe).agg(vol_ext=("volume_kg", "sum"), nb_ext=("volume_kg", "count")).reset_index())
            m = e[groupe].merge(ext, on=groupe, how="left")
            vol_ext, nb_ext = m["vol_ext"].fillna(0).to_numpy(), m["nb_ext"].fillna(0).to_numpy()
        hist = cap.shifts.rename(columns={"date": "date_shift"})[
            groupe + ["volume_annonce_kg", "nb_annonces", "nb_fermes", "capacite_utile_kg", "operateurs_presents"]]
        h = e[groupe].merge(hist, on=groupe, how="left")
        connu = h["capacite_utile_kg"].notna().to_numpy()
        theorique = np.array([cap.capacite_theorique(int(n), s, int(a))
                              for n, s, a in zip(e["nb_operateurs"], e["shift"], e["absences"])])
        capacite = np.where(connu, h["capacite_utile_kg"].fillna(0).to_numpy(), theorique)
        charge = h["volume_annonce_kg"].fillna(0).to_numpy() + vol_lot + vol_ext
        nb_total = h["nb_annonces"].fillna(0).to_numpy() + nb_lot + nb_ext
        ratio = np.where(capacite > 0, charge / np.where(capacite > 0, capacite, 1), np.inf)

        # Variables et prédictions (un seul appel par modèle)
        lignes = pd.concat([v.ligne_lot(r.volume_kg, r.sku_id, r.client_id, r.ferme_id, r.date_shift, r.shift,
                                        int(r.nb_operateurs), r.priorite, float(ratio[i]), int(nb_total[i]))
                            for i, r in enumerate(e.itertuples(index=False))], ignore_index=True)
        lignes["heure_arrivee_h"] = (arrivee.dt.hour + arrivee.dt.minute / 60).to_numpy()
        rendement = self._predire("rendement", lignes)
        cycle = self._predire("cycle", lignes)
        etapes = np.asarray(self._predire("etapes", lignes)).reshape(len(e), -1)

        # --- Décomposition COHÉRENTE du cycle (rendement ↔ temps de cycle) ----------------
        # Rendement = kg emballés ÷ heures de production effective (définition de la BDD).
        # Production effective = poids emballé prévu ÷ rendement prévu (exactement).
        # Cycle = attente avant traitement + (tri, attentes inter-postes, arrêts) + production
        #         effective + palettisation + préparation expédition.
        rendement = np.maximum(np.asarray(rendement, dtype=float), 1.0)
        cycle = np.asarray(cycle, dtype=float)
        etapes = np.clip(etapes, 0, None)
        total = etapes.sum(axis=1)
        etapes = etapes * np.where(total > 0, cycle / np.where(total > 0, total, 1), 0)[:, None]
        idx = {nom: i for i, nom in enumerate(self.ETAPES)}
        attente = etapes[:, idx["attente_reception_min"]]
        traitement = etapes[:, [idx["tri_min"], idx["attente_conditionnement_min"], idx["conditionnement_min"]]].sum(axis=1)
        palettisation = etapes[:, idx["palettisation_min"]]
        preparation = etapes[:, idx["expedition_min"]]
        dechet = lignes["taux_dechet_moyen_famille"].fillna(0.03).to_numpy(dtype=float)
        poids_emballe = e["volume_kg"].to_numpy(dtype=float) * (1 - dechet)
        production = poids_emballe / rendement * 60
        hors_production = np.maximum(traitement - production, 0)
        traitement = hors_production + production
        cycle = attente + traitement + palettisation + preparation
        proba = np.full(len(e), np.nan)
        if "surcharge" in self.modeles:
            m = self.modeles["surcharge"]
            presents = np.where(connu, h["operateurs_presents"].fillna(0).to_numpy(),
                                np.maximum(e["nb_operateurs"].to_numpy() - e["absences"].to_numpy(), 0))
            ls = pd.DataFrame({"shift": e["shift"], "mois": e["date_shift"].dt.month,
                               "jour_semaine": e["date_shift"].dt.dayofweek, "operateurs_presents": presents,
                               "absences": e["absences"], "volume_annonce_kg": charge, "nb_annonces": nb_total,
                               "nb_fermes": h["nb_fermes"].fillna(0).to_numpy() + 1,
                               "capacite_poste_limitant_kg": capacite})
            proba = m["pipeline"].predict_proba(ls[m["cat"] + m["num"]])[:, 1]

        # Heures : expédition prévue et heure limite du client
        expedition = arrivee + pd.to_timedelta(cycle, unit="min")
        limites = v.tables["clients"].set_index("client_id")["heure_limite_depart"]
        hl = e["client_id"].map(lambda c: int(str(limites[c])[:2]))
        ml = e["client_id"].map(lambda c: int(str(limites[c])[3:5]))
        limite = expedition.dt.normalize() + pd.to_timedelta(hl, unit="h") + pd.to_timedelta(ml, unit="min")
        marge = (limite - expedition).dt.total_seconds() / 60

        res = pd.DataFrame({
            "volume_kg": e["volume_kg"], "sku_id": e["sku_id"], "client_id": e["client_id"],
            "ferme_id": e["ferme_id"], "priorite": e["priorite"], "nb_operateurs": e["nb_operateurs"].astype(int),
            "heure_arrivee_prevue": arrivee.dt.floor("min"), "shift": e["shift"], "date_shift": e["date_shift"],
            "rendement_prevu_kg_h": np.round(rendement, 1), "duree_cycle_prevue_min": np.round(cycle, 1),
            "heure_expedition_prevue": expedition.dt.floor("min"), "heure_limite_client": limite,
            "marge_avant_limite_min": marge.round(0), "respect_heure_limite": marge >= 0,
            "volume_deja_annonce_kg": np.round(charge - e["volume_kg"].to_numpy(), 1),
            "volume_shift_kg": np.round(charge, 1), "capacite_utile_kg": np.round(capacite, 1),
            "taux_utilisation": np.round(ratio, 3),
            "statut_charge": [MoteurCapacite.statut_charge(x) for x in ratio],
            "risque_surcharge_proba": np.round(proba, 3),
            "quantite_a_arbitrer_kg": np.round(np.maximum(0.0, charge - capacite), 1),
        })
        for i, nom in enumerate(self.ETAPES):
            res[nom] = np.round(etapes[:, i], 1)
        res["taux_dechet_prevu"] = np.round(dechet, 4)
        res["poids_emballe_prevu_kg"] = np.round(poids_emballe, 1)
        res["temps_attente_min"] = np.round(attente, 1)
        res["temps_hors_production_min"] = np.round(hors_production, 1)
        res["temps_production_min"] = np.round(production, 1)
        res["temps_palettisation_min"] = np.round(palettisation, 1)
        res["temps_preparation_expedition_min"] = np.round(preparation, 1)
        minutes = lambda x: pd.to_timedelta(np.asarray(x), unit="min")
        res["heure_debut_traitement"] = (arrivee + minutes(attente)).dt.floor("min")
        res["heure_debut_production"] = (arrivee + minutes(attente + hors_production)).dt.floor("min")
        res["heure_fin_traitement"] = (arrivee + minutes(attente + traitement)).dt.floor("min")
        res["heure_fin_palettisation"] = (arrivee + minutes(attente + traitement + palettisation)).dt.floor("min")
        if not avec_reappro:
            return res
        articles = e["article_emballage"] if "article_emballage" in e.columns else pd.Series([None] * len(e))
        reappro = pd.DataFrame([self.reapprovisionnement(s, c, d, vol, a) for s, c, d, vol, a in
                                zip(e["sku_id"], e["client_id"], e["date_shift"], e["volume_kg"], articles)])
        return pd.concat([res, reappro], axis=1)

    def predire_lot(self, volume_kg, sku_id, client_id, ferme_id, date, shift, nb_operateurs,
                    priorite="NORMALE", absences=0, heure_arrivee=None, charge_externe=None,
                    journaliser=True) -> dict:
        """
        Prédit UN lot annoncé. Renvoie : rendement prévu, durée de cycle prévue, heure
        d'expédition prévue, risque de surcharge, quantité à arbitrer et date de
        réapprovisionnement conseillée. Rien n'est écrit dans les tables historiques.
        """
        entree = pd.DataFrame([{"volume_kg": volume_kg, "sku_id": sku_id, "client_id": client_id,
                                "ferme_id": ferme_id, "date": pd.Timestamp(date).normalize(), "shift": shift,
                                "nb_operateurs": int(nb_operateurs), "priorite": priorite, "absences": int(absences),
                                "heure_arrivee": pd.Timestamp(heure_arrivee) if heure_arrivee is not None else pd.NaT}])
        brut = self.predire_lots(entree, charge_externe, avec_reappro=True).iloc[0].to_dict()
        r = {}
        for cle, valeur in brut.items():
            if hasattr(valeur, "item"):
                valeur = valeur.item()
            if isinstance(valeur, float) and math.isnan(valeur):
                valeur = None
            r[cle] = valeur
        r["durees_etapes_min"] = {k: r.pop(k) for k in self.ETAPES}
        for cle in ("heure_arrivee_prevue", "heure_expedition_prevue", "heure_limite_client",
                    "heure_debut_traitement", "heure_debut_production", "heure_fin_traitement",
                    "heure_fin_palettisation"):
            r[cle] = pd.Timestamp(r[cle]).isoformat(timespec="minutes")
        r["date"] = pd.Timestamp(r.pop("date_shift")).date().isoformat()
        if journaliser:
            self._journaliser(r)
        return r

    def reapprovisionnement(self, sku_id, client_id, jour, volume_kg, article=None) -> dict:
        """Date conseillée = date + (stock actuel ÷ consommation moyenne − délai fournisseur)."""
        v = self.variables
        if not article or (isinstance(article, float) and math.isnan(article)):
            fmt = v.tables["produits_sku"].set_index("sku_id").at[sku_id, "format"]
            article = Config.FORMAT_VERS_ARTICLE.get(fmt)
        s = v.consommation_client[(v.consommation_client["article"] == article)
                                  & (v.consommation_client["client_id"] == client_id)]
        if s.empty:
            return {"article_emballage": article, "date_reappro_conseillee": None}
        avant = s[s["date_releve"] <= jour]
        derniere = (avant if len(avant) else s).iloc[-1]
        stock = float(derniere["stock_actuel"])
        conso = float(s["consommation_j"].mean()) or 1.0
        delai = int(derniere["delai_fournisseur_jours"])
        couverture = stock / conso
        jours_avant_commande = math.floor(couverture - delai)
        date_reappro = jour + pd.Timedelta(days=max(jours_avant_commande, 0))
        return {
            "article_emballage": article,
            "stock_actuel_unites": stock,
            "unite": derniere["unite"],
            "kg_supportes_stock": round(Config.kg_supportes(article, stock), 1),
            "stock_suffisant_pour_lot": bool(Config.kg_supportes(article, stock) >= volume_kg),
            "consommation_moy_unites_j": round(conso, 1),
            "couverture_j": round(couverture, 1),
            "delai_fournisseur_j": delai,
            "date_reappro_conseillee": date_reappro.date().isoformat(),
            "commande_urgente": bool(jours_avant_commande <= 0),
        }

    @staticmethod
    def _journaliser(resultat: dict) -> None:
        """Les prédictions vont dans sorties_moteur/, JAMAIS dans l'historique."""
        try:
            Config.creer_dossiers()
            ligne = {k: v for k, v in resultat.items() if not isinstance(v, dict)}
            ligne["horodatage"] = datetime.now().isoformat(timespec="seconds")
            fichier = Config.SORTIES_DIR / "journal_predictions.csv"
            pd.DataFrame([ligne]).to_csv(fichier, mode="a", index=False, header=not fichier.exists(),
                                         encoding=Config.ENCODAGE)
        except OSError:
            pass


# ======================================================================
#  8. RÉFÉRENTIEL D'EMBALLAGES (mapping SKU + client → article)
# ======================================================================
class ReferentielEmballages:
    """
    Proposition automatique de l'emballage (cahier V2 §8) :
      - le SKU a un emballage par défaut (selon son format) ;
      - une règle SKU + client a priorité sur le défaut SKU ;
      - le manager peut choisir une autre référence compatible (liste déroulante).
    Contrôle du stock : besoin cumulé des annonces futures vs stock du client.
    """

    @staticmethod
    def _client_sachet_personnalise(exigence: str) -> bool:
        texte = str(exigence).lower()
        return ("sachet" in texte and "spécifique" in texte) or "conditionnement personnalis" in texte

    @classmethod
    def generer_mapping(cls, skus: pd.DataFrame, clients: pd.DataFrame) -> pd.DataFrame:
        lignes = []

        def ajouter(sku, client, article, role, priorite, defaut):
            lignes.append({"mapping_id": f"MAP{len(lignes) + 1:04d}", "sku_id": sku, "client_id": client,
                           "article_emballage": article, "role": role, "priorite": priorite, "est_defaut": defaut})

        for _, s in skus.iterrows():
            defaut = Config.FORMAT_VERS_ARTICLE.get(s["format"], "EMBALLAGE_SUR_MESURE")
            categorie = Config.CATEGORIE_ARTICLE.get(defaut)
            alternatives = sorted(a for a, c in Config.CATEGORIE_ARTICLE.items()
                                  if c == categorie and a not in (defaut, "EMBALLAGE_SUR_MESURE"))
            # Règles génériques (client_id vide)
            ajouter(s["sku_id"], None, defaut, "primaire", 1, True)
            for i, a in enumerate(alternatives, start=2):
                ajouter(s["sku_id"], None, a, "primaire", i, False)
            ajouter(s["sku_id"], None, "EMBALLAGE_SUR_MESURE", "primaire", 9, False)
            # Règles spécifiques client
            for _, c in clients.iterrows():
                if str(s["format"]).startswith("Sachet") and cls._client_sachet_personnalise(c.get("exigence_emballage", "")):
                    ajouter(s["sku_id"], c["client_id"], "EMBALLAGE_SUR_MESURE", "primaire", 1, True)
                    ajouter(s["sku_id"], c["client_id"], defaut, "primaire", 2, False)
                etiquette = Config.ETIQUETTE_PAR_CLIENT.get(c["client_id"])
                if etiquette:
                    ajouter(s["sku_id"], c["client_id"], etiquette, "etiquette", 1, True)
        return pd.DataFrame(lignes)

    def __init__(self, tables: dict, variables: "ConstructeurVariables"):
        self.mapping = tables["mapping_emballages"]
        self.sku = tables["produits_sku"].set_index("sku_id")
        self.variables = variables
        self.articles_stock = set(tables["stock_emballages"]["article"])
        self.facteur = self._calibrer(tables)

    def _calibrer(self, tables) -> dict:
        """
        Les relevés de stock synthétiques sont exprimés en « unités de stock » (paquets, colis…)
        dont l'échelle ne correspond pas aux emballages physiques. Facteur calibré sur l'historique :
        facteur = besoin physique moyen / jour ÷ consommation moyenne observée / jour.
        """
        prod = tables["historique_production"]
        jours = (Config.DATE_FIN - Config.DATE_DEBUT).days + 1
        besoins: dict = {}
        for (sku, client), g in prod.groupby(["sku_id", "client_id"]):
            volume = float(g["poids_emballe_kg"].sum())
            for article in (self.proposer(sku, client), self.etiquette(sku, client)):
                if article:
                    besoins[(article, client)] = besoins.get((article, client), 0) + self.unites_physiques(article, sku, volume)
        conso = self.variables.consommation_client.groupby(["article", "client_id"])["consommation_j"].mean()
        self.paires_historiques = {cle for cle, v in besoins.items() if v > 0}
        facteurs = {cle: max(1.0, besoins.get(cle, 0) / jours / c) for cle, c in conso.items() if c > 0}
        # Repli par article (médiane des clients) pour les couples sans historique
        for a in set(a for a, _ in conso.index):
            valeurs = [v for (art, c), v in facteurs.items() if art == a and besoins.get((art, c), 0) > 0]
            facteurs[(a, None)] = float(np.median(valeurs)) if valeurs else 1.0
        return facteurs

    def _regles(self, sku_id, client_id, role="primaire") -> pd.DataFrame:
        m = self.mapping[(self.mapping["sku_id"] == sku_id) & (self.mapping["role"] == role)]
        specifiques = m[m["client_id"] == client_id].assign(rang=0)
        generiques = m[m["client_id"].isna()].assign(rang=1)
        return pd.concat([specifiques, generiques]).sort_values(["rang", "priorite"])

    def proposer(self, sku_id, client_id) -> str:
        """Emballage proposé automatiquement : règle SKU + client, sinon défaut SKU."""
        r = self._regles(sku_id, client_id)
        defauts = r[r["est_defaut"].astype(bool)]
        if len(defauts):
            return str(defauts.iloc[0]["article_emballage"])
        return str(r.iloc[0]["article_emballage"]) if len(r) else "EMBALLAGE_SUR_MESURE"

    def compatibles(self, sku_id, client_id) -> list:
        """Références compatibles (proposée en premier) pour la liste déroulante."""
        r = self._regles(sku_id, client_id)
        liste = [self.proposer(sku_id, client_id)] + r["article_emballage"].tolist()
        return list(dict.fromkeys(liste))

    def etiquette(self, sku_id, client_id) -> str | None:
        r = self._regles(sku_id, client_id, role="etiquette")
        return str(r.iloc[0]["article_emballage"]) if len(r) else Config.ETIQUETTE_PAR_CLIENT.get(client_id)

    def unites_necessaires(self, article, sku_id, volume_kg, client_id=None) -> float:
        """Besoin exprimé en unités de stock (après calibrage article × client) pour un volume en kg."""
        physiques = self.unites_physiques(article, sku_id, volume_kg)
        f = getattr(self, "facteur", {})
        facteur = f.get((article, client_id)) or f.get((article, None)) or 1.0
        return math.ceil(physiques / facteur)

    def unites_physiques(self, article, sku_id, volume_kg) -> float:
        """Nombre d'emballages physiques (sachets, cartons, étiquettes) pour un volume en kg."""
        poids_g = float(self.sku.at[sku_id, "poids_unitaire_g"]) if sku_id in self.sku.index else 0.0
        if str(article).startswith("ETIQUETTE"):
            return math.ceil(volume_kg / Config.KG_PAR_COLIS_ETIQUETE)       # 1 étiquette par colis
        if str(article).startswith("SACHET") or (article == "EMBALLAGE_SUR_MESURE" and poids_g and poids_g < 1000):
            return math.ceil(volume_kg * 1000 / poids_g) if poids_g else 0
        kg = Config.kg_par_unite(article) or 1.0
        return math.ceil(volume_kg / kg)

    def _stock(self, article, client_id, jour, receptions: pd.DataFrame | None = None) -> dict:
        """Stock disponible à une date : dernier relevé + réceptions saisies depuis ce relevé (jusqu'à la date)."""
        cons = self.variables.consommation_client
        s = cons[(cons["article"] == article) & (cons["client_id"] == client_id)]
        if s.empty:
            return {"stock": 0.0, "unite": "", "delai": 0, "conso": 0.0, "receptions": 0.0}
        avant = s[s["date_releve"] <= pd.Timestamp(jour)]
        derniere = (avant if len(avant) else s).iloc[-1]
        recu = 0.0
        if receptions is not None and len(receptions):
            r = receptions[(receptions["article"] == article) & (receptions["client_id"] == client_id)
                           & (receptions["date_reception"] > derniere["date_releve"])
                           & (receptions["date_reception"] <= pd.Timestamp(jour))]
            recu = float(r["quantite"].sum())
        return {"stock": float(derniere["stock_actuel"]) + recu, "unite": str(derniere["unite"]),
                "delai": int(derniere["delai_fournisseur_jours"]), "conso": float(s["consommation_j"].mean() or 0),
                "receptions": recu}

    def plan(self, vue: pd.DataFrame | None, receptions: pd.DataFrame | None = None,
             horizon_j: int | None = None, avec_fond: bool = False, aujourd_hui=None) -> tuple:
        """
        Plan de réapprovisionnement par article × client, jour par jour :
          stock projeté = dernier relevé + réceptions saisies − consommation prévue.
        Consommation prévue = besoins des annonces futures enregistrées ; avec « consommation de fond »,
        chaque jour consomme au moins la moyenne apprise (pour les arrivées pas encore annoncées).
        Date de commande = date de rupture − délai fournisseur.
        Quantité conseillée = de quoi finir l'horizon avec le stock de sécurité (seuil de couverture).
        Retourne (résumé, projections journalières).
        """
        horizon_j = int(horizon_j or Config.HORIZON_PLAN_J)
        auj = pd.Timestamp(aujourd_hui or date.today()).normalize()
        fin = auj + pd.Timedelta(days=horizon_j)
        lignes = []
        if vue is not None and len(vue):
            v = vue[vue["statut_prediction"] != "ARCHIVEE"] if "statut_prediction" in vue.columns else vue
            for l in v.itertuples(index=False):
                jour = pd.Timestamp(l.date_arrivee).normalize()
                for article in (l.article_emballage, self.etiquette(l.sku_id, l.client_id)):
                    if article and isinstance(article, str):
                        lignes.append({"date": jour, "article": article, "client_id": l.client_id,
                                       "besoin": self.unites_necessaires(article, l.sku_id, float(l.volume_kg), l.client_id)})
        besoins = pd.DataFrame(lignes, columns=["date", "article", "client_id", "besoin"])
        paires = set(zip(besoins["article"], besoins["client_id"]))
        if avec_fond:
            paires |= getattr(self, "paires_historiques", set())
        cons = self.variables.consommation_client
        derniers = cons.sort_values("date_releve").groupby(["article", "client_id"]).tail(1).set_index(["article", "client_id"])
        conso_moy = cons.groupby(["article", "client_id"])["consommation_j"].mean()
        rec = receptions if receptions is not None and len(receptions) else pd.DataFrame(
            columns=["date_reception", "article", "client_id", "quantite"])
        resume, projections = [], []
        for article, client in sorted(paires):
            if (article, client) not in derniers.index:
                continue
            d0 = derniers.loc[(article, client)]
            date0, stock = pd.Timestamp(d0["date_releve"]), float(d0["stock_actuel"])
            delai, unite = int(d0["delai_fournisseur_jours"]), str(d0["unite"])
            bj = besoins[(besoins["article"] == article) & (besoins["client_id"] == client)].groupby("date")["besoin"].sum()
            r = rec[(rec["article"] == article) & (rec["client_id"] == client) & (rec["date_reception"] > date0)]
            rj = r.groupby("date_reception")["quantite"].sum()
            fond = float(conso_moy.get((article, client), 0.0)) if avec_fond else 0.0
            rupture, stock_auj, conso_horizon, besoins_annonces = None, None, 0.0, 0.0
            for jour in pd.date_range(date0 + pd.Timedelta(days=1), fin, freq="D"):
                annonce = float(bj.get(jour, 0.0))
                consommation = max(fond, annonce)
                stock = stock + float(rj.get(jour, 0.0)) - consommation
                if jour <= auj:
                    stock_auj = stock
                if jour >= auj:
                    conso_horizon += consommation
                    besoins_annonces += annonce
                    projections.append({"date": jour, "article": article, "client_id": client, "stock_projete": stock,
                                        "consommation": consommation, "reception": float(rj.get(jour, 0.0))})
                if rupture is None and stock < 0:
                    rupture = jour
            if stock_auj is None:
                stock_auj = float(d0["stock_actuel"]) + float(rj[rj.index <= auj].sum())
            securite = Config.seuil_couverture(article) * (conso_horizon / horizon_j if horizon_j else 0)
            quantite = max(0.0, securite - stock)
            quantite = float(math.ceil(quantite / 10) * 10) if quantite > 0 else 0.0
            commande = rupture - pd.Timedelta(days=delai) if rupture is not None else None
            if rupture is not None and commande <= auj:
                statut, message = "critique", "à commander immédiatement (le délai fournisseur ne permet plus d'éviter la rupture)"
            elif rupture is not None and commande <= auj + pd.Timedelta(days=Config.ALERTE_COMMANDE_J):
                statut, message = "tension", f"commander avant le {commande:%d/%m/%Y}"
            elif rupture is not None:
                statut, message = "normal", f"commande à prévoir le {commande:%d/%m/%Y}"
            elif quantite > 0:
                statut, message = "tension", "pas de rupture, mais stock de sécurité entamé en fin d'horizon"
            else:
                statut, message = "normal", "stock suffisant sur l'horizon"
            resume.append({"article": article, "client_id": client, "unite": unite, "date_releve": date0,
                           "stock_releve": float(d0["stock_actuel"]), "receptions_saisies": float(r["quantite"].sum()),
                           "stock_aujourdhui": round(stock_auj, 0), "besoins_annonces": round(besoins_annonces, 0),
                           "consommation_fond_j": round(fond, 1), "consommation_horizon": round(conso_horizon, 0),
                           "stock_fin_horizon": round(stock, 0), "stock_securite": round(securite, 0),
                           "date_rupture": rupture, "delai_fournisseur_j": delai, "date_commande_limite": commande,
                           "quantite_conseillee": quantite, "statut": statut, "message": message})
        colonnes = ["article", "client_id", "unite", "date_releve", "stock_releve", "receptions_saisies",
                    "stock_aujourdhui", "besoins_annonces", "consommation_fond_j", "consommation_horizon",
                    "stock_fin_horizon", "stock_securite", "date_rupture", "delai_fournisseur_j",
                    "date_commande_limite", "quantite_conseillee", "statut", "message"]
        res = pd.DataFrame(resume, columns=colonnes)
        if len(res):
            res["ordre"] = res["statut"].map({"critique": 0, "tension": 1, "normal": 2})
            res = res.sort_values(["ordre", "date_commande_limite"], na_position="last").drop(columns="ordre")
        return res.reset_index(drop=True), pd.DataFrame(projections)

    def verifier(self, lots: pd.DataFrame, receptions: pd.DataFrame | None = None) -> pd.DataFrame:
        """
        Contrôle du stock d'emballage pour des lots futurs, dans l'ordre d'arrivée :
        les besoins se cumulent par article et client. Colonnes attendues : annonce_id,
        heure_arrivee, sku_id, client_id, volume_kg, article_emballage.
        Statut : ok / risque (couverture restante < seuil de couverture de l'article) / rupture (stock insuffisant).
        Le délai fournisseur sert à dater la commande conseillée.
        """
        sortie = []
        cumul: dict = {}
        ordre = lots.sort_values("heure_arrivee")
        for _, l in ordre.iterrows():
            resultat = {"annonce_id": l["annonce_id"]}
            pire, motifs = "ok", []
            for role, article in (("primaire", l["article_emballage"]), ("etiquette", self.etiquette(l["sku_id"], l["client_id"]))):
                if not article:
                    continue
                besoin = self.unites_necessaires(article, l["sku_id"], float(l["volume_kg"]), l["client_id"])
                st_ = self._stock(article, l["client_id"], l["heure_arrivee"], receptions)
                cle = (article, l["client_id"])
                cumul[cle] = cumul.get(cle, 0) + besoin
                restant = st_["stock"] - cumul[cle]
                couverture = restant / st_["conso"] if st_["conso"] > 0 else np.inf
                if restant < 0:
                    statut = "rupture"
                    motifs.append(f"{article} : besoin cumulé {cumul[cle]:,.0f} > stock {st_['stock']:,.0f} {st_['unite']}".replace(",", " "))
                elif couverture < Config.seuil_couverture(article):
                    statut = "risque"
                    motifs.append(f"{article} : couverture restante {couverture:.1f} j < seuil {Config.seuil_couverture(article)} j "
                                  f"(délai fournisseur {st_['delai']} j)")
                else:
                    statut = "ok"
                if {"ok": 0, "risque": 1, "rupture": 2}[statut] > {"ok": 0, "risque": 1, "rupture": 2}[pire]:
                    pire = statut
                    resultat["article_en_tension"] = article
                if role == "primaire":
                    jours_avant = math.floor(couverture - st_["delai"]) if np.isfinite(couverture) else 365
                    resultat.update({
                        "besoin_emballage_unites": besoin, "stock_emballage_unites": st_["stock"],
                        "unite_emballage": st_["unite"], "stock_restant_unites": round(restant, 0),
                        "couverture_restante_j": round(couverture, 1) if np.isfinite(couverture) else None,
                        "delai_fournisseur_j": st_["delai"],
                        "date_reappro_conseillee": (pd.Timestamp(l["heure_arrivee"]).normalize()
                                                    + pd.Timedelta(days=max(jours_avant, 0))),
                    })
            resultat["statut_emballage"] = pire
            resultat.setdefault("article_en_tension", None)
            resultat["motif_emballage"] = " ; ".join(motifs)
            sortie.append(resultat)
        return pd.DataFrame(sortie)


class GestionnaireReceptions:
    """Réapprovisionnements d'emballages reçus (saisis par le manager), stockés hors de la base."""

    COLONNES = ["reception_id", "date_reception", "article", "client_id", "quantite", "unite", "commentaire",
                "date_saisie"]

    def __init__(self, fichier: Path | None = None):
        self.fichier = Path(fichier) if fichier else Config.FICHIER_RECEPTIONS

    def charger(self) -> pd.DataFrame:
        if not self.fichier.exists():
            return pd.DataFrame(columns=self.COLONNES)
        try:
            df = pd.read_csv(self.fichier, encoding=Config.ENCODAGE, dtype=str, keep_default_na=False)
        except Exception:
            return pd.DataFrame(columns=self.COLONNES)
        df["date_reception"] = pd.to_datetime(df["date_reception"], errors="coerce").dt.normalize()
        df["date_saisie"] = pd.to_datetime(df["date_saisie"], errors="coerce")
        df["quantite"] = pd.to_numeric(df["quantite"], errors="coerce").fillna(0.0)
        return df.sort_values("date_reception", ascending=False)

    def _ecrire(self, df: pd.DataFrame) -> None:
        self.fichier.parent.mkdir(parents=True, exist_ok=True)
        ChargeurDonnees.format_iso(df.reindex(columns=self.COLONNES)).to_csv(self.fichier, index=False,
                                                                             encoding=Config.ENCODAGE)

    def ajouter(self, article: str, client_id: str, quantite: float, jour, unite: str = "",
                commentaire: str = "") -> str:
        if quantite is None or float(quantite) <= 0:
            raise ValueError("La quantité reçue doit être strictement positive.")
        df = self.charger()
        numeros = df["reception_id"].astype(str).str.extract(r"(\d+)$")[0].astype(float)
        rid = f"REC-{(int(numeros.max()) + 1) if len(numeros.dropna()) else 1:05d}"
        nouveau = pd.DataFrame([{"reception_id": rid, "date_reception": pd.Timestamp(jour).normalize(),
                                 "article": article, "client_id": client_id, "quantite": float(quantite),
                                 "unite": unite, "commentaire": commentaire,
                                 "date_saisie": pd.Timestamp.now().floor("s")}])
        self._ecrire(pd.concat([d for d in (df, nouveau) if len(d)], ignore_index=True))
        Config.journaliser_admin("reception_emballage", f"{rid} {article} {client_id} {quantite} {unite}")
        return rid

    def supprimer(self, ids) -> None:
        df = self.charger()
        self._ecrire(df[~df["reception_id"].isin(list(ids))])
        Config.journaliser_admin("suppression_reception", ", ".join(map(str, ids)))


# ======================================================================
#  9. ESPACE PRÉDICTION : ANNONCES FUTURES, PRÉDICTIONS, TRAÇABILITÉ
# ======================================================================
class GestionnairePrevisions:
    """
    Nouvelles annonces (saisies ou importées) et résultats du moteur prédictif.
    Stockage SÉPARÉ de l'historique d'apprentissage (cahier V2 §5.2) :
      sorties_moteur/annonces_futures.csv       → entrées futures
      sorties_moteur/predictions.csv            → prédiction courante de chaque annonce
      sorties_moteur/historique_predictions.csv → chaque calcul (entrées + sorties + versions)
      sorties_moteur/notifications.csv          → alertes de la cloche 🔔
    """

    COLONNES_ANNONCE = ["annonce_id", "lot_id", "date_annonce", "date_arrivee", "heure_arrivee",
                        "depart_camion_prevu", "ferme_id", "sku_id", "client_id", "volume_kg",
                        "article_emballage", "emballage_propose", "emballage_modifie", "priorite",
                        "origine_annonce", "nb_operateurs", "absences", "date_saisie"]
    COLONNES_PREDICTION = ["prediction_id", "annonce_id", "date_prediction", "version_modele", "version_parametres",
                           "shift", "date_shift", "temps_cycle_predit_min", "rendement_prevu_kg_h",
                           "taux_dechet_prevu", "poids_emballe_prevu_kg", "temps_attente_min",
                           "temps_hors_production_min", "temps_production_min", "temps_palettisation_min",
                           "temps_preparation_expedition_min",
                           "heure_arrivee_prevue", "heure_debut_traitement", "heure_debut_production", "heure_fin_predite",
                           "heure_fin_palettisation", "heure_expedition_predite", "heure_limite_client",
                           "marge_expedition_min", "capacite_disponible_kg", "charge_annoncee_kg",
                           "taux_utilisation_pct", "depassement_kg", "etat_charge", "risque_surcharge_proba",
                           "emballage_selectionne", "statut_emballage", "motif_emballage", "article_en_tension",
                           "besoin_emballage_unites", "stock_emballage_unites", "unite_emballage",
                           "date_reappro_conseillee", "niveau_risque", "motifs_risque", "statut_prediction"]
    COLONNES_IMPORT = ["lot_id", "date_annonce", "date_arrivee", "heure_arrivee", "depart_camion_prevu", "ferme_id",
                       "sku_id", "client_id", "volume_kg", "article_emballage", "priorite", "nb_operateurs", "absences"]
    ALIAS = {"date": "date_arrivee", "date_arrivee_prevue": "date_arrivee", "jour": "date_arrivee",
             "heure": "heure_arrivee", "heure_arrivee_prevue": "heure_arrivee", "arrivee": "heure_arrivee",
             "ferme": "ferme_id", "client": "client_id", "sku": "sku_id", "produit": "sku_id",
             "volume": "volume_kg", "volume_annonce_kg": "volume_kg", "quantite": "volume_kg",
             "quantite_kg": "volume_kg", "operateurs": "nb_operateurs", "lot": "lot_id",
             "emballage": "article_emballage", "article": "article_emballage", "depart_camion": "depart_camion_prevu",
             "depart": "depart_camion_prevu", "origine": "origine_annonce"}
    CHAMPS_SUIVIS = ["temps_cycle_predit_min", "heure_expedition_predite", "taux_utilisation_pct",
                     "emballage_selectionne", "niveau_risque", "statut_emballage"]

    def __init__(self, ctx: "ContexteDonnees", dossier: Path | None = None):
        self.ctx = ctx
        base = Path(dossier) if dossier else Config.SORTIES_DIR
        self.f_annonces = base / Config.FICHIER_ANNONCES_FUTURES.name
        self.f_predictions = base / Config.FICHIER_PREDICTIONS.name
        self.f_histo = base / Config.FICHIER_HISTO_PREDICTIONS.name
        self.notifications = GestionnaireNotifications(base / Config.FICHIER_NOTIFICATIONS.name)
        self.receptions = GestionnaireReceptions(base / Config.FICHIER_RECEPTIONS.name)
        self.dossier = base
        self.emballages = ctx.emballages
        self._migrer_ancien_format()

    def _migrer_ancien_format(self) -> None:
        """Reprend les annonces saisies avec la version 2.1 (colonne ref_prevision) au nouveau format."""
        if not self.f_annonces.exists():
            return
        try:
            entete = self.f_annonces.read_text(encoding=Config.ENCODAGE).splitlines()[0]
        except (OSError, IndexError):
            return
        if "ref_prevision" not in entete:
            return
        ancien = pd.read_csv(self.f_annonces, encoding=Config.ENCODAGE, dtype=str)
        sauvegarde = self.f_annonces.with_name("annonces_futures_v2_1_ancien.csv")
        self.f_annonces.replace(sauvegarde)
        colonnes = [c for c in ("date_arrivee", "heure_arrivee", "ferme_id", "client_id", "sku_id", "volume_kg",
                                "priorite", "nb_operateurs", "absences") if c in ancien.columns]
        valides, _ = self.valider(ancien[colonnes], origine_defaut="MANUEL")
        if len(valides):
            self.ajouter(valides)
    # --- Lecture -------------------------------------------------------------
    @staticmethod
    def _lire(fichier: Path, colonnes: list, dates: tuple = ()) -> pd.DataFrame:
        if not fichier.exists():
            return pd.DataFrame(columns=colonnes)
        try:
            df = pd.read_csv(fichier, encoding=Config.ENCODAGE, dtype=str, keep_default_na=False, na_values=[""])
        except Exception:
            return pd.DataFrame(columns=colonnes)
        for col in colonnes:
            if col not in df.columns:
                df[col] = np.nan
        for col in dates:
            df[col] = pd.to_datetime(df[col], errors="coerce")
        return df

    def annonces(self) -> pd.DataFrame:
        df = self._lire(self.f_annonces, self.COLONNES_ANNONCE, ("date_annonce", "date_arrivee", "date_saisie"))
        for col in ("volume_kg",):
            df[col] = pd.to_numeric(df[col], errors="coerce")
        for col in ("nb_operateurs", "absences"):
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
        df["emballage_modifie"] = df["emballage_modifie"].astype(str).str.lower().eq("true")
        return df

    def predictions(self) -> pd.DataFrame:
        dates = ("date_prediction", "date_shift", "heure_arrivee_prevue", "heure_debut_traitement", "heure_debut_production",
                 "heure_fin_predite", "heure_fin_palettisation", "heure_expedition_predite",
                 "heure_limite_client", "date_reappro_conseillee")
        df = self._lire(self.f_predictions, self.COLONNES_PREDICTION, dates)
        for col in ("temps_cycle_predit_min", "rendement_prevu_kg_h", "taux_dechet_prevu", "poids_emballe_prevu_kg",
                    "temps_attente_min", "temps_hors_production_min", "temps_production_min", "temps_palettisation_min",
                    "temps_preparation_expedition_min", "marge_expedition_min", "capacite_disponible_kg",
                    "charge_annoncee_kg", "taux_utilisation_pct", "depassement_kg", "risque_surcharge_proba",
                    "besoin_emballage_unites", "stock_emballage_unites"):
            df[col] = pd.to_numeric(df[col], errors="coerce")
        return df

    def vue(self) -> pd.DataFrame:
        """Annonces futures + leur prédiction courante (une ligne par annonce)."""
        a, p = self.annonces(), self.predictions()
        if a.empty:
            return a.reindex(columns=self.COLONNES_ANNONCE + [c for c in self.COLONNES_PREDICTION if c != "annonce_id"])
        return a.merge(p, on="annonce_id", how="left")

    def historique(self) -> pd.DataFrame:
        return self._lire(self.f_histo, ["date_prediction"], ("date_prediction",))

    # --- Écriture --------------------------------------------------------------
    def _ecrire(self, df: pd.DataFrame, fichier: Path) -> None:
        self.dossier.mkdir(parents=True, exist_ok=True)
        ChargeurDonnees.format_iso(df).to_csv(fichier, index=False, encoding=Config.ENCODAGE)

    def _version_modele(self) -> str:
        infos = self.ctx.ia.metriques.get("_infos", {}) if self.ctx.ia.metriques else {}
        return "M-" + str(infos.get("date_entrainement", "inconnue")).replace(":", "").replace("-", "")[:15]

    def recalculer(self, nouvelles: set | None = None) -> pd.DataFrame:
        """
        Recalcule la prédiction de toutes les annonces non archivées (les lots d'un même shift
        se cumulent), conserve les statuts, historise les changements et met à jour les notifications.
        """
        nouvelles = set(nouvelles or [])
        a = self.annonces()
        anciennes = self.predictions()
        archivees = set(anciennes.loc[anciennes["statut_prediction"] == "ARCHIVEE", "annonce_id"])
        actives = a[~a["annonce_id"].isin(archivees)].reset_index(drop=True)
        if actives.empty or not self.ctx.ia.prets:
            garder = anciennes[anciennes["annonce_id"].isin(archivees)]
            self._ecrire(garder.reindex(columns=self.COLONNES_PREDICTION), self.f_predictions)
            self.notifications.generer(self.vue(), set(), self._plan_alertes())
            return garder
        jours = actives["date_arrivee"].dt.strftime("%Y-%m-%d")
        arrivee = pd.to_datetime(jours + " " + actives["heure_arrivee"].astype(str), errors="coerce")
        entree = pd.DataFrame({
            "volume_kg": actives["volume_kg"].astype(float), "sku_id": actives["sku_id"],
            "client_id": actives["client_id"], "ferme_id": actives["ferme_id"], "priorite": actives["priorite"],
            "nb_operateurs": actives["nb_operateurs"].astype(int), "absences": actives["absences"].astype(int),
            "heure_arrivee": arrivee, "article_emballage": actives["article_emballage"]})
        r = self.ctx.ia.predire_lots(entree)
        emb = self.emballages.verifier(pd.DataFrame({
            "annonce_id": actives["annonce_id"], "heure_arrivee": arrivee, "sku_id": actives["sku_id"],
            "client_id": actives["client_id"], "volume_kg": actives["volume_kg"],
            "article_emballage": actives["article_emballage"]}), self.receptions.charger()
        ).set_index("annonce_id").reindex(actives["annonce_id"])
        maintenant = pd.Timestamp.now().floor("s")
        p = pd.DataFrame({
            "annonce_id": actives["annonce_id"].to_numpy(), "date_prediction": maintenant,
            "version_modele": self._version_modele(), "version_parametres": Config.version_parametres(),
            "shift": r["shift"].to_numpy(), "date_shift": r["date_shift"].to_numpy(),
            "temps_cycle_predit_min": r["duree_cycle_prevue_min"].to_numpy(),
            "rendement_prevu_kg_h": r["rendement_prevu_kg_h"].to_numpy(),
            "taux_dechet_prevu": r["taux_dechet_prevu"].to_numpy(),
            "poids_emballe_prevu_kg": r["poids_emballe_prevu_kg"].to_numpy(),
            "temps_attente_min": r["temps_attente_min"].to_numpy(),
            "temps_hors_production_min": r["temps_hors_production_min"].to_numpy(),
            "temps_production_min": r["temps_production_min"].to_numpy(),
            "temps_palettisation_min": r["temps_palettisation_min"].to_numpy(),
            "temps_preparation_expedition_min": r["temps_preparation_expedition_min"].to_numpy(),
            "heure_arrivee_prevue": r["heure_arrivee_prevue"].to_numpy(),
            "heure_debut_traitement": r["heure_debut_traitement"].to_numpy(),
            "heure_debut_production": r["heure_debut_production"].to_numpy(),
            "heure_fin_predite": r["heure_fin_traitement"].to_numpy(),
            "heure_fin_palettisation": r["heure_fin_palettisation"].to_numpy(),
            "heure_expedition_predite": r["heure_expedition_prevue"].to_numpy(),
            "heure_limite_client": r["heure_limite_client"].to_numpy(),
            "marge_expedition_min": r["marge_avant_limite_min"].to_numpy(),
            "capacite_disponible_kg": r["capacite_utile_kg"].to_numpy(),
            "charge_annoncee_kg": r["volume_shift_kg"].to_numpy(),
            "taux_utilisation_pct": (r["taux_utilisation"] * 100).round(1).to_numpy(),
            "depassement_kg": r["quantite_a_arbitrer_kg"].to_numpy(),
            "etat_charge": r["statut_charge"].to_numpy(),
            "risque_surcharge_proba": r["risque_surcharge_proba"].to_numpy(),
            "emballage_selectionne": actives["article_emballage"].to_numpy(),
        })
        for col in ("statut_emballage", "motif_emballage", "article_en_tension", "besoin_emballage_unites",
                    "stock_emballage_unites", "unite_emballage", "date_reappro_conseillee"):
            p[col] = emb[col].to_numpy() if col in emb.columns else np.nan
        niveaux, motifs = [], []
        for _, l in p.iterrows():
            m, niveau = [], "normal"
            if l["etat_charge"] in ("surcharge", "forte_surcharge"):
                m.append(f"surcharge : dépassement {l['depassement_kg']:,.0f} kg".replace(",", " "))
                niveau = "critique"
            if l["marge_expedition_min"] < 0:
                m.append(f"retard d'expédition de {-l['marge_expedition_min']:.0f} min")
                niveau = "critique"
            if niveau != "critique":
                if l["etat_charge"] == "tension":
                    m.append(f"charge proche de la capacité ({l['taux_utilisation_pct']:.0f} %)")
                    niveau = "tension"
                if 0 <= l["marge_expedition_min"] < Config.SEUIL_MARGE_EXPEDITION_MIN:
                    m.append(f"marge d'expédition faible ({l['marge_expedition_min']:.0f} min)")
                    niveau = "tension"
            if l["statut_emballage"] in ("rupture", "risque"):
                m.append(f"emballage : {l['statut_emballage']}")
                niveau = "tension" if niveau == "normal" else niveau
            niveaux.append(niveau)
            motifs.append(" ; ".join(m) if m else "aucun risque détecté")
        p["niveau_risque"], p["motifs_risque"] = niveaux, motifs

        # Statuts et identifiants : CREEE / RECALCULEE / VALIDEE conservée si rien ne change
        prec = anciennes.set_index("annonce_id") if len(anciennes) else pd.DataFrame()
        numeros = anciennes["prediction_id"].dropna().astype(str).str.extract(r"(\d+)$")[0].astype(float)
        prochain = int(numeros.max()) + 1 if len(numeros.dropna()) else 1
        ids, statuts, a_historiser = [], [], []
        for i, l in p.iterrows():
            aid = l["annonce_id"]
            if aid in nouvelles or aid not in prec.index or pd.isna(prec.at[aid, "prediction_id"]):
                ids.append(f"PRD-{prochain:06d}")
                prochain += 1
                statuts.append("CREEE")
                a_historiser.append(i)
                continue
            ancien = prec.loc[aid]
            change = any(str(self._normaliser(ancien[c])) != str(self._normaliser(l[c])) for c in self.CHAMPS_SUIVIS)
            ids.append(ancien["prediction_id"])
            if change:
                statuts.append("RECALCULEE")
                a_historiser.append(i)
            else:
                statuts.append(ancien["statut_prediction"] if pd.notna(ancien["statut_prediction"]) else "CREEE")
                p.at[i, "date_prediction"] = pd.to_datetime(ancien["date_prediction"])
        p["prediction_id"], p["statut_prediction"] = ids, statuts
        p = p.reindex(columns=self.COLONNES_PREDICTION)
        tout = pd.concat([d for d in (p, anciennes[anciennes["annonce_id"].isin(archivees)]
                                      .reindex(columns=self.COLONNES_PREDICTION)) if len(d)], ignore_index=True)
        self._ecrire(tout, self.f_predictions)
        if a_historiser:
            self._historiser(p.loc[a_historiser], actives)
        self.notifications.generer(self.vue(), nouvelles, self._plan_alertes())
        return tout

    def _plan_alertes(self) -> pd.DataFrame:
        """Plan de réapprovisionnement (annonces enregistrées uniquement) pour les alertes de commande."""
        try:
            resume, _ = self.emballages.plan(self.vue(), self.receptions.charger())
            return resume[resume["statut"].isin(["critique", "tension"]) & resume["date_rupture"].notna()]
        except Exception:
            return pd.DataFrame()

    def plan_reappro(self, horizon_j=None, avec_fond=False) -> tuple:
        return self.emballages.plan(self.vue(), self.receptions.charger(), horizon_j, avec_fond)

    def enregistrer_reception(self, article, client_id, quantite, jour, commentaire="") -> str:
        """Enregistre une réception puis recalcule prédictions, tensions et alertes."""
        unite = self.emballages._stock(article, client_id, jour)["unite"]
        rid = self.receptions.ajouter(article, client_id, quantite, jour, unite, commentaire)
        self.recalculer()
        return rid

    def supprimer_receptions(self, ids) -> None:
        self.receptions.supprimer(ids)
        self.recalculer()

    # --- Prochain batch conseillé ---------------------------------------------------
    def prochain_batch(self, date_shift, shift, vue: pd.DataFrame | None = None, _profondeur: int = 0) -> dict:
        """
        Quand demander aux fermes d'envoyer le batch suivant sur un shift :
          1. le poste goulot traite les lots l'un après l'autre (ordre de début de production) →
             heure à laquelle la ligne se libère ;
          2. arrivée conseillée = ligne libre − délai de préparation (attente + tri) d'un lot,
             pour que le lot soit prêt pile quand la ligne se libère ;
          3. départ de chaque ferme = arrivée conseillée − trajet (distance ÷ vitesse camion) ;
          4. quantité maximale = min(capacité restante du shift, temps restant × rendement prévu).
        Estimation : ordre d'arrivée, sans aléa.
        """
        v = self.vue() if vue is None else vue
        v = v[(v["statut_prediction"] != "ARCHIVEE") & v["date_shift"].notna()] if len(v) else v
        jour = pd.Timestamp(date_shift).normalize()
        g = v[(pd.to_datetime(v["date_shift"]) == jour) & (v["shift"] == shift)] if len(v) else v
        debut_shift = jour + pd.Timedelta(hours=Config.HEURE_DEBUT_SHIFT[shift])
        fin_fenetre = jour + pd.Timedelta(hours=22) if shift == "jour" else jour + pd.Timedelta(days=1, hours=6)
        maintenant = pd.Timestamp.now().ceil("5min")
        reference = g if len(g) else v
        mediane = lambda col, defaut: float(reference[col].median()) if len(reference) and reference[col].notna().any() else defaut
        preparation = mediane("temps_attente_min", 20) + mediane("temps_hors_production_min", 85)
        rendement = mediane("rendement_prevu_kg_h", 600)
        palettisation = mediane("temps_palettisation_min", 60)
        expedition = mediane("temps_preparation_expedition_min", 55)
        # Ordonnancement séquentiel du poste goulot
        ligne_libre, file_min = debut_shift, 0.0
        lots = g.copy()
        if len(lots):
            # Valeurs de repli si une prédiction est incomplète (ex. créée par une ancienne version)
            secours_debut = pd.to_datetime(lots["heure_debut_traitement"]).fillna(pd.to_datetime(lots["heure_arrivee_prevue"]))
            lots["debut_calc"] = pd.to_datetime(lots["heure_debut_production"]).fillna(secours_debut)
            secours_prod = lots["volume_kg"].astype(float) * 0.97 / max(rendement, 1) * 60
            lots["production_calc"] = pd.to_numeric(lots["temps_production_min"], errors="coerce").fillna(secours_prod)
            lots = lots.dropna(subset=["debut_calc", "production_calc"])
        for l in lots.sort_values("debut_calc").itertuples(index=False) if len(lots) else []:
            debut = max(ligne_libre, pd.Timestamp(l.debut_calc))
            file_min += (debut - pd.Timestamp(l.debut_calc)).total_seconds() / 60
            ligne_libre = debut + pd.Timedelta(minutes=float(l.production_calc))
        arrivee = max(ligne_libre - pd.Timedelta(minutes=preparation), debut_shift, maintenant).ceil("5min")
        if len(g):
            capacite, charge = float(g["capacite_disponible_kg"].iloc[0]), float(g["charge_annoncee_kg"].iloc[0])
        else:
            capacite = self.ctx.capacite.capacite_theorique(Config.OPERATEURS_DEFAUT[shift], shift)
            charge = 0.0
        restant = max(0.0, capacite - charge)
        debut_prod = max(arrivee + pd.Timedelta(minutes=preparation), ligne_libre)
        heures_restantes = max(0.0, (fin_fenetre - debut_prod).total_seconds() / 3600)
        quantite = min(restant, heures_restantes * rendement)
        quantite = float(math.floor(quantite / 50) * 50)
        sature = quantite < 200 or arrivee >= fin_fenetre
        fermes = self.ctx.tables["fermes"][["ferme_id", "distance_usine_km"]].copy()
        fermes["trajet_min"] = (fermes["distance_usine_km"] / Config.VITESSE_CAMION_KM_H * 60).round(0)
        fermes["depart_conseille"] = arrivee - pd.to_timedelta(fermes["trajet_min"], unit="min")
        fermes = fermes.sort_values("distance_usine_km").reset_index(drop=True)
        fin_production = debut_prod + pd.Timedelta(hours=quantite * 0.97 / rendement if rendement else 0)
        info = {"date_shift": jour, "shift": shift, "nb_lots": int(len(g)), "ligne_libre": ligne_libre,
                "file_attente_min": round(file_min, 0), "arrivee_conseillee": arrivee,
                "preparation_min": round(preparation, 0), "capacite_kg": round(capacite, 0),
                "charge_kg": round(charge, 0), "capacite_restante_kg": round(restant, 0),
                "rendement_kg_h": round(rendement, 0), "quantite_max_kg": quantite, "sature": bool(sature),
                "expedition_estimee": fin_production + pd.Timedelta(minutes=palettisation + expedition),
                "fermes": fermes, "suivant": None}
        if sature and _profondeur < 3:
            suivant = (jour, "nuit") if shift == "jour" else (jour + pd.Timedelta(days=1), "jour")
            info["suivant"] = self.prochain_batch(suivant[0], suivant[1], v, _profondeur + 1)
        return info

    @staticmethod
    def _normaliser(x):
        if isinstance(x, (pd.Timestamp, datetime)):
            return pd.Timestamp(x).strftime("%Y-%m-%d %H:%M")
        try:
            return round(float(x), 0)
        except (TypeError, ValueError):
            return x

    def _historiser(self, p: pd.DataFrame, annonces: pd.DataFrame) -> None:
        """Traçabilité : entrées utilisées + sorties + versions du modèle et des paramètres."""
        entrees = annonces.set_index("annonce_id")
        lignes = p.copy()
        lignes["entrees"] = [json.dumps({k: str(entrees.at[a, k]) for k in
                                         ("lot_id", "date_arrivee", "heure_arrivee", "ferme_id", "sku_id", "client_id",
                                          "volume_kg", "article_emballage", "priorite", "nb_operateurs", "absences",
                                          "origine_annonce")}, ensure_ascii=False) for a in lignes["annonce_id"]]
        self.dossier.mkdir(parents=True, exist_ok=True)
        ChargeurDonnees.format_iso(lignes).to_csv(self.f_histo, mode="a", index=False, header=not self.f_histo.exists(),
                                                  encoding=Config.ENCODAGE)

    # --- Actions du manager ------------------------------------------------------
    def _prochain_id_annonce(self, existantes: pd.DataFrame, n: int) -> list:
        """AN-AAAA-NNNNNN à la suite des annonces historiques (pas de collision avec l'historique)."""
        hist = self.ctx.tables["annonces_fermes"]["annonce_ref"].astype(str)
        tous = pd.concat([hist, existantes["annonce_id"].astype(str)])
        numeros = tous.str.extract(r"(\d+)$")[0].astype(float)
        depart = int(numeros.max()) + 1 if len(numeros.dropna()) else 1
        annee = date.today().year
        return [f"AN-{annee}-{depart + i:06d}" for i in range(n)]

    def lots_existants(self) -> set:
        return set(self.ctx.tables["historique_production"]["lot_id"]) | set(
            self.ctx.tables["annonces_fermes"]["lot_id"]) | set(self.annonces()["lot_id"].dropna())

    def ajouter(self, valides: pd.DataFrame, origine: str | None = None) -> pd.DataFrame:
        """Enregistre de nouvelles annonces futures puis lance la prédiction."""
        existantes = self.annonces()
        nouveaux = valides.copy().reset_index(drop=True)
        nouveaux["annonce_id"] = self._prochain_id_annonce(existantes, len(nouveaux))
        if origine:
            nouveaux["origine_annonce"] = origine
        nouveaux["date_saisie"] = pd.Timestamp.now().floor("s")
        tout = pd.concat([d.reindex(columns=self.COLONNES_ANNONCE) for d in (existantes, nouveaux) if len(d)],
                         ignore_index=True)
        self._ecrire(tout, self.f_annonces)
        self.recalculer(set(nouveaux["annonce_id"]))
        return self.vue()[lambda d: d["annonce_id"].isin(nouveaux["annonce_id"])]

    def remplacer(self, valides: pd.DataFrame) -> None:
        """Enregistre les annonces modifiées (avec leurs annonce_id) puis recalcule."""
        existantes = self.annonces().set_index("annonce_id")
        v = valides.set_index("annonce_id")
        for col in ("date_saisie", "origine_annonce"):
            v[col] = existantes[col].reindex(v.index).fillna(v[col] if col in v.columns else np.nan)
        self._ecrire(v.reset_index().reindex(columns=self.COLONNES_ANNONCE), self.f_annonces)
        self.recalculer()

    def changer_statut(self, ids, statut: str) -> None:
        if statut not in Config.STATUTS_PREDICTION:
            raise ValueError(f"Statut inconnu : {statut}")
        p = self.predictions()
        p.loc[p["annonce_id"].isin(list(ids)), "statut_prediction"] = statut
        self._ecrire(p, self.f_predictions)
        Config.journaliser_admin(f"statut_{statut.lower()}", ", ".join(map(str, ids)))
        self.recalculer()       # une annonce archivée ne charge plus son shift

    def supprimer(self, ids) -> None:
        ids = list(ids)
        a = self.annonces()
        self._ecrire(a[~a["annonce_id"].isin(ids)], self.f_annonces)
        p = self.predictions()
        self._ecrire(p[~p["annonce_id"].isin(ids)], self.f_predictions)
        Config.journaliser_admin("suppression_annonces", ", ".join(ids))
        self.recalculer()

    def tout_supprimer(self) -> None:
        for f in (self.f_annonces, self.f_predictions, self.notifications.fichier):
            if f.exists():
                f.unlink()
        Config.journaliser_admin("suppression_totale_annonces_futures")

    # --- Validation des saisies et des imports -----------------------------------
    @staticmethod
    def _manquant(x) -> bool:
        if x is None or x is pd.NaT:
            return True
        if isinstance(x, float) and math.isnan(x):
            return True
        return isinstance(x, str) and (not x.strip() or x.strip().lower() in ("nan", "none", "nat"))

    @classmethod
    def _date(cls, x):
        if cls._manquant(x):
            return None
        if isinstance(x, (datetime, date, pd.Timestamp)):
            return pd.Timestamp(x).normalize()
        texte = str(x).strip()
        d = pd.to_datetime(texte, dayfirst="/" in texte, errors="coerce")
        return None if pd.isna(d) else pd.Timestamp(d).normalize()

    @classmethod
    def _heure(cls, x, defaut="08:00"):
        import re
        if cls._manquant(x):
            return defaut
        if isinstance(x, (heure, datetime, pd.Timestamp)):
            return x.strftime("%H:%M")
        if isinstance(x, (int, float)) and 0 <= float(x) < 1:          # heure Excel (fraction de jour)
            minutes = round(float(x) * 24 * 60)
            return f"{minutes // 60:02d}:{minutes % 60:02d}"
        m = re.fullmatch(r"\s*(\d{1,2})\s*[:hH]\s*(\d{0,2})(:\d{2})?\s*", str(x))
        if not m:
            return None
        h, mi = int(m.group(1)), int(m.group(2) or 0)
        return f"{h:02d}:{mi:02d}" if 0 <= h <= 23 and 0 <= mi <= 59 else None

    @classmethod
    def _datetime(cls, x, jour_reference=None):
        """Date-heure complète, ou simple heure (rattachée à la date de référence)."""
        if cls._manquant(x):
            return None
        if isinstance(x, (datetime, pd.Timestamp)):
            return pd.Timestamp(x)
        h = cls._heure(x, defaut=None)
        if h and jour_reference is not None:
            return pd.Timestamp(jour_reference) + pd.Timedelta(hours=int(h[:2]), minutes=int(h[3:]))
        texte = str(x).strip()
        d = pd.to_datetime(texte, dayfirst="/" in texte, errors="coerce")
        return None if pd.isna(d) else pd.Timestamp(d)

    @classmethod
    def _nombre(cls, x, defaut=None):
        if cls._manquant(x):
            return defaut
        try:
            return float(str(x).replace("\u202f", "").replace("\xa0", "").replace(" ", "").replace(",", "."))
        except ValueError:
            return None

    def valider(self, brut: pd.DataFrame, origine_defaut: str = "MANUEL", controle_lots: bool = True) -> tuple:
        """
        Contrôle des champs obligatoires, des références et des formats (cahier V2 §7).
        Retourne (lignes valides, erreurs avec numéro de ligne et motif).
        """
        t = self.ctx.tables
        df = brut.copy()
        df.columns = [self.ALIAS.get(ChargeurDonnees._cle(c), ChargeurDonnees._cle(c)) for c in df.columns]
        colonnes_erreur = ["ligne", "colonne", "message"]
        manquantes = [c for c in ("date_arrivee", "ferme_id", "client_id", "sku_id", "volume_kg") if c not in df.columns]
        if manquantes:
            return (pd.DataFrame(columns=self.COLONNES_ANNONCE),
                    pd.DataFrame([{"ligne": "-", "colonne": c, "message": "colonne obligatoire absente"}
                                  for c in manquantes], columns=colonnes_erreur))
        fermes = {str(f).upper(): f for f in t["fermes"]["ferme_id"]}
        clients = {str(c).upper(): c for c in t["clients"]["client_id"]}
        clients.update({str(n).upper(): c for c, n in zip(t["clients"]["client_id"], t["clients"]["nom"])})
        skus = {str(s).upper(): s for s in t["produits_sku"]["sku_id"]}
        priorites = {"NORMALE": "NORMALE", "NORMAL": "NORMALE", "HAUTE": "HAUTE", "URGENTE": "URGENTE", "URGENT": "URGENTE"}
        lots_pris = self.lots_existants() if controle_lots else set()
        lots_fichier = set()
        valides, erreurs = [], []
        for position, (_, l) in enumerate(df.iterrows()):
            numero = position + 2
            pb = []
            jour = self._date(l.get("date_arrivee"))
            if jour is None:
                pb.append(("date_arrivee", "date d'arrivée invalide (AAAA-MM-JJ ou JJ/MM/AAAA)"))
            h = self._heure(l.get("heure_arrivee"))
            if h is None:
                pb.append(("heure_arrivee", "heure d'arrivée invalide (HH:MM)"))
            date_annonce = self._date(l.get("date_annonce")) or pd.Timestamp(date.today())
            ferme = fermes.get(str(l.get("ferme_id")).strip().upper())
            if ferme is None:
                pb.append(("ferme_id", f"ferme inconnue « {l.get('ferme_id')} »"))
            client = clients.get(str(l.get("client_id")).strip().upper())
            if client is None:
                pb.append(("client_id", f"client inconnu « {l.get('client_id')} »"))
            sku = skus.get(str(l.get("sku_id")).strip().upper())
            if sku is None:
                pb.append(("sku_id", f"SKU inconnu « {l.get('sku_id')} »"))
            volume = self._nombre(l.get("volume_kg"))
            if volume is None or not 0 < volume <= 50000:
                pb.append(("volume_kg", "volume invalide (nombre de kg > 0 et ≤ 50 000)"))
            texte_prio = "NORMALE" if self._manquant(l.get("priorite")) else str(l.get("priorite")).strip().upper()
            prio = priorites.get(texte_prio)
            if prio is None:
                pb.append(("priorite", "priorité attendue : Normale, Haute ou Urgente"))
            lot = None if self._manquant(l.get("lot_id")) else str(l.get("lot_id")).strip()
            if lot and controle_lots and (lot in lots_pris or lot in lots_fichier):
                pb.append(("lot_id", f"lot « {lot} » déjà existant"))
            depart = self._datetime(l.get("depart_camion_prevu"), jour)
            if not self._manquant(l.get("depart_camion_prevu")) and depart is None:
                pb.append(("depart_camion_prevu", "départ camion invalide (AAAA-MM-JJ HH:MM ou HH:MM)"))
            proposition, article = None, None
            if sku and client:
                proposition = self.emballages.proposer(sku, client)
                article = proposition if self._manquant(l.get("article_emballage")) else str(l.get("article_emballage")).strip().upper()
                if article not in self.emballages.compatibles(sku, client):
                    pb.append(("article_emballage", f"emballage « {article} » non compatible avec {sku} / {client}"))
            shift = "jour" if h and 6 <= int(h[:2]) < 22 else "nuit"
            ops = self._nombre(l.get("nb_operateurs"), Config.OPERATEURS_DEFAUT[shift])
            if ops is None or not 1 <= ops <= 30:
                pb.append(("nb_operateurs", "nombre d'opérateurs entre 1 et 30"))
            absences = self._nombre(l.get("absences"), 0)
            if absences is None or absences < 0 or (ops is not None and absences > ops):
                pb.append(("absences", "absences entre 0 et le nombre d'opérateurs"))
            if depart is not None and jour is not None and h and depart > jour + pd.Timedelta(hours=int(h[:2]), minutes=int(h[3:])):
                pb.append(("depart_camion_prevu", "le départ camion doit précéder l'arrivée"))
            if pb:
                erreurs += [{"ligne": numero, "colonne": c, "message": m} for c, m in pb]
                continue
            if not lot:
                lot = self._generer_lot(jour, lots_pris | lots_fichier)
            lots_fichier.add(lot)
            origine = str(l.get("origine_annonce")).strip().upper() if not self._manquant(l.get("origine_annonce")) else origine_defaut
            ligne = {"lot_id": lot, "date_annonce": date_annonce, "date_arrivee": jour, "heure_arrivee": h,
                     "depart_camion_prevu": depart, "ferme_id": ferme, "sku_id": sku, "client_id": client,
                     "volume_kg": round(volume, 1), "article_emballage": article, "emballage_propose": proposition,
                     "emballage_modifie": article != proposition, "priorite": prio,
                     "origine_annonce": origine if origine in Config.ORIGINES_SAISIE else origine_defaut,
                     "nb_operateurs": int(ops), "absences": int(absences)}
            if "annonce_id" in df.columns and not self._manquant(l.get("annonce_id")):
                ligne["annonce_id"] = l.get("annonce_id")
            valides.append(ligne)
        return pd.DataFrame(valides), pd.DataFrame(erreurs, columns=colonnes_erreur)

    @staticmethod
    def _generer_lot(jour, pris: set) -> str:
        n = 1
        while f"LOT-P-{pd.Timestamp(jour):%Y%m%d}-{n:03d}" in pris:
            n += 1
        return f"LOT-P-{pd.Timestamp(jour):%Y%m%d}-{n:03d}"

    @staticmethod
    def lire_fichier(nom: str, contenu: bytes) -> pd.DataFrame:
        """Lit un import .xlsx / .xls / .csv (séparateur et encodage détectés)."""
        if nom.lower().endswith((".xlsx", ".xlsm", ".xls")):
            return pd.read_excel(io.BytesIO(contenu))
        for encodage in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                return pd.read_csv(io.BytesIO(contenu), sep=None, engine="python", encoding=encodage)
            except (UnicodeDecodeError, pd.errors.ParserError, ValueError):
                continue
        raise ErreurDonnees("Fichier illisible : utilisez un CSV (UTF-8) ou un fichier Excel .xlsx.")

    @staticmethod
    def modele_import() -> pd.DataFrame:
        demain = date.today() + timedelta(days=1)
        apres = demain + timedelta(days=1)
        return pd.DataFrame([
            {"lot_id": "", "date_annonce": date.today().isoformat(), "date_arrivee": demain.isoformat(),
             "heure_arrivee": "07:30", "depart_camion_prevu": "", "ferme_id": "NG1", "sku_id": "SKU004",
             "client_id": "CLI001", "volume_kg": 1500, "article_emballage": "", "priorite": "HAUTE",
             "nb_operateurs": 10, "absences": 0},
            {"lot_id": "", "date_annonce": date.today().isoformat(), "date_arrivee": demain.isoformat(),
             "heure_arrivee": "10:15", "depart_camion_prevu": "09:30", "ferme_id": "NG3", "sku_id": "SKU001",
             "client_id": "Lidl", "volume_kg": 2200, "article_emballage": "", "priorite": "Normale",
             "nb_operateurs": 10, "absences": 0},
            {"lot_id": "", "date_annonce": date.today().isoformat(), "date_arrivee": apres.isoformat(),
             "heure_arrivee": "23:00", "depart_camion_prevu": "", "ferme_id": "NG8", "sku_id": "SKU016",
             "client_id": "CLI004", "volume_kg": 3075, "article_emballage": "CARTON_EXPORT_5KG",
             "priorite": "Urgente", "nb_operateurs": 9, "absences": 1},
        ])


# ======================================================================
#  10. NOTIFICATIONS (panneau de la cloche 🔔)
# ======================================================================
class GestionnaireNotifications:
    """
    Alertes produites par les prédictions (cahier V2 §12) :
      critique (surcharge, dépassement en kg) · avertissement (retard / marge insuffisante)
      · attention (tension emballage) · information (nouvelle prédiction).
    Les anomalies de données restent dans l'écran Qualité données.
    """

    COLONNES = ["notification_id", "date_creation", "niveau", "type", "annonce_id", "titre", "message", "lu"]
    NIVEAUX = {"critique": ("🔴", "Critique"), "avertissement": ("🟠", "Avertissement"),
               "attention": ("🟡", "Attention"), "information": ("ℹ️", "Information")}
    ORDRE = {"critique": 0, "avertissement": 1, "attention": 2, "information": 3}

    def __init__(self, fichier: Path | None = None):
        self.fichier = Path(fichier) if fichier else Config.FICHIER_NOTIFICATIONS

    def charger(self) -> pd.DataFrame:
        if not self.fichier.exists():
            return pd.DataFrame(columns=self.COLONNES)
        try:
            df = pd.read_csv(self.fichier, encoding=Config.ENCODAGE, dtype=str, keep_default_na=False)
        except Exception:
            return pd.DataFrame(columns=self.COLONNES)
        df["lu"] = df["lu"].astype(str).str.lower().eq("true")
        df["date_creation"] = pd.to_datetime(df["date_creation"], errors="coerce")
        df["ordre"] = df["niveau"].map(self.ORDRE).fillna(9)
        return df.sort_values(["lu", "ordre", "date_creation"], ascending=[True, True, False]).drop(columns="ordre")

    def _ecrire(self, df: pd.DataFrame) -> None:
        self.fichier.parent.mkdir(parents=True, exist_ok=True)
        ChargeurDonnees.format_iso(df.reindex(columns=self.COLONNES)).to_csv(self.fichier, index=False,
                                                                             encoding=Config.ENCODAGE)

    def generer(self, vue: pd.DataFrame, nouvelles: set, plan: pd.DataFrame | None = None) -> pd.DataFrame:
        """Reconstruit les alertes à partir des prédictions courantes (états lus conservés)."""
        anciennes = self.charger().set_index("notification_id") if self.fichier.exists() else pd.DataFrame()
        maintenant = pd.Timestamp.now().floor("s")
        lignes = []

        def ajouter(nid, niveau, type_, annonce, titre, message):
            lignes.append({"notification_id": nid, "niveau": niveau, "type": type_, "annonce_id": annonce,
                           "titre": titre, "message": message})

        v = vue[vue["statut_prediction"].notna() & (vue["statut_prediction"] != "ARCHIVEE")] if len(vue) else vue
        if len(v):
            groupes = v.groupby(["date_shift", "shift"])
            for (jour, shift), g in groupes:
                l = g.iloc[0]
                if l["etat_charge"] in ("surcharge", "forte_surcharge"):
                    ajouter(f"SURCHARGE-{pd.Timestamp(jour):%Y%m%d}-{shift}", "critique", "surcharge", "",
                            f"Surcharge prédite — {pd.Timestamp(jour):%d/%m/%Y}, shift {shift}",
                            f"Charge {l['charge_annoncee_kg']:,.0f} kg pour {l['capacite_disponible_kg']:,.0f} kg "
                            f"de capacité ({l['taux_utilisation_pct']:.0f} %) · dépassement "
                            f"{l['depassement_kg']:,.0f} kg · {len(g)} lot(s) prévu(s)".replace(",", " "))
            for _, l in v.iterrows():
                ref = f"{l['lot_id']} ({l['annonce_id']})"
                if l["marge_expedition_min"] < 0:
                    ajouter(f"RETARD-{l['annonce_id']}", "avertissement", "retard", l["annonce_id"],
                            f"Risque de retard — {ref}",
                            f"Expédition prévue {pd.Timestamp(l['heure_expedition_predite']):%d/%m %H:%M} après la limite "
                            f"client {pd.Timestamp(l['heure_limite_client']):%H:%M} ({-l['marge_expedition_min']:.0f} min)")
                elif l["marge_expedition_min"] < Config.SEUIL_MARGE_EXPEDITION_MIN:
                    ajouter(f"MARGE-{l['annonce_id']}", "avertissement", "marge", l["annonce_id"],
                            f"Marge d'expédition insuffisante — {ref}",
                            f"Seulement {l['marge_expedition_min']:.0f} min avant la limite client "
                            f"{pd.Timestamp(l['heure_limite_client']):%H:%M}")
            tensions = v[v["statut_emballage"].isin(["rupture", "risque"])]
            for (article, client), g in tensions.groupby([tensions["article_en_tension"].fillna(tensions["emballage_selectionne"]),
                                                          "client_id"]):
                pire = "rupture" if (g["statut_emballage"] == "rupture").any() else "risque"
                ajouter(f"EMBALLAGE-{article}-{client}", "attention", "emballage", ", ".join(g["annonce_id"]),
                        f"Tension emballage — {article} ({client})",
                        ("Stock insuffisant" if pire == "rupture" else "Risque de rupture future")
                        + f" · {len(g)} lot(s) : {', '.join(g['lot_id'].astype(str))} · {g['motif_emballage'].iloc[-1]}")
            for _, l in v[v["annonce_id"].isin(nouvelles)].iterrows():
                ajouter(f"INFO-{l['annonce_id']}", "information", "nouvelle_prediction", l["annonce_id"],
                        f"Nouvelle prédiction — {l['lot_id']}",
                        f"Cycle {Formats.duree(l['temps_cycle_predit_min'])} · expédition "
                        f"{pd.Timestamp(l['heure_expedition_predite']):%d/%m %H:%M} · ajoutée au planning")
        # Commandes d'emballage à passer (plan de réapprovisionnement, réceptions saisies incluses)
        if plan is not None and len(plan):
            for _, p in plan.iterrows():
                ajouter(f"COMMANDE-{p['article']}-{p['client_id']}", "attention", "commande", "",
                        f"Commande d'emballage — {p['article']} ({p['client_id']})",
                        f"Rupture prévue le {pd.Timestamp(p['date_rupture']):%d/%m/%Y} · {p['message']} · "
                        f"délai fournisseur {int(p['delai_fournisseur_j'])} j · quantité conseillée "
                        f"{p['quantite_conseillee']:,.0f} {p['unite']}".replace(",", " "))
        # Les informations déjà émises restent (si l'annonce existe toujours)
        if len(anciennes):
            annonces = set(vue["annonce_id"]) if len(vue) else set()
            infos = anciennes[(anciennes["type"] == "nouvelle_prediction") & anciennes["annonce_id"].isin(annonces)]
            deja = {l["notification_id"] for l in lignes}
            for nid, l in infos.iterrows():
                if nid not in deja:
                    lignes.append({"notification_id": nid, **{k: l[k] for k in ("niveau", "type", "annonce_id", "titre", "message")}})
        df = pd.DataFrame(lignes, columns=[c for c in self.COLONNES if c not in ("date_creation", "lu")])
        if len(df):
            if len(anciennes):
                prec = anciennes.reindex(df["notification_id"])
                meme = prec["message"].to_numpy() == df["message"].to_numpy()
                df["lu"] = np.where(meme, prec["lu"].fillna(False).to_numpy(), False)
                df["date_creation"] = np.where(meme & prec["date_creation"].notna().to_numpy(),
                                               prec["date_creation"].to_numpy(), maintenant)
            else:
                df["lu"] = False
                df["date_creation"] = maintenant
        self._ecrire(df)
        return df

    def marquer_lues(self, ids=None) -> None:
        df = self.charger()
        if len(df):
            df.loc[df["notification_id"].isin(ids) if ids is not None else slice(None), "lu"] = True
            self._ecrire(df)

    def non_lues(self) -> int:
        df = self.charger()
        return int((~df["lu"]).sum()) if len(df) else 0


class Formats:
    """Formats d'affichage communs."""

    @staticmethod
    def duree(minutes) -> str:
        """135 → « 02 h 15 »."""
        try:
            m = int(round(float(minutes)))
        except (TypeError, ValueError):
            return "—"
        return f"{m // 60:02d} h {m % 60:02d}"

    @staticmethod
    def kg(x) -> str:
        try:
            return f"{float(x):,.0f} kg".replace(",", " ")
        except (TypeError, ValueError):
            return "—"


# ======================================================================
#  11. CONTEXTE : TOUS LES RÉSULTATS CALCULÉS
# ======================================================================
@dataclass
class ContexteDonnees:
    tables: dict
    ecarts: pd.DataFrame
    rapport_cles: pd.DataFrame
    qualite: ControleQualite
    capacite: MoteurCapacite
    derive: AnalyseDerive
    variables: ConstructeurVariables
    ia: ModelesIA
    emballages: ReferentielEmballages
    empreintes_historiques: dict
    duree_chargement_s: float = 0.0
    version: str = VERSION

    @classmethod
    def construire(cls, exporter: bool = True, classeur: bool = False) -> "ContexteDonnees":
        debut = time.time()
        chargeur = ChargeurDonnees()
        tables = chargeur.charger(exporter=exporter, classeur=classeur)
        qualite = ControleQualite(tables)
        qualite.executer(exporter=exporter)
        capacite = MoteurCapacite(tables)
        derive = AnalyseDerive(tables, qualite.lots_aberrants())
        variables = ConstructeurVariables(tables, capacite, derive, qualite.lots_aberrants())
        ia = ModelesIA(variables)
        ia.charger()
        return cls(tables=tables, ecarts=pd.DataFrame(chargeur.ecarts), rapport_cles=chargeur.rapport_cles,
                   qualite=qualite, capacite=capacite, derive=derive, variables=variables, ia=ia,
                   emballages=ReferentielEmballages(tables, variables),
                   empreintes_historiques=ControlesRecette.empreintes(tables),
                   duree_chargement_s=round(time.time() - debut, 2))


# ======================================================================
#  12. CONTRÔLES DE RECETTE (cahier V2 §21)
# ======================================================================
class ControlesRecette:
    """Tests data (§21.1) et tests fonctionnels (§21.2), partagés par l'app et le mode « valider »."""

    TABLES_HISTORIQUES = ["historique_production", "historique_temps_de_cycle", "annonces_fermes"]
    MOTS_PREDICTIFS = ("predit", "prediction", "niveau_risque", "depassement", "taux_utilisation", "statut_charge")

    @staticmethod
    def empreintes(tables: dict) -> dict:
        """Empreinte de chaque table historique (pour prouver qu'elle n'est pas modifiée)."""
        return {t: (len(tables[t]), int(pd.util.hash_pandas_object(tables[t], index=False).sum()))
                for t in ControlesRecette.TABLES_HISTORIQUES}

    @classmethod
    def tests_data(cls, ctx: ContexteDonnees) -> pd.DataFrame:
        t, q, lignes = ctx.tables, ctx.qualite, []

        def ajouter(nom, ok, detail):
            lignes.append({"contrôle": nom, "résultat": "✓ OK" if ok else "✕ KO", "ok": bool(ok), "détail": detail})

        rc = ctx.rapport_cles
        dbl = int(rc.loc[rc["type"] == "clé primaire", "nb_problèmes"].sum())
        orph = int(rc.loc[rc["type"] == "clé étrangère", "nb_problèmes"].sum())
        ajouter("Unicité des clés primaires", dbl == 0, f"{dbl} doublon(s)")
        ajouter("Validité des clés étrangères", orph == 0, f"{orph} référence(s) inexistante(s)")
        chrono = q.taux_cycles_chronologiques()
        ajouter("Chronologie des timestamps historiques", chrono == 1.0, f"{chrono * 100:.1f} % de cycles valides")
        incoherents = q.anomalies[q.anomalies["regle_violee"].str.startswith("poids_emballe >", na=False)]
        ajouter("Cohérence poids reçu / déchets / poids emballé", incoherents.empty, f"{len(incoherents)} lot(s) incohérent(s)")
        ajouter("Détection de la dérive", ctx.derive.derive_detectee() and ctx.derive.redressement_detecte(),
                f"baisse {ctx.derive.fenetres.set_index('fenêtre').at['Baisse', 'dérive_moyenne_pct']:+.2f} %, "
                f"p = {ctx.derive.tests.iloc[0]['p_value']:.4f}")
        ajouter("Traçabilité des anomalies", Config.FICHIER_ANOMALIES.exists() or len(q.anomalies) > 0,
                f"{len(q.anomalies)} anomalie(s) dans anomalies.csv ({q.taux_anomalies() * 100:.2f} % des lots)")
        colonnes = [f"{tb}.{c}" for tb in cls.TABLES_HISTORIQUES for c in t[tb].columns
                    if any(m in c.lower() for m in cls.MOTS_PREDICTIFS)]
        intactes = cls.empreintes(t) == ctx.empreintes_historiques
        ajouter("Absence de sorties prédictives dans les historiques", not colonnes and intactes,
                "aucune colonne prédictive, empreintes inchangées" if not colonnes and intactes
                else f"colonnes suspectes : {colonnes}" if colonnes else "historique modifié depuis le chargement")
        m = t["mapping_emballages"]
        defauts = m[(m["role"] == "primaire") & m["client_id"].isna() & m["est_defaut"].astype(bool)].groupby("sku_id").size()
        problemes = []
        if set(defauts.index) != set(t["produits_sku"]["sku_id"]) or (defauts != 1).any():
            problemes.append("chaque SKU doit avoir exactement un emballage par défaut")
        inconnus = set(m["article_emballage"]) - set(t["stock_emballages"]["article"])
        if inconnus:
            problemes.append(f"articles absents du stock : {sorted(inconnus)}")
        clients_inconnus = set(m["client_id"].dropna()) - set(t["clients"]["client_id"])
        if clients_inconnus:
            problemes.append(f"clients inconnus : {sorted(clients_inconnus)}")
        ajouter("Validité des mappings SKU / client / emballage", not problemes,
                " ; ".join(problemes) if problemes else f"{len(m)} règles, {len(defauts)} SKU avec défaut")
        ts = ctx.capacite.taux_annonces_surcharge()
        lo, hi = Config.CIBLE_TAUX_SURCHARGE
        ajouter("Surcharge synthétique 10–15 %", lo <= ts <= hi, f"{ts * 100:.2f} % des annonces")
        ta = q.taux_anomalies()
        ajouter("Anomalies ≈ 1 %, toutes traçables", 0.003 <= ta <= 0.02, f"{ta * 100:.2f} % des lots")
        return pd.DataFrame(lignes)

    @classmethod
    def tests_fonctionnels(cls, ctx: ContexteDonnees) -> pd.DataFrame:
        """Rejoue les tests §21.2 dans un dossier temporaire (les vraies prévisions ne sont pas touchées)."""
        import tempfile
        lignes = []

        def ajouter(nom, ok, detail):
            lignes.append({"test": nom, "résultat": "✓ OK" if ok else "✕ KO", "ok": bool(ok), "détail": detail})

        if not ctx.ia.prets:
            ajouter("Modèles entraînés", False, "lancez d'abord le mode « entrainer »")
            return pd.DataFrame(lignes)
        avant = cls.empreintes(ctx.tables)
        with tempfile.TemporaryDirectory() as dossier:
            g = GestionnairePrevisions(ctx, dossier=Path(dossier))
            jour = date.today() + timedelta(days=30)
            # 1. Saisie
            v, e = g.valider(pd.DataFrame([{"date_arrivee": jour, "heure_arrivee": "08:15", "ferme_id": "NG8",
                                            "client_id": "CLI001", "sku_id": "SKU016", "volume_kg": 3075,
                                            "priorite": "HAUTE"}]))
            g.ajouter(v, "MANUEL")
            ok = len(g.annonces()) == 1 and cls.empreintes(ctx.tables) == avant
            ajouter("Saisie d'une nouvelle annonce", ok, "annonce future créée, historique réel inchangé")
            # 2. Import valide
            v, e = g.valider(g.modele_import(), "IMPORT")
            g.ajouter(v, "IMPORT")
            ajouter("Import de fichier valide", e.empty and len(g.annonces()) == 4,
                    f"{len(v)} ligne(s) enregistrée(s), origine IMPORT")
            # 3. Import avec erreur
            mauvais = g.modele_import()
            mauvais.loc[0, "ferme_id"], mauvais.loc[1, "heure_arrivee"], mauvais.loc[2, "volume_kg"] = "NG99", "25:00", -5
            v, e = g.valider(mauvais, "IMPORT")
            ajouter("Import avec erreur", v.empty and len(e) == 3,
                    " | ".join(f"ligne {r.ligne} : {r.message}" for r in e.itertuples()))
            # 4. SKU + client
            p1, p2 = ctx.emballages.proposer("SKU004", "CLI001"), ctx.emballages.proposer("SKU004", "CLI002")
            ajouter("SKU + client → emballage proposé", p1 == "EMBALLAGE_SUR_MESURE" and p2 == "SACHET_IMPORTE_250G",
                    f"SKU004 : Mercadona → {p1}, Lidl → {p2}")
            # 5. Modification de l'emballage
            v, e = g.valider(pd.DataFrame([{"date_arrivee": jour + timedelta(days=3), "heure_arrivee": "09:00",
                                            "ferme_id": "NG1", "client_id": "CLI001", "sku_id": "SKU004",
                                            "volume_kg": 800, "article_emballage": "SACHET_IMPORTE_250G"}]))
            r = g.ajouter(v, "MANUEL").iloc[0]
            ajouter("Modification de l'emballage conservée", bool(r["emballage_modifie"])
                    and r["emballage_selectionne"] == "SACHET_IMPORTE_250G",
                    f"proposé {r['emballage_propose']} → retenu {r['emballage_selectionne']}")
            # 6. Charge > capacité
            jour_s = jour + timedelta(days=5)
            lots = pd.DataFrame([{"date_arrivee": jour_s, "heure_arrivee": f"{7 + i}:00", "ferme_id": "NG2",
                                  "client_id": "CLI002", "sku_id": "SKU001", "volume_kg": 2600} for i in range(4)])
            v, e = g.valider(lots)
            r = g.ajouter(v, "MANUEL")
            notifs = g.notifications.charger()
            sur = notifs[notifs["type"] == "surcharge"]
            ajouter("Charge > capacité", (r["depassement_kg"] > 0).all() and len(sur) >= 1,
                    f"dépassement {Formats.kg(r['depassement_kg'].iloc[0])}, {len(sur)} notification(s) critique(s)")
            # 7. Charge proche de la capacité
            jour_t = jour + timedelta(days=8)
            v, e = g.valider(pd.DataFrame([{"date_arrivee": jour_t, "heure_arrivee": "07:00", "ferme_id": "NG4",
                                            "client_id": "CLI004", "sku_id": "SKU001", "volume_kg": 6700,
                                            "nb_operateurs": 10}]))
            r = g.ajouter(v, "MANUEL").iloc[0]
            ajouter("Charge proche de la capacité", r["etat_charge"] == "tension",
                    f"charge {r['taux_utilisation_pct']:.0f} % → état « {r['etat_charge']} »")
            # 8. Prédiction réussie
            vue = g.vue()
            ajouter("Prédiction réussie", vue["temps_cycle_predit_min"].gt(0).all() and vue["heure_expedition_predite"].notna().all(),
                    f"{len(vue)} prédiction(s) avec cycle et expédition")
            # 8 bis. Cohérence rendement ↔ temps de cycle
            prod = vue["poids_emballe_prevu_kg"] / vue["rendement_prevu_kg_h"] * 60
            somme = vue[["temps_attente_min", "temps_hors_production_min", "temps_production_min",
                         "temps_palettisation_min", "temps_preparation_expedition_min"]].sum(axis=1)
            ok = (abs(prod - vue["temps_production_min"]) <= 0.2).all() and (abs(somme - vue["temps_cycle_predit_min"]) <= 0.5).all()
            l0 = vue.iloc[0]
            ajouter("Cohérence rendement ↔ temps de cycle", ok,
                    f"ex. {l0['poids_emballe_prevu_kg']:.0f} kg ÷ {l0['rendement_prevu_kg_h']:.0f} kg/h = "
                    f"{Formats.duree(l0['temps_production_min'])} de production dans un cycle de "
                    f"{Formats.duree(l0['temps_cycle_predit_min'])} (somme des 5 étapes)")
            # 8 ter. Réception d'emballage et prochain batch
            res, _ = g.plan_reappro()
            critiques = res[res["statut"] == "critique"]
            if len(critiques):
                l1 = critiques.iloc[0]
                g.enregistrer_reception(l1["article"], l1["client_id"], l1["quantite_conseillee"] + 1000, date.today())
                notifs = g.notifications.charger()
                reste = notifs["notification_id"].eq(f"COMMANDE-{l1['article']}-{l1['client_id']}").any()
                ajouter("Réception d'emballage → alerte levée", not reste,
                        f"{l1['article']} ({l1['client_id']}) : alerte de commande supprimée après la réception")
            b = g.prochain_batch(jour_s, "jour")
            ajouter("Prochain batch conseillé", b["arrivee_conseillee"] is not None and (b["sature"] or b["quantite_max_kg"] > 0),
                    f"shift du {jour_s:%d/%m} : " + ("complet → " + (f"{b['suivant']['shift']} {b['suivant']['arrivee_conseillee']:%d/%m %H:%M}, "
                                                                       f"{b['suivant']['quantite_max_kg']:,.0f} kg" if b["suivant"] else "aucun")
                                                      if b["sature"] else f"arrivée {b['arrivee_conseillee']:%H:%M}, "
                                                                          f"≤ {b['quantite_max_kg']:,.0f} kg").replace(",", " "))
            # 9. Planning
            l = vue.iloc[0]
            ok = pd.Timestamp(l["heure_arrivee_prevue"]) == pd.Timestamp(jour) + pd.Timedelta(hours=8, minutes=15)
            ajouter("Planning : lot à la bonne date/heure", ok, f"{l['lot_id']} prévu le {pd.Timestamp(l['heure_arrivee_prevue']):%d/%m/%Y %H:%M}")
            # 10. Cloche
            types = set(g.notifications.charger()["type"])
            ajouter("Cloche : risques visibles", {"surcharge", "nouvelle_prediction"} <= types,
                    "types présents : " + ", ".join(sorted(types)))
            # 11. Historique
            ajouter("Historique : aucune prédiction présentée comme réelle", cls.empreintes(ctx.tables) == avant
                    and not set(vue["lot_id"]) & set(ctx.tables["historique_production"]["lot_id"]),
                    "historique_production et historique_temps_de_cycle inchangés")
        return pd.DataFrame(lignes)


# ======================================================================
#  13. INTERFACE : CSS, COMPOSANTS, FILTRES
# ======================================================================
@dataclass
class Filtres:
    """Filtres globaux persistants (st.session_state), appliqués aux annonces FUTURES et à leurs prédictions."""
    debut: pd.Timestamp
    fin: pd.Timestamp
    fermes: list = field(default_factory=list)
    clients: list = field(default_factory=list)
    familles: list = field(default_factory=list)
    skus: list = field(default_factory=list)
    shifts: list = field(default_factory=list)

    def previsions(self, v: pd.DataFrame, avec_periode: bool = True) -> pd.DataFrame:
        if v.empty:
            return v
        if avec_periode:
            jours = pd.to_datetime(v["date_arrivee"])
            v = v[(jours >= self.debut) & (jours <= self.fin)]
        for col, valeurs in (("ferme_id", self.fermes), ("client_id", self.clients), ("famille", self.familles),
                             ("sku_id", self.skus), ("shift", self.shifts)):
            if valeurs and col in v.columns:
                v = v[v[col].isin(valeurs)]
        return v

    def stocks(self, df):
        return df[df["client_id"].isin(self.clients)] if self.clients else df


class InterfaceUI:
    """Direction artistique : CSS injecté, cartes KPI, badges, panneau droit sombre vide."""

    C = Config.COULEURS
    STATUTS = {
        "normal": ("Normal", "✓", C["normal"]),
        "tension": ("Tension", "⚠", C["tension"]),
        "surcharge": ("Surcharge", "▲", C["critique"]),
        "forte_surcharge": ("Forte surcharge", "✕", C["critique"]),
        "critique": ("Critique", "✕", C["critique"]),
        "neutre": ("Neutre", "○", C["neutre"]),
        "avertissement": ("Avertissement", "⚠", "#E65100"),
        "attention": ("Attention", "◆", "#B8860B"),
        "information": ("Information", "ℹ", "#1565C0"),
    }
    MENU = [   # navigation : uniquement l'espace prédictif (la BDD sert à l'apprentissage)
        ("Dashboard", "◫", "page_dashboard"),
        ("Flux entrants", "⇄", "page_flux"),
        ("Planning prédictif", "▦", "page_planning"),
        ("Capacité & shifts", "◔", "page_capacite"),
        ("Emballages", "▣", "page_emballages"),
        ("Paramètres", "⚙", "page_parametres"),
    ]
    ICONES = {   # icônes linéaires (style « feather »)
        "camion": '<path d="M1 3h15v13H1z"/><path d="M16 8h4l3 3v5h-7z"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/>',
        "jauge": '<path d="M12 14l4-4"/><path d="M3.3 19a10 10 0 1 1 17.4 0"/>',
        "pourcent": '<line x1="19" y1="5" x2="5" y2="19"/><circle cx="6.5" cy="6.5" r="2.5"/><circle cx="17.5" cy="17.5" r="2.5"/>',
        "colis": '<path d="M21 16V8a2 2 0 0 0-1-1.7l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.7l7 4a2 2 0 0 0 2 0l7-4a2 2 0 0 0 1-1.7z"/><path d="M3.3 7L12 12l8.7-5"/><path d="M12 22V12"/>',
        "alerte": '<path d="M10.3 3.9L1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12" y2="17"/>',
        "horloge": '<circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/>',
        "activite": '<path d="M22 12h-4l-3 9L9 3l-3 9H2"/>',
        "valide": '<path d="M22 11.1V12a10 10 0 1 1-5.9-9.1"/><path d="M22 4L12 14l-3-3"/>',
        "cloche": '<path d="M18 8a6 6 0 0 0-12 0c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/>',
        "calendrier": '<rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>',
    }
    TOUS = "Tous"

    # --- CSS ----------------------------------------------------------------
    def injecter_css(self) -> None:
        c = self.C
        st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
html, body, .stApp, button, input, textarea, select, p, label, li, h1, h2, h3, h4,
[data-testid="stMarkdownContainer"], [data-testid="stWidgetLabel"] {{ font-family: 'Poppins', 'Inter', sans-serif; }}
/* Icônes Streamlit (flèches des menus dépliants, etc.) : garder leur police d'icônes,
   sinon leur nom s'affiche en texte (« arrow_right », « keyboard_arrow_down »…) */
[data-testid="stIconMaterial"], span[data-testid="stIconMaterial"], .material-symbols-rounded,
[class*="material-symbols"], [class*="material-icons"] {{
    font-family: 'Material Symbols Rounded', 'Material Icons' !important; font-weight: 400 !important;
    font-style: normal !important; letter-spacing: normal !important; text-transform: none !important; }}
/* ---- Lisibilité : texte toujours foncé sur fond clair (même si le navigateur est en mode sombre) ---- */
.stApp {{ background: {c['fond']}; color: {c['texte']}; }}
.stApp p, .stApp li, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp small,
[data-testid="stMarkdownContainer"], [data-testid="stCaptionContainer"], [data-testid="stWidgetLabel"],
[data-testid="stMetricLabel"], [data-testid="stMetricValue"], [data-testid="stMetricDelta"],
.stTabs [data-baseweb="tab"], [data-testid="stExpander"] summary {{ color: {c['texte']}; }}
[data-testid="stCaptionContainer"] {{ color: {c['texte2']}; }}
input, textarea, [data-baseweb="input"] > div, [data-baseweb="select"] > div, [data-baseweb="textarea"],
[data-testid="stDateInput"] input, [data-testid="stTimeInput"] input, [data-testid="stNumberInput"] input {{
    background: #FFFFFF !important; color: {c['texte']} !important; -webkit-text-fill-color: {c['texte']}; }}
[data-baseweb="select"] span, [data-baseweb="tag"] span {{ color: {c['texte']}; }}
[data-baseweb="popover"] ul, [data-baseweb="popover"] li, [data-baseweb="menu"], [role="listbox"], [role="option"],
[data-baseweb="calendar"], [data-baseweb="calendar"] * {{ background-color: #FFFFFF; color: {c['texte']}; }}
[role="option"]:hover, [aria-selected="true"][role="option"] {{ background-color: #EEF6EE !important; }}
.stButton button, .stDownloadButton button, .stFormSubmitButton button,
[data-testid="stBaseButton-secondary"], [data-testid="stBaseButton-secondaryFormSubmit"] {{
    background: #FFFFFF; color: {c['texte']}; border: 1px solid #D9E3D6; border-radius: 12px; }}
.stButton button:hover, .stDownloadButton button:hover {{ border-color: {c['vert']}; color: {c['vert']}; }}
.stButton button[kind="primary"], [data-testid="stBaseButton-primary"], [data-testid="stBaseButton-primaryFormSubmit"] {{
    background: {c['vert']}; color: #FFFFFF; border: 1px solid {c['vert']}; }}
.stButton button[kind="primary"] p, [data-testid="stBaseButton-primary"] p {{ color: #FFFFFF; }}
[data-testid="stExpander"] details, [data-testid="stFileUploaderDropzone"] {{ background: #FFFFFF; border-radius: 14px; }}
[data-testid="stFileUploaderDropzone"] * {{ color: {c['texte']}; }}
[data-testid="stAlert"] * {{ color: {c['texte']}; }}
header[data-testid="stHeader"] {{ background: transparent; }}
.block-container {{ padding-top: 1.6rem; padding-right: 2.2rem !important; padding-bottom: 3.5rem; max-width: 100% !important; }}
/* ---- Sidebar ---- */
section[data-testid="stSidebar"] {{ background: #FFFFFF; border-right: 1px solid #EEEAE0; min-width: 250px !important; max-width: 250px !important; }}
section[data-testid="stSidebar"] div[role="radiogroup"] label {{ padding: 8px 12px; border-radius: 12px; margin-bottom: 2px; width: 100%; }}
section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {{ background: #EEF6EE; }}
section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) {{ background: {c['vert']}; }}
section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) p {{ color: #FFFFFF !important; font-weight: 600; }}
section[data-testid="stSidebar"] div[role="radiogroup"] label > div:first-child {{ display: none; }}
/* ---- Cartes ---- */
div[data-testid="stVerticalBlockBorderWrapper"] {{ background: {c['carte']}; border-radius: 18px; border: none !important;
    box-shadow: 0 2px 14px rgba(30, 43, 34, .06); }}
.titre-page {{ font-size: 1.9rem; font-weight: 700; color: {c['texte']}; margin: 0; line-height: 1.2; }}
.sous-titre {{ color: {c['texte2']}; font-size: .95rem; margin: .2rem 0 1rem 0; }}
.kpi {{ border-radius: 18px; padding: 18px 20px; min-height: 132px; box-shadow: 0 2px 14px rgba(30,43,34,.07); }}
.kpi.vert {{ background: linear-gradient(135deg, #1F6B25 0%, {c['vert']} 100%); color: #FFFFFF; }}
.kpi.vert .entete, .kpi.vert .valeur, .kpi.vert .unite, .kpi.vert .detail {{ color: #FFFFFF; }}
.kpi.blanc {{ background: #FFFFFF; color: {c['texte']}; }}
.kpi.blanc .entete, .kpi.blanc .valeur, .kpi.blanc .unite {{ color: {c['texte']}; }}
.kpi.blanc .detail {{ color: {c['texte2']}; }}
.kpi .entete {{ display: flex; justify-content: space-between; align-items: center; font-size: .86rem; }}
.kpi .valeur {{ font-size: 1.75rem; font-weight: 700; margin-top: 8px; line-height: 1.1; }}
.kpi .unite {{ font-size: .95rem; font-weight: 500; margin-left: 4px; }}
.kpi .detail {{ font-size: .78rem; margin-top: 6px; }}
.kpi svg {{ width: 22px; height: 22px; fill: none; stroke: currentColor; stroke-width: 1.8; stroke-linecap: round; stroke-linejoin: round; }}
.badge {{ display: inline-flex; align-items: center; gap: 6px; padding: 3px 11px; border-radius: 999px;
    font-size: .78rem; font-weight: 600; border: 1px solid; }}
.logo {{ display: flex; align-items: center; gap: 10px; padding: 4px 4px 16px 4px; }}
.logo .rond {{ width: 38px; height: 38px; border-radius: 12px; background: {c['vert']}; color: #FFFFFF;
    display: flex; align-items: center; justify-content: center; font-weight: 700; }}
.logo .nom {{ font-weight: 700; font-size: 1.02rem; color: {c['texte']}; line-height: 1.1; }}
.logo .soc {{ font-size: .74rem; color: {c['texte2']}; }}
/* ---- Calendrier ---- */
.cal-entete {{ text-align: center; font-size: .78rem; font-weight: 600; color: {c['texte2']}; padding: 4px 0; }}
.cal-mois {{ text-align: center; font-weight: 700; font-size: 1.05rem; color: {c['texte']}; padding-top: 6px; }}
/* ---- Cloche + panneau vert coulissant des alertes ---- */
details.centre-alertes > summary {{ list-style: none; position: fixed; top: 64px; right: 26px; z-index: 1000002;
    width: 48px; height: 48px; border-radius: 50%; background: {c['vert']}; color: #FFFFFF; cursor: pointer;
    display: flex; align-items: center; justify-content: center; box-shadow: 0 4px 16px rgba(15,42,29,.30); }}
details.centre-alertes > summary::-webkit-details-marker {{ display: none; }}
details.centre-alertes > summary svg {{ width: 22px; height: 22px; fill: none; stroke: #FFFFFF; stroke-width: 2;
    stroke-linecap: round; stroke-linejoin: round; }}
details.centre-alertes[open] > summary {{ background: #FFFFFF; }}
details.centre-alertes[open] > summary svg {{ stroke: {c['panneau']}; }}
.pastille {{ position: absolute; top: -4px; right: -4px; min-width: 20px; height: 20px; padding: 0 5px; border-radius: 10px;
    background: {c['critique']}; color: #FFFFFF; font-size: .72rem; font-weight: 700; display: flex;
    align-items: center; justify-content: center; border: 2px solid #FFFFFF; }}
.panneau-alertes {{ position: fixed; top: 0; right: 0; width: 340px; max-width: 92vw; height: 100vh; overflow-y: auto;
    background: {c['panneau']}; color: #FFFFFF; z-index: 1000001; padding: 128px 20px 30px 20px;
    border-top-left-radius: 26px; border-bottom-left-radius: 26px; box-shadow: -8px 0 30px rgba(0,0,0,.25);
    animation: glisser .28s ease-out; }}
@keyframes glisser {{ from {{ transform: translateX(105%); }} to {{ transform: translateX(0); }} }}
.panneau-alertes .pa-titre {{ font-size: 1.15rem; font-weight: 700; color: #FFFFFF; margin-bottom: 2px; }}
.panneau-alertes .pa-sous {{ font-size: .78rem; color: #B9D3C2; margin-bottom: 16px; }}
.alerte {{ display: flex; gap: 10px; background: rgba(255,255,255,.07); border-radius: 14px; padding: 11px 12px;
    margin-bottom: 9px; border-left: 4px solid; }}
.alerte .ic {{ font-weight: 700; font-size: 1rem; line-height: 1.3; }}
.alerte .t {{ font-weight: 600; font-size: .86rem; color: #FFFFFF; }}
.alerte .d {{ font-size: .76rem; color: #CFE3D6; margin-top: 2px; }}
.alerte.critique {{ border-color: #FF6B6B; }} .alerte.critique .ic {{ color: #FF8A80; }}
.alerte.tension {{ border-color: #FFC94D; }} .alerte.tension .ic {{ color: #FFD54F; }}
.alerte.neutre {{ border-color: #9FB8A8; }} .alerte.neutre .ic {{ color: #C8D8CD; }}
.alerte.normal {{ border-color: #81C784; }} .alerte.normal .ic {{ color: #A5D6A7; }}
.pa-pied {{ font-size: .72rem; color: #9FB8A8; margin-top: 14px; text-align: center; }}
/* ---- Cloche 🔔 et panneau de notifications rétractable (conteneurs Streamlit à clé) ---- */
.st-key-zone_cloche {{ position: fixed; top: 62px; right: 22px; z-index: 1000003; width: auto !important; }}
.st-key-zone_cloche button {{ border-radius: 999px !important; background: {c['vert']} !important; border: none !important;
    min-height: 46px; padding: 0 18px !important; box-shadow: 0 4px 16px rgba(15,42,29,.30); }}
.st-key-zone_cloche button p {{ color: #FFFFFF !important; font-weight: 700; font-size: 1.05rem; }}
.st-key-panneau_notifications {{ position: fixed; top: 0; right: 0; width: 370px !important; max-width: 94vw; height: 100vh;
    overflow-y: auto; background: {c['panneau']}; z-index: 1000002; padding: 120px 18px 30px 18px;
    border-top-left-radius: 26px; border-bottom-left-radius: 26px; box-shadow: -8px 0 30px rgba(0,0,0,.25);
    animation: glisser .28s ease-out; }}
.st-key-panneau_notifications p, .st-key-panneau_notifications span {{ color: #FFFFFF; }}
.st-key-panneau_notifications button {{ background: rgba(255,255,255,.10) !important; border: 1px solid rgba(255,255,255,.28) !important; }}
.st-key-panneau_notifications button p {{ color: #FFFFFF !important; }}
.alerte.avertissement {{ border-color: #FF9E40; }} .alerte.avertissement .ic {{ color: #FFB74D; }}
.alerte.attention {{ border-color: #FFE066; }} .alerte.attention .ic {{ color: #FFE066; }}
.alerte.information {{ border-color: #82B1FF; }} .alerte.information .ic {{ color: #82B1FF; }}
.alerte .h {{ font-size: .68rem; color: #9FB8A8; margin-top: 3px; }}
.alerte .nouveau {{ background: #FF6B6B; color: #FFFFFF; border-radius: 6px; padding: 0 5px; font-size: .64rem; margin-left: 6px; }}
/* ---- Liens « Détail → » sous les KPI ---- */
[class*="st-key-lien_"] button {{ background: transparent !important; border: none !important; box-shadow: none !important;
    padding: 0 4px !important; min-height: 1.6rem; }}
[class*="st-key-lien_"] button p {{ color: {c['vert']} !important; font-weight: 600; font-size: .85rem; }}
.pied {{ position: fixed; bottom: 0; left: 0; right: 0; text-align: center; font-size: .76rem;
    color: {c['texte2']}; background: rgba(250, 248, 242, .94); padding: 6px; z-index: 80; }}
</style>""", unsafe_allow_html=True)

    @staticmethod
    def _container_cle_ok() -> bool:
        """Vrai si st.container accepte key= (Streamlit récent) : nécessaire pour le panneau fixe."""
        import inspect
        try:
            return "key" in inspect.signature(st.container).parameters
        except (TypeError, ValueError):
            return False

    def panneau_notifications(self, notif: "GestionnaireNotifications") -> None:
        """Cloche 🔔 (en haut à droite) qui ouvre/ferme le panneau sombre des notifications."""
        import html as html_mod
        ss = st.session_state
        ss.setdefault("panneau_ouvert", False)
        df = notif.charger()
        non_lues = int((~df["lu"]).sum()) if len(df) else 0

        def basculer():
            ss["panneau_ouvert"] = not ss.get("panneau_ouvert", False)

        def fermer():
            ss["panneau_ouvert"] = False

        cle_ok = self._container_cle_ok()
        zone = st.container(key="zone_cloche") if cle_ok else st.sidebar
        with zone:
            st.button(f"🔔 {non_lues}" if non_lues else "🔔", key="btn_cloche", on_click=basculer,
                      help="Notifications : surcharges, retards, tensions d'emballage, nouvelles prédictions")
        if not ss["panneau_ouvert"]:
            return
        panneau = st.container(key="panneau_notifications") if cle_ok else st.sidebar.container()
        with panneau:
            st.markdown(f'<div class="pa-titre">🔔 Notifications</div>'
                        f'<div class="pa-sous">{non_lues} non lue(s) · {len(df)} au total</div>', unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            c1.button("✓ Tout lire", key="notif_lire", on_click=notif.marquer_lues)
            c2.button("✕ Fermer", key="notif_fermer", on_click=fermer)
            if df.empty:
                cartes = ('<div class="alerte normal"><div class="ic">✓</div><div><div class="t">Aucune notification</div>'
                          '<div class="d">Les risques des prédictions apparaîtront ici.</div></div></div>')
            else:
                cartes = ""
                for _, n in df.head(60).iterrows():
                    emoji, libelle = GestionnaireNotifications.NIVEAUX.get(n["niveau"], ("•", n["niveau"]))
                    nouveau = '<span class="nouveau">nouveau</span>' if not n["lu"] else ""
                    quand = pd.Timestamp(n["date_creation"]).strftime("%d/%m %H:%M") if pd.notna(n["date_creation"]) else ""
                    cartes += (f'<div class="alerte {n["niveau"]}"><div class="ic">{emoji}</div><div>'
                               f'<div class="t">{libelle} — {html_mod.escape(str(n["titre"]))}{nouveau}</div>'
                               f'<div class="d">{html_mod.escape(str(n["message"]))}</div>'
                               f'<div class="h">{quand}</div></div></div>')
            st.markdown(cartes + '<div class="pa-pied">Cliquez sur la cloche 🔔 pour refermer le panneau</div>',
                        unsafe_allow_html=True)

    def pied_de_page(self) -> None:
        st.markdown(f'<div class="pied">🌿 {Config.PIED_DE_PAGE} · {Config.TITRE_APP}</div>',
                    unsafe_allow_html=True)

    # --- Composants ---------------------------------------------------------
    def icone(self, nom: str) -> str:
        return f'<svg viewBox="0 0 24 24">{self.ICONES.get(nom, self.ICONES["activite"])}</svg>'

    def titre(self, titre: str, sous_titre: str) -> None:
        st.markdown(f'<p class="titre-page">{titre}</p><p class="sous-titre">{sous_titre}</p>',
                    unsafe_allow_html=True)

    def carte_kpi(self, titre, valeur, unite, detail="", icone="activite", style="vert", statut=None) -> None:
        badge = f'<div style="margin-top:6px">{self.badge(statut, sur_fond_vert=style == "vert")}</div>' if statut else ""
        st.markdown(f"""<div class="kpi {style}">
  <div class="entete"><span>{titre}</span>{self.icone(icone)}</div>
  <div class="valeur">{valeur}<span class="unite">{unite}</span></div>
  <div class="detail">{detail}</div>{badge}</div>""", unsafe_allow_html=True)

    def badge(self, statut: str, sur_fond_vert: bool = False) -> str:
        libelle, symbole, couleur = self.STATUTS.get(statut, self.STATUTS["neutre"])
        fond = "rgba(255,255,255,.95)" if sur_fond_vert else f"{couleur}14"
        return (f'<span class="badge" style="color:{couleur};border-color:{couleur};background:{fond}">'
                f'{symbole} {libelle}</span>')

    def afficher_badge(self, statut: str, texte: str = "") -> None:
        st.markdown(f"{self.badge(statut)} <span style='margin-left:8px'>{texte}</span>", unsafe_allow_html=True)

    def libelle(self, statut: str) -> str:
        """Version texte du badge pour les tableaux (icône + libellé, jamais la couleur seule)."""
        libelle, symbole, _ = self.STATUTS.get(statut, self.STATUTS["neutre"])
        return f"{symbole} {libelle}"

    @staticmethod
    def nombre(x, decimales=0) -> str:
        try:
            if x is None or (isinstance(x, float) and math.isnan(x)):
                return "—"
            return f"{x:,.{decimales}f}".replace(",", " ").replace(".", ",")
        except (TypeError, ValueError):
            return str(x)

    def style_figure(self, fig, hauteur=360, titre=None):
        fig.update_layout(
            height=hauteur, margin=dict(l=10, r=10, t=50 if titre else 20, b=10),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Poppins, Inter, sans-serif", color=self.C["texte"], size=12),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            colorway=Config.PALETTE_GRAPHIQUES, title=titre)
        fig.update_xaxes(showgrid=False)
        fig.update_yaxes(gridcolor="#EFEBE0")
        return fig

    def graphique(self, fig, hauteur=360, titre=None) -> None:
        self.style_figure(fig, hauteur, titre)
        try:
            st.plotly_chart(fig, width="stretch", config={"displaylogo": False})
        except Exception:
            st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})

    def graphique_selection(self, fig, hauteur=360, cle="selection", titre=None) -> list:
        """Graphique cliquable : renvoie les valeurs x des points/barres sélectionnés."""
        self.style_figure(fig, hauteur, titre)
        fig.update_layout(clickmode="event+select")
        try:
            evenement = st.plotly_chart(fig, key=cle, on_select="rerun", selection_mode="points",
                                        config={"displaylogo": False},
                                        **({"width": "stretch"} if self._largeur_ok() else {"use_container_width": True}))
            points = evenement.selection.points if evenement else []
            return [p.get("x") for p in points if p.get("x") is not None]
        except Exception:
            self.graphique(fig, hauteur, titre)
            return []

    @staticmethod
    def tableau(df, hauteur=None, **options) -> None:
        params = dict(hide_index=True, **options)
        if hauteur:
            params["height"] = hauteur
        try:
            st.dataframe(df, width="stretch", **params)
        except Exception:
            st.dataframe(df, use_container_width=True, **params)

    def editeur(self, df, **options):
        if self._largeur_ok():
            return st.data_editor(df, width="stretch", hide_index=True, **options)
        return st.data_editor(df, use_container_width=True, hide_index=True, **options)

    @staticmethod
    def vide(df) -> bool:
        if df is None or len(df) == 0:
            st.info("Aucune donnée pour ces filtres.")
            return True
        return False

    # --- Sidebar : logo, menu, filtres ------------------------------------------
    def menu(self) -> str:
        if "page_cible" in st.session_state:
            st.session_state["menu"] = st.session_state.pop("page_cible")
        with st.sidebar:
            st.markdown('<div class="logo"><div class="rond">NG</div><div><div class="nom">PackFlow Predict</div>'
                        '<div class="soc">Nature Growers · Agadir</div></div></div>', unsafe_allow_html=True)
            noms = [m[0] for m in self.MENU]
            icones = {m[0]: m[1] for m in self.MENU}
            if st.session_state.get("menu") not in noms:
                st.session_state["menu"] = noms[0]
            return st.radio("Menu", noms, key="menu", format_func=lambda n: f"{icones[n]}  {n}",
                            label_visibility="collapsed")

    def filtres(self, ctx: "ContexteDonnees") -> Filtres:
        """Filtres globaux (listes déroulantes) appliqués aux prévisions, persistants entre les pages."""
        t = ctx.tables
        ss = st.session_state
        tous = self.TOUS
        ss.setdefault("f_periode_prev", (date.today(), date.today() + timedelta(days=60)))
        for cle in ("f_ferme", "f_client", "f_famille", "f_sku", "f_shift"):
            ss.setdefault(cle, tous)
        noms_clients = t["clients"].set_index("client_id")["nom"].to_dict()
        produits = t["produits_sku"].set_index("sku_id")
        with st.sidebar:
            st.markdown("---")
            with st.expander("Filtres globaux", expanded=False):
                st.date_input("Période (arrivées prévues)", key="f_periode_prev", format="DD/MM/YYYY")
                st.selectbox("Ferme", [tous] + t["fermes"]["ferme_id"].tolist(), key="f_ferme")
                st.selectbox("Client", [tous] + list(noms_clients), key="f_client",
                             format_func=lambda c: c if c == tous else f"{c} · {noms_clients.get(c, '')}")
                st.selectbox("Famille", [tous] + Config.FAMILLES, key="f_famille")
                st.selectbox("SKU", [tous] + produits.index.tolist(), key="f_sku",
                             format_func=lambda s: s if s == tous else f"{s} · {produits.at[s, 'produit']} ({produits.at[s, 'format']})")
                st.selectbox("Shift", [tous] + Config.SHIFTS, key="f_shift")
                if self.bouton("Réinitialiser les filtres"):
                    for cle in ("f_ferme", "f_client", "f_famille", "f_sku", "f_shift"):
                        ss[cle] = tous
                    ss["f_periode_prev"] = (date.today(), date.today() + timedelta(days=60))
                    st.rerun()
            st.caption("La base historique sert uniquement à l'apprentissage du modèle prédictif : "
                       "les écrans n'affichent que les annonces futures et leurs prédictions.")
            st.caption(f"Version {VERSION}")
        periode = ss["f_periode_prev"]
        if isinstance(periode, (list, tuple)):
            debut = periode[0] if len(periode) >= 1 else date.today()
            fin = periode[1] if len(periode) >= 2 else debut
        else:
            debut = fin = periode
        liste = lambda cle: [] if ss[cle] == tous else [ss[cle]]
        return Filtres(pd.Timestamp(debut), pd.Timestamp(fin), liste("f_ferme"), liste("f_client"),
                       liste("f_famille"), liste("f_sku"), liste("f_shift"))

    @staticmethod
    def _largeur_ok() -> bool:
        """Vrai si la version de Streamlit accepte width="stretch" sur les boutons."""
        try:
            version = tuple(int(x) for x in st.__version__.split(".")[:2])
            return version >= (1, 46)
        except Exception:
            return False

    def bouton(self, libelle, **options) -> bool:
        if self._largeur_ok():
            options.setdefault("width", "stretch")
        return st.button(libelle, **options)


# ======================================================================
#  14. PAGES DE L'APPLICATION
# ======================================================================
class PagesApplication:
    """Les écrans de PackFlow Predict."""

    def __init__(self, ctx: ContexteDonnees, ui: InterfaceUI, filtres: Filtres):
        self.ctx, self.ui, self.f = ctx, ui, filtres
        t = ctx.tables
        self.noms_clients = t["clients"].set_index("client_id")["nom"].to_dict()
        self.sku = t["produits_sku"].set_index("sku_id")
        self.prev = GestionnairePrevisions(ctx)
        self.distances = t["fermes"].set_index("ferme_id")["distance_usine_km"].to_dict()

    # --- Libellés des listes déroulantes ---------------------------------------
    def lib_sku(self, s):
        return f"{s} · {self.sku.at[s, 'produit']} ({self.sku.at[s, 'format']})" if s in self.sku.index else str(s)

    def lib_client(self, c):
        return f"{c} · {self.noms_clients.get(c, '')}"

    def lib_ferme(self, f):
        return f"{f} · {self.distances.get(f, 0):g} km"

    # --- Données de prédiction communes -----------------------------------------
    def vue_futures(self, avec_archivees=False, filtrer=False, avec_periode=True) -> pd.DataFrame:
        """Annonces futures + prédictions (hors archivées par défaut), filtres globaux optionnels."""
        v = self.prev.vue()
        if len(v):
            v = v.assign(famille=v["sku_id"].map(self.sku["famille"]))
            if not avec_archivees:
                v = v[v["statut_prediction"] != "ARCHIVEE"]
            if filtrer:
                v = self.f.previsions(v, avec_periode)
        return v

    @staticmethod
    def shifts_futurs(v: pd.DataFrame) -> pd.DataFrame:
        """Une ligne par shift prévu : charge, capacité, taux, dépassement, état."""
        if v.empty or v["date_shift"].isna().all():
            return pd.DataFrame(columns=["date_shift", "shift", "lots", "volume_lots", "charge", "capacite",
                                         "taux", "depassement", "etat"])
        return (v.groupby(["date_shift", "shift"])
                .agg(lots=("annonce_id", "count"), volume_lots=("volume_kg", "sum"),
                     charge=("charge_annoncee_kg", "first"), capacite=("capacite_disponible_kg", "first"),
                     taux=("taux_utilisation_pct", "first"), depassement=("depassement_kg", "first"),
                     etat=("etat_charge", "first"))
                .reset_index().sort_values(["date_shift", "shift"]))

    def lien(self, libelle, page=None, cle=None, ouvrir_panneau=False, **etat):
        """Petit bouton « Détail → » qui ouvre une page (ou le panneau des notifications)."""
        def aller():
            for k, val in etat.items():
                st.session_state[k] = val
            if ouvrir_panneau:
                st.session_state["panneau_ouvert"] = True
            if page:
                st.session_state["page_cible"] = page
        st.button(libelle, key=cle or f"lien_{libelle}_{page}", on_click=aller)

    def libelle_risque(self, niveau) -> str:
        return {"critique": "✕ Critique", "tension": "⚠ Tension", "normal": "✓ Normal"}.get(niveau, "—")

    def _aucune_prevision(self, cle="vide"):
        st.info("Aucune annonce future pour ces filtres : créez ou importez des annonces dans « Flux entrants » "
                "(ou élargissez la période dans les filtres globaux).")
        self.lien("✎ Créer une annonce →", "Flux entrants", f"lien_creer_{cle}")

    # ================================================================
    #  DASHBOARD (annonces futures et prédictions uniquement)
    # ================================================================
    def page_dashboard(self):
        ui = self.ui
        ui.titre("Dashboard", "Vue exécutive des annonces futures : charge, capacité, prédictions et risques")
        c1, c2, c3, c4 = st.columns([3, 1.3, 1.3, 1])
        recherche = c1.text_input("Recherche", placeholder="Rechercher un lot, une ferme, un SKU, un client…",
                                  label_visibility="collapsed")
        vue = c2.selectbox("Vue", ["Jour", "Semaine"], label_visibility="collapsed")
        shift_vu = c3.selectbox("Shift mis en évidence", ["Tous les shifts", "jour", "nuit"],
                                label_visibility="collapsed")
        if c4.button("↻ Rafraîchir"):
            st.cache_resource.clear()
            st.rerun()
        v = self.vue_futures(filtrer=True)
        v = v[v["heure_arrivee_prevue"].notna()] if len(v) else v
        sh = self.shifts_futurs(v)
        notifs = self.prev.notifications.charger()
        risques = int(notifs["niveau"].isin(["critique", "avertissement", "attention"]).sum()) if len(notifs) else 0
        volume = float(v["volume_kg"].sum()) if len(v) else 0.0
        cap_totale = float(sh["capacite"].sum()) if len(sh) else 0.0
        charge_totale = float(sh["charge"].sum()) if len(sh) else 0.0
        taux = charge_totale / cap_totale if cap_totale else 0.0
        nb_sur = int((sh["depassement"] > 0).sum()) if len(sh) else 0

        k = st.columns(5)
        with k[0]:
            ui.carte_kpi("Volume annoncé", ui.nombre(volume), "kg", f"{len(v)} annonce(s) future(s)", "camion")
            self.lien("Détail →", "Flux entrants", "lien_kpi_volume")
        with k[1]:
            ui.carte_kpi("Capacité disponible du shift", ui.nombre(cap_totale / len(sh) if len(sh) else 0), "kg",
                         f"moyenne sur {len(sh)} shift(s) prévu(s)", "jauge")
            self.lien("Détail →", "Capacité & shifts", "lien_kpi_capacite")
        with k[2]:
            ui.carte_kpi("Taux d'utilisation", ui.nombre(taux * 100, 1), "%", f"{nb_sur} shift(s) en surcharge",
                         "pourcent", statut=MoteurCapacite.statut_charge(taux) if len(sh) else "neutre")
            self.lien("Détail →", "Capacité & shifts", "lien_kpi_taux")
        with k[3]:
            ui.carte_kpi("Lots futurs", ui.nombre(len(v)), "lots",
                         f"{int((v['statut_prediction'] == 'VALIDEE').sum()) if len(v) else 0} validé(s)", "colis")
            self.lien("Détail →", "Planning prédictif", "lien_kpi_lots")
        with k[4]:
            ui.carte_kpi("Risques détectés", ui.nombre(risques), "alertes", "surcharge · retard · emballage",
                         "alerte", style="blanc", statut="critique" if risques else "normal")
            self.lien("Ouvrir 🔔 →", None, "lien_kpi_risques", ouvrir_panneau=True)
        if v.empty:
            st.write("")
            self._aucune_prevision("dashboard")
            return
        st.write("")
        g1, g2 = st.columns([2.1, 1])
        with g1:
            with st.container(border=True):
                st.markdown("**Flux entrant prévu vs capacité** — cliquez sur une barre pour filtrer le tableau")
                selection = self._graphique_flux(v, sh, vue, shift_vu)
        with g2:
            with st.container(border=True):
                st.markdown("**Répartition des volumes annoncés**")
                axe = st.radio("Répartition", ["Client", "Famille", "Priorité"], horizontal=True,
                               label_visibility="collapsed")
                self._anneau(v, axe)
        with st.container(border=True):
            st.markdown("**Annonces futures et prédictions**")
            self._tableau_futur(v, recherche, selection, vue)

    def _graphique_flux(self, v, sh, vue, shift_vu) -> list:
        regle = "D" if vue == "Jour" else "W-MON"
        vol = v.groupby([pd.Grouper(key="date_shift", freq=regle), "shift"])["volume_kg"].sum().reset_index()
        cap = sh.groupby(pd.Grouper(key="date_shift", freq=regle))["capacite"].sum().reset_index()
        fig = go.Figure()
        couleurs = {"jour": Config.COULEURS["vert"], "nuit": "#5D8A73"}
        for shift in Config.SHIFTS:
            d = vol[vol["shift"] == shift]
            fig.add_bar(x=d["date_shift"], y=d["volume_kg"], name=f"Volume annoncé — {shift}",
                        marker_color=couleurs[shift], opacity=1.0 if shift_vu in ("Tous les shifts", shift) else 0.25,
                        hovertemplate="%{x|%d/%m/%Y}<br>%{y:,.0f} kg<extra></extra>")
        fig.add_scatter(x=cap["date_shift"], y=cap["capacite"], mode="lines+markers", name="Capacité disponible",
                        line=dict(color=Config.COULEURS["tension"], width=2.5),
                        hovertemplate="%{x|%d/%m/%Y}<br>Capacité %{y:,.0f} kg<extra></extra>")
        sur = sh[sh["depassement"] > 0]
        if vue == "Jour" and len(sur):
            pts = sur.groupby("date_shift")["charge"].sum().reset_index()
            fig.add_scatter(x=pts["date_shift"], y=pts["charge"], mode="markers", name="▲ Surcharge (dépassement)",
                            marker=dict(color=Config.COULEURS["critique"], size=9, symbol="triangle-up"))
        fig.update_layout(barmode="stack", yaxis_title="kg", xaxis_title=None)
        return self.ui.graphique_selection(fig, 340, "sel_flux_futur")

    def _anneau(self, v, axe):
        col = {"Client": v["client_id"].map(self.noms_clients), "Famille": v["famille"],
               "Priorité": v["priorite"].str.capitalize()}[axe]
        d = v.assign(cle=col).groupby("cle", as_index=False)["volume_kg"].sum()
        fig = px.pie(d, names="cle", values="volume_kg", hole=0.62, color_discrete_sequence=Config.PALETTE_GRAPHIQUES)
        fig.update_traces(textinfo="percent", hovertemplate="%{label}<br>%{value:,.0f} kg<extra></extra>")
        fig.add_annotation(text=f"<b>{self.ui.nombre(d['volume_kg'].sum() / 1000, 1)}</b><br>tonnes",
                           showarrow=False, font=dict(size=15))
        self.ui.graphique(fig, 300)

    def _tableau_futur(self, v, recherche, selection, vue):
        ui = self.ui
        f = v.sort_values("heure_arrivee_prevue")
        tableau = pd.DataFrame({
            "_date": f["date_shift"], "Arrivée prévue": pd.to_datetime(f["heure_arrivee_prevue"]).dt.strftime("%d/%m/%Y %H:%M"),
            "Lot": f["lot_id"], "Ferme": f["ferme_id"], "SKU": f["sku_id"], "Client": f["client_id"].map(self.noms_clients),
            "Volume (kg)": f["volume_kg"].round(0), "Shift": f["shift"],
            "Statut": f["statut_prediction"].fillna("—").str.capitalize(),
            "Rendement prévu (kg/h)": f["rendement_prevu_kg_h"].round(0),
            "Cycle prédit": f["temps_cycle_predit_min"].map(Formats.duree),
            "Expédition prévue": pd.to_datetime(f["heure_expedition_predite"]).dt.strftime("%d/%m %H:%M"),
            "Risque": f["niveau_risque"].map(self.libelle_risque)})
        if selection:
            jours = pd.to_datetime(pd.Series(selection), errors="coerce").dt.normalize().dropna()
            if len(jours):
                if vue == "Semaine":
                    masque = tableau["_date"].apply(lambda d: any(j <= d < j + pd.Timedelta(days=7) for j in jours))
                else:
                    masque = tableau["_date"].isin(jours)
                tableau = tableau[masque]
                st.caption("Filtré sur la sélection du graphique : "
                           + ", ".join(j.strftime("%d/%m/%Y") for j in jours) + " (cliquez dans le vide pour annuler)")
        if recherche:
            masque = tableau.astype(str).apply(lambda col: col.str.contains(recherche, case=False, regex=False)).any(axis=1)
            tableau = tableau[masque]
        if ui.vide(tableau):
            return
        ui.tableau(tableau.drop(columns="_date"), hauteur=420)
        c1, c2 = st.columns([3, 1])
        lot = c1.selectbox("Ouvrir le détail d'un lot", tableau["Lot"].tolist(), label_visibility="collapsed")
        if c2.button("Voir dans le planning →", key="dash_detail"):
            jour = pd.Timestamp(v.loc[v["lot_id"] == lot, "date_arrivee"].iloc[0]).date()
            st.session_state["cal_jour"] = jour
            st.session_state["cal_mois"] = jour.replace(day=1)
            st.session_state["page_cible"] = "Planning prédictif"
            st.rerun()

    # ================================================================
    #  EXPLICATION DU CYCLE (rendement ↔ temps de cycle)
    # ================================================================
    def _explication_cycle(self, r):
        """Montre que rendement et temps de cycle sont cohérents : le cycle est une somme d'étapes."""
        ui = self.ui
        lignes = pd.DataFrame([
            ("1. Attente avant traitement", r["temps_attente_min"], "file d'attente à la réception"),
            ("2. Tri, attentes inter-postes et arrêts", r["temps_hors_production_min"], "temps de traitement hors production"),
            ("3. Production effective (conditionnement)", r["temps_production_min"],
             f"{ui.nombre(r['poids_emballe_prevu_kg'])} kg emballés ÷ {ui.nombre(r['rendement_prevu_kg_h'])} kg/h"),
            ("4. Palettisation", r["temps_palettisation_min"], ""),
            ("5. Préparation de l'expédition", r["temps_preparation_expedition_min"], ""),
        ], columns=["Étape", "minutes", "Calcul / commentaire"])
        lignes["Durée"] = lignes["minutes"].map(Formats.duree)
        total = lignes["minutes"].sum()
        st.markdown(
            f"**Rendement prévu = {ui.nombre(r['rendement_prevu_kg_h'])} kg/h** : c'est le débit de la ligne pendant la "
            f"production effective (kg emballés par heure de production, même définition que dans la base "
            f"d'apprentissage).  \n"
            f"Poids emballé prévu = {ui.nombre(r['volume_kg'])} kg × (1 − {ui.nombre(r['taux_dechet_prevu'] * 100, 1)} % "
            f"de déchets) = **{ui.nombre(r['poids_emballe_prevu_kg'])} kg** → production effective = "
            f"{ui.nombre(r['poids_emballe_prevu_kg'])} ÷ {ui.nombre(r['rendement_prevu_kg_h'])} = "
            f"**{Formats.duree(r['temps_production_min'])}**.  \n"
            f"**Temps de cycle = {Formats.duree(total)}** de l'arrivée du camion à l'expédition : la production "
            f"n'en représente que {ui.nombre(r['temps_production_min'] / total * 100 if total else 0, 0)} %.")
        ui.tableau(lignes[["Étape", "Durée", "Calcul / commentaire"]])

    # ================================================================
    #  CAPACITÉ & SHIFTS (shifts prévus uniquement)
    # ================================================================
    def page_capacite(self):
        ui = self.ui
        ui.titre("Capacité & shifts", "Charge prévue des shifts futurs vs capacité disponible (poste goulot, "
                                      "opérateurs, absences)")
        v = self.vue_futures(filtrer=True)
        sh = self.shifts_futurs(v)
        k = st.columns(4)
        with k[0]:
            ui.carte_kpi("Shifts prévus", ui.nombre(len(sh)), "shifts", "avec au moins une annonce", "horloge")
        with k[1]:
            n = int((sh["depassement"] > 0).sum()) if len(sh) else 0
            ui.carte_kpi("Shifts en surcharge", ui.nombre(n), "shifts", "charge > capacité", "alerte",
                         statut="critique" if n else "normal")
        with k[2]:
            ui.carte_kpi("Dépassement total", ui.nombre(sh["depassement"].sum() if len(sh) else 0), "kg",
                         "quantité à arbitrer, décaler ou absorber", "colis")
        with k[3]:
            ui.carte_kpi("Capacité nominale goulot", ui.nombre(Config.capacite_goulot_kg_h()), "kg/h",
                         f"{Config.POSTE_GOULOT} · {ui.nombre(Config.capacite_goulot_kg_h() * Config.DUREE_SHIFT_UTILE_H)} "
                         f"kg par shift de {Config.DUREE_SHIFT_UTILE_H} h", "jauge", style="blanc")
        st.write("")
        o1, o2 = st.tabs(["Charge des shifts prévus", "Capacité de référence par poste"])
        with o1:
            if sh.empty:
                self._aucune_prevision("capacite")
            else:
                sh = sh.assign(libelle=pd.to_datetime(sh["date_shift"]).dt.strftime("%d/%m") + " " + sh["shift"])
                couleurs = {"normal": Config.COULEURS["normal"], "tension": Config.COULEURS["tension"],
                            "surcharge": Config.COULEURS["critique"], "forte_surcharge": "#7F0000"}
                fig = go.Figure()
                fig.add_bar(x=sh["libelle"], y=sh["charge"], name="Charge prévue (kg)",
                            marker_color=sh["etat"].map(couleurs).fillna("#9E9E9E"),
                            text=sh["etat"].map(ui.libelle), textposition="outside")
                fig.add_scatter(x=sh["libelle"], y=sh["capacite"], mode="lines+markers", name="Capacité disponible (kg)",
                                line=dict(color=Config.COULEURS["panneau"], dash="dash"))
                fig.update_layout(yaxis_title="kg")
                with st.container(border=True):
                    ui.graphique(fig, 340, "Charge prévue vs capacité disponible par shift (kg)")
                ui.tableau(pd.DataFrame({
                    "Date": pd.to_datetime(sh["date_shift"]).dt.strftime("%d/%m/%Y"), "Shift": sh["shift"],
                    "Lots": sh["lots"], "Charge prévue (kg)": sh["charge"].round(0),
                    "Capacité disponible (kg)": sh["capacite"].round(0), "Taux d'utilisation (%)": sh["taux"],
                    "Dépassement (kg)": sh["depassement"].round(0), "État": sh["etat"].map(ui.libelle)}), hauteur=360)
                st.caption("Capacité disponible = capacité nominale du goulot × heures utiles × opérateurs présents ÷ "
                           f"effectif de référence ({Config.NB_OPERATEURS_REFERENCE[Config.POSTE_GOULOT]}) × facteur shift "
                           f"(nuit {Config.FACTEUR_NUIT}). Les opérateurs et absences sont ceux saisis avec l'annonce.")
        with o2:
            st_ = self.ctx.tables["stations"].copy()
            nominale = st_["capacite_nominale_kg_h"].astype(float).where(st_["poste"] != Config.POSTE_GOULOT,
                                                                         Config.capacite_goulot_kg_h())
            ref = pd.DataFrame({
                "Station": st_["station_id"], "Poste": st_["poste"], "Capacité nominale (kg/h)": nominale.round(0),
                "Capacité shift de jour (kg)": (nominale * Config.DUREE_SHIFT_UTILE_H).round(0),
                "Capacité shift de nuit (kg)": (nominale * Config.DUREE_SHIFT_UTILE_H * Config.FACTEUR_NUIT).round(0),
                "Effectif de référence": st_["poste"].map(Config.NB_OPERATEURS_REFERENCE)})
            ui.tableau(ref)
            fig = px.bar(ref, x="Poste", y="Capacité shift de jour (kg)", text="Capacité shift de jour (kg)")
            fig.update_traces(marker_color=[Config.COULEURS["critique"] if p == Config.POSTE_GOULOT
                                            else Config.COULEURS["vert"] for p in ref["Poste"]],
                              texttemplate="%{text:,.0f} kg")
            ui.graphique(fig, 300, f"Capacité par poste — le goulot ({Config.POSTE_GOULOT}) limite le shift")

    # ================================================================
    #  FLUX ENTRANTS : saisie, import, annonces futures, traçabilité, historique
    # ================================================================
    def page_flux(self):
        ui, ia = self.ui, self.ctx.ia
        ui.titre("Flux entrants", "Saisie et import des nouvelles annonces, lancement des prédictions et suivi")
        if not ia.prets:
            st.warning("Les modèles d'IA ne sont pas entraînés : les prédictions ne peuvent pas être calculées.")
            if st.button("Entraîner les modèles maintenant (≈ 1 min)", type="primary"):
                with st.spinner("Entraînement en cours (validation temporelle)…"):
                    ia.entrainer(journal=lambda *_: None)
                st.rerun()
        o = st.tabs(["✎ Nouvelle annonce", "⭱ Import Excel / CSV", "☰ Annonces futures & prédictions",
                     "🧾 Traçabilité des prédictions"])
        with o[0]:
            self._nouvelle_annonce()
        with o[1]:
            self._import_annonces()
        with o[2]:
            self._annonces_futures()
        with o[3]:
            self._tracabilite()

    def _nouvelle_annonce(self):
        ui, t, emb = self.ui, self.ctx.tables, self.ctx.emballages
        ss = st.session_state
        st.caption("Les champs se mettent à jour automatiquement : l'emballage est proposé selon le SKU et le client.")
        c1, c2, c3 = st.columns(3)
        date_annonce = c1.date_input("Date d'annonce", value=date.today(), format="DD/MM/YYYY", key="na_date_annonce")
        jour = c2.date_input("Date d'arrivée prévue", value=date.today() + timedelta(days=1), format="DD/MM/YYYY",
                             key="na_jour")
        h = c3.time_input("Heure d'arrivée prévue", value=heure(8, 0), step=300, key="na_heure")
        c4, c5, c6 = st.columns(3)
        ferme = c4.selectbox("Ferme", t["fermes"]["ferme_id"].tolist(), format_func=self.lib_ferme, key="na_ferme")
        sku = c5.selectbox("SKU / produit", self.sku.index.tolist(), format_func=self.lib_sku, key="na_sku")
        client = c6.selectbox("Client", list(self.noms_clients), format_func=self.lib_client, key="na_client")
        c7, c8, c9 = st.columns(3)
        volume = c7.number_input("Volume annoncé (kg)", min_value=1.0, max_value=50000.0, value=1500.0, step=50.0,
                                 key="na_volume")
        compatibles = emb.compatibles(sku, client)
        propose = compatibles[0]
        article = c8.selectbox("Emballage (proposé automatiquement, modifiable)", compatibles, index=0,
                               key=f"na_emb_{sku}_{client}",
                               format_func=lambda x: f"{x}  ★ proposé" if x == propose else x)
        prio = c9.selectbox("Priorité", Config.PRIORITES, format_func=str.capitalize, key="na_prio")
        c10, c11, c12 = st.columns(3)
        lot = c10.text_input("Lot (vide = identifiant généré)", key="na_lot", placeholder="ex. LOT-P-20261012-001")
        origine = c11.selectbox("Origine de l'annonce", [o for o in Config.ORIGINES_SAISIE if o != "IMPORT"],
                                format_func=str.capitalize, key="na_origine")
        distance = self.distances.get(ferme, 0)
        arrivee = pd.Timestamp(jour) + pd.Timedelta(hours=h.hour, minutes=h.minute)
        estime = arrivee - pd.Timedelta(hours=distance / Config.VITESSE_CAMION_KM_H)
        avec_depart = c12.checkbox(f"Renseigner le départ camion (estimé {estime:%H:%M})", key="na_dep_on")
        depart = None
        if avec_depart:
            d1, d2 = st.columns(2)
            jd = d1.date_input("Date de départ camion", value=estime.date(), format="DD/MM/YYYY", key="na_dep_jour")
            hd = d2.time_input("Heure de départ camion", value=estime.time().replace(second=0, microsecond=0),
                               step=300, key="na_dep_heure")
            depart = pd.Timestamp(jd) + pd.Timedelta(hours=hd.hour, minutes=hd.minute)
        shift = "jour" if 6 <= h.hour < 22 else "nuit"
        with st.expander(f"Shift déduit de l'heure : {shift} — opérateurs prévus et absences"):
            e1, e2 = st.columns(2)
            ops = e1.selectbox("Opérateurs au poste goulot (soudage)", list(range(1, 21)),
                               index=Config.OPERATEURS_DEFAUT[shift] - 1, key=f"na_ops_{shift}")
            absences = e2.selectbox("Absences prévues", list(range(0, 6)), key="na_abs")
        # Aperçu du contrôle de stock d'emballage
        besoin = emb.unites_necessaires(article, sku, volume, client)
        stock = emb._stock(article, client, jour, self.prev.receptions.charger())
        etat = "normal" if stock["stock"] >= besoin else "tension"
        ui.afficher_badge(etat, f"Emballage {article} : besoin ≈ {ui.nombre(besoin)} {stock['unite'] or 'unités'} · "
                                f"stock {self.noms_clients[client]} {ui.nombre(stock['stock'])} {stock['unite']} "
                                f"(dernier relevé) · délai fournisseur {stock['delai']} j")
        if article != propose:
            st.caption(f"Choix manuel : {article} au lieu de {propose} (sera conservé dans la prédiction).")
        if st.button("💾 Enregistrer et lancer la prédiction", type="primary", key="na_enregistrer",
                     disabled=not self.ctx.ia.prets):
            brut = pd.DataFrame([{"lot_id": lot, "date_annonce": date_annonce, "date_arrivee": jour,
                                  "heure_arrivee": h, "depart_camion_prevu": depart, "ferme_id": ferme,
                                  "sku_id": sku, "client_id": client, "volume_kg": volume,
                                  "article_emballage": article, "priorite": prio, "origine_annonce": origine,
                                  "nb_operateurs": ops, "absences": absences}])
            valides, erreurs = self.prev.valider(brut, origine_defaut=origine)
            if len(erreurs):
                for _, e in erreurs.iterrows():
                    st.error(f"{e['colonne']} : {e['message']}")
                return
            try:
                with st.spinner("Contrôles, calcul de charge et prédiction…"):
                    r = self.prev.ajouter(valides)
                ss["na_derniere"] = r["annonce_id"].iloc[0]
            except Exception as erreur:
                st.error(f"Enregistrement impossible : {erreur}")
                return
        dernier = ss.get("na_derniere")
        if dernier:
            v = self.prev.vue()
            v = v[v["annonce_id"] == dernier]
            if len(v):
                self._resultat_prediction(v.iloc[0])

    def _resultat_prediction(self, r):
        ui = self.ui
        st.success(f"Annonce {r['annonce_id']} (lot {r['lot_id']}) enregistrée comme donnée future, "
                   f"prédiction {r['prediction_id']} ajoutée au planning prédictif.")
        k = st.columns(5)
        with k[4]:
            ui.carte_kpi("Rendement prévu", ui.nombre(r["rendement_prevu_kg_h"]), "kg/h",
                         f"production effective {Formats.duree(r['temps_production_min'])} pour "
                         f"{ui.nombre(r['poids_emballe_prevu_kg'])} kg emballés", "activite", style="blanc")
        with k[0]:
            ui.carte_kpi("Temps de cycle prédit", Formats.duree(r["temps_cycle_predit_min"]), "",
                         "de l'arrivée du camion à l'expédition", "horloge")
        with k[1]:
            exp = pd.Timestamp(r["heure_expedition_predite"])
            ui.carte_kpi("Expédition prévue", exp.strftime("%H:%M"), "",
                         f"le {exp:%d/%m/%Y} · limite client {pd.Timestamp(r['heure_limite_client']):%H:%M}",
                         "camion", statut="critique" if r["marge_expedition_min"] < 0 else
                         "tension" if r["marge_expedition_min"] < Config.SEUIL_MARGE_EXPEDITION_MIN else "normal")
        with k[2]:
            ui.carte_kpi("Charge / capacité", ui.nombre(r["taux_utilisation_pct"], 0), "%",
                         f"{ui.nombre(r['charge_annoncee_kg'])} kg / {ui.nombre(r['capacite_disponible_kg'])} kg "
                         f"· dépassement {ui.nombre(r['depassement_kg'])} kg", "jauge", statut=r["etat_charge"])
        with k[3]:
            ui.carte_kpi("Niveau de risque", self.libelle_risque(r["niveau_risque"]), "", str(r["motifs_risque"]),
                         "alerte", style="blanc",
                         statut={"critique": "critique", "tension": "tension"}.get(r["niveau_risque"], "normal"))
        with st.expander("Comment le temps de cycle est-il obtenu à partir du rendement ?", expanded=True):
            self._explication_cycle(r)
        with st.container(border=True):
            st.markdown(f"**🚚 Prochain batch conseillé** — shift {r['shift']} du {pd.Timestamp(r['date_shift']):%d/%m/%Y}")
            self._bloc_prochain_batch(self.prev.prochain_batch(r["date_shift"], r["shift"]), f"res_{r['annonce_id']}")
        jour = pd.Timestamp(r["date_arrivee"]).date()
        self.lien("📅 Voir dans le planning prédictif →", "Planning prédictif", f"lien_planning_{r['annonce_id']}",
                  cal_jour=jour, cal_mois=jour.replace(day=1))

    def _import_annonces(self):
        ui = self.ui
        st.markdown("Importez plusieurs annonces en une opération (origine **IMPORT**). Colonnes : "
                    "`lot_id, date_annonce, date_arrivee, heure_arrivee, depart_camion_prevu, ferme_id, sku_id, "
                    "client_id, volume_kg, article_emballage, priorite, nb_operateurs, absences`. "
                    "Obligatoires : date_arrivee, ferme_id, sku_id, client_id, volume_kg. "
                    "Emballage vide = proposition automatique ; client accepté par son nom (ex. Mercadona).")
        modele = self.prev.modele_import()
        c1, c2 = st.columns(2)
        c1.download_button("⭳ Modèle CSV", modele.to_csv(index=False).encode("utf-8-sig"), "modele_annonces.csv",
                           "text/csv")
        c2.download_button("⭳ Modèle Excel", self._xlsx({"annonces": modele}), "modele_annonces.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        fichier = st.file_uploader("Fichier d'annonces (.csv, .xlsx)", type=["csv", "xlsx", "xls"])
        if fichier is None:
            return
        try:
            brut = self.prev.lire_fichier(fichier.name, fichier.getvalue())
        except Exception as erreur:
            st.error(f"Lecture impossible : {erreur}")
            return
        valides, erreurs = self.prev.valider(brut, origine_defaut="IMPORT")
        if len(valides):
            valides["origine_annonce"] = "IMPORT"
        c1, c2 = st.columns(2)
        with c1:
            ui.afficher_badge("normal" if len(valides) else "neutre", f"{len(valides)} ligne(s) valide(s)")
        with c2:
            nb = erreurs["ligne"].nunique() if len(erreurs) else 0
            ui.afficher_badge("critique" if nb else "normal", f"{nb} ligne(s) rejetée(s)")
        if len(erreurs):
            st.markdown("**Lignes rejetées et motifs**")
            ui.tableau(erreurs.rename(columns={"ligne": "Ligne du fichier", "colonne": "Colonne", "message": "Motif"}),
                       hauteur=220)
        if len(valides):
            st.markdown("**Lignes prêtes à être enregistrées**")
            ui.tableau(valides.assign(date_arrivee=lambda d: pd.to_datetime(d["date_arrivee"]).dt.strftime("%d/%m/%Y"))[
                ["lot_id", "date_arrivee", "heure_arrivee", "ferme_id", "sku_id", "client_id", "volume_kg",
                 "article_emballage", "emballage_modifie", "priorite"]].rename(columns={
                    "emballage_modifie": "emballage modifié"}), hauteur=260)
            if st.button(f"Importer les {len(valides)} annonce(s) et lancer les prédictions", type="primary",
                         disabled=not self.ctx.ia.prets):
                try:
                    with st.spinner("Import, contrôles et prédictions…"):
                        r = self.prev.ajouter(valides, "IMPORT")
                    st.success(f"{len(r)} annonce(s) importée(s) et prédite(s) : elles apparaissent dans le planning.")
                except Exception as erreur:
                    st.error(f"Import impossible : {erreur}")

    def _annonces_futures(self):
        ui, t = self.ui, self.ctx.tables
        v = self.vue_futures(avec_archivees=True)
        if v.empty:
            st.info("Aucune annonce future : utilisez « Nouvelle annonce » ou « Import Excel / CSV ».")
            return
        montrer = st.checkbox("Afficher les annonces archivées", value=False)
        if not montrer:
            v = v[v["statut_prediction"] != "ARCHIVEE"]
        if ui.vide(v):
            return
        v = v.sort_values(["date_arrivee", "heure_arrivee"]).reset_index(drop=True)
        articles = sorted(set(t["mapping_emballages"].loc[t["mapping_emballages"]["role"] == "primaire", "article_emballage"]))
        vue = pd.DataFrame({
            "Sélection": False, "annonce_id": v["annonce_id"], "lot_id": v["lot_id"],
            "date_arrivee": pd.to_datetime(v["date_arrivee"]).dt.date, "heure_arrivee": v["heure_arrivee"].astype(str),
            "ferme_id": v["ferme_id"], "sku_id": v["sku_id"], "client_id": v["client_id"],
            "volume_kg": v["volume_kg"].astype(float), "article_emballage": v["article_emballage"],
            "priorite": v["priorite"], "nb_operateurs": v["nb_operateurs"].astype(int),
            "absences": v["absences"].astype(int),
            "Shift": v["shift"], "Rendement (kg/h)": v["rendement_prevu_kg_h"].round(0),
            "Production effective": v["temps_production_min"].map(Formats.duree),
            "Cycle prédit": v["temps_cycle_predit_min"].map(Formats.duree),
            "Expédition prédite": pd.to_datetime(v["heure_expedition_predite"]).dt.strftime("%d/%m %H:%M"),
            "Charge (%)": v["taux_utilisation_pct"], "Dépassement (kg)": v["depassement_kg"],
            "Emballage": v["statut_emballage"].map({"ok": "✓ ok", "risque": "⚠ risque", "rupture": "✕ rupture"}),
            "Risque": v["niveau_risque"].map(self.libelle_risque), "Statut": v["statut_prediction"],
        })
        cc = st.column_config
        calcule = ["annonce_id", "lot_id", "Shift", "Rendement (kg/h)", "Production effective", "Cycle prédit",
                   "Expédition prédite", "Charge (%)",
                   "Dépassement (kg)", "Emballage", "Risque", "Statut"]
        st.caption("Modifiez les cellules (listes déroulantes), cochez des lignes pour les valider, archiver ou "
                   "supprimer. Toute modification relance la prédiction.")
        edite = ui.editeur(vue, key="editeur_futures", num_rows="fixed", disabled=calcule, height=430, column_config={
            "Sélection": cc.CheckboxColumn("Sélection", default=False),
            "annonce_id": cc.TextColumn("Annonce"), "lot_id": cc.TextColumn("Lot"),
            "date_arrivee": cc.DateColumn("Arrivée", format="DD/MM/YYYY", required=True),
            "heure_arrivee": cc.TextColumn("Heure", validate=r"^\d{1,2}:\d{2}$", required=True),
            "ferme_id": cc.SelectboxColumn("Ferme", options=t["fermes"]["ferme_id"].tolist(), required=True),
            "sku_id": cc.SelectboxColumn("SKU", options=self.sku.index.tolist(), required=True),
            "client_id": cc.SelectboxColumn("Client", options=list(self.noms_clients), required=True),
            "volume_kg": cc.NumberColumn("Volume (kg)", min_value=1, max_value=50000, step=10, required=True),
            "article_emballage": cc.SelectboxColumn("Emballage", options=articles, required=True),
            "priorite": cc.SelectboxColumn("Priorité", options=Config.PRIORITES, required=True),
            "nb_operateurs": cc.SelectboxColumn("Opérateurs", options=list(range(1, 21)), required=True),
            "absences": cc.SelectboxColumn("Absences", options=list(range(0, 6)), required=True),
        })
        choisis = edite.loc[edite["Sélection"].astype(bool), "annonce_id"].tolist()
        b = st.columns(5)
        if b[0].button("💾 Enregistrer & recalculer", type="primary"):
            saisie = edite.drop(columns=["Sélection"] + [c for c in calcule if c not in ("annonce_id", "lot_id")])
            saisie = saisie.merge(v[["annonce_id", "date_annonce", "depart_camion_prevu", "origine_annonce"]],
                                  on="annonce_id", how="left")
            valides, erreurs = self.prev.valider(saisie, controle_lots=False)
            if len(erreurs):
                st.error("Corrigez les erreurs avant d'enregistrer :")
                ui.tableau(erreurs)
            else:
                with st.spinner("Recalcul des prédictions…"):
                    archivees = self.prev.annonces()
                    garder = archivees[~archivees["annonce_id"].isin(valides["annonce_id"])]
                    self.prev.remplacer(pd.concat([valides, garder.reindex(columns=valides.columns)], ignore_index=True))
                st.success("Modifications enregistrées et prédictions recalculées.")
                st.rerun()
        for col, libelle, statut in ((b[1], "✓ Valider", "VALIDEE"), (b[2], "🗄 Archiver", "ARCHIVEE")):
            if col.button(f"{libelle} ({len(choisis)})", disabled=not choisis):
                self.prev.changer_statut(choisis, statut)
                st.rerun()
        if b[3].button(f"🗑 Supprimer ({len(choisis)})", disabled=not choisis):
            self.prev.supprimer(choisis)
            st.rerun()
        b[4].download_button("⭳ Export CSV", ChargeurDonnees.format_iso(v).to_csv(index=False).encode(Config.ENCODAGE),
                             "annonces_futures_predictions.csv", "text/csv")

    def _tracabilite(self):
        ui = self.ui
        h = self.prev.historique()
        st.caption("Chaque calcul est tracé : date, version du modèle et des paramètres, entrées utilisées et "
                   "résultats produits (sorties_moteur/historique_predictions.csv).")
        if ui.vide(h):
            return
        annonces = ["Toutes"] + sorted(h["annonce_id"].dropna().unique().tolist())
        choix = st.selectbox("Annonce", annonces)
        if choix != "Toutes":
            h = h[h["annonce_id"] == choix]
        colonnes = [c for c in ("date_prediction", "prediction_id", "annonce_id", "statut_prediction", "version_modele",
                                "version_parametres", "temps_cycle_predit_min", "heure_expedition_predite",
                                "taux_utilisation_pct", "depassement_kg", "emballage_selectionne", "niveau_risque",
                                "entrees") if c in h.columns]
        ui.tableau(h.sort_values("date_prediction", ascending=False)[colonnes], hauteur=420)

    # ================================================================
    #  PLANNING PRÉDICTIF (calendrier + détail journalier + frise)
    # ================================================================
    MOIS_FR = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre",
               "octobre", "novembre", "décembre"]
    ETATS_JOUR = {"surcharge": ("✕", "#FDECEA", "#C62828", "#7F1D1D"), "tension": ("⚠", "#FFF5D6", "#E0A100", "#5C4300"),
                  "normal": ("✓", "#E8F5E9", "#2E7D32", "#1B4D20")}

    def _decaler_mois(self, pas: int) -> None:
        m = st.session_state["cal_mois"]
        st.session_state["cal_mois"] = (pd.Timestamp(m) + pd.DateOffset(months=pas)).date().replace(day=1)

    def _choisir_jour(self, jour) -> None:
        st.session_state["cal_jour"] = jour

    @staticmethod
    def _etat_jour(etats) -> str:
        etats = set(etats.dropna())
        if etats & {"surcharge", "forte_surcharge"}:
            return "surcharge"
        return "tension" if "tension" in etats else "normal"

    def page_planning(self):
        ui = self.ui
        ui.titre("Planning prédictif", "Arrivées futures saisies ou importées et leurs prédictions — "
                                       "résultats prévisionnels, jamais présentés comme réels")
        v = self.vue_futures(filtrer=True, avec_periode=False)
        v = v[v["heure_arrivee_prevue"].notna()] if len(v) else v
        jours_sur = v.groupby(v["date_arrivee"].dt.date)["etat_charge"].apply(self._etat_jour) if len(v) else pd.Series(dtype=str)
        k = st.columns(4)
        with k[0]:
            ui.carte_kpi("Annonces prédites", ui.nombre(len(v)), "lots", f"{ui.nombre(v['volume_kg'].sum() if len(v) else 0)} kg",
                         "calendrier")
        with k[1]:
            ui.carte_kpi("Jours planifiés", ui.nombre(len(jours_sur)), "jours", "avec au moins une arrivée", "horloge")
        with k[2]:
            n = int((jours_sur == "surcharge").sum())
            ui.carte_kpi("Jours en surcharge", ui.nombre(n), "jours",
                         f"{int((jours_sur == 'tension').sum())} jour(s) sous tension", "alerte",
                         statut="critique" if n else "normal")
        with k[3]:
            n = int((v["niveau_risque"] != "normal").sum()) if len(v) else 0
            ui.carte_kpi("Lots à risque", ui.nombre(n), "lots", "surcharge, retard ou emballage", "colis", style="blanc",
                         statut="tension" if n else "normal")
        st.write("")
        o1, o2 = st.tabs(["📅 Calendrier et détail du jour", "⏱ Fiabilité du modèle prédictif"])
        with o1:
            if v.empty:
                st.info("Aucune annonce prédite : créez ou importez des annonces dans « Flux entrants ».")
                self.lien("✎ Créer une annonce →", "Flux entrants", "lien_planning_vide")
                return
            self._calendrier(v, jours_sur)
        with o2:
            self._metriques_modeles()

    def _calendrier(self, v, etats_jours):
        import calendar as cal_mod
        ui, ss = self.ui, st.session_state
        if "cal_mois" not in ss:
            a_venir = v[v["date_arrivee"] >= pd.Timestamp(date.today())]
            ref = (a_venir if len(a_venir) else v)["date_arrivee"].min()
            ss["cal_mois"] = pd.Timestamp(ref).date().replace(day=1)
        mois = ss["cal_mois"]
        volumes = v.groupby(v["date_arrivee"].dt.date)["volume_kg"].sum().to_dict()
        camions = v.groupby(v["date_arrivee"].dt.date)["annonce_id"].count().to_dict()
        choisi = ss.get("cal_jour")
        if choisi is None or (choisi.year, choisi.month) != (mois.year, mois.month):
            du_mois = sorted(j for j in volumes if (j.year, j.month) == (mois.year, mois.month))
            choisi = du_mois[0] if du_mois else None
        with st.container(border=True):
            n1, n2, n3 = st.columns([1, 3, 1])
            n1.button("◀ Mois précédent", key="cal_prec", on_click=self._decaler_mois, args=(-1,))
            n2.markdown(f'<div class="cal-mois">{self.MOIS_FR[mois.month - 1].capitalize()} {mois.year}</div>',
                        unsafe_allow_html=True)
            n3.button("Mois suivant ▶", key="cal_suiv", on_click=self._decaler_mois, args=(1,))
            st.markdown(" ".join([ui.badge("normal") + " charge absorbable", ui.badge("tension") + " sous tension",
                                  ui.badge("surcharge") + " surcharge (dépassement)"]), unsafe_allow_html=True)
            css = []
            entetes = st.columns(7)
            for col, nom in zip(entetes, ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]):
                col.markdown(f'<div class="cal-entete">{nom}</div>', unsafe_allow_html=True)
            for semaine in cal_mod.Calendar(firstweekday=0).monthdatescalendar(mois.year, mois.month):
                cols = st.columns(7)
                for col, jour in zip(cols, semaine):
                    if jour.month != mois.month:
                        col.write("")
                        continue
                    cle = f"cal_{jour.isoformat()}"
                    vol = volumes.get(jour, 0)
                    libelle = f"{jour.day}"
                    if vol:
                        etat = etats_jours.get(jour, "normal")
                        icone, fond, bord, texte = self.ETATS_JOUR[etat]
                        libelle = f"{jour.day} {icone} {vol / 1000:.1f} t".replace(".", ",")
                        css.append(f".st-key-{cle} button {{ background:{fond} !important; border:1.5px solid {bord} !important; }}"
                                   f" .st-key-{cle} button p {{ color:{texte} !important; font-weight:600; }}")
                    if jour == choisi:
                        css.append(f".st-key-{cle} button {{ outline:3px solid {Config.COULEURS['panneau']}; outline-offset:1px; }}")
                    aide = (f"{ui.nombre(vol)} kg · {camions.get(jour, 0)} camion(s)") if vol else "Aucune arrivée prévue"
                    col.button(libelle, key=cle, on_click=self._choisir_jour, args=(jour,), help=aide,
                               **({"width": "stretch"} if ui._largeur_ok() else {"use_container_width": True}))
            if css:
                st.markdown("<style>" + "\n".join(css) + "</style>", unsafe_allow_html=True)
        if choisi is None:
            st.info("Aucune arrivée prévue ce mois-ci : changez de mois.")
            return
        self._detail_jour(v[v["date_arrivee"].dt.date == choisi].sort_values("heure_arrivee_prevue"), choisi)

    def _detail_jour(self, d, jour):
        ui = self.ui
        st.markdown(f"### Arrivées prévues le {jour:%d/%m/%Y}")
        if ui.vide(d):
            return
        k = st.columns(4)
        k[0].metric("Quantité annoncée", f"{ui.nombre(d['volume_kg'].sum())} kg")
        k[1].metric("Camions", f"{len(d)}")
        k[2].metric("Première / dernière arrivée",
                    f"{pd.Timestamp(d['heure_arrivee_prevue'].min()):%H:%M} / {pd.Timestamp(d['heure_arrivee_prevue'].max()):%H:%M}")
        k[3].metric("Dernière expédition prévue", f"{pd.Timestamp(d['heure_expedition_predite'].max()):%d/%m %H:%M}")
        # Résumé par shift (charge vs capacité)
        shifts = d.groupby(["date_shift", "shift"]).agg(charge=("charge_annoncee_kg", "first"),
                                                        capacite=("capacite_disponible_kg", "first"),
                                                        taux=("taux_utilisation_pct", "first"),
                                                        depassement=("depassement_kg", "first"),
                                                        etat=("etat_charge", "first"), lots=("annonce_id", "count")).reset_index()
        for _, s in shifts.iterrows():
            texte = (f"Shift {s['shift']} du {pd.Timestamp(s['date_shift']):%d/%m} : charge {ui.nombre(s['charge'])} kg / "
                     f"capacité {ui.nombre(s['capacite'])} kg ({ui.nombre(s['taux'], 0)} %)")
            if s["depassement"] > 0:
                texte += f" · dépassement {ui.nombre(s['depassement'])} kg à arbitrer"
            ui.afficher_badge({"forte_surcharge": "forte_surcharge"}.get(s["etat"], s["etat"]), texte)
        st.write("")
        tableau = pd.DataFrame({
            "Heure d'arrivée": pd.to_datetime(d["heure_arrivee_prevue"]).dt.strftime("%H:%M"),
            "Lot": d["lot_id"], "Ferme": d["ferme_id"], "Client": d["client_id"].map(self.noms_clients),
            "SKU": d["sku_id"].map(lambda s: f"{self.sku.at[s, 'produit']} — {self.sku.at[s, 'format']}"),
            "Quantité annoncée (kg)": d["volume_kg"].round(0), "Priorité": d["priorite"].str.capitalize(),
            "Rendement prévu (kg/h)": d["rendement_prevu_kg_h"].round(0),
            "Production effective": d["temps_production_min"].map(Formats.duree),
            "Temps de cycle prédit": d["temps_cycle_predit_min"].map(Formats.duree),
            "Expédition prédite": pd.to_datetime(d["heure_expedition_predite"]).dt.strftime("%H:%M"),
            "Charge / capacité": d["taux_utilisation_pct"].map(lambda x: f"{x:.0f} %"),
            "État": d["etat_charge"].map(ui.libelle), "Risque": d["niveau_risque"].map(self.libelle_risque),
            "Emballage": d["emballage_selectionne"], "Statut": d["statut_prediction"].str.capitalize(),
        })
        ui.tableau(tableau, hauteur=min(420, 80 + 36 * len(d)))
        with st.expander("Détail du calcul rendement → temps de cycle d'un lot"):
            ids = d["annonce_id"].tolist()
            lot = st.selectbox("Lot", ids, key=f"explication_lot_{jour}",
                               format_func=lambda a: f"{d.loc[d['annonce_id'] == a, 'lot_id'].iloc[0]} · {a}"
                               if a in ids else str(a))
            if lot not in ids:
                lot = ids[0]
            self._explication_cycle(d[d["annonce_id"] == lot].iloc[0])
        # Frise : arrivée → début de traitement → fin de traitement → palettisation → expédition
        segments = []
        for _, l in d.iterrows():
            nom = f"{l['lot_id']} · {l['ferme_id']} · {l['sku_id']}"
            for etape, debut, fin in (("Attente avant traitement", "heure_arrivee_prevue", "heure_debut_traitement"),
                                      ("Tri, attentes inter-postes et arrêts", "heure_debut_traitement", "heure_debut_production"),
                                      ("Production effective (conditionnement)", "heure_debut_production", "heure_fin_predite"),
                                      ("Palettisation", "heure_fin_predite", "heure_fin_palettisation"),
                                      ("Préparation expédition", "heure_fin_palettisation", "heure_expedition_predite")):
                if pd.notna(l[debut]) and pd.notna(l[fin]) and pd.Timestamp(l[fin]) > pd.Timestamp(l[debut]):
                    segments.append({"Lot": nom, "Étape": etape, "Début": pd.Timestamp(l[debut]), "Fin": pd.Timestamp(l[fin])})
        if segments:
            seg = pd.DataFrame(segments)
            fig = px.timeline(seg, x_start="Début", x_end="Fin", y="Lot", color="Étape",
                              color_discrete_map={"Attente avant traitement": Config.COULEURS["tension"],
                                                  "Tri, attentes inter-postes et arrêts": "#A5D6A7",
                                                  "Production effective (conditionnement)": Config.COULEURS["vert"],
                                                  "Palettisation": "#5D8A73", "Préparation expédition": "#C5A572"})
            fig.update_yaxes(autorange="reversed", title=None)
            for _, l in d.iterrows():
                fig.add_scatter(x=[pd.Timestamp(l["heure_limite_client"])], y=[f"{l['lot_id']} · {l['ferme_id']} · {l['sku_id']}"],
                                mode="markers", marker=dict(symbol="line-ns-open", size=22, color=Config.COULEURS["critique"],
                                                            line=dict(width=3)), name="Heure limite client", showlegend=False,
                                hovertemplate="Limite client %{x|%H:%M}<extra></extra>")
            with st.container(border=True):
                ui.graphique(fig, max(260, 90 + 48 * len(d)), "Frise prédite des lots du jour (trait rouge = limite client)")
        # Prochain batch conseillé pour chaque shift du jour
        for _, s_ in shifts.iterrows():
            with st.expander(f"🚚 Prochain batch conseillé — shift {s_['shift']} du {pd.Timestamp(s_['date_shift']):%d/%m/%Y}"):
                self._bloc_prochain_batch(self.prev.prochain_batch(s_["date_shift"], s_["shift"]),
                                          f"jour_{pd.Timestamp(s_['date_shift']):%Y%m%d}_{s_['shift']}")
        # Actions sur les prédictions du jour
        with st.expander("Valider ou archiver des prédictions de ce jour"):
            choix = st.multiselect("Lots", d["annonce_id"].tolist(),
                                   format_func=lambda a: f"{a} · {d.loc[d['annonce_id'] == a, 'lot_id'].iloc[0]}")
            c1, c2 = st.columns(2)
            if c1.button("✓ Valider", disabled=not choix, key="jour_valider"):
                self.prev.changer_statut(choix, "VALIDEE")
                st.rerun()
            if c2.button("🗄 Archiver", disabled=not choix, key="jour_archiver"):
                self.prev.changer_statut(choix, "ARCHIVEE")
                st.rerun()

    # ================================================================
    #  EMBALLAGES
    # ================================================================
    def _bloc_prochain_batch(self, info: dict, cle: str):
        """Affiche quand demander le batch suivant aux fermes (et jusqu'à quelle quantité)."""
        ui = self.ui
        st.caption(f"Ligne libre (poste goulot) à {info['ligne_libre']:%d/%m %H:%M} après {info['nb_lots']} lot(s) prévu(s)"
                   + (f" · file d'attente estimée {Formats.duree(info['file_attente_min'])} entre lots"
                      if info["file_attente_min"] > 0 else "")
                   + f" · préparation d'un lot (attente + tri) ≈ {Formats.duree(info['preparation_min'])}.")
        if info["sature"]:
            ui.afficher_badge("tension", f"Shift {info['shift']} du {info['date_shift']:%d/%m} complet : capacité restante "
                                         f"{ui.nombre(info['capacite_restante_kg'])} kg — batch à prévoir sur le shift suivant")
            suivant = info.get("suivant")
            if suivant and not suivant["sature"]:
                info = suivant
            elif suivant:
                info = suivant
                while info.get("suivant") and info["sature"]:
                    info = info["suivant"]
            else:
                return
        k = st.columns(3)
        k[0].metric("Arrivée conseillée", f"{info['arrivee_conseillee']:%d/%m %H:%M}",
                    help="Heure à laquelle le camion doit arriver pour être prêt quand la ligne se libère")
        k[1].metric("Quantité maximale", f"{ui.nombre(info['quantite_max_kg'])} kg",
                    help="Min(capacité restante du shift, temps restant × rendement prévu)")
        k[2].metric("Expédition estimée", f"{info['expedition_estimee']:%d/%m %H:%M}",
                    help="Si le batch arrive à l'heure conseillée avec la quantité maximale")
        f = info["fermes"]
        ui.tableau(pd.DataFrame({"Ferme": f["ferme_id"], "Distance (km)": f["distance_usine_km"],
                                 "Trajet (min)": f["trajet_min"].astype(int),
                                 "Départ conseillé de la ferme": pd.to_datetime(f["depart_conseille"]).dt.strftime("%d/%m %H:%M"),
                                 "Arrivée à la station": info["arrivee_conseillee"].strftime("%d/%m %H:%M")}),
                   hauteur=min(330, 60 + 36 * len(f)))
        st.caption(f"Shift {info['shift']} du {info['date_shift']:%d/%m/%Y} — estimation : lots traités dans l'ordre, "
                   f"sans aléa ; trajet = distance ÷ {Config.VITESSE_CAMION_KM_H} km/h.")

    def _plan_reappro(self):
        ui = self.ui
        c1, c2 = st.columns([1, 2])
        horizon = c1.slider("Horizon (jours)", 14, 120, Config.HORIZON_PLAN_J, step=7, key="plan_horizon")
        fond = c2.checkbox("Ajouter la consommation de fond (moyenne apprise de la base) pour les arrivées pas encore "
                           "annoncées", value=False, key="plan_fond",
                           help="Chaque jour consomme au moins la moyenne apprise ; les annonces enregistrées la "
                                "remplacent quand elles sont plus fortes (pas de double comptage).")
        res, proj = self.prev.plan_reappro(horizon, fond)
        if res.empty:
            st.info("Aucun besoin d'emballage : enregistrez des annonces futures (ou activez la consommation de fond).")
            return
        res = res[res["client_id"].isin(self.f.clients)] if self.f.clients else res
        k = st.columns(3)
        with k[0]:
            n = int((res["statut"] == "critique").sum())
            ui.carte_kpi("À commander maintenant", ui.nombre(n), "articles", "délai fournisseur déjà atteint", "alerte",
                         statut="critique" if n else "normal")
        with k[1]:
            n = int((res["statut"] == "tension").sum())
            ui.carte_kpi(f"À commander sous {Config.ALERTE_COMMANDE_J} jours", ui.nombre(n), "articles",
                         "ou stock de sécurité entamé", "horloge", statut="tension" if n else "normal")
        with k[2]:
            ui.carte_kpi("Couverts sur l'horizon", ui.nombre(int((res["statut"] == "normal").sum())), "articles",
                         f"{horizon} jours", "valide", style="blanc")
        st.write("")
        dt = lambda x: pd.to_datetime(x).dt.strftime("%d/%m/%Y").fillna("—")
        ui.tableau(pd.DataFrame({
            "Article": res["article"], "Client": res["client_id"].map(self.noms_clients),
            "Stock aujourd'hui": res["stock_aujourdhui"].map(ui.nombre) + " " + res["unite"],
            "dont réceptions saisies": res["receptions_saisies"].map(ui.nombre),
            "Consommation prévue": res["consommation_horizon"].map(ui.nombre),
            "Rupture prévue": dt(res["date_rupture"]), "Commander avant le": dt(res["date_commande_limite"]),
            "Délai fournisseur (j)": res["delai_fournisseur_j"],
            "Quantité conseillée": res["quantite_conseillee"].map(ui.nombre) + " " + res["unite"],
            "Statut": res["statut"].map(ui.libelle), "Conseil": res["message"]}), hauteur=380)
        st.caption("Stock projeté = dernier relevé + réceptions saisies − consommation prévue. Commande = date de rupture − "
                   "délai fournisseur. Quantité conseillée = de quoi finir l'horizon avec le stock de sécurité "
                   "(seuil de couverture de l'article). Unités = unités de stock (cartons, sachets, rouleaux).")
        if len(proj):
            paires = list(zip(res["article"], res["client_id"]))
            choix = st.selectbox("Projection du stock", paires, key="plan_paire",
                                 format_func=lambda p: f"{p[0]} · {self.noms_clients.get(p[1], p[1])}")
            p = proj[(proj["article"] == choix[0]) & (proj["client_id"] == choix[1])]
            l = res[(res["article"] == choix[0]) & (res["client_id"] == choix[1])].iloc[0]
            fig = go.Figure()
            fig.add_scatter(x=p["date"], y=p["stock_projete"], mode="lines", name=f"Stock projeté ({l['unite']})",
                            line=dict(color=Config.COULEURS["vert"], width=3), fill="tozeroy", fillcolor="rgba(46,125,50,.08)")
            fig.add_bar(x=p["date"], y=p["consommation"], name="Consommation prévue", marker_color="#C5A572", opacity=.6)
            recu = p[p["reception"] > 0]
            if len(recu):
                fig.add_scatter(x=recu["date"], y=recu["stock_projete"], mode="markers", name="Réception saisie",
                                marker=dict(symbol="star", size=13, color=Config.COULEURS["panneau"]))
            fig.add_hline(y=l["stock_securite"], line_dash="dot", line_color=Config.COULEURS["tension"],
                          annotation_text="stock de sécurité")
            fig.add_hline(y=0, line_color=Config.COULEURS["critique"])
            if pd.notna(l["date_rupture"]):
                fig.add_vline(x=pd.Timestamp(l["date_rupture"]), line_dash="dash", line_color=Config.COULEURS["critique"])
            if pd.notna(l["date_commande_limite"]) and pd.Timestamp(l["date_commande_limite"]) >= p["date"].min():
                fig.add_vline(x=pd.Timestamp(l["date_commande_limite"]), line_dash="dash", line_color=Config.COULEURS["tension"])
            with st.container(border=True):
                ui.graphique(fig, 340, f"Stock projeté — {choix[0]} · {self.noms_clients.get(choix[1], choix[1])}")

    def _receptions_emballages(self):
        ui, t = self.ui, self.ctx.tables
        st.markdown("Déclarez ici un réapprovisionnement reçu : le stock projeté, les tensions des lots et les "
                    "alertes de la cloche 🔔 sont recalculés immédiatement.")
        articles = sorted(t["stock_emballages"]["article"].unique())
        c1, c2, c3, c4 = st.columns([1.6, 1.3, 1, 1])
        article = c1.selectbox("Article reçu", articles, key="rec_article")
        client = c2.selectbox("Client", list(self.noms_clients), format_func=self.lib_client, key="rec_client")
        unite = self.ctx.emballages._stock(article, client, date.today())["unite"] or "unités"
        quantite = c3.number_input(f"Quantité ({unite})", min_value=0.0, value=0.0, step=10.0, key="rec_quantite")
        jour = c4.date_input("Date de réception", value=date.today(), format="DD/MM/YYYY", key="rec_jour")
        commentaire = st.text_input("Commentaire (fournisseur, bon de livraison…)", key="rec_commentaire")
        if st.button("📦 Enregistrer la réception", type="primary", disabled=quantite <= 0):
            try:
                with st.spinner("Enregistrement et recalcul des alertes…"):
                    rid = self.prev.enregistrer_reception(article, client, quantite, jour, commentaire)
                st.success(f"Réception {rid} enregistrée : {ui.nombre(quantite)} {unite} de {article} pour "
                           f"{self.noms_clients[client]} le {jour:%d/%m/%Y}. Alertes mises à jour.")
            except Exception as erreur:
                st.error(f"Enregistrement impossible : {erreur}")
        r = self.prev.receptions.charger()
        st.markdown("**Réceptions enregistrées**")
        if r.empty:
            st.info("Aucune réception saisie.")
            return
        vue = pd.DataFrame({"Supprimer": False, "Réception": r["reception_id"],
                            "Date": pd.to_datetime(r["date_reception"]).dt.strftime("%d/%m/%Y"), "Article": r["article"],
                            "Client": r["client_id"].map(self.noms_clients),
                            "Quantité": r["quantite"].map(ui.nombre) + " " + r["unite"], "Commentaire": r["commentaire"]})
        edite = ui.editeur(vue, key="editeur_receptions", num_rows="fixed",
                           disabled=[c for c in vue.columns if c != "Supprimer"],
                           column_config={"Supprimer": st.column_config.CheckboxColumn("Supprimer", default=False)})
        choisis = edite.loc[edite["Supprimer"].astype(bool), "Réception"].tolist()
        if st.button(f"🗑 Supprimer la sélection ({len(choisis)})", disabled=not choisis, key="rec_supprimer"):
            self.prev.supprimer_receptions(choisis)
            st.rerun()

    def page_emballages(self):
        ui = self.ui
        ui.titre("Emballages", "Références, mapping SKU/client, stocks, couverture, tensions et délais fournisseur")
        o0, oR, o1, o2, o3 = st.tabs(["📋 Plan de réapprovisionnement", "📦 Réceptions d'emballages",
                                      "Stocks actuels & couverture", "Références & mapping", "Besoins des annonces futures"])
        with o0:
            self._plan_reappro()
        with oR:
            self._receptions_emballages()
        with o1:
            self._stocks_emballages()
        with o2:
            m = self.ctx.tables["mapping_emballages"].copy()
            c1, c2, c3 = st.columns(3)
            client = c1.selectbox("Client", ["Tous"] + list(self.noms_clients),
                                  format_func=lambda c: c if c == "Tous" else self.lib_client(c), key="map_client")
            article = c2.selectbox("Article", ["Tous"] + sorted(m["article_emballage"].unique()), key="map_article")
            sku = c3.selectbox("SKU", ["Tous"] + self.sku.index.tolist(),
                               format_func=lambda s: s if s == "Tous" else self.lib_sku(s), key="map_sku")
            if client != "Tous":
                m = m[m["client_id"].isna() | (m["client_id"] == client)]
            if article != "Tous":
                m = m[m["article_emballage"] == article]
            if sku != "Tous":
                m = m[m["sku_id"] == sku]
            if sku != "Tous" and client != "Tous":
                ui.afficher_badge("normal", f"Emballage proposé pour {sku} / {self.noms_clients[client]} : "
                                            f"{self.ctx.emballages.proposer(sku, client)} · étiquette "
                                            f"{self.ctx.emballages.etiquette(sku, client)}")
            m["client"] = m["client_id"].map(self.noms_clients).fillna("(règle générique)")
            m["défaut"] = m["est_defaut"].map({True: "★ oui", False: ""})
            ui.tableau(m[["mapping_id", "sku_id", "client", "article_emballage", "role", "priorite", "défaut"]], hauteur=420)
            st.caption("Règle SKU + client prioritaire sur le défaut SKU. Exigences clients : "
                       + " · ".join(f"{r.nom} : {r.exigence_emballage}" for r in self.ctx.tables["clients"].itertuples()))
        with o3:
            v = self.vue_futures(filtrer=True)
            if v.empty:
                self._aucune_prevision("emballages")
            else:
                ui.tableau(pd.DataFrame({
                    "Lot": v["lot_id"], "Arrivée": pd.to_datetime(v["heure_arrivee_prevue"]).dt.strftime("%d/%m %H:%M"),
                    "Client": v["client_id"].map(self.noms_clients), "SKU": v["sku_id"],
                    "Emballage": v["emballage_selectionne"],
                    "Modifié": v["emballage_modifie"].map({True: "✎ manuel", False: "auto"}),
                    "Besoin": v["besoin_emballage_unites"].map(ui.nombre) + " " + v["unite_emballage"].fillna(""),
                    "Stock client": v["stock_emballage_unites"].map(ui.nombre),
                    "État": v["statut_emballage"].map({"ok": "✓ ok", "risque": "⚠ risque de rupture",
                                                       "rupture": "✕ stock insuffisant"}),
                    "Motif": v["motif_emballage"].fillna(""),
                    "Réappro. conseillé": pd.to_datetime(v["date_reappro_conseillee"]).dt.strftime("%d/%m/%Y")}),
                    hauteur=400)
                st.caption("Une tension d'emballage est un avertissement distinct de la surcharge de capacité.")

    def _metriques_modeles(self):
        m = self.ctx.ia.metriques
        lignes = []
        for nom, libelle in (("rendement", "A — Rendement (kg/h)"), ("cycle", "B — Cycle total (min)")):
            if nom in m:
                for algo, sc in m[nom]["modeles"].items():
                    lignes.append({"Modèle": libelle, "Algorithme": algo, "MAE": sc["MAE"], "RMSE": sc["RMSE"],
                                   "R²": sc["R2"], "Retenu": "✓" if algo == m[nom]["meilleur_modele"] else ""})
                b = m[nom]["baseline_moyenne_sku"]
                lignes.append({"Modèle": libelle, "Algorithme": "Baseline (moyenne par SKU)", "MAE": b["MAE"],
                               "RMSE": b["RMSE"], "R²": b["R2"], "Retenu": ""})
        st.caption("Erreurs mesurées sur une période jamais vue pendant l'apprentissage (validation temporelle) : "
                   "MAE = erreur absolue moyenne (kg/h ou min), R² = part de la variabilité expliquée.")
        if lignes:
            self.ui.tableau(pd.DataFrame(lignes))
        if "surcharge" in m:
            sc = m["surcharge"]["scores"]
            st.caption(f"C — Surcharge : précision {sc['precision']}, rappel {sc['rappel']}, F1 {sc['F1']}, "
                       f"AUC {sc['ROC_AUC']} ({sc.get('validation', '')}).")
        if "rendement" in m:
            imp = pd.DataFrame(m["rendement"]["importances"][:10])
            if len(imp):
                fig = px.bar(imp, x="importance", y="variable", orientation="h")
                fig.update_traces(marker_color=Config.COULEURS["vert_clair"])
                fig.update_yaxes(autorange="reversed")
                self.ui.graphique(fig, 320, "Importance des variables — modèle de rendement")

    # ================================================================
    #  EMBALLAGES
    # ================================================================
    def _stocks_emballages(self):
        ui = self.ui
        s = self.f.stocks(self.ctx.variables.consommation_client)
        if ui.vide(s):
            return
        d = self.ctx.variables.etat_stocks(s, self.prev.receptions.charger())
        dernier = d["date_releve"].max()

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            ui.carte_kpi("Stock de référence", f"{dernier:%d/%m/%Y}", "", f"dernier relevé · {len(d)} lignes article × client", "colis")
        with k2:
            ui.carte_kpi("Articles critiques", ui.nombre((d["statut"] == "critique").sum()), "lignes",
                         "couverture < seuil", "alerte", statut="critique" if (d["statut"] == "critique").any() else "normal")
        with k3:
            sachets = d[d["categorie"] == "sachet_importe"]
            ui.carte_kpi("Sachets importés en tension", ui.nombre((sachets["statut"] != "normal").sum()), "lignes",
                         f"délai fournisseur jusqu'à {ui.nombre(sachets['delai_fournisseur_jours'].max())} jours", "camion")
        with k4:
            ui.carte_kpi("Couverture médiane", ui.nombre(d["couverture_j"].median(), 1), "jours",
                         "tous articles", "horloge", style="blanc")

        st.write("")
        unites = d["unite"].fillna("")
        ui.tableau(pd.DataFrame({
            "Article": d["article"], "Client": d["client_id"].map(self.noms_clients),
            "Stock": d["stock_actuel"].astype(int).astype(str) + " " + unites,
            "dont réceptions saisies": d["receptions_saisies"].map(ui.nombre),
            "kg supportés": d["kg_supportes"].round(0), "Conso. moyenne / jour": d["conso_moy_j"].round(1).astype(str) + " " + unites,
            "Couverture (jours)": d["couverture_j"], "Seuil (jours)": d["seuil_j"],
            "Délai fournisseur (jours)": d["delai_fournisseur_jours"], "Statut": d["statut"].map(ui.libelle)}), hauteur=420)
        with st.container(border=True):
            cov = d.groupby("article", as_index=False).agg(couverture_j=("couverture_j", "median"),
                                                           seuil_j=("seuil_j", "first"))
            fig = px.bar(cov, x="couverture_j", y="article", orientation="h",
                         labels={"couverture_j": "Couverture médiane (jours)", "article": ""})
            fig.update_traces(marker_color=Config.COULEURS["vert_clair"])
            fig.add_vline(x=7, line_dash="dot", line_color=Config.COULEURS["tension"], annotation_text="7 j")
            fig.add_vline(x=21, line_dash="dot", line_color=Config.COULEURS["critique"], annotation_text="21 j")
            ui.graphique(fig, 340, "Couverture actuelle par article (jours)")

    @staticmethod
    def _xlsx(feuilles: dict) -> bytes:
        tampon = io.BytesIO()
        with pd.ExcelWriter(tampon, engine="openpyxl") as ecrivain:
            for nom, df in feuilles.items():
                df.to_excel(ecrivain, sheet_name=nom[:31], index=False)
        return tampon.getvalue()

    # ================================================================
    #  PARAMÈTRES
    # ================================================================
    def page_parametres(self):
        ui, t = self.ui, self.ctx.tables
        ui.titre("Paramètres", "Hypothèses, seuils et paramètres centralisés — accès administrateur pour modifier")
        o1, o2, o4 = st.tabs(["Paramètres de calcul", "Table parametres (référence)", "Journal d'administration"])
        with o4:
            if Config.FICHIER_JOURNAL_ADMIN.exists():
                ui.tableau(pd.read_csv(Config.FICHIER_JOURNAL_ADMIN, encoding=Config.ENCODAGE).iloc[::-1], hauteur=360)
            else:
                st.info("Aucune action d'administration enregistrée.")
        with o1:
            ui.tableau(Config.tableau_parametres())
        with o2:
            ui.tableau(t["parametres"])
        with st.container(border=True):
            st.markdown("**Administration**")
            try:
                mot_de_passe = st.secrets.get("admin_password", Config.MOT_DE_PASSE_ADMIN)
            except Exception:
                mot_de_passe = Config.MOT_DE_PASSE_ADMIN
            if not st.session_state.get("admin_ok"):
                saisie = st.text_input("Mot de passe administrateur", type="password")
                if st.button("Se connecter"):
                    if saisie == mot_de_passe:
                        st.session_state["admin_ok"] = True
                        Config.journaliser_admin("connexion_admin")
                        st.rerun()
                    else:
                        st.error("Mot de passe incorrect.")
                return
            ui.afficher_badge("normal", "Connecté en administrateur")
            with st.form("parametres"):
                valeurs = {}
                c1, c2 = st.columns(2)
                for i, (cle, libelle) in enumerate(Config.PARAMETRES_MODIFIABLES.items()):
                    col = c1 if i % 2 == 0 else c2
                    actuel = getattr(Config, cle)
                    if cle == "SOURCE_CAPACITE_GOULOT":
                        valeurs[cle] = col.selectbox(libelle, ["cahier", "bdd"], index=0 if actuel == "cahier" else 1,
                                                     format_func=lambda s: {"cahier": "Cahier / Hypothèses (1 000 kg/h)",
                                                                            "bdd": "Table stations (1 300 kg/h)"}[s])
                    elif isinstance(actuel, bool):
                        valeurs[cle] = col.checkbox(libelle, value=actuel)
                    elif isinstance(actuel, int):
                        valeurs[cle] = int(col.number_input(libelle, value=actuel, step=1))
                    else:
                        valeurs[cle] = float(col.number_input(libelle, value=float(actuel), step=0.01, format="%.2f"))
                enregistrer = st.form_submit_button("Enregistrer les paramètres")
            if enregistrer:
                if valeurs["SEUIL_CHARGE_FORTE"] <= valeurs["SEUIL_CHARGE_SHIFT"]:
                    st.error("Le seuil de forte surcharge doit être supérieur au seuil de surcharge.")
                else:
                    Config.enregistrer_surcharges(valeurs)
                    st.cache_resource.clear()
                    st.success("Paramètres enregistrés. Pensez à réentraîner les modèles si la capacité a changé.")
                    st.rerun()
            c1, c2 = st.columns(2)
            if c1.button("Réinitialiser les valeurs par défaut"):
                Config.reinitialiser_surcharges()
                st.cache_resource.clear()
                st.success("Valeurs par défaut restaurées : redémarrez l'application pour les appliquer.")
            if c2.button("Se déconnecter"):
                st.session_state["admin_ok"] = False
                st.rerun()


# ======================================================================
#  15. APPLICATION STREAMLIT
# ======================================================================
def _construire_contexte() -> ContexteDonnees:
    return ContexteDonnees.construire(exporter=True)


if st is not None:
    _construire_contexte = st.cache_resource(show_spinner=False)(_construire_contexte)


class Application:
    """Orchestration : configuration de page, CSS, menu, filtres, cloche, page active, pied de page."""

    @staticmethod
    def _recalcul_si_necessaire(prev: "GestionnairePrevisions") -> None:
        """Recalcule les prédictions si le modèle ou les paramètres ont changé depuis le dernier calcul."""
        if not prev.ctx.ia.prets:
            return
        a, p = prev.annonces(), prev.predictions()
        if a.empty:
            return
        actives = p[p["statut_prediction"] != "ARCHIVEE"]
        manquantes = set(a["annonce_id"]) - set(p["annonce_id"])
        perimees = (actives["version_parametres"] != Config.version_parametres()).any() or \
                   (actives["version_modele"] != prev._version_modele()).any()
        # Prédictions créées par une version antérieure de l'application (nouvelles colonnes vides)
        anciennes = actives[["temps_production_min", "heure_debut_production", "heure_fin_predite"]].isna().any(axis=1).any()
        perimees = perimees or anciennes
        if manquantes or perimees:
            prev.recalculer()

    def executer(self) -> None:
        st.set_page_config(page_title=f"{Config.TITRE_APP} — Nature Growers", page_icon="🌿",
                           layout="wide", initial_sidebar_state="expanded")
        Config.charger_surcharges()
        ui = InterfaceUI()
        ui.injecter_css()
        page = ui.menu()
        try:
            with st.spinner("Chargement des données…"):
                ctx = _construire_contexte()
                if getattr(ctx, "version", None) != VERSION:      # ancienne version restée en cache
                    st.cache_resource.clear()
                    ctx = _construire_contexte()
        except ErreurDonnees as erreur:
            st.error(str(erreur))
            ui.pied_de_page()
            st.stop()
        except Exception as erreur:
            st.error(f"Erreur inattendue au chargement des données : {erreur}")
            with st.expander("Détails techniques"):
                st.exception(erreur)
            ui.pied_de_page()
            st.stop()
        filtres = ui.filtres(ctx)
        prev = GestionnairePrevisions(ctx)
        try:
            self._recalcul_si_necessaire(prev)
        except Exception as erreur:
            st.warning(f"Recalcul automatique des prédictions impossible : {erreur}")
        try:
            ui.panneau_notifications(prev.notifications)
        except Exception as erreur:
            st.sidebar.warning(f"Notifications indisponibles : {erreur}")
        pages = PagesApplication(ctx, ui, filtres)
        methode = {m[0]: m[2] for m in InterfaceUI.MENU}[page]
        try:
            getattr(pages, methode)()
        except Exception as erreur:
            st.error(f"Impossible d'afficher la page « {page} » : {erreur}")
            with st.expander("Détails techniques"):
                st.exception(erreur)
        ui.pied_de_page()


# ======================================================================
#  16. VALIDATION CONSOLE (MODE "valider")
# ======================================================================
class ValidationConsole:
    """Vérifie les critères de recette et affiche un résumé dans IDLE."""

    def __init__(self):
        self.resultats = []

    def _ok(self, libelle, condition, detail=""):
        self.resultats.append(bool(condition))
        print(f"  [{'OK ' if condition else 'KO '}] {libelle}{' — ' + detail if detail else ''}")

    def executer(self) -> bool:
        print("=" * 72)
        print(f"  {Config.TITRE_APP} — VALIDATION DE LA BASE ({Config.FICHIER_BDD.name})")
        print("=" * 72)
        if not BIBLIOS_DONNEES_OK:
            print("Bibliothèques manquantes : mettez MODE = \"installer\" puis F5.")
            return False
        try:
            debut = time.time()
            ctx = ContexteDonnees.construire(exporter=True, classeur=True)
        except ErreurDonnees as erreur:
            print(f"ERREUR : {erreur}")
            return False
        t, q = ctx.tables, ctx.qualite
        print(f"\nChargement + calculs : {time.time() - debut:.1f} s\n\n1. Volumes")
        for nom, attendu in Config.VOLUMES_ATTENDUS.items():
            self._ok(f"{nom:<27}", len(t[nom]) == attendu, f"{len(t[nom])} lignes (attendu {attendu})")
        self._ok("machines_assemblage générée", len(t["machines_assemblage"]) == 3)
        self._ok("mapping_emballages généré", len(t["mapping_emballages"]) > 0, f"{len(t['mapping_emballages'])} règles")
        print("\n2. Fichiers data_clean/")
        for nom in Config.TABLES_CAHIER + ["anomalies"]:
            chemin = Config.DATA_CLEAN_DIR / f"{nom}.csv"
            self._ok(f"{nom}.csv", chemin.exists())
        print("\n3. Intégrité")
        orphelins = int(ctx.rapport_cles.loc[ctx.rapport_cles["type"] == "clé étrangère", "nb_problèmes"].sum())
        doublons = int(ctx.rapport_cles.loc[ctx.rapport_cles["type"] == "clé primaire", "nb_problèmes"].sum())
        self._ok("0 clé étrangère orpheline", orphelins == 0, f"{orphelins} orphelin(s)")
        self._ok("0 clé primaire en double", doublons == 0, f"{doublons} doublon(s)")
        chrono = q.taux_cycles_chronologiques()
        self._ok("100 % des cycles chronologiques", chrono == 1.0, f"{chrono * 100:.1f} %")
        print("\n4. Indicateurs métier")
        taux_a = q.taux_anomalies()
        self._ok("Anomalies ≈ 1 % des lots", 0.003 <= taux_a <= 0.02, f"{taux_a * 100:.2f} % ({len(q.anomalies)} anomalies tracées)")
        ts = ctx.capacite.taux_annonces_surcharge()
        lo, hi = Config.CIBLE_TAUX_SURCHARGE
        self._ok("Annonces en surcharge 10–15 %", lo <= ts <= hi,
                 f"{ts * 100:.2f} % (goulot {Config.capacite_goulot_kg_h():.0f} kg/h)")
        f = ctx.derive.fenetres.set_index("fenêtre")
        self._ok("Dérive détectée (05/01 → 15/02/2026)", ctx.derive.derive_detectee(),
                 f"{f.at['Baisse', 'dérive_moyenne_pct']:+.2f} % vs référence, p = {ctx.derive.tests.iloc[0]['p_value']:.4f}")
        self._ok("Redressement détecté", ctx.derive.redressement_detecte(),
                 f"après redressement {f.at['Après redressement', 'dérive_moyenne_pct']:+.2f} %")
        print("\n5. Intelligence artificielle")
        if ctx.ia.prets:
            m = ctx.ia.metriques
            for nom in ("rendement", "cycle"):
                if nom in m:
                    best = m[nom]["meilleur_modele"]
                    sc = m[nom]["modeles"][best]
                    print(f"  {nom:<10} {best:<17} MAE={sc['MAE']}  RMSE={sc['RMSE']}  R²={sc['R2']}  "
                          f"(gain vs baseline {m[nom]['gain_vs_baseline_pct']} %)")
            try:
                r = ctx.ia.predire_lot(1500, "SKU004", "CLI001", "NG3", "2026-02-10", "jour", 10, journaliser=False)
                self._ok("predire_lot() fonctionne", True,
                         f"cycle {r['duree_cycle_prevue_min']:.0f} min, expédition {r['heure_expedition_prevue'][11:]}, "
                         f"charge {r['taux_utilisation'] * 100:.0f} %")
            except Exception as erreur:
                self._ok("predire_lot() fonctionne", False, str(erreur))
        else:
            print("  Modèles absents : mettez MODE = \"entrainer\" puis F5.")
        print("\n6. Tests data du cahier V2 (§21.1)")
        for _, l in ControlesRecette.tests_data(ctx).iterrows():
            self._ok(l["contrôle"], l["ok"], l["détail"])
        print("\n7. Tests fonctionnels du cahier V2 (§21.2) — dossier temporaire")
        for _, l in ControlesRecette.tests_fonctionnels(ctx).iterrows():
            self._ok(l["test"], l["ok"], l["détail"])
        print(f"\nClasseur complet (une feuille par table) : {Config.FICHIER_CLASSEUR}")
        print("\n" + "=" * 72)
        print(f"  RÉSULTAT : {sum(self.resultats)}/{len(self.resultats)} contrôles réussis")
        print("=" * 72)
        return all(self.resultats)


# ======================================================================
#  17. LANCEUR IDLE (installation + lancement de Streamlit)
# ======================================================================
class LanceurIDLE:
    """Permet de tout faire depuis IDLE avec F5, sans ligne de commande."""

    PAQUETS = ["pandas>=2.2.3", "numpy>=2.1.0", "openpyxl>=3.1.5", "scikit-learn>=1.6.0", "scipy>=1.14.1",
               "joblib>=1.4.2", "pyarrow>=18.0.0", "plotly>=5.24.0", "streamlit>=1.42.0", "flask>=3.0"]
    MODULES = {"pandas": "pandas", "numpy": "numpy", "openpyxl": "openpyxl", "sklearn": "scikit-learn",
               "scipy": "scipy", "joblib": "joblib", "plotly": "plotly", "streamlit": "streamlit"}

    @staticmethod
    def interpreteur() -> str:
        """IDLE tourne souvent sous pythonw.exe (sans console) : on préfère python.exe."""
        exe = Path(sys.executable)
        if exe.name.lower() == "pythonw.exe" and exe.with_name("python.exe").exists():
            return str(exe.with_name("python.exe"))
        return str(exe)

    def manquants(self) -> list:
        return [pip for module, pip in self.MODULES.items() if importlib.util.find_spec(module) is None]

    def installer(self) -> bool:
        print("Installation des bibliothèques (plusieurs minutes la première fois)…")
        subprocess.run([self.interpreteur(), "-m", "pip", "install", "--upgrade", "pip"])
        code = subprocess.run([self.interpreteur(), "-m", "pip", "install", *self.PAQUETS]).returncode
        if code == 0 and not self.manquants():
            print("\nInstallation terminée. Mettez MODE = \"valider\" puis F5.")
            return True
        print("\nL'installation a échoué. Vérifiez votre connexion internet puis relancez.")
        return False

    @staticmethod
    def port_ouvert(port: int) -> bool:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            return s.connect_ex(("127.0.0.1", port)) == 0

    def arreter_ancien_serveur(self, port: int) -> bool:
        """Arrête le serveur Streamlit déjà lancé sur ce port (ancienne version du code)."""
        if not self.port_ouvert(port):
            return True
        print(f"Un ancien serveur tourne sur le port {port} : arrêt pour charger la nouvelle version…")
        try:
            if os.name == "nt":
                sortie = subprocess.run(["netstat", "-ano", "-p", "TCP"], capture_output=True, text=True,
                                        errors="ignore").stdout
                pids = set()
                for ligne in sortie.splitlines():
                    morceaux = ligne.split()
                    # Colonnes : Proto  Adresse locale  Adresse distante  État  PID
                    if len(morceaux) >= 5 and morceaux[1].endswith(f":{port}") \
                            and morceaux[3].upper() in ("LISTENING", "ÉCOUTE", "ECOUTE"):
                        pids.add(morceaux[-1])
                for pid in pids:
                    if pid.isdigit() and int(pid) != os.getpid():
                        subprocess.run(["taskkill", "/PID", pid, "/T", "/F"], capture_output=True)
            else:
                subprocess.run(["fuser", "-k", f"{port}/tcp"], capture_output=True)
        except OSError:
            pass
        for _ in range(10):
            if not self.port_ouvert(port):
                return True
            time.sleep(0.5)
        return False

    def lancer(self) -> None:
        print("=" * 66)
        print(f"  {Config.TITRE_APP} — lancement de l'application")
        print(f"  Version : {VERSION}")
        print(f"  Fichier : {_fichier_courant()}")
        print("=" * 66)
        manquants = self.manquants()
        if manquants:
            print("Bibliothèques manquantes :", ", ".join(manquants))
            print("→ Mettez MODE = \"installer\" en haut du fichier puis appuyez sur F5.")
            return
        if not Config.FICHIER_BDD.exists():
            print(f"ATTENTION : base introuvable. Chemin attendu : {Config.FICHIER_BDD}")
        port = Config.PORT_STREAMLIT
        if not self.arreter_ancien_serveur(port):
            # Impossible d'arrêter l'ancien serveur : on prend un port libre
            port = next((p for p in range(port + 1, port + 20) if not self.port_ouvert(p)), port + 1)
            print(f"Ancien serveur impossible à arrêter : nouveau port {port}.")
        url = f"http://localhost:{port}"
        try:
            Config.ecrire_theme()
        except OSError:
            pass
        commande = [self.interpreteur(), "-m", "streamlit", "run", str(_fichier_courant()),
                    "--server.port", str(port), "--server.headless", "true",
                    "--browser.gatherUsageStats", "false"]
        for cle, valeur in Config.THEME.items():
            commande += [f"--theme.{cle}", valeur]
        options = {"creationflags": subprocess.CREATE_NEW_CONSOLE} if os.name == "nt" else {}
        processus = subprocess.Popen(commande, cwd=str(Config.BASE_DIR), **options)
        print("Démarrage du serveur Streamlit…")
        for _ in range(60):
            if processus.poll() is not None:
                print("ERREUR : Streamlit s'est arrêté. Consultez la fenêtre de console qui s'est ouverte.")
                return
            if self.port_ouvert(port):
                print(f"Application disponible : {url}  (version {VERSION})")
                webbrowser.open(url)
                print("Pour arrêter l'application : fermez la fenêtre de console noire.")
                return
            time.sleep(1)
        print(f"Le serveur ne répond pas. Ouvrez manuellement {url}")


# ======================================================================
#  18. POINT D'ENTRÉE
# ======================================================================
def _execute_par_streamlit() -> bool:
    """Vrai si ce fichier est exécuté par « streamlit run » (et non par IDLE)."""
    if st is None:
        return False
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx(suppress_warning=True) is not None
    except TypeError:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False



# ======================================================================
#  API WEB (Flask) — déploiement Cloud (Render) et futur lien avec Fusion AI
# ======================================================================
#  Routes (JSON) :
#    GET  /                           → présentation de l'API
#    GET  /api/sante                  → état (chargement, modèles prêts, version)
#    GET  /api/referentiels           → fermes, SKU, clients, emballages compatibles
#    POST /api/predire                → prédiction d'un lot (sans l'enregistrer)
#    GET  /api/annonces               → annonces futures et leurs prédictions
#    POST /api/annonces               → enregistre une ou plusieurs annonces (+ prédiction)
#    POST /api/annonces/<id>/statut   → VALIDEE ou ARCHIVEE
#    GET  /api/capacite?date=&shift=  → charge, capacité et prochain batch conseillé d'un shift
#    GET  /api/reapprovisionnement    → plan de réapprovisionnement des emballages
#    GET  /api/notifications          → alertes de la cloche
#  Sécurité : si la variable d'environnement PACKFLOW_API_KEY est définie, toutes les
#  routes /api/* (sauf /api/sante) exigent l'en-tête « X-API-Key: <clé> ».
#  Au premier démarrage sur le Cloud, les modèles sont entraînés automatiquement (≈ 1 à 3 min).
# ----------------------------------------------------------------------
class ServiceAPI:
    """Moteur PackFlow chargé une seule fois en mémoire et partagé par les requêtes."""

    def __init__(self):
        import threading
        self.verrou = threading.Lock()
        self.ctx = None
        self.etat = "en attente"
        self.erreur = None

    def charger(self) -> "ContexteDonnees":
        """Charge les données et les modèles (entraînés automatiquement s'ils sont absents)."""
        with self.verrou:
            if self.ctx is not None:
                return self.ctx
            try:
                self.etat = "chargement des données"
                ctx = ContexteDonnees.construire(exporter=False)
                if not ctx.ia.prets:
                    self.etat = "entraînement des modèles (premier démarrage)"
                    print("[API] Modèles absents : entraînement automatique…", flush=True)
                    ctx.ia.entrainer(journal=lambda *m: print("[API]", *m, flush=True))
                self.ctx, self.etat, self.erreur = ctx, "prêt", None
            except Exception as erreur:
                self.etat, self.erreur = "erreur", str(erreur)
                raise
            return self.ctx

    def prechauffer(self) -> None:
        """Charge le moteur en arrière-plan pour que le serveur réponde tout de suite."""
        import threading

        def tache():
            try:
                self.charger()
                print("[API] Moteur PackFlow prêt.", flush=True)
            except Exception as erreur:
                print(f"[API] Échec du chargement : {erreur}", flush=True)
        threading.Thread(target=tache, daemon=True).start()


def _json_propre(valeur):
    """Rend JSON-compatible : NaN → null, dates → ISO 8601, types numpy → Python."""
    if isinstance(valeur, dict):
        return {str(k): _json_propre(v) for k, v in valeur.items()}
    if isinstance(valeur, (list, tuple, set)):
        return [_json_propre(v) for v in valeur]
    if isinstance(valeur, pd.DataFrame):
        return json.loads(valeur.to_json(orient="records", date_format="iso", force_ascii=False))
    if isinstance(valeur, (pd.Timestamp, datetime, date)):
        return None if pd.isna(valeur) else valeur.isoformat()
    if valeur is pd.NaT:
        return None
    if hasattr(valeur, "item") and not isinstance(valeur, (str, bytes)):
        valeur = valeur.item()
    if isinstance(valeur, float) and (math.isnan(valeur) or math.isinf(valeur)):
        return None
    return valeur


def creer_app():
    """Application Flask (variable « app » utilisée par gunicorn sur Render)."""
    from flask import Flask, jsonify, request

    application = Flask(__name__)
    application.json.ensure_ascii = False
    service = ServiceAPI()
    application.config["SERVICE"] = service
    cle_attendue = os.environ.get("PACKFLOW_API_KEY", "").strip()

    def reponse(contenu, code=200):
        return jsonify(_json_propre(contenu)), code

    @application.before_request
    def securite():
        chemin = request.path.rstrip("/")
        if chemin.startswith("/api") and chemin != "/api/sante" and cle_attendue:
            if request.headers.get("X-API-Key", "") != cle_attendue:
                return reponse({"erreur": "clé API invalide ou absente (en-tête X-API-Key)"}, 401)
        return None

    def moteur():
        if service.ctx is None and service.etat.startswith(("chargement", "entraînement")):
            return None
        return service.charger()

    def pas_pret():
        return reponse({"erreur": "le moteur PackFlow démarre, réessayez dans une minute", "etat": service.etat}, 503)

    @application.errorhandler(Exception)
    def erreur_generale(erreur):
        from werkzeug.exceptions import HTTPException
        if isinstance(erreur, HTTPException):
            return reponse({"erreur": erreur.description}, erreur.code)
        return reponse({"erreur": str(erreur)}, 500)

    @application.get("/")
    def accueil():
        return reponse({"application": Config.TITRE_APP, "version": VERSION, "etat": service.etat,
                        "routes": ["/api/sante", "/api/referentiels", "/api/predire (POST)", "/api/annonces (GET/POST)",
                                   "/api/annonces/<id>/statut (POST)", "/api/capacite", "/api/reapprovisionnement",
                                   "/api/notifications"],
                        "note": Config.PIED_DE_PAGE})

    @application.get("/api/sante")
    def sante():
        pret = service.ctx is not None and service.ctx.ia.prets
        return reponse({"statut": "ok", "etat": service.etat, "modeles_prets": bool(pret), "erreur": service.erreur,
                        "version": VERSION, "securite": "clé API exigée" if cle_attendue else "aucune clé configurée",
                        "heure": datetime.now().isoformat(timespec="seconds")})

    @application.get("/api/referentiels")
    def referentiels():
        ctx = moteur()
        if ctx is None:
            return pas_pret()
        t = ctx.tables
        emballages = {f"{s}|{c}": ctx.emballages.compatibles(s, c)
                      for s in t["produits_sku"]["sku_id"] for c in t["clients"]["client_id"]}
        return reponse({"fermes": t["fermes"][["ferme_id", "distance_usine_km"]],
                        "produits": t["produits_sku"][["sku_id", "produit", "famille", "format", "poids_unitaire_g"]],
                        "clients": t["clients"][["client_id", "nom", "pays", "heure_limite_depart"]],
                        "priorites": Config.PRIORITES, "shifts": Config.SHIFTS,
                        "emballages_compatibles": emballages})

    @application.post("/api/predire")
    def predire():
        ctx = moteur()
        if ctx is None:
            return pas_pret()
        d = request.get_json(silent=True) or {}
        manquants = [c for c in ("volume_kg", "sku_id", "client_id", "ferme_id", "date") if c not in d]
        if manquants:
            return reponse({"erreur": "champs obligatoires manquants", "champs": manquants}, 400)
        shift = d.get("shift") or "jour"
        prev = GestionnairePrevisions(ctx)
        futures = prev.vue()
        externe = futures[["date_shift", "shift", "volume_kg"]].dropna() if len(futures) else None
        r = ctx.ia.predire_lot(float(d["volume_kg"]), d["sku_id"], d["client_id"], d["ferme_id"], d["date"], shift,
                               int(d.get("nb_operateurs") or Config.OPERATEURS_DEFAUT.get(shift, 10)),
                               d.get("priorite", "NORMALE"), int(d.get("absences") or 0),
                               heure_arrivee=d.get("heure_arrivee"), charge_externe=externe, journaliser=False)
        return reponse(r)

    @application.get("/api/annonces")
    def lister_annonces():
        ctx = moteur()
        if ctx is None:
            return pas_pret()
        v = GestionnairePrevisions(ctx).vue()
        if len(v) and request.args.get("archivees", "non") != "oui":
            v = v[v["statut_prediction"] != "ARCHIVEE"]
        jour = request.args.get("date")
        if jour and len(v):
            v = v[pd.to_datetime(v["date_arrivee"]).dt.date == pd.Timestamp(jour).date()]
        return reponse({"nombre": len(v), "annonces": v})

    @application.post("/api/annonces")
    def creer_annonces():
        ctx = moteur()
        if ctx is None:
            return pas_pret()
        d = request.get_json(silent=True) or {}
        lignes = d.get("annonces") if isinstance(d.get("annonces"), list) else [d]
        prev = GestionnairePrevisions(ctx)
        valides, erreurs = prev.valider(pd.DataFrame(lignes), origine_defaut=str(d.get("origine", "IMPORT")).upper())
        if len(erreurs):
            return reponse({"erreur": "annonces invalides", "details": erreurs}, 400)
        creees = prev.ajouter(valides)
        return reponse({"nombre": len(creees), "annonces": creees}, 201)

    @application.post("/api/annonces/<annonce_id>/statut")
    def changer_statut(annonce_id):
        ctx = moteur()
        if ctx is None:
            return pas_pret()
        statut = str((request.get_json(silent=True) or {}).get("statut", "")).upper()
        if statut not in ("VALIDEE", "ARCHIVEE"):
            return reponse({"erreur": "statut attendu : VALIDEE ou ARCHIVEE"}, 400)
        prev = GestionnairePrevisions(ctx)
        if annonce_id not in set(prev.annonces()["annonce_id"]):
            return reponse({"erreur": f"annonce inconnue : {annonce_id}"}, 404)
        prev.changer_statut([annonce_id], statut)
        return reponse({"annonce_id": annonce_id, "statut": statut})

    @application.get("/api/capacite")
    def capacite():
        ctx = moteur()
        if ctx is None:
            return pas_pret()
        jour = request.args.get("date") or date.today().isoformat()
        shift = request.args.get("shift", "jour")
        if shift not in Config.SHIFTS:
            return reponse({"erreur": "shift attendu : jour ou nuit"}, 400)
        info = GestionnairePrevisions(ctx).prochain_batch(jour, shift)
        info["fermes"] = info["fermes"].rename(columns={"depart_conseille": "depart_conseille_de_la_ferme"})
        if info.get("suivant"):
            info["suivant"] = {k: v for k, v in info["suivant"].items() if k not in ("fermes", "suivant")}
        return reponse(info)

    @application.get("/api/reapprovisionnement")
    def reapprovisionnement():
        ctx = moteur()
        if ctx is None:
            return pas_pret()
        horizon = int(request.args.get("horizon", Config.HORIZON_PLAN_J))
        fond = request.args.get("fond", "non") == "oui"
        resume, _ = GestionnairePrevisions(ctx).plan_reappro(horizon, fond)
        return reponse({"horizon_jours": horizon, "consommation_de_fond": fond, "articles": resume})

    @application.get("/api/notifications")
    def notifications():
        ctx = moteur()
        if ctx is None:
            return pas_pret()
        n = GestionnairePrevisions(ctx).notifications.charger()
        return reponse({"non_lues": int((~n["lu"]).sum()) if len(n) else 0, "notifications": n})

    if os.environ.get("PORT"):             # sur le Cloud : chargement immédiat en arrière-plan
        service.prechauffer()
    return application


try:
    app = creer_app() if BIBLIOS_DONNEES_OK else None
except ImportError:                         # Flask absent (usage local sans API) : l'application Streamlit reste utilisable
    app = None


def lancer_api_locale() -> None:
    """Lance l'API Flask sur le port fourni par l'hébergeur (variable PORT), 8000 par défaut en local."""
    if app is None:
        print("Flask n'est pas installé : py -m pip install flask (ou MODE = \"installer\").")
        return
    port = int(os.environ.get("PORT", 8000))
    app.config["SERVICE"].prechauffer()
    print(f"API PackFlow sur http://127.0.0.1:{port}/api/sante  (fermez la fenêtre pour arrêter)")
    app.run(host="0.0.0.0", port=port)


def main() -> None:
    if _execute_par_streamlit():
        Application().executer()
        return
    mode = str(MODE).strip().lower()
    if mode == "installer":
        LanceurIDLE().installer()
    elif not BIBLIOS_DONNEES_OK:
        print("Bibliothèques manquantes : mettez MODE = \"installer\" puis appuyez sur F5.")
    elif mode == "valider":
        ValidationConsole().executer()
    elif mode == "entrainer":
        print("=" * 72)
        print(f"  {Config.TITRE_APP} — ENTRAÎNEMENT DE L'IA")
        print("=" * 72)
        try:
            ctx = ContexteDonnees.construire(exporter=True)
            debut = time.time()
            ctx.ia.entrainer()
            print(f"Durée : {time.time() - debut:.0f} s — métriques : {Config.FICHIER_METRIQUES}")
        except ErreurDonnees as erreur:
            print(f"ERREUR : {erreur}")
    elif mode == "application":
        LanceurIDLE().lancer()
    elif mode == "api":
        lancer_api_locale()
    else:
        print(f"MODE inconnu : « {MODE} ». Valeurs possibles : installer, valider, entrainer, application, api.")


if __name__ == '__main__':
    if '--entrainer' in sys.argv:
        # Utilisé pendant la construction (build) sur Render : les modèles sont entraînés une fois
        # avec la version de scikit-learn installée, puis simplement chargés à chaque démarrage.
        _ctx = ContexteDonnees.construire(exporter=False)
        _ctx.ia.entrainer()
        sys.exit(0)
    if os.environ.get('PORT') or '--api' in sys.argv:
        # Hébergeur Cloud (Render) : récupère le port attribué (8000 par défaut en local)
        port = int(os.environ.get('PORT', 8000))
        if not os.environ.get('PORT'):
            app.config["SERVICE"].prechauffer()     # sur le Cloud, creer_app() l'a déjà lancé
        app.run(host='0.0.0.0', port=port)
    else:
        main()   # utilisation locale depuis IDLE (F5) : MODE en haut du fichier
