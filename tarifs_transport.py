"""
Module de gestion et de tarification du transport de marchandises Tradifoot.
Source des tarifs : Grille officielle GHERNAOUTI Walid (Dely Ibrahim, Chéraga - Alger).
Téléphone : 0550 40 55 02 | R.C : 16/00 - 4933691 A 13.
"""

import json
import os
import re
import unicodedata

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(CURRENT_DIR, "tarifs_transport.json")

# Chargement de la base JSON si disponible
_TARIFS_DATA = None

def _load_data():
    global _TARIFS_DATA
    if _TARIFS_DATA is None and os.path.exists(JSON_PATH):
        with open(JSON_PATH, "r", encoding="utf-8") as f:
            _TARIFS_DATA = json.load(f)
    return _TARIFS_DATA

def normalize_key(text: str) -> str:
    """Normalise un nom de wilaya ou ville (majuscules, sans accents, sans ponctuation)."""
    if not text:
        return ""
    text = unicodedata.normalize("NFD", str(text))
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = text.upper().strip()
    # Remplacer toute ponctuation par espace ou rien
    text = re.sub(r"[.\-_'/\\()]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

# Dictionnaire complet des tarifs H.T avec alias pour tolérance orthographique
TARIFS_RAW = {
    # Proche banlieue & Centre
    "BLIDA": 6000,
    "BOUMERDES": 7000,
    "TIPAZA": 8000,
    "MEDEA": 9000,
    "BOUIRA": 10000,
    "TIZI OUZOU": 10000,
    "AIN DEFLA": 11000,
    "CHLEF": 12000,

    # Centre-Est & Centre-Ouest
    "B.B.A": 15000,
    "BBA": 15000,
    "B B A": 15000,
    "BORDJ BOU ARRERIDJ": 15000,
    "BORDJ BOU ARRERIDJ": 15000,
    "TISSEMSILT": 15000,
    "BEJAIA": 15000,
    "M SILA": 15000,
    "MSILA": 15000,
    "M'SILA": 15000,
    "SETIF": 17000,
    "RELIZANE": 17000,
    "DJELFA": 17000,
    "EL EULMA": 18000,
    "TAJNANTE": 19000,
    "TADJENANET": 19000,
    "MOSTAGANEM": 19000,

    # Ouest & Est intérieur
    "TIARET": 22000,
    "JIJEL": 22000,
    "MASCARA": 22000,
    "ORAN": 22000,
    "CONSTANTINE": 22000,
    "SIDI B.ABESS": 24000,
    "SIDI B ABESS": 24000,
    "SIDI BEL ABBES": 24000,
    "SIDI BEL ABBÈS": 24000,
    "SBA": 24000,
    "LAGHOUAT": 24000,
    "BATNA": 25000,
    "MILA": 25000,
    "SAIDA": 28000,
    "AIN TMOUCHENT": 28000,
    "AIN TEMOUCHENT": 28000,
    "AÏN TÉMOUCHENT": 28000,
    "TLEMCEN": 28000,
    "OUM EL BOUAGHI": 28000,
    "OEB": 28000,
    "SKIKDA": 28000,

    # Sud & Est lointain
    "OULED DJELAL": 30000,
    "OULED DJELLAL": 30000,
    "BISKRA": 30000,
    "GUELMA": 32000,
    "KHENCHLA": 32000,
    "KHENCHELA": 32000,
    "SOUK AHRASS": 32000,
    "SOUK AHRAS": 32000,
    "MAGHNIA": 32000,
    "EL BEYADH": 32000,
    "EL BAYADH": 32000,
    "EL MEGHAIR": 32000,
    "EL M'GHAIR": 32000,
    "ANNABA": 32000,
    "TEBESSA": 35000,
    "GHARDAIA": 35000,
    "EL TARAF": 36000,
    "EL TARF": 36000,
    "NAAMA": 38000,
    "TOUGGOURT": 38000,
    "EL OUED": 40000,
    "OURGLA": 48000,
    "OUARGLA": 48000,
    "HASSI MESOUD": 55000,
    "HASSI MESSAOUD": 55000,
    "EL MENIA": 55000,
    "EL MENIAA": 55000,

    # Sud-Ouest & Grand Sud
    "BECHAR": 60000,
    "BENI ABBESE": 85000,
    "BENI ABBES": 85000,
    "BÉNI ABBÈS": 85000,
    "TIMIMOUNE": 85000,
    "TIMIMOUN": 85000,
    "IN SALAH": 90000,
    "ADRAR": 100000,
    "AIN AMINAS": 115000,
    "IN AMENAS": 115000,
    "TINDOUF": 125000,
    "ILIZI": 130000,
    "ILLIZI": 130000,
    "TAMANRASSET": 140000,
    "BORDJ BADJI M": 180000,
    "BORDJ BADJI MOKHTAR": 180000,
    "BBM": 180000,
    "DJANET": 180000,
    "IN GUIZZAM": 185000,
    "IN GUEZZAM": 185000,
}

# Clés normalisées
TARIFS_HT = {normalize_key(k): v for k, v in TARIFS_RAW.items()}

COUT_POINT_DANS_WILAYA = 2000
COUT_POINT_HORS_WILAYA = 4000

def get_transport_tariff(destination: str) -> int:
    """
    Retourne le tarif H.T en DZD pour une wilaya ou ville donnée.
    Renvoie None si la destination n'est pas trouvée.
    """
    if not destination:
        return None
    norm = normalize_key(destination)
    if norm in TARIFS_HT:
        return TARIFS_HT[norm]
    # Recherche partielle tolérante
    for key, price in TARIFS_HT.items():
        if key in norm or norm in key:
            return price
    return None

def calculate_trip_cost(destination: str, points_in_wilaya: int = 0, points_out_wilaya: int = 0, has_return: bool = False) -> dict:
    """
    Calcule le coût total d'une course de transport selon la grille officielle :
    - Tarif de base H.T selon la destination
    - Points supplémentaires dans wilaya : 2 000 DZD / point
    - Points supplémentaires hors wilaya : 4 000 DZD / point
    - Retour de marchandises : 100% du tarif de base
    """
    base_tariff = get_transport_tariff(destination)
    if base_tariff is None:
        raise ValueError(f"Destination inconnue : {destination}")

    cost_in_wilaya = max(0, points_in_wilaya) * COUT_POINT_DANS_WILAYA
    cost_out_wilaya = max(0, points_out_wilaya) * COUT_POINT_HORS_WILAYA
    cost_return = base_tariff if has_return else 0

    total_ht = base_tariff + cost_in_wilaya + cost_out_wilaya + cost_return
    tva = int(total_ht * 0.19)
    total_ttc = total_ht + tva

    return {
        "destination": destination,
        "tarif_base_ht": base_tariff,
        "points_dans_wilaya": points_in_wilaya,
        "cout_points_dans_wilaya": cost_in_wilaya,
        "points_hors_wilaya": points_out_wilaya,
        "cout_points_hors_wilaya": cost_out_wilaya,
        "retour_marchandises": has_return,
        "cout_retour_marchandises": cost_return,
        "total_ht": total_ht,
        "tva_19": tva,
        "total_ttc": total_ttc
    }

if __name__ == "__main__":
    print("=== TEST MODULE TARIFS TRANSPORT ===")
    test_destinations = ["BATNA", "SETIF", "ORAN", "HASSI MESOUD", "TADJENANET", "B.B.A", "IN GUIZZAM"]
    for d in test_destinations:
        t = get_transport_tariff(d)
        print(f"Tarif {d:<15} : {t:>7} DZD H.T")
    
    sim = calculate_trip_cost("BATNA", points_in_wilaya=2, points_out_wilaya=1, has_return=False)
    print("\nExemple simulation course BATNA (2 arrêts intra, 1 extra) :")
    print(f"  Base H.T : {sim['tarif_base_ht']} DZD")
    print(f"  Arrêts intra (+2) : {sim['cout_points_dans_wilaya']} DZD")
    print(f"  Arrêts extra (+1) : {sim['cout_points_hors_wilaya']} DZD")
    print(f"  TOTAL H.T : {sim['total_ht']} DZD")
    print(f"  TOTAL T.T.C (TVA 19%) : {sim['total_ttc']} DZD")
