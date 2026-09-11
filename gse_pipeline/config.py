"""Configuration for the current CMMS EDR cleaning workflow."""

from pathlib import Path

try:
    import config_local
except ImportError as exc:
    raise RuntimeError(
        "Missing config_local.py. Copy config_local.example.py to config_local.py."
    ) from exc

ROOT_DIR = Path(__file__).resolve().parent.parent

# SOURCE: this synced SharePoint path is local to the machine. If the workbook
# is copied beside the repository, comment this line and use the line below.
SOURCE_FILE = Path(config_local.SHAREPOINT_SOURCE_FILE)
# SOURCE_FILE = ROOT_DIR / "EDR (JAN-AUG) 2026.xls"

# OUTPUT: SharePoint publishing is active after local review.
OUTPUT_FILE = ROOT_DIR / "EDR_JANAUG_2026_CLEANED.xlsx"
# OUTPUT_FILE = Path(config_local.SHAREPOINT_OUTPUT_FILE)

# Display names for categorical values in the cleaned workbook. Matching is
# case-insensitive and limited to complete field values.
WORK_STATUS_MAP = {
    "waiting parts and ror": "Waiting For Spare",
    "flm h/o to workshop": "FLM H/O to Workshop",
}

WORK_ORDER_TRADE_MAP = {
    "welding shop": "Welding Shop",
    "paint shop": "Paint Shop",
    "jack - tow bar": "Jack - Tow Bar",
    "component overhaul": "Component Overhaul",
    "component electrical": "Component Electrical",
    "house keeping": "House Keeping",
    "electrical support": "Electrical Support",
    "foreman approval": "Foreman Approval",
    "tool crib": "Tool Crib",
    "material mgmnt": "Material Management",
}

# Known spelling variants are excluded from rare-word suggestions.
SPELLING_MAP = {
    "NEGLIGENE": "NEGLIGENCE", "NEGLEGENCE": "NEGLIGENCE", "NEGLGENCE": "NEGLIGENCE",
    "NELIGENCE": "NEGLIGENCE", "NGELIGENCE": "NEGLIGENCE", "NGLIGENCE": "NEGLIGENCE",
    "NEGELIGENCE": "NEGLIGENCE", "NEGLIGEIGENCE": "NEGLIGENCE", "NEGLIGNECE": "NEGLIGENCE",
    "NIGLIGENCE": "NEGLIGENCE", "ENGLIGENCE": "NEGLIGENCE", "NEGIGENCE": "NEGLIGENCE",
    "NEEGLIGENCE": "NEGLIGENCE", "NEELIGENCE": "NEGLIGENCE", "NEGILGENCE": "NEGLIGENCE",
    "NEGLIENCE": "NEGLIGENCE", "NEGLIGENGE": "NEGLIGENCE", "NEGLIGNCE": "NEGLIGENCE",
    "NEGLIIGENCE": "NEGLIGENCE", "NEGLOGENCE": "NEGLIGENCE", "NEGLUGENCE": "NEGLIGENCE",
    "NEHGLIGENCE": "NEGLIGENCE",
    "YTRE": "TYRE", "TYR": "TYRE", "TYE": "TYRE",
    "PUCNTURED": "PUNCTURED", "PUNCUTRED": "PUNCTURED", "PPUNCTURED": "PUNCTURED", "PUNCTURD": "PUNCTURED",
    "COMVEYOR": "CONVEYOR", "CONVOYER": "CONVEYOR", "CONEYOR": "CONVEYOR", "CONVER": "CONVEYOR", "CONVEYPOR": "CONVEYOR",
    "FRON": "FRONT", "FORNT": "FRONT", "FROMT": "FRONT", "FOWARD": "FORWARD", "FORWORD": "FORWARD",
    "BATERY": "BATTERY", "BATTRY": "BATTERY", "BATTTERY": "BATTERY", "MALFUNTION": "MALFUNCTION", "MULFUNCTION": "MALFUNCTION",
    "HIGHT": "HEIGHT", "PLATFROM": "PLATFORM", "PLAFORM": "PLATFORM", "TEMPEARTURE": "TEMPERATURE",
    "PRESURE": "PRESSURE", "PREESURE": "PRESSURE", "UNALBLE": "UNABLE", "UNBALE": "UNABLE",
    "STERING": "STEERING", "STERRING": "STEERING", "TRASMISSION": "TRANSMISSION", "HANBRAKE": "HANDBRAKE",
    "INTEMITTEN": "INTERMITTENT", "INTERMITTEN": "INTERMITTENT", "INTEMITTENT": "INTERMITTENT", "INTERMITENT": "INTERMITTENT",
    "UNSEVICEABLE": "UNSERVICEABLE", "TRANVERSE": "TRANSVERSE", "EKZOS": "EXHAUST", "EXZOS": "EXHAUST",
    "REERSE": "REVERSE", "REVERS": "REVERSE", "ACCELARATOR": "ACCELERATOR", "ACCERALATOR": "ACCELERATOR",
    "CLUCTH": "CLUTCH", "EXLE": "AXLE", "BOGGEY": "BOGIE", "BOGUE": "BOGIE",
    "RADIOTOR": "RADIATOR", "SHUTDWON": "SHUTDOWN", "CRAKED": "CRACKED", "MISSSING": "MISSING", "DAIHTSU": "DAIHATSU",
    "BONET": "BONNET", "LEAKNG": "LEAKING", "PADLE": "PEDAL", "BREAKE": "BRAKE", "SHAREBOLT": "SHEAR BOLT",
    "NONNEGLIGENCE": "NON-NEGLIGENCE", "HAEDLAMP": "HEADLAMP", "HEADLMAP": "HEADLAMP", "BECON": "BEACON",
    "MILLEAGE": "MILEAGE", "MILAGE": "MILEAGE", "ROUNDBOUT": "ROUNDABOUT", "ROUDABOUT": "ROUNDABOUT",
    "SATELITE": "SATELLITE", "SATALITE": "SATELLITE", "AUTOHUTDOWN": "AUTO SHUTDOWN", "DEISEL": "DIESEL",
    "FUNTIONAL": "FUNCTIONAL", "OUTTER": "OUTER", "PARKBRAKE": "PARK BRAKE", "PASSANGER": "PASSENGER", "PASSENGGER": "PASSENGER",
    "RIGH": "RIGHT", "SATGING": "STAGING", "SYSYTEM": "SYSTEM", "TROTTLE": "THROTTLE", "TYRESHOP": "TYRE SHOP",
    "WTERPOINT": "WATER POINT", "AICRAFT": "AIRCRAFT", "ALLIGNMENT": "ALIGNMENT", "BARCKET": "BRACKET",
    "BRAKEDOWN": "BREAKDOWN", "COULING": "COOLING", "ENGIBE": "ENGINE", "FLIGT": "FLIGHT", "JEMMED": "JAMMED",
    "LIGTH": "LIGHT", "RAILLING": "RAILING", "VACUM": "VACUUM", "WAEK": "WEAK", "WATR": "WATER", "WORNED": "WORN",
    "STOPER": "STOPPER", "HANDTHROTLE": "HAND THROTTLE", "COMPALAIN": "COMPLAINT", "PANER": "PANEL", "STRAT": "START",
}

# Valid words that should not be auto-corrected or reported as typo candidates.
KNOWN_CORRECT_WORDS = {
    "BOGEY",  # valid aircraft/GSE mechanical term
    "TIRE",   # valid American spelling; do not silently change regional usage
    "TIRES",  # valid American plural spelling
    "ALG",    # Air Algérie/IATA airport or airline operational code
    "KE",     # Korean Air operational code
    "KUL",    # Kuala Lumpur IATA airport code
}

# Terms known to be Malay, equipment names, abbreviations, or correct rare words.
MALAY_WORDS = [
    "BELAKANG", "DEPAN", "KIRI", "KANAN", "TAK", "BOLEH", "BOCOR", "PUTUS", "TERCABUT", "PANCIT",
    "TAYAR", "KENDUR", "KERAS", "MINYAK", "SKRU", "RUMAH", "HANTU", "BARANG", "BENGKOK", "BERGEGAR",
    "BIJI", "CAKAP", "DALAM", "DENGAR", "DI", "ENGSEL", "GOM", "HABIS", "KELUAR", "KON",
    "KONTENA", "LUAR", "MALAP", "NAIK", "OFIS", "PAKU", "PASIR", "PENDEK", "PENGIKAT", "PI",
    "PROSES", "PUSING", "SANGKUT", "SEJUK", "SEMPUT", "SYART", "TALI", "TEROWONG", "TUKAR", "TURUN",
    "ADA", "BERASAP",
]
