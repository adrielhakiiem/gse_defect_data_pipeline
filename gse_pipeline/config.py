import calendar
import re
from pathlib import Path

try:
    import config_local
except ImportError as exc:
    raise RuntimeError(
        "Missing config_local.py. Copy config_local.example.py to config_local.py "
        "and enter your private SharePoint paths there."
    ) from exc

# ---------------------------------------------------------------------------
# FILE PATHS
# ---------------------------------------------------------------------------
ROOT_DIR = Path(__file__).resolve().parent.parent

# INPUT: always read the private SharePoint source set in config_local.py.
SOURCE_FILE = Path(config_local.SHAREPOINT_SOURCE_FILE)

# OUTPUT: choose EXACTLY ONE of the following two lines.
# LOCAL REVIEW (active): output beside this repository.
# OUTPUT_FILE = ROOT_DIR / "EDR_GSE_Master_MERGED.xlsx"
# SHAREPOINT PUBLISHING: comment the local line above, then uncomment this line.
OUTPUT_FILE = Path(config_local.SHAREPOINT_OUTPUT_FILE)

# ---------------------------------------------------------------------------
# SHEET NAMES
# ---------------------------------------------------------------------------
MASTER_SHEET = "EDR GSE (FEB 2026 - MAY 2026)"
# DISABLED EQUIPMENT-CODE QA: uncomment this together with the marked blocks
# in pipeline.py, write_output.py, and run_pipeline.py to restore the review tab.
# EQUIPMENT_CODE_QA_SHEET = "QA - Equipment Code Mismatches"

# Monthly sheets are discovered from the open source workbook rather than
# maintained manually. New tabs such as "OCTOBER 2026" are included next run.
MONTH_SHEET_PATTERN = re.compile(r"^([A-Z]+)\s+(\d{4})$")
NON_MONTH_SHEETS = {
    MASTER_SHEET,
    "Equipment Pie Chart", "Top 3 Equipment By Month", "Equipment Type column Chart",
    "INSPECTION FEB MAR APR MAY 26", "INSPECTION MAY ", "DEFECTS SPECIFICATION",
    "MOTORIZED-NON MOTORIZED", "DEFECTS CATEGORIZATION", "No Baggage Defects",
    "UNIT NO EQUIPMENT DEFECTS ", "Sheet1", ".",
}


def detect_month_sheets(sheet_names):
    """Return (month-name label, sheet name) pairs in calendar order.

    Only uppercase MONTH YYYY sheet names qualify; known reporting/support
    tabs are explicitly ignored.
    """
    detected = []
    for sheet_name in sheet_names:
        if sheet_name in NON_MONTH_SHEETS:
            continue
        match = MONTH_SHEET_PATTERN.fullmatch(sheet_name)
        if not match:
            continue
        month_name, year_text = match.groups()
        try:
            month_number = list(calendar.month_name).index(month_name.title())
        except ValueError:
            continue
        detected.append(((int(year_text), month_number), month_name.title(), sheet_name))
    return [(label, sheet_name) for _, label, sheet_name in sorted(detected)]

# ---------------------------------------------------------------------------
# COLUMN SCHEMA
# ---------------------------------------------------------------------------
# The canonical set of columns every row will be normalized into, regardless
# of what the source sheet called them.
STANDARD_COLUMNS = [
    "Date",
    "By Month",
    "ADS Equipment No",
    "Equipment Type",
    "Maintenance By",
    "Motorized/Non-Motorized",
    "Defects Description",
    "Defects Categorization",
    "Defects Specification",
]

# Some monthly sheets use a different header for the same underlying field.
# The pipeline looks up each standard column's real header using this list,
# in order, and uses the first one that's actually present in that sheet.
# April 2026 used "Ownership" instead of "Maintenance By" - same TCR/GSE
# values, just a different header - so it's mapped here instead of being
# treated as a separate column.
COLUMN_ALIASES = {
    "Maintenance By": ["Maintenance By", "Ownership"],
    "Defects Categorization": ["Defects Categorization", "Defect Categorization"],
}

# ---------------------------------------------------------------------------
# MATCH KEY (for deduplication / "what's missing from Master" comparison)
# ---------------------------------------------------------------------------
# IMPORTANT: ADS Equipment No alone is NOT a safe match key - the same
# equipment can legitimately have multiple separate defect entries on the
# same day. The composite key below is what lets the pipeline tell a
# genuine repeat (same equipment, different defect) apart from an actual
# duplicate (same equipment, same defect, entered twice).
MATCH_KEY_FIELDS = ["Date", "ADS Equipment No", "Defects Description"]

# ---------------------------------------------------------------------------
# TEXT STANDARDIZATION
# ---------------------------------------------------------------------------
# Known misspellings -> correct spelling, applied only as whole-word,
# case-preserving replacements inside Defects Description. Equipment codes
# (e.g. DBT099) are never touched because they aren't in this dictionary -
# only exact known-bad words get replaced, never a general spellchecker.
SPELLING_MAP = {
    # negligence family (the single biggest source of typos in this dataset)
    "NEGLIGENE": "NEGLIGENCE", "NEGLEGENCE": "NEGLIGENCE", "NEGLGENCE": "NEGLIGENCE",
    "NELIGENCE": "NEGLIGENCE", "NGELIGENCE": "NEGLIGENCE", "NGLIGENCE": "NEGLIGENCE",
    "NEGELIGENCE": "NEGLIGENCE", "NEGLIGEIGENCE": "NEGLIGENCE", "NEGLIGNECE": "NEGLIGENCE",
    "NIGLIGENCE": "NEGLIGENCE", "ENGLIGENCE": "NEGLIGENCE", "NEGIGENCE": "NEGLIGENCE",
    "NEEGLIGENCE": "NEGLIGENCE", "NEELIGENCE": "NEGLIGENCE", "NEGILGENCE": "NEGLIGENCE",
    "NEGLIENCE": "NEGLIGENCE", "NEGLIGENGE": "NEGLIGENCE", "NEGLIGNCE": "NEGLIGENCE",
    "NEGLIIGENCE": "NEGLIGENCE", "NEGLOGENCE": "NEGLIGENCE", "NEGLUGENCE": "NEGLIGENCE",
    "NEHGLIGENCE": "NEGLIGENCE",
    # tyre family (standardizing to British spelling - the majority form in this data)
    "YTRE": "TYRE", "TYR": "TYRE", "TYE": "TYRE", "TIRE": "TYRE", "TIRES": "TYRES",
    # other recurring technical-term typos
    "PUCNTURED": "PUNCTURED", "PUNCUTRED": "PUNCTURED", "PPUNCTURED": "PUNCTURED", "PUNCTURD": "PUNCTURED",
    "COMVEYOR": "CONVEYOR", "CONVOYER": "CONVEYOR", "CONEYOR": "CONVEYOR", "CONVER": "CONVEYOR", "CONVEYPOR": "CONVEYOR",
    "FRON": "FRONT", "FORNT": "FRONT", "FROMT": "FRONT",
    "FOWARD": "FORWARD", "FORWORD": "FORWARD",
    "BATERY": "BATTERY", "BATTRY": "BATTERY", "BATTTERY": "BATTERY",
    "MALFUNTION": "MALFUNCTION", "MULFUNCTION": "MALFUNCTION",
    "HIGHT": "HEIGHT",
    "PLATFROM": "PLATFORM", "PLAFORM": "PLATFORM",
    "TEMPEARTURE": "TEMPERATURE",
    "PRESURE": "PRESSURE", "PREESURE": "PRESSURE",
    "UNALBLE": "UNABLE", "UNBALE": "UNABLE",
    "STERING": "STEERING", "STERRING": "STEERING",
    "TRASMISSION": "TRANSMISSION",
    "HANBRAKE": "HANDBRAKE",
    "INTEMITTEN": "INTERMITTENT", "INTERMITTEN": "INTERMITTENT", "INTEMITTENT": "INTERMITTENT", "INTERMITENT": "INTERMITTENT",
    "UNSEVICEABLE": "UNSERVICEABLE",
    "TRANVERSE": "TRANSVERSE",
    "EKZOS": "EXHAUST", "EXZOS": "EXHAUST",
    "REERSE": "REVERSE", "REVERS": "REVERSE",
    "ACCELARATOR": "ACCELERATOR", "ACCERALATOR": "ACCELERATOR",
    "CLUCTH": "CLUTCH",
    "EXLE": "AXLE",
    "BOGEY": "BOGIE", "BOGGEY": "BOGIE", "BOGUE": "BOGIE",
    "RADIOTOR": "RADIATOR",
    "SHUTDWON": "SHUTDOWN",
    "CRAKED": "CRACKED",
    "MISSSING": "MISSING",
    "DAIHTSU": "DAIHATSU",
    # found via discover_typos.py (word-frequency scan) after the first cleaning pass -
    # these were missed in the original manual review
    "BONET": "BONNET",
    "LEAKNG": "LEAKING",
    "PADLE": "PEDAL",
    "BREAKE": "BRAKE",
    "SHAREBOLT": "SHEAR BOLT",
    "NONNEGLIGENCE": "NON-NEGLIGENCE",  # zero-space variant the spacing-fix regex couldn't catch
    "HAEDLAMP": "HEADLAMP", "HEADLMAP": "HEADLAMP",
    "BECON": "BEACON",
    "MILLEAGE": "MILEAGE", "MILAGE": "MILEAGE",
    "ROUNDBOUT": "ROUNDABOUT", "ROUDABOUT": "ROUNDABOUT",
    "SATELITE": "SATELLITE", "SATALITE": "SATELLITE",
    "AUTOHUTDOWN": "AUTO SHUTDOWN",
    "DEISEL": "DIESEL",
    "FUNTIONAL": "FUNCTIONAL",
    "OUTTER": "OUTER",
    "PARKBRAKE": "PARK BRAKE",
    "PASSANGER": "PASSENGER", "PASSENGGER": "PASSENGER",
    "RIGH": "RIGHT",
    "SATGING": "STAGING",
    "SYSYTEM": "SYSTEM",
    "TROTTLE": "THROTTLE",
    "TYRESHOP": "TYRE SHOP",
    "WTERPOINT": "WATER POINT",
    "AICRAFT": "AIRCRAFT",
    "ALLIGNMENT": "ALIGNMENT",
    "BARCKET": "BRACKET",
    "BRAKEDOWN": "BREAKDOWN",
    "COULING": "COOLING",
    "ENGIBE": "ENGINE",
    "FLIGT": "FLIGHT",
    "JEMMED": "JAMMED",
    "LIGTH": "LIGHT",
    "RAILLING": "RAILING",
    "VACUM": "VACUUM",
    "WAEK": "WEAK",
    "WATR": "WATER",
    "WORNED": "WORN",
    "STOPER": "STOPPER",
    "HANDTHROTLE": "HAND THROTTLE",
    "COMPALAIN": "COMPLAINT",
    "PANER": "PANEL",
    "STRAT": "START",
}

# Words that showed up as rare during discovery but were confirmed NOT to be
# typos (abbreviations, brand/equipment names, genuinely correct rare terms).
# Kept here as a record so the next person doesn't have to re-investigate
# them from scratch when discover_typos.py surfaces them again.
CONFIRMED_NOT_TYPOS = {
    "ACKLIFT": "Equipment/unit type name, not a typo",
    "HANDGRILL": "Ambiguous (handrail? hand grip?) - needs domain expert, not guessed",
    "KINGPIN": "Correct mechanical term (steering component)",
    "HINO": "Vehicle brand name",
    "AXIA": "Vehicle model name (Perodua Axia)",
    "PERDANA": "Vehicle model name (Proton Perdana)",
    "MUDGUARD": "Correct word",
}

# Categorical (drop-down-style) field standardization: {column: {lowercased
# variant: canonical value}}. Applied as exact (not fuzzy) matches only.
CATEGORICAL_MAP = {
    # This relies on standardize_text()'s exact-match lookup (str(val).strip().lower()).
    # Never convert it to a regex or substring match: "motorized" in value.lower()
    # would incorrectly match inside "Non-Motorized" too.
    "Motorized/Non-Motorized": {
        "non - motorized": "Non-Motorized",
    },
    "Defects Categorization": {
        "body work": "Body Work",
    },
    "Defects Specification": {
        "item defect": "Item Defect",
        "item defects": "Item Defect",
        "electrical problem": "Electrical Problem",
        "incident issues": "Incident Issues",
    },
    "Equipment Type": {
        "ambulift": "Ambulift",
        "baggage tlractor": "Baggage Tractor",
        "catering highlift": "Catering Highlift",
    },
}

# Malay words found in Defects Description during review. These are NOT
# translated by the pipeline (translation needs a human who knows the
# intended technical meaning) - rows containing them are only flagged
# (highlighted yellow) for a later manual pass.
MALAY_WORDS = [
    "BELAKANG", "DEPAN", "KIRI", "KANAN", "TAK", "BOLEH", "BOCOR", "PUTUS", "TERCABUT", "PANCIT",
    "TAYAR", "KENDUR", "KERAS", "MINYAK", "SKRU", "RUMAH", "HANTU", "BARANG", "BENGKOK", "BERGEGAR",
    "BIJI", "CAKAP", "DALAM", "DENGAR", "DI", "ENGSEL", "GOM", "HABIS", "KELUAR", "KE", "KON",
    "KONTENA", "KUL", "LUAR", "MALAP", "NAIK", "OFIS", "PAKU", "PASIR", "PENDEK", "PENGIKAT", "PI",
    "PROSES", "PUSING", "SANGKUT", "SEJUK", "SEMPUT", "SYART", "TALI", "TEROWONG", "TUKAR", "TURUN",
    "ADA", "BERASAP",
]

# ---------------------------------------------------------------------------
# STYLING
# ---------------------------------------------------------------------------
FONT_NAME = "Arial"
COLOR_APPENDED_ROW = "FFF2CC"   # light amber - row was missing from Master and got appended
COLOR_MALAY_FLAG = "FFFF00"     # bright yellow - description contains unreviewed Malay text
COLOR_HEADER_FILL = "2F5496"

# ---------------------------------------------------------------------------
# MOTORIZED/NON-MOTORIZED CONTAMINATION
# ---------------------------------------------------------------------------
# Values that belong in Motorized/Non-Motorized but recurringly get typed
# into Maintenance By by mistake.
MAINTENANCE_BY_CONTAMINATION_VALUES = {"MOTORIZED", "NON-MOTORIZED", "NON MOTORIZED", "NONMOTORIZED"}
VALID_MAINTENANCE_BY = {"TCR", "GSE"}

# Minimum share of agreement among an equipment's OTHER valid Maintenance By
# entries before we trust it enough to auto-fill. Below this, leave blank
# and flag instead of guessing.
MAINTENANCE_BY_CONFIDENCE_THRESHOLD = 0.85
