"""Edit the filenames here, then save and rerun the Streamlit app.

Change only the quoted names on the RIGHT. Keep the dictionary keys on the LEFT.
GEO filename settings are first. All output files are CSV.
GEO names follow: base name + Area suffix + Email/Viber suffix + .csv.
"""

# ==================== EDIT YOUR GEO FILENAMES HERE ====================
# These five base names apply to both Mindanao and R3, combining ALL levels.
GEO_OUTPUT_FILENAMES = {
    # QUALIFIED Christmas Amnesty accounts:
    "qualified_cards": "CS AMNESTY GEO CARDS",
    "qualified_bel": "CS AMNESTY GEO BEL",

    # BLANK Christmas Amnesty accounts:
    "blank_cards": "GEO_BLANK_CARDS",
    "blank_bel": "GEO_BLANK_BEL",
    "blank_other_products": "GEO_BLANK_OTHER_PRODUCTS",

    # Used only if qualified accounts have products other than CARDS and BEL.
    "qualified_review": "GEO_QUALIFIED_OTHER_PRODUCTS_REVIEW",
}

# Edit these to rename the Area part of each GEO filename.
GEO_AREA_SUFFIXES = {
    "mindanao": "_MINDANAO",
    "r3": "_R3",
    "area_review": "_AREA_REVIEW",
}

# Name of the GEO ZIP download.
GEO_ZIP_FILENAME = "GEO_AREA_CHRISTMAS_AMNESTY_FILES.zip"

# Example: "CS AMNESTY GEO CARDS" becomes
# CS AMNESTY GEO CARDS_MINDANAO_Email.csv and CS AMNESTY GEO CARDS_R3_Email.csv.
# CHANNEL_SUFFIXES below controls the Email/Viber part for both workflows.


# ==================== STANDARD NEGATIVE FILENAMES ====================
OUTPUT_FILENAMES = {
    "qualified_cards_level_1": "CS AMNESTY NEGA CARDS - L1",
    "qualified_cards_level_6": "CS AMNESTY NEGA CARDS - L6",
    "qualified_bel": "CS AMNESTY NEGA BEL - PL",
    "blank_cards_level_1": "NEGATIVE_BLANK_CARDS_LEVEL_1",
    "blank_cards_level_6": "NEGATIVE_BLANK_CARDS_LEVEL_6",
    "blank_bel": "NEGATIVE_BLANK_BEL",
    "blank_other_products": "NEGATIVE_BLANK_OTHER_PRODUCTS",
    # Used only if qualified rows do not match any of the three rules above.
    "qualified_review": "NEGATIVE_QUALIFIED_REVIEW",
    # Used only for blank-status CARDS at levels other than Level 1 or Level 6.
    "blank_review": "NEGATIVE_BLANK_REVIEW",
}

# These are appended to EVERY base filename above (before the file extension).
CHANNEL_SUFFIXES = {
    "email": "_Email",
    "viber": "_Viber",
}

# Filename for the download containing all of the separate files.
ZIP_FILENAME = "NEGATIVE_CHRISTMAS_AMNESTY_FILES.zip"

# Default destination for "Save files to folder". Relative paths start here in the project.
# You can also change this folder in the app before saving.
OUTPUT_FOLDER = "output"
