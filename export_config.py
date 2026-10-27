"""Edit the filenames here, then save and rerun the Streamlit app.

Change only the quoted names on the RIGHT. Keep the dictionary keys on the LEFT.
GEO filename settings are first. All output files are CSV.
GEO names below already include the Area. The app adds _Email or _Viber and .csv.
"""

# ==================== EDIT YOUR GEO FILENAMES HERE ====================
# Each entry is the full base filename for one group. All levels are combined.
GEO_OUTPUT_FILENAMES = {
    # MINDANAO: edit the five filenames on the right.
    "mindanao_qualified_cards": "CS AMNESTY GEO CARDS_MINDANAO",
    "mindanao_qualified_bel": "CS AMNESTY GEO BEL_MINDANAO",
    "mindanao_blank_cards": "GEO_BLANK_CARDS_MINDANAO",
    "mindanao_blank_bel": "GEO_BLANK_BEL_MINDANAO",
    "mindanao_blank_other_products": "GEO_BLANK_OTHER_PRODUCTS_MINDANAO",

    # REGION 3: edit the five filenames on the right.
    "r3_qualified_cards": "CS AMNESTY GEO CARDS_R3",
    "r3_qualified_bel": "CS AMNESTY GEO BEL_R3",
    "r3_blank_cards": "GEO_BLANK_CARDS_R3",
    "r3_blank_bel": "GEO_BLANK_BEL_R3",
    "r3_blank_other_products": "GEO_BLANK_OTHER_PRODUCTS_R3",

    # Review files, used only when unmatched products or Areas are present.
    "mindanao_qualified_review": "GEO_QUALIFIED_OTHER_PRODUCTS_REVIEW_MINDANAO",
    "r3_qualified_review": "GEO_QUALIFIED_OTHER_PRODUCTS_REVIEW_R3",
    "area_review_qualified_cards": "CS AMNESTY GEO CARDS_AREA_REVIEW",
    "area_review_qualified_bel": "CS AMNESTY GEO BEL_AREA_REVIEW",
    "area_review_blank_cards": "GEO_BLANK_CARDS_AREA_REVIEW",
    "area_review_blank_bel": "GEO_BLANK_BEL_AREA_REVIEW",
    "area_review_blank_other_products": "GEO_BLANK_OTHER_PRODUCTS_AREA_REVIEW",
    "area_review_qualified_review": "GEO_QUALIFIED_OTHER_PRODUCTS_REVIEW_AREA_REVIEW",
}

# Name of the GEO ZIP download.
GEO_ZIP_FILENAME = "GEO_AREA_CHRISTMAS_AMNESTY_FILES.zip"

# Example: "CS AMNESTY GEO CARDS_MINDANAO" becomes
# CS AMNESTY GEO CARDS_MINDANAO_Email.csv and CS AMNESTY GEO CARDS_MINDANAO_Viber.csv.
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
