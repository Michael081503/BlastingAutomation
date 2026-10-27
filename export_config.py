"""Edit the filenames here, then save and rerun the Streamlit app.

Change only the quoted names on the RIGHT. Keep the dictionary keys on the LEFT.
The app adds a channel suffix, then .csv. All output files are CSV.
"""

# ==================== EDIT YOUR OUTPUT FILENAMES HERE ====================
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
