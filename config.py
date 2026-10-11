INCOMING_FOLDER = r"C:\Users\Vonotec\OneDrive - Vonotec (1)\Automation"
PROCESSED_LOG = "processed_files_log.txt"
GEMINI_API_KEY = ""
SHAREPOINT_HOSTNAME = "vonotecph.sharepoint.com"
SHAREPOINT_SITE_NAME = "Vonotec"
SHAREPOINT_FOLDER_PATH = "07 - Employee Folders (to be reviewed)/Lorenz/Whiteboard Excel Files"

JSON_KEYS = [
    "Submitter", "Date", "Elevation", "Drop", "Floor", "Tower",
    "Sealant Damage", "Sealant Dimension",
    "Concrete Damage", "Concrete Dimension",
    "Paint Damage", "Paint Dimension",
    "Gasket Damage", "Gasket Dimension",
    "With Film Damage", "With Film Dimension"
]

# ==========================================================================
# REPAIR FORMULAS: edit ONLY this block when a repair formula changes.
#
# Each template is an Excel formula written WITHOUT the leading "=":
#   {h}         is replaced by the cell holding the board dimension
#   {modifier}  is replaced by the sealant location code (CC, CF, GG, ...)
# ==========================================================================

def _limit(code):
    """Excel lookup of a code's Repair Limit (column F of the Reference Data sheet)."""
    return f'VLOOKUP("{code}", \'Reference Data\'!$A:$F, 6, FALSE)'


F_CRACK = "ROUNDUP(MROUND({h}*1.5, 10)/100, 2)"
F_SURFACE = "ROUNDUP({h}*1.15/1000, 2)"
F_PAINT = "IF(ROUNDUP({h}*1.3/1000, 2)<1, 1, ROUNDUP({h}*1.3/1000, 2))"
F_GASKET = "ROUNDUP(IF(MROUND({h}*8, 10)>" + _limit("DG") + ", " + _limit("DG") + ", MROUND({h}*8, 10))/100, 2)"
F_SEALANT = "MROUND({h}*0.08, 0.05)"  # no location code written
F_SEALANT_WITH_LOCATION = (
    "MROUND(IF(({h}*8/100)>" + _limit("{modifier}") + ", " + _limit("{modifier}") + ", ({h}*8/100)), 0.05)"
)

# ==========================================================================
# REFERENCE DATA: one row per defect code.
#
# Columns A-F are copied into the "Reference Data" sheet of every Excel file
# (the lookup formulas read them). Columns G onward are rules for the app only
# and are NOT written to Excel:
#   G Board Unit                     unit of the dimension written on the board
#   H Repair Unit                    unit of the calculated repair dimension
#   I Repair Formula                 template ("" = no calculated repair dimension)
#   J Repair Formula (with location) used instead of I when a location code follows
#                                    the defect (DS / MS only)
# ==========================================================================

REFERENCE_SHEET_COLUMNS = 6  # how many columns (from the left) go into the Excel sheet

REFERENCE_DATA = [
    ["Defect Code", "Defect Type", "Likely Cause(s)", "Justification", "Recommended Repair", "Repair Limit",
     "Board Unit", "Repair Unit", "Repair Formula", "Repair Formula (with location)"],
    ["CC-", "Minor Concrete Crack", "Act of God or Wear and Tear", "Typically form due to drying shrinkage or building movement. Can be exacerbated by environmental stress or thermal cycling. Occurences beyond human control", "Apply epoxy mortar.", "",
     "CM", "L.M.", F_CRACK, ""],
    ["C-", "Minor Concrete Crack", "Act of God or Wear and Tear", "Typically form due to drying shrinkage or building movement. Can be exacerbated by environmental stress or thermal cycling. Occurences beyond human control", "Apply epoxy mortar.", "",
     "CM", "L.M.", F_CRACK, ""],
    ["CC+", "Major Concrete Crack", "Act of God or Wear and Tear", "Often indicate structural stress or abnormal movement; may also reflect long-term degradation in structural load paths. Soil movements blow causing structure movements.", "Perform epoxy injection.", "",
     "CM", "L.M.", F_CRACK, ""],
    ["C+", "Major Concrete Crack", "Act of God or Wear and Tear", "Often indicate structural stress or abnormal movement; may also reflect long-term degradation in structural load paths. Soil movements blow causing structure movements.", "Perform epoxy injection.", "",
     "CM", "L.M.", F_CRACK, ""],
    ["BH", "Dented Surface or Bugholes", "Workmanship", "Result from improper concrete casting, lack of consolidation, or poor formwork. Appear early and are usually non-structural.", "Apply repair mortar with concrete epoxy (A&B). (Repainting upon building admin decision)", "",
     "CM²", "SQ.M.", F_SURFACE, ""],
    ["US", "Dented Surface or Bugholes", "Workmanship", "Result from improper concrete casting, lack of consolidation, or poor formwork. Appear early and are usually non-structural.", "Apply repair mortar with concrete epoxy (A&B). (Repainting upon building admin decision)", "",
     "CM²", "SQ.M.", F_SURFACE, ""],
    ["DP", "Discolored Paint", "Wear and Tear", "Fades or stains over time due to sun exposure, moisture, and pollution. A predictable result of long-term exposure.", "Retouch paint; ensure proper surface prep.", "",
     "CM²", "SQ.M.", F_PAINT, ""],
    ["FP", "Flaky Paint", "Wear and Tear", "Lapsed warranty; degrades due to age, UV, and temperature shifts. Failures are common after several years of exposure.", "Remove loose paint, prime, and repaint.", "",
     "CM²", "SQ.M.", F_PAINT, ""],
    ["BP", "Bubbly Paint", "Wear and Tear", "Lapsed warranty; degrades due to age, UV, and temperature shifts. Failures are common after several years of exposure.", "Remove bubbles, dry surface, prime, and repaint.", "",
     "CM²", "SQ.M.", F_PAINT, ""],
    ["NP", "No Paint", "Workmanship", "Skipped by previous workers", "Include in next beautification projects.", "",
     "", "", "", ""],
    ["DG", "Damaged Gasket", "Wear and Tear", "Gaskets degrade due to age, UV, and temperature shifts. Failures are common after several years of exposure", "Remove gasket and reseal.", 3.2,
     "CM", "L.M.", F_GASKET, ""],
    ["DS", "Damaged Sealant", "Wear and Tear", "Lapsed warranty; Cracking or tearing due to joint movement, UV aging, or thermal cycles. Material failure over time is expected even with proper application.", "Remove existing sealant, clean joint, install backer rod and apply new sealant.", "",
     "CM", "L.M.", F_SEALANT, F_SEALANT_WITH_LOCATION],
    ["MS", "Missing Sealant", "Workmanship", "Indicates that sealant was never applied during installation. This reflects an omission in workmanship or quality control at the time of construction.", "Clean joint and install backer rod and applynew sealant.", "",
     "CM", "L.M.", F_SEALANT, F_SEALANT_WITH_LOCATION],
    ["BG", "Broken Glass (BG)", "Unclassified (Circumstantial)", "Typically the result of impact, high wind pressure, or spontaneous breakage due to thermal stress. Often sudden and isolated.", "Replace glass and reseal edges.", "",
     "", "", "", ""],
    ["GLASS-FRAME", "Glass Panel still with Film", "Workmanship", "Removal skipped by previous workers", "Include in next facade works", "",
     "CM", "L.M.", "", ""],
    ["FRAME", "Glass Panel still with Film", "Workmanship", "Removal skipped by previous workers", "Include in next facade works", "",
     "CM", "L.M.", "", ""],
    ["CC", "Sealant Location Modifier", "-", "-", "-", 4.8, "", "", "", ""],
    ["CF", "Sealant Location Modifier", "-", "-", "-", 3.2, "", "", "", ""],
    ["FC", "Sealant Location Modifier (Same as CF)", "-", "-", "-", 3.2, "", "", "", ""],
    ["GG", "Sealant Location Modifier", "-", "-", "-", 1.8, "", "", "", ""],
    ["FG", "Sealant Location Modifier", "-", "-", "-", 1.8, "", "", "", ""],
    ["GF", "Sealant Location Modifier (Same as FG)", "-", "-", "-", 1.8, "", "", "", ""],
    ["FF", "Sealant Location Modifier", "-", "-", "-", 3.2, "", "", "", ""]
]

# STRICT MATERIAL VALIDATION DICTIONARY
MATERIAL_RULES = {
    "Concrete": ["CC+", "CC-", "C+", "C-", "BH", "US"],
    "Paint": ["DP", "FP", "BP", "B", "NP"],
    "Gasket": ["DG"],
    "Sealant": ["DS", "MS"],
    "Broken Glass": ["BG"],
    "With Film": ["GLASS-FRAME", "FRAME"]
}
