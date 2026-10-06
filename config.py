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
    "Gasket Damage", "Gasket Dimension"
]

# Added C+ and C- variations to accommodate different worker handwriting styles
REFERENCE_DATA = [
    ["Defect Code", "Defect Type", "Likely Cause(s)", "Justification", "Recommended Repair"],
    ["CC-", "Minor Concrete Crack", "Act of God or Wear and Tear", "Typically form due to drying shrinkage or building movement. Can be exacerbated by environmental stress or thermal cycling. Occurences beyond human control", "Apply epoxy mortar."],
    ["C-", "Minor Concrete Crack", "Act of God or Wear and Tear", "Typically form due to drying shrinkage or building movement. Can be exacerbated by environmental stress or thermal cycling. Occurences beyond human control", "Apply epoxy mortar."],
    ["CC+", "Major Concrete Crack", "Act of God or Wear and Tear", "Often indicate structural stress or abnormal movement; may also reflect long-term degradation in structural load paths. Soil movements blow causing structure movements.", "Perform epoxy injection."],
    ["C+", "Major Concrete Crack", "Act of God or Wear and Tear", "Often indicate structural stress or abnormal movement; may also reflect long-term degradation in structural load paths. Soil movements blow causing structure movements.", "Perform epoxy injection."],
    ["BH", "Dented Surface or Bugholes", "Workmanship", "Result from improper concrete casting, lack of consolidation, or poor formwork. Appear early and are usually non-structural.", "Apply repair mortar with concrete epoxy (A&B). (Repainting upon building admin decision)"],
    ["US", "Dented Surface or Bugholes", "Workmanship", "Result from improper concrete casting, lack of consolidation, or poor formwork. Appear early and are usually non-structural.", "Apply repair mortar with concrete epoxy (A&B). (Repainting upon building admin decision)"],
    ["DP", "Discolored Paint", "Wear and Tear", "Fades or stains over time due to sun exposure, moisture, and pollution. A predictable result of long-term exposure.", "Retouch paint; ensure proper surface prep."],
    ["FP", "Flaky Paint", "Wear and Tear", "Lapsed warranty; degrades due to age, UV, and temperature shifts. Failures are common after several years of exposure.", "Remove loose paint, prime, and repaint."],
    ["BP", "Bubbly Paint", "Wear and Tear", "Lapsed warranty; degrades due to age, UV, and temperature shifts. Failures are common after several years of exposure.", "Remove bubbles, dry surface, prime, and repaint."],
    ["DG", "Damaged Gasket", "Wear and Tear", "Gaskets degrade due to age, UV, and temperature shifts. Failures are common after several years of exposure", "Remove gasket and reseal."],
    ["DS", "Damaged Sealant", "Wear and Tear", "Lapsed warranty; Cracking or tearing due to joint movement, UV aging, or thermal cycles. Material failure over time is expected even with proper application.", "Remove existing sealant, clean joint, install backer rod and apply new sealant."],
    ["MS", "Missing Sealant", "Workmanship", "Indicates that sealant was never applied during installation. This reflects an omission in workmanship or quality control at the time of construction.", "Clean joint and install backer rod and applynew sealant."],
    ["BG", "Broken Glass (BG)", "Unclassified (Circumstantial)", "Typically the result of impact, high wind pressure, or spontaneous breakage due to thermal stress. Often sudden and isolated.", "Replace glass and reseal edges."]
]

# STRICT MATERIAL VALIDATION DICTIONARY
MATERIAL_RULES = {
    "Concrete": ["CC+", "CC-", "C+", "C-", "BH", "US"],
    "Paint": ["DP", "FP", "BP", "B"], 
    "Gasket": ["DG"],
    "Sealant": ["DS", "MS"],
    "Broken Glass": ["BG"]
}
