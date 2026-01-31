import random
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle

# --- CONFIGURATION ---
filename = "Smart_Availability_Matrix.xlsx"
regions = ["USA", "Europe", "UAE", "APAC", "Saudi"]
languages = ["English", "Hindi", "Spanish", "German", "Arabic", "French", "Italian"]
services = [
    "ASR",
    "Redaction",
    "Conversation Facts",
    "Gen AI Disposition",
    "ITN",
    "Intent (UniFit)",
    "Entity (NER/Spacy)",
    "Gen AI Summary",
    "Pro-active Service",
    "SSA-LLM",
    "KaaS"
]
products = ["SSA", "RTGA", "CIA", "CRA"]
statuses = [
    "Full Support (In-Region)",
    "Full Support (Cross-Region)",
    "Limited Support (In-Region)",
    "Limited Support (Cross-Region)",
    "Not Supported"
]

# Dependency Logic (For Reference/Comments)
# SSA: ASR, SSA-LLM, KaaS
# RTGA: ASR, Intent, Entity, Gen AI Summary, Gen AI Disposition, Redaction, KaaS, Pro-active Service
# CIA: ASR, Gen AI Disposition, Redaction, Conversation Facts
# CRA: ASR, Redaction

wb = Workbook()

# ==========================================
# SHEET 1: SERVICE INPUT (The Source of Truth)
# ==========================================
ws_data = wb.active
ws_data.title = "Service_Input"
headers = ["Region", "Service", "Language", "Status", "Lookup_Key"]
ws_data.append(headers)

# Formatting
header_fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)
for cell in ws_data[1]:
    cell.fill = header_fill
    cell.font = header_font

# Generate Data (Full Factorial for SERVICES only)
row_num = 2
for r in regions:
    for s in services:
        for l in languages:
            # Random Status (5 options: Full In-Region, Full Cross-Region, Limited In-Region, Limited Cross-Region, Not Supported)
            status = random.choices(statuses, weights=[25, 20, 15, 15, 25], k=1)[0]
            # Key: Region|Service|Language
            key_formula = f'=A{row_num}&"|"&B{row_num}&"|"&C{row_num}'
            ws_data.append([r, s, l, status, key_formula])
            row_num += 1

ws_data.column_dimensions['E'].hidden = True

# ==========================================
# STYLING HELPERS
# ==========================================
def apply_dark_theme_formatting(ws, start_row, start_col, num_cols, num_rows):
    # Background
    ws.sheet_view.showGridLines = False

    # Conditional Formatting Ranges
    matrix_range = f"{get_column_letter(start_col)}{start_row+1}:{get_column_letter(start_col+num_cols-1)}{start_row+num_rows}"

    # Dark Green (Full Support - In-Region)
    dark_green_fill = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
    dxf_dark_green = DifferentialStyle(fill=dark_green_fill, font=Font(color="FFFFFF"))
    rule_dark_green = Rule(type="containsText", operator="containsText", text="Full Support (In-Region)", dxf=dxf_dark_green)
    rule_dark_green.formula = [f'NOT(ISERROR(SEARCH("Full Support (In-Region)", {matrix_range})))']
    ws.conditional_formatting.add(matrix_range, rule_dark_green)

    # Light Green (Full Support - Cross-Region)
    light_green_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    dxf_light_green = DifferentialStyle(fill=light_green_fill, font=Font(color="FFFFFF"))
    rule_light_green = Rule(type="containsText", operator="containsText", text="Full Support (Cross-Region)", dxf=dxf_light_green)
    rule_light_green.formula = [f'NOT(ISERROR(SEARCH("Full Support (Cross-Region)", {matrix_range})))']
    ws.conditional_formatting.add(matrix_range, rule_light_green)

    # Dark Yellow/Orange (Limited Support - In-Region)
    dark_yellow_fill = PatternFill(start_color="BF8F00", end_color="BF8F00", fill_type="solid")
    dxf_dark_yellow = DifferentialStyle(fill=dark_yellow_fill, font=Font(color="FFFFFF"))
    rule_dark_yellow = Rule(type="containsText", operator="containsText", text="Limited Support (In-Region)", dxf=dxf_dark_yellow)
    rule_dark_yellow.formula = [f'NOT(ISERROR(SEARCH("Limited Support (In-Region)", {matrix_range})))']
    ws.conditional_formatting.add(matrix_range, rule_dark_yellow)

    # Light Yellow (Limited Support - Cross-Region)
    light_yellow_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    dxf_light_yellow = DifferentialStyle(fill=light_yellow_fill)
    rule_light_yellow = Rule(type="containsText", operator="containsText", text="Limited Support (Cross-Region)", dxf=dxf_light_yellow)
    rule_light_yellow.formula = [f'NOT(ISERROR(SEARCH("Limited Support (Cross-Region)", {matrix_range})))']
    ws.conditional_formatting.add(matrix_range, rule_light_yellow)

    # Red (Not Supported)
    red_fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
    dxf_red = DifferentialStyle(fill=red_fill, font=Font(color="FFFFFF"))
    rule_red = Rule(type="containsText", operator="containsText", text="Not Supported", dxf=dxf_red)
    rule_red.formula = [f'NOT(ISERROR(SEARCH("Not Supported", {matrix_range})))']
    ws.conditional_formatting.add(matrix_range, rule_red)

# ==========================================
# SHEET 2: SERVICE DASHBOARD
# ==========================================
ws_serv = wb.create_sheet("Service_Dashboard")

# Header Controls
ws_serv['B2'] = "SERVICE AVAILABILITY MATRIX"
ws_serv['B2'].font = Font(size=16, bold=True, color="203764")
ws_serv['B4'] = "Select Region:"
ws_serv['C4'] = regions[0] # Default
ws_serv['C4'].font = Font(bold=True)
ws_serv['C4'].border = Border(bottom=Side(style='thin'))

# Dropdown
dv = DataValidation(type="list", formula1=f'"{",".join(regions)}"', allow_blank=False)
ws_serv.add_data_validation(dv)
dv.add(ws_serv['C4'])

# Grid Setup
start_row = 7
start_col = 2
ws_serv.cell(row=start_row, column=1, value="Language").font = Font(bold=True)

# Headers (Services)
for i, s in enumerate(services):
    c = start_col + i
    cell = ws_serv.cell(row=start_row, column=c, value=s)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color="2F5597", fill_type="solid")
    cell.alignment = Alignment(horizontal='center')
    ws_serv.column_dimensions[get_column_letter(c)].width = 20

# Rows (Languages) & Logic
for i, lang in enumerate(languages):
    r = start_row + 1 + i
    ws_serv.cell(row=r, column=1, value=lang).font = Font(bold=True)
    
    for j, s in enumerate(services):
        c = start_col + j
        # INDEX MATCH to look up Service Status directly
        formula = f'=INDEX(Service_Input!$D:$D, MATCH($C$4&"|"&{get_column_letter(c)}${start_row}&"|"&$A{r}, Service_Input!$E:$E, 0))'
        cell = ws_serv.cell(row=r, column=c, value=formula)
        cell.alignment = Alignment(horizontal='center')
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

apply_dark_theme_formatting(ws_serv, start_row, start_col, len(services), len(languages))

# ==========================================
# SHEET 3: PRODUCT DASHBOARD (The Smart One)
# ==========================================
ws_prod = wb.create_sheet("Product_Dashboard")

# Header Controls
ws_prod['B2'] = "PRODUCT AVAILABILITY MATRIX (Auto-Calculated)"
ws_prod['B2'].font = Font(size=16, bold=True, color="203764")
ws_prod['B4'] = "Select Region:"
ws_prod['C4'] = regions[0] 
ws_prod['C4'].font = Font(bold=True)
ws_prod['C4'].border = Border(bottom=Side(style='thin'))

# Dropdown
ws_prod.add_data_validation(dv) # Use same validation
dv.add(ws_prod['C4'])

# Legend Note
ws_prod['E4'] = "*Calculated based on dependencies (e.g. SSA requires ASR, GenAI, KaaS)"
ws_prod['E4'].font = Font(italic=True, size=9, color="555555")

# Grid Setup
start_row = 7
start_col = 2
ws_prod.cell(row=start_row, column=1, value="Language").font = Font(bold=True)

# Headers (Products)
for i, p in enumerate(products):
    c = start_col + i
    cell = ws_prod.cell(row=start_row, column=c, value=p)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color="203764", fill_type="solid") # Darker blue for products
    cell.alignment = Alignment(horizontal='center')
    ws_prod.column_dimensions[get_column_letter(c)].width = 20

# HELPER FUNCTION FOR FORMULA GENERATION
def get_lookup_formula(service_name, row_idx):
    # Returns the INDEX/MATCH part for a specific service
    return f'INDEX(Service_Input!$D:$D, MATCH($C$4&"|{service_name}|"&$A{row_idx}, Service_Input!$E:$E, 0))'

# Rows (Languages) & DEPENDENCY LOGIC
for i, lang in enumerate(languages):
    r = start_row + 1 + i
    ws_prod.cell(row=r, column=1, value=lang).font = Font(bold=True)

    # Service lookup formulas
    f_asr = get_lookup_formula("ASR", r)
    f_redaction = get_lookup_formula("Redaction", r)
    f_conv_facts = get_lookup_formula("Conversation Facts", r)
    f_gen_disp = get_lookup_formula("Gen AI Disposition", r)
    f_itn = get_lookup_formula("ITN", r)
    f_intent = get_lookup_formula("Intent (UniFit)", r)
    f_entity = get_lookup_formula("Entity (NER/Spacy)", r)
    f_gen_summary = get_lookup_formula("Gen AI Summary", r)
    f_proactive = get_lookup_formula("Pro-active Service", r)
    f_ssa_llm = get_lookup_formula("SSA-LLM", r)
    f_kaas = get_lookup_formula("KaaS", r)

    # Helper: check if status contains "Not Supported", "Limited Support", or "Cross-Region"
    def is_not_supported(f):
        return f'ISNUMBER(SEARCH("Not Supported", {f}))'

    def is_limited(f):
        return f'ISNUMBER(SEARCH("Limited Support", {f}))'

    def is_cross_region(f):
        return f'ISNUMBER(SEARCH("Cross-Region", {f}))'

    # 1. SSA (ASR, SSA-LLM, KaaS)
    # Logic: If ANY are "Not Supported" -> Not Supported
    # Else if ANY are "Limited" -> Limited Support (Cross-Region if any Cross-Region, else In-Region)
    # Else Full Support (Cross-Region if any Cross-Region, else In-Region)
    formula_ssa = (
        f'=IF(OR({is_not_supported(f_asr)}, {is_not_supported(f_ssa_llm)}, {is_not_supported(f_kaas)}), "Not Supported", '
        f'IF(OR({is_limited(f_asr)}, {is_limited(f_ssa_llm)}, {is_limited(f_kaas)}), '
        f'IF(OR({is_cross_region(f_asr)}, {is_cross_region(f_ssa_llm)}, {is_cross_region(f_kaas)}), "Limited Support (Cross-Region)", "Limited Support (In-Region)"), '
        f'IF(OR({is_cross_region(f_asr)}, {is_cross_region(f_ssa_llm)}, {is_cross_region(f_kaas)}), "Full Support (Cross-Region)", "Full Support (In-Region)")))'
    )
    ws_prod.cell(row=r, column=2, value=formula_ssa).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 2. RTGA (ASR, Intent, Entity, Gen AI Summary, Gen AI Disposition, Redaction, KaaS, Pro-active Service)
    formula_rtga = (
        f'=IF(OR({is_not_supported(f_asr)}, {is_not_supported(f_intent)}, {is_not_supported(f_entity)}, {is_not_supported(f_gen_summary)}, {is_not_supported(f_gen_disp)}, {is_not_supported(f_redaction)}, {is_not_supported(f_kaas)}, {is_not_supported(f_proactive)}), "Not Supported", '
        f'IF(OR({is_limited(f_asr)}, {is_limited(f_intent)}, {is_limited(f_entity)}, {is_limited(f_gen_summary)}, {is_limited(f_gen_disp)}, {is_limited(f_redaction)}, {is_limited(f_kaas)}, {is_limited(f_proactive)}), '
        f'IF(OR({is_cross_region(f_asr)}, {is_cross_region(f_intent)}, {is_cross_region(f_entity)}, {is_cross_region(f_gen_summary)}, {is_cross_region(f_gen_disp)}, {is_cross_region(f_redaction)}, {is_cross_region(f_kaas)}, {is_cross_region(f_proactive)}), "Limited Support (Cross-Region)", "Limited Support (In-Region)"), '
        f'IF(OR({is_cross_region(f_asr)}, {is_cross_region(f_intent)}, {is_cross_region(f_entity)}, {is_cross_region(f_gen_summary)}, {is_cross_region(f_gen_disp)}, {is_cross_region(f_redaction)}, {is_cross_region(f_kaas)}, {is_cross_region(f_proactive)}), "Full Support (Cross-Region)", "Full Support (In-Region)")))'
    )
    ws_prod.cell(row=r, column=3, value=formula_rtga).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 3. CIA (ASR, Gen AI Disposition, Redaction, Conversation Facts)
    formula_cia = (
        f'=IF(OR({is_not_supported(f_asr)}, {is_not_supported(f_gen_disp)}, {is_not_supported(f_redaction)}, {is_not_supported(f_conv_facts)}), "Not Supported", '
        f'IF(OR({is_limited(f_asr)}, {is_limited(f_gen_disp)}, {is_limited(f_redaction)}, {is_limited(f_conv_facts)}), '
        f'IF(OR({is_cross_region(f_asr)}, {is_cross_region(f_gen_disp)}, {is_cross_region(f_redaction)}, {is_cross_region(f_conv_facts)}), "Limited Support (Cross-Region)", "Limited Support (In-Region)"), '
        f'IF(OR({is_cross_region(f_asr)}, {is_cross_region(f_gen_disp)}, {is_cross_region(f_redaction)}, {is_cross_region(f_conv_facts)}), "Full Support (Cross-Region)", "Full Support (In-Region)")))'
    )
    ws_prod.cell(row=r, column=4, value=formula_cia).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 4. CRA (ASR, Redaction)
    formula_cra = (
        f'=IF(OR({is_not_supported(f_asr)}, {is_not_supported(f_redaction)}), "Not Supported", '
        f'IF(OR({is_limited(f_asr)}, {is_limited(f_redaction)}), '
        f'IF(OR({is_cross_region(f_asr)}, {is_cross_region(f_redaction)}), "Limited Support (Cross-Region)", "Limited Support (In-Region)"), '
        f'IF(OR({is_cross_region(f_asr)}, {is_cross_region(f_redaction)}), "Full Support (Cross-Region)", "Full Support (In-Region)")))'
    )
    ws_prod.cell(row=r, column=5, value=formula_cra).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

apply_dark_theme_formatting(ws_prod, start_row, start_col, len(products), len(languages))
ws_prod.column_dimensions['A'].width = 15
ws_serv.column_dimensions['A'].width = 15

# Save
wb.save(filename)
print(f"File '{filename}' created successfully.")