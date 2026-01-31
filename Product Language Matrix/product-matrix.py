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

# Service statuses (for Service Dashboard)
service_statuses = ["General Availability", "Limited Availability", "Not Supported"]

# Product statuses (for Product Dashboard)
product_statuses = [
    "Full Support (In-Region)",
    "Full Support (Cross-Region)",
    "Limited Availability",
    "Not Supported"
]

# Dependency Logic (For Reference/Comments)
# SSA: ASR, SSA-LLM, KaaS
# RTGA: ASR, Intent, Entity, Gen AI Summary, Gen AI Disposition, Redaction, KaaS, Pro-active Service
# CIA: ASR, Gen AI Disposition, Redaction, Conversation Facts
# CRA: ASR, Redaction

wb = Workbook()

# ==========================================
# SHEET 0: README (Instructions & Logic)
# ==========================================
ws_readme = wb.active
ws_readme.title = "README"

# Styling
title_font = Font(size=18, bold=True, color="203764")
section_font = Font(size=14, bold=True, color="2F5597")
subsection_font = Font(size=12, bold=True)
normal_font = Font(size=11)
italic_font = Font(size=11, italic=True, color="555555")

# Set column widths
ws_readme.column_dimensions['A'].width = 5
ws_readme.column_dimensions['B'].width = 80
ws_readme.column_dimensions['C'].width = 30

# Hide gridlines
ws_readme.sheet_view.showGridLines = False

# Content
readme_content = [
    ("", ""),
    ("PRODUCT & SERVICE AVAILABILITY MATRIX", "title"),
    ("", ""),
    ("OVERVIEW", "section"),
    ("This workbook helps you track and visualize the availability of AI services and products across different regions and languages.", "normal"),
    ("", ""),
    ("SHEETS IN THIS WORKBOOK", "section"),
    ("", ""),
    ("1. Service_Input (Source of Truth)", "subsection"),
    ("   - This is where you enter/update the availability status of each service", "normal"),
    ("   - Columns: Region, Service, Language, Status", "normal"),
    ("   - Status options: General Availability, Limited Availability, Not Supported", "normal"),
    ("   - All other sheets automatically pull data from here", "normal"),
    ("", ""),
    ("2. Service_Dashboard", "subsection"),
    ("   - Visual matrix showing service availability for a selected region", "normal"),
    ("   - Use the dropdown at the top to select a region", "normal"),
    ("   - Shows status for each service x language combination", "normal"),
    ("", ""),
    ("3. Product_Dashboard", "subsection"),
    ("   - Visual matrix showing product availability for a selected region", "normal"),
    ("   - Product status is AUTO-CALCULATED based on service dependencies", "normal"),
    ("   - Use the dropdown at the top to select a region", "normal"),
    ("", ""),
    ("SERVICE STATUS COLOR CODING", "section"),
    ("", ""),
    ("General Availability", "status_green"),
    ("   Service is fully mature and available in the region", "normal"),
    ("", ""),
    ("Limited Availability", "status_yellow"),
    ("   Service is available but not fully mature", "normal"),
    ("", ""),
    ("Not Supported", "status_red"),
    ("   Service is not available in the region", "normal"),
    ("", ""),
    ("PRODUCT STATUS COLOR CODING", "section"),
    ("", ""),
    ("Full Support (In-Region)", "status_dark_green"),
    ("   All required services are available locally (General Availability)", "normal"),
    ("", ""),
    ("Full Support (Cross-Region)", "status_light_green"),
    ("   Product works, but some services are fetched from another region", "normal"),
    ("", ""),
    ("Limited Availability", "status_yellow"),
    ("   Some services are Limited Availability or unavailable (but <=50%)", "normal"),
    ("", ""),
    ("Not Supported", "status_red"),
    ("   More than 50% of required services are unavailable globally", "normal"),
    ("", ""),
    ("PRODUCT DEPENDENCY LOGIC", "section"),
    ("", ""),
    ("Each product requires specific services to function:", "normal"),
    ("", ""),
    ("SSA (Speech & Sentiment Analytics)", "subsection"),
    ("   Required services: ASR, SSA-LLM, KaaS", "normal"),
    ("", ""),
    ("RTGA (Real-Time Guidance Agent)", "subsection"),
    ("   Required services: ASR, Intent (UniFit), Entity (NER/Spacy), Gen AI Summary,", "normal"),
    ("   Gen AI Disposition, Redaction, KaaS, Pro-active Service", "normal"),
    ("", ""),
    ("CIA (Customer Interaction Analytics)", "subsection"),
    ("   Required services: ASR, Gen AI Disposition, Redaction, Conversation Facts", "normal"),
    ("", ""),
    ("CRA (Call Recording Analytics)", "subsection"),
    ("   Required services: ASR, Redaction", "normal"),
    ("", ""),
    ("HOW PRODUCT STATUS IS CALCULATED", "section"),
    ("", ""),
    ("Priority order (highest to lowest):", "normal"),
    ("", ""),
    ("1. NOT SUPPORTED", "subsection"),
    ("   If more than 50% of required services are 'Not Supported' in ALL regions", "normal"),
    ("", ""),
    ("2. FULL SUPPORT (CROSS-REGION)", "subsection"),
    ("   If any required service is 'Not Supported' locally but available in another region", "normal"),
    ("", ""),
    ("3. LIMITED AVAILABILITY", "subsection"),
    ("   If any required service is 'Limited Availability' locally, OR", "normal"),
    ("   If any service is 'Not Supported' globally (but <=50% of dependencies)", "normal"),
    ("", ""),
    ("4. FULL SUPPORT (IN-REGION)", "subsection"),
    ("   If all required services are 'General Availability' in the selected region", "normal"),
    ("", ""),
    ("HOW TO USE", "section"),
    ("", ""),
    ("1. Go to 'Service_Input' sheet and update service statuses as needed", "normal"),
    ("2. Go to 'Service_Dashboard' to view service availability by region", "normal"),
    ("3. Go to 'Product_Dashboard' to see auto-calculated product availability", "normal"),
    ("4. Use the region dropdown on each dashboard to switch regions", "normal"),
    ("", ""),
    ("Note: Product status updates automatically when you change Service_Input data.", "italic"),
]

row = 1
for text, style in readme_content:
    cell = ws_readme.cell(row=row, column=2, value=text)

    if style == "title":
        cell.font = title_font
    elif style == "section":
        cell.font = section_font
    elif style == "subsection":
        cell.font = subsection_font
    elif style == "italic":
        cell.font = italic_font
    elif style == "status_green":
        cell.font = Font(size=11, bold=True)
        cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
        cell.font = Font(size=11, bold=True, color="FFFFFF")
    elif style == "status_dark_green":
        cell.font = Font(size=11, bold=True)
        cell.fill = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
        cell.font = Font(size=11, bold=True, color="FFFFFF")
    elif style == "status_light_green":
        cell.font = Font(size=11, bold=True)
        cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
        cell.font = Font(size=11, bold=True, color="FFFFFF")
    elif style == "status_yellow":
        cell.font = Font(size=11, bold=True)
        cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    elif style == "status_red":
        cell.font = Font(size=11, bold=True)
        cell.fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
        cell.font = Font(size=11, bold=True, color="FFFFFF")
    else:
        cell.font = normal_font

    row += 1

# ==========================================
# SHEET 1: SERVICE INPUT (The Source of Truth)
# ==========================================
ws_data = wb.create_sheet("Service_Input")
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
            # Random Status
            status = random.choices(service_statuses, weights=[45, 30, 25], k=1)[0]
            # Key: Region|Service|Language
            key_formula = f'=A{row_num}&"|"&B{row_num}&"|"&C{row_num}'
            ws_data.append([r, s, l, status, key_formula])
            row_num += 1

ws_data.column_dimensions['E'].hidden = True

# ==========================================
# STYLING HELPERS - Service Dashboard (3 statuses)
# ==========================================
def apply_service_formatting(ws, start_row, start_col, num_cols, num_rows):
    ws.sheet_view.showGridLines = False
    matrix_range = f"{get_column_letter(start_col)}{start_row+1}:{get_column_letter(start_col+num_cols-1)}{start_row+num_rows}"

    # Green (General Availability)
    green_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    dxf_green = DifferentialStyle(fill=green_fill, font=Font(color="FFFFFF"))
    rule_green = Rule(type="containsText", operator="containsText", text="General Availability", dxf=dxf_green)
    rule_green.formula = [f'NOT(ISERROR(SEARCH("General Availability", {matrix_range})))']
    ws.conditional_formatting.add(matrix_range, rule_green)

    # Yellow (Limited Availability)
    yellow_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    dxf_yellow = DifferentialStyle(fill=yellow_fill)
    rule_yellow = Rule(type="containsText", operator="containsText", text="Limited Availability", dxf=dxf_yellow)
    rule_yellow.formula = [f'NOT(ISERROR(SEARCH("Limited Availability", {matrix_range})))']
    ws.conditional_formatting.add(matrix_range, rule_yellow)

    # Red (Not Supported)
    red_fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
    dxf_red = DifferentialStyle(fill=red_fill, font=Font(color="FFFFFF"))
    rule_red = Rule(type="containsText", operator="containsText", text="Not Supported", dxf=dxf_red)
    rule_red.formula = [f'NOT(ISERROR(SEARCH("Not Supported", {matrix_range})))']
    ws.conditional_formatting.add(matrix_range, rule_red)

# ==========================================
# STYLING HELPERS - Product Dashboard (4 statuses)
# ==========================================
def apply_product_formatting(ws, start_row, start_col, num_cols, num_rows):
    ws.sheet_view.showGridLines = False
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

    # Yellow (Limited Availability)
    yellow_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    dxf_yellow = DifferentialStyle(fill=yellow_fill)
    rule_yellow = Rule(type="containsText", operator="containsText", text="Limited Availability", dxf=dxf_yellow)
    rule_yellow.formula = [f'NOT(ISERROR(SEARCH("Limited Availability", {matrix_range})))']
    ws.conditional_formatting.add(matrix_range, rule_yellow)

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
ws_serv['C4'] = regions[0]  # Default
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

apply_service_formatting(ws_serv, start_row, start_col, len(services), len(languages))

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
ws_prod.add_data_validation(dv)  # Use same validation
dv.add(ws_prod['C4'])

# Legend Note
ws_prod['E4'] = "*Cross-Region = service not available locally but available in another region"
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
    cell.fill = PatternFill(start_color="203764", fill_type="solid")
    cell.alignment = Alignment(horizontal='center')
    ws_prod.column_dimensions[get_column_letter(c)].width = 25

# Total number of regions (for Cross-Region check)
num_regions = len(regions)

# HELPER FUNCTIONS FOR FORMULA GENERATION
def get_local_status(service_name, row_idx):
    """Get status of service in selected region"""
    return f'INDEX(Service_Input!$D:$D, MATCH($C$4&"|{service_name}|"&$A{row_idx}, Service_Input!$E:$E, 0))'

def get_global_not_supported_count(service_name, row_idx):
    """Count how many regions have this service as 'Not Supported' for the language"""
    # COUNTIFS to count rows where Service matches, Language matches, and Status is "Not Supported"
    return f'COUNTIFS(Service_Input!$B:$B,"{service_name}",Service_Input!$C:$C,$A{row_idx},Service_Input!$D:$D,"Not Supported")'

def is_not_supported_locally(service_name, row_idx):
    """Check if service is Not Supported in selected region"""
    return f'{get_local_status(service_name, row_idx)}="Not Supported"'

def is_limited_locally(service_name, row_idx):
    """Check if service is Limited Availability in selected region"""
    return f'{get_local_status(service_name, row_idx)}="Limited Availability"'

def is_not_supported_globally(service_name, row_idx):
    """Check if service is Not Supported in ALL regions (count = num_regions)"""
    return f'{get_global_not_supported_count(service_name, row_idx)}={num_regions}'

def needs_cross_region(service_name, row_idx):
    """Check if service needs cross-region (Not Supported locally but available somewhere)"""
    # Not supported locally AND not supported globally count < num_regions (meaning available somewhere)
    return f'AND({is_not_supported_locally(service_name, row_idx)}, {get_global_not_supported_count(service_name, row_idx)}<{num_regions})'

# Product dependencies
product_deps = {
    "SSA": ["ASR", "SSA-LLM", "KaaS"],
    "RTGA": ["ASR", "Intent (UniFit)", "Entity (NER/Spacy)", "Gen AI Summary", "Gen AI Disposition", "Redaction", "KaaS", "Pro-active Service"],
    "CIA": ["ASR", "Gen AI Disposition", "Redaction", "Conversation Facts"],
    "CRA": ["ASR", "Redaction"]
}

# Rows (Languages) & DEPENDENCY LOGIC
for i, lang in enumerate(languages):
    r = start_row + 1 + i
    ws_prod.cell(row=r, column=1, value=lang).font = Font(bold=True)

    for prod_idx, (product, deps) in enumerate(product_deps.items()):
        col = start_col + prod_idx
        total_deps = len(deps)
        threshold = total_deps / 2  # More than 50%

        # Build conditions for this product
        # 1. Count dependencies Not Supported globally (in ALL regions)
        globally_not_supported_conditions = [f'IF({is_not_supported_globally(dep, r)},1,0)' for dep in deps]
        count_globally_not_supported = f'({"+".join(globally_not_supported_conditions)})'

        # 2. Any dependency needs Cross-Region → Full Support (Cross-Region)
        needs_cross = [needs_cross_region(dep, r) for dep in deps]

        # 3. Any dependency Limited locally OR any dependency not supported globally (but ≤50%) → Limited Availability
        limited_locally = [is_limited_locally(dep, r) for dep in deps]
        globally_not_supported = [is_not_supported_globally(dep, r) for dep in deps]

        # 4. Otherwise → Full Support (In-Region)

        # Build the formula
        # Priority:
        # 1. >50% dependencies Not Supported globally → Not Supported
        # 2. Any needs Cross-Region → Full Support (Cross-Region)
        # 3. Any Limited locally OR any Not Supported globally (≤50%) → Limited Availability
        # 4. All General Availability locally → Full Support (In-Region)
        formula = (
            f'=IF({count_globally_not_supported}>{threshold}, "Not Supported", '
            f'IF(OR({",".join(needs_cross)}), "Full Support (Cross-Region)", '
            f'IF(OR({",".join(limited_locally)},{",".join(globally_not_supported)}), "Limited Availability", '
            f'"Full Support (In-Region)")))'
        )

        cell = ws_prod.cell(row=r, column=col, value=formula)
        cell.alignment = Alignment(horizontal='center')
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

apply_product_formatting(ws_prod, start_row, start_col, len(products), len(languages))
ws_prod.column_dimensions['A'].width = 15
ws_serv.column_dimensions['A'].width = 15

# Save
wb.save(filename)
print(f"File '{filename}' created successfully.")
