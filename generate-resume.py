from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import os
import json

# --- LOAD DATA FROM JSON FILES ---
def load_json(filename):
    filepath = os.path.join('data', filename)
    with open(filepath, 'r', encoding='utf-8') as f:
        return json.load(f)

# Load all data files
personal_info = load_json('personal-info.json')
summary_data = load_json('summary.json')
skills_data = load_json('skills.json')
experience_data = load_json('experience.json')
education_data = load_json('education.json')
certifications_data = load_json('certifications.json')
section_labels = load_json('section-labels.json')

# Initialize Document
doc = Document()

# --- PAGE LAYOUT SETTINGS ---
# Set margins to 0.5 inches for all sides
sections = doc.sections
for section in sections:
    section.top_margin = Inches(0.2)
    section.bottom_margin = Inches(0.2)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

# --- DOCUMENT HEADER ---
# Enable different first page header and use first_page_header for first page only
for section in sections:
    section.different_first_page_header_footer = True

# Access the first page header (appears only on first page)
header = sections[0].first_page_header
header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()

# Create 3-column table in header: Email/Phone | Name | LinkedIn/GitHub
header_table = header.add_table(rows=1, cols=3, width=Inches(7.5))
header_table.allow_autofit = False
header_table.columns[0].width = Inches(2.5)
header_table.columns[1].width = Inches(2.5)
header_table.columns[2].width = Inches(2.5)

# Left column: Email and Phone (stacked)
left_cell = header_table.cell(0, 0)
left_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
left_p = left_cell.paragraphs[0]
left_p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
email_run = left_p.add_run(personal_info['email'])
email_run.font.size = Pt(9)
left_p.add_run('\n')
phone_run = left_p.add_run(personal_info['phone'])
phone_run.font.size = Pt(9)

# Center column: Name
center_cell = header_table.cell(0, 1)
center_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
center_p = center_cell.paragraphs[0]
center_p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
name_run = center_p.add_run(personal_info['name'].upper())
name_run.bold = True
name_run.font.size = Pt(12)
name_run.font.color.rgb = RGBColor(46, 64, 83)

# Right column: LinkedIn and GitHub (stacked)
right_cell = header_table.cell(0, 2)
right_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
right_p = right_cell.paragraphs[0]
right_p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
linkedin_run = right_p.add_run(personal_info['linkedin'])
linkedin_run.font.size = Pt(9)
right_p.add_run('\n')
github_run = right_p.add_run(personal_info['github'])
github_run.font.size = Pt(9)

# Remove the empty first paragraph in header to reduce space
if header_para.text == '':
    p_element = header_para._element
    p_element.getparent().remove(p_element)

# --- STYLES & FORMATTING ---
style = doc.styles['Normal']
font = style.font
font.name = 'Arial'
font.size = Pt(10)

# Function to add a shaded background to headings (Sanjay Style)
def add_section_header(document, text):
    p = document.add_paragraph()
    runner = p.add_run(text.upper())
    runner.bold = True
    runner.font.size = Pt(12)
    runner.font.color.rgb = RGBColor(255, 255, 255)  # White text

    # Add dark blue shading (Professional, similar to high-end templates)
    shading_elm = parse_xml(r'<w:shd {} w:fill="2E4053"/>'.format(nsdecls('w')))
    p._p.get_or_add_pPr().append(shading_elm)

    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)

def add_role_header(document, title, company_location, date):
    # Table for layout: Title (Left) | Date (Right)
    table = document.add_table(rows=1, cols=2)
    table.allow_autofit = False

    # Set table width to full available width (8.5" - 0.5" - 0.5" = 7.5")
    table.width = Inches(7.5)

    # Set column widths with proper cell width assignment
    table.columns[0].width = Inches(5.5)
    table.columns[1].width = Inches(2.0)

    # Also set the cell widths directly to ensure they're respected
    table.cell(0, 0).width = Inches(5.5)
    table.cell(0, 1).width = Inches(2.0)

    # Title Cell
    cell_1 = table.cell(0, 0)
    p1 = cell_1.paragraphs[0]
    p1.paragraph_format.space_before = Pt(1)
    p1.paragraph_format.space_after = Pt(1)
    r1 = p1.add_run(title)
    r1.bold = True
    r1.font.size = Pt(11)
    r1.font.color.rgb = RGBColor(46, 64, 83) # Dark Blue/Grey

    # Date Cell
    cell_2 = table.cell(0, 1)

    # Set cell vertical alignment
    cell_2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    p2 = cell_2.paragraphs[0]
    p2.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    r2 = p2.add_run(date)
    r2.bold = True
    r2.font.size = Pt(11)

    # Company Line below title (only if provided - for backward compatibility)
    if company_location:
        p_sub = document.add_paragraph()
        r_sub = p_sub.add_run(company_location)
        r_sub.italic = True
        p_sub.paragraph_format.space_after = Pt(1)

# --- PROFESSIONAL SUMMARY ---
add_section_header(doc, section_labels['professional_summary'])
summary = doc.add_paragraph(summary_data['text'])
summary.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

# --- KEY SKILLS ---
add_section_header(doc, section_labels['key_skills'])

for skill_category in skills_data['categories']:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    cat_run = p.add_run(skill_category['category'] + " ")
    cat_run.bold = True
    p.add_run(skill_category['items'])

# --- PROFESSIONAL EXPERIENCE ---
add_section_header(doc, section_labels['professional_experience'])

for company in experience_data['companies']:
    # Add company header
    company_p = doc.add_paragraph()
    company_run = company_p.add_run(f"{company['company']} | {company['location']}")
    company_run.bold = True
    company_run.font.size = Pt(12)
    company_run.font.color.rgb = RGBColor(46, 64, 83)
    company_p.paragraph_format.space_before = Pt(3)
    company_p.paragraph_format.space_after = Pt(1)

    # Add roles under this company
    for role in company['roles']:
        add_role_header(doc, role['title'], "", role['date'])
        for bullet in role['bullets']:
            bullet_p = doc.add_paragraph(bullet, style='List Bullet')
            bullet_p.paragraph_format.space_after = Pt(1)
            bullet_p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

# Earlier Career
add_section_header(doc, section_labels['earlier_career'])
ec_title = doc.add_paragraph(experience_data['earlier_career']['title'], style='Normal')
ec_title.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
ec_desc = doc.add_paragraph(experience_data['earlier_career']['description'])
ec_desc.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

# --- EDUCATION ---
add_section_header(doc, section_labels['education'])
for degree in education_data['degrees']:
    edu_p = doc.add_paragraph()
    edu_p.paragraph_format.space_after = Pt(1)

    # Degree name in bold
    degree_run = edu_p.add_run(degree['degree'])
    degree_run.bold = True

    # Institution in regular text on same line
    edu_p.add_run(f", {degree['institution']}")

# --- CERTIFICATIONS ---
add_section_header(doc, section_labels['certifications'])
for cert in certifications_data['certifications']:
    doc.add_paragraph(cert)

# Save document with versioning
output_dir = "generated"
# Create filename from name (replace spaces with underscores, remove special chars)
name_for_file = personal_info['name'].replace(' ', '_').title()
base_filename = f"{name_for_file}_Resume"

# Create the generate folder if it doesn't exist
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Find the next available version number
version = 1
while True:
    filename = f"{base_filename}_{version}.docx"
    filepath = os.path.join(output_dir, filename)
    if not os.path.exists(filepath):
        break
    version += 1

# Save the document
doc.save(filepath)
print(f"Resume generated successfully: {filepath}")