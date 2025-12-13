from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
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
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

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
    
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)

def add_role_header(document, title, company_location, date):
    # Table for layout: Title (Left) | Date (Right)
    table = document.add_table(rows=1, cols=2)
    table.allow_autofit = False

    # Set column widths with proper cell width assignment
    table.columns[0].width = Inches(5.0)
    table.columns[1].width = Inches(2.0)

    # Also set the cell widths directly to ensure they're respected
    table.cell(0, 0).width = Inches(5.0)
    table.cell(0, 1).width = Inches(2.0)

    # Title Cell
    cell_1 = table.cell(0, 0)
    p1 = cell_1.paragraphs[0]
    r1 = p1.add_run(title)
    r1.bold = True
    r1.font.size = Pt(11)
    r1.font.color.rgb = RGBColor(46, 64, 83) # Dark Blue/Grey

    # Date Cell
    cell_2 = table.cell(0, 1)
    p2 = cell_2.paragraphs[0]
    p2.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    r2 = p2.add_run(date)
    r2.bold = True
    r2.font.size = Pt(11)

    # Company Line below title
    p_sub = document.add_paragraph()
    r_sub = p_sub.add_run(company_location)
    r_sub.italic = True
    p_sub.paragraph_format.space_after = Pt(2)

# --- HEADER SECTION ---
header = doc.add_paragraph()
header.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
name = header.add_run(personal_info['name'])
name.bold = True
name.font.size = Pt(22)
name.font.color.rgb = RGBColor(46, 64, 83) # Dark Slate Blue

contact = doc.add_paragraph()
contact.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
contact.add_run(f"{personal_info['location']} | {personal_info['phone']} | {personal_info['email']}")
contact.add_run(f"\n{personal_info['linkedin']} | {personal_info['github']}")
contact.paragraph_format.space_after = Pt(10)

# --- PROFESSIONAL SUMMARY ---
add_section_header(doc, section_labels['professional_summary'])
summary = doc.add_paragraph(summary_data['text'])
summary.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

# --- KEY SKILLS ---
add_section_header(doc, section_labels['key_skills'])

for skill_category in skills_data['categories']:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    cat_run = p.add_run(skill_category['category'] + " ")
    cat_run.bold = True
    p.add_run(skill_category['items'])

# --- PROFESSIONAL EXPERIENCE ---
add_section_header(doc, section_labels['professional_experience'])

for role in experience_data['roles']:
    add_role_header(doc, role['title'], role['company_location'], role['date'])
    for bullet in role['bullets']:
        doc.add_paragraph(bullet, style='List Bullet')

# Earlier Career
add_section_header(doc, section_labels['earlier_career'])
doc.add_paragraph(experience_data['earlier_career']['title'], style='Normal')
doc.add_paragraph(experience_data['earlier_career']['description'])

# --- EDUCATION ---
add_section_header(doc, section_labels['education'])
num_degrees = len(education_data['degrees'])
edu_table = doc.add_table(rows=num_degrees, cols=2)
edu_table.autofit = False
edu_table.columns[0].width = Inches(5.0)
edu_table.columns[1].width = Inches(2.0)

for idx, degree in enumerate(education_data['degrees']):
    edu_table.cell(idx, 0).text = f"{degree['degree']}\n{degree['institution']}"

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