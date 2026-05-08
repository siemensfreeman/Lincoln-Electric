#!/usr/bin/env python3
"""
Generate lincoln_AEP.pptx matching the superior 3-column layout design
"""

import sys
from io import BytesIO
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
except ImportError:
    print("Error: python-pptx not installed. Install with: pip install python-pptx")
    sys.exit(1)

# Colors matching Image 2
ACCENT_COLORS = {
    'r5': RGBColor(100, 116, 139),    # Slate
    'r4': RGBColor(59, 130, 246),     # Blue
    'r3': RGBColor(139, 92, 246),     # Purple
    'r2': RGBColor(6, 182, 212),      # Cyan
    'r1': RGBColor(0, 169, 157),      # Teal
}

DARK_BG = RGBColor(7, 16, 31)
LIGHT_TEXT = RGBColor(255, 255, 255)
MUTED_TEXT = RGBColor(148, 163, 184)
TAG_BG = RGBColor(20, 30, 50)

ROWS = [
    {
        'row_id': 'r5',
        'persona': 'CHARLES JOHNSON • FACILITIES & MAINTENANCE',
        'pain': 'SharePoint for maintenance — ungoverned, unauditable, no workflow enforcement or policy compliance',
        'cap_num': 'CAP 05',
        'cap_title': 'ENTERPRISE AI & WORKFLOW GOVERNANCE',
        'cap_product': 'Platform-wide • Graph Studio contributes lineage, audit & access control',
        'cap_tags': ['Every decision traceable', 'Audit trails', 'Policy enforcement', 'Data lineage', 'Layer-level security'],
        'outcome_icon': '◉',
        'outcome_title': 'Governed, Auditable Workflows',
        'outcome_desc': 'Every maintenance action logged, policy-enforced, and fully auditable — SharePoint replaced with a governed, compliant platform',
    },
    {
        'row_id': 'r4',
        'persona': 'CHARLES JOHNSON • FACILITIES & MAINTENANCE',
        'pain': 'Work orders disconnected from drawings & asset records — manual stitching required every time',
        'cap_num': 'CAP 04',
        'cap_title': 'AGENTIC PROCESS ORCHESTRATION',
        'cap_product': 'App Studio (Mendix) • Hybrid workforce coordination',
        'cap_tags': ['Governed workflow apps', 'Human + agent coordination', 'Write-back via MCP', 'Replaces ad-hoc tools'],
        'outcome_icon': '◉',
        'outcome_title': 'Work Orders ↔ Drawings ↔ Assets Unified',
        'outcome_desc': 'Mendix orchestrates the full workflow — work orders, drawings, and asset records connected in one automated, traceable flow',
    },
    {
        'row_id': 'r3',
        'persona': 'CHARLES JOHNSON • FACILITIES & MAINTENANCE',
        'pain': 'Automation needed outside Teamcenter — no low-code path, no way to extend scope to facilities & HR data',
        'cap_num': 'CAP 03',
        'cap_title': 'AI ASSISTED APP DEVELOPMENT',
        'cap_product': 'App Studio (Mendix) • Low-code agentic applications',
        'cap_tags': ['Low-code agents', 'Enterprise deployment', 'Connects non-TC data', 'No scope limits', 'Weeks to production'],
        'outcome_icon': '◉',
        'outcome_title': 'Automation Beyond Teamcenter',
        'outcome_desc': 'Low-code Mendix apps connect facilities, maintenance & HR data alongside TC — no scope limits, production-ready in weeks',
    },
    {
        'row_id': 'r2',
        'persona': 'JEANNIE WHITED • MULTI-SYSTEM REPORTING',
        'pain': 'Facility-wide retrieval needs AI-powered reasoning across asset & operational data — no intelligent layer exists today',
        'cap_num': 'CAP 02',
        'cap_title': 'AI / ML MODEL DEVELOPMENT',
        'cap_product': 'AI Studio • Precision AI grounded in the Knowledge Graph',
        'cap_tags': ['Governed ML models', 'Semantic grounding', 'Predictive analytics', 'Explainable AI'],
        'outcome_icon': '◉',
        'outcome_title': 'AI-Powered Facility Intelligence',
        'outcome_desc': 'ML models grounded in real asset & operational data — predictive maintenance, anomaly detection, explainable outputs traceable to source',
    },
    {
        'row_id': 'r1',
        'persona': 'JEANNIE WHITED • MULTI-SYSTEM REPORTING',
        'pain': 'Manual data aggregation — one ETL per question, weeks of engineering per new cross-domain report. External integration complexity — data moves, context breaks, no unified cross-domain view',
        'cap_num': 'CAP 01',
        'cap_title': 'INSTITUTIONAL KNOWLEDGE GRAPH',
        'cap_product': 'Rapidminer Graph Studio • The data foundation every other capability depends on',
        'cap_tags': ['One query • Any system', 'No ETL per question', 'Data stays in place', 'W3C RDF / SPARQL / OWL', 'MPP • 10s of billions', 'Scales across every domain'],
        'outcome_icon': '◉',
        'outcome_title': 'One Query • Any System • No ETL',
        'outcome_desc': 'Semantic overlay connects all sources — cross-domain questions answered in seconds, data stays where it lives, expand without re-architecture',
        'foundation': '◆ Rapidminer Graph Studio — Foundation Layer',
    },
]

def create_presentation():
    """Create PowerPoint with 3-column layout (Pain | Capability | Outcome)"""
    prs = Presentation()
    prs.slide_width = Inches(16)
    prs.slide_height = Inches(10)
    
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = DARK_BG
    
    # Siemens logo placeholder (teal box)
    logo_box = slide.shapes.add_shape(1, Inches(0.3), Inches(0.35), Inches(0.9), Inches(0.45))
    logo_box.fill.solid()
    logo_box.fill.fore_color.rgb = ACCENT_COLORS['r1']
    logo_box.line.color.rgb = ACCENT_COLORS['r1']
    logo_tf = logo_box.text_frame
    logo_p = logo_tf.paragraphs[0]
    logo_p.text = "SIEMENS"
    logo_p.font.size = Pt(14)
    logo_p.font.bold = True
    logo_p.font.color.rgb = DARK_BG
    logo_p.alignment = PP_ALIGN.CENTER
    logo_tf.vertical_anchor = 1  # Middle
    
    # Eyebrow + Title
    eyebrow_box = slide.shapes.add_textbox(Inches(1.4), Inches(0.35), Inches(8), Inches(0.25))
    tf = eyebrow_box.text_frame
    p = tf.paragraphs[0]
    p.text = "PLATFORM RESPONSE"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = ACCENT_COLORS['r1']
    
    title_box = slide.shapes.add_textbox(Inches(1.4), Inches(0.6), Inches(13), Inches(0.7))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "The Agentic Enterprise Platform — From Pain to Outcome"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = LIGHT_TEXT
    
    # Column headers
    headers = ['BUSINESS PAIN POINT', 'AEP CAPABILITY', 'PLATFORM OUTCOME']
    header_left = [Inches(0.3), Inches(5.8), Inches(11.3)]
    
    for i, header in enumerate(headers):
        hbox = slide.shapes.add_textbox(header_left[i], Inches(1.5), Inches(5), Inches(0.25))
        hf = hbox.text_frame
        hp = hf.paragraphs[0]
        hp.text = header
        hp.font.size = Pt(9)
        hp.font.bold = True
        hp.font.color.rgb = MUTED_TEXT
    
    # Row height and spacing
    row_height = Inches(1.55)
    row_top_start = Inches(1.85)
    row_gap = Inches(0.08)
    
    for idx, row_data in enumerate(ROWS):
        row_top = row_top_start + (idx * (row_height + row_gap))
        create_row(slide, row_top, row_height, row_data)
    
    return prs

def create_row(slide, top, height, row_data):
    """Create one row with 3 columns"""
    color = ACCENT_COLORS[row_data['row_id']]
    
    # LEFT: Pain Point (with accent bar)
    accent_bar = slide.shapes.add_shape(1, Inches(0.3), top, Inches(0.08), height)
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = color
    accent_bar.line.width = Pt(0)
    
    pain_box = slide.shapes.add_textbox(Inches(0.5), top + Inches(0.1), Inches(5.1), height - Inches(0.2))
    tf = pain_box.text_frame
    tf.word_wrap = True
    
    # Persona
    p = tf.paragraphs[0]
    p.text = row_data['persona']
    p.font.size = Pt(7)
    p.font.color.rgb = MUTED_TEXT
    p.space_after = Pt(4)
    
    # Pain text
    p = tf.add_paragraph()
    p.text = row_data['pain']
    p.font.size = Pt(8)
    p.font.color.rgb = LIGHT_TEXT
    p.line_spacing = 1.15
    
    # CENTER: Capability
    cap_box = slide.shapes.add_textbox(Inches(5.8), top + Inches(0.1), Inches(5.3), height - Inches(0.2))
    tf = cap_box.text_frame
    tf.word_wrap = True
    
    # CAP number and title
    p = tf.paragraphs[0]
    p.text = f"{row_data['cap_num']}  {row_data['cap_title']}"
    p.font.size = Pt(8)
    p.font.bold = True
    p.font.color.rgb = color
    p.space_after = Pt(3)
    
    # Product
    p = tf.add_paragraph()
    p.text = row_data['cap_product']
    p.font.size = Pt(7)
    p.font.color.rgb = MUTED_TEXT
    p.space_after = Pt(4)
    
    # Tags
    for tag in row_data['cap_tags']:
        p = tf.add_paragraph()
        p.text = f"  {tag}"
        p.font.size = Pt(7)
        p.font.color.rgb = color
        p.space_after = Pt(1)
    
    # Foundation (r1 only)
    if row_data.get('foundation'):
        p = tf.add_paragraph()
        p.text = ""
        p.space_after = Pt(2)
        p = tf.add_paragraph()
        p.text = f"  {row_data['foundation']}"
        p.font.size = Pt(7)
        p.font.bold = True
        p.font.color.rgb = color
    
    # RIGHT: Outcome
    outcome_box = slide.shapes.add_textbox(Inches(11.3), top + Inches(0.1), Inches(4.4), height - Inches(0.2))
    tf = outcome_box.text_frame
    tf.word_wrap = True
    
    # Icon + Title
    p = tf.paragraphs[0]
    p.text = f"{row_data['outcome_icon']}  {row_data['outcome_title']}"
    p.font.size = Pt(8)
    p.font.bold = True
    p.font.color.rgb = color
    p.space_after = Pt(3)
    
    # Description
    p = tf.add_paragraph()
    p.text = row_data['outcome_desc']
    p.font.size = Pt(7)
    p.font.color.rgb = MUTED_TEXT
    p.line_spacing = 1.15

if __name__ == '__main__':
    prs = create_presentation()
    prs.save('lincoln_AEP.pptx')
    print("✓ Presentation created: lincoln_AEP.pptx")
