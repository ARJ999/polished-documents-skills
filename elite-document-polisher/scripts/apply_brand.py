#!/usr/bin/env python3
"""
Elite Document Polisher - World-Class Brand Style Application Tool

The most sophisticated document styling system for transforming DOCX documents
into professionally-styled masterpieces using 10 premium brand identities.

Features:
- Single brand application
- Multi-brand batch generation
- All-brand variant creation
- Intelligent content preservation

Usage:
    # Single brand
    python apply_brand.py <input.docx> <brand_name> <output.docx>

    # Multiple brands (batch)
    python apply_brand.py <input.docx> <brand1,brand2,brand3> <output_prefix>

    # All brands
    python apply_brand.py <input.docx> all <output_prefix>

    # List available brands
    python apply_brand.py --list

Examples:
    python apply_brand.py report.docx mckinsey polished_report.docx
    python apply_brand.py proposal.docx mckinsey,deloitte,stripe proposal
    python apply_brand.py document.docx all document

Available Brands:
    economist  - The Economist (Editorial, serif typography)
    mckinsey   - McKinsey & Company (Consulting, executive presence)
    deloitte   - Deloitte (Consulting, modern teal)
    kpmg       - KPMG (Consulting, corporate blue)
    stripe     - Stripe (Tech, developer-focused)
    apple      - Apple (Tech, minimalist premium)
    ibm        - IBM (Tech, enterprise authority)
    linear     - Linear (Tech, modern purple)
    notion     - Notion (Productivity, clean blue)
    figma      - Figma (Design, vibrant multi-color)

Author: Elite Document Polisher
Version: 2.0.0
"""

import json
import sys
import os
from pathlib import Path
from typing import List, Dict, Optional

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches, Twips
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("=" * 70)
    print("ERROR: python-docx is required")
    print("=" * 70)
    print("\nInstall with:")
    print("  pip install python-docx")
    print("\nOr in a virtual environment:")
    print("  python -m venv venv")
    print("  source venv/bin/activate  # On Windows: venv\\Scripts\\activate")
    print("  pip install python-docx")
    print("=" * 70)
    sys.exit(1)


# ============================================================================
# COLOR UTILITIES
# ============================================================================

def hex_to_rgb(hex_color: str) -> RGBColor:
    """Convert hex color (#RRGGBB) to RGBColor object."""
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return RGBColor(r, g, b)


# ============================================================================
# FONT UTILITIES
# ============================================================================

def set_font_name(run, font_name: str):
    """Set font name with full XML compatibility for cross-platform rendering."""
    run.font.name = font_name
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:cs'), font_name)


def set_style_font(style, font_name: str):
    """Set font name on a document style element."""
    style.font.name = font_name
    element = style.element
    rPr = element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:cs'), font_name)


# ============================================================================
# BRAND CONFIGURATION
# ============================================================================

def get_script_dir() -> Path:
    """Get the directory containing this script."""
    return Path(__file__).parent.parent


def load_brand_config(brand_name: str) -> Dict:
    """Load brand configuration from brand-mapping.json."""
    script_dir = get_script_dir()
    config_path = script_dir / 'templates' / 'brand-mapping.json'

    if not config_path.exists():
        raise FileNotFoundError(f"Brand configuration not found at: {config_path}")

    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    brand_name_lower = brand_name.lower()

    if brand_name_lower not in config['brands']:
        available = ', '.join(sorted(config['brands'].keys()))
        raise ValueError(
            f"Brand '{brand_name}' not found.\n"
            f"Available brands: {available}"
        )

    return config['brands'][brand_name_lower]


def get_all_brands() -> List[str]:
    """Get list of all available brand names."""
    script_dir = get_script_dir()
    config_path = script_dir / 'templates' / 'brand-mapping.json'

    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    return list(config['brands'].keys())


# ============================================================================
# DOCUMENT STYLING ENGINE
# ============================================================================

def apply_brand_to_docx(input_path: str, brand_name: str, output_path: str) -> str:
    """
    Apply brand styling to a DOCX file by recreating it with proper styles.

    Args:
        input_path: Path to source DOCX file
        brand_name: Name of the brand to apply
        output_path: Path for the styled output file

    Returns:
        Path to the created output file
    """
    brand = load_brand_config(brand_name)

    # Load source document
    source_doc = Document(input_path)

    # Create new document
    doc = Document()

    # Extract brand settings
    heading_font = brand['typography']['headingFont']
    body_font = brand['typography']['bodyFont']

    h1_style = brand['styles']['h1']
    h2_style = brand['styles']['h2']
    h3_style = brand['styles']['h3']
    body_style = brand['styles']['body']
    caption_style = brand['styles'].get('caption', {'size': 9, 'color': '#666666'})

    primary_color = hex_to_rgb(brand['colors']['primary'])
    text_color = hex_to_rgb(brand['colors']['textPrimary'])
    secondary_color = hex_to_rgb(brand['colors']['textSecondary'])
    accent_color = hex_to_rgb(brand['colors']['accent'])

    # ========================================================================
    # APPLY DOCUMENT STYLES
    # ========================================================================

    styles = doc.styles

    # Normal style (base for body text)
    normal = styles['Normal']
    set_style_font(normal, body_font)
    normal.font.size = Pt(body_style['size'])
    normal.font.color.rgb = text_color
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(10)
    normal.paragraph_format.line_spacing = 1.15

    # Title style (document title, centered)
    title_style = styles['Title']
    set_style_font(title_style, heading_font)
    title_style.font.size = Pt(h1_style['size'] + 8)
    title_style.font.bold = h1_style.get('bold', True)
    title_style.font.color.rgb = primary_color
    title_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_style.paragraph_format.space_before = Pt(0)
    title_style.paragraph_format.space_after = Pt(24)

    # Heading 1 - Chapter/Major section headings
    h1 = styles['Heading 1']
    set_style_font(h1, heading_font)
    h1.font.size = Pt(h1_style['size'])
    h1.font.bold = h1_style.get('bold', True)
    h1.font.color.rgb = hex_to_rgb(h1_style['color'])
    h1.paragraph_format.space_before = Pt(0)
    h1.paragraph_format.space_after = Pt(18)
    h1.paragraph_format.line_spacing = 1.0
    h1.paragraph_format.keep_with_next = True

    # Heading 2 - Section headings
    h2 = styles['Heading 2']
    set_style_font(h2, heading_font)
    h2.font.size = Pt(h2_style['size'])
    h2.font.bold = h2_style.get('bold', True)
    h2.font.color.rgb = hex_to_rgb(h2_style['color'])
    h2.paragraph_format.space_before = Pt(18)
    h2.paragraph_format.space_after = Pt(8)
    h2.paragraph_format.line_spacing = 1.0
    h2.paragraph_format.keep_with_next = True

    # Heading 3 - Sub-section headings
    h3 = styles['Heading 3']
    set_style_font(h3, heading_font)
    h3.font.size = Pt(h3_style['size'])
    h3.font.bold = h3_style.get('bold', True)
    h3.font.color.rgb = hex_to_rgb(h3_style['color'])
    h3.paragraph_format.space_before = Pt(12)
    h3.paragraph_format.space_after = Pt(6)
    h3.paragraph_format.line_spacing = 1.0
    h3.paragraph_format.keep_with_next = True

    # List Bullet style
    try:
        list_bullet = styles['List Bullet']
        set_style_font(list_bullet, body_font)
        list_bullet.font.size = Pt(body_style['size'])
        list_bullet.font.color.rgb = text_color
        list_bullet.paragraph_format.space_before = Pt(2)
        list_bullet.paragraph_format.space_after = Pt(2)
    except KeyError:
        pass

    # List Number style
    try:
        list_number = styles['List Number']
        set_style_font(list_number, body_font)
        list_number.font.size = Pt(body_style['size'])
        list_number.font.color.rgb = text_color
        list_number.paragraph_format.space_before = Pt(2)
        list_number.paragraph_format.space_after = Pt(2)
    except KeyError:
        pass

    # ========================================================================
    # SET PAGE MARGINS
    # ========================================================================

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.25)
        section.right_margin = Inches(1.25)

    # ========================================================================
    # COPY CONTENT WITH STYLING
    # ========================================================================

    first_h1_seen = False

    for para in source_doc.paragraphs:
        style_name = para.style.name if para.style else 'Normal'
        para_text = para.text.strip()

        # Skip empty paragraphs but preserve structure
        if not para_text and not para.runs:
            doc.add_paragraph()
            continue

        # Handle Heading 1 (add page break before, except first one)
        if style_name.startswith('Heading 1') or style_name == 'Heading 1':
            if first_h1_seen:
                doc.add_page_break()
            first_h1_seen = True
            new_para = doc.add_heading(para_text, level=1)

        # Handle Heading 2
        elif style_name.startswith('Heading 2') or style_name == 'Heading 2':
            new_para = doc.add_heading(para_text, level=2)

        # Handle Heading 3
        elif style_name.startswith('Heading 3') or style_name == 'Heading 3':
            new_para = doc.add_heading(para_text, level=3)

        # Handle Title
        elif style_name == 'Title':
            new_para = doc.add_paragraph()
            new_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            new_para.paragraph_format.space_before = Pt(100)
            new_para.paragraph_format.space_after = Pt(24)
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                set_font_name(new_run, heading_font)
                new_run.font.size = Pt(h1_style['size'] + 14)
                new_run.font.bold = True
                new_run.font.color.rgb = primary_color

        # Handle Subtitle
        elif style_name == 'Subtitle':
            new_para = doc.add_paragraph()
            new_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            new_para.paragraph_format.space_before = Pt(0)
            new_para.paragraph_format.space_after = Pt(36)
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                set_font_name(new_run, body_font)
                new_run.font.size = Pt(body_style['size'] + 2)
                new_run.font.color.rgb = secondary_color

        # Handle List Bullet
        elif 'List Bullet' in style_name or style_name == 'List Bullet':
            new_para = doc.add_paragraph(style='List Bullet')
            new_para.paragraph_format.space_before = Pt(2)
            new_para.paragraph_format.space_after = Pt(2)
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                set_font_name(new_run, body_font)
                new_run.font.size = Pt(body_style['size'])
                new_run.font.color.rgb = text_color
                if run.bold:
                    new_run.bold = True
                if run.italic:
                    new_run.italic = True

        # Handle List Number
        elif 'List Number' in style_name or style_name == 'List Number':
            new_para = doc.add_paragraph(style='List Number')
            new_para.paragraph_format.space_before = Pt(2)
            new_para.paragraph_format.space_after = Pt(2)
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                set_font_name(new_run, body_font)
                new_run.font.size = Pt(body_style['size'])
                new_run.font.color.rgb = text_color
                if run.bold:
                    new_run.bold = True
                if run.italic:
                    new_run.italic = True

        # Handle Quote/Block Quote
        elif 'Quote' in style_name:
            new_para = doc.add_paragraph()
            new_para.paragraph_format.left_indent = Inches(0.5)
            new_para.paragraph_format.right_indent = Inches(0.5)
            new_para.paragraph_format.space_before = Pt(12)
            new_para.paragraph_format.space_after = Pt(12)
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                set_font_name(new_run, body_font)
                new_run.font.size = Pt(body_style['size'])
                new_run.font.italic = True
                new_run.font.color.rgb = secondary_color

        # Handle Caption
        elif 'Caption' in style_name:
            new_para = doc.add_paragraph()
            new_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            new_para.paragraph_format.space_before = Pt(6)
            new_para.paragraph_format.space_after = Pt(12)
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                set_font_name(new_run, body_font)
                new_run.font.size = Pt(caption_style['size'])
                new_run.font.color.rgb = secondary_color
                if caption_style.get('italic', False):
                    new_run.font.italic = True

        # Handle regular paragraphs (Normal and others)
        else:
            new_para = doc.add_paragraph()

            # Preserve alignment
            if para.alignment:
                new_para.alignment = para.alignment

            new_para.paragraph_format.space_before = Pt(0)
            new_para.paragraph_format.space_after = Pt(10)

            # Copy runs with formatting
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                set_font_name(new_run, body_font)
                new_run.font.size = Pt(body_style['size'])
                new_run.font.color.rgb = text_color

                # Preserve character formatting
                if run.bold:
                    new_run.bold = True
                if run.italic:
                    new_run.italic = True
                if run.underline:
                    new_run.underline = True

    # ========================================================================
    # COPY TABLES WITH STYLING
    # ========================================================================

    for table in source_doc.tables:
        rows = len(table.rows)
        cols = len(table.columns)

        new_table = doc.add_table(rows=rows, cols=cols)
        new_table.style = 'Table Grid'
        new_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                new_cell = new_table.rows[i].cells[j]

                for para_idx, para in enumerate(cell.paragraphs):
                    if para_idx == 0 and new_cell.paragraphs:
                        new_para = new_cell.paragraphs[0]
                        new_para.clear()
                    else:
                        new_para = new_cell.add_paragraph()

                    for run in para.runs:
                        new_run = new_para.add_run(run.text)
                        set_font_name(new_run, body_font)
                        new_run.font.size = Pt(body_style['size'])

                        # Header row gets accent styling
                        if i == 0:
                            new_run.font.bold = True
                            new_run.font.color.rgb = accent_color
                        else:
                            new_run.font.color.rgb = text_color

                        if run.bold:
                            new_run.bold = True
                        if run.italic:
                            new_run.italic = True

    # ========================================================================
    # SAVE DOCUMENT
    # ========================================================================

    doc.save(output_path)
    return output_path


# ============================================================================
# BATCH PROCESSING
# ============================================================================

def apply_multiple_brands(input_path: str, brands: List[str], output_prefix: str) -> List[str]:
    """
    Apply multiple brand styles to create variants of the same document.

    Args:
        input_path: Path to source DOCX file
        brands: List of brand names to apply
        output_prefix: Prefix for output files (brand name will be appended)

    Returns:
        List of paths to created output files
    """
    created_files = []

    for brand in brands:
        brand_lower = brand.lower().strip()

        # Determine output filename
        if output_prefix.endswith('.docx'):
            # User provided full filename, insert brand before extension
            base = output_prefix[:-5]
            output_path = f"{base}_{brand_lower}.docx"
        else:
            output_path = f"{output_prefix}_{brand_lower}.docx"

        try:
            result = apply_brand_to_docx(input_path, brand_lower, output_path)
            created_files.append(result)
            print(f"  [+] Created: {result}")
        except Exception as e:
            print(f"  [!] Error applying {brand}: {e}")

    return created_files


# ============================================================================
# DISPLAY UTILITIES
# ============================================================================

def display_brand_menu():
    """Display the brand selection menu with detailed information."""
    script_dir = get_script_dir()
    config_path = script_dir / 'templates' / 'brand-mapping.json'

    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    print()
    print("=" * 78)
    print("                      ELITE DOCUMENT POLISHER")
    print("                  World-Class Brand Styling System")
    print("=" * 78)
    print()

    categories = {
        'editorial': 'EDITORIAL EXCELLENCE',
        'consulting': 'CONSULTING AUTHORITY',
        'tech': 'TECH INNOVATION',
        'productivity': 'PRODUCTIVITY & DESIGN',
        'design': 'DESIGN & CREATIVITY'
    }

    # Group brands by category
    brands_by_category = {}
    for brand_id, brand in config['brands'].items():
        cat = brand['category']
        if cat not in brands_by_category:
            brands_by_category[cat] = []
        brands_by_category[cat].append((brand_id, brand))

    # Display by category
    for cat_key, cat_name in categories.items():
        if cat_key in brands_by_category:
            print(f"  {cat_name}")
            print(f"  {'─' * len(cat_name)}")

            for brand_id, brand in brands_by_category[cat_key]:
                print(f"    {brand_id:12} │ {brand['name']}")
                print(f"    {' ':12} │ {brand['description'][:60]}...")
                print(f"    {' ':12} │ Primary: {brand['colors']['primary']}, "
                      f"Accent: {brand['colors']['accent']}")
                print()

    print("-" * 78)
    print()
    print("USAGE:")
    print("  python apply_brand.py <input.docx> <brand> <output.docx>")
    print("  python apply_brand.py <input.docx> <brand1,brand2> <output_prefix>")
    print("  python apply_brand.py <input.docx> all <output_prefix>")
    print()
    print("EXAMPLES:")
    print("  python apply_brand.py report.docx mckinsey polished_report.docx")
    print("  python apply_brand.py proposal.docx mckinsey,deloitte,stripe proposal")
    print("  python apply_brand.py document.docx all variants/document")
    print()
    print("-" * 78)
    print()


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main():
    """Main entry point for the Elite Document Polisher."""

    # Handle help/list commands
    if len(sys.argv) == 1 or sys.argv[1] in ['-h', '--help', 'help', 'list', '--list', '-l']:
        display_brand_menu()
        return 0

    # Validate arguments
    if len(sys.argv) < 4:
        print("ERROR: Insufficient arguments")
        print()
        print("Usage:")
        print("  python apply_brand.py <input.docx> <brand_name> <output.docx>")
        print("  python apply_brand.py <input.docx> <brand1,brand2> <output_prefix>")
        print("  python apply_brand.py <input.docx> all <output_prefix>")
        print()
        print("Run 'python apply_brand.py --list' to see available brands")
        return 1

    input_path = sys.argv[1]
    brand_arg = sys.argv[2].lower()
    output_arg = sys.argv[3]

    # Validate input file exists
    if not os.path.exists(input_path):
        print(f"ERROR: Input file not found: {input_path}")
        return 1

    # Validate input is a DOCX file
    if not input_path.lower().endswith('.docx'):
        print(f"ERROR: Input file must be a .docx file: {input_path}")
        return 1

    print()
    print("=" * 60)
    print("ELITE DOCUMENT POLISHER")
    print("=" * 60)
    print(f"  Input: {input_path}")
    print()

    try:
        # Handle "all" brands
        if brand_arg == 'all':
            brands = get_all_brands()
            print(f"  Mode: All Brands ({len(brands)} variants)")
            print(f"  Output prefix: {output_arg}")
            print()
            print("  Generating variants...")
            print("-" * 60)

            created = apply_multiple_brands(input_path, brands, output_arg)

            print("-" * 60)
            print(f"  Successfully created {len(created)} document(s)")

        # Handle multiple brands (comma-separated)
        elif ',' in brand_arg:
            brands = [b.strip() for b in brand_arg.split(',')]
            print(f"  Mode: Multiple Brands ({len(brands)} variants)")
            print(f"  Brands: {', '.join(brands)}")
            print(f"  Output prefix: {output_arg}")
            print()
            print("  Generating variants...")
            print("-" * 60)

            created = apply_multiple_brands(input_path, brands, output_arg)

            print("-" * 60)
            print(f"  Successfully created {len(created)} document(s)")

        # Handle single brand
        else:
            brand_config = load_brand_config(brand_arg)
            print(f"  Mode: Single Brand")
            print(f"  Brand: {brand_config['name']}")
            print(f"  Output: {output_arg}")
            print()
            print("  Applying brand styling...")
            print("-" * 60)

            result = apply_brand_to_docx(input_path, brand_arg, output_arg)

            print(f"  [+] Successfully created: {result}")
            print("-" * 60)
            print(f"  Brand '{brand_config['name']}' applied successfully!")

        print()
        print("=" * 60)
        print()
        return 0

    except ValueError as e:
        print(f"ERROR: {e}")
        return 1
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        return 1
    except Exception as e:
        print(f"ERROR: Unexpected error - {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
