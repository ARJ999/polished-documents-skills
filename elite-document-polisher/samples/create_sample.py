#!/usr/bin/env python3
"""
Sample Document Generator for Elite Document Polisher

This script creates a sample DOCX document with various elements
to demonstrate the brand styling capabilities.

Usage:
    python create_sample.py

Output:
    sample_report.docx
"""

try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print("Error: python-docx is required. Install with: pip install python-docx")
    exit(1)


def create_sample_document():
    """Create a comprehensive sample document for testing brand styles."""

    doc = Document()

    # Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Q4 2024 Strategic Performance Report")
    run.bold = True
    title.style = 'Title'

    # Subtitle
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("Comprehensive Analysis and Forward-Looking Projections")
    subtitle.style = 'Subtitle'

    # Executive Summary (H1)
    doc.add_heading("Executive Summary", level=1)

    doc.add_paragraph(
        "This report presents a comprehensive analysis of our Q4 2024 performance "
        "across all business units. Key highlights include a 23% year-over-year "
        "revenue growth, successful expansion into three new markets, and the "
        "launch of our flagship product line which exceeded initial projections by 40%."
    )

    doc.add_paragraph(
        "Our strategic initiatives delivered exceptional results, with customer "
        "satisfaction scores reaching an all-time high of 94%. The following sections "
        "provide detailed analysis and actionable recommendations for sustaining "
        "this momentum into 2025."
    )

    # Key Metrics (H2)
    doc.add_heading("Key Performance Metrics", level=2)

    # Bullet list
    doc.add_paragraph("Revenue: $47.3M (+23% YoY)", style='List Bullet')
    doc.add_paragraph("Gross Margin: 68.2% (+3.1pp)", style='List Bullet')
    doc.add_paragraph("Customer Acquisition: 12,400 new customers", style='List Bullet')
    doc.add_paragraph("Net Promoter Score: 72 (+8 points)", style='List Bullet')
    doc.add_paragraph("Employee Satisfaction: 4.6/5.0", style='List Bullet')

    # Market Analysis (H1)
    doc.add_heading("Market Analysis", level=1)

    doc.add_paragraph(
        "The global market environment presented both challenges and opportunities "
        "during Q4. Despite macroeconomic headwinds, our positioning in key growth "
        "segments allowed us to capture significant market share while maintaining "
        "pricing discipline."
    )

    # Competitive Landscape (H2)
    doc.add_heading("Competitive Landscape", level=2)

    doc.add_paragraph(
        "Our competitive position strengthened considerably this quarter. Market "
        "research indicates we now hold the #2 position in our primary segment, "
        "closing the gap with the market leader from 12% to just 7%."
    )

    # Regional Performance (H3)
    doc.add_heading("Regional Performance Breakdown", level=3)

    # Numbered list
    doc.add_paragraph("North America: 45% of revenue, +18% growth", style='List Number')
    doc.add_paragraph("Europe: 32% of revenue, +27% growth", style='List Number')
    doc.add_paragraph("Asia-Pacific: 18% of revenue, +34% growth", style='List Number')
    doc.add_paragraph("Rest of World: 5% of revenue, +12% growth", style='List Number')

    # Table
    doc.add_heading("Quarterly Revenue by Segment", level=2)

    table = doc.add_table(rows=5, cols=4)
    table.style = 'Table Grid'

    # Header row
    headers = ["Segment", "Q4 2024", "Q4 2023", "Change"]
    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header

    # Data rows
    data = [
        ["Enterprise", "$22.1M", "$17.8M", "+24.2%"],
        ["Mid-Market", "$14.7M", "$12.1M", "+21.5%"],
        ["SMB", "$8.2M", "$6.9M", "+18.8%"],
        ["Consumer", "$2.3M", "$1.6M", "+43.8%"]
    ]

    for row_idx, row_data in enumerate(data, start=1):
        for col_idx, cell_text in enumerate(row_data):
            table.rows[row_idx].cells[col_idx].text = cell_text

    # Strategic Initiatives (H1)
    doc.add_heading("Strategic Initiatives", level=1)

    doc.add_paragraph(
        "Our strategic roadmap for 2024 focused on three pillars: product innovation, "
        "market expansion, and operational excellence. Each pillar delivered "
        "measurable results that position us well for continued growth."
    )

    # Product Innovation (H2)
    doc.add_heading("Product Innovation", level=2)

    para = doc.add_paragraph()
    para.add_run("Our R&D investments yielded significant breakthroughs. ").bold = False
    para.add_run("The new AI-powered analytics module").bold = True
    para.add_run(" launched in October and has already been adopted by 67% of enterprise customers.")

    # Operational Excellence (H2)
    doc.add_heading("Operational Excellence", level=2)

    doc.add_paragraph(
        "Process improvements and automation initiatives reduced operational costs "
        "by 15% while improving service delivery times by 22%. Customer support "
        "resolution rates improved from 89% to 96%."
    )

    # Financial Overview (H1)
    doc.add_heading("Financial Overview", level=1)

    doc.add_heading("Revenue Composition", level=2)

    doc.add_paragraph("Subscription Revenue: $38.4M (81.2%)", style='List Bullet')
    doc.add_paragraph("Professional Services: $6.2M (13.1%)", style='List Bullet')
    doc.add_paragraph("Partner Ecosystem: $2.7M (5.7%)", style='List Bullet')

    doc.add_heading("Investment Areas", level=3)

    doc.add_paragraph(
        "Strategic investments totaled $12.3M during Q4, focused primarily on "
        "product development (62%), go-to-market expansion (24%), and infrastructure "
        "modernization (14%)."
    )

    # Recommendations (H1)
    doc.add_heading("Recommendations and Next Steps", level=1)

    para = doc.add_paragraph()
    para.add_run("Based on our Q4 performance and market analysis, we recommend the following ").bold = False
    para.add_run("priority initiatives").bold = True
    para.add_run(" for Q1 2025:").bold = False

    doc.add_paragraph("Accelerate enterprise sales team expansion in EMEA", style='List Number')
    doc.add_paragraph("Launch Phase 2 of the AI analytics platform", style='List Number')
    doc.add_paragraph("Complete integration of recently acquired technology assets", style='List Number')
    doc.add_paragraph("Implement customer success program enhancements", style='List Number')
    doc.add_paragraph("Finalize strategic partnership with key ecosystem players", style='List Number')

    # Conclusion (H1)
    doc.add_heading("Conclusion", level=1)

    doc.add_paragraph(
        "Q4 2024 demonstrated the strength of our strategic direction and the "
        "effectiveness of our execution. With strong fundamentals, a clear roadmap, "
        "and a dedicated team, we are well-positioned to achieve our ambitious "
        "goals for 2025 and beyond."
    )

    para = doc.add_paragraph()
    para.add_run("We appreciate your continued support and look forward to delivering ").bold = False
    para.add_run("exceptional results").italic = True
    para.add_run(" in the coming quarters.").bold = False

    # Save document
    output_path = "sample_report.docx"
    doc.save(output_path)
    print(f"Sample document created: {output_path}")
    print("\nDocument includes:")
    print("  - Title and Subtitle")
    print("  - Heading 1, 2, and 3 levels")
    print("  - Regular paragraphs")
    print("  - Bullet lists")
    print("  - Numbered lists")
    print("  - Tables with headers")
    print("  - Bold and italic formatting")
    print("\nReady for brand styling!")


if __name__ == "__main__":
    create_sample_document()
