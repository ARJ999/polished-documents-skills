---
name: elite-document-polisher
description: "The world's most sophisticated DOCX document styling system v3.0 - God-Level Flawless Edition. Transform any document into a professionally-styled masterpiece using 10 premium brand aesthetics from the world's finest organizations (The Economist, McKinsey, Deloitte, KPMG, Stripe, Apple, IBM, Linear, Notion, Figma). Features: single-brand styling, multi-brand batch generation, QUALITY VALIDATION ensuring only flawless documents, professional table formatting with consistent widths, golden-ratio typography, and orphan/widow control. Use when users want to: (1) Apply world-class brand styling to documents, (2) Generate multiple brand variants of the same document, (3) Make documents look professionally polished, (4) Match specific brand visual identities, (5) Create executive-ready reports, proposals, or presentations."
---

# Elite Document Polisher v3.0

**God-Level Flawless Document Styling**

Transform any document into a professionally-styled masterpiece using world-class brand aesthetics from the finest organizations on the planet. Now with quality validation ensuring only flawless documents are delivered.

## What This Skill Does

The Elite Document Polisher is the definitive tool for applying premium brand styling to DOCX documents. It analyzes your document's content and purpose, then applies sophisticated typography, color schemes, and formatting from 10 meticulously researched brand identities representing the pinnacle of professional document design.

**Key Capabilities:**
- Apply single brand styling with one command
- Generate multiple brand variants simultaneously (batch mode)
- Receive intelligent brand recommendations based on document content
- Preserve all original formatting, structure, and content integrity
- Professional table styling, heading hierarchies, and whitespace management

**v3.0 God-Level Features:**
- **Quality Validation**: Pre-output checks ensure only flawless documents are presented
- **Professional Table Formatting**: All tables span full width with intelligent column sizing
- **Minimal Vertical Borders**: Clean, modern table aesthetic without visual clutter
- **Alternating Row Colors**: Subtle shading for improved readability
- **Golden Ratio Typography**: Spacing calculated using mathematical harmony
- **Orphan/Widow Control**: No lonely lines at page breaks
- **Consistent Cell Padding**: Professional spacing in all table cells

## When to Use

Invoke this skill when the user:
- Wants to make a document look more professional or polished
- Mentions styling, branding, or formatting a Word document
- Asks for a document to look like McKinsey, The Economist, Stripe, etc.
- Needs executive-ready reports, proposals, or presentations
- Wants to generate multiple styled versions of the same document
- Asks "which style would look best" for their document
- Mentions keywords: polish, style, brand, professional, format, executive, premium

## Brand Selection Menu

**ALWAYS display this menu when a user wants to polish a document:**

```
╔═══════════════════════════════════════════════════════════════════════════════════╗
║                           ELITE DOCUMENT POLISHER                                 ║
║                      World-Class Brand Styling System                             ║
╠═══════════════════════════════════════════════════════════════════════════════════╣
║                                                                                   ║
║  EDITORIAL EXCELLENCE                                                             ║
║  ────────────────────                                                             ║
║   1. The Economist    │ Iconic serif typography + signature red accents           ║
║                       │ Best for: Analysis, thought leadership, editorial content ║
║                       │ Typography: Georgia │ Colors: Navy #1F2E7A + Red #E3120B  ║
║                                                                                   ║
║  CONSULTING AUTHORITY                                                             ║
║  ────────────────────                                                             ║
║   2. McKinsey         │ Sharp precision + bold blue commanding presence           ║
║                       │ Best for: Strategy decks, executive summaries, C-suite    ║
║                       │ Typography: Georgia/Calibri │ Colors: Blue #2251FF        ║
║                                                                                   ║
║   3. Deloitte         │ Modern professional + teal-blue sophistication            ║
║                       │ Best for: Audits, assessments, formal business reports    ║
║                       │ Typography: Calibri │ Colors: Teal #007CB0                ║
║                                                                                   ║
║   4. KPMG             │ Corporate gravitas + two-tone blue authority              ║
║                       │ Best for: Financial reports, compliance, governance docs  ║
║                       │ Typography: Calibri │ Colors: Blue #005EB8 + #00338D      ║
║                                                                                   ║
║  TECH INNOVATION                                                                  ║
║  ───────────────                                                                  ║
║   5. Stripe           │ Developer-focused + dark blue/purple gradients            ║
║                       │ Best for: API docs, technical guides, product specs       ║
║                       │ Typography: Arial │ Colors: Navy #0A2540 + Purple #635BFF ║
║                                                                                   ║
║   6. Apple            │ Minimalist premium + generous breathing room              ║
║                       │ Best for: Product documentation, user guides, manuals     ║
║                       │ Typography: Arial │ Colors: Blue #0071E3 + Gray #1D1D1F   ║
║                                                                                   ║
║   7. IBM              │ Enterprise authority + Carbon design system               ║
║                       │ Best for: Technical documentation, enterprise reports     ║
║                       │ Typography: Arial │ Colors: Blue #0F62FE + Black #161616  ║
║                                                                                   ║
║   8. Linear           │ Ultra-modern precision + purple sophistication            ║
║                       │ Best for: Product specs, changelogs, engineering docs     ║
║                       │ Typography: Arial │ Colors: Purple #5E6AD2                ║
║                                                                                   ║
║  PRODUCTIVITY & DESIGN                                                            ║
║  ────────────────────                                                             ║
║   9. Notion           │ Clean productivity + subtle blue highlights               ║
║                       │ Best for: Wikis, documentation, project plans, SOPs       ║
║                       │ Typography: Segoe UI │ Colors: Blue #2383E2               ║
║                                                                                   ║
║  10. Figma            │ Design-forward + vibrant multi-color creativity           ║
║                       │ Best for: Creative briefs, design docs, brand guidelines  ║
║                       │ Typography: Arial │ Colors: Purple #A259FF + Multi        ║
║                                                                                   ║
╠═══════════════════════════════════════════════════════════════════════════════════╣
║  OPTIONS:                                                                         ║
║  • Enter a number (1-10) for single brand styling                                 ║
║  • Enter multiple numbers (e.g., "1,3,5") for batch generation                    ║
║  • Enter "all" to generate all 10 brand variants                                  ║
║  • Enter "recommend" for AI-powered brand suggestion based on your content        ║
╚═══════════════════════════════════════════════════════════════════════════════════╝
```

## How It Works

### Step 1: Display Menu and Gather Input

When a user requests document polishing:
1. Display the brand selection menu above
2. Ask: **"Which brand style would you like? You can select one, multiple (comma-separated), or type 'recommend' for a suggestion based on your document."**
3. If user says "recommend", analyze the document content and suggest the best brand match

### Step 2: Apply Brand Styling

**CRITICAL: Always use the Python script with python-docx. Never use direct XML/OOXML manipulation.**

#### Single Brand Application

```bash
# Ensure python-docx is installed
pip install python-docx

# Apply single brand
python scripts/apply_brand.py <input.docx> <brand_name> <output.docx>

# Examples:
python scripts/apply_brand.py report.docx mckinsey report_mckinsey.docx
python scripts/apply_brand.py proposal.docx economist proposal_economist.docx
python scripts/apply_brand.py guide.docx stripe guide_stripe.docx
```

#### Multi-Brand Batch Generation

```bash
# Generate multiple brand variants at once
python scripts/apply_brand.py <input.docx> <brand1,brand2,brand3> <output_prefix>

# Examples:
python scripts/apply_brand.py report.docx mckinsey,deloitte,kpmg report
# Creates: report_mckinsey.docx, report_deloitte.docx, report_kpmg.docx

python scripts/apply_brand.py proposal.docx all proposal
# Creates all 10 brand variants with prefix "proposal_"
```

### Step 3: Verify Output

After applying styling:
1. Confirm the output file(s) were created successfully
2. Report which brand(s) were applied
3. Provide the output file path(s) to the user

## Brand Quick Reference

| ID | Brand | Category | Best For | Primary Color |
|----|-------|----------|----------|---------------|
| `economist` | The Economist | Editorial | Reports, analysis, thought leadership | #E3120B (Red) |
| `mckinsey` | McKinsey & Company | Consulting | Strategy, executive summaries | #2251FF (Blue) |
| `deloitte` | Deloitte | Consulting | Audits, assessments, formal reports | #007CB0 (Teal) |
| `kpmg` | KPMG | Consulting | Financial reports, compliance | #005EB8 (Blue) |
| `stripe` | Stripe | Tech | API docs, developer guides | #0A2540 (Navy) |
| `apple` | Apple | Tech | Product docs, user guides | #0071E3 (Blue) |
| `ibm` | IBM | Tech | Technical docs, enterprise | #0F62FE (Blue) |
| `linear` | Linear | Tech | Product specs, changelogs | #5E6AD2 (Purple) |
| `notion` | Notion | Productivity | Wikis, project plans, docs | #2383E2 (Blue) |
| `figma` | Figma | Design | Creative briefs, design docs | #A259FF (Purple) |

## Intelligent Brand Recommendations

When user asks for a recommendation, analyze the document and suggest based on:

| Document Type | Recommended Brand | Why |
|---------------|-------------------|-----|
| Financial/Audit Report | KPMG or Deloitte | Corporate authority, compliance aesthetic |
| Strategy/Board Deck | McKinsey | Executive presence, sharp professionalism |
| Analysis/Editorial | The Economist | Intellectual gravitas, editorial excellence |
| API/Technical Docs | Stripe or Linear | Developer-focused, modern tech aesthetic |
| Product Documentation | Apple | Clean minimalism, user-friendly |
| Enterprise Technical | IBM | Authority, comprehensive structure |
| Internal Wiki/SOPs | Notion | Clean productivity, accessible |
| Creative/Design Brief | Figma | Visual creativity, design-forward |
| General Business | McKinsey or Deloitte | Safe professional choices |
| Startup/Modern Tech | Linear or Stripe | Contemporary, innovative feel |

## What the Script Does

The `apply_brand.py` script performs these operations:

1. **Loads Brand Configuration** from `templates/brand-mapping.json`
2. **Creates Fresh Document** (prevents XML corruption)
3. **Applies Document-Level Styles:**
   - Title style (centered, prominent, branded)
   - Heading 1, 2, 3 (font, size, color, spacing)
   - Normal/body text (font, size, line height)
   - List Bullet and List Number styles
   - Caption style for figures/tables
4. **Sets Page Layout:**
   - 1" top/bottom margins
   - 1.25" left/right margins
   - Professional spacing
5. **Copies Content with Formatting:**
   - Preserves heading hierarchy
   - Maintains bold, italic, underline
   - Keeps bullet and numbered lists
   - Preserves paragraph alignment
   - Copies tables with styling
6. **Applies Professional Table Formatting** (v3.0):
   - Sets all tables to 100% content width
   - Optimizes column widths based on content analysis
   - Applies minimal vertical borders (clean modern look)
   - Formats header row with accent color
   - Adds alternating row shading
   - Sets consistent cell padding
7. **Runs Quality Validation** (v3.0):
   - Checks typography for orphans/widows
   - Validates table structure
   - Ensures spacing consistency
   - Verifies heading hierarchy
   - Reports quality level (PERFECT, ACCEPTABLE, NEEDS_ATTENTION, FAILED)
8. **Saves as New File** (never overwrites original)

## Quality Validation System (v3.0)

The script includes a comprehensive quality validation system that ensures only flawless documents are delivered:

### Quality Levels

| Level | Description |
|-------|-------------|
| **PERFECT** | No issues detected - document is flawless |
| **ACCEPTABLE** | Minor issues detected but auto-corrected |
| **NEEDS_ATTENTION** | Some issues require manual review |
| **FAILED** | Critical issues prevent quality output |

### What Gets Validated

| Category | Checks |
|----------|--------|
| **Typography** | Orphan/widow lines, runt lines (single short words) |
| **Tables** | Consistent column counts, non-empty headers |
| **Spacing** | No excessive empty paragraphs |
| **Structure** | Proper heading hierarchy (no skipped levels) |
| **Layout** | Professional margins maintained |

### Quality Output Example

```
======================================================================
ELITE DOCUMENT POLISHER v3.0 - God-Level Flawless Edition
======================================================================
  Input: report.docx

  Mode: Single Brand
  Brand: McKinsey & Company
  Output: report_mckinsey.docx

  Applying brand styling with quality validation...
----------------------------------------------------------------------
  [✓] Created: report_mckinsey.docx
  Quality Level: PERFECT
----------------------------------------------------------------------
  Brand 'McKinsey & Company' applied successfully!
======================================================================
```

## Professional Table Formatting (v3.0)

Tables are formatted with world-class standards:

### Table Features

| Feature | Implementation |
|---------|----------------|
| **Width** | 100% of content area (full page width) |
| **Column Sizing** | Intelligent proportional based on content |
| **Min/Max Width** | 15% minimum, 60% maximum per column |
| **Borders** | Horizontal only, minimal vertical (nil) |
| **Header Border** | 1.5pt accent-colored bottom border |
| **Body Borders** | 0.5pt subtle horizontal separators |
| **Cell Padding** | 6pt vertical, 8pt horizontal |
| **Row Height** | 28pt header, 24pt body (minimum) |
| **Alternating Rows** | Subtle #F8F9FA gray shading |

### Before vs After

**Before (Misaligned Tables):**
- Different tables have inconsistent widths
- Column widths don't align across tables
- Cluttered with vertical borders

**After (Professional Tables):**
- All tables span full content width
- Clean horizontal lines only
- Consistent cell padding throughout
- Professional header styling with accent colors

## Style Specifications by Brand

### The Economist
```
Heading 1: Georgia 28pt, #E3120B (Red), Bold
Heading 2: Georgia 18pt, #1F2E7A (Navy), Bold
Heading 3: Georgia 13pt, #0D0D0D (Black), Bold
Body: Georgia 11pt, #0D0D0D
Tables: Minimal borders, serif typography
```

### McKinsey & Company
```
Heading 1: Georgia 52pt, #051C2C (Dark Navy), Bold
Heading 2: Georgia 32pt, #051C2C, Bold
Heading 3: Georgia 20pt, #051C2C, Bold
Body: Calibri 11pt, #051C2C
Tables: Clean lines, sharp corners
```

### Deloitte
```
Heading 1: Calibri 40pt, #007CB0 (Teal), Bold
Heading 2: Calibri 28pt, #007CB0, Bold
Heading 3: Calibri 16pt, #000000, Bold
Body: Calibri 11pt, #000000
Tables: Rounded corners, modern aesthetic
```

### KPMG
```
Heading 1: Calibri 42pt, #00338D (Blue), Bold
Heading 2: Calibri 38pt, #00338D, Bold
Heading 3: Calibri 16pt, #00338D, Bold
Body: Calibri 11pt, #00338D
Tables: Rounded modern style
```

### Stripe
```
Heading 1: Arial 32pt, #0A2540 (Navy), Bold
Heading 2: Arial 24pt, #0A2540, Bold
Heading 3: Arial 16pt, #0A2540, Bold
Body: Arial 11pt, #0A2540
Tables: Subtle shadow, modern tech
```

### Apple
```
Heading 1: Arial 32pt, #1D1D1F (Near Black), Regular
Heading 2: Arial 24pt, #1D1D1F, Regular
Heading 3: Arial 16pt, #1D1D1F, Bold
Body: Arial 12pt, #1D1D1F
Tables: Minimal clean, generous spacing
```

### IBM
```
Heading 1: Arial 36pt, #161616 (Black), Bold
Heading 2: Arial 28pt, #161616, Bold
Heading 3: Arial 18pt, #161616, Bold
Body: Arial 12pt, #161616
Tables: Grid style, enterprise structure
```

### Linear
```
Heading 1: Arial 32pt, #1A1A1A (Near Black), Bold
Heading 2: Arial 24pt, #1A1A1A, Bold
Heading 3: Arial 16pt, #1A1A1A, Bold
Body: Arial 11pt, #1A1A1A
Tables: Modern subtle, clean lines
```

### Notion
```
Heading 1: Segoe UI 48pt, #191918 (Near Black), Bold
Heading 2: Segoe UI 32pt, #191918, Bold
Heading 3: Segoe UI 20pt, #191918, Bold
Body: Segoe UI 12pt, #191918
Tables: Block-based style
```

### Figma
```
Heading 1: Arial 36pt, #A259FF (Purple), Bold
Heading 2: Arial 26pt, #1ABCFE (Blue), Bold
Heading 3: Arial 16pt, #0ACF83 (Green), Bold
Body: Arial 11pt, #1E1E1E
Tables: Colorful accents
```

## Required Libraries

- **python-docx** - Core library for DOCX manipulation (`pip install python-docx`)
- **Python 3.8+** - Required runtime environment

## Example Usage

### Example 1: Single Brand Application

**User prompt**: "Make my quarterly report look like a McKinsey document"

**Claude will**:
1. Display the brand selection menu
2. Confirm McKinsey selection
3. Run the styling command:
```bash
python scripts/apply_brand.py quarterly_report.docx mckinsey quarterly_report_mckinsey.docx
```
4. Confirm successful creation

**Output**: `quarterly_report_mckinsey.docx`

---

### Example 2: Multi-Brand Batch Generation

**User prompt**: "Generate this proposal in McKinsey, Deloitte, and Stripe styles"

**Claude will**:
1. Acknowledge multi-brand request
2. Run batch command:
```bash
python scripts/apply_brand.py proposal.docx mckinsey,deloitte,stripe proposal
```
3. Report all three files created

**Output files**:
- `proposal_mckinsey.docx`
- `proposal_deloitte.docx`
- `proposal_stripe.docx`

---

### Example 3: Generate All Variants

**User prompt**: "Create all 10 brand versions of my document"

**Claude will**:
1. Confirm all-brand generation
2. Run:
```bash
python scripts/apply_brand.py document.docx all document
```
3. List all 10 generated files

**Output files**:
- `document_economist.docx`
- `document_mckinsey.docx`
- `document_deloitte.docx`
- `document_kpmg.docx`
- `document_stripe.docx`
- `document_apple.docx`
- `document_ibm.docx`
- `document_linear.docx`
- `document_notion.docx`
- `document_figma.docx`

---

### Example 4: Brand Recommendation

**User prompt**: "Which style would work best for my financial audit report?"

**Claude will**:
1. Analyze the document type (financial audit)
2. Recommend: **KPMG** or **Deloitte**
3. Explain: "For financial audit reports, I recommend KPMG or Deloitte styling. Both convey corporate authority and compliance credibility. KPMG offers bold two-tone blue for maximum impact, while Deloitte provides modern teal sophistication."
4. Ask which the user prefers
5. Apply the selected brand

---

## Tips for Best Results

1. **Match Brand to Audience**: Use consulting brands (McKinsey, Deloitte, KPMG) for executives; tech brands (Stripe, Linear) for developers
2. **Consider Document Purpose**: The Economist for thought leadership; Apple for user-facing docs; IBM for enterprise technical content
3. **Use Proper Source Formatting**: Documents with proper Heading 1/2/3 styles transform best
4. **Test with PDF Export**: Final appearance may vary slightly between Word and PDF
5. **Batch for Comparison**: Generate 2-3 variants to compare before choosing final version
6. **Don't Over-Brand**: Goal is professional polish, not exact brand copying

## Troubleshooting

### "Module not found: docx"
```bash
pip install python-docx
```

### Document Styling Looks Wrong
- Ensure source document uses standard Word styles (Heading 1, 2, 3, Normal)
- Check that lists use built-in List Bullet/List Number styles
- Verify no custom styles are overriding defaults

### Script Path Issues
Always run from project root or use absolute paths:
```bash
python /path/to/elite-document-polisher/scripts/apply_brand.py input.docx brand output.docx
```

### Brand Not Found Error
Available brand names (case-insensitive):
`economist`, `mckinsey`, `deloitte`, `kpmg`, `stripe`, `apple`, `ibm`, `linear`, `notion`, `figma`

## File Structure

```
elite-document-polisher/
├── SKILL.md                          # This file
├── demo-prompt.txt                   # Quick usage reference
├── SETUP.md                          # Installation guide
├── scripts/
│   └── apply_brand.py                # Main styling engine
├── templates/
│   └── brand-mapping.json            # Brand configurations
├── brands/
│   ├── economist.md                  # The Economist brand details
│   ├── mckinsey.md                   # McKinsey brand details
│   ├── deloitte.md                   # Deloitte brand details
│   ├── kpmg.md                       # KPMG brand details
│   ├── stripe.md                     # Stripe brand details
│   ├── apple.md                      # Apple brand details
│   ├── ibm.md                        # IBM brand details
│   ├── linear.md                     # Linear brand details
│   ├── notion.md                     # Notion brand details
│   └── figma.md                      # Figma brand details
└── samples/
    └── sample_report.docx            # Test document
```

## Adding Custom Brands

To add a new brand to the system:

1. **Extract Brand Identity** (if from website):
   ```
   Use FireCrawl MCP: mcp__firecrawl__firecrawl_extract_branding
   ```

2. **Create Brand Reference** in `brands/<brand_name>.md`

3. **Add to Configuration** in `templates/brand-mapping.json`:
   ```json
   "brandname": {
     "name": "Brand Display Name",
     "description": "Description of the brand aesthetic",
     "category": "consulting|tech|editorial|productivity|design",
     "colors": {
       "primary": "#HEX",
       "accent": "#HEX",
       "background": "#FFFFFF",
       "textPrimary": "#HEX",
       "textSecondary": "#HEX"
     },
     "typography": {
       "headingFont": "Font Name",
       "bodyFont": "Font Name",
       "headingWeight": "bold",
       "bodyWeight": "normal"
     },
     "styles": {
       "h1": {"size": 28, "color": "#HEX", "bold": true},
       "h2": {"size": 22, "color": "#HEX", "bold": true},
       "h3": {"size": 16, "color": "#HEX", "bold": true},
       "body": {"size": 11, "color": "#HEX"},
       "caption": {"size": 9, "color": "#HEX"}
     },
     "elements": {
       "borderRadius": 0,
       "tableStyle": "minimal",
       "accentUse": "headers_only"
     }
   }
   ```

4. **Test** with sample documents before production use
