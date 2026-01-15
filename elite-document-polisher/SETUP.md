# Elite Document Polisher - Setup Guide

The world's most sophisticated document styling system. Transform any DOCX document into a professionally-styled masterpiece using 10 premium brand aesthetics.

## Quick Start

1. Install the skill
2. Install Python dependencies
3. Use natural language: *"Polish my report.docx with McKinsey style"*

---

## Installation

### For Claude.ai (Browser)

1. Download the `elite-document-polisher` folder as a zip file
2. Go to **Settings** > **Capabilities**
3. Click **"Upload Skill"** and select the zip file
4. The skill is now available in your conversations

### For Claude Code (CLI)

**Global Installation (Recommended)**
```bash
# Copy to personal skills directory
cp -r elite-document-polisher ~/.claude/skills/
```

**Project-Specific Installation**
```bash
# Copy to project's skills directory
mkdir -p .claude/skills
cp -r elite-document-polisher .claude/skills/
```

### For Claude API

Use the `/v1/skills` endpoint to upload the skill package as a zip file.

---

## Python Dependencies

The skill requires `python-docx` for document manipulation.

### Quick Install
```bash
pip install python-docx
```

### Recommended: Virtual Environment
```bash
# Create virtual environment
python3 -m venv venv

# Activate (Linux/Mac)
source venv/bin/activate

# Activate (Windows)
venv\Scripts\activate

# Install dependency
pip install python-docx
```

### Verify Installation
```bash
python -c "from docx import Document; print('python-docx installed successfully')"
```

---

## What's Included

```
elite-document-polisher/
├── SKILL.md                    # Main skill definition
├── demo-prompt.txt             # Quick usage reference
├── SETUP.md                    # This installation guide
├── scripts/
│   └── apply_brand.py          # Brand styling engine
├── templates/
│   └── brand-mapping.json      # Brand configurations
├── brands/
│   └── *.md                    # Individual brand references
└── samples/
    └── sample_report.docx      # Test document
```

---

## Available Brands

| Brand | Category | Best For |
|-------|----------|----------|
| The Economist | Editorial | Reports, analysis, thought leadership |
| McKinsey | Consulting | Strategy decks, executive summaries |
| Deloitte | Consulting | Audits, assessments, formal reports |
| KPMG | Consulting | Financial reports, compliance docs |
| Stripe | Tech | API docs, developer guides |
| Apple | Tech | Product docs, user guides |
| IBM | Tech | Technical docs, enterprise reports |
| Linear | Tech | Product specs, changelogs |
| Notion | Productivity | Wikis, project plans, documentation |
| Figma | Design | Creative briefs, design docs |

---

## Usage Examples

### Single Brand
```
"Polish my quarterly_report.docx with McKinsey style"
```
Output: `quarterly_report_mckinsey.docx`

### Multiple Brands
```
"Generate my proposal.docx in McKinsey, Deloitte, and Stripe styles"
```
Output:
- `proposal_mckinsey.docx`
- `proposal_deloitte.docx`
- `proposal_stripe.docx`

### All Brands
```
"Create all 10 brand versions of my document.docx"
```
Output: 10 styled variants

### Brand Recommendation
```
"Which style would work best for my financial audit report?"
```
Claude will recommend KPMG or Deloitte based on document type.

---

## Command Line Usage

If running the script directly:

```bash
# Single brand
python scripts/apply_brand.py input.docx mckinsey output.docx

# Multiple brands (comma-separated)
python scripts/apply_brand.py input.docx mckinsey,deloitte,stripe output_prefix

# All brands
python scripts/apply_brand.py input.docx all output_prefix

# List available brands
python scripts/apply_brand.py --list
```

---

## Permissions Configuration

Add to your `settings.local.json` to auto-approve script execution:

```json
{
  "permissions": {
    "allow": [
      "Bash(python:*)",
      "Bash(python3:*)",
      "Skill(elite-document-polisher)"
    ]
  }
}
```

---

## Troubleshooting

### "Module not found: docx"
```bash
pip install python-docx
```

### Script Path Issues
Use absolute paths or run from project root:
```bash
python /path/to/elite-document-polisher/scripts/apply_brand.py input.docx brand output.docx
```

### Document Styling Looks Wrong
- Ensure source document uses standard Word styles (Heading 1, 2, 3, Normal)
- Check that lists use built-in List Bullet/List Number styles

### Brand Not Found
Available brands (case-insensitive):
`economist`, `mckinsey`, `deloitte`, `kpmg`, `stripe`, `apple`, `ibm`, `linear`, `notion`, `figma`

---

## Tips for Best Results

1. **Use Proper Headings**: Documents with Heading 1/2/3 styles transform best
2. **Match Brand to Audience**: Consulting for executives, tech for developers
3. **Batch Compare**: Generate 2-3 variants to compare before choosing
4. **Test PDF Export**: Final appearance may vary between Word and PDF

---

## Support

For issues or feature requests, consult the SKILL.md documentation or the brand reference files in the `brands/` directory.
