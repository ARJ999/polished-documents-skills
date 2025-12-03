# Setup Guide

This guide walks you through setting up the Claude Code Skills Collection from scratch.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Skill Installation Methods](#skill-installation-methods)
- [Python Environment Setup](#python-environment-setup)
- [Permissions Configuration](#permissions-configuration)
- [Verification](#verification)

---

## Prerequisites

### Required

- **Claude Code** - Installed and configured ([Installation Guide](https://claude.com/claude-code))
- **Python 3.8+** - For document generation scripts
- **pip** - Python package manager

### Optional (for MCP features)

- **Node.js 18+** - For running MCP servers
- **npm** - Node package manager
- **FireCrawl API Key** - For brand extraction ([Get one here](https://firecrawl.dev))

### Check Prerequisites

```bash
# Check Claude Code
claude --version

# Check Python
python3 --version

# Check Node.js (optional)
node --version
npm --version
```

---

## Installation

### Step 1: Clone the Repository

```bash
git clone https://github.com/YOUR_USERNAME/claude-code-skills.git
cd claude-code-skills
```

### Step 2: Set Up Python Environment

We recommend using a virtual environment to avoid dependency conflicts:

```bash
# Create virtual environment
python3 -m venv venv

# Activate it
source venv/bin/activate  # macOS/Linux
# or
.\venv\Scripts\activate   # Windows

# Install required packages
pip install python-docx

# Verify installation
python -c "import docx; print('python-docx installed successfully')"
```

### Step 3: Install Skills

See [Skill Installation Methods](#skill-installation-methods) below.

### Step 4: Configure Permissions

```bash
# Copy example settings
cp .claude/settings.example.json .claude/settings.local.json
```

---

## Skill Installation Methods

### Method 1: Project-Specific Installation

Install skills for a single project:

```bash
# Navigate to your project
cd /path/to/your/project

# Create .claude directory if it doesn't exist
mkdir -p .claude/skills

# Copy skills from this repository
cp -r /path/to/claude-code-skills/.claude/skills/* .claude/skills/
```

### Method 2: Global Installation (Recommended)

Install skills globally so they're available in all projects:

```bash
# Create global skills directory
mkdir -p ~/.claude/skills

# Copy all skills
cp -r .claude/skills/* ~/.claude/skills/

# Verify
ls ~/.claude/skills/
# Should show: document-polisher docx pdf xlsx pptx ...
```

### Method 3: Selective Installation

Install only the skills you need:

```bash
# Install only document-polisher
cp -r .claude/skills/document-polisher ~/.claude/skills/

# Install document-polisher + core document skills
for skill in document-polisher docx pdf xlsx pptx; do
  cp -r .claude/skills/$skill ~/.claude/skills/
done
```

---

## Python Environment Setup

### Why Virtual Environment?

Using a virtual environment:
- Isolates dependencies from system Python
- Prevents conflicts with other projects
- Makes dependencies reproducible

### Setup Commands

```bash
# Create (one-time)
python3 -m venv venv

# Activate (every session)
source venv/bin/activate  # macOS/Linux
.\venv\Scripts\activate   # Windows

# Install dependencies
pip install python-docx

# Deactivate when done
deactivate
```

### Using with Claude Code

When using skills that require Python packages, ensure your virtual environment is activated:

```bash
# Activate first
source venv/bin/activate

# Then run Claude Code
claude
```

Alternatively, configure Claude Code to use the virtual environment Python:

```json
{
  "permissions": {
    "allow": [
      "Bash(./venv/bin/python:*)"
    ]
  }
}
```

---

## Permissions Configuration

### Understanding Permissions

Claude Code uses a permission system to control what actions it can perform. The `settings.local.json` file defines these permissions.

### Example Configuration

```json
{
  "permissions": {
    "allow": [
      "Bash(python:*)",
      "Bash(python3:*)",
      "Bash(pip install:*)",
      "Bash(pip3 install:*)",
      "Bash(source venv/bin/activate)",
      "Bash(ls:*)",
      "Bash(mkdir:*)",
      "Bash(cat:*)",
      "Skill(document-polisher)",
      "Skill(docx)",
      "Skill(pdf)",
      "Skill(xlsx)",
      "Skill(pptx)",
      "Bash(python \".claude/skills/document-polisher/scripts/apply_brand.py\":*)"
    ],
    "deny": [],
    "ask": []
  }
}
```

### Permission Types

| Type | Description | Example |
|------|-------------|---------|
| `allow` | Always permitted | `"Bash(ls:*)"` |
| `deny` | Always blocked | `"Bash(rm -rf:*)"` |
| `ask` | Prompt before running | `"Bash(pip install:*)"` |

### Wildcard Patterns

- `*` - Match anything
- `Bash(python:*)` - Allow any python command
- `Skill(docx)` - Allow specific skill

---

## Verification

### Test Skill Installation

```bash
# Start Claude Code
claude

# In Claude Code, run:
/skill document-polisher
```

You should see the brand selection menu.

### Test Document Generation

```bash
# Activate virtual environment
source venv/bin/activate

# Run the apply_brand script
python .claude/skills/document-polisher/scripts/apply_brand.py --list
```

Expected output:
```
======================================================================
DOCUMENT POLISHER - Available Brand Styles
======================================================================

  economist    | The Economist
               | Iconic editorial style with signature red and serif typography...
...
```

### Test Full Workflow

1. Create a test document:
```bash
# In Claude Code
Create a simple test document with a title, heading, and paragraph
```

2. Apply brand styling:
```bash
Polish this document with the McKinsey style
```

3. Verify the output document has correct styling.

---

## Next Steps

- [MCP Server Setup](MCP-SERVERS.md) - Configure FireCrawl for brand extraction
- [Troubleshooting](TROUBLESHOOTING.md) - Common issues and solutions
- [README](../README.md) - Full feature documentation

---

## Quick Reference

### Activate Environment
```bash
source venv/bin/activate
```

### Apply Brand Styling
```bash
python .claude/skills/document-polisher/scripts/apply_brand.py input.docx brand output.docx
```

### List Brands
```bash
python .claude/skills/document-polisher/scripts/apply_brand.py --list
```
