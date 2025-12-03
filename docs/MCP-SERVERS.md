# MCP Server Configuration Guide

This guide covers setting up MCP (Model Context Protocol) servers to extend Claude Code's capabilities, specifically the **FireCrawl MCP server** for extracting brand identity from websites.

## Table of Contents

- [What is MCP?](#what-is-mcp)
- [FireCrawl MCP Server](#firecrawl-mcp-server)
- [Installation Methods](#installation-methods)
- [Configuration](#configuration)
- [Available Tools](#available-tools)
- [Usage Examples](#usage-examples)
- [Other Useful MCP Servers](#other-useful-mcp-servers)

---

## What is MCP?

The **Model Context Protocol (MCP)** is an open standard that allows AI assistants like Claude to connect with external tools and services. MCP servers provide Claude Code with additional capabilities like:

- Web scraping and brand extraction (FireCrawl)
- Database access (Supabase, Postgres)
- File system operations
- API integrations

---

## FireCrawl MCP Server

### Overview

FireCrawl is a web scraping API that extracts structured data from websites. The FireCrawl MCP server enables Claude Code to:

- **Extract brand identity** - Colors, fonts, logos from any website
- **Scrape content** - Convert web pages to markdown
- **Compare brands** - Analyze differences between two websites
- **Capture screenshots** - Visual references of web pages

### Prerequisites

1. **Node.js 18+** - Required to run MCP servers
2. **FireCrawl API Key** - Get one at [firecrawl.dev](https://firecrawl.dev)

### Get Your API Key

1. Go to [firecrawl.dev](https://firecrawl.dev)
2. Click "Sign Up" or "Get Started"
3. Create an account (free tier available)
4. Navigate to Dashboard → API Keys
5. Generate a new API key
6. Copy and save it securely

> **Security Note**: Never commit API keys to version control. Store them securely in environment variables or use a secrets manager.

---

## Installation Methods

### Method 1: Claude Code CLI (Recommended)

The simplest way to add an MCP server:

```bash
claude mcp add firecrawl \
  --command "npx" \
  --args "-y" "@anthropic/firecrawl-mcp" \
  --env "FIRECRAWL_API_KEY=fc-your-api-key-here"
```

### Method 2: Manual Configuration

Add to your MCP configuration file:

**Location:**
- macOS/Linux: `~/.claude/mcp_servers.json`
- Windows: `%USERPROFILE%\.claude\mcp_servers.json`

```json
{
  "mcpServers": {
    "firecrawl": {
      "command": "npx",
      "args": ["-y", "@anthropic/firecrawl-mcp"],
      "env": {
        "FIRECRAWL_API_KEY": "fc-your-api-key-here"
      }
    }
  }
}
```

### Method 3: Project-Specific Configuration

For project-specific MCP servers, create `.claude/mcp.json` in your project:

```json
{
  "mcpServers": {
    "firecrawl": {
      "command": "npx",
      "args": ["-y", "@anthropic/firecrawl-mcp"],
      "env": {
        "FIRECRAWL_API_KEY": "${FIRECRAWL_API_KEY}"
      }
    }
  }
}
```

Then set the environment variable:
```bash
export FIRECRAWL_API_KEY="fc-your-api-key-here"
```

---

## Configuration

### Full Configuration Example

```json
{
  "mcpServers": {
    "firecrawl": {
      "command": "npx",
      "args": ["-y", "@anthropic/firecrawl-mcp"],
      "env": {
        "FIRECRAWL_API_KEY": "fc-your-api-key-here"
      },
      "disabled": false
    }
  }
}
```

### Configuration Options

| Option | Type | Description |
|--------|------|-------------|
| `command` | string | The command to run (usually `npx`) |
| `args` | array | Arguments for the command |
| `env` | object | Environment variables |
| `disabled` | boolean | Temporarily disable the server |

### Verify Installation

After configuration, restart Claude Code and check:

```bash
# List configured MCP servers
claude mcp list
```

Or in Claude Code, the MCP tools should appear with prefix `mcp__firecrawl__`:
- `mcp__firecrawl__firecrawl_extract_branding`
- `mcp__firecrawl__firecrawl_scrape_with_branding`
- `mcp__firecrawl__firecrawl_compare_branding`

---

## Available Tools

### 1. Extract Branding

Extract complete brand identity from a website.

**Tool**: `mcp__firecrawl__firecrawl_extract_branding`

**Parameters**:
| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `url` | string | Yes | Website URL to analyze |
| `response_format` | string | No | `markdown` or `json` (default: `markdown`) |

**Returns**:
- Color scheme (light/dark)
- Color palette (primary, secondary, accent, background, text)
- Typography (font families, sizes, weights)
- Spacing values
- UI component styles
- Brand assets (logo, favicon, OG image)
- Brand personality traits

**Example**:
```
Extract the branding from https://stripe.com in JSON format
```

### 2. Scrape with Branding

Combine brand extraction with content scraping.

**Tool**: `mcp__firecrawl__firecrawl_scrape_with_branding`

**Parameters**:
| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `url` | string | Yes | Website URL |
| `include_markdown` | boolean | No | Include page content |
| `include_screenshot` | boolean | No | Include page screenshot |
| `include_links` | boolean | No | Include extracted links |
| `only_main_content` | boolean | No | Exclude headers/footers |

**Example**:
```
Scrape https://linear.app with branding and markdown content
```

### 3. Compare Branding

Compare brand identity between two websites.

**Tool**: `mcp__firecrawl__firecrawl_compare_branding`

**Parameters**:
| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `url1` | string | Yes | First website URL |
| `url2` | string | Yes | Second website URL |
| `response_format` | string | No | `markdown` or `json` |

**Example**:
```
Compare the branding of stripe.com and square.com
```

---

## Usage Examples

### Extract and Add New Brand Theme

```
1. Extract branding from https://notion.com
2. Add it as a new brand theme called "notion" to the document-polisher skill
```

Claude Code will:
1. Use FireCrawl to extract colors, fonts, and styling
2. Update `brand-mapping.json` with the new theme
3. Create a reference file in `brands/notion.md`

### Analyze Competitor Branding

```
Compare the branding of our website https://ourcompany.com with competitor https://competitor.com and summarize the key differences
```

### Create Brand-Consistent Documents

```
1. Extract the branding from https://client-website.com
2. Create a project proposal document styled with their brand
```

### Bulk Brand Extraction

```
Extract branding from these websites and create a comparison table:
- stripe.com
- square.com
- paypal.com
```

---

## Other Useful MCP Servers

### Supabase MCP

Database access and management:

```bash
claude mcp add supabase \
  --command "npx" \
  --args "-y" "@supabase/mcp-server-supabase@latest" \
  --env "SUPABASE_ACCESS_TOKEN=your-token"
```

### 21st Magic (UI Components)

UI component generation:

```bash
claude mcp add magic \
  --command "npx" \
  --args "-y" "@anthropic/21st-magic-mcp"
```

### File System MCP

Enhanced file operations:

```bash
claude mcp add filesystem \
  --command "npx" \
  --args "-y" "@anthropic/filesystem-mcp" \
  --args "/path/to/allowed/directory"
```

---

## Troubleshooting

### MCP Server Not Loading

1. **Check Node.js version**: Requires Node.js 18+
   ```bash
   node --version
   ```

2. **Verify configuration syntax**: JSON must be valid
   ```bash
   cat ~/.claude/mcp_servers.json | python -m json.tool
   ```

3. **Check API key**: Ensure it's set correctly
   ```bash
   echo $FIRECRAWL_API_KEY
   ```

### "Tool not found" Errors

1. Restart Claude Code after adding MCP servers
2. Check that the server isn't disabled
3. Verify the MCP server is listed: `claude mcp list`

### Rate Limiting

FireCrawl has rate limits on the free tier:
- Consider upgrading for production use
- Add delays between bulk requests
- Cache results when possible

### Timeout Errors

Large websites may timeout:
- Try extracting from specific pages instead of homepage
- Use `only_main_content: true` to reduce processing

---

## Security Best Practices

1. **Never commit API keys** - Use environment variables
2. **Use project-specific configs** - Limit scope of credentials
3. **Rotate keys regularly** - Especially if exposed
4. **Review MCP permissions** - Understand what each server can access
5. **Use `.gitignore`** - Exclude configuration files with secrets

---

## Resources

- [MCP Specification](https://modelcontextprotocol.io/)
- [FireCrawl Documentation](https://docs.firecrawl.dev/)
- [Claude Code Documentation](https://claude.com/claude-code)
- [Available MCP Servers](https://github.com/anthropics/mcp-servers)
