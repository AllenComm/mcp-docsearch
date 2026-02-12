# docsearch

MCP server for searching and reading binary document files from Claude Code.

## Install

```bash
claude mcp add --scope user docsearch -- uvx --from git+https://github.com/AllenComm/mcp-docsearch docsearch
```

Or per-project:

```bash
claude mcp add docsearch -- uvx --from git+https://github.com/AllenComm/mcp-docsearch docsearch
```

Or add directly to your MCP config (`~/.claude/.mcp.json` for user-scope, `.mcp.json` for project-scope):

```json
{
  "mcpServers": {
    "docsearch": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/AllenComm/mcp-docsearch", "docsearch"]
    }
  }
}
```

Restart Claude Code after changing the config.

## CLAUDE.md

Add to `~/.claude/CLAUDE.md` so Claude Code knows when to use these tools:

```
Use the docgrep and docread MCP tools instead of grep/read for binary documents (PDF, DOCX, PPTX, XLSX, ODT, ODS, ODP, RTF, EPUB).
```

## Supported Formats

| Format | Extension | Extraction |
|--------|-----------|------------|
| PDF | `.pdf` | Page-by-page text |
| Word | `.docx` | Paragraphs + tables |
| PowerPoint | `.pptx` | Slide-by-slide text frames + tables |
| Excel | `.xlsx` | Sheet-by-sheet, tab-separated rows |
| OpenDocument Text | `.odt` | Paragraphs |
| OpenDocument Spreadsheet | `.ods` | Sheet-by-sheet, tab-separated rows |
| OpenDocument Presentation | `.odp` | Slide-by-slide text |
| Rich Text Format | `.rtf` | Full text |
| EPUB | `.epub` | Chapter-by-chapter text (spine order) |

## Tools

### `docgrep`

Search through documents for text matching a regex pattern.

**Parameters:**
- `directory` (required) — path to search recursively
- `pattern` (required) — regex pattern to match
- `case_sensitive` — default `false`
- `file_types` — filter to specific extensions, e.g. `["pdf", "docx"]`
- `max_results` — default `100`

**Output format:** `filepath:section:matching_line`

### `docread`

Read full text content from a single document.

**Parameters:**
- `filepath` (required) — path to the document
- `range` — filter to specific sections by format:
  - **PDF:** page numbers, e.g. `"1-5"`, `"3"`, `"1,3,5-7"`
  - **PPTX/ODP:** slide numbers, e.g. `"2-3"`
  - **XLSX/ODS:** sheet name or 1-based index, with optional row range after colon, e.g. `"1"`, `"Sheet1"`, `"1:1-100"`, `"Revenue:50-200"`
  - **EPUB:** chapter numbers, e.g. `"1-5"`
  - **DOCX/ODT/RTF:** line numbers (e.g. `"1-50"`, `"100-200"`)

Output is auto-truncated at 40,000 characters. Use `range` to narrow results for large documents.

## Development

### Testing with MCP Inspector

```bash
npx @modelcontextprotocol/inspector uv run docsearch
```

### Testing with Claude Code

```bash
claude mcp add --scope user docsearch -- uv run --directory /path/to/docsearch docsearch
```

## Examples

```
# Search all docs in a directory for "quarterly revenue"
docgrep(directory="/home/user/reports", pattern="quarterly revenue")

# Read a specific PDF, pages 1-3
docread(filepath="/home/user/reports/q4.pdf", range="1-3")

# Search only Word docs for a pattern
docgrep(directory="/home/user/docs", pattern="TODO|FIXME", file_types=["docx"])

# Read first 100 rows of the first sheet in an Excel file
docread(filepath="/home/user/data/sales.xlsx", range="1:1-100")

# Read rows 50-200 of a sheet by name
docread(filepath="/home/user/data/sales.xlsx", range="Revenue:50-200")
```
