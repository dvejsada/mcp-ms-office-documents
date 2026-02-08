# MCP Office Documents Server

This server lets AI assistants generate professional PowerPoint, Word, Excel, XML, and EML email drafts through the Model Context Protocol (MCP).

## 1) Installation

- Requirements: Docker and Docker Compose
- Prepare folders (create if missing):
  - `output/` – where files are saved for LOCAL strategy
  - `custom_templates/` – optional custom Office/email templates
  - `config/` – configuration (e.g., `email_templates.yaml`, credentials)
- Get docker-compose.yml (copy or download into this folder):
```cmd
curl -L -o docker-compose.yml https://raw.githubusercontent.com/dvejsada/mcp-ms-office-docs/main/docker-compose.yml
```
  (If you already cloned this repository, you can just copy the existing `docker-compose.yml`.)
- Configure environment:
  - Copy `.env.example` to `.env`
  - Edit `.env`:
    - `LOG_LEVEL=INFO|DEBUG`
    - `UPLOAD_STRATEGY=LOCAL|S3|GCS|AZURE|MINIO`
    - For S3/GCS/AZURE/MINIO, fill the required credentials
- Start the server:
```bash
docker-compose up -d
```
- MCP endpoint: `http://localhost:8958/mcp`


## 2) Tool overview

The server exposes these MCP tools:

- create_powerpoint_presentation
  - Creates a .pptx from structured slides (title, section, content) with optional templates
  - Format: `4:3`  or `16:9`(default)

- create_word_from_markdown
  - Converts Markdown to .docx, supporting headers, lists, tables, inline formatting, links, block quotes
  - **Markdown formatting:** All placeholder values support full markdown:
    - Inline: `**bold**`, `*italic*`, `` `code` ``, `[links](url)`
    - Block-level: headings (`#`), bullet lists (`-`, `*`, `+`), numbered lists (`1.`, `2.`)

- create_excel_from_markdown
  - Converts Markdown tables and headers to .xlsx
  - Supports formulas and relative/table references (e.g., `=B[0]`, `T1.SUM(B[0]:E[0])`)

- create_email_draft
  - Creates an EML draft with an HTML body using a preset wrapper template
  - Accepts subject, to/cc/bcc, priority, language, and raw content (no <html>/<body>/<style>)

- create_xml_file
  - Creates an XML file from provided XML content
  - Validates that input is well-formed XML before saving
  - XML declaration (`<?xml version="1.0"?>`) is added automatically if missing

Dynamic email tools (optional):
- If `config/email_templates.yaml` exists, each entry is registered as its own email-draft tool at startup. See below for details.

- Short explanation: Dynamic email templates are reusable, parameterized HTML email layouts defined in `config/email_templates.yaml`. At startup the server registers each template as an individual MCP tool and automatically adds standard fields (subject, to, cc, bcc, priority, language). Template-specific arguments (for example `first_name` or `promo_code`) are exposed as tool parameters so AI assistants can call a single, strongly-typed tool to produce consistent, production-ready emails without composing full HTML bodies.

Dynamic DOCX template tools (optional):
- If `config/docx_templates.yaml` exists, each entry is registered as its own document generation tool at startup. See below for details.

- Short explanation: Dynamic DOCX templates are reusable Word documents with placeholders (`{{placeholder_name}}`) defined in `config/docx_templates.yaml`. At startup, the server registers each template as an individual MCP tool. Template-specific arguments are exposed as tool parameters. Placeholder values support the same markdown formatting as described above for `create_word_from_markdown`.

Outputs:
- LOCAL: files saved to `output/` and reported back
- S3/GCS/AZURE/MINIO: a time-limited download link is returned (TTL via `SIGNED_URL_EXPIRES_IN`)

### MinIO private storage

You can point the server to a self-hosted MinIO instance (or any other S3 compatible storage) by setting `UPLOAD_STRATEGY=MINIO` and providing:

```
MINIO_ENDPOINT=https://minio.example.com
MINIO_ACCESS_KEY=...
MINIO_SECRET_KEY=...
MINIO_BUCKET=office-documents
MINIO_REGION=us-east-1          # optional, defaults to us-east-1
MINIO_VERIFY_SSL=true           # optional, set to false for self-signed endpoints
MINIO_PATH_STYLE=true           # optional, defaults to true (recommended for MinIO)
SIGNED_URL_EXPIRES_IN=3600      # optional TTL in seconds for download links
```

The MinIO backend reuses the existing boto3 dependency and generates pre-signed download links just like the AWS S3 strategy. Ensure the bucket exists and the provided credentials have `s3:PutObject`/`s3:GetObject` permissions.

## 3) Custom templates

You can provide custom templates for PowerPoint, Word, and email.

Place files in `custom_templates/`.

- PowerPoint: `custom_pptx_template_4_3.pptx`, `custom_pptx_template_16_9.pptx`

- Word: `custom_docx_template.docx`

- Email wrapper : `custom_email_template.html` - base your template on `default_templates/default_email_template.html`

Dynamic email templates (optional):

- Create `config/email_templates.yaml` and reference HTML files by filename only (no paths).
- Example entry:
```yaml
templates:
  - name: welcome_email
    description: Welcome email with optional promo code
    html_path: welcome_email.html  # file must exist in custom_templates/ or default_templates/
    annotations:
      title: Welcome Email
    args:
      - name: first_name
        type: string
        description: Recipient's first name
        required: true
      - name: promo_code
        type: string
        description: Optional promotional code (html formatted)
        required: false
```
- Minimal example HTML (place in `custom_templates/welcome_email.html`):
```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
</head>
<body>
  <h2>Welcome {{first_name}}!</h2>
  <p>We’re excited to have you on board.</p>
  {{{promo_code_block}}}
  <p>Regards,<br/>Support Team</p>
</body>
</html>
```
- Subject, to, cc, bcc, priority, and language are handled automatically and added to each template tool.
- Tip: use `{{variable}}` for escaped text; `{{{variable}}}` for raw HTML

### Dynamic DOCX Templates

Dynamic DOCX templates allow you to create reusable Word document templates with placeholders that support full markdown formatting.

#### Setup

1. Create `config/docx_templates.yaml` and reference DOCX template files by filename only (no paths).
2. Place your `.docx` template files in `custom_templates/` directory.
3. Use `{{placeholder_name}}` syntax in your Word document for placeholders.
4. Placeholders can be placed in document body, tables, headers, and footers.

#### Example YAML configuration

Create `config/docx_templates.yaml`:
```yaml
templates:
  - name: formal_letter
    description: Generate a formal business letter. Placeholder values support markdown formatting (see tool overview).
    docx_path: letter_template.docx  # must exist in custom_templates/ or default_templates/
    annotations:
      title: Formal Letter Generator
    args:
      - name: recipient_name
        type: string
        description: Full name of the recipient
        required: true
      - name: recipient_address
        type: string
        description: Recipient's address (use line breaks for multiple lines)
        required: true
      - name: subject
        type: string
        description: Letter subject
        required: true
      - name: body
        type: string
        description: Main body of the letter (supports markdown)
        required: true
      - name: sender_name
        type: string
        description: Sender's name
        required: true
      - name: date
        type: string
        description: Letter date
        required: false
        default: ""
```

#### Example DOCX template

Create in Word and save as `custom_templates/letter_template.docx`:
```
{{date}}

{{recipient_name}}
{{recipient_address}}

Subject: {{subject}}

{{salutation}}

{{body}}

{{closing}}

{{sender_name}}
{{sender_title}}
```

#### Markdown Formatting Support

**Inline formatting** (works within any text):
| Markdown | Result |
|----------|--------|
| `**bold text**` | Bold text |
| `*italic text*` | Italic text |
| `` `code` `` | Monospace font (Courier New) |
| `[link text](url)` | Clickable hyperlink |
| `**bold with *nested italic***` | Nested formatting |

**Block-level markdown** (creates new paragraphs):

| Markdown | Description |
|----------|-------------|
| `# Heading 1` through `###### Heading 6` | Heading levels 1-6 |
| `- item` or `* item` or `+ item` | Bullet list |
| `1. item`, `2. item` | Numbered list |
| 3 spaces + list marker | Nested list (up to 3 levels) |

Example of block-level content in a placeholder value:
```markdown
Here are the key points:

1. First important item
2. Second important item
   - Sub-point A
   - Sub-point B
3. Third item with **bold** emphasis

## Next Steps

- Review the proposal
- Schedule a follow-up meeting
```

#### Word Styles Used

When creating custom templates, ensure these Word styles exist for proper formatting:

**Heading styles:**
- Heading 1 through Heading 6 (used for markdown `#` headings)

**List styles:**
- List Bullet, List Bullet 2, List Bullet 3 (for nested bullet lists)
- List Number, List Number 2, List Number 3 (for nested numbered lists)

**Other styles:**
- Quote (for blockquotes - base tool only)
- Table Grid (for markdown tables - base tool only)
- Normal (regular paragraphs)

> **Tip:** You can customize these styles in your template (font, size, color, spacing) and the system will use your customizations.

#### Additional Features

- **Font preservation:** The original font size and name from the placeholder location in the template are preserved for inline replacement text.
- **Table support:** Placeholders work inside table cells with full markdown formatting.
- **Header/Footer support:** Placeholders in document headers and footers are replaced (inline formatting only).

