# MCP Office Documents Server

This server lets AI assistants generate professional PowerPoint, Word, Excel, and EML email drafts through the Model Context Protocol (MCP).

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
  - Required input: slides array with slide_type and slide_title; optional author and slide_text for content slides
  - Format: `4:3` (default) or `16:9`

- create_word_from_markdown
  - Converts Markdown to .docx, supporting headers, lists, tables, inline formatting, links, block quotes

- create_excel_from_markdown
  - Converts Markdown tables and headers to .xlsx
  - Supports formulas and relative/table references (e.g., `=B[0]`, `T1.SUM(B[0]:E[0])`)

- create_email_draft
  - Creates an EML draft with an HTML body using a preset wrapper template
  - Accepts subject, to/cc/bcc, priority, language, and raw content (no <html>/<body>/<style>)

Dynamic email tools (optional):
- If `config/email_templates.yaml` exists, each entry is registered as its own email-draft tool at startup. See below for details.

- Short explanation: Dynamic email templates are reusable, parameterized HTML email layouts defined in `config/email_templates.yaml`. At startup the server registers each template as an individual MCP tool and automatically adds standard fields (subject, to, cc, bcc, priority, language). Template-specific arguments (for example `first_name` or `promo_code`) are exposed as tool parameters so AI assistants can call a single, strongly-typed tool to produce consistent, production-ready emails without composing full HTML bodies.

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
      - name: subject
        type: string
        required: true
      - name: first_name
        type: string
        required: true
      - name: promo_code
        type: string
        required: false
```
- Minimal example HTML (place in `custom_templates/welcome_email.html`):
```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <title>{{subject}}</title>
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
