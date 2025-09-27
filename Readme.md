# MCP MS Office Documents Server

A [Model Context Protocol](https://modelcontextprotocol.io/) server that empowers AI assistants to create professional Microsoft Office documents on your behalf.

## 🚀 What You Can Do

Transform your conversations with AI into professional documents:

- **📊 PowerPoint Presentations** - Create professional slideshows with title slides, section dividers, and content slides featuring bullet points and proper formatting
- **📄 Word Documents** - Generate formatted documents from markdown including headers, tables, lists, and styling - perfect for reports, contracts, and documentation  
- **📧 Email Drafts** - Compose professional email drafts in EML format with proper HTML formatting and styling
- **📈 Excel Spreadsheets** - Build data-rich spreadsheets from markdown tables with formula support and cross-table references
- **⚙️ Dynamic Email Template Tools** - Auto-generate additional specialized email draft tools via a simple Mustache-based YAML configuration

All documents are created with professional templates and can be customized with your own branding.

## 🛠️ Quick Setup

### 1. Run the Container

Using Docker Compose (recommended):

```bash
# Download the docker-compose.yml file
curl -O https://raw.githubusercontent.com/dvejsada/mcp-ms-office-documents/main/docker-compose.yml

# Edit the environment variables (see configuration below)
nano docker-compose.yml

# Start the server
docker-compose up -d
```

The server will be available at `http://localhost:8958`

### 2. Configure File Upload Strategy

Choose how generated files are delivered to you by setting the `UPLOAD_STRATEGY` environment variable:

**Option A: Local Files** (files saved to your computer)
```yaml
environment:
  UPLOAD_STRATEGY: LOCAL
volumes:
  - ./output:/app/output  # Files will be saved here
```

**Option B: AWS S3** (files uploaded to cloud storage)
```yaml
environment:
  UPLOAD_STRATEGY: S3
  AWS_ACCESS_KEY: your_access_key
  AWS_SECRET_ACCESS_KEY: your_secret_key  
  AWS_REGION: your-region
  S3_BUCKET: your-bucket-name
```

## 🔗 Connect to MCP Clients

### LibreChat

Add this configuration to your `librechat.yaml` file:

```yaml
mcpServers:
  office-docs:
    type: streamable-http
    url: http://localhost:8958/mcp  # Adjust URL if running on different host
    timeout: 120000  # Allow extra time for document generation
```

After updating the configuration:
1. Restart LibreChat
2. Create a new agent or edit an existing one
3. Add the MCP Office Documents tools to your agent
4. Start creating documents by asking your agent!

### Claude Desktop

Add this configuration to your Claude Desktop MCP settings:

```json
{
  "mcpServers": {
    "office-docs": {
      "command": "node",
      "args": ["-e", "require('http').get('http://localhost:8958/mcp')"]
    }
  }
}
```

### Other MCP Clients

The server exposes a streamable-http endpoint at `/mcp` and follows the standard MCP protocol. Consult your MCP client's documentation for connection details.

## 🎨 Advanced Features

### Custom Templates

Use your own company templates and branding:

1. Create template files:
   - `template_4_3.pptx` (4:3 aspect ratio PowerPoint)
   - `template_16_9.pptx` (16:9 aspect ratio PowerPoint)  
   - `template.docx` (Word document)  
   - `general_template.html` (email HTML wrapper / styling) – mount under `templates/` to override built‑in email appearance.

2. Mount the template directory:
```yaml
volumes:
  - ./templates:/app/templates
```

**Template Requirements:**
- PowerPoint: Title slide layout must be 3rd, content slide layout must be 5th, section slide layout must be 8th in master slides
- Word: Must contain standard Word styles for proper formatting
- Email: `general_template.html` must include Mustache placeholders `{{{content}}}` for body HTML, `{{subject}}` (optional for <title>), and may include `{{language}}`.

### Dynamic Email Template Tools (Simplified Mustache-Only)

Define additional specialized email draft tools without writing Python code by placing an `email_templates.yaml` file in `config/` (mounted at `/app/config/email_templates.yaml`). On server startup each entry becomes its own MCP tool.

Example `config/email_templates.yaml`:
```yaml
templates:
  - name: welcome_email
    description: Welcome email with optional promo code
    html_path: templates/welcome_format.html
    annotations:
      title: Welcome Email (Dynamic)
    args:
      - name: first_name
        type: string
        required: true
        description: Recipient first name
      - name: promo_code
        type: string
        required: false
        description: Optional promotional code
```

Template snippet (`templates/welcome_format.html`):
```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <title>{{subject}}</title>
  <style>
    body { font-family: Calibri, Arial, sans-serif; font-size: 14px; color: #222; }
    h2 { font-size: 18px; margin-bottom: 4px; }
    .promo { background:#f5f5f5; padding:8px 12px; border-left:4px solid #0066cc; margin-top:16px; }
  </style>
</head>
<body>
  <h2>Welcome {{first_name}}!</h2>
  <p>We're excited to have you on board.</p>
  {{{promo_code_block}}}
  <p style="margin-top:24px;">Regards,<br/>Support Team</p>
</body>
</html>
```

#### Placeholder Escaping vs Raw HTML

Mustache offers two syntaxes for inserting values:

- `{{variable}}` (double braces): Inserts the value with HTML escaping. Use this for normal text (names, emails, links, notes, etc.).
- `{{{variable}}}` (triple braces): Inserts the value without escaping (raw HTML). Use only for values intended to contain simple HTML markup.

#### Enumerations (enum)
Add `enum: [value1, value2, ...]` to an argument in `email_templates.yaml` to restrict its accepted values. At runtime the tool will validate the value; invalid options are rejected before rendering. Example from `welcome_email`:
```yaml
- name: tone
  type: string
  required: false
  enum: ["casual", "formal", "friendly"]
  description: Tone variant inserted into template (enum)
```
If a `default` is provided it must be one of the listed values; otherwise it is ignored.

#### Defaults (default)
You can supply a `default:` value for any argument (enum or non‑enum). Notes:
- If `required: false` and a default is present, the default is used when the caller omits the argument.
- If `required: true` and you also give a default, the field effectively becomes optional (the default is applied when omitted).
- For enum arguments the default must be one of the enum values (otherwise it is ignored and the field remains required/optional as specified).
- Omit `default` entirely if you want the tool to force the caller to provide a value (set `required: true`).

Example with enum default (shown in `welcome_email`):
```yaml
- name: tone
  type: string
  required: false
  enum: ["casual", "formal", "friendly"]
  default: "friendly"
  description: Tone variant inserted into template
```

Example non‑enum default:
```yaml
- name: footer_note
  type: string
  required: false
  default: "This message is confidential."
  description: Optional footer line appended at the end
```

### Usage Tips

For best results when working with AI assistants:

- **Presentations**: Ask for structured content with clear sections and bullet points
- **Documents**: Use markdown formatting in your requests for better results
- **Emails**: Specify tone, recipients, and key points you want to cover
- **Dynamic Email Tools**: Provide only the defined parameters; the server handles HTML assembly and uploading
- **Spreadsheets**: Describe your data structure and any calculations needed

See `instructions_template.md` for detailed agent configuration examples.

## 🔧 Requirements

- Docker and Docker Compose
- An MCP-compatible client (LibreChat, Claude Desktop, etc.)
- For S3 upload: AWS account with S3 access

## 🤝 Contributing

Contributions are welcome! Feel free to submit issues, feature requests, or pull requests.

### Development Roadmap

- [x] PowerPoint presentations (pptx) 
- [x] Word documents (docx)
- [x] Email drafts (eml)
- [x] Excel spreadsheets (xlsx)
- [x] Dynamic email template tools
