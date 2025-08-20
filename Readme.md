# MCP MS Office Documents Server

A [Model Context Protocol](https://modelcontextprotocol.io/) server that empowers AI assistants to create professional Microsoft Office documents on your behalf.

## 🚀 What You Can Do

Transform your conversations with AI into professional documents:

- **📊 PowerPoint Presentations** - Create professional slideshows with title slides, section dividers, and content slides featuring bullet points and proper formatting
- **📄 Word Documents** - Generate formatted documents from markdown including headers, tables, lists, and styling - perfect for reports, contracts, and documentation  
- **📧 Email Drafts** - Compose professional email drafts in EML format with proper HTML formatting and styling
- **📈 Excel Spreadsheets** - Build data-rich spreadsheets from markdown tables with formula support and cross-table references

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
    url: http://localhost:8958/sse  # Adjust URL if running on different host
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
      "args": ["-e", "require('http').get('http://localhost:8958/sse')"]
    }
  }
}
```

### Other MCP Clients

The server exposes an SSE (Server-Sent Events) endpoint at `/sse` and follows the standard MCP protocol. Consult your MCP client's documentation for connection details.

## 🎨 Advanced Features

### Custom Templates

Use your own company templates and branding:

1. Create template files:
   - `template_4_3.pptx` (4:3 aspect ratio PowerPoint)
   - `template_16_9.pptx` (16:9 aspect ratio PowerPoint)  
   - `template.docx` (Word document)

2. Mount the template directory:
```yaml
volumes:
  - ./templates:/app/templates
```

**Template Requirements:**
- PowerPoint: Title slide layout must be 3rd, content slide layout must be 5th, section slide layout must be 8th in master slides
- Word: Must contain standard Word styles for proper formatting

### Email Styling Customization

Customize the appearance of generated email drafts with your own styling:

1. Create a `config.yaml` file with your email styles (see `config.yaml` template in the repository)
2. Mount the config directory:
```yaml
volumes:
  - ./config:/app/config
```

**Example config.yaml for corporate branding:**
```yaml
email:
  styles:
    base:
      body:
        font-family: "Calibri, Arial, sans-serif"
        font-size: "11pt"
        color: "#333333"
        line-height: "1.6"
    elements:
      h2:
        font-weight: "bold"
        font-size: "14pt"
        color: "#1f4788"  # Corporate blue
        border-bottom: "2px solid #1f4788"
      h3:
        font-weight: "bold"
        color: "#1f4788"
        text-decoration: "none"
```

If no config file is provided, emails will use the default styling.

### Usage Tips

For best results when working with AI assistants:

- **Presentations**: Ask for structured content with clear sections and bullet points
- **Documents**: Use markdown formatting in your requests for better results
- **Emails**: Specify tone, recipients, and key points you want to cover
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
- [x] Email styling customization via config.yaml
- [ ] Additional template customization options