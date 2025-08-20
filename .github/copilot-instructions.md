# Copilot Instructions for MCP MS Office Documents Server

## Repository Overview

This repository contains an MCP (Model Context Protocol) server that empowers AI assistants to create professional Microsoft Office documents. The server provides tools for generating PowerPoint presentations, Word documents, Excel spreadsheets, and email drafts with professional formatting and customizable templates.

## Architecture

### Core Components

- **MCP Server** (`src/main.py`): FastMCP-based server that exposes document creation tools
- **Document Creators**: Specialized modules for each document type
  - `src/create_pptx.py`: PowerPoint presentation generation
  - `src/create_docx.py`: Word document generation from Markdown
  - `src/create_xlsx.py`: Excel spreadsheet creation with formula support
  - `src/create_msg.py`: Email draft creation in EML format
- **File Management** (`src/upload_file.py`): Handles file storage (local or S3)
- **Templates**: Built-in and custom template support for professional formatting

### Key Technologies

- **FastMCP**: MCP server framework for tool exposure
- **Pydantic**: Data validation and serialization
- **python-pptx**: PowerPoint file manipulation
- **python-docx**: Word document creation
- **openpyxl**: Excel file handling
- **boto3**: AWS S3 integration for cloud storage

## Code Patterns and Conventions

### MCP Tool Definitions

Tools are defined using the `@mcp.tool` decorator with comprehensive metadata:

```python
@mcp.tool(
    name="tool_name",
    description="Clear description of what the tool does",
    tags={"category", "type"},
    annotations={"title": "Human Readable Title"}
)
async def tool_function(
    param: Annotated[Type, Field(description="Detailed parameter description")]
) -> str:
    """Tool implementation with error handling and logging"""
```

### Error Handling Pattern

All tools follow consistent error handling:

```python
try:
    # Document creation logic
    result = create_document(parameters)
    print(f"Document created successfully")
    return result
except Exception as e:
    print(f"Error creating document: {e}")
    return f"Error creating document: {str(e)}"
```

### Template Loading Pattern

Document creators implement template discovery with fallbacks:

```python
def load_templates():
    """Loads templates with multiple fallback locations"""
    # 1. Custom templates in /app/templates (production)
    # 2. Development templates in ../templates
    # 3. Built-in templates in src/
    # 4. Fallback to default library templates
```


## File Structure

```
├── src/
│   ├── main.py              # MCP server and tool definitions
│   ├── create_pptx.py       # PowerPoint generation
│   ├── create_docx.py       # Word document creation
│   ├── create_xlsx.py       # Excel spreadsheet handling
│   ├── create_msg.py        # Email draft creation
│   ├── upload_file.py       # File storage management
│   └── template_*           # Built-in templates
├── .github/
│   └── copilot-instructions.md
├── Dockerfile               # Container configuration
├── docker-compose.yml       # Service orchestration
├── requirements.txt         # Python dependencies
├── instructions_template.md # AI agent usage guide
└── Readme.md               # User documentation
```

## Testing and Deployment

- Docker-based deployment with volume mounts for templates and output
- Compatible with LibreChat, Claude Desktop, and other MCP clients
- Configurable file storage (local filesystem or AWS S3)
- Port 8958 for HTTP streaming connections

## AI Assistant Integration

This server is designed to work with AI assistants through MCP. The `instructions_template.md` provides specific guidance for PowerPoint creation workflows. When extending functionality, consider:

- Clear parameter descriptions for AI understanding
- Structured data models for complex inputs
- Comprehensive error messages for debugging
- Examples in documentation for AI reference

## Best Practices

1. **Parameter Validation**: Use Pydantic models for complex structured input
2. **Template Flexibility**: Support both custom and built-in templates
3. **Error Recovery**: Graceful fallbacks when templates or configurations fail
4. **Documentation**: Keep tool descriptions and parameter docs comprehensive
5. **Logging**: Include detailed logging for debugging document creation issues
6. **Security**: Validate file paths and sanitize inputs for file operations
