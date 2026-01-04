from fastmcp import FastMCP
from pydantic import BaseModel, Field
from typing import Annotated, List, Dict, Optional, Literal
from xlsx_tools import markdown_to_excel
from docx_tools import markdown_to_word
from pptx_tools import create_presentation
from email_tools import create_eml
from email_tools.dynamic_email_tools import register_email_template_tools_from_yaml
from pathlib import Path
import logging
from config import get_config

mcp = FastMCP("MCP Office Documents")

# Initialize config and logging
config = get_config()
logger = logging.getLogger(__name__)

# Look for dynamic email templates in production and local locations.
# Production (container): /app/config/email_templates.yaml
# Local development: <project_root>/config/email_templates.yaml
APP_CONFIG_PATH = Path("/app/config") / "email_templates.yaml"
LOCAL_CONFIG_PATH = Path(__file__).resolve().parent / "config" / "email_templates.yaml"

# Prefer the production path when present, otherwise fall back to local config.
_primary_yaml = None
for candidate in (APP_CONFIG_PATH, LOCAL_CONFIG_PATH):
    if candidate.exists():
        _primary_yaml = candidate
        logger.info("[dynamic-email] Found email templates file: %s", candidate)
        break

if _primary_yaml:
    try:
        register_email_template_tools_from_yaml(mcp, _primary_yaml)
    except Exception as e:
        logger.exception("[dynamic-email] Failed to register email templates from %s: %s", _primary_yaml, e)
else:
    logger.info(
        "[dynamic-email] No dynamic email templates file found at /app/config/email_templates.yaml or config/email_templates.yaml - skipping"
    )

class PowerPointSlide(BaseModel):
    """PowerPoint slide - can be title, section, content, or table slide based on slide_type."""
    slide_type: Literal["title", "section", "content", "table"] = Field(description="Type of slide: 'title' for presentation opening, 'section' for dividers, 'content' for slide with bullet points, 'table' for slide with a table")
    slide_title: str = Field(description="Title text for the slide")

    # Optional fields based on slide type
    author: Optional[str] = Field(default="", description="Author name for title slides - appears in subtitle placeholder. Leave empty for section/content/table slides.")
    slide_text: Optional[List[Dict]] = Field(
        default=None,
        description="Array of bullet points for content slides. Each bullet point must have 'text' (string) and 'indentation_level' (integer 1-5). Leave empty/null for title, section, and table slides."
    )
    table_data: Optional[List[List[str]]] = Field(
        default=None,
        description="Table data for table slides. A list of rows where each row is a list of cell values (strings). The first row is treated as the header row. Leave empty/null for title, section, and content slides."
    )

@mcp.tool(
    name="create_excel_from_markdown",
    description="Converts markdown content with tables and formulas to Excel (.xlsx) format.",
    tags={"excel", "spreadsheet", "data"},
    annotations={"title": "Markdown to Excel Converter"}
)
async def create_excel_document(
    markdown_content: Annotated[str, Field(description="Markdown content containing tables, headers, and formulas. Use T1.B[0] for cross-table references and B[0] for current row references. ALWAYS use [0], [1], [2] notation, NEVER use absolute row numbers like B2, B3. Do NOT count table header as first row, first row has index [0]. Supports cell formatting: **bold**, *italic*.")]
) -> str:
    """
    Converts markdown to Excel with advanced formula support.
    """

    logger.info("Converting markdown to Excel document")

    try:
        result = markdown_to_excel(markdown_content)
        logger.info("Excel document uploaded successfully")
        return result
    except Exception as e:
        logger.error(f"Error creating Excel document: {e}")
        return f"Error creating Excel document: {str(e)}"

@mcp.tool(
    name="create_word_from_markdown",
    description="Converts markdown content to Word (.docx) format. Supports headers, tables, lists, formatting, hyperlinks, and block quotes.",
    tags={"word", "document", "text", "legal", "contract"},
    annotations={"title": "Markdown to Word Converter"}
)
async def create_word_document(
    markdown_content: Annotated[str, Field(description="Markdown content. For LEGAL CONTRACTS use numbered lists (1., 2., 3.) for sections and nested lists for provisions - DO NOT use headers (except for contract title). For other documents use headers (# ## ###).")]
) -> str:
    """
    Converts markdown to professionally formatted Word document.

    """

    logger.info("Converting markdown to Word document")

    try:
        result = markdown_to_word(markdown_content)
        logger.info("Word document uploaded successfully")
        return result
    except Exception as e:
        logger.error(f"Error creating Word document: {e}")
        return f"Error creating Word document: {str(e)}"

@mcp.tool(
    name="create_powerpoint_presentation",
    description="Creates PowerPoint presentations from structured slides.",
    tags={"powerpoint", "presentation", "slides"},
    annotations={"title": "PowerPoint Presentation Creator"}
)
async def create_powerpoint_presentation(
    slides: Annotated[List[dict], Field(
        description="""List of slide objects. Each slide requires 'slide_type' (str) and type-specific fields:

- title: {slide_type: "title", slide_title: str, author?: str}
- section: {slide_type: "section", slide_title: str}
- content: {slide_type: "content", slide_title: str, slide_text: [{text: str, indentation_level: int (1-3)}]}
- table: {slide_type: "table", slide_title: str, table_data: [[str]] (first row = header), header_color?: str (hex), alternate_rows?: bool}
- image: {slide_type: "image", slide_title?: str, image_url: str, image_caption?: str}
- two_column: {slide_type: "two_column", slide_title: str, left_column: [{text: str, indentation_level: int}], right_column: [{text: str, indentation_level: int}], left_heading?: str, right_heading?: str}
- chart: {slide_type: "chart", slide_title: str, chart_type: str (bar|column|line|pie|doughnut|stacked_bar|area), chart_data: {categories: [str], series: [{name: str, values: [number]}]}, has_legend?: bool, legend_position?: str}
- quote: {slide_type: "quote", slide_title?: str, quote_text: str, quote_author?: str}

All slides support optional 'speaker_notes': str field."""
    )],
    format: Annotated[Literal["4:3", "16:9"], Field(
        default="16:9",
        description="Aspect ratio: '16:9' (widescreen) or '4:3' (traditional)"
    )] = "16:9"
) -> str:
    """Creates PowerPoint presentations with structured slide models and professional templates."""

    logger.info(f"Creating PowerPoint presentation with {len(slides)} slides in {format} format")

    try:
        result = create_presentation(slides, format)
        logger.info(f"PowerPoint presentation created: {result}")
        return result
    except Exception as e:
        logger.error(f"Error creating PowerPoint presentation: {e}")
        return f"Error creating PowerPoint presentation: {str(e)}"

@mcp.tool(
    name="create_email_draft",
    description="Creates an email draft in EML format with HTML content using preset professional styling.",
    tags={"email", "eml", "communication"},
    annotations={"title": "Email Draft Creator"}
)
async def create_email_draft(
    content: Annotated[str, Field(description="BODY CONTENT ONLY - Do NOT include HTML structure tags like <html>, <head>, <body>, or <style>. Do NOT include any CSS styling. Use <p> for greetings and for signatures, never headers. Use <h2> for section headers (will be bold), <h3> for subsection headers (will be underlined). HTML tags allowed: <p>, <h2>, <h3>, <ul>, <li>, <strong>, <em>, <div>.")],
    subject: Annotated[str, Field(description="Email subject line")],
    to: Annotated[Optional[List[str]], Field(description="List of recipient email addresses", default=None)],
    cc: Annotated[Optional[List[str]], Field(description="List of CC recipient email addresses", default=None)],
    bcc: Annotated[Optional[List[str]], Field(description="List of BCC recipient email addresses", default=None)],
    priority: Annotated[str, Field(description="Email priority: 'low', 'normal', or 'high'", default="normal")],
    language: Annotated[str, Field(description="Language code for proofreading in Outlook (e.g., 'cs-CZ' for Czech, 'en-US' for English, 'de-DE' for German, 'sk-SK' for Slovak)", default="cs-CZ")]
) -> str:
    """
    Creates professional email drafts in EML format with preset styling and language settings.
    """

    logger.info(f"Creating email draft with subject: {subject}")

    try:
        result = create_eml(
            to=to,
            cc=cc,
            bcc=bcc,
            re=subject,
            content=content,
            priority=priority,
            language=language
        )
        logger.info(f"Email draft created: {result}")
        return result
    except Exception as e:
        logger.error(f"Error creating email draft: {e}")
        return f"Error creating email draft: {str(e)}"

if __name__ == "__main__":
    mcp.run(
        transport="streamable-http",
        host="0.0.0.0",
        port=8958,
        log_level=config.logging.mcp_level_str,
        path="/mcp"
    )
