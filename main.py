from fastmcp import FastMCP
from pydantic import BaseModel, Field
from typing import Annotated, List, Dict, Optional, Literal
from create_xlsx import markdown_to_excel
from docx_tools import markdown_to_word
from create_pptx import create_presentation
from email_tools import create_eml
from email_tools.dynamic_email_tools import register_email_template_tools_from_yaml
from pathlib import Path
import logging
from config import get_config

mcp = FastMCP("MCP Office Documents")

# Initialize config and logging
config = get_config()
logger = logging.getLogger(__name__)

# Dynamic email tools: ONLY load from config/email_templates.yaml
_config_dir = Path(__file__).resolve().parent.parent / "config"
_primary_yaml = _config_dir / "email_templates.yaml"
if _primary_yaml.exists():
    register_email_template_tools_from_yaml(mcp, _primary_yaml)
else:
    logger.info("[dynamic-email] No dynamic email templates file found at config/email_templates.yaml - skipping")

class PowerPointSlide(BaseModel):
    """PowerPoint slide - can be title, section, or content slide based on slide_type."""
    slide_type: Literal["title", "section", "content"] = Field(description="Type of slide: 'title' for presentation opening, 'section' for dividers, 'content' for slide with bullet points")
    slide_title: str = Field(description="Title text for the slide")

    # Optional fields based on slide type
    author: Optional[str] = Field(default="", description="Author name for title slides - appears in subtitle placeholder. Leave empty for section/content slides.")
    slide_text: Optional[List[Dict]] = Field(
        default=None,
        description="Array of bullet points for content slides. Each bullet point must have 'text' (string) and 'indentation_level' (integer 1-5). Leave empty/null for title and section slides."
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
    description="Creates PowerPoint presentations with professional templates using structured slide models.",
    tags={"powerpoint", "presentation", "slides"},
    annotations={"title": "PowerPoint Presentation Creator"}
)
async def create_powerpoint_presentation(
    slides: List[PowerPointSlide],
    format: Annotated[Literal["4:3", "16:9"], Field(
        default="4:3",
        description="Presentation formating: '4:3' for traditional or '16:9' for widescreen"
    )]
) -> str:
    """Creates PowerPoint presentations with structured slide models and professional templates."""

    logger.info(f"Creating PowerPoint presentation with {len(slides)} slides in {format} format")

    try:
        slides_data = [slide.model_dump() for slide in slides]
        result = create_presentation(slides_data, format)
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
