import io
from pathlib import Path
from email.mime.text import MIMEText  # fixed module path
from email.utils import formatdate
from email.header import Header
import pystache
import html

# Robust import for upload_file regardless of sys.path setup
try:
    from upload_file import upload_file  # type: ignore
except ImportError:  # pragma: no cover
    import os, sys
    sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    from upload_file import upload_file  # type: ignore

# Template file name
TEMPLATE_FILENAME = "general_template.html"
PROD_TEMPLATE_PATH = Path("/app/templates") / TEMPLATE_FILENAME  # external override (production / mounted)
# Built-in template now expected directly in this package directory for simplicity
INTERNAL_TEMPLATE_PATH = Path(__file__).parent / TEMPLATE_FILENAME

def _load_template() -> str:
    """Load the email HTML template.

    Priority:
      1. External override: /app/templates/general_template.html (e.g., Docker volume mount)
      2. Built-in package template: email_tools/general_template.html (current design)

    Raises FileNotFoundError if none exist.
    """
    if PROD_TEMPLATE_PATH.exists():
        with open(PROD_TEMPLATE_PATH, "r", encoding="utf-8") as f:
            return f.read()
    if INTERNAL_TEMPLATE_PATH.exists():
        with open(INTERNAL_TEMPLATE_PATH, "r", encoding="utf-8") as f:
            return f.read()


    raise FileNotFoundError(
        "Email template not found."
    )


def create_eml(to=None, cc=None, bcc=None, re=None, content=None, priority="normal", language="cs-CZ"):
    """Create an unsent email draft (EML) using a Mustache HTML template.

    Template variables:
      {{language}}  - inserted into lang attributes (sanitized)
      {{subject}}   - inserted (HTML-escaped) into <title>
      {{{content}}} - raw HTML fragment for email body (caller restricted to allowed tags)
    """

    # Validate priority
    if priority and priority.lower() not in ["low", "normal", "high"]:
        raise ValueError("Priority must be 'low', 'normal', or 'high'")

    if not content:
        raise ValueError("Email content is required")
    if not re:
        raise ValueError("Email subject is required")

    template_html = _load_template()

    # Prepare context
    safe_language = (language or "").replace('"', '').replace("'", '')
    escaped_subject = html.escape(re or "")

    renderer = pystache.Renderer(escape=lambda u: u)  # We'll manually escape where needed
    context = {
        "language": safe_language,   # safe for attribute insertion
        "subject": escaped_subject,  # already escaped
        "content": content,          # inserted unescaped via triple braces {{{content}}}
    }
    complete_html = renderer.render(template_html, context)

    buffer = None
    try:
        msg = MIMEText(complete_html, 'html', 'utf-8')
        # Ensure proper encoding (base64 avoids quoted-printable soft breaks generating '=')
        if 'Content-Transfer-Encoding' in msg:
            msg.replace_header('Content-Transfer-Encoding', 'base64')
        else:
            msg.add_header('Content-Transfer-Encoding', 'base64')

        if to:
            msg['To'] = ", ".join(to)
        if cc:
            msg['Cc'] = ", ".join(cc)
        if bcc:
            msg['Bcc'] = ", ".join(bcc)

        msg['Subject'] = Header(re, 'utf-8')
        msg['Date'] = formatdate(localtime=True)
        msg['Content-Language'] = safe_language
        msg['Accept-Language'] = safe_language

        if priority.lower() == 'high':
            msg['X-Priority'] = '1 (Highest)'
            msg['X-MSMail-Priority'] = 'High'
            msg['Importance'] = 'High'
        elif priority.lower() == 'low':
            msg['X-Priority'] = '5 (Lowest)'
            msg['X-MSMail-Priority'] = 'Low'
            msg['Importance'] = 'Low'

        msg['X-Unsent'] = '1'

        buffer = io.BytesIO()
        msg_bytes = msg.as_bytes()
        buffer.write(msg_bytes)
        buffer.seek(0)

        return upload_file(buffer, "eml")
    except Exception as e:
        raise Exception(f"Failed to create email draft: {e}")
    finally:
        if buffer:
            buffer.close()
