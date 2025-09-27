"""Dynamic registration of email draft MCP tools from a YAML configuration (Mustache-only). """
from __future__ import annotations

import io
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Dict, Optional, Literal

import yaml
import pystache
from pydantic import Field, create_model
from fastmcp import FastMCP

try:  # local import pattern consistent with other modules
    from upload_file import upload_file  # type: ignore
except ImportError:  # pragma: no cover
    import os, sys
    sys.path.append(os.path.abspath(Path(__file__).resolve().parent.parent))
    from upload_file import upload_file  # type: ignore

__all__ = ["register_email_template_tools_from_yaml"]

TYPE_MAP = {
    "string": str, "str": str,
    "int": int, "integer": int,
    "float": float,
    "bool": bool, "boolean": bool,
    "list": list[str], "list[str]": list[str], "list[string]": list[str],
    "dict": dict, "object": dict,
}

BASE_FIELDS: Dict[str, Any] = {
    "subject": (str, Field(..., description="Email subject line (also sets Subject header)")),
    "to": (Optional[list[str]], Field(None, description="List of recipient email addresses")),
    "cc": (Optional[list[str]], Field(None, description="List of CC recipient email addresses")),
    "bcc": (Optional[list[str]], Field(None, description="List of BCC recipient email addresses")),
}


def register_email_template_tools_from_yaml(mcp: FastMCP, yaml_path: Path) -> None:

    try:
        cfg = yaml.safe_load(yaml_path.read_text(encoding="utf-8")) or {}
    except Exception as e:  # pragma: no cover
        print(f"[dynamic-email] Failed to load YAML '{yaml_path}': {e}")
        return

    templates = cfg.get("templates") or []
    if not isinstance(templates, list):
        print("[dynamic-email] 'templates' key must be a list – skipping.")
        return

    for spec in templates:
        try:
            # Load tool definition
            name = spec["name"]
            description = spec.get("description")
            annotations = spec.get("annotations")
            meta = spec.get("meta")
            html_path = spec.get("html_path")

            if not html_path:
                print(f"[dynamic-email] Missing html_path for {name}, skipping.")
                continue

            template_path = (yaml_path.parent / html_path).resolve()

            if not template_path.exists():
                print(f"[dynamic-email] Template file missing for {name}: {template_path}")
                continue

            html_source = template_path.read_text(encoding="utf-8")

            # Build Pydantic model fields starting with base fields
            fields: Dict[str, Any] = dict(BASE_FIELDS)

            for arg in spec.get("args", []):
                arg_name = arg.get("name")
                if not arg_name or arg_name in fields:
                    continue  # skip duplicates / base overrides

                enum_values = arg.get("enum")
                if enum_values and isinstance(enum_values, list) and enum_values:
                    # Infer literal value types (all ints -> int, all floats -> float, else str)
                    if all(isinstance(v, int) for v in enum_values):
                        lit_values = tuple(int(v) for v in enum_values)
                    elif all(isinstance(v, (int, float)) for v in enum_values):
                        lit_values = tuple(float(v) for v in enum_values)
                    else:
                        lit_values = tuple(str(v) for v in enum_values)
                    py_type = Literal[lit_values]  # type: ignore[index]
                    required = bool(arg.get("required", True))
                    default = arg.get("default", (Ellipsis if required else None))
                    if default is not Ellipsis and default is not None and default not in lit_values:
                        print(f"[dynamic-email] Default '{default}' not in enum for {arg_name}; ignoring default.")
                        default = Ellipsis if required else None
                    desc = arg.get("description") or f"One of: {', '.join(map(str, lit_values))}"
                    fields[arg_name] = (py_type, Field(default, description=desc))
                    continue

                py_type = TYPE_MAP.get(str(arg.get("type", "string")).lower(), str)
                required = bool(arg.get("required", True))
                field_type = py_type if required else Optional[py_type]  # type: ignore[index]
                default = arg["default"] if "default" in arg else (Ellipsis if required else None)
                desc = arg.get("description")
                fields[arg_name] = (field_type, Field(default, description=desc) if desc is not None else default)

            model = create_model(f"{name}_Args", **fields)  # type: ignore
            globals()[model.__name__] = model  # expose unique model name globally for type resolution

            # Mustache renderer (partials relative to template directory)
            renderer = pystache.Renderer(search_dirs=[str(template_path.parent)], file_encoding="utf-8")

            def make_tool_fn(_model=model, _html=html_source, _renderer=renderer, _name=name):
                def tool_impl(data):  # annotation applied after definition to avoid forward-ref issues
                    payload = data.model_dump()
                    safe_payload = {k: ("" if v is None else v) for k, v in payload.items()}

                    # Derived promo_code_block
                    if "promo_code" in safe_payload and "promo_code_block" not in safe_payload:
                        promo_val = safe_payload.get("promo_code")
                        safe_payload["promo_code_block"] = (
                            f"<div class=\"promo\">Use promo code <strong>{promo_val}</strong>.</div>" if promo_val else ""
                        )

                    try:
                        html_rendered = _renderer.render(_html, safe_payload)
                    except Exception as e:  # pragma: no cover
                        return f"Error rendering template {_name}: {e}"

                    msg = EmailMessage()
                    subject = str(safe_payload.get("subject", ""))
                    if subject:
                        msg["Subject"] = subject
                    for hdr in ("To", "Cc", "Bcc"):
                        key = hdr.lower()
                        val = safe_payload.get(key)
                        if isinstance(val, list) and val:
                            msg[hdr] = ", ".join(val)
                        elif isinstance(val, str) and val:
                            msg[hdr] = val
                    msg['X-Unsent'] = '1'
                    msg.set_content("This is an HTML email. Your client should display the HTML part.")
                    msg.add_alternative(html_rendered, subtype="html")

                    buffer = io.BytesIO()
                    try:
                        buffer.write(msg.as_bytes())
                        buffer.seek(0)
                        upload_result = upload_file(buffer, "eml")
                        print(f"[dynamic-email] Generated email draft via '{_name}'")
                        return upload_result
                    except Exception as e:  # pragma: no cover
                        return f"Error creating email draft for template '{_name}': {e}"
                    finally:
                        buffer.close()

                # Apply annotations explicitly (FastMCP will read them after registration)
                tool_impl.__annotations__['data'] = _model  # type: ignore[index]
                tool_impl.__annotations__['return'] = str  # type: ignore[index]
                return tool_impl

            mcp.tool(name=name, description=description, annotations=annotations, meta=meta)(make_tool_fn())
            print(f"[dynamic-email] Registered tool: {name}")
        except Exception as e:  # pragma: no cover
            print(f"[dynamic-email] Failed to register template spec: {e}")
