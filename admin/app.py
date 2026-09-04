"""FastHTML admin UI for managing dynamic docx / email / pptx templates.

Mounted in the same ASGI process as the FastMCP server (see
:func:`build_combined_app`) so that saving a template takes effect immediately —
no restart. The UI is gated by a single shared password (:mod:`admin.auth`) and
persists templates through :class:`admin.store.TemplateStore`.

Authoring model (chosen with the maintainer): the user uploads a real ``.docx``
(or email ``.html``); the UI auto-detects placeholders/conditionals
(:mod:`admin.analysis`), pre-builds the argument form, and previews
(:mod:`admin.preview`) without ever hitting the upload backend.

PowerPoint is the odd one out and the views branch on it in a few places. A
docx or email template is a parameterised document that becomes an MCP tool of
its own, so it has arguments to declare and "live" means the tool is
registered. A pptx template is a *design* — layouts, theme, aspect — with no
arguments at all; it becomes one more value for the ``template`` argument of
the single ``create_powerpoint_presentation`` tool, and "live" means the
template registry has re-read it. Its form offers layout roles and deck
defaults where the others offer arguments, and its preview builds a fixed
sample deck instead of substituting values.
"""
from __future__ import annotations

import hashlib
import json
import logging
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

import metrics

from fasthtml.common import (
    FastHTML, Title, Main, Header, Div, P, A, H1, H2, H3, Form, Input, Button,
    Label, Span, Table, Thead, Tbody, Tr, Th, Td, Select, Option, Textarea,
    Details, Summary, Meta, Style, Script, NotStr, to_xml,
    RedirectResponse, Response, HTMLResponse, Hidden,
)
from starlette.routing import Mount

from config import Config
from admin import auth
from admin.analysis import analyze, reconcile, Analysis, PptxAnalysis
from admin.preview import (
    sample_values, render_docx_preview, render_email_preview, render_pptx_preview,
)
from admin.store import (
    FileTemplateStore, KIND_DOCX, KIND_EMAIL, KIND_PPTX,
    TemplateStoreError, kind_meta, validate_name,
)
from template_registry import gather_specs

logger = logging.getLogger(__name__)

KINDS = (KIND_DOCX, KIND_EMAIL, KIND_PPTX)
_KIND_LABEL = {KIND_DOCX: "Word", KIND_EMAIL: "Email", KIND_PPTX: "PowerPoint"}
_KIND_ICON = {KIND_DOCX: "📝", KIND_EMAIL: "✉️", KIND_PPTX: "📊"}
_ARG_TYPES = ["string", "int", "float", "bool", "list"]
# Style-mapping keys surfaced in the UI (subset of the full recognised set).
_STYLE_KEYS = ["heading_1", "list_number", "list_bullet", "quote", "table"]

# Reject uploads larger than this (read fully into memory before validation).
MAX_UPLOAD_BYTES = 10 * 1024 * 1024  # 10 MB


def _csrf_input(token: str):
    return Hidden(name="csrf", value=token or "")

# Self-contained theme — no CDN dependency, so the UI looks right offline / in
# locked-down deployments (FastHTML's default pico.css is CDN-loaded).
ADMIN_CSS = """
:root{
  --bg:#f1f5f9; --card:#ffffff; --ink:#0f172a; --muted:#64748b; --line:#e2e8f0;
  --brand:#4f46e5; --brand-d:#4338ca; --ok:#16a34a; --warn:#b45309; --err:#dc2626;
  --ok-bg:#ecfdf5; --warn-bg:#fffbeb; --err-bg:#fef2f2; --info-bg:#eff6ff;
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
  font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
  font-size:15px;line-height:1.5}
a{color:var(--brand);text-decoration:none}
a:hover{text-decoration:underline}
.topbar{background:var(--brand);color:#fff;padding:.75rem 1.25rem;
  display:flex;align-items:center;justify-content:space-between;gap:1rem;
  box-shadow:0 1px 3px rgba(0,0,0,.15)}
.topbar .brand{font-weight:700;font-size:1.05rem;color:#fff;display:flex;align-items:center;gap:.5rem}
.topbar nav a{color:#e0e7ff;margin-left:1rem;font-size:.92rem}
.topbar nav a:hover{color:#fff}
.container{max-width:980px;margin:1.5rem auto;padding:0 1.25rem}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;
  padding:1.25rem 1.5rem;margin-bottom:1.25rem;box-shadow:0 1px 2px rgba(15,23,42,.04)}
h1{font-size:1.5rem;margin:.2rem 0 1rem}
h2{font-size:1.15rem;margin:0 0 .75rem}
h3{font-size:1rem;margin:1.25rem 0 .5rem}
.muted{color:var(--muted);font-size:.9rem}
.field{margin-bottom:1rem}
.field label{display:block;font-weight:600;font-size:.85rem;margin-bottom:.3rem}
input,select,textarea{width:100%;padding:.5rem .6rem;border:1px solid var(--line);
  border-radius:8px;font:inherit;background:#fff;color:var(--ink)}
input:focus,select:focus,textarea:focus{outline:2px solid var(--brand);
  outline-offset:-1px;border-color:var(--brand)}
input[disabled]{background:#f8fafc;color:var(--muted)}
textarea{resize:vertical}
.btn{display:inline-flex;align-items:center;gap:.4rem;cursor:pointer;border-radius:8px;
  padding:.5rem .9rem;font:inherit;font-weight:600;border:1px solid transparent;
  background:#fff;color:var(--ink);text-decoration:none}
.btn:hover{text-decoration:none}
.btn-primary{background:var(--brand);color:#fff}
.btn-primary:hover{background:var(--brand-d);color:#fff}
.btn-secondary{background:#fff;color:var(--brand);border-color:var(--brand)}
.btn-secondary:hover{background:#eef2ff}
.btn-danger{background:#fff;color:var(--err);border-color:#fecaca}
.btn-danger:hover{background:var(--err-bg)}
.btn-sm{padding:.3rem .55rem;font-size:.82rem}
.btn-icon{padding:.25rem .5rem;background:#fff;border:1px solid var(--line);color:var(--err)}
.btn-icon:hover{background:var(--err-bg);border-color:#fecaca}
.actions{display:flex;gap:.6rem;flex-wrap:wrap;align-items:center}
table{width:100%;border-collapse:collapse}
th,td{text-align:left;padding:.55rem .5rem;border-bottom:1px solid var(--line);vertical-align:middle}
th{font-size:.78rem;text-transform:uppercase;letter-spacing:.03em;color:var(--muted)}
.tpl-table td{font-size:.95rem}
.args-table td{padding:.3rem .35rem;border-bottom:none}
.args-table th{padding-bottom:.25rem}
.args-table input,.args-table select{padding:.4rem .5rem;font-size:.9rem}
.col-name{width:22%}.col-type{width:12%}.col-req{width:13%}.col-def{width:16%}.col-x{width:38px}
.badge{display:inline-block;padding:.12rem .5rem;border-radius:999px;font-size:.74rem;
  font-weight:600;line-height:1.5}
.badge-live{background:var(--ok-bg);color:var(--ok)}
.badge-off{background:#f1f5f9;color:var(--muted)}
.badge-ro{background:#f1f5f9;color:var(--muted)}
.badge-if{background:#eef2ff;color:var(--brand);margin-left:.4rem}
.chip{display:inline-block;padding:.15rem .55rem;margin:.15rem .25rem .15rem 0;border-radius:6px;
  background:#eef2ff;color:var(--brand-d);font-size:.82rem;font-family:ui-monospace,Menlo,monospace}
.chip-if{background:#fef9c3;color:#854d0e}
.flash{padding:.6rem .85rem;border-radius:8px;margin:.5rem 0;border:1px solid transparent}
.flash-info{background:var(--info-bg);border-color:#bfdbfe}
.flash-ok{background:var(--ok-bg);border-color:#a7f3d0}
.flash-warn{background:var(--warn-bg);border-color:#fde68a;color:#854d0e}
.flash-err{background:var(--err-bg);border-color:#fecaca;color:#991b1b}
.empty{text-align:center;padding:2rem 1rem;color:var(--muted)}
/* PowerPoint template views */
.role-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:0 1rem}
.facts .field{margin-bottom:.6rem}
.warn-text{color:var(--warn)}
.swatch-wrap{display:inline-flex;align-items:center;gap:.35rem;margin:.15rem .6rem .15rem 0}
.swatch{display:inline-block;width:14px;height:14px;border-radius:3px;border:1px solid var(--line)}
.swatch-label{font-size:.78rem;color:var(--muted);font-family:ui-monospace,Menlo,monospace}
.field label input[type=checkbox]{width:auto;margin-right:.4rem;vertical-align:-2px}
details{margin:1rem 0;border:1px solid var(--line);border-radius:8px;padding:.5rem .85rem}
summary{cursor:pointer;font-weight:600}
.login-wrap{max-width:380px;margin:8vh auto}
.inline-form{display:inline}
hr{border:none;border-top:1px solid var(--line);margin:1.25rem 0}
.stats{display:flex;gap:1rem;flex-wrap:wrap;margin-bottom:1rem}
.stat{flex:1;min-width:150px;background:var(--card);border:1px solid var(--line);
  border-radius:10px;padding:.85rem 1rem}
.stat .num{font-size:1.55rem;font-weight:700;line-height:1.2}
.stat .lbl{font-size:.78rem;text-transform:uppercase;letter-spacing:.03em;color:var(--muted)}
.logs{font-family:ui-monospace,Menlo,Consolas,monospace;font-size:.82rem;
  max-height:460px;overflow:auto;border:1px solid var(--line);border-radius:8px}
.logs table{width:100%}
.logs td{padding:.25rem .5rem;border-bottom:1px solid #f1f5f9;white-space:nowrap}
.logs td.msg{white-space:normal;word-break:break-word}
.logs .ts{color:var(--muted)}
.lvl{font-weight:700}
.lvl-ERROR,.lvl-CRITICAL{color:var(--err)}
.lvl-WARNING{color:var(--warn)}
.lvl-INFO{color:#2563eb}
.lvl-DEBUG{color:var(--muted)}
.toggle-row{display:flex;gap:.5rem;align-items:center;margin-bottom:.6rem}
.num-err{color:var(--err)}
"""

# Vanilla JS for dynamic argument rows (add / remove) — avoids a CDN htmx dep.
ARG_ROWS_JS = """
function adminAddArgRow(){
  var tb=document.getElementById('argrows');
  if(!tb) return;
  tb.insertAdjacentHTML('beforeend', window.__ARG_ROW_HTML__);
}
function adminRemoveRow(btn){
  var tr=btn.closest('tr'); if(tr) tr.remove();
}
"""


class AdminContext:
    """Shared services the views depend on."""

    def __init__(self, mcp, config: Config):
        self.mcp = mcp
        self.config = config
        self.path = config.admin.path.rstrip("/")
        self.store = FileTemplateStore.from_config()

    def u(self, path: str = "") -> str:
        """Absolute (mount-prefixed) URL for an admin-relative *path*."""
        return f"{self.path}{path}"

    @property
    def global_style_mapping(self) -> Dict[str, Any]:
        """Master-YAML ``style_mapping``, re-read each access so edits on disk
        take effect without a restart (the admin UI is about live editing)."""
        master = self.store.config_dir / "docx_templates.yaml"
        _templates, cfg = gather_specs(master, None)
        return cfg.get("style_mapping") or {}

    # -- live MCP tool registration ----------------------------------------

    # A docx or email template *is* an MCP tool, so saving one registers a tool.
    # A PowerPoint template is not: there is a single
    # ``create_powerpoint_presentation`` tool and a template is one more value
    # for its ``template`` argument. "Going live" therefore means the registry
    # re-reads the spec directory, which is what clear_cache forces. The
    # registry already notices the change on its own via a directory
    # fingerprint; dropping the cache here just makes it immediate rather than
    # on the next call that happens to look.

    def register(self, kind: str, spec: Dict[str, Any]) -> bool:
        if kind == KIND_PPTX:
            return self._refresh_pptx_registry(spec.get("name"))
        if kind == KIND_DOCX:
            from docx_tools.dynamic_docx_tools import register_docx_template
            return register_docx_template(self.mcp, spec, self.global_style_mapping)
        from email_tools.dynamic_email_tools import register_email_template
        return register_email_template(self.mcp, spec)

    def unregister(self, kind: str, name: str) -> bool:
        if kind == KIND_PPTX:
            from pptx_tools.templates import clear_cache
            clear_cache()
            return True
        if kind == KIND_DOCX:
            from docx_tools.dynamic_docx_tools import unregister_docx_template
            return unregister_docx_template(self.mcp, name)
        from email_tools.dynamic_email_tools import unregister_email_template
        return unregister_email_template(self.mcp, name)

    def live_names(self, kind: str) -> List[str]:
        if kind == KIND_PPTX:
            from pptx_tools.templates import template_names
            try:
                return template_names()
            except Exception:
                logger.exception("[admin] Could not read the pptx template registry")
                return []
        if kind == KIND_DOCX:
            from docx_tools.dynamic_docx_tools import registered_docx_template_names
            return registered_docx_template_names()
        from email_tools.dynamic_email_tools import registered_email_template_names
        return registered_email_template_names()

    @staticmethod
    def _refresh_pptx_registry(name: Optional[str]) -> bool:
        """Reload the pptx registry and report whether *name* is now in it.

        Returning False here means the saved spec did not come back out of the
        registry — almost always because its file is missing from the template
        directories — and the admin is told so rather than being shown a
        success message for a template the tool cannot use.
        """
        from pptx_tools.templates import clear_cache, template_names
        clear_cache()
        try:
            return name in template_names()
        except Exception:
            logger.exception("[admin] Could not reload the pptx template registry")
            return False


# ---------------------------------------------------------------------------
# Form parsing helpers
# ---------------------------------------------------------------------------

def _coerce_default(atype: str, raw: str) -> Any:
    """Coerce a string default from the form into the arg's declared type."""
    raw = (raw or "").strip()
    atype = (atype or "string").lower()
    if atype in ("bool", "boolean"):
        return raw.lower() in ("1", "true", "yes", "on")
    if atype in ("int", "integer") and raw:
        try:
            return int(raw)
        except ValueError:
            return raw
    if atype == "float" and raw:
        try:
            return float(raw)
        except ValueError:
            return raw
    return raw


def _parse_args_from_form(form) -> List[Dict[str, Any]]:
    """Build the ``args`` list from the parallel-indexed form fields."""
    names = form.getlist("arg_name")
    types = form.getlist("arg_type")
    reqs = form.getlist("arg_required")
    defs = form.getlist("arg_default")
    descs = form.getlist("arg_desc")
    args: List[Dict[str, Any]] = []
    for i, raw_name in enumerate(names):
        name = (raw_name or "").strip()
        if not name:
            continue
        atype = types[i] if i < len(types) else "string"
        required = (reqs[i] if i < len(reqs) else "true") == "true"
        default_raw = defs[i] if i < len(defs) else ""
        desc = descs[i] if i < len(descs) else ""
        arg: Dict[str, Any] = {"name": name, "type": atype, "required": required, "description": desc}
        if not required or (default_raw or "").strip():
            arg["default"] = _coerce_default(atype, default_raw)
        args.append(arg)
    return args


def _parse_style_mapping(form) -> Dict[str, str]:
    """Collect non-default style-mapping selections."""
    mapping: Dict[str, str] = {}
    for key in _STYLE_KEYS:
        val = (form.get(f"style_{key}") or "").strip()
        if val and val != "__default__":
            mapping[key] = val
    return mapping


def _checked(form, field: str) -> bool:
    """True when a checkbox was submitted ticked.

    An unticked HTML checkbox sends nothing at all, so absence has to mean
    False rather than "unset".
    """
    return (form.get(field) or "").lower() in ("1", "true", "yes", "on")


def _build_pptx_spec(form) -> Dict[str, Any]:
    """Assemble a PowerPoint template spec from the submitted edit form.

    Nothing like the docx/email spec: no ``args``, because a presentation
    template takes no arguments. What it carries instead is how the deck should
    be laid out — which layout plays each role, and the deck-wide defaults.
    """
    from pptx_tools.layouts import ROLES

    name = validate_name((form.get("name") or "").strip())
    spec: Dict[str, Any] = {
        "name": name,
        "description": (form.get("description") or "").strip(),
        "pptx_path": (form.get("asset_filename") or "").strip(),
    }
    if _checked(form, "is_default"):
        spec["default"] = True
    # Only write strip_slides when it departs from the default, so a spec file
    # stays as small as what the admin actually chose.
    if not _checked(form, "strip_slides"):
        spec["strip_slides"] = False

    layouts = {
        role: (form.get(f"layout_{role}") or "").strip()
        for role in ROLES
        if (form.get(f"layout_{role}") or "").strip()
    }
    if layouts:
        spec["layouts"] = layouts

    defaults: Dict[str, Any] = {}
    footer = (form.get("default_footer_text") or "").strip()
    if footer:
        defaults["footer_text"] = footer
    language = (form.get("default_language") or "").strip()
    if language:
        defaults["language"] = language
    if _checked(form, "default_slide_numbers"):
        defaults["show_slide_numbers"] = True
    if defaults:
        spec["defaults"] = defaults
    return spec


def _build_spec(kind: str, form) -> Dict[str, Any]:
    """Assemble a template spec dict from the submitted edit form."""
    if kind == KIND_PPTX:
        return _build_pptx_spec(form)
    name = validate_name((form.get("name") or "").strip())
    asset_filename = (form.get("asset_filename") or "").strip()
    description = (form.get("description") or "").strip()
    title = (form.get("title") or "").strip()

    spec: Dict[str, Any] = {"name": name, "description": description or f"Generate {name}"}
    if title:
        spec["annotations"] = {"title": title}
    spec[_path_key(kind)] = asset_filename
    spec["args"] = _parse_args_from_form(form)
    if kind == KIND_DOCX:
        style_mapping = _parse_style_mapping(form)
        if style_mapping:
            spec["style_mapping"] = style_mapping
    return spec


def _path_key(kind: str) -> str:
    return kind_meta(kind)["path_key"]


def _asset_ext(kind: str) -> str:
    """The canonical extension, used when deriving a filename from the name."""
    return kind_meta(kind)["asset_ext"]


def _accept_exts(kind: str) -> str:
    """Every extension the upload field accepts, as an HTML accept list."""
    return ",".join(kind_meta(kind)["asset_exts"])


# ---------------------------------------------------------------------------
# View fragments
# ---------------------------------------------------------------------------

def _topbar(ctx: AdminContext, authed: bool = True):
    from fasthtml.common import Nav
    nav = Nav(
        A("All templates", href=ctx.u("/")),
        A("New Word", href=ctx.u("/new/docx")),
        A("New Email", href=ctx.u("/new/email")),
        A("Status", href=ctx.u("/status")),
        A("Log out", href=ctx.u("/logout")),
    ) if authed else Span()
    return Header(Span("📄 Template Admin", cls="brand"), nav, cls="topbar")


def _page(ctx: AdminContext, title: str, *content, authed: bool = True):
    """Wrap page *content* in the themed shell (title + topbar + container)."""
    return (
        Title(f"{title} · Template Admin"),
        _topbar(ctx, authed),
        Main(Div(*content, cls="container")),
    )


def _flash(msg: Optional[str], kind: str = "info"):
    if not msg:
        return None
    return Div(msg, cls=f"flash flash-{kind}")


def _status_badge(live: bool):
    return Span("● Live", cls="badge badge-live") if live else Span("Not live", cls="badge badge-off")


def _template_table(ctx: AdminContext, kind: str, csrf: str = ""):
    managed = ctx.store.list_specs(kind)
    managed_names = {s.get("name") for s in managed}
    live = set(ctx.live_names(kind))

    # "Args" is meaningless for a presentation template, which declares none;
    # what distinguishes one there is whether it is the default.
    detail_header = "Default" if kind == KIND_PPTX else "Args"

    rows = []
    for spec in managed:
        name = spec.get("name")
        if kind == KIND_PPTX:
            detail = "★ default" if spec.get("default") else "—"
        else:
            detail = str(len(spec.get("args") or []))
        rows.append(Tr(
            Td(A(name, href=ctx.u(f"/{kind}/{name}/edit"))),
            Td(detail),
            Td(_status_badge(name in live)),
            Td(Div(
                A("Edit", href=ctx.u(f"/{kind}/{name}/edit"), cls="btn btn-secondary btn-sm"),
                Form(_csrf_input(csrf),
                     Button("Delete", type="submit", cls="btn btn-danger btn-sm",
                            onclick="return confirm('Delete this template? This removes the live tool. "
                                    "The source file is kept in custom_templates/.')"),
                     action=ctx.u(f"/{kind}/{name}/delete"), method="post", cls="inline-form"),
                cls="actions",
            )),
        ))
    # Read-only templates from the master YAML (live but not UI-managed).
    for name in sorted(live - managed_names):
        rows.append(Tr(
            Td(name), Td("—"),
            Td(Span("● Live", cls="badge badge-live")),
            Td(Span("from master YAML — read-only", cls="badge badge-ro")),
        ))

    if not rows:
        return Div(
            P(f"No {_KIND_LABEL[kind]} templates yet."),
            A(f"{_KIND_ICON[kind]} Create your first {_KIND_LABEL[kind]} template",
              href=ctx.u(f"/new/{kind}"), cls="btn btn-primary"),
            cls="empty",
        )
    return Table(
        Thead(Tr(Th("Name"), Th(detail_header), Th("Status"), Th("Actions"))),
        Tbody(*rows),
        cls="tpl-table",
    )


def _arg_row(arg: Dict[str, Any] = None, cond: bool = False):
    """Render one editable argument row (parallel-indexed fields + remove button)."""
    arg = arg or {}
    name = arg.get("name", "")
    atype = str(arg.get("type", "string")).lower()
    required = bool(arg.get("required", True))
    default = arg.get("default", "")
    desc = arg.get("description", "")

    type_opts = [Option(t, value=t, selected=(t == atype)) for t in _ARG_TYPES]
    req_opts = [
        Option("required", value="true", selected=required),
        Option("optional", value="false", selected=not required),
    ]
    name_cell = [Input(name="arg_name", value=name, placeholder="argument_name")]
    if cond:
        name_cell.append(Span("if", cls="badge badge-if", title="Conditional flag"))
    return Tr(
        Td(*name_cell, cls="col-name"),
        Td(Select(*type_opts, name="arg_type"), cls="col-type"),
        Td(Select(*req_opts, name="arg_required"), cls="col-req"),
        Td(Input(name="arg_default", value="" if default in (None,) else str(default),
                 placeholder="optional"), cls="col-def"),
        Td(Input(name="arg_desc", value=desc, placeholder="what this is, for the AI")),
        Td(Button("✕", type="button", cls="btn-icon", title="Remove",
                  onclick="adminRemoveRow(this)"), cls="col-x"),
        cls="arg-row",
    )


def _chips(items, cond_set=None):
    cond_set = cond_set or set()
    if not items:
        return Span("none", cls="muted")
    return Span(*[Span(it, cls="chip chip-if" if it in cond_set else "chip") for it in items])


def _style_mapping_block(analysis: Optional[Analysis], spec: Dict[str, Any]):
    styles = (analysis.styles_present if analysis else []) or []
    current = (spec or {}).get("style_mapping") or {}
    selects = []
    for key in _STYLE_KEYS:
        cur = current.get(key, "")
        opts = [Option("(use built-in)", value="__default__", selected=not cur)]
        opts += [Option(s, value=s, selected=(s == cur)) for s in styles]
        selects.append(Div(Label(key), Select(*opts, name=f"style_{key}"), cls="field"))
    return Details(
        Summary("Advanced — map markdown styles to your template's own style names"),
        P("Only needed if your Word template renames the built-in styles "
          "(Heading 1, List Number, …). Otherwise leave these alone.", cls="muted"),
        *selects,
    )


def _analysis_report(analysis, spec: Dict[str, Any]):
    if isinstance(analysis, PptxAnalysis):
        return _pptx_analysis_report(analysis)
    rec = reconcile(analysis, spec.get("args") or [])
    cond_set = set(analysis.conditionals)
    items = [
        Div(Label("Detected placeholders"), _chips(analysis.placeholders), cls="field"),
        Div(Label("Detected conditionals"),
            _chips(analysis.conditionals, cond_set) if analysis.conditionals else Span("none", cls="muted"),
            cls="field"),
    ]
    if analysis.missing_required_styles:
        items.append(_flash(
            "⚠ The document is missing some Word styles the renderer uses: "
            + ", ".join(analysis.missing_required_styles)
            + ". Lists/headings using them will fall back to the default style.", "warn"))
    for w in analysis.warnings:
        items.append(_flash("⚠ " + w, "warn"))
    if rec.orphan_args:
        items.append(_flash("Args with no matching placeholder (they'll be ignored by the document): "
                            + ", ".join(rec.orphan_args), "warn"))
    if rec.non_bool_conditions:
        items.append(_flash("Conditional flags should be type 'bool': "
                            + ", ".join(rec.non_bool_conditions), "warn"))
    return Div(H3("What we found in the document"),
               *[i for i in items if i is not None], cls="card")


def _swatch(name: str, value: str):
    """One theme colour, shown rather than described."""
    return Span(
        Span(cls="swatch", style=f"background:{value}"),
        Span(f"{name} {value}", cls="swatch-label"),
        cls="swatch-wrap", title=f"{name}: {value}",
    )


def _pptx_analysis_report(analysis: PptxAnalysis):
    """What we found in a PowerPoint template: size, layouts, roles, theme."""
    rows = []
    for layout in analysis.layouts:
        rows.append(Tr(
            Td(layout.name),
            Td(Span(layout.role, cls="chip") if layout.role
               else Span("name it explicitly", cls="muted")),
            Td(_chips(layout.placeholders)),
            Td("—" if layout.has_footer else Span("no footer", cls="muted")),
        ))

    facts = Div(
        Div(Label("Slide size"),
            Span(f"{analysis.slide_size} ({analysis.aspect})" if analysis.slide_size
                 else "unknown"), cls="field"),
        Div(Label("Theme fonts"),
            Span(", ".join(f"{k}: {v}" for k, v in analysis.theme_fonts.items())
                 or "not declared", cls="muted" if not analysis.theme_fonts else ""),
            cls="field"),
        Div(Label("Theme colours"),
            Span(*[_swatch(k, v) for k, v in analysis.theme_colors.items()])
            if analysis.theme_colors else Span("not declared", cls="muted"),
            cls="field"),
        cls="facts",
    )

    items = [H3("What we found in the template"), facts]
    for w in analysis.warnings:
        items.append(_flash("⚠ " + w, "warn"))
    if rows:
        items.append(Table(
            Thead(Tr(Th("Layout"), Th("Detected role"), Th("Placeholders"), Th("Footer"))),
            Tbody(*rows), cls="tpl-table",
        ))
    return Div(*items, cls="card")


def _pptx_edit_form(ctx: AdminContext, spec: Dict[str, Any],
                    analysis: Optional[PptxAnalysis], is_new: bool, csrf: str = ""):
    """The PowerPoint template form.

    Deliberately not the argument editor. A presentation template declares no
    arguments; what an admin needs to control is which layout plays each slide
    role, and the deck-wide defaults.
    """
    from pptx_tools.layouts import ROLES

    kind = KIND_PPTX
    asset_filename = spec.get("pptx_path", "")
    configured = spec.get("layouts") or {}
    defaults = spec.get("defaults") or {}
    detected = analysis.role_map if analysis else {}
    layout_names = analysis.layout_names if analysis else []

    name_field = (
        Div(Label("Template name"),
            Input(name="name", value=spec.get("name", ""), required=True,
                  placeholder="e.g. corporate_16_9"),
            P("Letters, digits and underscores. This is the value passed as "
              "the 'template' argument when generating a deck.", cls="muted"),
            cls="field")
        if is_new else
        Div(Label("Template name"), Input(value=spec.get("name", ""), disabled=True), cls="field")
    )

    details_card = Div(
        H3("Template details"),
        name_field,
        Div(Label("Description"),
            Textarea(spec.get("description", ""), name="description", rows="3",
                     placeholder="When should the AI reach for this deck? "
                                 "e.g. 'Brand deck, widescreen. Use for client-facing work.'"),
            P("This is what the AI sees when choosing between templates.", cls="muted"),
            cls="field"),
        Div(Label(Input(name="is_default", type="checkbox", value="1",
                        checked=bool(spec.get("default"))),
                  " Use this template when none is named"), cls="field"),
        Div(Label(Input(name="strip_slides", type="checkbox", value="1",
                        checked=bool(spec.get("strip_slides", True))),
                  " Drop any slides the template file itself contains"),
            P("Sample slides a designer left in the file never belong in generated "
              "output. Leave this on unless you know otherwise.", cls="muted"),
            cls="field"),
        cls="card",
    )

    # One dropdown per role, populated from the template's real layout names.
    role_fields = []
    for role in ROLES:
        current = configured.get(role, "")
        auto = detected.get(role)
        auto_label = f"auto-detected: {auto}" if auto else "not detected"
        opts = [Option(f"({auto_label})", value="", selected=not current)]
        opts += [Option(n, value=n, selected=(n == current)) for n in layout_names]
        role_fields.append(Div(
            Label(role),
            Select(*opts, name=f"layout_{role}"),
            None if auto else P("No layout matches this role — slides needing it fall "
                                "back by position unless you pick one.",
                                cls="muted warn-text"),
            cls="field",
        ))

    layouts_card = Div(
        H3("Layout for each slide role"),
        P("Layouts are matched automatically by their placeholders, so most templates "
          "need nothing here. Override a role only when the automatic choice is wrong, "
          "or when it says 'not detected'.", cls="muted"),
        Div(*role_fields, cls="role-grid"),
        cls="card",
    )

    defaults_card = Div(
        H3("Deck defaults"),
        P("Applied when the tool call does not set the same option.", cls="muted"),
        Div(Label("Footer text"),
            Input(name="default_footer_text", value=defaults.get("footer_text", "") or "",
                  placeholder="e.g. ACME s.r.o. · Confidential"),
            cls="field"),
        Div(Label("Language"),
            Input(name="default_language", value=defaults.get("language", "") or "",
                  placeholder="e.g. cs-CZ"),
            P("BCP-47 tag stamped on every text run, so the deck is proof-read in "
              "the right language.", cls="muted"),
            cls="field"),
        Div(Label(Input(name="default_slide_numbers", type="checkbox", value="1",
                        checked=bool(defaults.get("show_slide_numbers"))),
                  " Show slide numbers"), cls="field"),
        cls="card",
    )

    return Form(
        _csrf_input(csrf),
        Hidden(name="kind", value=kind),
        Hidden(name="asset_filename", value=asset_filename),
        Hidden(name="name", value=spec.get("name", "")) if not is_new else None,
        details_card,
        layouts_card,
        defaults_card,
        Div(
            Div(
                Button("Save & make live", type="submit", cls="btn btn-primary"),
                Button("Preview sample deck", type="submit",
                       formaction=ctx.u(f"/{kind}/preview"),
                       formtarget="_blank", cls="btn btn-secondary"),
                A("Cancel", href=ctx.u("/"), cls="btn"),
                cls="actions",
            ),
            P("Preview builds a six-slide sample deck on this template, using the "
              "layout choices above — including ones you have not saved yet.",
              cls="muted"),
            cls="card",
        ),
        action=ctx.u(f"/{kind}/save"), method="post",
    )


def _edit_form(ctx: AdminContext, kind: str, spec: Dict[str, Any],
               analysis: Optional[Analysis], is_new: bool, csrf: str = ""):
    if kind == KIND_PPTX:
        return _pptx_edit_form(ctx, spec, analysis, is_new, csrf)
    asset_filename = spec.get(_path_key(kind), "")
    annotations = spec.get("annotations") or {}
    cond_set = set(analysis.conditionals) if analysis else set()

    # Build arg rows: declared args first, then detected-but-missing. No blank
    # spares — users add rows with the "Add argument" button.
    existing = spec.get("args") or []
    existing_names = {a.get("name") for a in existing if isinstance(a, dict)}
    rows = [_arg_row(a, cond=a.get("name") in cond_set) for a in existing]
    if analysis:
        for ph in analysis.placeholders:
            if ph not in existing_names:
                is_cond = ph in cond_set
                rows.append(_arg_row(
                    {"name": ph, "type": "bool" if is_cond else "string",
                     "required": not is_cond, "description": ""}, cond=is_cond))
                existing_names.add(ph)
        for c in analysis.conditionals:
            if c not in existing_names:
                rows.append(_arg_row({"name": c, "type": "bool", "required": False,
                                      "description": ""}, cond=True))
                existing_names.add(c)
    if not rows:
        rows = [_arg_row()]

    name_field = (
        Div(Label("Tool name"),
            Input(name="name", value=spec.get("name", ""), required=True,
                  placeholder="e.g. formal_letter"),
            P("Letters, digits and underscores. This is what the AI calls.", cls="muted"),
            cls="field")
        if is_new else
        Div(Label("Tool name"), Input(value=spec.get("name", ""), disabled=True), cls="field")
    )

    details_card = Div(
        H3("Tool details"),
        name_field,
        Div(Label("Title"), Input(name="title", value=annotations.get("title", ""),
            placeholder="Friendly name shown to the user"), cls="field"),
        Div(Label("Description"),
            Textarea(spec.get("description", ""), name="description", rows="3",
                     placeholder="Tell the AI when and how to use this template."),
            P("This is the instruction the AI sees when choosing the tool.", cls="muted"),
            cls="field"),
        cls="card",
    )

    args_card = Div(
        H3("Arguments"),
        P("Every placeholder and conditional in your document needs an argument. "
          "We pre-filled them from the document — adjust types and descriptions, then save.",
          cls="muted"),
        Table(
            Thead(Tr(Th("Name", cls="col-name"), Th("Type", cls="col-type"),
                     Th("Required", cls="col-req"), Th("Default", cls="col-def"),
                     Th("Description"), Th("", cls="col-x"))),
            Tbody(*rows, id="argrows"),
            cls="args-table",
        ),
        Div(Button("+ Add argument", type="button", cls="btn btn-secondary btn-sm",
                   onclick="adminAddArgRow()"), style="margin-top:.6rem"),
        cls="card",
    )

    return Form(
        _csrf_input(csrf),
        Hidden(name="kind", value=kind),
        Hidden(name="asset_filename", value=asset_filename),
        Hidden(name="name", value=spec.get("name", "")) if not is_new else None,
        details_card,
        args_card,
        Div(_style_mapping_block(analysis, spec), cls="card") if kind == KIND_DOCX else None,
        Div(
            Div(
                Button("Save & make live", type="submit", cls="btn btn-primary"),
                Button("Preview", type="submit", formaction=ctx.u(f"/{kind}/preview"),
                       formtarget="_blank", cls="btn btn-secondary"),
                A("Cancel", href=ctx.u("/"), cls="btn"),
                cls="actions",
            ),
            cls="card",
        ),
        action=ctx.u(f"/{kind}/save"), method="post",
    )


# ---------------------------------------------------------------------------
# App factory
# ---------------------------------------------------------------------------

def _fmt_ts(ts: Optional[float]) -> str:
    if not ts:
        return "—"
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(ts))


def _fmt_uptime(seconds: float) -> str:
    seconds = int(max(0, seconds))
    d, rem = divmod(seconds, 86400)
    h, rem = divmod(rem, 3600)
    m, _ = divmod(rem, 60)
    if d:
        return f"{d}d {h}h {m}m"
    if h:
        return f"{h}h {m}m"
    return f"{m}m"


def _stat(label: str, value, num_cls: str = ""):
    return Div(Div(str(value), cls=f"num {num_cls}".strip()), Div(label, cls="lbl"), cls="stat")


def _replace_card(ctx: AdminContext, kind: str, name: str, csrf: str = ""):
    return Div(
        H3("Replace document"),
        P("Upload a new version of the source file. We'll re-scan it for placeholders "
          "and keep the arguments you've already configured.", cls="muted"),
        Form(
            _csrf_input(csrf),
            Div(Input(name="file", type="file", accept=_asset_ext(kind), required=True), cls="field"),
            Button("Upload & re-scan", type="submit", cls="btn btn-secondary"),
            action=ctx.u(f"/{kind}/{name}/reupload"), method="post", enctype="multipart/form-data",
        ),
        cls="card",
    )


def _status_view(ctx: AdminContext, level: str = "info"):
    live_docx = ctx.live_names(KIND_DOCX)
    live_email = ctx.live_names(KIND_EMAIL)
    lvl_counts = metrics.counts_by_level()
    err_count = lvl_counts.get("ERROR", 0) + lvl_counts.get("CRITICAL", 0)

    stats = Div(
        _stat("Uptime", _fmt_uptime(time.time() - metrics.START_TIME)),
        _stat("Upload backend", ctx.config.storage.strategy.value),
        _stat("Live Word tools", len(live_docx)),
        _stat("Live Email tools", len(live_email)),
        _stat("Errors logged", err_count, num_cls="num-err" if err_count else ""),
        cls="stats",
    )

    # Per-template usage.
    rows = []
    for st in metrics.tool_stats():
        rows.append(Tr(
            Td(st.name), Td(st.kind), Td(str(st.calls)),
            Td(str(st.errors), cls="num-err" if st.errors else ""),
            Td(_fmt_ts(st.last_called)),
            Td(st.last_error or "—", cls="msg"),
        ))
    usage = (Table(
        Thead(Tr(Th("Tool"), Th("Kind"), Th("Calls"), Th("Errors"),
                 Th("Last used"), Th("Last error"))),
        Tbody(*rows),
        cls="tpl-table",
    ) if rows else P("No template tools have been called yet this session.", cls="muted"))

    # Recent logs, optionally filtered to errors+warnings.
    errors_only = level == "error"
    min_level = logging.WARNING if errors_only else logging.INFO
    log_rows = [
        Tr(
            Td(_fmt_ts(r["time"]), cls="ts"),
            Td(r["level"], cls=f"lvl lvl-{r['level']}"),
            Td(r["logger"]),
            Td(r["message"], cls="msg"),
        )
        for r in metrics.recent_logs(min_level, limit=150)
    ]
    log_block = (Div(Table(Tbody(*log_rows)), cls="logs") if log_rows
                 else P("No log records captured yet.", cls="muted"))
    toggle = Div(
        Span("Show:", cls="muted"),
        A("All", href=ctx.u("/status"),
          cls="btn btn-sm " + ("btn-secondary" if errors_only else "btn-primary")),
        A("Warnings & errors", href=ctx.u("/status?level=error"),
          cls="btn btn-sm " + ("btn-primary" if errors_only else "btn-secondary")),
        A("Refresh", href=ctx.u(f"/status{'?level=error' if errors_only else ''}"), cls="btn btn-sm"),
        cls="toggle-row",
    )

    return _page(
        ctx, "Status",
        H1("Status"),
        stats,
        Div(H2("Template usage (this session)"), usage, cls="card"),
        Div(H2("Recent activity & errors"), toggle, log_block, cls="card"),
    )


def _new_page(ctx: AdminContext, kind: str, csrf: str = "", error: str = None):
    return _page(
        ctx, f"New {_KIND_LABEL[kind]} template",
        H1(f"{_KIND_ICON[kind]} New {_KIND_LABEL[kind]} template"),
        _flash(error, "err"),
        Div(
            Form(
                _csrf_input(csrf),
                Div(Label("Template name" if kind == KIND_PPTX else "Tool name"),
                    Input(name="name", required=True,
                          placeholder="e.g. corporate_16_9" if kind == KIND_PPTX
                          else "e.g. formal_letter"),
                    P("Letters, digits and underscores — this is the value passed as the "
                      "'template' argument when generating a deck." if kind == KIND_PPTX
                      else "Letters, digits and underscores — this is what the AI calls.",
                      cls="muted"),
                    cls="field"),
                Div(Label(f"{_KIND_LABEL[kind]} file ({_accept_exts(kind)})"),
                    Input(name="file", type="file", accept=_accept_exts(kind), required=True),
                    cls="field"),
                Button("Upload & analyze", type="submit", cls="btn btn-primary"),
                action=ctx.u(f"/{kind}/draft"), method="post", enctype="multipart/form-data",
            ),
            cls="card",
        ),
        _pptx_help() if kind == KIND_PPTX else Details(
            Summary("How does this work?"),
            P("Author your document in ",
              "Word" if kind == KIND_DOCX else "any HTML editor",
              " and insert placeholders like ", Span("{{recipient_name}}", cls="chip"),
              " wherever a value should go."),
            P("Wrap optional content in ", Span("{{#if include_clause}}", cls="chip"),
              " … ", Span("{{/if}}", cls="chip"), " to show it only when a flag is set.",
              cls="muted") if kind == KIND_DOCX else None,
            P("When you upload, we scan the file, list every placeholder, and pre-build the "
              "argument form for you. Nothing goes live until you press Save.", cls="muted"),
        ),
    )


def _pptx_help():
    """How a presentation template differs from the other two kinds."""
    return Details(
        Summary("How does this work?"),
        P("A PowerPoint template is a ", Span("design", cls="chip"),
          ", not a fill-in-the-blanks document. It has no placeholders and takes no "
          "arguments: what it contributes is the slide master, the layouts, the theme "
          "fonts and the colour palette."),
        P("So this does not become a tool of its own, the way a Word template does. It "
          "becomes one more choice for the ", Span("template", cls="chip"),
          " argument of the presentation tool, alongside any others you register."),
        P("Design your deck in PowerPoint and save it as .pptx or .potx. Every layout you "
          "want used should carry the placeholders it needs — a title, a content box, a "
          "picture frame — and we work out which slide role each layout plays from those. "
          "Sample slides left in the file are dropped from generated decks.", cls="muted"),
        P("When you upload, we report the slide size, every layout with the role we detected "
          "for it, and the theme. Nothing goes live until you press Save.", cls="muted"),
    )


def build_admin_app(mcp, config: Config) -> FastHTML:
    ctx = AdminContext(mcp, config)
    expected_pw = config.admin_password_effective
    login_path = ctx.u("/login")

    # Capture recent logs for the Status page (only when the admin UI is on).
    metrics.install_log_capture(level=config.logging.level_no)

    # Stable session secret derived from the password so cookies survive
    # restarts; falls back to a constant when no password is configured (the
    # login can't succeed in that case anyway).
    secret = hashlib.sha256(f"mcp-office-admin:{expected_pw or ''}".encode()).hexdigest()

    # Self-contained headers (no CDN): meta + theme CSS + the row JS, with the
    # blank-row HTML injected so "Add argument" can clone it client-side.
    blank_row_html = json.dumps(to_xml(_arg_row()))
    hdrs = (
        Meta(charset="utf-8"),
        Meta(name="viewport", content="width=device-width, initial-scale=1"),
        Style(ADMIN_CSS),
        Script(NotStr(f"window.__ARG_ROW_HTML__ = {blank_row_html};\n{ARG_ROWS_JS}")),
    )

    app = FastHTML(secret_key=secret, before=auth.make_before(login_path),
                   default_hdrs=False, htmx=False, surreal=False, hdrs=hdrs)
    rt = app.route

    def _csrf_guard(sess, form):
        """Return a 403 response when the form's CSRF token is invalid, else None."""
        if not auth.valid_csrf(sess, form.get("csrf")):
            logger.warning("[admin] Rejected POST with invalid CSRF token")
            return Response("CSRF validation failed — reload the page and try again.",
                            status_code=403)
        return None

    @rt("/login", methods=["get", "post"])
    async def login(req, sess):
        error = None
        if req.method == "POST":
            form = await req.form()
            if auth.check_password(form.get("password"), expected_pw):
                sess[auth.SESSION_KEY] = True
                auth.ensure_csrf(sess)
                return RedirectResponse(ctx.u("/"), status_code=303)
            client = req.client.host if req.client else "?"
            logger.warning("[admin] Failed login attempt from %s", client)
            error = "Incorrect password."
        return _page(
            ctx, "Sign in",
            Div(
                Div(
                    H1("📄 Template Admin"),
                    P("Sign in to manage document templates.", cls="muted"),
                    _flash(error, "err"),
                    Form(
                        Div(Label("Password"),
                            Input(name="password", type="password", required=True, autofocus=True),
                            cls="field"),
                        Button("Sign in", type="submit", cls="btn btn-primary"),
                        action=ctx.u("/login"), method="post",
                    ),
                    cls="card",
                ),
                cls="login-wrap",
            ),
            authed=False,
        )

    @rt("/logout", methods=["get", "post"])
    async def logout(req, sess):
        # Only a CSRF-validated POST clears the session, so a cross-origin GET
        # (e.g. <img src="/admin/logout">) can't force-logout an admin.
        if req.method == "POST":
            form = await req.form()
            bad = _csrf_guard(sess, form)
            if bad:
                return bad
            sess.pop(auth.SESSION_KEY, None)
            return RedirectResponse(login_path, status_code=303)
        csrf = auth.ensure_csrf(sess)
        return _page(
            ctx, "Sign out",
            Div(
                Div(
                    H1("Sign out?"),
                    P("You'll need your password to sign back in.", cls="muted"),
                    Form(
                        _csrf_input(csrf),
                        Div(
                            Button("Sign out", type="submit", cls="btn btn-primary"),
                            A("Cancel", href=ctx.u("/"), cls="btn"),
                            cls="actions",
                        ),
                        action=ctx.u("/logout"), method="post",
                    ),
                    cls="card",
                ),
                cls="login-wrap",
            ),
        )

    @rt("/")
    def index(sess):
        csrf = auth.ensure_csrf(sess)
        return _page(
            ctx, "Templates",
            H1("Templates"),
            P("Create reusable Word and email templates — each becomes a tool the AI can "
              "call — and register PowerPoint templates for it to build decks on.",
              cls="muted"),
            Div(H2(f"{_KIND_ICON[KIND_DOCX]} Word templates"),
                _template_table(ctx, KIND_DOCX, csrf), cls="card"),
            Div(H2(f"{_KIND_ICON[KIND_EMAIL]} Email templates"),
                _template_table(ctx, KIND_EMAIL, csrf), cls="card"),
            Div(H2(f"{_KIND_ICON[KIND_PPTX]} PowerPoint templates"),
                P("Designs the presentation tool can build on. These are not tools of "
                  "their own — they are choices for its 'template' argument.", cls="muted"),
                _template_table(ctx, KIND_PPTX, csrf), cls="card"),
        )

    @rt("/status")
    def status(level: str = "info"):
        return _status_view(ctx, level=level)

    @rt("/new/{kind}")
    def new(sess, kind: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        return _new_page(ctx, kind, csrf=auth.ensure_csrf(sess))

    @rt("/{kind}/draft", methods=["post"])
    async def draft(req, sess, kind: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        csrf = auth.ensure_csrf(sess)
        try:
            name = validate_name((form.get("name") or "").strip())
        except TemplateStoreError as e:
            return _new_page(ctx, kind, csrf=csrf, error=str(e))
        upload = form.get("file")
        data = await upload.read() if upload is not None else b""
        if not data:
            return _new_page(ctx, kind, csrf=csrf, error="Please choose a file to upload.")
        if len(data) > MAX_UPLOAD_BYTES:
            return _new_page(ctx, kind, csrf=csrf,
                             error=f"File too large (max {MAX_UPLOAD_BYTES // (1024*1024)} MB).")

        analysis = analyze(kind, data)
        if any("Could not open" in w for w in analysis.warnings):
            return _new_page(ctx, kind, csrf=csrf, error=analysis.warnings[0])

        # Keep the uploaded extension (a .potx stays a .potx) rather than
        # forcing the canonical one onto a file that is not in that format.
        suffix = Path(getattr(upload, "filename", "") or "").suffix.lower()
        if suffix not in kind_meta(kind)["asset_exts"]:
            suffix = _asset_ext(kind)
        filename = f"{name}{suffix}"
        try:
            ctx.store.write_asset(kind, filename, data)
        except TemplateStoreError as e:
            return _new_page(ctx, kind, csrf=csrf, error=str(e))
        spec: Dict[str, Any] = {"name": name, "description": "", _path_key(kind): filename}
        if kind == KIND_PPTX:
            spec["strip_slides"] = True
        else:
            spec["args"] = []
        return _page(
            ctx, f"Configure {name}",
            H1(f"Configure {name}"),
            _flash(
                f"Analyzed {filename} — check the layouts below, preview a sample deck, "
                "then save." if kind == KIND_PPTX else
                f"Analyzed {filename} — review the arguments below, preview, then save.",
                "ok"),
            _analysis_report(analysis, spec),
            _edit_form(ctx, kind, spec, analysis, is_new=False, csrf=csrf),
        )

    @rt("/{kind}/{name}/edit")
    def edit(sess, kind: str, name: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        spec = ctx.store.get_spec(kind, name)
        if spec is None:
            return _page(ctx, "Not found", H1("Not found"),
                         _flash(f"No managed template named '{name}'.", "err"),
                         A("← Back to all templates", href=ctx.u("/")))
        csrf = auth.ensure_csrf(sess)
        analysis = None
        asset = spec.get(_path_key(kind))
        if asset and ctx.store.asset_exists(kind, asset):
            analysis = analyze(kind, ctx.store.read_asset(kind, asset))
        return _page(
            ctx, f"Edit {name}",
            H1(f"Edit {name}"),
            _analysis_report(analysis, spec) if analysis else
            _flash("The template's source file is missing — arguments can still be edited.", "warn"),
            _edit_form(ctx, kind, spec, analysis, is_new=False, csrf=csrf),
            _replace_card(ctx, kind, name, csrf),
        )

    @rt("/{kind}/{name}/reupload", methods=["post"])
    async def reupload(req, sess, kind: str, name: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        spec = ctx.store.get_spec(kind, name)
        if spec is None:
            return _page(ctx, "Not found", H1("Not found"),
                         _flash(f"No managed template named '{name}'.", "err"),
                         A("← Back to all templates", href=ctx.u("/")))
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        csrf = auth.ensure_csrf(sess)
        upload = form.get("file")
        data = await upload.read() if upload is not None else b""
        analysis = analyze(kind, data) if data else None
        error = None
        if not data:
            error = "Please choose a file to upload."
        elif len(data) > MAX_UPLOAD_BYTES:
            error = f"File too large (max {MAX_UPLOAD_BYTES // (1024*1024)} MB)."
        elif analysis and any("Could not open" in w for w in analysis.warnings):
            error = analysis.warnings[0]
        if error:
            asset = spec.get(_path_key(kind))
            cur = analyze(kind, ctx.store.read_asset(kind, asset)) if asset and ctx.store.asset_exists(kind, asset) else None
            return _page(ctx, f"Edit {name}", H1(f"Edit {name}"), _flash(error, "err"),
                         _edit_form(ctx, kind, spec, cur, is_new=False, csrf=csrf),
                         _replace_card(ctx, kind, name, csrf))

        filename = spec.get(_path_key(kind)) or f"{name}{_asset_ext(kind)}"
        ctx.store.write_asset(kind, filename, data)
        return _page(
            ctx, f"Edit {name}",
            H1(f"Edit {name}"),
            _flash(f"Re-scanned {filename}. New placeholders (if any) were added below — "
                   "review and save to apply.", "ok"),
            _analysis_report(analysis, spec),
            _edit_form(ctx, kind, spec, analysis, is_new=False, csrf=csrf),
            _replace_card(ctx, kind, name, csrf),
        )

    @rt("/{kind}/save", methods=["post"])
    async def save(req, sess, kind: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        try:
            spec = _build_spec(kind, form)
            ctx.store.save_spec(kind, spec)
        except TemplateStoreError as e:
            return _page(ctx, "Save failed", H1("Save failed"), _flash(str(e), "err"),
                         A("← Back to all templates", href=ctx.u("/")))
        ok = ctx.register(kind, spec)
        if kind == KIND_PPTX:
            # No tool is created for a presentation template, so saying one is
            # "live" would be a lie; what changed is the set of templates the
            # presentation tool can build on.
            msg = (
                f"Saved — '{spec['name']}' is now one of the templates the presentation "
                "tool can build on."
                if ok else
                f"Saved, but '{spec['name']}' did not come back out of the template "
                "registry — check that its file is in custom_templates/ and see the logs."
            )
        else:
            msg = (f"Saved — the tool '{spec['name']}' is now live and ready for the AI to use."
                   if ok else
                   f"Saved, but the tool '{spec['name']}' could not be registered (check the logs).")
        return _page(
            ctx, "Saved",
            H1("✓ Saved" if ok else "Saved with a warning"),
            _flash(msg, "ok" if ok else "warn"),
            Div(A("Back to all templates", href=ctx.u("/"), cls="btn btn-primary"),
                A("Keep editing", href=ctx.u(f"/{kind}/{spec['name']}/edit"), cls="btn btn-secondary"),
                cls="actions"),
        )

    @rt("/{kind}/preview", methods=["post"])
    async def preview(req, sess, kind: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        spec = _build_spec(kind, form)
        asset = spec.get(_path_key(kind))
        if not asset or not ctx.store.asset_exists(kind, asset):
            return HTMLResponse("<p>Nothing to preview yet — save the template first.</p>",
                                status_code=400)
        if kind == KIND_PPTX:
            # Rendered from the file on disk, not from bytes: open_template needs
            # a path to fall back to its .potx rewrite.
            try:
                out, warnings = render_pptx_preview(ctx.store.asset_path(kind, asset), spec)
            except Exception as e:
                logger.exception("[admin] pptx preview failed")
                return HTMLResponse(f"<p>Could not build a preview: {e}</p>", status_code=400)
            headers = {
                "Content-Disposition": f'attachment; filename="{spec["name"]}_preview.pptx"',
            }
            if warnings:
                # Surfaced in a header because the response body is the deck
                # itself; the same warnings reach the model on a real call.
                headers["X-Preview-Warnings"] = " | ".join(warnings)[:900]
            return Response(
                content=out,
                media_type=(
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                ),
                headers=headers,
            )

        data = ctx.store.read_asset(kind, asset)
        analysis = analyze(kind, data)
        values = sample_values(spec.get("args") or [], analysis.conditionals)
        if kind == KIND_DOCX:
            out = render_docx_preview(data, spec, values, ctx.global_style_mapping)
            return Response(
                content=out,
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                headers={"Content-Disposition": f'attachment; filename="{spec["name"]}_preview.docx"'},
            )
        html = render_email_preview(data, spec, values)
        return HTMLResponse(html)

    @rt("/{kind}/{name}/delete", methods=["post"])
    async def delete(req, sess, kind: str, name: str):
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        if kind in KINDS:
            ctx.store.delete_spec(kind, name, delete_asset=False)
            ctx.unregister(kind, name)
        return RedirectResponse(ctx.u("/"), status_code=303)

    return app


# ---------------------------------------------------------------------------
# Combined ASGI app (admin + MCP in one process)
# ---------------------------------------------------------------------------

def build_combined_app(mcp, config: Config):
    """Build a single ASGI app serving the MCP endpoint and the admin UI.

    The MCP app (with its required lifespan / session manager) is mounted at the
    root; the admin UI is mounted under ``config.admin.path``. Mount order puts
    the admin prefix first so it wins over the catch-all MCP mount.
    """
    from starlette.applications import Starlette

    mcp_app = mcp.http_app(path="/mcp", stateless_http=config.stateless_http)
    admin_app = build_admin_app(mcp, config)
    routes = [
        Mount(config.admin.path, app=admin_app),
        Mount("/", app=mcp_app),
    ]
    return Starlette(routes=routes, lifespan=mcp_app.lifespan)
