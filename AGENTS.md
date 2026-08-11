# AGENTS.md

## Project Overview

MCP (Model Context Protocol) server built with **FastMCP 3.0** that exposes Office document generation as MCP tools. Runs as a Docker container (Python 3.12, Alpine) on port **8958** at `/mcp` using streamable-HTTP transport. Entry point: `main.py`.

## Architecture

```
main.py                  ← Registers all MCP tools on a single FastMCP instance
├── {docx,xlsx,pptx,email,xml}_tools/
│   ├── __init__.py      ← Re-exports the public function (e.g. markdown_to_word)
│   ├── base_*_tool.py   ← Core conversion logic (markdown → document bytes)
│   ├── helpers.py        ← Parsing, formatting, shared utilities
│   └── dynamic_*_tools.py  ← YAML-driven tool registration (docx, email only)
├── upload_tools/
│   ├── main.py          ← upload_file() dispatches to strategy backend
│   └── backends/{local,s3,gcs,azure,minio}.py
├── config.py            ← Singleton Config from env vars (Pydantic v2), logging setup
├── template_utils.py    ← Template resolution: custom_templates/ → default_templates/
└── middleware.py         ← Optional API key auth (Bearer / x-api-key header)
```

**Data flow:** Every tool converts input → in-memory bytes → calls `upload_file(file_obj, suffix)` → backend saves/uploads → returns URL or path string to the MCP client.

## Key Conventions

- **Config is centralized in `config.py`** — no module reads `os.environ` directly. Access via `get_config()` singleton.
- **Template resolution** (`template_utils.py`): searches `custom_templates/` before `default_templates/`, with `/app/*` container paths tried first, then local paths. Never hardcode template paths.
- **Dynamic tool registration**: YAML files in `config/` define parameterized email/docx templates. Each entry becomes a separate MCP tool at startup via `register_*_template_tools_from_yaml(mcp, path)`. Placeholders use Mustache syntax `{{name}}`. See `config/docx_templates.yaml` for the canonical example.
  - **Merged loading**: the master `config/<kind>_templates.yaml` is merged with UI-managed per-template files in `config/<kind>_templates.d/<name>.yaml` (the `.d` file wins on a name clash). The documented master files are never rewritten by tooling. Merge logic lives in `template_registry.py`.
  - **Live (un)registration**: `register_docx_template` / `register_email_template` (and their `unregister_*` counterparts) can add/replace/remove tools on a running `FastMCP` instance via `local_provider`, so the admin UI applies template edits without a restart. A module-level registry tracks live dynamic tool names.
- **Admin UI** (`admin/`, opt-in via `ADMIN_ENABLED`): a FastHTML template-admin for creating/editing dynamic docx+email templates without YAML. When enabled, `main.py` runs the **combined ASGI app** (`admin.app.build_combined_app`) under uvicorn — the MCP endpoint mounted at `/` and the admin UI under `ADMIN_PATH` (default `/admin`) in one process, so saving a template registers its MCP tool live. Modules: `store.py` (`TemplateStore` + `FileTemplateStore`, abstracted for a future shared backend), `analysis.py` (placeholder/conditional/style detection on uploaded assets), `preview.py` (render with sample values, no upload backend), `auth.py` (shared-password gate), `app.py` (views + app factory, including the Status page and document re-upload/re-scan). When `ADMIN_ENABLED` is unset, `main.py` keeps the original `mcp.run()` path unchanged.
- **Metrics** (`metrics.py`, project root): tiny always-on in-process counters (per-template calls/errors/last-used, recorded by the dynamic tool wrappers) plus a bounded recent-log ring buffer (installed by the admin app at startup). Surfaced on the admin Status page. No external deps; safe to import from the core tool modules.
- **Pydantic models** for tool arguments are created dynamically with `create_model()` in `dynamic_*_tools.py`. The `TYPE_MAP` dict maps YAML type strings to Python types.
- **Error handling in tools**: raise `fastmcp.exceptions.ToolError` for user-facing errors; use `RuntimeError` in upload/backend layers.
- **Logging**: use `logging.getLogger(__name__)` everywhere. Level controlled by `DEBUG` env var only.

## Filename Generation

**UUID prefix control** (`add_unique_prefix` parameter):
- All MCP tools accept an optional `add_unique_prefix: bool` parameter.
- The **default behavior depends on the storage strategy**:
  - **Traditional backends** (LOCAL, S3, GCS, AZURE, MINIO): defaults to `True` for collision safety in shared storage.
  - **LIBRECHAT**: defaults to `False` because LibreChat adds its own UUID prefix during file storage (`3deca384-6c9a-492a-ae2b-08ce22122cba__My_Report.docx`).
- When `True`: adds an 8-character UUID prefix for server-side uniqueness (e.g., `ff8ae81d_My_Report.docx`).
- When `False`: filenames are clean without UUID prefix (e.g., `My_Report.docx`).
- **Implementation**: The parameter flows through `upload_file_async()` → `upload_file()` → `generate_named_object_name()` in `upload_tools/`. When the parameter is `None` (not explicitly set), the default is resolved based on the storage strategy.
- **Dynamic tools**: YAML-defined tools may declare `add_unique_prefix` as an argument. `dynamic_docx_tools.py` and `dynamic_email_tools.py` read it with `.get("add_unique_prefix")` — deliberately **without** a default — so a template that does not declare it yields `None` and inherits the strategy-based default, exactly like the static tools. Do not reintroduce a `False` default there: the payload is built solely from the template's declared args, so `False` would reach `upload_file()` for every template, satisfy its `is None` guard, and silently suppress the prefix on every backend.

## Adding a New Document Tool

1. Create `<type>_tools/` package with `__init__.py`, `base_<type>_tool.py`, and optional `helpers.py`.
2. The base tool function should: accept content → produce an `io.BytesIO` → call `upload_file(buffer, "<ext>")` → return the result string.
3. Register the async wrapper in `main.py` using `@mcp.tool(name=..., description=..., tags=..., annotations=...)`.
4. Use `Annotated[<type>, Field(description=...)]` for all tool parameters — the descriptions are critical because MCP clients (AI models) rely on them.

## Tests

```bash
pytest                        # Run all tests (asyncio_mode=auto in pytest.ini)
pytest tests/test_docx_base.py  # Single module
```

- Tests live in `tests/` and output generated files to `tests/output/{docx,pptx,xlsx}/` for manual inspection.
- Upload is mocked in tests — patch `upload_file` or the specific `*_tool.upload_file` to capture bytes without needing a real backend. See `test_xlsx_creation.py::_create_workbook_from_markdown` for the pattern.
- PPTX tests instantiate `PowerpointPresentation` directly and call `.save()` to get a buffer, bypassing upload entirely.
- No `.env` required for tests — `config.py` defaults to `LOCAL` strategy and INFO logging.
