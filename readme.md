# Presentation Generation Tool

![Python](https://img.shields.io/badge/Python-3.11+-blue)
![Open WebUI](https://img.shields.io/badge/Open%20WebUI-tool-orange)

This repository contains the source code for the presentation generation tool developed as part of the bachelor's thesis *"Design and Implementation of a Presentation Generation Pipeline Using Large Language Models"* (University of Tartu, Institute of Computer Science, 2026).

The tool integrates into Icosagen's self-hosted Open WebUI + Ollama infrastructure and generates `.pptx` presentations from natural-language prompts using a planning -> rendering workflow driven by a large language model.

## Table of Contents

- [Features](#-features)
- [How It Works](#-how-it-works)
- [Requirements](#-requirements)
- [Installation](#-installation)
- [Usage](#-usage)
- [Citation](#-citation)

## Features

- **Two-phase workflow** - the model first proposes a slide outline based on the template, then renders the presentation after the user confirms.
- **Template-driven rendering** — uses Icosagen's existing PowerPoint templates, preserving corporate branding and typography without per-slide styling.
- **Local-only** — runs entirely inside the Open WebUI + Ollama container. Data is not sent to external services.
- **Document and prompt input** — generates presentations from a text prompt, an attached document, or a combination of both.
- **Chart and table support** — produces native PowerPoint charts and tables.
- **Multi-template** — users can switch between available templates from a dropdown in the chat UI.


## How It Works

The tool exposes two methods to the LLM through Open WebUI:

1. **`get_template_manifest()`** inspects the active template and returns slide dimensions, available layouts, placeholder indices and types, and supported chart types. The model uses this to plan an outline.
2. **`render_presentation(spec)`** takes a validated `PresentationSpec` (a structured slide specification) and produces a `.pptx` file. The model calls this only after the user confirms the outline.

### Slide specification

Each slide is defined by a layout index, a set of placeholder keys mapped to text, and optionally a chart or a table:

```json
{
  "slides": [
    {
      "layout": 0,
      "placeholders": {
        "title": "Quarterly results",
        "subtitle": "Q4 2025"
      }
    },
    {
      "layout": 5,
      "placeholders": {
        "title": "Revenue by region"
      },
      "chart": {
        "chart_type": "column",
        "categories": ["EU", "US", "APAC"],
        "series": {"Revenue (M€)": [12.4, 8.1, 4.7]}
      }
    }
  ]
}
```

Placeholder keys may be:
- **Semantic**: `title`, `subtitle`, `body`, `body2`, `body3` (resolved by placeholder type)
- **Explicit**: a placeholder index as a string (e.g. `"13"`)

Bullet points within a single placeholder are separated by `\n`. The full schema is defined by the Pydantic models in `tool.py` (`PresentationSpec`, `SlideSpec`, `ChartSpec`, `TableSpec`).

## Requirements

**Runtime environment:**
- Python 3.11 or later
- Open WebUI
- Ollama with at least one tool-capable model installed

**Python libraries:**
- `python-pptx`
- `pydantic`

Both are bundled with Open WebUI's standard Docker image;

**Template:**
- A valid `.pptx` file in the directory configured by `TEMPLATES_DIR` (default: `/app/backend/data/templates`).

## Installation

The tool is deployed as an Open WebUI tool, not run as a standalone script.

1. Open the Open WebUI workspace as an administrator and navigate to **Workspace → Tools**.
2. Click **+** to create a new tool and paste the contents of `tool.py`.
3. Save the tool. It will appear in the tool list as *Presentation generating tool*.
4. Place a `.pptx` template into the directory configured by the `TEMPLATES_DIR` valve.
5. Configure the `STATIC_DIR` and `STATIC_URL` valves so generated files are written to a directory served publicly by Open WebUI's static file route.
6. Create a model in **Workspace → Models** that uses the system prompt from `system_prompt.txt` and enable the tool for it.

### Valve configuration

| Valve | Scope | Description |
|---|---|---|
| `STATIC_DIR` | Admin | Filesystem path inside the container where generated `.pptx` files are written. |
| `STATIC_URL` | Admin | Public base URL mapping to `STATIC_DIR`, used to construct download links. |
| `TEMPLATES_DIR` | Admin | Directory containing `.pptx` template files. |
| `template` | User | The active template, selected via dropdown in the chat UI. |

## Usage

1. Start a new chat using the *Presentation Generator* model.
2. Provide a prompt describing the desired presentation (optionally attaching documents).
3. Review the outline the tool proposes and confirm or request changes.
4. Download the generated `.pptx` via the link returned by the tool.
