# Presentation builder for Icosagen

This project generates PowerPoint presentations from a JSON slide schema and a PowerPoint template.

There are two ways to use it:

1. **OpenWebUI tool** – main way for normal users.
2. **Local (Python)** – for development / debugging.

---

## Slide JSON schema

All decks use the same structure:

```json
{
  "slides": [
    {
      "layout": 0,
      "placeholders": {
        "Title 1": "My Title",
        "Subtitle 2": "My subtitle"
      }
    }
  ]
}
```

* `slides`: non-empty list of slides.
* Each slide:

  * `layout`: integer index of a layout in the PPTX template.
  * `placeholders`: `{ "<placeholder name>": "<text or \n-separated bullets>" }`
* Bullets = one string with items separated by `\n`.

Examples (must match the template):

* Layout 0 (Title slide): `"Title 1"`, `"Subtitle 2"`
* Layout 1 (Title + content): `"Title 1"`, `"Content Placeholder 2"`
* Layout 2 (Section header): `"Title 1"`, `"Text Placeholder 2"`

See `prompt.json` for a fuller example.

---

## OpenWebUI: tool + model

### Tool (`tool_openwebui.py`)

* Exposes `create_presentation(slides)` to OpenWebUI.
* Does **only** JSON → PPTX:

  * loads `icosagen-template.pptx`,
  * fills placeholders,
  * saves `presentation_YYYYMMDD_HHMMSS.pptx` into `STATIC_DIR`,
  * returns a `file_url` in a `GenerationResult`.

**Valves:**

* `STATIC_DIR`: where files are written inside the container
  e.g. `/app/static`
* `STATIC_URL`: public base URL mapping to that directory
  e.g. `https://icosagenai.hpc.ut.ee/static/`
* `template_path`: absolute path to the PPTX template
  e.g. `/app/backend/data/icosagen-template.pptx`

### Model in OpenWebUI

A separate chat model knows how to build the JSON and call the tool.

* Name: `Presentation Generator`
* Base model: `gemma3:27b`
* Tool enabled: `create_presentation`

**System prompt behavior (short):**

* When the user asks for a *presentation/slides/deck*:

  1. Build a `slides` object with the previously shown shape:
  2. Call `create_presentation(slides=...)`.
  3. After the tool result:

     * If `status == "ok"` → reply: `Your presentation is ready: <file_url>`.
     * If `status == "error"` → show the error message.

* Placeholder rules:

  * Use only placeholders from the template (e.g. `"Title 1"`, `"Content Placeholder 2"`).
  * Bullets separated by `\n`.
  * Don’t write into Date/Footer/Slide Number placeholders.
  * Text only, no images.

---

## Local usage (dev)

For quick testing without OpenWebUI:
1. Check and install needed dependencies
2. Requires `.pptx` template in project root folder
3. Write a small script that imports the JSON → PPTX logic from `tool_local.py` and calls it with a manual `slides` dict, e.g.:

   ```python
   slides = {
       "slides": [
           {
               "layout": 0,
               "placeholders": {
                   "Title 1": "Local Test Deck",
                   "Subtitle 2": "Manual JSON"
               }
           }
       ]
   }
   ```

4. Run the script and open the generated `.pptx` to check layouts and placeholders.
