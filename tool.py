from pydantic import BaseModel, Field
from typing import Any, Dict, Optional
from datetime import datetime
from pptx import Presentation as PPTXPresentation
import os
import json
import requests
from dotenv import load_dotenv
load_dotenv()


class GenerationResult(BaseModel):
    status: str
    message: str
    file_url: str
    mimetype: str
    slides: Optional[Any] = None

def make_llm_callable():
    base = "https://icosagenai.hpc.ut.ee"
    token = os.getenv("OPENWEBUI_TOKEN")
    if not token:
      raise RuntimeError("Missing OPENWEBUI_TOKEN")
    model = "gemma3:27b"
    
    def _call(prompt: str) -> str:
        payload = {
            "model": model,
            "messages": [{"role": "user", "content": prompt}],
        }
        r = requests.post(
            f"{base}/api/chat/completions",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload,
        )
        r.raise_for_status()
        data = r.json()
        return data["choices"][0]["message"]["content"]
    return _call

class Tools:
    class Valves(BaseModel):
        STATIC_DIR: str = Field("output")
        STATIC_URL: str = Field("https://localhost:8000/static/")
        template_path: str = Field("icosagen-template.pptx")
       
    
    def __init__(self, llm_callable=None):
        self.valves = self.Valves()
        os.makedirs(self.valves.STATIC_DIR, exist_ok=True)
        self.llm_callable = llm_callable or make_llm_callable()
    
    def _call(self, prompt: str) -> str:
        return self.llm_callable(prompt)
    
    
    async def create_presentation(
        self,
        prompt: str,
        __event_emitter__=None
    ) -> GenerationResult:
        try:
            data = self._parse_prompt_to_json(prompt)
            result = self._build_and_save_from_json(
                data=data,
                filename_prefix="presentation",
                template_path=None,
            )
            return GenerationResult(
                status="ok",
                message=result["message"],
                file_url=result["file_url"],
                mimetype=result["mimetype"],
                slides=data,
            )
        except Exception as e:
            return GenerationResult(
                status="error",
                message=str(e),
                file_url="",
                mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                slides=None,
            )

    def _parse_prompt_to_json(self, prompt: str) -> Dict[str, Any]:
        instruction = '''
            You are a converter that outputs ONLY a JSON object for python-pptx slide generation.
REQUIRED OUTPUT FORMAT (no code fences, no extra text):
{
  "slides": [
    { "layout": <int>, "placeholders": { "<Placeholder Name>": "<text or \\n-separated bullets>" } }
  ]
}
RULES
- Use ONLY the layouts and placeholder names listed below.
- Put text ONLY into non-meta placeholders (Title, Subtitle, Text, Content, Vertical Text, Picture Caption Text).
- Do NOT set values for Date, Footer, or Slide Number placeholders (these are auto-handled by the template).
- Prefer:
  - layout 0 for the deck title slide,
  - layout 2 for section headers,
  - layout 1 for single-column content,
  - layout 3 for two-column content,
  - layout 4 for comparisons,
  - layout 5 for title-only,
  - layout 7 for content with a caption,
  - layout 9/10 for vertical text when explicitly requested,
  - layout 6 (blank) only if the user insists on a blank slide.
- Bullet lists should be newline-separated within a single string value.
- Do NOT include images or files. Avoid layout 8 (Picture with Caption) unless text-only is acceptable in its text placeholder.
- Output must be valid JSON (UTF-8), no comments, no trailing commas.

ALLOWED LAYOUTS AND PLACEHOLDERS
layout 0: Title Slide
  - Title 1
  - Subtitle 2
  - Date Placeholder 3
  - Footer Placeholder 4
  - Slide Number Placeholder 5

layout 1: Title and Content
  - Title 1
  - Content Placeholder 2
  - Date Placeholder 3
  - Footer Placeholder 4
  - Slide Number Placeholder 5

layout 2: Section Header
  - Title 1
  - Text Placeholder 2
  - Date Placeholder 3
  - Footer Placeholder 4
  - Slide Number Placeholder 5

layout 3: Two Content
  - Title 1
  - Content Placeholder 2
  - Content Placeholder 3
  - Date Placeholder 4
  - Footer Placeholder 5
  - Slide Number Placeholder 6

layout 4: Comparison
  - Title 1
  - Text Placeholder 2
  - Content Placeholder 3
  - Text Placeholder 4
  - Content Placeholder 5
  - Date Placeholder 6
  - Footer Placeholder 7
  - Slide Number Placeholder 8

layout 5: Title Only
  - Title 1
  - Date Placeholder 2
  - Footer Placeholder 3
  - Slide Number Placeholder 4

layout 6: Blank
  - Date Placeholder 1
  - Footer Placeholder 2
  - Slide Number Placeholder 3

layout 7: Content with Caption
  - Title 1
  - Content Placeholder 2
  - Text Placeholder 3
  - Date Placeholder 4
  - Footer Placeholder 5
  - Slide Number Placeholder 6

layout 8: Picture with Caption
  - Title 1
  - Picture Placeholder 2
  - Text Placeholder 3
  - Date Placeholder 4
  - Footer Placeholder 5
  - Slide Number Placeholder 6

layout 9: Title and Vertical Text
  - Title 1
  - Vertical Text Placeholder 2
  - Date Placeholder 3
  - Footer Placeholder 4
  - Slide Number Placeholder 5

layout 10: Vertical Title and Text
  - Vertical Title 1
  - Vertical Text Placeholder 2
  - Date Placeholder 3
  - Footer Placeholder 4
  - Slide Number Placeholder 5

TASK
Convert the user's input into a coherent slide deck using the schema above. Choose appropriate layouts and populate only the allowed placeholders with text. Output ONLY the JSON object.
    '''.strip()

        llm_prompt = f"{instruction}\n\n=== USER INPUT ===\n{prompt}\n=== END ==="
        raw = self.llm_callable(llm_prompt)

        if raw.strip().startswith("```"):
            raw = raw.strip().strip("`")
            if raw.lower().startswith("json"):
                raw = raw[4:].lstrip()

        return json.loads(raw)

    def _build_and_save_from_json(
        self,
        data: Dict[str, Any],
        filename_prefix: str,
        template_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        final_template = template_path or self.valves.template_path
        if not os.path.exists(final_template):
            raise FileNotFoundError(f"Template not found: {final_template}")

        prs = PPTXPresentation(final_template)

        slides_data = data.get("slides", [])

        # fill first existing slide
        if prs.slides and slides_data:
            first_slide = prs.slides[0]
            self._fill_slide_placeholders(first_slide, slides_data[0].get("placeholders", {}))
        #create new slides
        for slide_data in slides_data[1:]:
            layout_idx = slide_data.get("layout", 1)
            layout = prs.slide_layouts[layout_idx]
            slide = prs.slides.add_slide(layout)
            self._fill_slide_placeholders(slide, slide_data.get("placeholders", {}))

        # save
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{filename_prefix}_{timestamp}.pptx"
        output_path = os.path.join(self.valves.STATIC_DIR, filename)
        prs.save(output_path)
        file_url = self.valves.STATIC_URL + filename
        return {
            "status": "success",
            "message": f"Generated: {file_url}",
            "file_url": file_url,
            "mimetype": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        }

   
    def _fill_slide_placeholders(self, slide, placeholders: Dict[str, Any]):
        for shape in slide.shapes:
            if not getattr(shape, "is_placeholder", False):
                continue
            name = shape.name
            if name in placeholders and hasattr(shape, "text_frame"):
                shape.text_frame.clear()
                shape.text_frame.text = str(placeholders[name])
                
if __name__ == "__main__":
    import asyncio
    async def _run():
        tools = Tools()
        prompt = ('''Generate slideshow based on the following text (10 slides)
                  Antibody discovery is undergoing a shift from artisanal, target-by-target biology toward an engineered, scalable discipline. Classical methods—animal immunization, hybridoma generation, low-throughput screening—built the first generations of successful biologics, but they are slow, resource-intensive, and poorly suited to today’s targets, which are often weakly immunogenic, highly conserved, or structurally complex. As pipelines expand and timelines tighten, the industry can’t afford discovery processes where success depends on chance immune responses or laborious manual selection. The core problem isn’t that old methods don’t work; it’s that they don’t explore enough sequence space, they don’t generate enough comparative data, and they don’t give early visibility into developability risks.
To address that, discovery is moving toward high-diversity, high-throughput systems. Display technologies—phage, yeast, ribosome, and increasingly mammalian display—let researchers start from libraries with billions of variants and repeatedly enrich for binders that meet predefined criteria. When this is coupled with next-generation sequencing of input and output libraries, teams can track clonal expansion, identify enriched CDR motifs, and understand why certain binders outperform others. That turns discovery from “find whatever sticks” into “systematically converge on the best binders.” Even more importantly, the same infrastructure can be reused across targets, so each campaign is not a bespoke one-off but part of a repeatable platform.
In parallel, single-cell technologies and B-cell repertoire mining are making it much easier to harvest antibodies that biology has already optimized. By capturing natural heavy–light chain pairs from immunized animals, human donors, or convalescent patients, researchers can shortcut a lot of the reformatting and pairing problems that plagued older approaches. This matters in infectious disease, oncology, and autoimmunity, where the winning antibodies can be rare, highly somatically mutated, or narrowly specific. The bottleneck is no longer “can we capture the cells?” but “can we prioritize thousands of plausible antibodies quickly and objectively?” That’s pushing groups to standardize functional assays, expression tests, and early liability screening so bad candidates are eliminated before expensive characterization.
The biggest acceleration, however, is coming from computation and AI. Structure prediction, in silico affinity maturation, paratope–epitope modeling, and ML models trained on large antibody–antigen datasets make it possible to rank or even generate candidates before wet-lab work. This directly attacks one of the costliest problems in biologics R&D: late-stage attrition due to poor biophysics, immunogenicity flags, or developability issues. If a model can say, “these 200 variants are likely to express, fold, and avoid common liabilities,” labs can focus their bench time where it counts. The most advanced groups are already running closed-loop systems—design → express → assay → feed data back to the model → redesign—so every campaign improves the model, and every model iteration improves the next campaign. That’s how discovery stops being a linear pipeline and becomes a learning system.
Put together, these trends point to an integrated, automated antibody discovery workflow where data, not intuition, drives decisions. Robotics and standardized assay panels reduce human variability; LIMS and FAIR data practices make results comparable across campaigns and sites; and shared scoring frameworks let R&D, CMC, and even downstream clinical teams speak the same language about “what good looks like.” The payoff isn’t just more antibodies—it’s better antibodies, found earlier, with clearer rationale and lower downstream risk. That is what “revolutionizing antibody discovery” actually means in practice, and it’s the story worth telling on your slides.
                  '''
        )
        result = await tools.create_presentation(prompt)
        print(result.status)
        print(result.message)
        print(result.file_url)
        print(result.mimetype)

    asyncio.run(_run())
