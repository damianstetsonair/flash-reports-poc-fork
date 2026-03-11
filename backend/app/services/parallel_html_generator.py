"""
Parallel HTML Generation Service

Optimized PPTX -> HTML conversion that processes each slide independently
and in parallel via Claude Vision, then merges all slide HTMLs with
unified CSS.

Pipeline:
  PPTX -> PDF -> split pages -> parallel Claude calls per slide -> CSS unify -> merge HTML
"""

import asyncio
import anthropic
import base64
import json
import re
import time
import tinycss2
from typing import List, Tuple, Dict, Any
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor

from app.config import ANTHROPIC_API_KEY, CLAUDE_MODEL_FAST, CLAUDE_MAX_TOKENS_FAST
from app.services.claude_html import HTML_TEMPLATE_PROMPT_BASE, fix_character_encoding


# Shared client and thread pool
_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
_executor = ThreadPoolExecutor(max_workers=10)


# ---------------------------------------------------------------------------
# Prompt construction
# ---------------------------------------------------------------------------

# We reuse the full HTML_TEMPLATE_PROMPT_BASE from claude_html.py which contains
# all quality sections: role, character_encoding, design_principles, layout_rules,
# pptx_compatibility (with CONTENT RICHNESS), text_handling, list_formatting,
# boxes_and_containers, tables_and_grids, quality_checklist.
#
# We only REPLACE <task> and <output_format> with single-slide versions.

SINGLE_SLIDE_TASK = """<task>
Analyze this single slide image and generate HTML/CSS that replicates EXACTLY its visual appearance.
This is SLIDE {slide_number} of {total_slides} in a presentation.

Your goal is to create a replica so faithful that when placed side-by-side with the original,
they are completely indistinguishable.

CRITICAL - EXACT REPLICATION:
This is NOT a template with placeholders. You must replicate the slide EXACTLY as it appears,
including ALL the original text, numbers, names, dates, and values shown in the image.

- Copy all text EXACTLY as shown (project names, person names, dates, numbers, percentages)
- Do NOT replace any text with placeholders like {{{{field_name}}}}
- Do NOT modify, summarize, or change any content
- The HTML should look IDENTICAL to the original slide

<visual_richness>
CRITICAL - PROFESSIONAL VISUAL QUALITY:
Each slide must look like it was designed by a professional presentation designer.
You MUST produce visually DENSE, RICH output. Do NOT simplify or produce minimal CSS.

REQUIRED visual elements (use ALL that appear in the image):
- Colored top-bar and footer-bar with solid brand colors or gradients
- Section headers with colored border-top separator lines (3-4px solid)
- Section boxes with borders, padding, and subtle background colors
- Bullet items with colored square bullets via CSS ::before pseudo-elements
- Progress bars with nested divs (outer background + inner colored fill with width in %)
- Status indicators using small colored spans with border-radius: 50%
- Trend arrows and KPI indicators with colored backgrounds
- Bold sub-labels within content sections
- Table cells with alternating row colors and header backgrounds
- Visual separators, dividers, and spacing that create a polished layout
- Box shadows on cards and containers for depth
- Precise color matching (#hex values extracted from the image, not generic colors)

DO NOT produce flat, unstyled HTML. Every element must have deliberate styling
including colors, borders, padding, margins, font-weights, and backgrounds.
The CSS for this single slide should be comprehensive (typically 40-80 CSS rules).
</visual_richness>
</task>"""

SINGLE_SLIDE_OUTPUT_FORMAT = """<output_format>
You must output TWO things for this slide:

1. **slide_html**: A single <div class="slide" data-slide-number="{slide_number}"> element containing
   all the slide content. Use the standard CSS class names on elements:
   .top-bar, .date-box, .main-title, .section-header, .section-title, .section-box,
   .bullet-item, .sub-label, .trend-box, .trend-item, .footer-bar, .page-number,
   .logo, .link-text
   You may also use inline styles for positioning and slide-specific values (top, left, width, height, colors).

2. **slide_css**: ALL CSS rules needed for this slide. EVERY rule MUST be scoped with the prefix:
   .slide[data-slide-number="{slide_number}"]

   The CSS MUST be comprehensive and include:
   - Pseudo-elements (::before, ::after) for bullets, decorative elements, and separators
   - Gradients (linear-gradient) on bars and backgrounds where the image shows them
   - Box-shadows on cards, containers, and elevated elements
   - Border properties (border, border-radius, border-top, border-left) for section styling
   - Font properties (font-size, font-weight, color, line-height, letter-spacing)
   - Background colors on ALL colored elements (do NOT leave backgrounds as default white)
   - Padding and margins for proper spacing inside containers

   Examples:
   .slide[data-slide-number="{slide_number}"] .top-bar {{ background: #003366; position: absolute; top: 0; left: 0; width: 960px; height: 8px; }}
   .slide[data-slide-number="{slide_number}"] .main-title {{ position: absolute; top: 15px; left: 20px; font-size: 22px; font-weight: bold; color: #003366; }}
   .slide[data-slide-number="{slide_number}"] .bullet-item::before {{ content: ""; display: inline-block; width: 6px; height: 6px; background: #003366; margin-right: 8px; vertical-align: middle; }}
   .slide[data-slide-number="{slide_number}"] .section-box {{ border: 1px solid #e0e0e0; border-radius: 4px; padding: 12px; background: #fafafa; }}
   .slide[data-slide-number="{slide_number}"] .footer-bar {{ position: absolute; bottom: 0; left: 0; width: 960px; height: 30px; background: #003366; }}

   Do NOT output unscoped rules. Every selector MUST start with .slide[data-slide-number="{slide_number}"].

Do NOT include <!DOCTYPE>, <html>, <head>, <body>, or <style> tags.
The slide div should have: class="slide" data-slide-number="{slide_number}"
</output_format>"""

SINGLE_SLIDE_FINAL = """
Generate the HTML and CSS for SLIDE {slide_number} of {total_slides}.

<overflow_prevention>
CRITICAL - ASPECT RATIO AND OVERFLOW RULES:
The slide container is EXACTLY 960px wide x 540px tall (16:9 aspect ratio).
ALL content MUST fit within these boundaries. Nothing may overflow or be clipped.

Layout boundaries:
- Usable content area: 20px padding on all sides = 920px wide x 480px tall
- Top bar: position absolute, top: 0, left: 0, width: 960px, height: 6-8px
- Footer bar: position absolute, bottom: 0, left: 0, width: 960px, height: 28-32px
- Content zone: from y=40px to y=508px (above footer)
- Two-column layout: left column ~460px, right column ~460px, with 20px gap

Title and header layout (CRITICAL - prevent overlaps):
- .main-title MUST span the full available width: left: 20px, width: 700-760px
  If there is a .date-box on the right, set title width so it STOPS 20px before the date-box left edge
  Example: date-box at left:780px -> title width: 740px (780 - 20 - 20)
- .date-box (if present): position absolute, top: 10-18px, RIGHT-aligned (left: 760-820px)
- .section-header spans full width of its parent container (no fixed narrow widths)
- .section-title inside .section-header: width: 100%, no fixed px width that could truncate text
- NO element may overlap another element at the same z-level. Check all top/left/width/height
  combinations to ensure bounding boxes do not intersect
- Long titles: use word-wrap: break-word and reduce font-size (18-20px) rather than letting text overflow

Font size guidelines (to prevent overflow):
- Slide title: 20-26px, font-weight: bold
- Section header title: 13-16px, font-weight: 600, text-transform: uppercase
- Body text / bullet items: 11-13px, line-height: 1.3-1.5
- Table cells: 10-12px
- Footer text / captions: 9-11px
- Page numbers: 10-12px

Element sizing rules:
- EVERY absolutely-positioned element MUST have explicit width in px
- EVERY absolutely-positioned element MUST have explicit height in px OR use bottom constraint
- Tables MUST have width in px on both the table element and each td/th
- Text containers must account for padding: usable_width = container_width - padding_left - padding_right
- Progress bars: outer div has explicit width in px, inner div uses percentage width

Overflow safety:
- Set overflow: hidden on ALL section-box and content container elements
- Use word-wrap: break-word on ALL text elements
- If content is dense, prefer smaller font sizes (11px body, 10px table) over clipping
</overflow_prevention>

<self_check>
MANDATORY - BEFORE RETURNING YOUR RESPONSE, mentally review every absolutely-positioned element
you created and verify there are NO collisions:
1. List each element's bounding box: (top, left, width, height) → compute bottom = top+height, right = left+width
2. For every pair of elements at the same depth, confirm their bounding boxes do NOT intersect:
   - Two boxes collide if: A.left < B.right AND A.right > B.left AND A.top < B.bottom AND A.bottom > B.top
3. If any pair collides, ADJUST positions/sizes before outputting. Common fixes:
   - Move the lower element's top below the upper element's bottom + 4px gap
   - Narrow an element's width so it stops before the neighbor's left edge
   - Reduce font-size to shrink an element's height
4. Pay special attention to: title vs date-box, section-box vs section-box, footer vs last content element
5. Verify the LAST element's bottom edge is ABOVE the footer-bar's top edge (at least 4px gap)
</self_check>

REMEMBER:
1. Replicate ALL text EXACTLY as shown - do NOT use placeholders
2. Use clean ASCII/HTML entities for bullets and special characters
3. Ensure pixel-perfect accuracy in positioning and styling
4. Use the required CSS class names (.top-bar, .main-title, .section-header, .section-box, .footer-bar, etc.)
5. Scope ALL CSS rules with .slide[data-slide-number="{slide_number}"]
6. The output must be professional, high-quality HTML+CSS with RICH visual styling
7. Include pseudo-elements (::before, ::after) for bullets and decorations
8. ALL elements must fit within the 960x540px container - verify no overflow
9. Use pixel values (px) for ALL positioning, widths, heights, and font sizes
10. CSS should be comprehensive: expect 40-80 rules per slide, not 10-15
"""


def _build_single_slide_prompt(slide_number: int, total_slides: int) -> str:
    """
    Build the full prompt for a single slide by taking the original
    HTML_TEMPLATE_PROMPT_BASE and replacing <task> and <output_format>.
    """
    base = HTML_TEMPLATE_PROMPT_BASE

    # Replace <task>...</task>
    base = re.sub(
        r"<task>.*?</task>",
        SINGLE_SLIDE_TASK.format(
            slide_number=slide_number, total_slides=total_slides
        ),
        base,
        flags=re.DOTALL,
    )

    # Replace <output_format>...</output_format>
    base = re.sub(
        r"<output_format>.*?</output_format>",
        SINGLE_SLIDE_OUTPUT_FORMAT.format(slide_number=slide_number),
        base,
        flags=re.DOTALL,
    )

    return base


# ---------------------------------------------------------------------------
# CSS unification (using tinycss2 for proper CSS parsing)
# ---------------------------------------------------------------------------

# Regex to strip the slide-number scope prefix from selectors
_SCOPE_RE = re.compile(
    r'\.slide\[data-slide-number="?\d+"?\]\s*',
)


def _normalize_declarations(body_str: str) -> str:
    """
    Normalize CSS declaration block for comparison.
    Uses tinycss2 to properly parse declarations (handles quoted strings,
    !important, complex values like content: "..." with special chars).
    """
    declarations = tinycss2.parse_declaration_list(body_str)
    decls = []
    for d in declarations:
        if d.type == 'declaration':
            name = d.name
            value = tinycss2.serialize(d.value).strip()
            important = " !important" if d.important else ""
            decls.append(f"{name}: {value}{important}")
    decls.sort()
    return "; ".join(decls)


def _unify_css(slide_css_list: List[str], total_slides: int) -> str:
    """
    Unify CSS from all slides using tinycss2 for proper parsing.

    Strategy:
    1. Parse all CSS rules from each slide using tinycss2
    2. Strip the [data-slide-number] scope to get the base selector
    3. Group by base selector
    4. If ALL slides define identical properties for a selector -> emit ONE unscoped rule
    5. If properties differ -> keep scoped versions per slide

    Handles pseudo-elements (::before, ::after), quoted strings with braces,
    @keyframes, @media at-rules, and complex selectors.

    Returns the unified CSS string.
    """
    # Map: base_selector -> { slide_idx -> (full_selector, normalized_declarations) }
    selector_map: Dict[str, Dict[int, Tuple[str, str]]] = defaultdict(dict)
    # At-rules kept per slide (cannot be meaningfully deduplicated)
    at_rules_collected: List[str] = []

    for slide_idx, css_text in enumerate(slide_css_list, 1):
        if not css_text or not css_text.strip():
            continue

        rules = tinycss2.parse_stylesheet(
            css_text, skip_comments=True, skip_whitespace=True
        )

        for rule in rules:
            if rule.type == 'qualified-rule':
                full_selector = tinycss2.serialize(rule.prelude).strip()
                body_str = tinycss2.serialize(rule.content).strip()

                # Strip scope to get base selector
                base_selector = _SCOPE_RE.sub("", full_selector).strip()
                if not base_selector:
                    base_selector = full_selector

                normalized = _normalize_declarations(body_str)
                selector_map[base_selector][slide_idx] = (full_selector, normalized)

            elif rule.type == 'at-rule':
                # @keyframes, @media, etc. -- pass through unmodified
                at_rule_str = tinycss2.serialize([rule]).strip()
                if at_rule_str:
                    at_rules_collected.append(at_rule_str)

    # Build unified output
    shared_rules = []
    specific_rules = []
    total_rules = 0
    deduped_count = 0

    for base_sel, slides_data in selector_map.items():
        total_rules += len(slides_data)
        unique_props = set(props for _, props in slides_data.values())

        if len(unique_props) == 1 and len(slides_data) == total_slides:
            # All slides have identical rules -> shared
            _, props = next(iter(slides_data.values()))
            shared_rules.append(f"{base_sel} {{ {props}; }}")
            deduped_count += len(slides_data) - 1
        else:
            # Different across slides -> keep scoped
            for slide_num in sorted(slides_data.keys()):
                full_sel, props = slides_data[slide_num]
                specific_rules.append(f"{full_sel} {{ {props}; }}")

    shared_css = "\n".join(shared_rules)
    specific_css = "\n".join(specific_rules)
    at_rules_css = "\n".join(at_rules_collected)

    total_size = len(shared_css) + len(specific_css) + len(at_rules_css)

    print(f"[parallel-html] --- CSS Unification ---")
    print(f"[parallel-html] Total CSS rules collected: {total_rules}")
    print(f"[parallel-html] Shared rules (deduplicated): {len(shared_rules)}")
    print(f"[parallel-html] Slide-specific rules: {len(specific_rules)}")
    if at_rules_collected:
        print(f"[parallel-html] At-rules (keyframes/media): {len(at_rules_collected)}")
    print(f"[parallel-html] Final CSS size: {total_size / 1024:.1f} KB")
    if deduped_count > 0:
        print(f"[parallel-html] Deduplicated {deduped_count} redundant rules")

    parts = []
    if shared_rules:
        parts.append("/* Shared across all slides */")
        parts.append(shared_css)
    if specific_rules:
        parts.append("\n/* Slide-specific */")
        parts.append(specific_css)
    if at_rules_collected:
        parts.append("\n/* At-rules (keyframes, media queries) */")
        parts.append(at_rules_css)

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Base CSS for the final document
# ---------------------------------------------------------------------------

BASE_CSS = """* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    background: #1a1a1a;
    padding: 20px;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
}
.slide {
    width: 960px;
    height: 540px;
    position: relative;
    overflow: hidden;
    margin: 20px auto;
    box-shadow: 0 8px 32px rgba(0,0,0,0.4);
    border-radius: 8px;
    background: #ffffff;
}"""


# ---------------------------------------------------------------------------
# Single slide generation (runs in thread pool)
# ---------------------------------------------------------------------------

def _generate_single_slide_html(
    image_bytes: bytes,
    media_type: str,
    slide_number: int,
    total_slides: int,
) -> Dict[str, str]:
    """
    Call Claude Vision for a single slide image.
    Returns dict with 'slide_html' and 'slide_css'.
    Runs synchronously (called from thread pool).
    """
    slide_start = time.time()
    tag = f"[Slide {slide_number}/{total_slides}]"

    print(f"[parallel-html] {tag} Starting Claude Vision call...")

    img_base64 = base64.b64encode(image_bytes).decode("utf-8")

    prompt_text = _build_single_slide_prompt(slide_number, total_slides)

    final_text = SINGLE_SLIDE_FINAL.format(
        slide_number=slide_number, total_slides=total_slides
    )

    content = [
        {"type": "text", "text": prompt_text},
        {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": media_type,
                "data": img_base64,
            },
        },
        {"type": "text", "text": final_text},
    ]

    collected = ""
    with _client.beta.messages.stream(
        model=CLAUDE_MODEL_FAST,
        max_tokens=CLAUDE_MAX_TOKENS_FAST,
        temperature=0.2,
        betas=["structured-outputs-2025-11-13"],
        messages=[{"role": "user", "content": content}],
        output_format={
            "type": "json_schema",
            "schema": {
                "type": "object",
                "properties": {
                    "slide_html": {
                        "type": "string",
                        "description": (
                            "The <div class='slide' data-slide-number='...'> element "
                            "with all slide content using CSS class names."
                        ),
                    },
                    "slide_css": {
                        "type": "string",
                        "description": (
                            "All CSS rules for this slide, each scoped with "
                            ".slide[data-slide-number='N'] prefix."
                        ),
                    },
                },
                "required": ["slide_html", "slide_css"],
                "additionalProperties": False,
            },
        },
    ) as stream:
        for text in stream.text_stream:
            collected += text

    result = json.loads(collected)

    slide_html = fix_character_encoding(result["slide_html"])
    slide_css = result["slide_css"]

    elapsed = time.time() - slide_start
    print(
        f"[parallel-html] {tag} Claude response received "
        f"({len(slide_html)} chars HTML, {len(slide_css)} chars CSS)"
    )
    print(f"[parallel-html] {tag} Completed in {elapsed:.1f}s")

    return {"slide_html": slide_html, "slide_css": slide_css, "time": elapsed}


# ---------------------------------------------------------------------------
# Main parallel orchestrator
# ---------------------------------------------------------------------------

async def generate_html_parallel(
    images: List[Tuple[bytes, str]],
) -> Dict[str, Any]:
    """
    Generate HTML from slide images in parallel.

    Each slide is sent to Claude independently, then all results are merged
    into a complete HTML document with unified CSS.

    Args:
        images: List of (image_bytes, media_type) tuples, one per slide

    Returns:
        Dictionary with:
          - full_html: Complete HTML document with unified CSS
          - slide_count: Number of slides
          - timings: Per-slide and total timing info
    """
    total_slides = len(images)
    start_time = time.time()

    print(f"[parallel-html] === Starting parallel generation for {total_slides} slides ===")
    print(f"[parallel-html] Launching {total_slides} parallel Claude Vision calls...")

    # Launch all slides in parallel using the thread pool
    loop = asyncio.get_event_loop()
    tasks = []
    for i, (img_bytes, media_type) in enumerate(images, 1):
        task = loop.run_in_executor(
            _executor,
            _generate_single_slide_html,
            img_bytes,
            media_type,
            i,
            total_slides,
        )
        tasks.append(task)

    # Wait for all to complete
    results = await asyncio.gather(*tasks, return_exceptions=True)

    # Collect results
    slide_htmls = []
    slide_css_list = []
    ok_count = 0
    fail_count = 0
    slide_times = []

    for i, result in enumerate(results, 1):
        tag = f"[Slide {i}/{total_slides}]"
        if isinstance(result, Exception):
            fail_count += 1
            print(f"[parallel-html] {tag} FAILED: {result}")
            slide_htmls.append(
                f'<div class="slide" data-slide-number="{i}" '
                f'style="width:960px; height:540px; position:relative; overflow:hidden; background:#ffffff;">'
                f'<div style="position:absolute; top:50%; left:50%; transform:translate(-50%,-50%); '
                f'color:#cc0000; font-size:18px; text-align:center;">'
                f"Error generating slide {i}:<br>{result}</div>"
                f"</div>"
            )
            slide_css_list.append("")
        else:
            ok_count += 1
            slide_htmls.append(result["slide_html"])
            slide_css_list.append(result["slide_css"])
            slide_times.append(result["time"])

    parallel_elapsed = time.time() - start_time
    print(
        f"[parallel-html] === All {total_slides} slides completed "
        f"({ok_count} OK, {fail_count} failed) in {parallel_elapsed:.1f}s ==="
    )

    # --- CSS Unification ---
    css_start = time.time()
    unified_css = _unify_css(slide_css_list, total_slides)
    css_elapsed = time.time() - css_start

    # --- Final HTML Assembly ---
    print(f"[parallel-html] --- Final HTML Assembly ---")

    slides_joined = "\n\n".join(slide_htmls)

    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Flash Report</title>
<style>
{BASE_CSS}

{unified_css}
</style>
</head>
<body>
{slides_joined}
</body>
</html>"""

    total_elapsed = time.time() - start_time
    avg_per_slide = total_elapsed / max(total_slides, 1)

    print(f"[parallel-html] Final HTML size: {len(full_html) / 1024:.1f} KB")
    print(
        f"[parallel-html] === Pipeline complete in {total_elapsed:.1f}s "
        f"(avg {avg_per_slide:.1f}s/slide) ==="
    )

    return {
        "full_html": full_html,
        "slide_count": total_slides,
        "ok_count": ok_count,
        "fail_count": fail_count,
        "timings": {
            "total_seconds": round(total_elapsed, 2),
            "parallel_seconds": round(parallel_elapsed, 2),
            "css_unify_seconds": round(css_elapsed, 4),
            "avg_per_slide": round(avg_per_slide, 2),
            "per_slide": [round(t, 2) for t in slide_times],
        },
    }
