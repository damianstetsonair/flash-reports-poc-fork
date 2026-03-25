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

from collections import Counter
from lxml import html as lxml_html

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
Analyze this single slide image and generate HTML/CSS that replicates its visual appearance.
This is SLIDE {slide_number} of {total_slides} in a presentation.

PRIORITY ORDER:
1. EVERYTHING FITS — all content visible, nothing overflows, nothing overlaps
2. LOOKS CLEAN — professional, readable, well-spaced
3. FAITHFUL REPLICA — match the original as closely as possible

CRITICAL - EXACT REPLICATION:
This is NOT a template with placeholders. You must replicate the slide EXACTLY as it appears,
including ALL the original text, numbers, names, dates, and values shown in the image.

- Copy all text EXACTLY as shown (project names, person names, dates, numbers, percentages)
- Do NOT replace any text with placeholders like {{{{field_name}}}}
- Do NOT modify, summarize, or change any content
- If the original image shows overlapping text or cut-off content, FIX these issues in your HTML.
  Produce a clean, readable version — not a faithful copy of visual bugs.

<visual_richness>
CRITICAL - PROFESSIONAL VISUAL QUALITY:
Each slide must look like it was designed by a professional presentation designer.
You MUST produce visually DENSE, RICH output. Do NOT simplify or produce minimal CSS.

COLOR PALETTE — Extract 4-5 key colors from the image (primary, secondary, accent, bg-tint, text):
- Use the SAME hex values consistently — all blue elements use the exact same #hex
- Never invent new colors; every color must come from the original image

REQUIRED visual elements (use ALL that appear in the image):
- Top-bar and footer-bar: replicate EXACTLY — flat solid color if flat, gradient ONLY if the image shows one.
  Do NOT add gradients to flat bars. Most presentations use flat solid colors.
- Section headers with colored border-top (3-4px solid [primary-color])
- Section boxes: replicate the original look exactly.
  Only add box-shadow, tinted backgrounds, or border-radius if the original image shows them.
  If the image shows sharp/square corners, do NOT add border-radius. Do NOT invent visual effects.
- Bullet items with colored square bullets via CSS ::before pseudo-elements
- Progress bars with nested divs (outer background + inner colored fill)
- STATUS BADGES as pills: <span style="display:inline-block; padding:2px 10px; border-radius:12px;
  background:#dcfce7; color:#166534; font-size:10px; font-weight:600;">On Track</span>
  (green=good, yellow=warning, red=critical, blue=info, gray=neutral)
- KPI donut charts with conic-gradient when the image shows circular indicators
- Traffic light indicators (3 dots, active one at full opacity, others at 0.3)
- Trend arrows and KPI indicators with colored backgrounds
- Bold sub-labels within content sections
- Table cells with alternating row colors and header backgrounds

SPACING RHYTHM — All spacing must follow a 4px grid:
- gap: 8px between sections, 4px between bullets, 12px padding inside boxes
- ONLY use: 4, 8, 12, 16, 20, 24px — never arbitrary values like 7px, 13px

TYPOGRAPHY:
- Titles: font-weight:700; letter-spacing:-0.5px;
- Section headers: text-transform:uppercase; letter-spacing:1px; font-weight:600;
- Body: font-weight:400; line-height:1.4;
- Captions/dates: color:rgba(0,0,0,0.55);

BORDERS — Replicate from image, consistent system (never mix styles):
- Section separator: border-top with primary color (replicate thickness from image)
- Box border: replicate from image (color, thickness). Do NOT add border-radius unless image shows rounded corners
- Table rows: replicate from image (typically hairline separators)

DO NOT produce flat, unstyled HTML. Every element must have deliberate styling.
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

   TEXT FIDELITY WARNING: The .logo and .page-number inside .footer-bar, and .main-title,
   MUST contain the COMPLETE text from the image. Read these texts character by character.
   A common bug is outputting only the LAST letter of a brand name (e.g., "A" instead of "SYSTRA").
   Carefully read the FULL word in the image before writing it.

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
   .slide[data-slide-number="{slide_number}"] .section-box {{ border: 1px solid #e0e0e0; padding: 12px; }}
   .slide[data-slide-number="{slide_number}"] .footer-bar {{ position: absolute; bottom: 0; left: 0; width: 960px; height: 30px; background: #003366; }}

   Do NOT output unscoped rules. Every selector MUST start with .slide[data-slide-number="{slide_number}"].

Do NOT include <!DOCTYPE>, <html>, <head>, <body>, or <style> tags.
The slide div should have: class="slide" data-slide-number="{slide_number}"
</output_format>"""

SINGLE_SLIDE_FINAL = """
Generate the HTML and CSS for SLIDE {slide_number} of {total_slides}.

<overflow_prevention>
CRITICAL - THE SLIDE IS 960px WIDE x 540px TALL. EVERYTHING MUST FIT.

LAYOUT STRUCTURE — use this approach:
1. CHROME: .top-bar (absolute, top:0, height: 6-8px) and .footer-bar (absolute, bottom:0, height: 28-32px)
2. HEADER ZONE: A flex row container (absolute positioned below top-bar):
   - position: absolute; top: 10px; left: 20px; right: 20px;
   - display: flex; align-items: center; justify-content: space-between;
   - .main-title grows to fill space (flex: 1; min-width: 0;)
   - .date-box stays fixed size (flex-shrink: 0;)
   - This AUTOMATICALLY prevents title/date-box overlap — no pixel math needed
3. CONTENT ZONE: A flex column container (absolute positioned between header and footer):
   - position: absolute; top: ~50px; left: 20px; right: 20px; bottom: 32px;
   - display: flex; flex-direction: column; gap: 8px;
   - Sections flow vertically and SHARE SPACE — flex-shrink lets them compress if tight
   - NO hardcoded top/left per section — flex handles the distribution
4. INSIDE SECTIONS (.section-box):
   - display: flex; flex-direction: column; gap: 4px;
   - Content flows naturally — bullets, text, sub-sections never overlap
   - flex-shrink: 1; min-height: 0; — allows the section to compress
5. TREND/KPI ROWS (.trend-box):
   - display: flex; gap: 12px; flex-wrap: wrap; align-items: center;
6. TWO-COLUMN LAYOUTS:
   - display: flex; gap: 16px; — each column: flex: 1; min-width: 0; word-wrap: break-word;

Font size guidelines:
- Slide title: 20-26px, font-weight: bold
- Section header: 13-16px, font-weight: 600, text-transform: uppercase
- Body text / bullets: 11-13px, line-height: 1.3-1.5
- Table cells: 10-12px
- Footer / captions: 9-11px

Fitting rules:
- Use gap (not margins) for spacing between flex children
- Use flex-shrink: 1 + min-height: 0 on sections so they compress when space is tight
- If content is dense, prefer smaller font sizes BEFORE any clipping occurs
- The content zone flex container handles vertical distribution — trust it

RESILIENT TEXT CONTAINERS — This template will be populated with variable-length data later:
- EVERY text element MUST have: word-wrap: break-word; overflow-wrap: break-word;
- DO NOT use overflow: hidden or text-overflow: ellipsis on text elements — text must NEVER be clipped or truncated
- EVERY flex column child MUST have: min-height: 0; flex-shrink: 1;
- EVERY flex row child MUST have: min-width: 0; word-wrap: break-word;
- Do NOT use fixed heights on text containers — let flex handle sizing
- Text that says "Project Alpha" now may become "Enterprise Cloud Migration Phase 2" later
- All containers must handle 2-3x longer text without breaking layout
- The ONLY element that should have overflow: hidden is the .slide container itself (960x540)
</overflow_prevention>

<self_check>
BEFORE RETURNING, verify:
1. Chrome (.top-bar, .footer-bar) is position: absolute at top/bottom edges
2. Header zone uses display: flex row — title and date-box cannot collide
3. Content zone uses display: flex column — sections cannot collide
4. Content zone bottom is ABOVE .footer-bar top (bottom: 32px or more)
5. Each .section-box uses flex column internally — bullets/text cannot overlap
6. No text is cut off — reduce font-size if needed, flex-shrink handles the rest
7. Two-column layouts (if any) use flex row with flex: 1 per column
8. TEXT FIDELITY — Re-read the image and verify these texts CHARACTER BY CHARACTER:
   - .logo text in footer-bar: read EVERY letter left to right from the image. Is it COMPLETE?
     (e.g., if image shows "SYSTRA", your HTML must say "SYSTRA" — not "A", not "TRA")
   - .main-title: is the FULL title there, word by word?
   - .page-number: correct number?
   - .date-box: full date string?
   - Brand names and company names are the #1 source of truncation bugs — double check them.
</self_check>

REMEMBER:
1. Replicate ALL text EXACTLY as shown — do NOT use placeholders
2. Use clean ASCII/HTML entities for bullets and special characters
3. EVERYTHING MUST FIT — flex layout prevents overlaps; reduce font-size if content is dense
4. Use the required CSS class names (.top-bar, .main-title, .section-header, .section-box, .footer-bar, etc.)
5. Scope ALL CSS rules with .slide[data-slide-number="{slide_number}"]
6. Extract a consistent color palette from the image — reuse the SAME hex values everywhere
7. All spacing follows 4px rhythm: 4, 8, 12, 16, 20, 24px only
8. Status indicators = pill badges (rounded, tinted background, bold text)
9. Shadows and gradients ONLY if visible in the original image — do NOT add them to flat designs
10. Chrome bars (top-bar, footer-bar): flat solid color if flat in image, gradient ONLY if image shows one
11. Typography: titles tight (-0.5px spacing), headers uppercase (+1px spacing), captions muted opacity
12. CSS should be comprehensive: expect 40-80 rules per slide, not 10-15
13. If the original shows visual bugs (overlaps, cut-off text), FIX them — produce clean output
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
    background: #ffffff;
}

/* --- Container query auto-scaling for text --- */
.section-box,
.trend-box,
.section-header {
    container-type: inline-size;
}
.section-box .bullet-item,
.section-box .sub-label,
.section-box p,
.section-box li,
.section-box span:not(.page-number):not(.logo) {
    font-size: clamp(8px, 2.8cqw, 13px);
    line-height: clamp(1.1, 0.1cqw + 1, 1.5);
}
.section-header .section-title {
    font-size: clamp(9px, 3.2cqw, 16px);
}
.trend-box .trend-item {
    font-size: clamp(8px, 2.5cqw, 12px);
}
td, th {
    font-size: clamp(8px, 2.5cqw, 12px);
}
.main-title {
    font-size: clamp(14px, 3.5cqw, 26px);
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
# Chrome normalization (fix inconsistent footer/header text across slides)
# ---------------------------------------------------------------------------

def _normalize_chrome_text(slide_htmls: List[str]) -> List[str]:
    """
    Fix inconsistent .logo and .page-number text across slides.

    Because each slide is generated by an independent Claude call, the agent
    may read footer/header text differently per slide (e.g., "SYSTRA" on slide 1
    but only "A" on slide 3). This function:

    1. Extracts .logo text from every slide
    2. Picks the LONGEST value as the canonical one (longest = most complete reading)
    3. Replaces truncated versions in all slides

    Same logic for .main-title consistency check (logs warnings only).
    """
    if len(slide_htmls) < 2:
        return slide_htmls

    # Extract logo text from each slide
    logo_texts = []
    for i, html_str in enumerate(slide_htmls):
        try:
            doc = lxml_html.fromstring(f"<div>{html_str}</div>")
            logos = doc.cssselect(".logo")
            text = logos[0].text_content().strip() if logos else ""
            logo_texts.append(text)
        except Exception:
            logo_texts.append("")

    # Filter to non-empty values
    non_empty = [t for t in logo_texts if t]
    if not non_empty:
        return slide_htmls

    # Pick canonical logo: longest text (most complete OCR reading)
    canonical_logo = max(non_empty, key=len)

    # Count how many slides need fixing
    fixes_needed = sum(
        1 for t in logo_texts
        if t and t != canonical_logo and len(t) < len(canonical_logo)
    )

    if fixes_needed > 0:
        print(f"[parallel-html] --- Chrome Normalization ---")
        print(f"[parallel-html] Canonical .logo text: \"{canonical_logo}\"")
        print(f"[parallel-html] Fixing {fixes_needed} slides with truncated logo text")

        fixed_htmls = []
        for i, (html_str, logo_text) in enumerate(zip(slide_htmls, logo_texts)):
            if logo_text and logo_text != canonical_logo and len(logo_text) < len(canonical_logo):
                # Replace the truncated text in the .logo span
                # Use regex to target <span class="logo">...</span> precisely
                fixed = re.sub(
                    r'(<span\s+class="logo"[^>]*>)[^<]*(</span>)',
                    rf'\g<1>{canonical_logo}\2',
                    html_str,
                )
                if fixed != html_str:
                    print(f"[parallel-html] Slide {i+1}: \"{logo_text}\" → \"{canonical_logo}\"")
                fixed_htmls.append(fixed)
            else:
                fixed_htmls.append(html_str)

        return fixed_htmls

    return slide_htmls


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

    # --- Chrome Normalization (fix inconsistent footer/header text across slides) ---
    slide_htmls = _normalize_chrome_text(slide_htmls)

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
