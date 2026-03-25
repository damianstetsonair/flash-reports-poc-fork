"""
Claude Vision HTML Generation Service

Uses Claude's vision capabilities to convert slide images into pixel-perfect HTML templates.
"""

import anthropic
import base64
import json
from typing import List, Tuple, Dict, Any

from app.config import ANTHROPIC_API_KEY, CLAUDE_MODEL, CLAUDE_MAX_TOKENS


# Initialize Anthropic client
client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


# Base prompt for generating HTML template from slide images
HTML_TEMPLATE_PROMPT_BASE = """<role>
You are an elite frontend developer with 15+ years of experience in pixel-perfect HTML/CSS replication.
Your specialty is converting visual designs into flawless, production-ready code.
</role>

<task>
Analyze each slide image and generate HTML/CSS that replicates EXACTLY its visual appearance.
Your goal is to create a replica so faithful that when placed side-by-side with the original,
they are completely indistinguishable.

CRITICAL - EXACT REPLICATION:
This is NOT a template with placeholders. You must replicate the slides EXACTLY as they appear,
including ALL the original text, numbers, names, dates, and values shown in the images.

- Copy all text EXACTLY as shown (project names, person names, dates, numbers, percentages)
- Do NOT replace any text with placeholders like {{field_name}}
- Do NOT modify, summarize, or change any content
- The HTML should look IDENTICAL to the original slides
</task>

<character_encoding>
CRITICAL - CHARACTER HANDLING:
- Use standard ASCII characters for all bullets and symbols
- For bullet points, use these HTML entities or CSS:
  • Use "•" (bullet) or CSS list-style-type: disc
  • Use "–" for en-dash, "—" for em-dash
  • Use "→" for arrows
- NEVER output garbled characters like "â–ª", "â€"", "â€™"
- If you see special Unicode characters in the image, convert them to their HTML entity equivalents
- For checkmarks use ✓ or &#10003;
- For X marks use ✗ or &#10007;
- Ensure the HTML has: <meta charset="UTF-8">
</character_encoding>

<design_principles>
COLOR PALETTE — Extract and reuse consistently:
- Identify the 4-5 key colors from the slide: primary (brand), secondary, accent, background tint, text color
- Use hex codes (#FF5733) — extract from the image, do NOT invent new colors
- Reuse the SAME palette across all elements: top-bar, section-header borders, bullet colors, status badges
- If the slide uses a dark blue (#003366), ALL blue elements must use that exact shade — no variations

TYPOGRAPHY SYSTEM:
- Font family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif
- Titles (.main-title): font-weight: 700; letter-spacing: -0.5px; (tight, professional)
- Section headers (.section-title): text-transform: uppercase; letter-spacing: 1px; font-weight: 600;
- Body text / bullets: font-weight: 400; line-height: 1.4;
- Captions, footnotes, dates: color with reduced opacity (e.g., rgba(0,0,0,0.55)); font-weight: 400;
- Preserve text alignment (left, center, right, justified)

SPACING RHYTHM — All spacing follows a 4px grid:
- Use ONLY these values for gap, padding, margin: 4px, 8px, 12px, 16px, 20px, 24px
- gap: 8px between sections in the content zone
- gap: 4px between bullet items inside a section-box
- padding: 12px inside section-boxes and cards
- padding: 16px-20px for larger containers
- NEVER use arbitrary values like 7px, 13px, 18px — stick to the 4px rhythm

VISUAL HIERARCHY:
- Maintain clear levels: large titles → medium section headers → small body text
- Match proportions and relative positioning of every element
</design_principles>

<visual_polish>
PROFESSIONAL VISUAL ENRICHMENT — Apply these patterns to make slides look premium:

1. STATUS BADGES (pill-shaped) — For any status, phase, or category indicator:
   <span style="display:inline-block; padding: 2px 10px; border-radius: 12px;
     background:#dcfce7; color:#166534; font-size:10px; font-weight:600;">On Track</span>
   Color mapping:
   - Green pill: background:#dcfce7; color:#166534; (on track, completed, good)
   - Yellow pill: background:#fef9c3; color:#854d0e; (at risk, in progress, warning)
   - Red pill: background:#fee2e2; color:#991b1b; (delayed, blocked, critical)
   - Blue pill: background:#dbeafe; color:#1e40af; (planned, info, neutral)
   - Gray pill: background:#f3f4f6; color:#374151; (N/A, not started, unknown)

2. ELEVATION SHADOWS — ONLY when detected in the original image:
   - If the original slide shows shadows on cards or containers, replicate them faithfully
   - If the original slide has a FLAT design with no shadows, do NOT add any box-shadow
   - Do NOT invent shadows that are not visible in the source image
   - The .slide container itself may have a shadow for the page preview, but inner elements should only
     have shadows if the original image shows them

3. CHROME BARS (top-bar, footer-bar) — Replicate EXACTLY as seen in the image:
   - If the image shows a FLAT solid color bar → use a flat background: #hex; (NO gradient)
   - If the image shows a visible gradient → replicate it with linear-gradient
   - Do NOT add gradients to flat-colored bars — this changes the design intent
   - Most presentation footers are flat solid colors — keep them flat

4. BORDERS AS VISUAL LANGUAGE — Consistent border system:
   - Section header separator: border-top with the primary color (replicate thickness and color from image)
   - Section box border: replicate exactly from image (color, thickness, radius)
   - Table borders: replicate from image (typically hairline row separators)
   - NEVER mix border styles — all borders in a slide follow the same language
   - Do NOT add border-radius if the original image shows sharp/square corners
   - Only use border-radius when the original image clearly shows rounded corners

5. CSS MICRO-VISUALIZATIONS — For KPIs and metrics, use CSS-only visual elements:

   Donut chart (for percentage completion):
   <div style="width:48px; height:48px; border-radius:50%;
     background: conic-gradient([primary] 0% 75%, #e5e7eb 75% 100%);
     display:flex; align-items:center; justify-content:center;">
     <span style="width:32px; height:32px; border-radius:50%; background:white;
       display:flex; align-items:center; justify-content:center; font-size:11px; font-weight:700;">75%</span>
   </div>

   Horizontal comparison bar (actual vs planned):
   <div style="display:flex; flex-direction:column; gap:2px; width:100%;">
     <div style="height:6px; border-radius:3px; background:#e5e7eb; overflow:hidden;">
       <div style="height:100%; width:75%; background:[primary]; border-radius:3px;"></div>
     </div>
     <div style="display:flex; justify-content:space-between; font-size:9px; color:rgba(0,0,0,0.5);">
       <span>Actual: 75%</span><span>Target: 100%</span>
     </div>
   </div>

   Traffic light (3-state indicator):
   <div style="display:flex; gap:4px; align-items:center;">
     <span style="width:10px; height:10px; border-radius:50%; background:#ef4444; opacity:0.3;"></span>
     <span style="width:10px; height:10px; border-radius:50%; background:#eab308; opacity:0.3;"></span>
     <span style="width:10px; height:10px; border-radius:50%; background:#22c55e; opacity:1;"></span>
   </div>

   Use these ONLY when the original slide shows similar visual elements (charts, gauges, indicators).
   Do NOT add visualizations that are not present in the original image.

6. SECTION-BOX STYLING — Replicate what you see in the image:
   - If the original shows white/plain section boxes, keep them white/plain
   - If the original shows tinted backgrounds, replicate the exact tint color
   - If the original shows sharp/square corners, use border-radius: 0 (or omit it)
   - If the original shows rounded corners, replicate the exact radius
   - Do NOT add border-radius, shadows, or tinted backgrounds not visible in the source image
</visual_polish>

<layout_rules>
SLIDE DIMENSIONS AND LAYOUT:
- FIXED dimensions: 960px width × 540px height (16:9 aspect ratio)
- ALL elements must fit WITHIN the 960×540 container — NOTHING may overflow or be clipped
- Use overflow: hidden on each .slide container as a safety net
- Layer elements with z-index only when intentional overlap is needed

LAYOUT STRATEGY — FIT EVERYTHING FIRST:
The #1 priority is that ALL content is visible, readable, and nothing overlaps or gets cut off.
Use this layout approach:

1. CHROME (top-bar, footer-bar): Use position: absolute — they are fixed at top/bottom edges.
2. HEADER ZONE (title + date-box): Use a flex row container with position: absolute.
   - display: flex; align-items: center; justify-content: space-between;
   - This PREVENTS title/date-box overlap automatically.
3. CONTENT ZONE (sections between header and footer): Use a flex column container with position: absolute.
   - position: absolute; top: [below header]; left: 20px; right: 20px; bottom: [above footer];
   - display: flex; flex-direction: column; gap: 8px;
   - Sections FLOW vertically and share space — NO hardcoded top values needed.
   - If content is too tall, flex-shrink distributes the squeeze across sections.
4. INSIDE SECTIONS (.section-box): Use flex column for internal content.
   - display: flex; flex-direction: column; gap: 4px;
   - Bullets, paragraphs, and sub-sections flow naturally without overlapping.
5. TREND/KPI ROWS (.trend-box): Use flex row.
   - display: flex; gap: 12px; flex-wrap: wrap; align-items: center;
6. TWO-COLUMN LAYOUTS: Use flex row with flex: 1 per column.
   - display: flex; gap: 16px; — each column: flex: 1; min-width: 0;

WHEN TO USE position: absolute:
- .top-bar (top: 0), .footer-bar (bottom: 0) — fixed chrome
- The header zone container and content zone container — anchored to slide edges
- Individual elements that must overlap others (badges, floating indicators)

WHEN TO USE flexbox:
- Content flow inside the slide (sections stacking vertically)
- Items inside section-box (bullets, text blocks, lists)
- Rows of KPIs, status indicators, trend items
- Title + date-box alignment
- Two-column layouts
- Any place where elements need to share space without overlapping

SIZING RULES:
- Use px for widths on containers and font-sizes
- Use flex: 1 / flex-shrink for proportional sizing within flex containers
- Percentages allowed for progress bar fills (width: 75%)
- gap property is preferred over margins for spacing between siblings
</layout_rules>

<fit_everything>
CRITICAL — THE #1 PRIORITY IS THAT EVERYTHING FITS AND LOOKS CLEAN:
If you must choose between pixel-perfect replica and content visibility, ALWAYS choose visibility.

1. ALL text must be fully readable — never cut off, never overlapping adjacent elements
2. If the original slide image shows overlapping text, truncated content, or misaligned elements,
   FIX these issues in your HTML. Produce a clean, readable version — not a copy of visual bugs.
3. Use flex layouts to let content flow and share space naturally
4. Reduce font-size when content is dense (11px body, 10px tables) BEFORE it overflows
5. Use flex-shrink and min-height: 0 on flex children so they can compress when space is tight
6. Never leave empty whitespace while content elsewhere is being clipped
7. Test mentally: "if every text field had 2x the content, would flex handle it?" If not, add safeguards.
</fit_everything>

<structure_rules>
HTML STRUCTURE STANDARDS:

REQUIRED CSS CLASS NAMES (use these exact names):
- .top-bar — colored bar at the top of each slide
- .date-box — date/period box (in the header, right-aligned via flex)
- .main-title — slide title text (in the header, grows to fill available space)
- .footer-bar — footer bar at the bottom (NOT .bottom-bar, NOT .footer)
- .page-number — page number INSIDE .footer-bar
- .logo — logo text INSIDE .footer-bar
- .section-header — section header with border-top separator line
- .section-title — title text INSIDE .section-header
- .section-box — bordered content box (use flex column inside for content flow)
- .bullet-item — individual bullet point item (with colored square bullet via CSS ::before)
- .sub-label — bold sub-label/category within a section-box
- .trend-box / .trend-item — KPI/trend indicators row (use flex row)
- .link-text — styled hyperlinks

STRUCTURE RULES:

1. SECTION GROUPING — Each content section MUST keep .section-header and .section-box together
   as children of the SAME parent container:
   CORRECT:
     <div class="section">
       <div class="section-header"><span class="section-title">BUDGET</span></div>
       <div class="section-box">...content here...</div>
     </div>
   WRONG:
     <div class="section-header"><span class="section-title">BUDGET</span></div>
     <!-- gap or other elements -->
     <div class="section-box">...content...</div>

2. FOOTER NESTING — .page-number and .logo MUST be children of .footer-bar:
   CORRECT: <div class="footer-bar"><span class="page-number">1</span><span class="logo">Air</span></div>
   WRONG:   <div class="footer-bar"></div><div class="page-number">1</div>

3. FOOTER / HEADER TEXT PRESERVATION — CRITICAL:
   Brand names, logos, and company names in footer-bar and top-bar MUST be reproduced
   CHARACTER BY CHARACTER. Read the text from the image VERY CAREFULLY.
   - Zoom in mentally on the footer area — read EVERY letter left to right
   - Common errors: reading only the last letter(s) of a word, dropping leading characters
   - Example: if the footer shows "SYSTRA", you MUST output "SYSTRA" — not "A", not "TRA", not "STRA"
   - Example: if the footer shows "ACCENTURE", you MUST output "ACCENTURE" — not "E", not "URE"
   - The .logo span must contain the FULL, COMPLETE text — never a partial word
   - The .page-number span must contain the FULL page number text
   - After writing footer HTML, RE-READ the image footer and VERIFY every character matches
   - Apply the same care to .main-title, .date-box, and any text in .top-bar

4. TABLES — Use semantic <table><tr><th>/<td> for tabular data. Set column widths in px.

4. PROGRESS BARS — Nested divs: outer (background, border-radius) + inner (fill color, width in %).

CONTENT RICHNESS — Be CREATIVE with the visual design:
- Use colored progress bars, timeline indicators, trend arrows
- Create rich section content with bullet points, sub-labels, bold text
- Use colored status indicators (small spans with background-color and border-radius)
- Add visual separators, borders, and section dividers for professional look
- Every slide should feel dense with useful information, not empty
</structure_rules>

<text_handling>
CRITICAL — TEXT CONTAINERS MUST BE RESILIENT TO VARIABLE-LENGTH CONTENT:
This HTML template will later be populated with real project data by another agent.
Text that is "Project Alpha" in the original may become "Enterprise Cloud Migration Phase 2 - EMEA Region".
Every text container MUST handle longer content gracefully.

MANDATORY on ALL text-containing elements (.main-title, .section-title, .bullet-item, .sub-label, td, span, p):
- word-wrap: break-word; overflow-wrap: break-word; (break long words)
- DO NOT use overflow: hidden or text-overflow: ellipsis on text elements — text must NEVER be clipped or truncated
- If text is too long, REDUCE font-size until it fits — this is always preferred over cutting text

MANDATORY on ALL flex column containers (.section-box, content zone):
- min-height: 0; (allow flex children to shrink below content size)
- flex-shrink: 1; (participate in compression when space is tight)

MANDATORY on ALL flex row items (trend-items, header title+date, two-column children):
- min-width: 0; (allow flex children to shrink — prevents horizontal overflow)
- word-wrap: break-word; overflow-wrap: break-word; (wrap long words instead of clipping)

Font-size approach:
- Use appropriate line-height (typically 1.3-1.5) for readability
- Titles: set a font-size that works if the title were 2-3x longer
- Body text: 11-13px is safer than 14px+ for variable content
- Table cells: 10-12px to fit varying data widths

DO NOT use fixed heights on text containers unless absolutely necessary.
Let flex handle the sizing. DO NOT use overflow: hidden on text containers — text must always be fully visible.
The ONLY element that should have overflow: hidden is the .slide container itself (960x540).
</text_handling>

<list_formatting>
CRITICAL - PROPER HTML LISTS:
- For ANY bullet points or list items, you MUST use proper HTML structure:
  <ul>
    <li>First item</li>
    <li>Second item</li>
  </ul>
- NEVER use raw bullet characters (*, -, etc.) in plain text
- NEVER output lists as: "* Item 1 * Item 2" in a single paragraph
- Style lists with CSS:
  ul {{ list-style-type: disc; padding-left: 20px; margin: 10px 0; }}
  li {{ margin-bottom: 5px; }}
- For numbered lists, use <ol> with list-style-type: decimal
- For custom bullets, use ::before pseudo-element or list-style-image
- Nested lists should be properly indented with nested <ul>/<ol> elements
</list_formatting>

<boxes_and_containers>
When replicating boxes, cards, or bordered containers:
- Content must be completely INSIDE the border with proper padding
- Use padding from the 4px rhythm (8px, 12px, 16px, 20px) to separate content from borders
- For lists inside boxes: account for bullet width + text width
- Replicate border-radius ONLY if the original image shows rounded corners — do NOT round square corners
- Replicate box-shadow ONLY if the original image shows shadows — do NOT add shadows to flat designs
- Match background colors precisely — do NOT add tints or gradients not in the original
- The rule is simple: if you can see it in the image, replicate it. If you cannot, do not add it.
</boxes_and_containers>

<tables_and_grids>
For tabular data:
- Use semantic HTML <table>, <tr>, <th>, <td> elements
- Set column widths in PIXELS directly on <td>/<th> elements via inline style (e.g., style="width: 240px;")
- Replicate header styling (background color, font weight, borders)
- Alternate row colors if present in the original
- Match cell padding and text alignment
- Set table width in pixels via inline style

For non-tabular grid layouts (cards, KPI grids, project overview grids):
- Use display: flex with flex-wrap: wrap and gap
- Each card/item: flex: 0 0 auto with a fixed width, or flex: 1 for equal distribution
- This is more flexible and prevents overflow compared to tables
</tables_and_grids>

<quality_checklist>
Before finalizing, verify each slide against this checklist (in priority order):
1. ✓ ALL content fits — nothing overflows, nothing is cut off, nothing overlaps
2. ✓ Flex containers are used for content flow — sections stack via flex column, not hardcoded top values
3. ✓ All text is visible and fully readable at a comfortable font size
4. ✓ Color palette is consistent — same primary/secondary hex used across all elements (no random color variations)
5. ✓ Spacing follows 4px rhythm — all gaps/padding are multiples of 4px
6. ✓ Status indicators use pill badges with appropriate color (green/yellow/red/blue/gray)
7. ✓ Shadows ONLY present if the original image shows them — no invented shadows on flat designs
8. ✓ Chrome bars (top-bar, footer-bar) match the image: flat color if flat, gradient only if visible
9. ✓ Typography system applied — titles tight (letter-spacing: -0.5px), headers uppercase with spacing
10. ✓ All text matches EXACTLY what is shown in the original slides
11. ✓ Footer/header text verified CHARACTER BY CHARACTER — logo, brand name, page number are COMPLETE (not truncated)
12. ✓ Lists and bullets are properly formatted with clean characters
13. ✓ No garbled Unicode characters (â–ª, â€", etc.)
</quality_checklist>

<output_format>
Generate a complete, valid HTML5 document with:
- <!DOCTYPE html>
- <html lang="en">
- <head> with <meta charset="UTF-8"> and <style> tag
- <body> with background: #1a1a1a and padding: 20px
- Each slide as: <div class="slide" data-slide-number="N">
</output_format>"""


FINAL_INSTRUCTIONS = """

<final_instructions>
Generate the complete HTML document that perfectly replicates each slide shown above.

REQUIRED BASE CSS:
```css
* { box-sizing: border-box; margin: 0; padding: 0; }
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

/* Container query auto-scaling for text */
.section-box, .trend-box, .section-header {
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
}
```

REMEMBER:
1. Replicate ALL text EXACTLY as shown - do NOT use placeholders
2. Use clean ASCII/HTML entities for bullets and special characters
3. Ensure pixel-perfect accuracy in positioning and styling
4. The output must be production-ready HTML

Your response must be professional, high-quality HTML that is visually IDENTICAL to the original images.
</final_instructions>"""


def build_long_text_instructions(strategy: str = 'summarize') -> str:
    """
    Build instructions for handling long text based on user's selected strategy.

    Args:
        strategy: 'summarize' | 'ellipsis' | 'omit'

    Returns:
        String with long text handling instructions for the prompt
    """
    base_intro = """<long_text_handling>
IMPORTANT: The user has selected a specific strategy for handling long text content.
When you encounter text fields that contain lengthy content (descriptions, notes, comments, etc.),
you MUST apply this strategy:
"""

    if strategy == 'summarize':
        return base_intro + """
STRATEGY: SUMMARIZE
- For long descriptions, project notes, or detailed text: CREATE A CONCISE SUMMARY
- Keep the essential meaning but reduce to 1-2 sentences max
- Focus on key points, outcomes, and important metrics
- Example: A 500-word project description → "Cloud migration project focused on reducing infrastructure costs by 40% while improving system reliability."
- Preserve critical data like numbers, dates, and status information
- The summary should fit comfortably in the designated space without overflow
</long_text_handling>"""

    elif strategy == 'ellipsis':
        return base_intro + """
STRATEGY: TRUNCATE WITH ELLIPSIS
- For long text: TRUNCATE and add "..." at the end
- Cut the text at a natural break point (end of word/sentence) that fits the container
- Always end truncated text with "..."
- Example: "This is a very long project description that..."
- Ensure the truncated text + ellipsis fits within the element's boundaries
- Do NOT summarize or paraphrase - just cut and add ellipsis
- Preserve the beginning of the text as-is
</long_text_handling>"""

    elif strategy == 'omit':
        return base_intro + """
STRATEGY: OMIT LONG TEXT
- For long text fields: REPLACE with a short placeholder or leave minimal content
- Use placeholders like "-", "See details", or "N/A" for lengthy content
- Keep only short, essential text (titles, status, dates, numbers)
- Example: A long description field → "-" or "Details available"
- This keeps the slides clean and focused on key metrics
- Short text (under ~50 characters) can remain as-is
</long_text_handling>"""

    else:
        # Default to summarize if unknown strategy
        return build_long_text_instructions('summarize')


def build_field_instructions(mapping_json: Dict[str, Any] = None) -> str:
    """
    Build dynamic field instructions based on user's mapping configuration.

    Args:
        mapping_json: The user's mapping configuration

    Returns:
        String with field naming instructions for the prompt
    """
    if not mapping_json:
        # Fallback to generic instructions if no mapping provided
        return """Use descriptive field names in snake_case:
- "Project Alpha" → {{project_name}}
- "John Smith" → {{owner_name}} or {{manager_name}}
- "2024-01-15" → {{start_date}} or {{end_date}}
- "85%" → {{progress_percentage}} or {{completion_rate}}
- "In Progress" → {{status}} or {{project_status}}
- "$150,000" → {{budget_amount}} or {{total_budget}}
- "Q1 2024" → {{quarter}} or {{period}}
- List items → {{item_1_name}}, {{item_2_name}}, {{item_3_name}}, etc.
- Milestone names → {{milestone_1}}, {{milestone_2}}, etc."""

    # Build instructions from the user's mapping
    instructions = ["USE EXACTLY THESE FIELD NAMES from the user's mapping configuration:"]
    instructions.append("")

    for field_name, field_config in mapping_json.items():
        if isinstance(field_config, dict):
            data_path = field_config.get('path', field_config.get('source', 'unknown'))
            description = field_config.get('description', '')
        else:
            data_path = str(field_config)
            description = ''

        example = f"  - {{{{{{field_name}}}}}} → maps to: {data_path}"
        if description:
            example += f" ({description})"
        instructions.append(example)

    instructions.append("")
    instructions.append("For any ADDITIONAL dynamic content not in the mapping above,")
    instructions.append("use descriptive snake_case names like: {{additional_field_1}}, {{metric_value}}, etc.")
    instructions.append("")
    instructions.append("IMPORTANT: The field names MUST match the mapping exactly so data population works correctly.")

    return "\n".join(instructions)


def generate_html_template(
    images: List[Tuple[bytes, str]],
    mapping_json: Dict[str, Any] = None,
    long_text_strategy: str = 'summarize'
) -> Dict[str, Any]:
    """
    Use Claude Vision to generate an exact HTML replica from slide images.

    Args:
        images: List of (image_bytes, media_type) tuples
        mapping_json: Not used (kept for backward compatibility)
        long_text_strategy: Not used (kept for backward compatibility)

    Returns:
        Dictionary with 'full_html' key (exact replica, no placeholders)
    """
    content = []

    # Use the base prompt (no placeholder instructions)
    prompt_text = HTML_TEMPLATE_PROMPT_BASE

    # Add the main prompt
    content.append({
        "type": "text",
        "text": prompt_text
    })

    # Add each slide image
    for i, (img_bytes, media_type) in enumerate(images, 1):
        img_base64 = base64.b64encode(img_bytes).decode('utf-8')
        content.append({
            "type": "text",
            "text": f"\n--- SLIDE {i} of {len(images)} ---"
        })
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": media_type,
                "data": img_base64
            }
        })

    # Add final instructions
    content.append({
        "type": "text",
        "text": FINAL_INSTRUCTIONS
    })

    # Call Claude Opus 4.5 with structured output
    collected_text = ""

    with client.beta.messages.stream(
        model=CLAUDE_MODEL,  # claude-opus-4-5-20251101
        max_tokens=CLAUDE_MAX_TOKENS,
        temperature=0.2,  # Small amount of creativity for better visual interpretation
        betas=["structured-outputs-2025-11-13"],
        messages=[
            {
                "role": "user",
                "content": content
            }
        ],
        output_format={
            "type": "json_schema",
            "schema": {
                "type": "object",
                "properties": {
                    "full_html": {
                        "type": "string",
                        "description": "The complete HTML code that exactly replicates the slides, including DOCTYPE, head with styles, and body with all slides. Use clean ASCII characters for bullets."
                    }
                },
                "required": ["full_html"],
                "additionalProperties": False
            }
        }
    ) as stream:
        for text in stream.text_stream:
            collected_text += text
            print(".", end="", flush=True)

    print()  # Newline after progress dots

    result = json.loads(collected_text)

    # Post-process to fix any remaining character encoding issues
    if "full_html" in result:
        result["full_html"] = fix_character_encoding(result["full_html"])

    # Add empty fields for backward compatibility
    result["fields"] = []

    return result


def fix_character_encoding(html: str) -> str:
    """
    Fix common character encoding issues in generated HTML.
    Replaces garbled UTF-8 characters with clean ASCII/HTML entities.
    """
    result = html

    # FIRST: Fix mojibake (UTF-8 interpreted as Latin-1/Windows-1252)
    # Using Unicode escape sequences to avoid syntax errors
    mojibake_fixes = [
        # Bullets - mojibake patterns
        ("\u00e2\u20ac\u00a2", "*"),      # â€¢ -> • bullet
        ("\u00e2\u0096\u00aa", "*"),      # â–ª -> ▪ small square
        ("\u00e2\u0097\u00a6", "*"),      # â—¦ -> ◦ white bullet
        ("\u00e2\u0097\u2039", "*"),      # â—‹ -> ○ white circle
        ("\u00e2\u0097", "*"),            # â— -> ● black circle prefix
        # Dashes - mojibake patterns
        ("\u00e2\u20ac\u201c", "-"),      # â€" -> – en-dash
        ("\u00e2\u20ac\u201d", "-"),      # â€" -> — em-dash (different encoding)
        # Quotes - mojibake patterns
        ("\u00e2\u20ac\u02dc", "'"),      # â€˜ -> ' left single quote
        ("\u00e2\u20ac\u2122", "'"),      # â€™ -> ' right single quote
        ("\u00e2\u20ac\u0153", '"'),      # â€œ -> " left double quote
        ("\u00e2\u20ac\u009d", '"'),      # â€ -> " right double quote
        # Arrows - mojibake patterns
        ("\u00e2\u2020\u2019", "->"),     # â†' -> → right arrow
        ("\u00e2\u2020\u0090", "<-"),     # â† -> ← left arrow
        # Spaces - mojibake
        ("\u00c2\u00a0", " "),            # Â  -> non-breaking space
        # French/Spanish accents - mojibake (Ã + second byte)
        ("\u00c3\u00a9", "e"),            # Ã© -> é
        ("\u00c3\u00a8", "e"),            # Ã¨ -> è
        ("\u00c3\u00aa", "e"),            # Ãª -> ê
        ("\u00c3\u00a0", "a"),            # Ã  -> à
        ("\u00c3\u00a2", "a"),            # Ã¢ -> â
        ("\u00c3\u00a1", "a"),            # Ã¡ -> á
        ("\u00c3\u00ae", "i"),            # Ã® -> î
        ("\u00c3\u00af", "i"),            # Ã¯ -> ï
        ("\u00c3\u00ad", "i"),            # Ã­ -> í
        ("\u00c3\u00b4", "o"),            # Ã´ -> ô
        ("\u00c3\u00b3", "o"),            # Ã³ -> ó
        ("\u00c3\u00b9", "u"),            # Ã¹ -> ù
        ("\u00c3\u00bb", "u"),            # Ã» -> û
        ("\u00c3\u00ba", "u"),            # Ãº -> ú
        ("\u00c3\u00bc", "u"),            # Ã¼ -> ü
        ("\u00c3\u00a7", "c"),            # Ã§ -> ç
        ("\u00c3\u00b1", "n"),            # Ã± -> ñ
        ("\u00c3\u00a4", "a"),            # Ã¤ -> ä
        ("\u00c3\u00b6", "o"),            # Ã¶ -> ö
        ("\u00c5\u0093", "oe"),           # Å" -> œ
        ("\u00c3\u0178", "ss"),           # ÃŸ -> ß
    ]

    for bad, good in mojibake_fixes:
        result = result.replace(bad, good)

    # SECOND: Replace Unicode characters with ASCII equivalents
    # Bullets
    result = result.replace("\u2022", "*")  # bullet •
    result = result.replace("\u25aa", "*")  # small square ▪
    result = result.replace("\u25cf", "*")  # black circle ●
    result = result.replace("\u2023", "*")  # triangular bullet ‣
    result = result.replace("\u2043", "-")  # hyphen bullet ⁃
    result = result.replace("\u25e6", "*")  # white bullet ◦

    # Dashes
    result = result.replace("\u2013", "-")  # en-dash –
    result = result.replace("\u2014", "-")  # em-dash —
    result = result.replace("\u2015", "-")  # horizontal bar ―

    # Quotes
    result = result.replace("\u2018", "'")  # left single quote '
    result = result.replace("\u2019", "'")  # right single quote '
    result = result.replace("\u201c", '"')  # left double quote "
    result = result.replace("\u201d", '"')  # right double quote "
    result = result.replace("\u201a", ",")  # single low quote ‚
    result = result.replace("\u201e", '"')  # double low quote „

    # Arrows
    result = result.replace("\u2192", "->")  # right arrow →
    result = result.replace("\u2190", "<-")  # left arrow ←
    result = result.replace("\u2191", "^")   # up arrow ↑
    result = result.replace("\u2193", "v")   # down arrow ↓

    # Spaces
    result = result.replace("\u00a0", " ")   # non-breaking space
    result = result.replace("\u202f", " ")   # narrow no-break space

    # Checkmarks and crosses
    result = result.replace("\u2713", "[x]")  # check mark ✓
    result = result.replace("\u2714", "[x]")  # heavy check mark ✔
    result = result.replace("\u2717", "[ ]")  # ballot x ✗
    result = result.replace("\u2718", "[ ]")  # heavy ballot x ✘

    return result


def extract_template_fields(html_content: str) -> List[dict]:
    """
    Extract all {{field_name}} placeholders from the HTML template.

    Returns:
        List of field dictionaries with name and context
    """
    import re

    fields = []
    pattern = r'\{\{(\w+)\}\}'

    # Find all unique field names
    matches = set(re.findall(pattern, html_content))

    for field_name in matches:
        fields.append({
            "field_name": field_name,
            "placeholder": f"{{{{{field_name}}}}}"
        })

    return fields
