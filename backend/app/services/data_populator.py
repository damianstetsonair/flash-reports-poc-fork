"""
Data Population Service

Populates HTML templates with actual project data using Claude to:
1. Understand the mapping between template fields and project data
2. Generate the populated HTML with correct data placement
3. Creatively expand templates for multiple projects while preserving design
"""

import anthropic
import json
import re
from typing import Dict, List, Any, Optional

from app.config import ANTHROPIC_API_KEY, CLAUDE_MODEL, CLAUDE_MAX_TOKENS


# Initialize Anthropic client
client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


def fix_mojibake(text: str) -> str:
    """
    Fix mojibake (UTF-8 misinterpreted as Latin-1/Windows-1252) in text.
    This is common when data comes from external APIs with encoding issues.
    """
    if not isinstance(text, str):
        return text

    # Common mojibake patterns - using Unicode escape sequences
    fixes = [
        # Bullets - mojibake patterns
        ("\u00e2\u20ac\u00a2", "*"),      # â€¢ -> • bullet
        ("\u00e2\u0096\u00aa", "*"),      # â–ª -> ▪ small square
        ("\u00e2\u0097\u00a6", "*"),      # â—¦ -> ◦ white bullet
        ("\u00e2\u0097\u2039", "*"),      # â—‹ -> ○ white circle
        ("\u00e2\u0097", "*"),            # â— -> ● black circle prefix
        # Dashes - mojibake patterns
        ("\u00e2\u20ac\u201c", "-"),      # â€" -> – en-dash
        ("\u00e2\u20ac\u201d", "-"),      # â€" -> — em-dash
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
        # Spanish punctuation
        ("\u00c2\u00bf", "?"),            # Â¿ -> ¿
        ("\u00c2\u00a1", "!"),            # Â¡ -> ¡
    ]

    result = text
    for bad, good in fixes:
        result = result.replace(bad, good)

    return result


def clean_project_data(data: Any) -> Any:
    """
    Recursively clean mojibake from project data (dicts, lists, strings).
    """
    if isinstance(data, str):
        return fix_mojibake(data)
    elif isinstance(data, dict):
        return {k: clean_project_data(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [clean_project_data(item) for item in data]
    else:
        return data


def get_nested_value(data: Dict[str, Any], path: str) -> Any:
    """
    Get a value from nested dictionary using dot notation path.

    Examples:
        get_nested_value(data, "project.name") → data["project"]["name"]
        get_nested_value(data, "milestones") → data["milestones"]
    """
    keys = path.split('.')
    value = data

    for key in keys:
        if isinstance(value, dict):
            value = value.get(key)
        elif isinstance(value, list) and key.isdigit():
            idx = int(key)
            value = value[idx] if idx < len(value) else None
        else:
            return None

        if value is None:
            return None

    return value


def apply_mapping_to_project(
    project_data: Dict[str, Any],
    mapping_json: Dict[str, Any],
    template_fields: List[str]
) -> Dict[str, str]:
    """
    Apply the mapping configuration to extract values for each template field.

    Args:
        project_data: The fetched project data from AirSaas
        mapping_json: The mapping configuration from the mapping step
        template_fields: List of field names found in the HTML template

    Returns:
        Dictionary mapping field_name → actual_value
    """
    field_values = {}

    # The mapping_json structure from mapping-batch-submit:
    # {
    #   "slides": {
    #     "slide_1": {
    #       "field_id": { "source": "project.name", "status": "ok" }
    #     }
    #   },
    #   "missing_fields": []
    # }

    # Build a flat mapping from field_id to source path
    field_to_source = {}

    if "slides" in mapping_json:
        for slide_key, slide_fields in mapping_json["slides"].items():
            for field_id, field_config in slide_fields.items():
                if isinstance(field_config, dict) and field_config.get("source"):
                    field_to_source[field_id] = field_config["source"]

    # For each template field, try to find its value
    for field_name in template_fields:
        # Check if we have a mapping for this field
        source_path = field_to_source.get(field_name)

        if source_path and source_path != "none":
            value = get_nested_value(project_data, source_path)
            if value is not None:
                # Convert value to string representation
                if isinstance(value, list):
                    # For arrays, format as bullet points or comma-separated
                    field_values[field_name] = format_array_value(value)
                elif isinstance(value, dict):
                    # For objects, use a sensible string representation
                    field_values[field_name] = format_dict_value(value)
                else:
                    field_values[field_name] = str(value)
            else:
                field_values[field_name] = ""
        else:
            # No mapping found, leave empty or use placeholder
            field_values[field_name] = ""

    return field_values


def format_array_value(arr: List[Any], max_items: int = 5) -> str:
    """Format an array value for display."""
    if not arr:
        return ""

    items = arr[:max_items]
    formatted = []

    for item in items:
        if isinstance(item, dict):
            # Try common name fields
            name = item.get("name") or item.get("title") or item.get("label")
            if name:
                formatted.append(str(name))
            else:
                formatted.append(str(item))
        else:
            formatted.append(str(item))

    return ", ".join(formatted)


def format_dict_value(d: Dict[str, Any]) -> str:
    """Format a dictionary value for display."""
    # Try common name fields
    name = d.get("name") or d.get("title") or d.get("label") or d.get("full_name")
    if name:
        return str(name)

    # Fallback to first string value
    for key, value in d.items():
        if isinstance(value, str):
            return value

    return str(d)


def simple_populate_html(html_template: str, field_values: Dict[str, str]) -> str:
    """
    Simple string replacement to populate HTML template.

    Replaces all {{field_name}} with corresponding values.
    """
    result = html_template

    for field_name, value in field_values.items():
        placeholder = f"{{{{{field_name}}}}}"
        result = result.replace(placeholder, value or "")

    # Remove any remaining unmatched placeholders
    result = re.sub(r'\{\{[\w_]+\}\}', '', result)

    return result


# Advanced prompt for intelligent HTML population
POPULATION_PROMPT = """<role>
You are an expert presentation designer who populates HTML slide templates with project data.
You create professional, visually balanced presentations that effectively communicate project information.
</role>

<css_rules>
CSS AND LAYOUT RULES:
1. PRESERVE the template's <style> block — copy it exactly
2. You MAY add flex properties via inline styles to IMPROVE layout (prevent overlaps, fit content)
3. Preserve ALL existing class names: .top-bar, .date-box, .main-title,
   .footer-bar, .page-number, .logo, .section-header, .section-title, .section-box,
   .bullet-item, .sub-label, .trend-box, .trend-item, .link-text
4. FOOTER — .page-number and .logo MUST stay inside .footer-bar as children
5. SECTION GROUPING — .section-header and .section-box MUST be children of the same parent
6. TABLES — Keep as <table><tr><td> for tabular data

FLEX LAYOUT IMPROVEMENTS — You are ENCOURAGED to add flex to improve content flow:
- If the template uses position: absolute for content sections and they risk overlapping,
  convert the content zone to display: flex; flex-direction: column; gap: 8px;
- Inside .section-box: use display: flex; flex-direction: column; gap: 4px; for content flow
- For .trend-box: use display: flex; gap: 12px; flex-wrap: wrap;
- For title + date-box: use display: flex; justify-content: space-between; align-items: center;
- Use flex-shrink: 1 and min-height: 0 on sections so they compress when space is tight
- Keep .top-bar and .footer-bar as position: absolute (fixed chrome)
</css_rules>

<fit_everything>
CRITICAL — THE #1 PRIORITY IS THAT ALL CONTENT FITS AND IS READABLE:
1. ALL text must be visible — never cut off, never overlapping adjacent elements
2. Use flex layouts to let content flow and share space naturally
3. Reduce font-size when content is dense (11px body, 10px tables) BEFORE overflow occurs
4. Use flex-shrink and min-height: 0 on flex children so they compress when space is tight
5. Use word-wrap: break-word; overflow-wrap: break-word; on text containers to wrap long words
6. NEVER let text from one element visually overlap or cover text from another
7. If the text is 2x longer than the template placeholder, the layout MUST still hold — flex handles this
8. NEVER use overflow: hidden on text containers — text must ALWAYS be fully visible, never clipped
9. The ONLY element that should have overflow: hidden is the .slide container itself (960x540)
</fit_everything>

<icon_safety>
ICONS AND SPECIAL CHARACTERS - CRITICAL:
1. ONLY use basic ASCII characters and standard HTML entities
2. DO NOT use emoji unicode characters (they render inconsistently across systems)
3. DO NOT use icon fonts (FontAwesome, Material Icons, etc.) unless already in the template
4. For status indicators use simple text or HTML symbols:
   - Good: "OK", "Yes", "No", "-", "N/A", "&#9679;" (bullet), "&#9650;" (triangle up), "&#9660;" (triangle down)
   - Good: "&#10003;" (checkmark), "&#10005;" (cross), "&#9733;" (star)
   - BAD: emoji like 🟢🔴⚠️🎯✅❌ (these may render as broken/null characters in PDF)
5. For colored indicators, use a <span> with background-color and border-radius instead of emoji:
   Example: <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#22c55e;"></span>
6. NEVER use characters from Private Use Area (U+E000–U+F8FF) or rare Unicode blocks
</icon_safety>

<what_you_CAN_do>
You ARE ALLOWED and ENCOURAGED to make these adjustments:
1. Add display: flex (column or row) to content containers to prevent overlaps
2. Add gap, flex-shrink, min-height: 0 for proper spacing and compression
3. Reduce font-size on long titles/text so everything fits — this is the PRIMARY strategy for long text
4. Add word-wrap: break-word; overflow-wrap: break-word; on text containers
5. DO NOT use overflow: hidden or text-overflow: ellipsis on text elements — text must never be cut
6. Adjust line-height if text is cramped
7. Convert absolute-positioned content sections to flex layout if it prevents overlaps
8. Keep .top-bar and .footer-bar as position: absolute — do not change chrome elements
</what_you_CAN_do>

<visual_polish>
VISUAL ENRICHMENT — Apply when populating data to make slides look premium:

1. STATUS VALUES → render as PILL BADGES (not plain text):
   <span style="display:inline-block; padding:2px 10px; border-radius:12px;
     background:#dcfce7; color:#166534; font-size:10px; font-weight:600;">On Track</span>
   Colors: green=good/on-track, yellow=at-risk/warning, red=delayed/critical, blue=info/planned, gray=N/A

2. PERCENTAGE VALUES → consider donut mini-charts when space allows:
   <div style="width:44px; height:44px; border-radius:50%;
     background:conic-gradient([primary] 0% {pct}%, #e5e7eb {pct}% 100%);
     display:flex; align-items:center; justify-content:center;">
     <span style="width:30px; height:30px; border-radius:50%; background:white;
       display:flex; align-items:center; justify-content:center; font-size:10px; font-weight:700;">{pct}%</span>
   </div>
   Use these for key metrics like project completion — not for every percentage.

3. SPACING — All gaps and padding must follow 4px rhythm: 4, 8, 12, 16, 20, 24px only

4. SECTION-BOX STYLING — Preserve the template's original look:
   - If the template has flat boxes with no shadow, keep them flat — do NOT add box-shadow
   - If the template has shadows, preserve them exactly
   - If the template has tinted backgrounds, keep them; if white/plain, keep them white/plain
   - Do NOT invent shadows, gradients, or tinted backgrounds not present in the template

5. TYPOGRAPHY when inserting data:
   - Captions and dates: color: rgba(0,0,0,0.55); (muted, not black)
   - Values/numbers: font-weight: 600; (semi-bold to stand out)
   - Keep consistent with the template's type hierarchy
</visual_polish>

<html_template>
{html_template}
</html_template>

<template_fields>
{template_fields}
</template_fields>

<project_data>
```json
{project_data}
```
</project_data>

<mapping_configuration>
```json
{mapping_json}
```
</mapping_configuration>

<slide_structure_spec>
REPORT STRUCTURE - MANDATORY REQUIREMENTS:

═══════════════════════════════════════════════════════════════════════════════
CRITICAL: YOU MUST FOLLOW THIS EXACT STRUCTURE. NO EXCEPTIONS.
═══════════════════════════════════════════════════════════════════════════════

FOR SINGLE PROJECT:
- The template defines the slide types (Card, Progress, Planning, etc.)
- Populate ALL slides with the project's data

FOR MULTI-PROJECT REPORTS (2+ projects):
════════════════════════════════════════
THIS STRUCTURE IS MANDATORY AND NON-NEGOTIABLE:

**SLIDE 1 - PORTFOLIO OVERVIEW (REQUIRED - MUST BE FIRST)**
- This slide MUST exist and MUST be the first slide
- Shows a summary table/grid of ALL projects at a glance
- Include for each project: Name, Status indicator (color/icon), Mood/Weather, Progress %
- Use a compact table or grid layout to fit all projects
- This is an EXECUTIVE SUMMARY - gives leadership a quick view of portfolio health
- Example format:
  | Project Name | Status | Mood | Progress |
  |--------------|--------|------|----------|
  | Project A    | 🟢 On Track | ☀️ | 85% |
  | Project B    | 🟡 At Risk  | ⛅ | 60% |
  | Project C    | 🔴 Delayed  | 🌧️ | 40% |

**SLIDES 2 to N - INDIVIDUAL PROJECT SLIDES (REQUIRED)**
- Repeat the FULL template structure for EACH project
- Each project gets its own complete set of slides (Card, Progress, Planning, etc.)
- Maintain consistent formatting across all projects
- Number of slides per project = number of slides in the template

**FINAL SLIDE - DATA NOTES (REQUIRED - MUST BE LAST)**
- This slide MUST exist and MUST be the last slide
- Lists any fields that could not be populated due to missing data
- Include generation timestamp
- Example: "Report generated on Feb 3, 2026. Missing data: Budget for Project B, End date for Project C"

════════════════════════════════════════
FAILURE TO INCLUDE THESE 3 SECTIONS (Overview, Project Slides, Data Notes) IS UNACCEPTABLE.
════════════════════════════════════════

COMMON SLIDE TYPES (per project, as defined by template):
1. PROJECT CARD: Name, Budget, Achievements, Status, Mood/Weather, Risk level, Key dates
2. PROGRESS SLIDE: Completion percentage, KPIs, Key metrics
3. PLANNING SLIDE: Milestones timeline, Team efforts, Resource allocation

DATA FIELD PRIORITIES (what to show when available):
- Project identification: name, short_id, program
- Status indicators: status, mood/weather, risk level
- Financial: budget (initial, current, EAC), expenses
- Progress: completion %, milestones status
- Timeline: start_date, end_date, key milestone dates
- People: owner, manager, team members
- Achievements: recent accomplishments, decisions made
- Risks/Issues: attention points, blockers
</slide_structure_spec>

<population_task>
1. COPY the entire HTML template structure exactly
2. For each {{field_name}} placeholder:
   a. Find the mapping: mapping_json tells you which data field to use
   b. Get the value from project_data using the mapped path
   c. Replace {{field_name}} with the actual value

3. TEXT FITTING - Critical:
   - If a title/text is too long for its container, REDUCE the font-size inline
   - Example: <span style="font-size: 14px;">Very Long Project Name Here</span>
   - Titles should NEVER overflow or get cut off
   - Text must NEVER be clipped or truncated — reduce font-size until all text is fully visible
   - DO NOT use "..." truncation or overflow: hidden on any text element

4. ABSOLUTELY NO EMPTY SLIDES OR BLANK FIELDS - THIS IS CRITICAL:
   - EVERY slide must have meaningful, visible content
   - EVERY text field must be filled with real data
   - If a {{field_name}} has no direct mapping, FIND relevant data from project_data to fill it
   - Look for related fields: name, title, description, status, dates, owner, budget, progress, etc.
   - If truly no data exists, use sensible placeholders like "N/A", "-", or "Not specified"
   - NEVER leave a visible text area empty or with just whitespace
   - A slide with blank content is UNACCEPTABLE - always populate with something meaningful
   - A slide with ONLY headers/footers and no section content is UNACCEPTABLE
   - Every section-box must contain visible content (bullet-items, text, tables, indicators)
   - Fill section-boxes with bullet-items, sub-labels, progress info, team data, timelines, etc.

5. INTELLIGENT DATA FILLING (when no direct mapping exists):
   - Analyze ALL available data in project_data
   - Match fields intelligently: "project_title" can fill a "name" placeholder
   - Use context: a "description" field can fill "summary", "overview", "details" placeholders
   - Dates: use start_date, end_date, created_at, updated_at as appropriate
   - Numbers: use budget, progress, completion_rate, etc.
   - Status fields: use status, phase, state interchangeably
   - Owner/Manager: use owner, manager, lead, responsible, assignee

6. DATA FORMATTING:
   - Dates: "Jan 15, 2024" or "15/01/2024"
   - Percentages: "85%" (include % symbol)
   - Currency: "$150,000" or "150,000 EUR"
   - Status: Capitalize properly ("In Progress", "Completed", "On Hold")
   - Numbers: Use thousand separators (1,500 not 1500)

7. LISTS AND BULLET POINTS - CRITICAL:
   - ALWAYS use proper HTML structure for lists: <ul><li>Item</li></ul>
   - NEVER output raw bullet characters like "* Item 1 * Item 2" in plain text
   - For milestones, tasks, or any list data, convert to proper HTML:
     WRONG: "* Milestone 1 * Milestone 2 * Milestone 3"
     CORRECT: <ul><li>Milestone 1</li><li>Milestone 2</li><li>Milestone 3</li></ul>
   - Style lists appropriately within their containers

8. VISUAL BALANCE:
   - Text should not overlap with other elements
   - Maintain readable spacing
   - Keep the professional look of the template
</population_task>

{long_text_strategy_instructions}

<output>
Return ONLY the complete populated HTML document.
- No explanations
- No markdown code blocks
- No ```html wrapper
- Just raw HTML starting with <!DOCTYPE html>
</output>"""


LONG_TEXT_STRATEGY_INSTRUCTIONS = {
    'summarize': """<long_text_strategy>
USER-SELECTED STRATEGY FOR LONG TEXT: **SUMMARIZE**
This is a USER CHOICE that you MUST respect - it overrides your own judgment about text length.

Rules:
1. ANY text field longer than 2 sentences MUST be condensed to a maximum of 2 sentences
2. Preserve the key meaning and most important information
3. Write in the same language as the original text
4. Do NOT simply truncate - actually summarize the content intelligently
5. This applies to ALL text fields: descriptions, achievements, comments, notes, attention points, etc.
6. Even if the text fits visually, still summarize it if it exceeds 2 sentences - the USER wants concise content
</long_text_strategy>""",

    'ellipsis': """<long_text_strategy>
USER-SELECTED STRATEGY FOR LONG TEXT: **TRUNCATE WITH ELLIPSIS**
This is a USER CHOICE that you MUST respect - it overrides your own judgment about text length.

Rules:
1. ANY text field longer than 100 characters MUST be cut at ~100 characters and end with "..."
2. Cut at a word boundary when possible (don't cut mid-word)
3. This applies to ALL text fields: descriptions, achievements, comments, notes, attention points, etc.
4. Do NOT summarize or rephrase - just cut the original text and add "..."
5. Even if the text fits visually, still truncate it if it exceeds 100 characters - the USER wants short content
</long_text_strategy>""",

    'omit': """<long_text_strategy>
USER-SELECTED STRATEGY FOR LONG TEXT: **OMIT**
This is a USER CHOICE that you MUST respect - it overrides your own judgment about text length.

Rules:
1. ANY text field longer than 100 characters MUST be replaced with "-" or left as "N/A"
2. Do NOT show the long text at all - the user explicitly chose to skip long content
3. Short text (under 100 characters) should still be shown normally
4. This applies to ALL text fields: descriptions, achievements, comments, notes, attention points, etc.
5. Even if the text fits visually, still omit it if it exceeds 100 characters - the USER wants to skip long content
</long_text_strategy>""",
}


def populate_html_with_claude(
    html_template: str,
    project_data: Dict[str, Any],
    mapping_json: Dict[str, Any],
    long_text_strategy: str = 'summarize'
) -> str:
    """
    Use Claude Opus 4.5 to intelligently populate the HTML template with project data.

    This approach is sophisticated - Claude understands the context and can:
    1. Handle complex data transformations
    2. Format data appropriately for each field type
    3. Handle missing data gracefully
    4. Maintain visual consistency
    """
    # Clean project data to fix any mojibake encoding issues
    cleaned_project_data = clean_project_data(project_data)

    # First, extract template fields
    template_fields = list(set(re.findall(r'\{\{(\w+)\}\}', html_template)))

    # Build the prompt
    strategy_instructions = LONG_TEXT_STRATEGY_INSTRUCTIONS.get(
        long_text_strategy, LONG_TEXT_STRATEGY_INSTRUCTIONS['summarize']
    )
    prompt = POPULATION_PROMPT.format(
        html_template=html_template,
        template_fields=json.dumps(template_fields, indent=2),
        project_data=json.dumps(cleaned_project_data, indent=2, ensure_ascii=False),
        mapping_json=json.dumps(mapping_json, indent=2),
        long_text_strategy_instructions=strategy_instructions
    )

    # Use Claude Opus 4.5 with streaming
    html_content = ""
    token_count = 0

    print(f"         [populate] Prompt size: {len(prompt)} chars")
    print(f"         [populate] Model: {CLAUDE_MODEL}, Max tokens: {CLAUDE_MAX_TOKENS}")
    print(f"         [populate] Starting Claude API call...", flush=True)

    import time as _time
    api_start = _time.time()

    with client.messages.stream(
        model=CLAUDE_MODEL,  # claude-opus-4-5-20251101
        max_tokens=CLAUDE_MAX_TOKENS,
        temperature=0.1,  # Minimal creativity for smart data presentation
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ]
    ) as stream:
        for text in stream.text_stream:
            html_content += text
            token_count += 1
            if token_count % 500 == 0:
                elapsed = _time.time() - api_start
                print(f"\n         [stream] {token_count} chunks, {elapsed:.1f}s elapsed, HTML: {len(html_content)} chars", flush=True)

    api_elapsed = _time.time() - api_start
    print(f"\n         [populate] Completed in {api_elapsed:.1f}s, chunks: {token_count}, HTML: {len(html_content)} chars", flush=True)

    # Clean up - extract just the HTML if wrapped in code blocks
    if "```html" in html_content:
        match = re.search(r'```html\s*([\s\S]*?)\s*```', html_content)
        if match:
            html_content = match.group(1)
    elif "```" in html_content:
        match = re.search(r'```\s*([\s\S]*?)\s*```', html_content)
        if match:
            html_content = match.group(1)

    # Fix any mojibake in the generated HTML output
    return fix_mojibake(html_content.strip())


# Advanced prompt for multi-project HTML generation
MULTI_PROJECT_PROMPT = """<role>
You are an expert presentation designer creating a multi-project report.
You will generate slides for MULTIPLE projects, each with the same professional design but different data.
Your presentations are visually polished, well-balanced, and effectively communicate project information.
</role>

<css_rules>
CSS AND LAYOUT RULES:
1. PRESERVE the template's <style> block — copy it exactly
2. You MAY add flex properties via inline styles to IMPROVE layout (prevent overlaps, fit content)
3. Preserve ALL existing class names: .top-bar, .date-box, .main-title,
   .footer-bar, .page-number, .logo, .section-header, .section-title, .section-box,
   .bullet-item, .sub-label, .trend-box, .trend-item, .link-text
4. FOOTER — .page-number and .logo MUST stay inside .footer-bar as children
5. SECTION GROUPING — .section-header and .section-box MUST be children of the same parent
6. TABLES — Keep as <table><tr><td> for tabular data

FLEX LAYOUT IMPROVEMENTS — You are ENCOURAGED to add flex to improve content flow:
- If the template uses position: absolute for content sections and they risk overlapping,
  convert the content zone to display: flex; flex-direction: column; gap: 8px;
- Inside .section-box: use display: flex; flex-direction: column; gap: 4px; for content flow
- For .trend-box: use display: flex; gap: 12px; flex-wrap: wrap;
- For title + date-box: use display: flex; justify-content: space-between; align-items: center;
- Use flex-shrink: 1 and min-height: 0 on sections so they compress when space is tight
- Keep .top-bar and .footer-bar as position: absolute (fixed chrome)
</css_rules>

<fit_everything>
CRITICAL — THE #1 PRIORITY IS THAT ALL CONTENT FITS AND IS READABLE:
1. ALL text must be visible — never cut off, never overlapping adjacent elements
2. Use flex layouts to let content flow and share space naturally
3. Reduce font-size when content is dense (11px body, 10px tables) BEFORE overflow occurs
4. Use flex-shrink and min-height: 0 on flex children so they compress when space is tight
5. Use word-wrap: break-word; overflow-wrap: break-word; on text containers to wrap long words
6. NEVER let text from one element visually overlap or cover text from another
7. If the text is 2x longer than the template placeholder, the layout MUST still hold — flex handles this
8. NEVER use overflow: hidden on text containers — text must ALWAYS be fully visible, never clipped
9. The ONLY element that should have overflow: hidden is the .slide container itself (960x540)
</fit_everything>

<icon_safety>
ICONS AND SPECIAL CHARACTERS - CRITICAL:
1. ONLY use basic ASCII characters and standard HTML entities
2. DO NOT use emoji unicode characters (they render inconsistently across systems)
3. DO NOT use icon fonts (FontAwesome, Material Icons, etc.) unless already in the template
4. For status indicators use simple text or HTML symbols:
   - Good: "OK", "Yes", "No", "-", "N/A", "&#9679;" (bullet), "&#9650;" (triangle up), "&#9660;" (triangle down)
   - Good: "&#10003;" (checkmark), "&#10005;" (cross), "&#9733;" (star)
   - BAD: emoji like 🟢🔴⚠️🎯✅❌ (these may render as broken/null characters in PDF)
5. For colored indicators, use a <span> with background-color and border-radius instead of emoji:
   Example: <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#22c55e;"></span>
6. NEVER use characters from Private Use Area (U+E000–U+F8FF) or rare Unicode blocks
</icon_safety>

<what_you_CAN_do>
You ARE ALLOWED and ENCOURAGED to make these adjustments:
1. Add display: flex (column or row) to content containers to prevent overlaps
2. Add gap, flex-shrink, min-height: 0 for proper spacing and compression
3. Reduce font-size on long titles/text so everything fits — this is the PRIMARY strategy for long text
4. Add word-wrap: break-word; overflow-wrap: break-word; on text containers
5. DO NOT use overflow: hidden or text-overflow: ellipsis on text elements — text must never be cut
6. Adjust line-height if text is cramped
7. Convert absolute-positioned content sections to flex layout if it prevents overlaps
8. Keep .top-bar and .footer-bar as position: absolute — do not change chrome elements
</what_you_CAN_do>

<visual_polish>
VISUAL ENRICHMENT — Apply when populating data to make slides look premium:

1. STATUS VALUES → render as PILL BADGES (not plain text):
   <span style="display:inline-block; padding:2px 10px; border-radius:12px;
     background:#dcfce7; color:#166534; font-size:10px; font-weight:600;">On Track</span>
   Colors: green=good/on-track, yellow=at-risk/warning, red=delayed/critical, blue=info/planned, gray=N/A

2. PERCENTAGE VALUES → consider donut mini-charts when space allows:
   <div style="width:44px; height:44px; border-radius:50%;
     background:conic-gradient([primary] 0% {pct}%, #e5e7eb {pct}% 100%);
     display:flex; align-items:center; justify-content:center;">
     <span style="width:30px; height:30px; border-radius:50%; background:white;
       display:flex; align-items:center; justify-content:center; font-size:10px; font-weight:700;">{pct}%</span>
   </div>
   Use these for key metrics like project completion — not for every percentage.

3. SPACING — All gaps and padding must follow 4px rhythm: 4, 8, 12, 16, 20, 24px only

4. SECTION-BOX STYLING — Preserve the template's original look:
   - If the template has flat boxes with no shadow, keep them flat — do NOT add box-shadow
   - If the template has shadows, preserve them exactly
   - If the template has tinted backgrounds, keep them; if white/plain, keep them white/plain
   - Do NOT invent shadows, gradients, or tinted backgrounds not present in the template

5. TYPOGRAPHY when inserting data:
   - Captions and dates: color: rgba(0,0,0,0.55); (muted, not black)
   - Values/numbers: font-weight: 600; (semi-bold to stand out)
   - Keep consistent with the template's type hierarchy

6. PORTFOLIO OVERVIEW SLIDE — Extra polish:
   - Use flex card grid (display:flex; flex-wrap:wrap; gap:8px) for project cards
   - Each card: pill badge for status, donut for progress, muted text for dates
   - If many projects (6+), reduce card size and font to fit all in 960x540
</visual_polish>

<original_template>
{html_template}
</original_template>

<projects_data>
Here is the data for ALL projects. Create slides for EACH project:
```json
{projects_data}
```
</projects_data>

<mapping_configuration>
This tells you which template field maps to which data path:
```json
{mapping_json}
```
</mapping_configuration>

<slide_structure_spec>
═══════════════════════════════════════════════════════════════════════════════
MULTI-PROJECT REPORT STRUCTURE - MANDATORY AND NON-NEGOTIABLE
═══════════════════════════════════════════════════════════════════════════════

YOU MUST GENERATE EXACTLY THIS STRUCTURE. FAILURE TO DO SO IS UNACCEPTABLE.

**SLIDE 1 - PORTFOLIO OVERVIEW (REQUIRED - MUST BE FIRST SLIDE)**
════════════════════════════════════════════════════════════════
- This slide MUST exist and MUST be the very first slide in the report
- Purpose: Executive summary giving leadership a quick view of all projects
- Content REQUIRED:
  * Table or card grid showing ALL projects at a glance
  * Each row/card must show: Project Name, Status (color indicator), Mood/Weather, Progress %
  * Use HTML <table> for few projects, or flex card grid for many projects
- Example flex card grid (preferred for 4+ projects):
  <div style="display: flex; flex-wrap: wrap; gap: 8px;">
    <div style="flex: 1 1 200px; border: 1px solid #e0e0e0; padding: 8px;">
      <strong>Project A</strong> <span style="color:#22c55e">On Track</span> 85%
    </div>
    <div style="flex: 1 1 200px; border: 1px solid #e0e0e0; padding: 8px;">
      <strong>Project B</strong> <span style="color:#eab308">At Risk</span> 60%
    </div>
  </div>
- DO NOT SKIP THIS SLIDE. IT IS MANDATORY.

**SLIDES 2 to N - INDIVIDUAL PROJECT SLIDES (REQUIRED)**
════════════════════════════════════════════════════════════════
- For EACH project in the data, generate a COMPLETE set of template slides
- If template has 3 slides, and there are 5 projects, generate 15 project slides (3 × 5)
- Each project gets the full template structure: Card, Progress, Planning, etc.
- Add data-project-index and data-project-name attributes to each slide
- Maintain consistent formatting across all projects

**FINAL SLIDE - DATA NOTES (REQUIRED - MUST BE LAST SLIDE)**
════════════════════════════════════════════════════════════════
- This slide MUST exist and MUST be the very last slide in the report
- Purpose: Transparency about data quality and gaps
- Content REQUIRED:
  * Generation timestamp (e.g., "Generated: Feb 3, 2026 at 15:30 UTC")
  * List of any fields that could not be populated
  * Which projects had missing or incomplete data
- Example: "Missing data: Budget not available for Project C, End date not set for Project B"
- If no data is missing, state: "All fields populated successfully"
- DO NOT SKIP THIS SLIDE. IT IS MANDATORY.

════════════════════════════════════════════════════════════════
CHECKLIST - YOUR OUTPUT MUST INCLUDE:
[ ] First slide = Portfolio Overview with ALL projects in a table/grid
[ ] Middle slides = Complete template set for EACH project
[ ] Last slide = Data Notes with timestamp and missing data list
════════════════════════════════════════════════════════════════

COMMON SLIDE TYPES (per project):
1. PROJECT CARD: Name, Budget, Achievements, Status, Mood/Weather, Risk, Dates
2. PROGRESS SLIDE: Completion %, KPIs, Metrics, Progress bars
3. PLANNING SLIDE: Milestones timeline, Team efforts table, Resource allocation

DATA PRIORITIES (populate these fields first):
- Identity: project name, short_id, program name
- Status: current status, mood/weather, risk level
- Financial: budget values (initial, current, EAC)
- Progress: completion percentage, milestone counts
- Timeline: start_date, end_date, next milestone
- People: owner name, manager, team size
- Key info: achievements, decisions, attention points
</slide_structure_spec>

<generation_task>
═══════════════════════════════════════════════════════════════════════════════
GENERATION STEPS - FOLLOW IN EXACT ORDER
═══════════════════════════════════════════════════════════════════════════════

1. COPY the <style> block from the template EXACTLY — you may ADD flex rules but not remove existing ones

2. **FIRST SLIDE - PORTFOLIO OVERVIEW (MANDATORY)**:
   - This MUST be the first slide in your output
   - Create a new slide (not from template) with class="slide portfolio-overview"
   - Include a table or flex card grid showing ALL projects:
     * Project Name, Status (color: green=On Track, yellow=At Risk, red=Delayed), Mood/Weather, Progress %
   - Use flex layout for the content zone: display: flex; flex-direction: column; to stack header + grid
   - For the project grid itself: display: flex; flex-wrap: wrap; gap: 8px; for cards
     or a <table> for a compact row-based view
   - Style it to match the template's look and feel
   - EVERYTHING must fit within 960x540 — reduce font-size if many projects
   - DO NOT SKIP THIS STEP

3. **PROJECT SLIDES (MANDATORY)**:
   For EACH PROJECT in projects_data, create a complete set of slides:
   a. Copy all slide <div>s from the template body
   b. Add attributes: data-project-index="N" data-project-name="Project Name"
   c. Replace ALL {{field_name}} placeholders with actual data from that project
   d. Use the mapping_configuration to find the correct data path for each field

4. **LAST SLIDE - DATA NOTES (MANDATORY)**:
   - This MUST be the last slide in your output
   - Create a new slide with class="slide data-notes"
   - Include:
     * Generation timestamp
     * List of any fields that could not be populated
     * Which projects had missing data (if any)
   - DO NOT SKIP THIS STEP

4. TEXT FITTING AND LAYOUT - Critical for each slide:
   - Use flex column for the content zone so sections flow without overlapping
   - Use flex-shrink: 1 + min-height: 0 on sections so they compress if space is tight
   - Long titles: Reduce font-size inline (e.g., style="font-size: 14px;")
   - Long descriptions: Reduce font-size further (down to 8px minimum) — NEVER truncate or clip text
   - Text must NEVER overlap other elements — flex layout prevents this
   - Text must NEVER be clipped or hidden — all injected text must be fully visible
   - DO NOT use overflow: hidden on text containers — only the .slide container should clip

5. ABSOLUTELY NO EMPTY SLIDES OR BLANK FIELDS - THIS IS CRITICAL:
   - EVERY slide for EVERY project must have meaningful, visible content
   - EVERY text field must be filled with real data
   - If a {{field_name}} has no direct mapping, FIND relevant data from that project to fill it
   - Look for related fields: name, title, description, status, dates, owner, budget, progress, etc.
   - If truly no data exists, use sensible placeholders like "N/A", "-", or "Not specified"
   - NEVER leave a visible text area empty or with just whitespace
   - A slide with blank content is UNACCEPTABLE - always populate with something meaningful
   - A slide with ONLY headers/footers and no section content is UNACCEPTABLE
   - Every section-box must contain visible content (bullet-items, text, tables, indicators)
   - Fill section-boxes with bullet-items, sub-labels, progress info, team data, timelines, etc.

6. INTELLIGENT DATA FILLING (when no direct mapping exists):
   - Analyze ALL available data in each project
   - Match fields intelligently: "project_title" can fill a "name" placeholder
   - Use context: a "description" field can fill "summary", "overview", "details" placeholders
   - Dates: use start_date, end_date, created_at, updated_at as appropriate
   - Numbers: use budget, progress, completion_rate, etc.
   - Status fields: use status, phase, state interchangeably
   - Owner/Manager: use owner, manager, lead, responsible, assignee

7. DATA FORMATTING:
   - Dates: "Jan 15, 2024" format
   - Percentages: "85%" (always include %)
   - Currency: "$150,000" or "EUR 150,000"
   - Status: "In Progress", "Completed", "On Hold" (capitalized)
   - Numbers: Use thousand separators

8. LISTS AND BULLET POINTS - CRITICAL:
   - ALWAYS use proper HTML structure for lists: <ul><li>Item</li></ul>
   - NEVER output raw bullet characters like "* Item 1 * Item 2" in plain text
   - For milestones, tasks, or any list data, convert to proper HTML:
     WRONG: "* Milestone 1 * Milestone 2 * Milestone 3"
     CORRECT: <ul><li>Milestone 1</li><li>Milestone 2</li><li>Milestone 3</li></ul>
   - Style lists appropriately within their containers

9. VISUAL QUALITY per slide:
   - Professional appearance
   - Readable text sizes (minimum 10px)
   - Proper spacing between elements
   - Consistent formatting across all projects
   - Each project's slides should look as polished as the template

10. CONTENT DECISIONS:
    - Analyze what data is available for each project
    - Present the most relevant and important information
    - If a project has more data than fits, prioritize key metrics
    - Ensure consistency in what data appears across all projects

11. CONTENT DECISIONS:
    - Analyze what data is available for each project
    - Present the most relevant and important information
    - If a project has more data than fits, prioritize key metrics
    - Ensure consistency in what data appears across all projects
</generation_task>

<output_structure>
YOUR OUTPUT MUST FOLLOW THIS EXACT STRUCTURE:

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Portfolio Flash Report</title>
    [EXACT COPY of template's <style> block]
</head>
<body>
    <!-- ═══════════════════════════════════════════════════════════ -->
    <!-- SLIDE 1: PORTFOLIO OVERVIEW (MANDATORY - MUST BE FIRST)    -->
    <!-- ═══════════════════════════════════════════════════════════ -->
    <div class="slide portfolio-overview" data-slide-number="1">
        <h1>Portfolio Overview</h1>
        <table>
            <tr><th>Project</th><th>Status</th><th>Mood</th><th>Progress</th></tr>
            <tr><td>Project A</td><td>On Track</td><td>Sunny</td><td>85%</td></tr>
            <!-- ... all projects ... -->
        </table>
    </div>

    <!-- ═══════════════════════════════════════════════════════════ -->
    <!-- SLIDES 2-N: INDIVIDUAL PROJECT SLIDES                       -->
    <!-- ═══════════════════════════════════════════════════════════ -->
    <!-- Project 1 slides -->
    <div class="slide" data-slide-number="2" data-project-index="0" data-project-name="Project A">...</div>
    <div class="slide" data-slide-number="3" data-project-index="0" data-project-name="Project A">...</div>

    <!-- Project 2 slides -->
    <div class="slide" data-slide-number="4" data-project-index="1" data-project-name="Project B">...</div>
    <div class="slide" data-slide-number="5" data-project-index="1" data-project-name="Project B">...</div>

    <!-- ... more projects ... -->

    <!-- ═══════════════════════════════════════════════════════════ -->
    <!-- FINAL SLIDE: DATA NOTES (MANDATORY - MUST BE LAST)          -->
    <!-- ═══════════════════════════════════════════════════════════ -->
    <div class="slide data-notes" data-slide-number="LAST">
        <h1>Data Notes</h1>
        <p>Report generated: [TIMESTAMP]</p>
        <h3>Missing Data:</h3>
        <ul>
            <li>Project B: Budget information not available</li>
            <li>Project C: End date not specified</li>
        </ul>
        <p>Or if no missing data: "All fields populated successfully."</p>
    </div>
</body>
</html>
</output_structure>

<final_checklist>
BEFORE RETURNING YOUR OUTPUT, VERIFY:
[ ] First slide is Portfolio Overview with ALL projects listed
[ ] Each project has a complete set of template slides
[ ] Last slide is Data Notes with timestamp and missing data info
[ ] All {{placeholders}} have been replaced with actual data
[ ] No empty slides or blank sections
</final_checklist>

<output>
Return ONLY the complete HTML document.
- No explanations or commentary
- No markdown code blocks (no ```html)
- Just raw HTML starting with <!DOCTYPE html>
The CSS must be IDENTICAL to the template. Only slide content changes.
</output>

{long_text_strategy_instructions}"""


def generate_multi_project_html(
    html_template: str,
    projects_data: List[Dict[str, Any]],
    mapping_json: Dict[str, Any],
    use_claude: bool = True,
    long_text_strategy: str = 'summarize'
) -> str:
    """
    Generate HTML with slides for multiple projects using Claude Opus 4.5.

    Args:
        html_template: The HTML template with placeholders
        projects_data: List of project data dictionaries
        mapping_json: The mapping configuration
        use_claude: Whether to use Claude for population (vs simple replacement)

    Returns:
        Complete HTML with slides for all projects
    """
    if not projects_data:
        return html_template

    if not use_claude:
        # Fallback to simple replacement for each project
        return _simple_multi_project_generation(html_template, projects_data, mapping_json)

    # Clean all project data to fix any mojibake encoding issues
    cleaned_projects_data = [clean_project_data(proj) for proj in projects_data]

    # Use Claude Opus 4.5 for intelligent multi-project generation
    strategy_instructions = LONG_TEXT_STRATEGY_INSTRUCTIONS.get(
        long_text_strategy, LONG_TEXT_STRATEGY_INSTRUCTIONS['summarize']
    )
    prompt = MULTI_PROJECT_PROMPT.format(
        html_template=html_template,
        projects_data=json.dumps(cleaned_projects_data, indent=2, ensure_ascii=False),
        mapping_json=json.dumps(mapping_json, indent=2),
        long_text_strategy_instructions=strategy_instructions
    )

    html_content = ""
    token_count = 0

    print(f"         Generating slides for {len(projects_data)} projects...")
    print(f"         Template HTML size: {len(html_template)} chars")
    print(f"         Projects data size: {len(json.dumps(cleaned_projects_data))} chars")
    print(f"         Total prompt size: {len(prompt)} chars")
    print(f"         Model: {CLAUDE_MODEL}, Max tokens: {CLAUDE_MAX_TOKENS}")
    print(f"         Starting Claude API call...", flush=True)

    import time as _time
    api_start = _time.time()

    with client.messages.stream(
        model=CLAUDE_MODEL,  # claude-opus-4-5-20251101
        max_tokens=CLAUDE_MAX_TOKENS,
        temperature=0.15,  # Slight creativity for smart data presentation
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ]
    ) as stream:
        for text in stream.text_stream:
            html_content += text
            token_count += 1
            if token_count % 500 == 0:
                elapsed = _time.time() - api_start
                print(f"\n         [stream] {token_count} chunks received, {elapsed:.1f}s elapsed, HTML size: {len(html_content)} chars", flush=True)

    api_elapsed = _time.time() - api_start
    print(f"\n         Claude API completed in {api_elapsed:.1f}s, total chunks: {token_count}, HTML size: {len(html_content)} chars", flush=True)

    # Clean up response
    if "```html" in html_content:
        match = re.search(r'```html\s*([\s\S]*?)\s*```', html_content)
        if match:
            html_content = match.group(1)
    elif "```" in html_content:
        match = re.search(r'```\s*([\s\S]*?)\s*```', html_content)
        if match:
            html_content = match.group(1)

    # Fix any mojibake in the generated HTML output
    return fix_mojibake(html_content.strip())


def _simple_multi_project_generation(
    html_template: str,
    projects_data: List[Dict[str, Any]],
    mapping_json: Dict[str, Any]
) -> str:
    """
    Simple fallback for multi-project generation without Claude.
    """
    template_fields = list(set(re.findall(r'\{\{(\w+)\}\}', html_template)))
    all_slides_html = []

    for idx, project_data in enumerate(projects_data):
        field_values = apply_mapping_to_project(project_data, mapping_json, template_fields)
        populated = simple_populate_html(html_template, field_values)

        # Extract slide content and add project attributes
        slide_match = re.search(r'<body[^>]*>([\s\S]*)</body>', populated)
        if slide_match:
            slides_content = slide_match.group(1)
            project_name = project_data.get("project", {}).get("name", f"Project {idx+1}")
            slides_content = re.sub(
                r'<div class="slide"',
                f'<div class="slide" data-project-index="{idx}" data-project-name="{project_name}"',
                slides_content
            )
            all_slides_html.append(slides_content)
        else:
            all_slides_html.append(populated)

    # Combine all slides
    combined_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Flash Report - {len(projects_data)} Projects</title>
    <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{
            background: #1a1a1a;
            padding: 20px;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
        }}
        .slide {{
            width: 960px;
            height: 540px;
            position: relative;
            overflow: hidden;
            margin: 20px auto;
            background: #ffffff;
        }}
        /* Container query auto-scaling for text */
        .section-box, .trend-box, .section-header {{
            container-type: inline-size;
        }}
        .section-box .bullet-item,
        .section-box .sub-label,
        .section-box p,
        .section-box li,
        .section-box span:not(.page-number):not(.logo) {{
            font-size: clamp(8px, 2.8cqw, 13px);
            line-height: clamp(1.1, 0.1cqw + 1, 1.5);
        }}
        .section-header .section-title {{
            font-size: clamp(9px, 3.2cqw, 16px);
        }}
        .trend-box .trend-item {{
            font-size: clamp(8px, 2.5cqw, 12px);
        }}
        td, th {{
            font-size: clamp(8px, 2.5cqw, 12px);
        }}
        .main-title {{
            font-size: clamp(14px, 3.5cqw, 26px);
        }}
        .project-divider {{
            text-align: center;
            color: #888;
            font-size: 14px;
            padding: 30px 0 10px;
        }}
    </style>
</head>
<body>
    {"".join(all_slides_html)}
</body>
</html>"""

    return combined_html
