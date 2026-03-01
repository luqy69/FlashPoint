"""
Content Organizer Module - 2-Stage Pipeline
Stage 1: Raw Gemini research text comes in
Stage 2: Send it back to Gemini with slide_formating_prompt to get structured SLIDE format
Stage 3: parse_slides() parses the structured output into slide data for ppt_builder

Pipeline:
    Raw Research Text -> slide_formating_prompt -> Gemini -> Structured SLIDE output -> parse_slides() -> slides list
    
Fallback: If Gemini formatting fails, use regex-based section detection
"""

import re
import os
import sys
from colorama import Fore, Style


def load_slide_formatting_prompt():
    """
    Load the slide_formating_prompt file using the same priority system as research prompts.
    Priority: config/ folder > root folder > bundled > dev directory
    """
    filename = 'slide_formating_prompt'
    
    # Priority 1: config/ folder next to executable
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        config_path = os.path.join(exe_dir, 'config', filename)
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    print(f"{Fore.CYAN}[*] Loaded config/{filename} (user-edited){Style.RESET_ALL}")
                    return content
            except Exception as e:
                print(f"{Fore.YELLOW}[!] Error reading config/{filename}: {e}{Style.RESET_ALL}")
    
    # Priority 2: Root directory next to executable (legacy)
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        legacy_path = os.path.join(exe_dir, filename)
        if os.path.exists(legacy_path):
            try:
                with open(legacy_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    print(f"{Fore.CYAN}[*] Loaded {filename} from root{Style.RESET_ALL}")
                    return content
            except Exception as e:
                print(f"{Fore.YELLOW}[!] Error reading {filename}: {e}{Style.RESET_ALL}")
    
    # Priority 3: Bundled files or dev directory
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
        prompt_path = os.path.join(base_dir, filename)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        base_dir = os.path.dirname(script_dir)  # Parent = project root
        prompt_path = os.path.join(base_dir, 'config', filename)
        
    try:
        with open(prompt_path, 'r', encoding='utf-8') as f:
            content = f.read()
            print(f"{Fore.GREEN}[+] Loaded {filename}{Style.RESET_ALL}")
            return content
    except Exception as e:
        print(f"{Fore.YELLOW}[!] Could not load {filename}: {e}{Style.RESET_ALL}")
    
    # Fallback: hardcoded minimal prompt
    return """ROLE: Act as a master academic lecturer converting content into PowerPoint slides.
TASK: Convert the following content into slide format.
STRICT OUTPUT FORMAT:
SLIDE 1
TYPE: Concept
TITLE: <title>
BULLETS:
- Bullet 1
- Bullet 2
RULES:
- Maximum 5 bullets per slide
- Each bullet under 12 words
- One core idea per slide
OUTPUT ONLY SLIDES IN THE SPECIFIED FORMAT."""


def parse_slides(raw_text):
    """
    Parses structured slide output from Gemini.
    Tries multiple parsing strategies since Gemini may use different formats.
    Returns a list of slide dictionaries.
    """
    if not raw_text or len(raw_text) < 50:
        return []
    
    # Strategy 1: Strict SLIDE format (SLIDE 1 / TYPE: / TITLE: / BULLETS:)
    slides = _parse_strict_slide_format(raw_text)
    if slides and len(slides) >= 2:
        print(f"    Parsed {len(slides)} slides using strict SLIDE format")
        return slides
    
    # Strategy 2: Markdown-style headers with bullet points (## Title + - bullets)
    slides = _parse_markdown_format(raw_text)
    if slides and len(slides) >= 2:
        print(f"    Parsed {len(slides)} slides using markdown format")
        return slides
    
    # Strategy 3: Numbered sections (1. Title\n- bullet)
    slides = _parse_numbered_format(raw_text)
    if slides and len(slides) >= 2:
        print(f"    Parsed {len(slides)} slides using numbered format")
        return slides
    
    # Strategy 4: Title/header detection with content chunks
    slides = _parse_generic_format(raw_text)
    if slides and len(slides) >= 2:
        print(f"    Parsed {len(slides)} slides using generic format")
        return slides
    
    return []


def _parse_strict_slide_format(text):
    """Parse structured slide format from Gemini output.
    Handles all output format variants:
      - STANDARD: BODY: / - bullets
      - MECHANISM: SEQUENCE: / STEP N →
      - CLASSIFICATION: CATEGORIES: / CATEGORY X —
      - DEEP DIVE: DEFINITION: / COMPONENTS: / CLINICAL RELEVANCE: / REAL-WORLD EXAMPLE:
      - Pipe-separated metadata: TYPE: X | TITLE: Y | SUBTITLE: Z
    """
    slides = []
    
    # Strip Gemini UI chrome
    text = re.sub(r'^.*?(?=---\s*SLIDE|SLIDE\s*\[?\d)', '', text, flags=re.DOTALL | re.IGNORECASE)
    
    # Split on slide boundaries
    slide_blocks = re.split(r'-{0,3}\s*SLIDE\s*\[?\d+\+?\]?\s*-{0,3}', text, flags=re.IGNORECASE)
    
    for block in slide_blocks:
        block = block.strip()
        if not block:
            continue
        
        slide = {
            "type": None, "title": None, "subtitle": None, 
            "visual": None, "bullets": [], "notes": "",
            "explanation": ""
        }
        lines = block.split("\n")
        
        # State tracking
        current_section = None  # body, sequence, categories, components, clinical_relevance, definition, notes, explanation
        current_category = None
        notes_lines = []
        explanation_lines = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            upper = line.upper()
            
            # --- Handle pipe-separated metadata line ---
            # e.g. "TYPE: MECHANISM | TITLE: Glomerular Filtration | SUBTITLE: Part 1"
            if '|' in line and upper.startswith('TYPE:'):
                parts = line.split('|')
                for part in parts:
                    part = part.strip()
                    part_upper = part.upper()
                    if part_upper.startswith('TYPE:'):
                        slide["type"] = part.split(':', 1)[1].strip()
                    elif part_upper.startswith('TITLE:'):
                        slide["title"] = part.split(':', 1)[1].strip()
                    elif part_upper.startswith('SUBTITLE:'):
                        sub = part.split(':', 1)[1].strip()
                        if sub and not sub.startswith('<'):
                            slide["subtitle"] = sub
                    elif part_upper.startswith('VISUAL:'):
                        vis = part.split(':', 1)[1].strip()
                        if vis and not vis.startswith('<'):
                            slide["visual"] = vis
                continue
            
            # --- Section keywords (change current_section state) ---
            
            # SPEAKER NOTES — discard entirely (user doesn't want them)
            if upper.startswith("SPEAKER NOTES:"):
                current_section = "notes"
                continue
            
            # EXPLANATION (inline or multi-line)
            if upper.startswith("EXPLANATION:"):
                current_section = "explanation"
                after = line.split(":", 1)[1].strip()
                if after and not after.startswith('<'):
                    explanation_lines.append(after)
                continue
            
            # BODY / BULLETS
            if upper.startswith("BODY:") or upper.startswith("BULLETS:"):
                current_section = "body"
                after = line.split(":", 1)[1].strip()
                if after and len(after) > 3:
                    slide["bullets"].append(after)
                continue
            
            # SEQUENCE (MECHANISM / DEEP DIVE Part 2)
            if upper.startswith("SEQUENCE:"):
                current_section = "sequence"
                continue
            
            # CATEGORIES (CLASSIFICATION)
            if upper.startswith("CATEGORIES:"):
                current_section = "categories"
                continue
            
            # COMPONENTS (DEEP DIVE Part 1)
            if upper.startswith("COMPONENTS:"):
                current_section = "components"
                continue
            
            # CLINICAL RELEVANCE (DEEP DIVE Part 3)
            if upper.startswith("CLINICAL RELEVANCE:"):
                current_section = "clinical_relevance"
                continue
            
            # DEFINITION (DEEP DIVE Part 1 — inline text → goes to bullets as context)
            if upper.startswith("DEFINITION:"):
                current_section = "definition"
                continue
            
            # REAL-WORLD EXAMPLE (DEEP DIVE Part 3 — inline text → goes to bullets)
            if upper.startswith("REAL-WORLD EXAMPLE:"):
                current_section = "real_world"
                continue
            
            # Non-pipe metadata (fallback for single-line fields)
            if upper.startswith("TYPE:") and not slide["type"]:
                slide["type"] = line.split(":", 1)[1].strip()
                continue
            if upper.startswith("TITLE:") and not slide["title"]:
                slide["title"] = line.split(":", 1)[1].strip()
                continue
            if upper.startswith("SUBTITLE:") and not slide["subtitle"]:
                sub = line.split(":", 1)[1].strip()
                if sub and not sub.startswith('<'):
                    slide["subtitle"] = sub
                continue
            if upper.startswith("VISUAL:") and not slide["visual"]:
                vis = line.split(":", 1)[1].strip()
                if vis and not vis.startswith('<'):
                    slide["visual"] = vis
                continue
            
            # --- Content lines based on current_section ---
            
            if current_section == "notes":
                # Discard speaker notes entirely
                pass
            
            elif current_section == "explanation":
                if not line.startswith('<'):
                    explanation_lines.append(line)
            
            elif current_section == "definition":
                # Definition text can be useful context — skip it (it's already implied by bullets)
                pass
            
            elif current_section == "real_world":
                # Real-world example text — skip (already distinct from explanation)
                pass
            
            elif current_section == "sequence":
                # STEP N → text
                step_match = re.match(r'STEP\s*\d+\s*[→>]\s*(.*)', line, re.IGNORECASE)
                if step_match:
                    step_text = step_match.group(1).strip()
                    if step_text and len(step_text) > 3:
                        slide["bullets"].append(f"→ {step_text}")
                elif '→' in line or '>' in line:
                    # Handle inline step chains: "STEP 1 → STEP 2 → STEP 3"
                    parts = re.split(r'[→>]', line)
                    for part in parts:
                        part = part.strip()
                        # Remove "STEP N" prefix
                        part = re.sub(r'^STEP\s*\d+\s*', '', part, flags=re.IGNORECASE).strip()
                        if part and len(part) > 3:
                            slide["bullets"].append(f"→ {part}")
            
            elif current_section == "categories":
                # CATEGORY X — Name
                cat_match = re.match(r'CATEGORY\s+\w+\s*[—\-]\s*(.*)', line, re.IGNORECASE)
                if cat_match:
                    current_category = cat_match.group(1).strip()
                    if current_category:
                        slide["bullets"].append(f"▸ {current_category}")
                elif line.startswith("-") or line.startswith("•") or line.startswith("*"):
                    bullet = line.lstrip("-•●* ").strip()
                    if bullet and len(bullet) > 3:
                        prefix = f"  " if current_category else ""
                        slide["bullets"].append(f"{prefix}{bullet}")
            
            elif current_section in ("body", "components", "clinical_relevance"):
                # Standard bullet lines
                if line.startswith("-") or line.startswith("•") or line.startswith("*"):
                    bullet = line.lstrip("-•●* ").strip()
                    if bullet and len(bullet) > 3:
                        slide["bullets"].append(bullet)
                elif len(line) > 10:
                    slide["bullets"].append(line)
            
            else:
                # No current section — try to parse as bullet anyway
                if line.startswith("-") or line.startswith("•") or line.startswith("*"):
                    bullet = line.lstrip("-•●* ").strip()
                    if bullet and len(bullet) > 3:
                        slide["bullets"].append(bullet)
        
        # Store explanation (ONLY from EXPLANATION: lines, NOT DEFINITION/REAL-WORLD)
        if explanation_lines:
            slide["explanation"] = " ".join(explanation_lines)
        
        # Speaker notes — discard, user doesn't want them
        # Notes field only gets explanation text for reference
        if explanation_lines:
            slide["notes"] = ""
        
        # Skip decorative slides
        slide_type = (slide.get("type") or "").upper()
        if slide_type in ("DIVIDER", "TITLE"):
            continue
        
        if slide["title"] and slide["bullets"]:
            slides.append(slide)
    
    # DEDUPLICATE: Remove slides with identical title+bullets (extraction sometimes captures 2-3x)
    seen = set()
    unique_slides = []
    for s in slides:
        key = (s.get('title', ''), tuple(s.get('bullets', [])[:3]))
        if key not in seen:
            seen.add(key)
            unique_slides.append(s)
    
    if len(unique_slides) < len(slides):
        print(f"    Deduplicated: {len(slides)} → {len(unique_slides)} slides")
    
    return unique_slides


def _parse_markdown_format(text):
    """Parse ## Header + bullet lists format"""
    slides = []
    # Split on markdown headers (##, ###, or bold **Title**)
    sections = re.split(r'\n(?=##\s|###\s|\*\*[A-Z])', text)
    
    for section in sections:
        section = section.strip()
        if not section or len(section) < 20:
            continue
        
        lines = section.split("\n")
        title = None
        bullets = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Header detection
            if line.startswith("#"):
                title = re.sub(r'^#+\s*', '', line).strip()
            elif line.startswith("**") and line.endswith("**"):
                title = line.strip("* ").strip()
            elif line.startswith("-") or line.startswith("•") or line.startswith("*"):
                bullet = line.lstrip("-•*● ").strip()
                if bullet and len(bullet) > 3:
                    bullets.append(bullet)
            elif not title and len(line) < 80 and len(line) > 3:
                title = line
            elif len(line) > 15:
                bullets.append(line)
        
        if title and bullets:
            slides.append({"type": "Concept", "title": title, "bullets": bullets[:5]})
    
    return slides


def _parse_numbered_format(text):
    """Parse numbered sections: 1. Title\ncontent"""
    slides = []
    # Split on numbered sections (1., 2., etc.)
    sections = re.split(r'\n(?=\d+[.)]\s+[A-Z])', text)
    
    for section in sections:
        section = section.strip()
        if not section or len(section) < 20:
            continue
        
        lines = section.split("\n")
        title = None
        bullets = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # First line with number is the title
            if not title and re.match(r'^\d+[.)]\s+', line):
                title = re.sub(r'^\d+[.)]\s+', '', line).strip()
                if title.endswith(":"):
                    title = title[:-1]
            elif line.startswith("-") or line.startswith("•") or line.startswith("*"):
                bullet = line.lstrip("-•●* ").strip()
                if bullet and len(bullet) > 3:
                    bullets.append(bullet)
            elif len(line) > 15 and title:
                # Treat longer lines as content/bullets
                bullets.append(line)
        
        if title and bullets:
            slides.append({"type": "Concept", "title": title, "bullets": bullets[:5]})
    
    return slides


def _parse_generic_format(text):
    """Last resort: split text into chunks by short lines (likely titles) followed by longer lines"""
    slides = []
    lines = text.split("\n")
    
    # If no newlines, try splitting by common patterns
    if len(lines) <= 1 and len(text) > 200:
        # Try to split by capitalized phrases that look like headers
        # Look for patterns like "TitleContent" where Title starts with capital
        parts = re.split(r'(?=[A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,5}(?:[:\-]|(?=[A-Z])))', text)
        if len(parts) > 3:
            lines = parts
    
    current_title = None
    current_bullets = []
    
    for line in lines:
        line = line.strip()
        if not line or len(line) < 4:
            continue
        
        # Short lines are likely titles, long lines are content
        if len(line) < 70 and not line.endswith('.') and len(line) > 5:
            # Save previous slide
            if current_title and current_bullets:
                slides.append({
                    "type": "Concept",
                    "title": current_title,
                    "bullets": current_bullets[:5]
                })
            current_title = line.strip(":").strip()
            current_bullets = []
        elif len(line) > 10:
            # Split long lines into sentences if needed
            if len(line) > 150:
                sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z])', line)
                for sent in sentences:
                    sent = sent.strip()
                    if len(sent) > 10:
                        current_bullets.append(sent if len(sent) < 150 else sent[:147] + "...")
            else:
                current_bullets.append(line)
    
    # Save last slide
    if current_title and current_bullets:
        slides.append({
            "type": "Concept",
            "title": current_title,
            "bullets": current_bullets[:5]
        })
    
    return slides


class ContentOrganizer:
    """
    2-Stage Content Organizer.
    Takes raw Gemini research → sends to Gemini with slide formatting prompt → parses structured output.
    """
    
    # UI artifacts to remove during fallback cleaning
    JUNK_PATTERNS = [
        r'^sign in$', r'^menu$', r'^settings$', r'^share$', r'^copy link$',
        r'^edit$', r'^new chat$', r'^google$', r'^gemini$', r'^search$',
        r'^home$', r'^send$', r'^stop$', r'^regenerate$', r'^retry$',
        r'^dark theme$', r'^light theme$', r'^feedback$', r'^help$',
        r'^about$', r'^privacy$', r'^terms$', r'^star$', r'^more$',
        r'^expand$', r'^collapse$', r'^close$', r'^back$', r'^forward$',
        r'^next$', r'^previous$', r'^show more$', r'^show less$',
        r'^drafts$', r'^modify response$', r'^share & export$',
        r'^copy$', r'^report legal issue$', r'^flag$',
        r'^skip to main content$', r'^cookie', r'^accept all$',
        r'^\d{1,2}:\d{2}$', r'^\d{1,2}:\d{2}\s*(am|pm)$',
        r'^just now$', r'^\d+ (second|minute|hour|day)s? ago$',
    ]
    
    def __init__(self):
        self._junk_re = [re.compile(p, re.IGNORECASE) for p in self.JUNK_PATTERNS]
        self._formatting_prompt = load_slide_formatting_prompt()
    
    def organize(self, research_data, topic, researcher=None, target_slides=20):
        """
        Main entry point. Checks for pre-formatted slides from research.py first,
        then falls back to regex-based detection.
        """
        raw_text = research_data.get('raw_text', '')
        existing_sections = research_data.get('sections', [])
        formatted_text = research_data.get('formatted_slides', '')
        output_dir = research_data.get('output_dir', 'output')
        
        print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
        print(f"{Fore.CYAN}CONTENT ORGANIZER{Style.RESET_ALL}")
        print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
        print(f"{Fore.CYAN}  Raw text: {len(raw_text)} chars{Style.RESET_ALL}")
        print(f"{Fore.CYAN}  Formatted slides text: {len(formatted_text)} chars{Style.RESET_ALL}")
        print(f"{Fore.CYAN}  Existing sections: {len(existing_sections)}{Style.RESET_ALL}")
        
        # PRIMARY: Try to parse pre-formatted slides from research.py (multiple formats)
        formatted_slides = None
        if formatted_text and len(formatted_text) > 100:
            print(f"{Fore.CYAN}  [*] Parsing pre-formatted slide data from Gemini...{Style.RESET_ALL}")
            formatted_slides = parse_slides(formatted_text)
            
            if formatted_slides and len(formatted_slides) >= 3:
                print(f"{Fore.GREEN}  [+] Parsed {len(formatted_slides)} slides from formatted data{Style.RESET_ALL}")
            else:
                print(f"{Fore.YELLOW}  [!] Strict parsing found {len(formatted_slides) if formatted_slides else 0} slides{Style.RESET_ALL}")
                formatted_slides = None
        
        if formatted_slides and len(formatted_slides) >= 3:
            slides = self._build_slide_list(formatted_slides, topic, target_slides)
        elif formatted_text and len(formatted_text) > 100:
            # FALLBACK A: Use the Stage 2 text through regex organizer (it's more organized than raw)
            print(f"{Fore.YELLOW}  [*] Using Stage 2 text with fallback organizer (better than raw){Style.RESET_ALL}")
            slides = self._fallback_organize(formatted_text, [], topic, target_slides)
        else:
            # FALLBACK B: Use raw Stage 1 text
            print(f"{Fore.YELLOW}  [!] No formatted text, using raw research with fallback{Style.RESET_ALL}")
            slides = self._fallback_organize(raw_text, existing_sections, topic, target_slides)
        
        print(f"{Fore.GREEN}  [+] Final slide count: {len(slides)}{Style.RESET_ALL}")
        print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
        
        # Save detailed debug log of parsed slides
        self._save_parse_debug_log(slides, output_dir)
        
        return slides
    
    def _save_parse_debug_log(self, slides, output_dir):
        """Save detailed per-slide parse info to debug log file."""
        try:
            import os, time as _time
            os.makedirs(output_dir, exist_ok=True)
            log_path = os.path.join(output_dir, 'slide_parse_debug.txt')
            with open(log_path, 'w', encoding='utf-8') as f:
                f.write(f"=== SLIDE PARSE DEBUG LOG ===\n")
                f.write(f"Timestamp: {_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Total slides parsed: {len(slides)}\n")
                f.write(f"{'='*70}\n\n")
                
                for i, s in enumerate(slides):
                    f.write(f"--- PARSED SLIDE {i+1} ---\n")
                    f.write(f"  Type:        {s.get('type', 'N/A')}\n")
                    f.write(f"  Title:       {s.get('title', 'N/A')}\n")
                    f.write(f"  Subtitle:    {s.get('subtitle', 'N/A')}\n")
                    f.write(f"  Visual:      {s.get('visual', 'N/A')}\n")
                    
                    bullets = s.get('bullets', [])
                    f.write(f"  Bullets ({len(bullets)}):\n")
                    for j, b in enumerate(bullets):
                        f.write(f"    [{j+1}] {b}\n")
                    
                    explanation = s.get('explanation', '')
                    f.write(f"  Explanation: {explanation[:150]}{'...' if len(explanation) > 150 else ''}\n")
                    
                    notes = s.get('notes', '')
                    f.write(f"  Notes:       {notes[:150]}{'...' if len(notes) > 150 else ''}\n")
                    f.write(f"\n")
            
            print(f"{Fore.GREEN}  [+] Parse debug saved to: {log_path}{Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.YELLOW}  [!] Could not save parse debug log: {e}{Style.RESET_ALL}")
    
    def _gemini_format_slides(self, raw_text, topic, researcher, target_slides):
        """
        Stage 2: Send raw research text back to Gemini with the slide formatting prompt.
        Returns parsed slide list or None on failure.
        """
        print(f"{Fore.CYAN}  [*] Stage 2: Sending research to Gemini for slide formatting...{Style.RESET_ALL}")
        
        # Truncate if too long (Gemini has input limits)
        max_input = 12000
        content_to_format = raw_text[:max_input] if len(raw_text) > max_input else raw_text
        
        # Build the formatting prompt with the research content injected
        formatting_request = f"""{self._formatting_prompt}

CONTENT TO FORMAT (Topic: {topic}):
---
{content_to_format}
---

Generate approximately {target_slides} slides from the above content. OUTPUT ONLY SLIDES IN THE SPECIFIED FORMAT."""
        
        try:
            response = researcher.ask_gemini(formatting_request)
            
            if response and len(response) > 50:
                print(f"{Fore.GREEN}  [+] Gemini returned {len(response)} chars of formatted slides{Style.RESET_ALL}")
                
                # Parse the structured slide output
                parsed = parse_slides(response)
                
                if parsed and len(parsed) >= 2:
                    print(f"{Fore.GREEN}  [+] Parsed {len(parsed)} slides from Gemini output{Style.RESET_ALL}")
                    return parsed
                else:
                    print(f"{Fore.YELLOW}  [!] Could not parse Gemini output into slides{Style.RESET_ALL}")
                    return None
            else:
                print(f"{Fore.YELLOW}  [!] Gemini returned insufficient response{Style.RESET_ALL}")
                return None
                
        except Exception as e:
            print(f"{Fore.RED}  [!] Gemini formatting failed: {e}{Style.RESET_ALL}")
            return None
    
    def _build_slide_list(self, parsed_slides, topic, target_slides):
        """
        Convert parse_slides() output into the final slide list format for ppt_builder.
        No speaker notes — only explanation text is preserved.
        """
        slides = []
        
        # 1. Title slide
        slides.append({
            'title': topic,
            'type': 'TITLE',
            'bullets': ['Comprehensive Research-Based Presentation',
                        'Generated from AI-Powered Analysis'],
            'notes': ''
        })
        
        # 2. Table of Contents (from parsed slide titles — unique only)
        toc_items = []
        for s in parsed_slides:
            title = s.get('title', '')
            stype = (s.get('type') or '').upper()
            if title and title not in toc_items and stype != 'SUMMARY':
                toc_items.append(title)
        
        if toc_items:
            slides.append({
                'title': 'Table of Contents',
                'type': 'TOC',
                'bullets': toc_items[:10],
                'notes': ''
            })
        
        # 3. Content slides from parsed Gemini output
        for s in parsed_slides:
            title = s.get('title', 'Topic')
            bullets = s.get('bullets', [])
            slide_type = (s.get('type') or 'CONCEPT')
            subtitle = s.get('subtitle', '')
            explanation = s.get('explanation', '')
            
            if not bullets:
                continue
            
            # If subtitle exists, append to title
            if subtitle:
                title = f"{title} — {subtitle}"
            
            # CLASSIFICATION and DEEP DIVE can have more bullets
            max_bullets = 10 if slide_type.upper() in ('CLASSIFICATION', 'DEEP DIVE') else 6
            
            slide_data = {
                'title': title,
                'type': slide_type,
                'bullets': bullets[:max_bullets],
                'notes': '',  # No speaker notes
                'explanation': explanation
            }
            
            # Map TYPE: Summary -> Key Takeaways slide
            if slide_type.upper() == 'SUMMARY':
                slide_data['title'] = title if 'synthesis' in title.lower() or 'takeaway' in title.lower() else 'Key Takeaways'
            
            slides.append(slide_data)
        
        # 4. Conclusion
        slides.append({
            'title': 'Conclusion',
            'type': 'CONCLUSION',
            'bullets': ['Thank you for your attention',
                        'Questions and Discussion',
                        'Contact Information'],
            'notes': ''
        })
        
        # 5. References
        import datetime
        slides.append({
            'title': 'References',
            'type': 'REFERENCES',
            'bullets': ['Research compiled using Google Gemini AI',
                        'Web sources and academic publications',
                        f'Current as of {datetime.datetime.now().strftime("%B %Y")}'],
            'notes': ''
        })
        
        return slides
    
    def _fallback_organize(self, raw_text, existing_sections, topic, target_slides):
        """
        Fallback: regex-based section detection when Gemini formatting is unavailable.
        """
        print(f"{Fore.CYAN}  [*] Using fallback regex-based organizer{Style.RESET_ALL}")
        
        # Clean the raw text
        cleaned = self._clean_text(raw_text)
        
        # Detect sections
        sections = self._detect_sections(cleaned)
        
        # Use existing sections if our detection is poor
        if len(sections) < 3 and len(existing_sections) >= 3:
            sections = existing_sections
            for section in sections:
                section['content'] = [self._clean_line(line) for line in section.get('content', []) if self._clean_line(line)]
        
        # Extract bullets
        sections = self._extract_bullets(sections)
        
        # Structure into slides
        return self._structure_slides(sections, topic, target_slides)
    
    def _clean_text(self, text):
        """Clean raw scraped text."""
        if not text:
            return ""
        
        lines = text.split('\n')
        cleaned_lines = []
        seen = set()
        
        for line in lines:
            line = line.strip()
            if not line or len(line) < 4:
                continue
            if self._is_junk(line):
                continue
            
            # Clean markdown
            line = re.sub(r'\*\*\*(.+?)\*\*\*', r'\1', line)
            line = re.sub(r'\*\*(.+?)\*\*', r'\1', line)
            line = re.sub(r'\*(.+?)\*', r'\1', line)
            line = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', line)
            
            line = self._clean_line(line)
            if not line or len(line) < 3:
                continue
            
            line_key = line.lower().strip()
            if line_key in seen:
                continue
            seen.add(line_key)
            cleaned_lines.append(line)
        
        return '\n'.join(cleaned_lines)
    
    def _is_junk(self, line):
        for pattern in self._junk_re:
            if pattern.match(line.strip()):
                return True
        return False
    
    def _clean_line(self, line):
        if not line:
            return ""
        line = re.sub(r'^[\s]*[-*+>]+\s*', '', line)
        line = re.sub(r'^[\s]*\d+[.)]\s*', '', line)
        line = re.sub(r'^#+\s*', '', line)
        return line.strip()
    
    def _detect_sections(self, text):
        lines = text.split('\n')
        sections = []
        current_title = "Introduction"
        current_content = []
        
        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue
            
            if self._is_header(line, i, lines):
                if current_content:
                    sections.append({'title': self._clean_title(current_title), 'content': current_content})
                current_title = line
                current_content = []
            elif len(line) > 8:
                current_content.append(line)
        
        if current_content:
            sections.append({'title': self._clean_title(current_title), 'content': current_content})
        
        if not sections and text:
            lines = [l.strip() for l in text.split('\n') if l.strip() and len(l.strip()) > 15]
            for i in range(0, len(lines), 5):
                chunk = lines[i:i+5]
                if chunk:
                    title = chunk[0] if len(chunk[0]) < 60 else f"Section {len(sections)+1}"
                    content = chunk[1:] if len(chunk[0]) < 60 else chunk
                    if content:
                        sections.append({'title': title, 'content': content})
        
        return sections
    
    def _is_header(self, line, index, all_lines):
        stripped = line.strip()
        if len(stripped) > 120:
            return False
        if stripped.startswith('#'):
            return True
        if re.match(r'^\d+[.)]\s+[A-Z]', stripped) and len(stripped) < 100:
            return True
        if stripped.isupper() and 5 < len(stripped) < 80:
            return True
        if stripped.endswith(':') and 8 < len(stripped) < 80:
            words = stripped.rstrip(':').split()
            if len(words) <= 8:
                return True
        if len(stripped) < 60 and index + 1 < len(all_lines):
            next_line = all_lines[index + 1].strip()
            if len(next_line) > len(stripped) * 1.5 and len(next_line) > 40:
                if not stripped.endswith('.') and stripped[0].isupper() and len(stripped.split()) <= 8:
                    return True
        if re.match(r'^(Introduction|Conclusion|Summary|Overview|Background|Discussion|'
                    r'Methodology|Results|Analysis|References|Key\s+\w+|Findings|'
                    r'Recommendations|Implementation|Future\s+\w+)\b', stripped, re.IGNORECASE):
            if len(stripped) < 60:
                return True
        return False
    
    def _clean_title(self, title):
        title = re.sub(r'^#+\s*', '', title)
        title = re.sub(r'^\d+[.)]\s*', '', title)
        title = title.rstrip(':').strip()
        if title.isupper():
            title = title.title()
        if len(title) > 60:
            title = title[:57] + "..."
        return title if title else "Topic"
    
    def _extract_bullets(self, sections):
        processed = []
        for section in sections:
            title = section.get('title', 'Topic')
            content = section.get('content', [])
            if not content:
                continue
            
            bullets = []
            for line in content:
                line = line.strip()
                if not line:
                    continue
                if len(line) > 200:
                    sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z])', line)
                    for sent in sentences:
                        sent = sent.strip()
                        if len(sent) > 15:
                            if len(sent) > 180:
                                sent = sent[:177] + "..."
                            bullets.append(sent)
                elif len(line) > 10:
                    if len(line) > 180:
                        line = line[:177] + "..."
                    bullets.append(line)
            
            if bullets:
                processed.append({'title': title, 'content': bullets})
        return processed
    
    def _structure_slides(self, sections, topic, target_slides):
        slides = []
        
        slides.append({
            'title': topic,
            'bullets': ['Comprehensive Research-Based Presentation', 'Generated from AI-Powered Analysis'],
            'notes': f'Welcome to this presentation on {topic}'
        })
        
        toc_items = [s.get('title', 'Topic') for s in sections[:10] 
                     if s.get('title') not in ('Introduction', 'Topic')]
        if toc_items:
            slides.append({'title': 'Table of Contents', 'bullets': toc_items[:8], 'notes': 'Overview'})
        
        max_bullets = 5
        for section in sections:
            title = section.get('title', 'Topic')
            content = section.get('content', [])
            if not content:
                continue
            
            for chunk_idx in range(0, len(content), max_bullets):
                chunk = content[chunk_idx:chunk_idx + max_bullets]
                if not chunk:
                    continue
                slide_title = title if chunk_idx == 0 else f"{title} (continued)"
                remaining = content[chunk_idx + max_bullets:]
                notes = ' '.join(remaining[:3]) if remaining else ''
                slides.append({'title': slide_title, 'bullets': chunk, 'notes': notes})
                if len(slides) >= target_slides + 3:
                    break
            if len(slides) >= target_slides + 3:
                break
        
        key_points = [s['content'][0] for s in sections[:5] if s.get('content')]
        if key_points:
            slides.append({'title': 'Key Takeaways', 'bullets': key_points[:5], 'notes': 'Summary'})
        
        slides.append({
            'title': 'Conclusion',
            'bullets': ['Thank you for your attention', 'Questions and Discussion', 'Contact Information'],
            'notes': 'End of presentation'
        })
        
        import datetime
        slides.append({
            'title': 'References',
            'bullets': ['Research compiled using Google Gemini AI', 'Web sources and academic publications',
                        f'Current as of {datetime.datetime.now().strftime("%B %Y")}'],
            'notes': 'All information sourced from reputable sources'
        })
        
        print(f"  Total slides structured: {len(slides)}")
        return slides


def organize_content(research_data, topic, researcher=None, target_slides=20):
    """Convenience function - main entry point for external use."""
    organizer = ContentOrganizer()
    return organizer.organize(research_data, topic, researcher, target_slides)
