"""
Content Generator Module
Uses research data to create structured presentation content
"""

from colorama import Fore, Style
import re


class ContentGenerator:
    """Generate structured presentation content using research data"""
    
    def __init__(self, researcher):
        """Initialize with active GeminiResearcher instance"""
        self.researcher = researcher
    
    def create_outline(self, topic, research_data):
        """Generate presentation structure from research data"""
        print(f"{Fore.CYAN}[*] Creating presentation outline...{Style.RESET_ALL}")
        
        # Check if we have research data
        if not research_data or not research_data.get('sections'):
            print(f"{Fore.YELLOW}[!] No research data available, using minimal outline{Style.RESET_ALL}")
            return self.create_minimal_outline(topic)
        
        # Create outline directly from research data
        print(f"{Fore.CYAN}   Using research data ({research_data['total_sections']} sections found){Style.RESET_ALL}")
        outline = self.create_outline_from_research(topic, research_data)
        
        # Only ask Gemini if we have very little content
        if len(outline) < 10:
            print(f"{Fore.YELLOW}[!] Limited content, asking Gemini for more...{Style.RESET_ALL}")
            gemini_outline = self.ask_gemini_for_outline(topic)
            if gemini_outline and len(gemini_outline) > len(outline):
                outline = gemini_outline
        
        print(f"{Fore.GREEN}[+] Created outline with {len(outline)} slides{Style.RESET_ALL}")
        return outline
    
    def create_outline_from_research(self, topic, research_data):
        """Create outline directly from extracted research data"""
        slides = []
        
        # Title slide
        slides.append({
            'title': topic,
            'bullets': ['Comprehensive Research-Based Presentation', 'Generated from AI Analysis'],
            'notes': 'Welcome to this presentation on ' + topic
        })
        
        # Table of contents
        section_titles = [s['title'] for s in research_data['sections'][:8]]
        slides.append({
            'title': 'Table of Contents',
            'bullets': section_titles if section_titles else ['Overview', 'Analysis', 'Conclusion'],
            'notes': 'Presentation structure and key topics'
        })
        
        # Process each research section into slides
        for section in research_data['sections']:
            section_title = section.get('title', 'Topic')
            section_content = section.get('content', [])
            
            if not section_content:
                continue
            
            # If section has lots of content, split into multiple slides
            if len(section_content) > 10:
                # First slide for this section
                slides.append({
                    'title': section_title,
                    'bullets': section_content[:6],
                    'notes': ' '.join(section_content[6:10]) if len(section_content) > 6 else ''
                })
                
                # Additional slides if needed
                remaining = section_content[6:]
                for i in range(0, len(remaining), 6):
                    chunk = remaining[i:i+6]
                    if chunk:
                        slides.append({
                            'title': f"{section_title} (continued)",
                            'bullets': chunk,
                            'notes': ''
                        })
                        if len(slides) >= 25:  # Limit total slides
                            break
            else:
                # Single slide for this section
                slides.append({
                    'title': section_title,
                    'bullets': section_content[:6],
                    'notes': ' '.join(section_content[6:]) if len(section_content) > 6 else ''
                })
            
            # Stop if we have enough slides
            if len(slides) >= 25:
                break
        
        # Add conclusion if we have room
        if len(slides) < 30:
            # Try to extract key points from research
            key_points = []
            if research_data.get('sections'):
                for section in research_data['sections'][:3]:
                    if section.get('content'):
                        key_points.append(section['content'][0])
            
            slides.append({
                'title': 'Key Takeaways',
                'bullets': key_points[:5] if key_points else ['Comprehensive Overview Provided', 'Multiple Perspectives Explored', 'Evidence-Based Analysis'],
                'notes': 'Summary of main points from this presentation'
            })
            
            slides.append({
                'title': 'Conclusion',
                'bullets': ['Thank you for your attention', 'Questions and Discussion', 'Contact Information'],
                'notes': 'End of presentation'
            })
        
        # References
        slides.append({
            'title': 'References',
            'bullets': ['Research compiled using Google Gemini', 'Web sources and academic publications', 'Current as of ' + __import__('datetime').datetime.now().strftime('%B %Y')],
            'notes': 'All information sourced from reputable sources'
        })
        
        return slides
    
    def ask_gemini_for_outline(self, topic):
        """Ask Gemini to generate presentation outline (used as fallback only)"""
        prompt = f"""Create a PowerPoint presentation outline for '{topic}' with 20-25 slides.

Format as:
SLIDE 1: Title
- Bullet 1
- Bullet 2

Make it professional and comprehensive."""
        
        response = self.researcher.ask_gemini(prompt)
        
        if response and len(response) > 200:
            return self.parse_outline(response)
        return []
    
    def parse_outline(self, outline_text):
        """Parse Gemini's outline into structured slides"""
        slides = []
        current_slide = None
        
        lines = outline_text.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Detect slide headers
            if line.upper().startswith('SLIDE') or re.match(r'^\d+\.', line):
                # Save previous slide
                if current_slide and current_slide.get('bullets'):
                    slides.append(current_slide)
                
                # Extract title
                title = re.sub(r'^SLIDE\s+\d+:|^\d+\.', '', line, flags=re.IGNORECASE).strip()
                title = title.lstrip(':').strip()
                
                if title:
                    current_slide = {
                        'title': title,
                        'bullets': [],
                        'notes': ''
                    }
            
            elif current_slide and line.startswith(('-', '-', '*', '-')):
                # Bullet point
                bullet = line.lstrip('--*- ').strip()
                if bullet:
                    current_slide['bullets'].append(bullet)
        
        # Add last slide
        if current_slide and current_slide.get('bullets'):
            slides.append(current_slide)
        
        return slides
    
    def create_minimal_outline(self, topic):
        """Create minimal outline when no research data available"""
        return [
            {'title': topic, 'bullets': ['Overview', 'Analysis'], 'notes': ''},
            {'title': 'Table of Contents', 'bullets': ['Introduction', 'Main Topics', 'Conclusion'], 'notes': ''},
            {'title': 'Introduction', 'bullets': ['Background', 'Objectives', 'Scope'], 'notes': ''},
            {'title': 'Overview', 'bullets': ['Key concepts', 'Important definitions'], 'notes': ''},
            {'title': 'Analysis', 'bullets': ['Current state', 'Trends', 'Implications'], 'notes': ''},
            {'title': 'Conclusion', 'bullets': ['Summary', 'Key takeaways', 'Thank you'], 'notes': ''},
            {'title': 'References', 'bullets': ['Sources cited'], 'notes': ''},
        ]
    
    def enhance_with_speaker_notes(self, slides):
        """Add speaker notes (skip Gemini to save time, use bullet content)"""
        print(f"{Fore.CYAN}[*] Adding speaker notes...{Style.RESET_ALL}")
        
        # Instead of asking Gemini (which is slow and unreliable),
        # generate notes from the bullet points
        for slide in slides:
            if not slide.get('notes') and slide.get('bullets'):
                # Create notes from bullets
                bullets_text = '. '.join(slide['bullets'][:3])
                slide['notes'] = f"This slide covers: {bullets_text}. Expand on each point with relevant details and examples."
        
        print(f"{Fore.GREEN}[+] Speaker notes added{Style.RESET_ALL}")
        return slides
    


def generate_presentation_content(topic, research_data, researcher):
    """Main function to generate all presentation content"""
    generator = ContentGenerator(researcher)
    
    print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}CONTENT GENERATION{Style.RESET_ALL}")
    print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
    
    # Show what we're working with
    if research_data:
        print(f"{Fore.GREEN}[+] Research data available:{Style.RESET_ALL}")
        print(f"  - Total text: {len(research_data.get('raw_text', ''))} characters")
        print(f"  - Sections: {research_data.get('total_sections', 0)}")
        print(f"  - Formatted slides: {len(research_data.get('formatted_slides', ''))} characters")
    else:
        print(f"{Fore.RED}[!] No research data available{Style.RESET_ALL}")
    
    # USE NEW CONTENT ORGANIZER for intelligent processing (2-stage pipeline)
    try:
        from modules.content_organizer import organize_content
        print(f"{Fore.CYAN}[*] Using Content Organizer (2-Stage Pipeline)...{Style.RESET_ALL}")
        slides = organize_content(research_data, topic, researcher=researcher, target_slides=20)
        print(f"{Fore.GREEN}[+] Content Organizer produced {len(slides)} slides{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.YELLOW}[!] Content Organizer failed: {e}{Style.RESET_ALL}")
        print(f"{Fore.YELLOW}    Falling back to basic outline...{Style.RESET_ALL}")
        slides = generator.create_outline(topic, research_data)
    
    # Clean up: remove Gemini UI artifacts from slide bullets
    ui_junk = ['Show thinking', 'Gemini said', 'Gemini 2.0', 'Show thinking\n']
    for slide in slides:
        if 'bullets' in slide:
            slide['bullets'] = [b for b in slide['bullets'] if b.strip() not in ui_junk and len(b.strip()) > 5]
        # Ensure no speaker notes
        slide['notes'] = ''
    
    print(f"{Fore.GREEN}[+] Generated {len(slides)} total slides{Style.RESET_ALL}")
    
    return slides, []

