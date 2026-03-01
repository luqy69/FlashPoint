"""
Automated PowerPoint Presentation Generator
Main script with CLI interface

Author: AI Research Assistant
Version: 1.0.0
"""

import os
import sys
from datetime import datetime
from colorama import init, Fore, Style
from tqdm import tqdm
import time

# Initialize colorama for Windows
init(autoreset=True)

# Add modules to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'modules'))

from modules.research import gather_research
from modules.content_generator import generate_presentation_content
from modules.visualizations import generate_visualizations
from modules.ppt_builder import build_presentation
from modules.research_doc import generate_supplemental_research
from modules.themes import get_theme_choice, apply_theme_to_presentation
from modules.export import get_export_preferences, export_to_ppt, export_to_pdf


class ProgressTracker:
    """Simple default-style progress bar similar to Go's progressbar.Default"""
    
    def __init__(self, total_steps=100):
        self.total_steps = total_steps
        self.current = 0
        self.pbar = None
        
    def start(self, description="Processing"):
        """Initialize the progress bar"""
        print(f"\n{Fore.CYAN}{'━' * 80}{Style.RESET_ALL}")
        print(f"{Fore.CYAN}  {description}{Style.RESET_ALL}")
        print(f"{Fore.CYAN}{'━' * 80}{Style.RESET_ALL}\n")
        
        self.pbar = tqdm(
            total=self.total_steps,
            bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]',
            colour='cyan',
            ncols=80
        )
    
    def update(self, amount=1):
        """Increment progress by amount"""
        if self.pbar:
            self.pbar.update(amount)
            self.current += amount
            time.sleep(0.04)  # 40ms delay like the Go example
    
    def set_description(self, desc):
        """Update the description"""
        if self.pbar:
            self.pbar.set_description(desc)
    
    def complete(self):
        """Mark progress as complete"""
        if self.pbar:
            self.pbar.close()
        print(f"\n{Fore.GREEN}{'━' * 80}{Style.RESET_ALL}\n")


def print_animated_header():
    """Print animated header with gradient effect"""
    colors = [Fore.CYAN, Fore.LIGHTCYAN_EX, Fore.WHITE, Fore.LIGHTCYAN_EX, Fore.CYAN]
    
    lines = [
        "╔════════════════════════════════════════════════════════════════════════════════╗",
        "║                                                                                ║",
        "║   ██████╗ ██████╗ ████████╗     ██████╗ ███████╗███╗   ██╗███████╗██████╗    ║",
        "║   ██╔══██╗██╔══██╗╚══██╔══╝    ██╔════╝ ██╔════╝████╗  ██║██╔════╝██╔══██╗   ║",
        "║   ██████╔╝██████╔╝   ██║       ██║  ███╗█████╗  ██╔██╗ ██║█████╗  ██████╔╝   ║",
        "║   ██╔═══╝ ██╔═══╝    ██║       ██║   ██║██╔══╝  ██║╚██╗██║██╔══╝  ██╔══██╗   ║",
        "║   ██║     ██║        ██║       ╚██████╔╝███████╗██║ ╚████║███████╗██║  ██║   ║",
        "║   ╚═╝     ╚═╝        ╚═╝        ╚═════╝ ╚══════╝╚═╝  ╚═══╝╚══════╝╚═╝  ╚═╝   ║",
        "║                                                                                ║",
        "║              🤖 AI-Powered Presentation Generator v2.0 🚀                      ║",
        "║                                                                                ║",
        "╚════════════════════════════════════════════════════════════════════════════════╝"
    ]
    
    print("\n")
    for i, line in enumerate(lines):
        color = colors[i % len(colors)]
        print(f"{color}{line}{Style.RESET_ALL}")
        time.sleep(0.05)  # Animated reveal
    print("\n")


def print_step_header(step_num, total_steps, title):
    """Print polished step header with progress indicator"""
    percentage = int((step_num / total_steps) * 100)
    
    # Create mini progress indicator
    filled = int(20 * step_num / total_steps)
    mini_bar = "█" * filled + "░" * (20 - filled)
    
    print(f"\n{Fore.MAGENTA}{'═' * 80}{Style.RESET_ALL}")
    print(f"{Fore.MAGENTA}║{Style.RESET_ALL} ", end="")
    print(f"{Fore.CYAN}STEP {step_num}/{total_steps}{Style.RESET_ALL} ", end="")
    print(f"{Fore.WHITE}[{mini_bar}]{Style.RESET_ALL} ", end="")
    print(f"{Fore.YELLOW}{percentage}%{Style.RESET_ALL} ", end="")
    print(f"{Fore.MAGENTA}║{Style.RESET_ALL}")
    print(f"{Fore.MAGENTA}║{Style.RESET_ALL} {Fore.GREEN}► {title}{Style.RESET_ALL}")
    print(f"{Fore.MAGENTA}{'═' * 80}{Style.RESET_ALL}\n")
    time.sleep(0.1)


def print_success_box(title, items):
    """Print success message in a styled box"""
    print(f"\n{Fore.GREEN}╔{'═' * 78}╗{Style.RESET_ALL}")
    print(f"{Fore.GREEN}║{Style.RESET_ALL} {Fore.WHITE}✓ {title}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}╠{'═' * 78}╣{Style.RESET_ALL}")
    
    for item in items:
        print(f"{Fore.GREEN}║{Style.RESET_ALL} {item}")
    
    print(f"{Fore.GREEN}╚{'═' * 78}╝{Style.RESET_ALL}\n")


def print_spinner(message, duration=2):
    """Show a spinner animation"""
    spinner_chars = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏']
    end_time = time.time() + duration
    
    i = 0
    while time.time() < end_time:
        sys.stdout.write(f'\r{Fore.CYAN}{spinner_chars[i % len(spinner_chars)]} {message}{Style.RESET_ALL}')
        sys.stdout.flush()
        time.sleep(0.1)
        i += 1
    
    sys.stdout.write(f'\r{Fore.GREEN}✓ {message}{Style.RESET_ALL}\n')
    sys.stdout.flush()





def print_banner():
    """Print banner with ASCII art logo"""
    print(f"\n{Fore.CYAN}{'=' * 90}")
    print(f"{Fore.YELLOW}██████  ██████  ████████{Fore.CYAN}      {Fore.GREEN}██████  ███████ ███    ██ ██████   █████  ████████  ██████  ██████  {Fore.CYAN}")
    print(f"{Fore.YELLOW}██   ██ ██   ██    ██   {Fore.CYAN}     {Fore.GREEN}██       ██      ████   ██ ██   ██ ██   ██    ██    ██    ██ ██   ██ {Fore.CYAN}")
    print(f"{Fore.YELLOW}██████  ██████     ██   {Fore.CYAN}     {Fore.GREEN}██   ███ █████   ██ ██  ██ ██████  ███████    ██    ██    ██ ██████  {Fore.CYAN}")
    print(f"{Fore.YELLOW}██      ██         ██   {Fore.CYAN}     {Fore.GREEN}██    ██ ██      ██  ██ ██ ██   ██ ██   ██    ██    ██    ██ ██   ██ {Fore.CYAN}")
    print(f"{Fore.YELLOW}██      ██         ██   {Fore.CYAN}     {Fore.GREEN} ██████  ███████ ██   ████ ██   ██ ██   ██    ██     ██████  ██   ██ {Fore.CYAN}")
    print(f"")
    print(f"  {Fore.WHITE}AI-Powered Presentation Generator{Fore.CYAN}")
    print(f"  {Fore.MAGENTA}Automated Research • Professional Slides • Smart Visuals{Fore.CYAN}")
    print(f"{'=' * 90}{Style.RESET_ALL}\n")


def print_progress_header(step, description):
    """Print progress step header with icon"""
    icons = {
        1: "RESEARCH",
        2: "CONTENT", 
        3: "VISUALS",
        4: "BUILD"
    }
    
    icon_text = icons.get(step, "STEP")
    
    print(f"\n{Fore.CYAN}{'=' * 76}")
    print(f"  [{icon_text}] STEP {step}/4: {description}")
    print(f"{'=' * 76}{Style.RESET_ALL}\n")


def print_box(title, items, color=Fore.GREEN, title_icon=""):
    """Print a formatted box with title and items"""
    width = 74
    
    # Top border
    print(f"\n{color}{'=' * 76}")
    
    # Title line
    icon_prefix = f"{title_icon} " if title_icon else ""
    title_text = f"  {icon_prefix}{title}"
    print(f"{title_text}{' ' * (76 - len(title_text))}")
    
    # Separator
    print(f"{'-' * 76}")
    
    # Items
    for item in items:
        # Handle long lines
        if len(item) > 72:
            print(f"  {item[:72]}")
            remaining = item[72:]
            while remaining:
                print(f"  {remaining[:72]}")
                remaining = remaining[72:]
        else:
            print(f"  {item}{' ' * (74 - len(item))}")
    
    # Bottom border
    print(f"{'=' * 76}{Style.RESET_ALL}")


def get_user_input():
    """Get topic and preferences from user"""
    print(f"\n{Fore.CYAN}{'=' * 76}")
    print(f"  CONFIGURATION")
    print(f"{'=' * 76}{Style.RESET_ALL}\n")
    
    # Get topic
    while True:
        topic = input(f"{Fore.GREEN}  Enter presentation topic: {Style.RESET_ALL}").strip()
        if topic:
            print(f"{Fore.CYAN}  > Topic set: {Fore.WHITE}{topic}{Style.RESET_ALL}")
            break
        print(f"{Fore.RED}  X Topic cannot be empty. Please try again.{Style.RESET_ALL}")
    
    # Optional settings
    print(f"\n{Fore.CYAN}  Optional Settings {Fore.YELLOW}(press Enter for defaults){Style.RESET_ALL}")
    
    slide_count_input = input(f"{Fore.GREEN}  Target slide count [{Fore.YELLOW}default: 20-25{Fore.GREEN}]: {Style.RESET_ALL}").strip()
    try:
        slide_count = int(slide_count_input) if slide_count_input else 22
    except:
        slide_count = 22
    
    print(f"{Fore.CYAN}  > Slides: {Fore.WHITE}{slide_count}{Style.RESET_ALL}")
    
    # Theme selection
    theme_choice, theme_info = get_theme_choice()
    
    # Export preferences
    export_prefs = get_export_preferences()
    
    return {
        'topic': topic,
        'slide_count': slide_count,
        'theme_choice': theme_choice,
        'theme_info': theme_info,
        'export_prefs': export_prefs
    }


def print_success_box(title, details):
    """Print a success message in a box"""
    print_box(title, details, Fore.GREEN, "✓")


def simulate_progress(description, duration=2):
    """Show a progress bar for a task"""
    with tqdm(total=100, desc=description, bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt}') as pbar:
        for i in range(100):
            time.sleep(duration / 100)
            pbar.update(1)


def main():
    """Main application flow"""
    try:
        # Print banner
        print_banner()
        
        # Get user input
        config = get_user_input()
        topic = config['topic']
        
        
        # Configuration summary
        print_box("CONFIGURATION COMPLETE", [
            f"{Fore.CYAN}Topic:{Style.RESET_ALL} {Fore.WHITE}{topic}{Style.RESET_ALL}",
            f"{Fore.CYAN}Target Slides:{Style.RESET_ALL} {Fore.WHITE}{config['slide_count']}{Style.RESET_ALL}"
        ], Fore.GREEN, "✓")
        
        # Create output directory
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        # Sanitize topic for safe file path
        safe_topic = topic.strip().replace(" ", "_").replace("\\", "").replace("/", "")
        output_dir = os.path.join(os.path.dirname(__file__), 'output', f'{safe_topic}_{timestamp}')
        os.makedirs(output_dir, exist_ok=True)
        
        charts_dir = os.path.join(output_dir, 'charts')
        os.makedirs(charts_dir, exist_ok=True)
        
        # Initialize progress tracker
        progress = ProgressTracker(total_steps=5)
        progress.start("Generating Presentation")
        
        # STEP 1: Research
        print_step_header(1, 5, "Web Research Using Gemini")
        progress.set_description("Researching")
        progress.update(1)
        
        research_data, researcher = gather_research(topic, lambda: progress.update(0.04))
        
        if not research_data or not researcher:
            print(f"{Fore.RED}✗ Research failed. Cannot proceed.{Style.RESET_ALL}")
            progress.complete()
            return 1
        
        # Pass output_dir to research_data for debug logging
        research_data['output_dir'] = output_dir
        
        print(f"\n{Fore.GREEN}✓ Research completed successfully!{Style.RESET_ALL}")
        print(f"{Fore.CYAN}  Sections found: {research_data.get('total_sections', 0)}{Style.RESET_ALL}")
        
        # STEP 2: Generate Content
        print_step_header(2, 5, "Generating Presentation Content")
        progress.set_description("Generating")
        progress.update(1)
        
        try:
            slides, viz_suggestions = generate_presentation_content(topic, research_data, researcher)
            
            print(f"\n{Fore.GREEN}✓ Content generation completed!{Style.RESET_ALL}")
            print(f"{Fore.CYAN}  Total slides: {len(slides)}{Style.RESET_ALL}")
            print(f"{Fore.CYAN}  Visualizations planned: {len(viz_suggestions)}{Style.RESET_ALL}")
            
        except Exception as e:
            print(f"{Fore.RED}✗ Content generation failed: {e}{Style.RESET_ALL}")
            progress.complete()
            return 1
        
        # STEP 3: Generate Visualizations
        print_step_header(3, 5, "Creating Visualizations")
        progress.set_description("Downloading")
        progress.update(1)
        
        try:
            # Generate visualizations (downloads from Google Images + sample fallback)
            image_paths = generate_visualizations(topic, researcher, charts_dir)
            
            print(f"\n{Fore.GREEN}✓ Visualizations created!{Style.RESET_ALL}")
            print(f"{Fore.CYAN}  Images downloaded: {len(image_paths)}{Style.RESET_ALL}")
            
        except Exception as e:
            print(f"{Fore.YELLOW}⚠ Visualization generation had issues: {e}{Style.RESET_ALL}")
            image_paths = []
        
        
        # STEP 4: Build Presentation
        print_step_header(4, 5, "Building PowerPoint Presentation")
        progress.set_description("Building PPT")
        progress.update(1)
        
        output_filename = os.path.join(output_dir, f'{topic.replace(" ", "_")}_presentation.pptx')
        
        try:
            # Build presentation
            prs = build_presentation(topic, slides, image_paths, output_filename, config)
            
            # Apply selected theme
            apply_theme_to_presentation(prs, config['theme_choice'])
            
            # Save with theme
            prs.save(output_filename)
            
            print(f"\n{Fore.GREEN}✓ PowerPoint created successfully!{Style.RESET_ALL}")
            print(f"{Fore.CYAN}  Theme: {config['theme_info']['name']}{Style.RESET_ALL}")
            print(f"{Fore.CYAN}  File: {os.path.basename(output_filename)}{Style.RESET_ALL}")
            
            # Export to additional formats
            exported_files = [output_filename]
            
            if config['export_prefs'].get('ppt'):
                ppt_file = export_to_ppt(output_filename)
                if ppt_file:
                    exported_files.append(ppt_file)
            
            if config['export_prefs'].get('pdf'):
                pdf_file = export_to_pdf(output_filename)
                if pdf_file:
                    exported_files.append(pdf_file)
            
        except Exception as e:
            print(f"{Fore.RED}✗ Error building presentation: {e}{Style.RESET_ALL}")
            # Close browser before exiting
            if researcher:
                researcher.close_browser()
            progress.complete()
            return 1
        
        # STEP 5: Generate Supplemental Research Document
        print_step_header(5, 5, "Creating Supplemental Research Document")
        progress.set_description("Creating Doc")
        progress.update(1)
        
        research_doc_path = None
        try:
            research_doc_path = generate_supplemental_research(topic, research_data, output_dir)
            if research_doc_path:
                print(f"\n{Fore.GREEN}✓ Supplemental research document created!{Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.YELLOW}⚠ Could not create research document: {e}{Style.RESET_ALL}")
        
        # Close browser
        if researcher:
            researcher.close_browser()
        
        # Success display
        print_box("PRESENTATION GENERATED SUCCESSFULLY!", [], Fore.GREEN, "✓")
        
        # Output files
        output_files = [
            f"{Fore.WHITE}Presentation:{Style.RESET_ALL} {os.path.basename(output_filename)}",
            f"{Fore.WHITE}Location:{Style.RESET_ALL} {output_dir}"
        ]
        if research_doc_path:
            output_files.append(f"{Fore.WHITE}Research Doc:{Style.RESET_ALL} {os.path.basename(research_doc_path)}")
        
        print_box("OUTPUT FILES", output_files, Fore.CYAN, "")
        
        # Statistics
        print_box("PRESENTATION STATISTICS", [
            f"{Fore.GREEN}✓{Style.RESET_ALL} Total Slides: {Fore.WHITE}{len(slides) + len(image_paths) + 2}{Style.RESET_ALL}",
            f"{Fore.GREEN}✓{Style.RESET_ALL} Content Slides: {Fore.WHITE}{len(slides)}{Style.RESET_ALL}",
            f"{Fore.GREEN}✓{Style.RESET_ALL} Visualizations: {Fore.WHITE}{len(image_paths)}{Style.RESET_ALL}",
            f"{Fore.GREEN}✓{Style.RESET_ALL} Research Sections: {Fore.WHITE}{research_data.get('total_sections', 0)}{Style.RESET_ALL}"
        ], Fore.MAGENTA, "")
        
        # Next steps
        print_box("NEXT STEPS", [
            "1. Open the presentation in Microsoft PowerPoint or Google Slides",
            "2. Review speaker notes for additional context",
            "3. Customize colors, fonts, and layout as needed"
        ], Fore.YELLOW, "")
        
        # Open file location
        try:
            os.startfile(output_dir)
        except:
            print(f"{Fore.CYAN}Open the output directory to view your files.{Style.RESET_ALL}")
        
        return 0
        
    except KeyboardInterrupt:
        print(f"\n\n{Fore.RED}✗ Operation cancelled by user{Style.RESET_ALL}")
        return 1
    
    except Exception as e:
        print(f"\n{Fore.RED}✗ Unexpected error: {e}{Style.RESET_ALL}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit_code = main()
    
    print(f"\n{Fore.CYAN}Press Enter to exit...{Style.RESET_ALL}")
    input()
    
    sys.exit(exit_code)
