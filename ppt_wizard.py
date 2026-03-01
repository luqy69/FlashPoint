"""
AI PowerPoint Generator - GUI Wizard
Professional wizard-style interface for presentation generation

Made by Luqman (2024-ag-8738 DVM)
Discord: sad_memer.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
import sys
import ctypes
import time
from pathlib import Path
from datetime import datetime

# Enable high-DPI awareness for crisp text on Windows
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Windows 8.1+
except:
    try:
        ctypes.windll.user32.SetProcessDPIAware()  # Windows Vista+  
    except:
        pass  # DPI awareness not available

# Add modules to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'modules'))

from modules.research import GeminiResearcher, gather_research
from modules.content_generator import generate_presentation_content
from modules.visualizations import generate_visualizations
from modules.ppt_builder import build_presentation
from modules.research_doc import generate_supplemental_research
from modules.themes import BUILTIN_THEMES, apply_theme_to_presentation
from modules.export import export_to_ppt, export_to_pdf


class PPTWizard(tk.Tk):
    """Main wizard application window"""
    
    def __init__(self):
        super().__init__()
        
        self.title("AI PowerPoint Generator")
        self.geometry("850x750")  # Increased height: 850 wide x 750 tall
        self.resizable(False, False)
        
        # Configuration
        self.config_data = {
            'topic': '',
            'slide_count': 25,  # New: default 25 slides
            'research_level': 3,  # New: Level 3 (Master Thesis) by default
            'theme_choice': '1',
            'export_ppt': True,
            'export_pdf': False
        }
        
        # Current step tracking
        self.current_step = 0
        self.total_steps = 8  # Increased from 7 (added Super Research Options step)
        
        # Setup UI
        self.setup_ui()
        self.show_step(0)
        
    def setup_ui(self):
        """Initialize the Windows wizard-style UI"""
        # Configure window
        self.configure(bg='white')
        
        # Main container
        main_container = tk.Frame(self, bg='white')
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # TOP HEADER BAR (Navy - Minimalist Design)
        header_bar = tk.Frame(main_container, bg='#0F172A', height=65)
        header_bar.pack(side=tk.TOP, fill=tk.X)
        header_bar.pack_propagate(False)  # Maintain fixed height
        
        # Header title
        header_title = tk.Label(
            header_bar,
            text="FlashPoint AI",
            font=("Segoe UI", 18, "bold"),
            bg='#0F172A',
            fg='white'
        )
        header_title.pack(pady=18)
        
        # RIGHT CONTENT AREA (White)
        content_container = tk.Frame(main_container, bg='white')
        content_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Title area
        self.title_frame = tk.Frame(content_container, bg='white')
        self.title_frame.pack(fill=tk.X, padx=20, pady=(20, 10))
        
        self.title_label = tk.Label(
            self.title_frame,
            text="AI PowerPoint Generator",
            font=("Segoe UI", 12, "bold"),
            bg='white',
            fg='#0F172A',  # Navy - minimalist
            anchor='w'
        )
        self.title_label.pack(anchor='w')
        
        # Subtitle/step indicator
        self.subtitle_label = tk.Label(
            self.title_frame,
            text="Welcome",
            font=("Segoe UI", 16, "bold"),
            bg='white',
            fg='#0F172A',  # Navy - minimalist
            anchor='w'
        )
        self.subtitle_label.pack(anchor='w', pady=(5, 0))
        
        # Separator line
        separator = tk.Frame(content_container, height=1, bg='#D3D3D3')
        separator.pack(fill=tk.X, padx=0)
        
        # Bottom button area - PACK THIS FIRST FROM BOTTOM TO ENSURE IT'S ALWAYS VISIBLE
        button_container = tk.Frame(content_container, bg='white', height=65)
        button_container.pack(fill=tk.X, side=tk.BOTTOM)
        button_container.pack_propagate(False)
        
        # Button frame with padding
        button_frame = tk.Frame(button_container, bg='white')
        button_frame.pack(side=tk.RIGHT, padx=10, pady=10)
        
        # Cancel button (right-most) - Gray flat
        self.cancel_btn = tk.Button(
            button_frame,
            text="Cancel",
            width=10,
            command=self.destroy,
            bg='#E5E7EB',
            fg='#374151',
            activebackground='#D1D5DB',
            relief=tk.FLAT,
            bd=0,
            font=("Segoe UI", 9),
            cursor='hand2'
        )
        self.cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        # Next button - Orange flat (minimalist accent)
        self.next_btn = tk.Button(
            button_frame,
            text="Next >",
            width=12,
            command=self.go_next,
            bg='#F97316',
            fg='white',
            activebackground='#EA580C',
            relief=tk.FLAT,
            bd=0,
            font=("Segoe UI", 9, "bold"),
            default=tk.ACTIVE,
            cursor='hand2'
        )
        self.next_btn.pack(side=tk.RIGHT, padx=5)
        
        # Back button - Gray flat
        self.back_btn = tk.Button(
            button_frame,
            text="< Back",
            width=10,
            command=self.go_back,
            state=tk.DISABLED,
            bg='#E5E7EB',
            fg='#374151',
            activebackground='#D1D5DB',
            relief=tk.FLAT,
            bd=0,
            font=("Segoe UI", 9),
            cursor='hand2'
        )
        self.back_btn.pack(side=tk.RIGHT, padx=5)
        
        # Branding in bottom left
        branding_label = tk.Label(
            button_container,
            text="By Luqman (2024-ag-8738 DVM)",
            font=("Segoe UI", 8, "bold"),
            bg='#F0F0F0',
            fg='#333',
            anchor=tk.W
        )
        branding_label.pack(side=tk.LEFT, padx=10)
        
        # Bottom separator ABOVE buttons
        bottom_separator = tk.Frame(content_container, height=1, bg='#D3D3D3')
        bottom_separator.pack(fill=tk.X, side=tk.BOTTOM, padx=0)
        
        # Content frame - PACK LAST so it fills remaining space
        self.content_frame = tk.Frame(content_container, bg='white')
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=15)
        
    
    def show_step(self, step_num):
        """Display specific wizard step"""
        self.current_step = step_num
        
        # Update subtitle based on current step
        step_titles = [
            "Welcome",
            "Topic Input",
            "Advanced Options",  # Slide count & research depth
            "Super Research Options",  # NEW - Toggles for Super Research & Auto-Images
            "Theme Selection",
            "Export Options",
            "Generating...",
            "Complete"
        ]
        self.subtitle_label.config(text=step_titles[step_num])
        
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        # Show appropriate step content
        if step_num == 0:
            self.show_welcome()
        elif step_num == 1:
            self.show_topic_input()
        elif step_num == 2:
            self.show_advanced_options()  # Slide count & research depth
        elif step_num == 3:
            self.show_super_research_options()  # NEW - Super Research & Auto-Images toggles
        elif step_num == 4:
            self.show_theme_selection()
        elif step_num == 5:
            self.show_export_options()
        elif step_num == 6:
            self.show_progress()
        elif step_num == 7:
            self.show_completion()
        
        # Update button states
        # Disable back on welcome (step 0) and generation/completion (step 6+)
        if step_num == 0 or step_num >= 6:
            self.back_btn.config(state=tk.DISABLED)
        else:
            self.back_btn.config(state=tk.NORMAL)
        
        # Disable next on generation (step 6) and completion (step 7)
        if step_num >= 6:
            self.next_btn.config(state=tk.DISABLED)
        else:
            self.next_btn.config(text="Next >", command=self.go_next, state=tk.NORMAL)
        
        # Force update to ensure everything renders
        self.update_idletasks()
    
    def show_welcome(self):
        """Step 0: Welcome screen - Windows wizard style"""
        # Main instruction text
        instruction = tk.Label(
            self.content_frame,
            text="This wizard will help you create professional PowerPoint presentations\npowered by AI research and stunning visuals.",
            font=("Segoe UI", 9),
            bg='white',
            justify=tk.LEFT,
            wraplength=450
        )
        instruction.pack(anchor='w', pady=(0, 20))
        
        # Features section
        features_text = """Features:
- AI-powered research using Google Gemini
- 10+ professional PowerPoint themes
- High-quality visualizations
- Supplemental research document
- Export to PPTX, PPT, and PDF formats"""
        
        features = tk.Label(
            self.content_frame,
            text=features_text,
            font=("Segoe UI", 9),
            bg='white',
            justify=tk.LEFT
        )
        features.pack(anchor='w', pady=(0, 20))
        
        # Important note
        note_frame = tk.LabelFrame(
            self.content_frame,
            text="  Important  ",
            font=("Segoe UI", 9, "bold"),
            bg='white',
            fg='#000080',
            relief=tk.GROOVE,
            borderwidth=1
        )
        note_frame.pack(fill=tk.X, pady=20)
        
        note_text = tk.Label(
            note_frame,
            text="- You must be logged into Google Gemini\n- Google Chrome will be used for automation\n- Generation may take 2-5 minutes",
            font=("Segoe UI", 9),
            bg='white',
            justify=tk.LEFT
        )
        note_text.pack(padx=10, pady=10, anchor='w')
        
        # Bottom instruction
        continue_text = tk.Label(
            self.content_frame,
            text="To continue, click Next.",
            font=("Segoe UI", 9),
            bg='white'
        )
        continue_text.pack(side=tk.BOTTOM, anchor='w', pady=(20, 0))
    
    def show_topic_input(self):
        """Step 1: Topic input with Gemini login check"""
        ttk.Label(
            self.content_frame,
            text="Enter Your Presentation Topic",
            font=("Arial", 14, "bold")
        ).pack(pady=20)
        
        ttk.Label(
            self.content_frame,
            text="What would you like to create a presentation about?",
            font=("Arial", 10)
        ).pack()
        
        # Topic entry
        self.topic_var = tk.StringVar(value=self.config_data.get('topic', ''))
        
        entry_frame = ttk.Frame(self.content_frame)
        entry_frame.pack(pady=20, padx=50, fill=tk.X)
        
        ttk.Label(entry_frame, text="Topic:", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        
        self.topic_entry = ttk.Entry(entry_frame, textvariable=self.topic_var, font=("Arial", 11))
        self.topic_entry.pack(fill=tk.X, pady=5)
        self.topic_entry.focus()
        
        # Examples
        examples_frame = ttk.LabelFrame(self.content_frame, text="Examples", padding=10)
        examples_frame.pack(pady=10, padx=50, fill=tk.X)
        
        examples = [
            "Artificial Intelligence in Healthcare",
            "Climate Change and Global Warming",
            "Quantum Computing Fundamentals"
        ]
        
        for example in examples:
            btn = ttk.Button(
                examples_frame,
                text=f"- {example}",
                command=lambda e=example: self.topic_var.set(e),
                style="Link.TButton"
            )
            btn.pack(anchor=tk.W, pady=2)
        
        # Gemini login status
        self.login_status_label = ttk.Label(
            self.content_frame,
            text="[OK] Ready to proceed",
            font=("Arial", 9),
            foreground="green"
        )
        self.login_status_label.pack(pady=10)
    
    def show_advanced_options(self):
        """Step 2: Advanced Options - Slide count and research depth ONLY"""
        ttk.Label(
            self.content_frame,
            text="Advanced Options",
            font=("Arial", 14, "bold")
        ).pack(pady=10)

        
        # Slide Count Selection
        slide_frame = ttk.LabelFrame(self.content_frame, text="Slide Count", padding=15)
        slide_frame.pack(pady=10, padx=50, fill=tk.X)
        
        ttk.Label(
            slide_frame,
            text="How many slides do you want?",
            font=("Arial", 10)
        ).pack(anchor=tk.W, pady=5)
        
        # Combobox for slide count
        self.slide_count_var = tk.StringVar(value=str(self.config_data.get('slide_count', 25)))
        slide_combo = ttk.Combobox(
            slide_frame,
            textvariable=self.slide_count_var,
            values=["10", "15", "20", "25", "30"],
            state="readonly",
            width=10,
            font=("Arial", 11)
        )
        slide_combo.pack(anchor=tk.W, pady=5)
        
        # Research Depth Selection
        research_frame = ttk.LabelFrame(self.content_frame, text="Research Depth", padding=15)
        research_frame.pack(pady=10, padx=50, fill=tk.X)  # Don't expand vertically
        
        ttk.Label(
            research_frame,
            text="Choose the level of detail for AI research:",
            font=("Arial", 10)
        ).pack(anchor=tk.W, pady=5)
        
        self.research_level_var = tk.IntVar(value=self.config_data.get('research_level', 3))
        
        # Level 1: Basic
        level1_frame = ttk.Frame(research_frame)
        level1_frame.pack(anchor=tk.W, pady=8)
        
        ttk.Radiobutton(
            level1_frame,
            text="Level 1: Basic",
            variable=self.research_level_var,
            value=1
        ).pack(anchor=tk.W)
        
        ttk.Label(
            level1_frame,
            text="   -> For general audiences - Simple explanations - Fun facts",
            font=("Arial", 9),
            foreground="#666"
        ).pack(anchor=tk.W)
        
        # Level 2: Professional
        level2_frame = ttk.Frame(research_frame)
        level2_frame.pack(anchor=tk.W, pady=8)
        
        ttk.Radiobutton(
            level2_frame,
            text="Level 2: Professional",
            variable=self.research_level_var,
            value=2
        ).pack(anchor=tk.W)
        
        ttk.Label(
            level2_frame,
            text="   -> For technical presentations - Data-driven - 2024-2026 studies",
            font=("Arial", 9),
            foreground="#666"
        ).pack(anchor=tk.W)
        
        # Level 3: Master Thesis
        level3_frame = ttk.Frame(research_frame)
        level3_frame.pack(anchor=tk.W, pady=8)
        
        ttk.Radiobutton(
            level3_frame,
            text="Level 3: Master Thesis",
            variable=self.research_level_var,
            value=3
        ).pack(anchor=tk.W)
        
        ttk.Label(
            level3_frame,
            text="   -> Maximum detail - 5000-word depth - Expert-level content",
            font=("Arial", 9),
            foreground="#666"
        ).pack(anchor=tk.W)
        
        # Save to config when user changes values
        def save_advanced_options():
            self.config_data['slide_count'] = int(self.slide_count_var.get())
            self.config_data['research_level'] = self.research_level_var.get()
        
        # Initialize toggle variables if not already set (will be set in next step)
        if not hasattr(self, 'super_research_var'):
            self.super_research_var = tk.BooleanVar(value=self.config_data.get('super_research', False))
        if not hasattr(self, 'auto_images_var'):
            self.auto_images_var = tk.BooleanVar(value=self.config_data.get('auto_images', False))
        
        # Save initial values immediately
        save_advanced_options()
        
        slide_combo.bind("<<ComboboxSelected>>", lambda e: save_advanced_options())
        self.research_level_var.trace_add("write", lambda *args: save_advanced_options())
    
    def show_super_research_options(self):
        """Step 3: Super Research & Auto-Images Toggles - OWN PAGE"""
        ttk.Label(
            self.content_frame,
            text="Extra Features",
            font=("Arial", 14, "bold")
        ).pack(pady=10)
        
        ttk.Label(
            self.content_frame,
            text="Configure additional features for your presentation:",
            font=("Arial", 10),
            foreground="#666"
        ).pack(pady=5)
        
        # Save function
        def save_super_research_options():
            self.config_data['super_research'] = False  # Always disabled
            self.config_data['auto_images'] = self.auto_images_var.get()
        
        # Initialize variables (if coming from back button)
        if not hasattr(self, 'super_research_var'):
            self.super_research_var = tk.BooleanVar(value=False)
        if not hasattr(self, 'auto_images_var'):
            self.auto_images_var = tk.BooleanVar(value=self.config_data.get('auto_images', False))
        
        # SUPER RESEARCH MODE - DISABLED (Coming Soon)
        super_research_frame = ttk.LabelFrame(self.content_frame, text="Super Research Mode", padding=15)
        super_research_frame.pack(pady=15, padx=50, fill=tk.X)
        
        super_cb = ttk.Checkbutton(
            super_research_frame,
            text="Enable Super Research (Coming Soon)",
            variable=self.super_research_var,
            state=tk.DISABLED
        )
        super_cb.pack(anchor=tk.W, pady=5)
        
        ttk.Label(
            super_research_frame,
            text="This feature is currently being improved.\n"
                 "It will be available in a future update.",
            font=("Arial", 9),
            foreground="#999",
            justify=tk.LEFT
        ).pack(anchor=tk.W, padx=20, pady=5)
        
        # Save initial values
        save_super_research_options()
    
    def show_theme_selection(self):
        """Step 2: Theme gallery - 6 Premium Themes with scrollbar"""
        ttk.Label(
            self.content_frame,
            text="Choose a Presentation Theme",
            font=("Arial", 14, "bold")
        ).pack(pady=10)
        
        ttk.Label(
            self.content_frame,
            text="Select a premium visual style for your slides:",
            font=("Arial", 10),
            foreground="#666"
        ).pack(pady=5)
        
        # Create scrollable frame for themes
        scroll_container = ttk.Frame(self.content_frame)
        scroll_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        canvas = tk.Canvas(scroll_container, height=280, highlightthickness=0)
        
        scrollbar = tk.Scrollbar(
            scroll_container,
            orient="vertical",
            command=canvas.yview,
            bg='#D0D0D0',
            troughcolor='#F0F0F0',
            activebackground='#A0A0A0',
            width=16
        )
        
        themes_frame = ttk.Frame(canvas)
        
        themes_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=themes_frame, anchor="nw",
                             width=canvas.winfo_reqwidth())
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Mousewheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        def bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        def unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
        
        canvas.bind("<Enter>", bind_mousewheel)
        canvas.bind("<Leave>", unbind_mousewheel)
        
        # Resize canvas interior width dynamically
        def on_canvas_configure(event):
            canvas.itemconfig(canvas.find_all()[0], width=event.width)
        canvas.bind("<Configure>", on_canvas_configure)
        
        # Theme radio buttons
        self.theme_var = tk.StringVar(value=self.config_data.get('theme_choice', '1'))
        
        for key, theme in BUILTIN_THEMES.items():
            theme_frame = ttk.Frame(themes_frame, relief=tk.RIDGE, borderwidth=1)
            theme_frame.pack(fill=tk.X, padx=10, pady=4)
            
            radio = ttk.Radiobutton(
                theme_frame,
                text=f"{theme['name']}",
                variable=self.theme_var,
                value=key
            )
            radio.pack(anchor=tk.W, padx=10, pady=5)
            
            ttk.Label(
                theme_frame,
                text=theme['description'],
                font=("Arial", 9),
                foreground="#666"
            ).pack(anchor=tk.W, padx=30, pady=(0, 5))
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def show_export_options(self):
        """Step 3: Export format selection"""
        ttk.Label(
            self.content_frame,
            text="Select Export Formats",
            font=("Arial", 14, "bold")
        ).pack(pady=20)
        
        ttk.Label(
            self.content_frame,
            text="Choose which file formats you want to generate:",
            font=("Arial", 10)
        ).pack(pady=10)
        
        # Options frame
        options_frame = ttk.Frame(self.content_frame)
        options_frame.pack(pady=20)
        
        # PPTX (always included)
        ttk.Label(
            options_frame,
            text="[OK] PPTX (PowerPoint 2007+)",
            font=("Arial", 11, "bold"),
            foreground="green"
        ).pack(anchor=tk.W, pady=5)
        ttk.Label(
            options_frame,
            text="   Modern PowerPoint format - Always included",
            font=("Arial", 9),
            foreground="#666"
        ).pack(anchor=tk.W, pady=(0, 15))
        
        # PPT option
        self.ppt_var = tk.BooleanVar(value=self.config_data.get('export_ppt', True))
        ppt_check = ttk.Checkbutton(
            options_frame,
            text="PPT (PowerPoint 97-2003)",
            variable=self.ppt_var
        )
        ppt_check.pack(anchor=tk.W, pady=5)
        ttk.Label(
            options_frame,
            text="   Legacy format for older PowerPoint versions",
            font=("Arial", 9),
            foreground="#666"
        ).pack(anchor=tk.W, pady=(0, 15))
        
        # PDF option
        self.pdf_var = tk.BooleanVar(value=self.config_data.get('export_pdf', False))
        pdf_check = ttk.Checkbutton(
            options_frame,
            text="PDF (Portable Document Format)",
            variable=self.pdf_var
        )
        pdf_check.pack(anchor=tk.W, pady=5)
        ttk.Label(
            options_frame,
            text="   Read-only format for sharing and printing",
            font=("Arial", 9),
            foreground="#666"
        ).pack(anchor=tk.W, pady=(0, 15))
        
        # Note
        note_frame = ttk.LabelFrame(self.content_frame, text="Note", padding=10)
        note_frame.pack(pady=20, padx=50, fill=tk.X)
        
        ttk.Label(
            note_frame,
            text="PPT and PDF conversion requires Microsoft PowerPoint to be installed.",
            font=("Arial", 9),
            foreground="#666"
        ).pack()
    
    def show_progress(self):
        """Step 4: Generation progress"""
        ttk.Label(
            self.content_frame,
            text="Generating Your Presentation...",
            font=("Arial", 14, "bold")
        ).pack(pady=20)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(
            self.content_frame,
            variable=self.progress_var,
            maximum=100,
            length=500,
            mode='determinate'
        )
        progress_bar.pack(pady=20)
        
        # Status label
        self.status_label = ttk.Label(
            self.content_frame,
            text="Initializing...",
            font=("Arial", 10)
        )
        self.status_label.pack(pady=10)
        
        # Details text
        self.details_text = tk.Text(
            self.content_frame,
            height=12,
            width=70,
            font=("Consolas", 9),
            wrap=tk.WORD,
            bg="#F5F5F5"
        )
        self.details_text.pack(pady=10)
        self.details_text.config(state=tk.DISABLED)
        
        # Start generation in background thread
        self.generation_thread = threading.Thread(target=self.run_generation, daemon=True)
        self.generation_thread.start()
    
    def show_completion(self):
        """Step 5: Completion screen"""
        ttk.Label(
            self.content_frame,
            text="[SUCCESS] Presentation Generated!",
            font=("Arial", 16, "bold"),
            foreground="green"
        ).pack(pady=30)
        
        # Summary
        summary_frame = ttk.LabelFrame(self.content_frame, text="Generated Files", padding=15)
        summary_frame.pack(pady=10, padx=50, fill=tk.BOTH, expand=True)
        
        if hasattr(self, 'output_files'):
            for file_path in self.output_files:
                file_frame = ttk.Frame(summary_frame)
                file_frame.pack(fill=tk.X, pady=5)
                
                ttk.Label(
                    file_frame,
                    text=f"FILE: {os.path.basename(file_path)}",
                    font=("Arial", 10)
                ).pack(side=tk.LEFT)
                
                ttk.Button(
                    file_frame,
                    text="Open Folder",
                    command=lambda p=file_path: os.startfile(os.path.dirname(p))
                ).pack(side=tk.RIGHT)
        
        # Next steps
        next_steps_text = tk.Text(
            self.content_frame,
            height=5,
            font=("Arial", 10),
            wrap=tk.WORD,
            borderwidth=0
        )
        next_steps_text.insert("1.0", "What's next?\n\n- Open the files in PowerPoint\n- Review and customize as needed\n- Share your presentation!")
        next_steps_text.tag_configure("center", justify="center")
        next_steps_text.tag_add("center", "1.0", "end")
        next_steps_text.config(state=tk.DISABLED)
        next_steps_text.pack(pady=20)
        
        # FINISH BUTTON - Opens output folder automatically
        finish_button_frame = ttk.Frame(self.content_frame)
        finish_button_frame.pack(pady=20)
        
        def open_output_folder_and_exit():
            """Open the output folder and close the application"""
            try:
                if hasattr(self, 'output_files') and self.output_files:
                    # Get the directory of the first output file
                    output_dir = os.path.dirname(self.output_files[0])
                    # Open the directory in Windows Explorer
                    os.startfile(output_dir)
                else:
                    # Fallback: open the default output directory
                    output_dir = os.path.join(os.getcwd(), 'output')
                    if os.path.exists(output_dir):
                        os.startfile(output_dir)
            except Exception as e:
                print(f"Could not open folder: {e}")
            
            # Close the application
            self.master.quit()
        
        finish_btn = ttk.Button(
            finish_button_frame,
            text="Finish & Open Folder",
            command=open_output_folder_and_exit
        )
        finish_btn.pack()
        
        # CRITICAL: Explicitly disable Next and Back buttons to prevent looping
        self.next_btn.config(state=tk.DISABLED)
        self.back_btn.config(state=tk.DISABLED)
    
    def log_message(self, message):
        """Add message to details log"""
        if hasattr(self, 'details_text'):
            self.details_text.config(state=tk.NORMAL)
            self.details_text.insert(tk.END, message + "\n")
            self.details_text.see(tk.END)
            self.details_text.config(state=tk.DISABLED)
    
    def update_status(self, status, progress=None):
        """Update status label and progress bar with smooth animation"""
        if hasattr(self, 'status_label'):
            self.status_label.config(text=status)
        
        if progress is not None and hasattr(self, 'progress_var'):
            # Animate progress bar smoothly from current to target
            current = self.progress_var.get()
            target = progress
            
            if abs(target - current) > 0.5:
                # Calculate smooth animation steps
                diff = target - current
                steps = max(10, min(30, int(abs(diff))))  # 10-30 steps
                increment = diff / steps
                
                for i in range(steps):
                    current += increment
                    self.progress_var.set(current)
                    self.update()
                    time.sleep(0.01)  # 10ms delay for smooth animation
            
            # Set final value
            self.progress_var.set(target)
    
    def check_gemini_login(self):
        """Check if user is logged into Gemini, prompt if not"""
        self.log_message("Checking Gemini login status...")
        
        try:
            researcher = GeminiResearcher()
            researcher.init_browser()
            
            if not researcher.navigate_to_gemini():
                # Not logged in - keep browser open for manual login
                self.log_message("Not logged into Gemini")
                self.log_message("Please log in to Gemini in the browser window...")
                
                # Wait for login
                messagebox.showinfo(
                    "Gemini Login Required",
                    "Please log in to Google Gemini in the browser window.\n\nClick OK after you've logged in."
                )
                
                # Give user time to complete login and retry verification
                self.log_message("Verifying login...")
                import time
                time.sleep(2)  # Wait for page to load
                
                # Retry verification multiple times
                login_verified = False
                for attempt in range(3):
                    self.log_message(f"Verification attempt {attempt + 1}/3...")
                    if researcher.navigate_to_gemini():
                        login_verified = True
                        break
                    time.sleep(2)
                
                if not login_verified:
                    raise Exception("Failed to verify Gemini login. Please make sure you're logged in.")
                
                self.log_message("[OK] Logged in successfully")
            else:
                self.log_message("[OK] Already logged into Gemini")
            
            researcher.close_browser()
            return True
            
        except Exception as e:
            self.log_message(f"[X] Login check failed: {e}")
            return False
    
    def start_continuous_progress(self, start_percent, end_percent, duration_seconds):
        """Start continuous progress animation from start to end percent over duration"""
        self._progress_stop_flag = False
        self._progress_thread = threading.Thread(
            target=self._animate_continuous_progress,
            args=(start_percent, end_percent, duration_seconds),
            daemon=True
        )
        self._progress_thread.start()
    
    def stop_continuous_progress(self):
        """Stop continuous progress animation"""
        self._progress_stop_flag = True
        if hasattr(self, '_progress_thread'):
            self._progress_thread.join(timeout=1)
    
    def _animate_continuous_progress(self, start_percent, end_percent, duration_seconds):
        """Animate progress bar continuously"""
        steps = duration_seconds  # 1% per second
        increment = (end_percent - start_percent) / steps
        current = start_percent
        
        for i in range(steps):
            if self._progress_stop_flag:
                break
            current += increment
            self.progress_var.set(current)
            self.update()
            time.sleep(1)  # 1 second per step
    
    def run_generation(self):
        """Main generation process (runs in background thread)"""
        try:
            # Collect configuration
            topic = self.config_data['topic']
            theme_choice = self.config_data['theme_choice']
            export_ppt_enabled = self.config_data['export_ppt']
            export_pdf_enabled = self.config_data['export_pdf']
            
            self.output_files = []
            
            # Create output directory in package/output folder
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_topic = topic.strip().replace(" ", "_").replace("\\", "").replace("/", "")
            
            # Get executable directory (works for both .py and .exe)
            if getattr(sys, 'frozen', False):
                # Running as compiled executable
                exe_dir = os.path.dirname(sys.executable)
            else:
                # Running as script
                exe_dir = os.path.dirname(os.path.abspath(__file__))
            
            # Use package/output folder
            output_base = os.path.join(exe_dir, 'output')
            output_dir = os.path.join(output_base, f'{safe_topic}_{timestamp}')
            os.makedirs(output_dir, exist_ok=True)
            charts_dir = os.path.join(output_dir, 'charts')
            os.makedirs(charts_dir, exist_ok=True)
            
            self.log_message(f"Output directory: {output_dir}")
            
            # Step 1: Check Gemini login and Research (0-20%)
            self.update_status("Checking Gemini login...", 0)
            if not self.check_gemini_login():
                raise Exception("Gemini login required")
            
            self.update_status("Conducting AI research...", 5)
            self.log_message(f"\n=== Step 1/5: Research ===")
            self.log_message(f"Topic: {topic}")
            
            # Start continuous progress animation (5% to 20% over ~80 seconds)
            self.start_continuous_progress(5, 20, 80)
            
            research_data, researcher = gather_research(
                topic,
                research_level=self.config_data['research_level'],
                super_research=self.config_data.get('super_research', False),
                progress_callback=None,  # Animation thread handles progress
                output_folder=output_dir  # Pass output folder for super_research
            )
            
            # Stop animation and set to target
            self.stop_continuous_progress()
            
            if not research_data:
                raise Exception("Research failed")
            
            # Pass output_dir for debug logging
            research_data['output_dir'] = output_dir
            
            self.update_status("Research completed!", 20)
            self.log_message(f"[+] Found {research_data.get('total_sections', 0)} sections")
            
            # Step 2: Generate Content (20-40%)
            self.update_status("Generating presentation content...", 25)
            self.log_message(f"\n=== Step 2/5: Content Generation ===")
            
            slides, viz_suggestions = generate_presentation_content(topic, research_data, researcher)
            
            self.update_status("Content generated!", 40)
            self.log_message(f"[+] Created {len(slides)} slides")
            
            # Step 3: Visualizations (40-60%)
            self.update_status("Downloading visualizations...", 45)
            self.log_message(f"\n=== Step 3/5: Visualizations ===")
            
            image_paths = generate_visualizations(topic, researcher, charts_dir)
            
            self.update_status("Visualizations ready!", 60)
            self.log_message(f"[+] Downloaded {len(image_paths)} images")
            
            # Step 4: Build Presentation (60-75%)
            self.update_status("Building PowerPoint...", 65)
            self.log_message(f"\n=== Step 4/5: Building Presentation ===")
            
            output_filename = os.path.join(output_dir, f'{safe_topic}_presentation.pptx')
            prs = build_presentation(topic, slides, image_paths, output_filename, self.config_data)
            
            # Apply theme
            self.log_message(f"Applying theme: {BUILTIN_THEMES[theme_choice]['name']}")
            apply_theme_to_presentation(prs, theme_choice)
            prs.save(output_filename)
            
            self.output_files.append(output_filename)
            self.update_status("PowerPoint created!", 75)
            self.log_message(f"[+] Saved: {os.path.basename(output_filename)}")
            
            # Step 5: Research Document (75-85%)
            self.update_status("Creating research document...", 78)
            self.log_message(f"\n=== Step 5/5: Research Document ===")
            
            research_doc_path = generate_supplemental_research(topic, research_data, output_dir)
            if research_doc_path:
                self.output_files.append(research_doc_path)
                self.log_message(f"[+] Saved: {os.path.basename(research_doc_path)}")
            
            self.update_status("Research document created!", 85)
            
            # Export formats (85-100%)
            current_progress = 85
            
            if export_ppt_enabled:
                self.update_status("Exporting to PPT format...", current_progress)
                self.log_message(f"\n=== Exporting to PPT ===")
                
                ppt_file = export_to_ppt(output_filename)
                if ppt_file:
                    self.output_files.append(ppt_file)
                    self.log_message(f"[+] Saved: {os.path.basename(ppt_file)}")
                
                current_progress += 7
                self.update_status("PPT export complete!", current_progress)
            
            if export_pdf_enabled:
                self.update_status("Exporting to PDF format...", current_progress)
                self.log_message(f"\n=== Exporting to PDF ===")
                
                pdf_file = export_to_pdf(output_filename)
                if pdf_file:
                    self.output_files.append(pdf_file)
                    self.log_message(f"[+] Saved: {os.path.basename(pdf_file)}")
                
                current_progress += 8
                self.update_status("PDF export complete!", current_progress)
            
            # Close browser
            if researcher:
                researcher.close_browser()
            
            # Complete!
            self.update_status("[+] All done!", 100)
            self.log_message(f"\n{'='*50}")
            self.log_message(f"[+] GENERATION COMPLETE!")
            self.log_message(f"{'='*50}")
            self.log_message(f"Output directory: {output_dir}")
            
            # Move to completion step (step 7, NOT step 6 which is progress!)
            self.after(2000, lambda: self.show_step(7))
            
        except Exception as e:
            self.log_message(f"\n[!] ERROR: {str(e)}")
            messagebox.showerror("Generation Failed", f"An error occurred:\n\n{str(e)}")
            self.after(0, lambda: self.show_step(0))  # Go back to start
    
    def go_next(self):
        """Next button handler"""
        if self.current_step == 1:
            # Validate topic
            topic = self.topic_var.get().strip()
            if not topic:
                messagebox.showwarning("Topic Required", "Please enter a presentation topic.")
                return
            self.config_data['topic'] = topic
        
        elif self.current_step == 2:
            # Step 2: Advanced Options - values already auto-saved
            pass
        
        elif self.current_step == 3:
            # Step 3: Super Research Options - values already auto-saved
            pass
        
        elif self.current_step == 4:
            # Step 4: Save theme choice
            self.config_data['theme_choice'] = self.theme_var.get()
        
        elif self.current_step == 5:
            # Step 5: Save export preferences
            self.config_data['export_ppt'] = self.ppt_var.get()
            self.config_data['export_pdf'] = self.pdf_var.get()
        
        # Move to next step
        if self.current_step < self.total_steps - 1:
            self.show_step(self.current_step + 1)
    
    def go_back(self):
        """Back button handler"""
        if self.current_step > 0:
            self.show_step(self.current_step - 1)
    
    def finish(self):
        """Finish button handler"""
        # Properly close the application
        self.quit()
        self.destroy()


def main():
    """Launch the wizard"""
    app = PPTWizard()
    app.mainloop()


if __name__ == "__main__":
    main()
