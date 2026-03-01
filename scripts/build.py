"""
Build script to create standalone executable
Packages the AI PowerPoint Generator into a single .exe file

Usage: python build.py
"""

import os
import sys
import subprocess
import shutil

def main():
    print("="*70)
    print("AI PowerPoint Generator - Build Script")
    print("="*70)
    print()
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
        print("[OK] PyInstaller found")
    except ImportError:
        print("[X] PyInstaller not found")
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("[OK] PyInstaller installed")
    
    print()
    print("Building executable...")
    print()
    
    # PyInstaller command
    cmd = [
        "pyinstaller",
        "--onefile",                    # Single executable
        "--windowed",                   # No console window
        "--name=FlashPoint_AI",         # Executable name (Renamed to fix cache)
        "--icon=assets/icon.ico",              # Application icon (Original)
        "--add-data=modules;modules",   # Include modules folder
        "--add-data=config/prompt_lvl_1;.",    # Include prompt templates
        "--add-data=config/prompt_lvl_2;.",
        "--add-data=config/prompt_lvl_3;.",
        "--add-data=config/slide_formating_prompt;.",  # Slide formatting prompt
        "--hidden-import=selenium",
        "--hidden-import=webdriver_manager",
        "--hidden-import=beautifulsoup4",
        "--hidden-import=python-pptx",
        "--hidden-import=python-docx",
        "--hidden-import=matplotlib",
        "--hidden-import=PIL",
        "--hidden-import=comtypes",
        "--clean",                      # Clean before building
        "ppt_wizard.py"
    ]
    
    try:
        # CLEANUP OLD BUILDS
        print("Cleaning up old builds...")
        cleanup_files = [
            "package\\PPT_Generator.exe",
            "dist\\PPT_Generator.exe",
            "package\\FlashPoint_v2.exe",
            "dist\\FlashPoint_v2.exe",
            "package\\FlashPoint_v3.exe",
            "dist\\FlashPoint_v3.exe",
            "package\\FlashPoint.exe",    # Remove previous build
            "dist\\FlashPoint.exe"
        ]
        for f in cleanup_files:
            if os.path.exists(f):
                try:
                    os.remove(f)
                    print(f"[CLEANUP] Removed: {f}")
                except Exception as e:
                    print(f"[WARNING] Could not remove {f}: {e}")

        subprocess.check_call(cmd)
        print()
        print("="*70)
        print("[OK] BUILD SUCCESSFUL!")
        print("="*70)
        print()
        print(f"Executable created: dist\\FlashPoint_AI.exe")
        print()
        
        # Copy to package folder
        print("Copying to package folder...")
        
        package_dir = "package"
        os.makedirs(package_dir, exist_ok=True)
        
        # Copy executable
        shutil.copy("dist\\FlashPoint_AI.exe", f"{package_dir}\\FlashPoint_AI.exe")
        print(f"[OK] Copied: {package_dir}\\FlashPoint_AI.exe")
        
        # Create output folder
        output_dir = f"{package_dir}\\output"
        os.makedirs(output_dir, exist_ok=True)
        print(f"[OK] Created: {output_dir}")
        
        # Create config folder for prompt customization
        config_dir = f"{package_dir}\\config"
        os.makedirs(config_dir, exist_ok=True)
        print(f"[OK] Created: {config_dir}")
        
        # Copy prompt files to config folder ONLY
        print("\\nCopying prompt templates to config folder...")
        for level in [1, 2, 3]:
            prompt_file = f"config\\prompt_lvl_{level}"
            if os.path.exists(prompt_file):
                shutil.copy(prompt_file, f"{config_dir}\\prompt_lvl_{level}")
                print(f"[OK] Copied: config\\prompt_lvl_{level}")
        
        # Copy slide formatting prompt
        if os.path.exists("config\\slide_formating_prompt"):
            shutil.copy("config\\slide_formating_prompt", f"{config_dir}\\slide_formating_prompt")
            print(f"[OK] Copied: config\\slide_formating_prompt")
        
        # Copy customization guide
        if os.path.exists("docs\\PROMPT_CUSTOMIZATION.md"):
            shutil.copy("docs\\PROMPT_CUSTOMIZATION.md", f"{config_dir}\\PROMPT_CUSTOMIZATION.md")
            print(f"[OK] Copied: config\\PROMPT_CUSTOMIZATION.md")
        
        print("\\nTIP: Edit files in the 'config' folder to customize research prompts!")
        
        # Create README
        readme_content = """FlashPoint - AI PowerPoint Generator
==========================================

Made by Luqman (2024-ag-8738 DVM 4th Semester)
Discord Support: sad_memer.

------------------------------------------------------------------
👋 INTRODUCTION
------------------------------------------------------------------
Hello! My name is Luqman from DVM 4th Semester.

As we all know, everyone hates doing presentations. The research, the content, 
and the diagrams... all require some level of computer expertise.

But today I would like to share a program made by me to somewhat help you guys 
in generating, making, and collecting data and content for your presentation slides!

Here is an early-beta demo of my program.

------------------------------------------------------------------
🚀 HOW TO RUN IT (Step-by-Step)
------------------------------------------------------------------
1. Run the Program:
   Pretty simple! Just double-click on the "FlashPoint.exe" file.

2. Enter Topic:
   A window opens asking you about your topic for the presentation.

3. Specify Slides & Depth:
   - Specify the number of slides you wish to make (e.g., 20-25).
   - Select research level (Basic, Professional, Master Thesis).

4. Select Theme:
   Choose your preferred design theme from the gallery.

5. AI Research:
   The program automatically opens Gemini AI and does research for you on the regarding topic.

6. Auto-Collect Diagrams:
   The program automatically collects diagrams and charts related to the topic.

7. Output:
   Finally, your presentation will be prepared in the "output" folder containing:
   - A .pptx file with your slides
   - A PDF file of the slides
   - A charts folder for additional diagrams
   - A research document textual file containing extra info related to the topic

------------------------------------------------------------------
✨ KEY FEATURES
------------------------------------------------------------------
1. AI Deep Research: Uses Google Gemini for up-to-date content (2024-2026).
2. Smart Visualizations: Automatically finds charts and images.
3. Professional Themes: 10+ built-in styles.
4. Multi-Format: Exports PPTX, PDF, and Research Document.

------------------------------------------------------------------
🔧 TROUBLESHOOTING
------------------------------------------------------------------
- "Gemini Login Required": Ensure you are logged into Google in the opened browser.
- Browser Closes: Do not manually close the automation window!
- Stuck: Check internet connection or restart if hung for >5 mins.

------------------------------------------------------------------
CONTACT & SUPPORT
------------------------------------------------------------------
Start generating and enjoy!

For feedback or bugs:
Discord: sad_memer.
"""
        
        with open(f"{package_dir}\\README.txt", "w", encoding="utf-8") as f:
            f.write(readme_content)
        print(f"[OK] Created: {package_dir}\\README.txt")
        
        print()
        print("="*70)
        print("PACKAGE READY!")
        print("="*70)
        print()
        print(f"Package location: {os.path.abspath(package_dir)}")
        print()
        print("You can now:")
        print("1. Test the executable: package\\FlashPoint.exe")
        print("2. Share the entire 'package' folder with friends")
        print()
        
    except subprocess.CalledProcessError as e:
        print()
        print("="*70)
        print("[ERROR] BUILD FAILED")
        print("="*70)
        print()
        print(f"Error: {e}")
        print()
        print("Try running manually:")
        print(" ".join(cmd))
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
