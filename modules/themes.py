"""
Premium PowerPoint Themes Module
6 professionally designed theme palettes for stunning presentations.
Each theme controls ALL visual elements: backgrounds, accents, text colors, fonts.
"""

from pptx.dml.color import RGBColor
from colorama import Fore, Style


# ============================================================================
# PREMIUM THEME DEFINITIONS
# ============================================================================

PREMIUM_THEMES = {
    '1': {
        'name': 'Midnight',
        'description': 'Dark blue-black with cyan & purple accents',
        'bg_primary': RGBColor(15, 15, 25),
        'bg_card': RGBColor(30, 32, 48),
        'accent_1': RGBColor(100, 149, 237),     # Cornflower blue
        'accent_2': RGBColor(0, 206, 209),        # Cyan
        'accent_3': RGBColor(147, 112, 219),       # Purple
        'text_primary': RGBColor(240, 240, 248),
        'text_secondary': RGBColor(180, 185, 200),
        'text_dim': RGBColor(120, 125, 145),
        'title_bar_color': RGBColor(0, 206, 209),  # Accent bar under titles
        'font_heading': 'Segoe UI',
        'font_body': 'Segoe UI',
    },
    '2': {
        'name': 'Arctic',
        'description': 'Clean white with blue & teal modern accents',
        'bg_primary': RGBColor(248, 249, 252),
        'bg_card': RGBColor(235, 240, 248),
        'accent_1': RGBColor(37, 99, 235),         # Royal blue
        'accent_2': RGBColor(14, 165, 165),         # Teal
        'accent_3': RGBColor(99, 102, 241),         # Indigo
        'text_primary': RGBColor(30, 30, 45),
        'text_secondary': RGBColor(75, 85, 99),
        'text_dim': RGBColor(140, 148, 160),
        'title_bar_color': RGBColor(37, 99, 235),
        'font_heading': 'Segoe UI',
        'font_body': 'Segoe UI',
    },
    '3': {
        'name': 'Ember',
        'description': 'Dark charcoal with warm orange & amber accents',
        'bg_primary': RGBColor(24, 24, 27),
        'bg_card': RGBColor(39, 39, 42),
        'accent_1': RGBColor(249, 115, 22),         # Orange
        'accent_2': RGBColor(245, 158, 11),          # Amber
        'accent_3': RGBColor(239, 68, 68),            # Red
        'text_primary': RGBColor(245, 245, 245),
        'text_secondary': RGBColor(180, 180, 185),
        'text_dim': RGBColor(130, 130, 135),
        'title_bar_color': RGBColor(249, 115, 22),
        'font_heading': 'Segoe UI',
        'font_body': 'Segoe UI',
    },
    '4': {
        'name': 'Forest',
        'description': 'Deep green-black with emerald & lime accents',
        'bg_primary': RGBColor(10, 20, 15),
        'bg_card': RGBColor(25, 40, 30),
        'accent_1': RGBColor(16, 185, 129),          # Emerald
        'accent_2': RGBColor(132, 204, 22),           # Lime
        'accent_3': RGBColor(34, 197, 94),             # Green
        'text_primary': RGBColor(240, 248, 240),
        'text_secondary': RGBColor(180, 200, 185),
        'text_dim': RGBColor(120, 145, 130),
        'title_bar_color': RGBColor(16, 185, 129),
        'font_heading': 'Segoe UI',
        'font_body': 'Segoe UI',
    },
    '5': {
        'name': 'Royal',
        'description': 'Deep purple with gold & lavender elegant accents',
        'bg_primary': RGBColor(20, 10, 35),
        'bg_card': RGBColor(35, 25, 55),
        'accent_1': RGBColor(168, 85, 247),           # Violet
        'accent_2': RGBColor(251, 191, 36),            # Gold/Amber
        'accent_3': RGBColor(196, 181, 253),           # Lavender
        'text_primary': RGBColor(245, 240, 255),
        'text_secondary': RGBColor(190, 180, 210),
        'text_dim': RGBColor(140, 130, 165),
        'title_bar_color': RGBColor(251, 191, 36),
        'font_heading': 'Segoe UI',
        'font_body': 'Segoe UI',
    },
    '6': {
        'name': 'Scholar',
        'description': 'Clean white professional education theme',
        'bg_primary': RGBColor(255, 255, 255),
        'bg_card': RGBColor(243, 244, 246),
        'accent_1': RGBColor(31, 41, 55),              # Charcoal
        'accent_2': RGBColor(59, 130, 246),             # Blue
        'accent_3': RGBColor(107, 114, 128),            # Gray
        'text_primary': RGBColor(17, 24, 39),
        'text_secondary': RGBColor(55, 65, 81),
        'text_dim': RGBColor(107, 114, 128),
        'title_bar_color': RGBColor(59, 130, 246),
        'font_heading': 'Segoe UI',
        'font_body': 'Segoe UI',
    },
}

# Backward compatibility alias
BUILTIN_THEMES = PREMIUM_THEMES


def get_theme(theme_choice):
    """Get theme palette by ID, default to Midnight."""
    return PREMIUM_THEMES.get(str(theme_choice), PREMIUM_THEMES['1'])


def display_theme_menu():
    """Display available themes for user selection."""
    print(f"\n{Fore.CYAN}{'=' * 60}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}SELECT PRESENTATION THEME{Style.RESET_ALL}")
    print(f"{Fore.CYAN}{'=' * 60}{Style.RESET_ALL}\n")
    
    for key, theme in PREMIUM_THEMES.items():
        print(f"{Fore.GREEN}{key}.{Style.RESET_ALL} {Fore.WHITE}{theme['name']}{Style.RESET_ALL}")
        print(f"   -> {theme['description']}\n")
    
    print(f"{Fore.CYAN}{'=' * 60}{Style.RESET_ALL}\n")


def get_theme_choice():
    """Get theme selection from user."""
    display_theme_menu()
    
    while True:
        choice = input(f"{Fore.YELLOW}Enter theme number (1-6) [default: 1]: {Style.RESET_ALL}").strip()
        
        if not choice:
            choice = '1'
        
        if choice in PREMIUM_THEMES:
            theme = PREMIUM_THEMES[choice]
            print(f"{Fore.GREEN}[+] Selected: {theme['name']}{Style.RESET_ALL}\n")
            return choice, theme
        else:
            print(f"{Fore.RED}Invalid choice. Please enter a number between 1-6.{Style.RESET_ALL}\n")


def apply_theme_to_presentation(prs, theme_choice):
    """
    Apply theme is now handled during slide creation by ppt_builder.
    This function exists for backward compatibility.
    """
    theme = get_theme(theme_choice)
    print(f"{Fore.CYAN}[*] Theme: {theme['name']} (applied during slide creation){Style.RESET_ALL}")
    print(f"{Fore.GREEN}[+] Theme applied successfully{Style.RESET_ALL}")
