"""
Export Module
Handles PPT and PDF export using PowerPoint COM automation
"""

import os
import sys
from colorama import Fore, Style


def export_to_ppt(pptx_path):
    """
    Convert PPTX to PPT (PowerPoint 97-2003 format)
   
    Args:
        pptx_path: Path to the PPTX file
   
    Returns:
        Path to generated PPT file or None if failed
    """
    try:
        print(f"{Fore.CYAN}[*] Converting to PPT format...{Style.RESET_ALL}")
       
        # Use comtypes for Windows PowerPoint automation
        import comtypes.client
       
        # Start PowerPoint
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
       
        # Open the PPTX file
        presentation = powerpoint.Presentations.Open(pptx_path)
       
        # Save as PPT (format 1 = PowerPoint 97-2003)
        ppt_path = pptx_path.replace('.pptx', '.ppt')
        presentation.SaveAs(ppt_path, 1)
       
        # Close presentation
        presentation.Close()
        powerpoint.Quit()
       
        print(f"{Fore.GREEN}[+] PPT file created: {os.path.basename(ppt_path)}{Style.RESET_ALL}")
        return ppt_path
       
    except ImportError:
        print(f"{Fore.YELLOW}[!] comtypes not installed. Installing...{Style.RESET_ALL}")
        os.system("pip install comtypes")
        return export_to_ppt(pptx_path)  # Retry after install
       
    except Exception as e:
        print(f"{Fore.RED}[!] PPT conversion failed: {e}{Style.RESET_ALL}")
        print(f"{Fore.YELLOW}  Note: PowerPoint must be installed for PPT conversion{Style.RESET_ALL}")
        return None


def export_to_pdf(pptx_path):
    """
    Convert PPTX to PDF using PowerPoint COM automation
   
    Args:
        pptx_path: Path to the PPTX file
   
    Returns:
        Path to generated PDF file or None if failed
    """
    try:
        print(f"{Fore.CYAN}[*] Converting to PDF format...{Style.RESET_ALL}")
       
        # Use comtypes for Windows PowerPoint automation
        import comtypes.client
       
        # Start PowerPoint
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
       
        # Open the PPTX file
        presentation = powerpoint.Presentations.Open(pptx_path)
       
        # Save as PDF (format 32 = PDF)
        pdf_path = pptx_path.replace('.pptx', '.pdf')
        presentation.SaveAs(pdf_path, 32)
       
        # Close presentation
        presentation.Close()
        powerpoint.Quit()
       
        print(f"{Fore.GREEN}[+] PDF file created: {os.path.basename(pdf_path)}{Style.RESET_ALL}")
        return pdf_path
       
    except ImportError:
        print(f"{Fore.YELLOW}[!] comtypes not installed. Installing...{Style.RESET_ALL}")
        os.system("pip install comtypes")
        return export_to_pdf(pptx_path)  # Retry after install
       
    except Exception as e:
        print(f"{Fore.RED}[!] PDF conversion failed: {e}{Style.RESET_ALL}")
        print(f"{Fore.YELLOW}  Note: PowerPoint must be installed for PDF conversion{Style.RESET_ALL}")
        return None


def get_export_preferences():
    """
    Ask user which export formats they want
   
    Returns:
        dict with 'ppt' and 'pdf' boolean flags
    """
    print(f"\n{Fore.CYAN}{'=' * 80}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}EXPORT OPTIONS{Style.RESET_ALL}")
    print(f"{Fore.CYAN}{'=' * 80}{Style.RESET_ALL}\n")
   
    preferences = {
        'ppt': False,
        'pdf': False
    }
   
    # Ask about PPT format
    while True:
        ppt_choice = input(f"{Fore.YELLOW}Export as PPT format (PowerPoint 97-2003)? (y/n) [default: y]: {Style.RESET_ALL}").strip().lower()
       
        if not ppt_choice or ppt_choice == 'y':
            preferences['ppt'] = True
            break
        elif ppt_choice == 'n':
            break
        else:
            print(f"{Fore.RED}Invalid input. Please enter 'y' or 'n'.{Style.RESET_ALL}")
   
    # Ask about PDF format
    while True:
        pdf_choice = input(f"{Fore.YELLOW}Export as PDF format? (y/n) [default: n]: {Style.RESET_ALL}").strip().lower()
       
        if not pdf_choice or pdf_choice == 'n':
            break
        elif pdf_choice == 'y':
            preferences['pdf'] = True
            break
        else:
            print(f"{Fore.RED}Invalid input. Please enter 'y' or 'n'.{Style.RESET_ALL}")
   
    print(f"\n{Fore.GREEN}[+] Export preferences saved{Style.RESET_ALL}\n")
    return preferences
