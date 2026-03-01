"""
Research Module - Browser Automation with Gemini
Uses Selenium to automate Chrome and leverage Gemini's deep research feature
"""

import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
from colorama import Fore, Style
import re
import sys  # Added for path resolution in frozen app


# Load research prompt templates
def load_prompt_templates():
    """Load the 3 research prompt level templates from external files
    
    Priority order:
    1. config/ folder next to executable (recommended for user customization)
    2. Root directory next to executable (backward compatibility)
    3. Bundled files in _MEIPASS (default templates from installation)
    4. Development directory (when running as .py)
    """
    templates = {}
    
    for level in [1, 2, 3]:
        content = None
        
        # Priority 1: Check config folder next to executable (RECOMMENDED)
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
            config_prompt = os.path.join(exe_dir, 'config', f'prompt_lvl_{level}')
            if os.path.exists(config_prompt):
                try:
                    with open(config_prompt, 'r', encoding='utf-8') as f:
                        content = f.read()
                        print(f"{Fore.CYAN}[*] Loaded config/prompt_lvl_{level} (user-edited){Style.RESET_ALL}")
                except Exception as e:
                    print(f"{Fore.YELLOW}[!] Error reading config/prompt_lvl_{level}: {e}{Style.RESET_ALL}")
        
        # Priority 2: Check root directory next to executable (LEGACY)
        if content is None and getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
            legacy_prompt = os.path.join(exe_dir, f'prompt_lvl_{level}')
            if os.path.exists(legacy_prompt):
                try:
                    with open(legacy_prompt, 'r', encoding='utf-8') as f:
                        content = f.read()
                        print(f"{Fore.CYAN}[*] Loaded prompt_lvl_{level} from root (legacy location){Style.RESET_ALL}")
                except Exception as e:
                    print(f"{Fore.YELLOW}[!] Error reading prompt_lvl_{level}: {e}{Style.RESET_ALL}")
        
        # Priority 3: Check bundled files in _MEIPASS (or dev directory)
        if content is None:
            if getattr(sys, 'frozen', False):
                base_dir = sys._MEIPASS  # Bundled files
                prompt_file = os.path.join(base_dir, f'prompt_lvl_{level}')
            else:
                script_dir = os.path.dirname(os.path.abspath(__file__))
                base_dir = os.path.dirname(script_dir)  # Parent directory in dev
                prompt_file = os.path.join(base_dir, 'config', f'prompt_lvl_{level}')
            
            try:
                with open(prompt_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                    print(f"{Fore.GREEN}[+] Loaded prompt_lvl_{level} (bundled default){Style.RESET_ALL}")
            except Exception as e:
                print(f"{Fore.YELLOW}[!] Could not load prompt_lvl_{level}: {e}{Style.RESET_ALL}")
                content = "Generate comprehensive research on [Insert Topic Here]."
        
        templates[level] = content
    
    return templates

# Load templates at module initialization
PROMPT_TEMPLATES = load_prompt_templates()


class GeminiResearcher:
    """Handles browser automation for Gemini research"""
    
    def __init__(self):
        self.driver = None
        self.wait = None
        
    def init_browser(self, headless=False):
        """Initialize Chrome browser with Selenium"""
        print(f"{Fore.CYAN}[*] Initializing Chrome browser...{Style.RESET_ALL}")
        
        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument('--headless')
        
        # Add user data directory to persist login
        user_data_dir = os.path.join(os.path.expanduser('~'), '.ppt_generator_chrome')
        options.add_argument(f'--user-data-dir={user_data_dir}')
        options.add_argument('--profile-directory=Default')
        
        # Windows 11 compatibility fixes
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')  # Critical for Windows 11
        options.add_argument('--disable-software-rasterizer')
        options.add_argument('--no-first-run')
        options.add_argument('--no-default-browser-check')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        
        try:
            # Auto-install ChromeDriver
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.maximize_window()
            self.wait = WebDriverWait(self.driver, 30)
            
            print(f"{Fore.GREEN}[+] Browser initialized{Style.RESET_ALL}")
            print(f"{Fore.CYAN}   Session data saved in: {user_data_dir}{Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.RED}[!] Error initializing browser: {e}{Style.RESET_ALL}")
            print(f"{Fore.YELLOW}   Trying with additional compatibility options...{Style.RESET_ALL}")
            
            # Fallback: Try without profile persistence for Windows 11
            options = webdriver.ChromeOptions()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-dev-shm-usage')
            
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.maximize_window()
            self.wait = WebDriverWait(self.driver, 30)
            
            print(f"{Fore.GREEN}[+] Browser initialized (compatibility mode){Style.RESET_ALL}")
        
    def navigate_to_gemini(self):
        """Navigate to Gemini web interface"""
        print(f"{Fore.CYAN}[*] Navigating to Gemini...{Style.RESET_ALL}")
        
        try:
            self.driver.get("https://gemini.google.com/")
            time.sleep(3)
            
            # Check if login is required
            if self.check_login_required():
                print(f"{Fore.YELLOW}[!] Login required. Please log in manually.{Style.RESET_ALL}")
                print(f"{Fore.YELLOW}Press ENTER after you've logged in...{Style.RESET_ALL}")
                input()
                time.sleep(2)
            
            print(f"{Fore.GREEN}[+] Ready to use Gemini{Style.RESET_ALL}")
            return True
            
        except Exception as e:
            print(f"{Fore.RED}[!] Error navigating to Gemini: {e}{Style.RESET_ALL}")
            return False
    
    def check_login_required(self):
        """Check if user needs to log in - IMPROVED for Windows 11"""
        try:
            # Wait a moment for page to fully load
            time.sleep(3)
            
            # FIRST: Try to find the input box - if found and interactable, user IS logged in
            # This is the most reliable indicator
            input_selectors = [
                "rich-textarea",
                "div[contenteditable='true']",
                "textarea",
                "div[role='textbox']",
                ".ql-editor"
            ]
            
            for selector in input_selectors:
                try:
                    element = self.driver.find_element(By.CSS_SELECTOR, selector)
                    if element.is_displayed():
                        print(f"{Fore.GREEN}   [+] Found chat input field - user is logged in{Style.RESET_ALL}")
                        return False  # User IS logged in!
                except:
                    continue
            
            # SECOND: Check URL - if on accounts.google.com, definitely need login
            current_url = self.driver.current_url
            if 'accounts.google.com' in current_url:
                print(f"{Fore.YELLOW}   On Google login page - login required{Style.RESET_ALL}")
                return True
            
            # THIRD: Check for specific login page indicators (only if no input found)
            page_source = self.driver.page_source
            login_page_indicators = [
                'Choose an account',
                'Use another account',
                'Sign in to continue'
            ]
            
            for indicator in login_page_indicators:
                if indicator in page_source:
                    print(f"{Fore.YELLOW}   Login page detected: {indicator}{Style.RESET_ALL}")
                    return True
            
            # If we're on gemini.google.com but no input found, wait more and try again
            if 'gemini.google.com' in current_url:
                print(f"{Fore.CYAN}   On Gemini page, waiting for UI to load...{Style.RESET_ALL}")
                time.sleep(3)
                
                # Try again to find input
                for selector in input_selectors:
                    try:
                        element = self.driver.find_element(By.CSS_SELECTOR, selector)
                        if element.is_displayed():
                            print(f"{Fore.GREEN}   [+] Found chat input field after wait{Style.RESET_ALL}")
                            return False
                    except:
                        continue
            
            # Last resort - assume not logged in if nothing found
            print(f"{Fore.YELLOW}   Could not confirm login status{Style.RESET_ALL}")
            return True
                
        except Exception as e:
            print(f"{Fore.RED}   Error checking login: {e}{Style.RESET_ALL}")
            return True  # On error, assume login required
    
    def perform_deep_research(self, topic, research_level=3, super_research=False, progress_callback=None, output_folder=None):
        """Submit topic with deep research mode to Gemini
        
        Args:
            topic: The presentation topic
            research_level: 1 (Basic), 2 (Professional), or 3 (Master Thesis)
            super_research: Boolean - Enable academic API integration
            progress_callback: Optional callback for progress updates
            output_folder: Output folder path for super_research data
        """
        level_names = {1: 'Basic', 2: 'Professional', 3: 'Master Thesis'}
        print(f"{Fore.CYAN}[*] Researching: {topic}{Style.RESET_ALL}")
        print(f"{Fore.CYAN}   Research Level: {research_level} ({level_names.get(research_level, 'Unknown')}){Style.RESET_ALL}")
        
        if super_research:
            print(f"{Fore.MAGENTA}[*] Super Research Mode: ENABLED{Style.RESET_ALL}")
        
        try:
            # Find the input box (Gemini's prompt textarea)
            # This selector may need adjustment based on Gemini's current UI
            input_selectors = [
                "rich-textarea[placeholder*='Enter']",
                "div[contenteditable='true']",
                "textarea",
                ".ql-editor"
            ]
            
            prompt_box = None
            for selector in input_selectors:
                try:
                    prompt_box = self.wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    if prompt_box:
                        break
                except:
                    continue
            
            if not prompt_box:
                raise Exception("Could not find Gemini input box")
            
            # Click on the input box
            prompt_box.click()
            time.sleep(1)
            
            # Load the appropriate research prompt based on level
            research_prompt = PROMPT_TEMPLATES.get(research_level, PROMPT_TEMPLATES[3])
            # Replace topic placeholders (handle both old and new prompt formats)
            research_prompt = research_prompt.replace("[Insert Topic Here]", topic)
            research_prompt = research_prompt.replace("[Insert Topic]", topic)
            
            # SUPER RESEARCH MODE: DISABLED (Coming Soon)
            # This feature is temporarily disabled and will be re-enabled in a future update.
            
            
            # IMPROVED SUBMISSION: JavaScript Injection
            # Use JavaScript injection for large text (more reliable than typing 5000+ chars)
            print(f"{Fore.CYAN}   Injecting research prompt via JavaScript (fast & reliable)...{Style.RESET_ALL}")
            
            try:
                # Clear existing content first
                try:
                    prompt_box.clear()
                except:
                    self.driver.execute_script("arguments[0].value = '';", prompt_box)
                
                time.sleep(1)
                
                # Inject text directly into the element
                # Try multiple properties for compatibility
                self.driver.execute_script(
                    """
                    var element = arguments[0];
                    var text = arguments[1];
                    element.value = text;
                    element.innerText = text;
                    element.textContent = text;
                    element.dispatchEvent(new Event('input', { bubbles: true }));
                    element.dispatchEvent(new Event('change', { bubbles: true }));
                    """, 
                    prompt_box, 
                    research_prompt
                )
                time.sleep(2)
                
                # Click send button directly
                print(f"{Fore.CYAN}   Clicking send button...{Style.RESET_ALL}")
                
                # Try finding send button with multiple selectors
                send_clicked = False
                send_button_selectors = [
                    "button[aria-label*='Send']",
                    "button[class*='send-button']", 
                    "button[data-test-id='send-button']",
                    ".send-button",
                    "button[type='submit']"
                ]
                
                for selector in send_button_selectors:
                    try:
                        btns = self.driver.find_elements(By.CSS_SELECTOR, selector)
                        for btn in btns:
                            if btn.is_displayed():
                                # Scroll into view and click
                                self.driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                                time.sleep(0.5)
                                
                                # JavaScript click is often more reliable
                                self.driver.execute_script("arguments[0].click();", btn)
                                print(f"{Fore.GREEN}   [+] Clicked send button via JS{Style.RESET_ALL}")
                                send_clicked = True
                                break
                        if send_clicked: break
                    except:
                        continue
                
                if not send_clicked:
                    # Fallback: Type a space and enter to trigger if JS failed to trigger 'dirty' state
                    print(f"{Fore.YELLOW}   Could not find send button, trying keyboard submit...{Style.RESET_ALL}")
                    prompt_box.send_keys(" ")
                    time.sleep(0.5)
                    prompt_box.send_keys(Keys.RETURN)
                    
            except Exception as js_e:
                print(f"{Fore.RED}   JavaScript injection failed: {js_e}{Style.RESET_ALL}")
                # Last resort fallback
                print(f"{Fore.YELLOW}   Falling back to keyboard input...{Style.RESET_ALL}")
                prompt_box.send_keys(research_prompt)
                prompt_box.send_keys(Keys.RETURN)
                
            print(f"{Fore.YELLOW}[...] Waiting for Gemini's response (this may take a minute)...{Style.RESET_ALL}")
            
            # Wait for response to complete
            return self.wait_for_research_completion()
            
        except Exception as e:
            print(f"{Fore.RED}[!] Error during research: {e}{Style.RESET_ALL}")
            return None
    
    def wait_for_research_completion(self, progress_callback=None):
        """Wait for Gemini to complete its response with intelligent polling"""
        print(f"{Fore.YELLOW}[...] Waiting for Gemini to finish generating response...{Style.RESET_ALL}")
        
        # First, wait for response to start appearing (20 seconds)
        print(f"{Fore.CYAN}Waiting for response to start...{Style.RESET_ALL}")
        for i in range(4):  # 20 seconds / 5 second chunks
            time.sleep(5)
            if progress_callback:
                progress_callback()
        
        # Now poll to detect when Gemini stops typing
        max_checks = 12  # Check up to 12 times (60 seconds additional max)
        stable_count_needed = 2  # Need 2 consecutive stable checks
        stable_count = 0
        last_length = 0
        
        for check_num in range(max_checks):
            # Wait 5 seconds between checks
            time.sleep(5)
            if progress_callback:
                progress_callback()
            
            # Check if Gemini is still generating
            is_generating = self._is_gemini_generating()
            current_length = self._get_response_length()
            
            print(f"{Fore.CYAN}Check {check_num + 1}/{max_checks}: Response length = {current_length}, Still generating = {is_generating}{Style.RESET_ALL}")
            
            # If generation stopped AND text is stable
            if not is_generating and current_length == last_length and current_length > 500:
                stable_count += 1
                if stable_count >= stable_count_needed:
                    print(f"{Fore.GREEN}[+] Response complete! Extracting {current_length} characters...{Style.RESET_ALL}")
                    research_data = self.extract_research_data()
                    # STAGE 2: Send formatting prompt while browser is still open
                    research_data = self.format_research_into_slides(research_data)
                    return research_data
            else:
                stable_count = 0  # Reset if still changing
            
            last_length = current_length
        
        # Timeout - extract whatever we have
        print(f"{Fore.YELLOW}Timeout reached, extracting current response...{Style.RESET_ALL}")
        research_data = self.extract_research_data()
        # STAGE 2: Send formatting prompt while browser is still open
        research_data = self.format_research_into_slides(research_data)
        return research_data
    
    def format_research_into_slides(self, research_data):
        """
        STAGE 2: Send the raw research back to Gemini with slide_formating_prompt
        to get structured SLIDE output. This happens while browser is still open.
        
        Uses JavaScript injection (same as Stage 1) for reliable prompt submission,
        and polls for completion (same wait logic as Stage 1) instead of fixed timeout.
        """
        raw_text = research_data.get('raw_text', '')
        
        if not raw_text or len(raw_text) < 100:
            print(f"{Fore.YELLOW}[!] Not enough research text to format into slides{Style.RESET_ALL}")
            return research_data
        
        print(f"\n{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
        print(f"{Fore.CYAN}STAGE 2: Formatting Research into Slides{Style.RESET_ALL}")
        print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
        
        try:
            # Load the slide formatting prompt
            from modules.content_organizer import load_slide_formatting_prompt
            formatting_prompt = load_slide_formatting_prompt()
            
            # Build the formatting request — NO research text appended!
            # Gemini already has the research output in context (same chat session)
            # The prompt says "Convert the above academic content"
            formatting_request = f"""{formatting_prompt}

Generate approximately 15-20 slides from the content above. OUTPUT ONLY SLIDES IN THE SPECIFIED FORMAT."""
            
            print(f"{Fore.CYAN}[*] Sending formatting prompt to Gemini...{Style.RESET_ALL}")
            print(f"{Fore.CYAN}    Prompt length: {len(formatting_request)} chars{Style.RESET_ALL}")
            
            # Find the input box (same selectors as perform_deep_research)
            input_selectors = [
                "rich-textarea[placeholder*='Enter']",
                "div[contenteditable='true']",
                "textarea",
                ".ql-editor"
            ]
            
            prompt_box = None
            for selector in input_selectors:
                try:
                    prompt_box = self.driver.find_element(By.CSS_SELECTOR, selector)
                    if prompt_box:
                        break
                except:
                    continue
            
            if not prompt_box:
                print(f"{Fore.RED}[!] Could not find Gemini input box for Stage 2{Style.RESET_ALL}")
                research_data['formatted_slides'] = ''
                return research_data
            
            # Clear and inject via JavaScript
            prompt_box.click()
            time.sleep(2)
            
            try:
                prompt_box.clear()
            except:
                self.driver.execute_script("arguments[0].value = '';", prompt_box)
            
            time.sleep(1)
            
            # Method 1: JavaScript injection with paragraph elements for contenteditable
            # Gemini uses a rich-textarea (contenteditable div) that needs proper DOM manipulation
            print(f"{Fore.CYAN}    Injecting formatting prompt via JavaScript...{Style.RESET_ALL}")
            
            injection_success = False
            
            try:
                self.driver.execute_script("""
                    var element = arguments[0];
                    var text = arguments[1];
                    
                    // Clear existing content
                    element.innerHTML = '';
                    
                    // Create proper paragraph elements for contenteditable
                    var lines = text.split('\\n');
                    for (var i = 0; i < lines.length; i++) {
                        var p = document.createElement('p');
                        p.textContent = lines[i] || '\\u00A0';  // nbsp for empty lines
                        element.appendChild(p);
                    }
                    
                    // Also set text properties as fallback
                    element.value = text;
                    
                    // Fire all relevant events for React/Angular frameworks
                    element.dispatchEvent(new Event('input', { bubbles: true }));
                    element.dispatchEvent(new Event('change', { bubbles: true }));
                    element.dispatchEvent(new Event('keyup', { bubbles: true }));
                    element.dispatchEvent(new InputEvent('input', { bubbles: true, data: text, inputType: 'insertText' }));
                """, 
                    prompt_box, 
                    formatting_request
                )
                time.sleep(2)
                
                # Verify text was actually injected
                injected_text = self.driver.execute_script(
                    "return arguments[0].innerText || arguments[0].textContent || arguments[0].value || '';",
                    prompt_box
                )
                if len(injected_text.strip()) > 100:
                    injection_success = True
                    print(f"{Fore.GREEN}    [+] JS injection verified ({len(injected_text)} chars in input){Style.RESET_ALL}")
                else:
                    print(f"{Fore.YELLOW}    [!] JS injection may have failed ({len(injected_text)} chars detected){Style.RESET_ALL}")
            except Exception as js_e:
                print(f"{Fore.YELLOW}    [!] JS injection error: {js_e}{Style.RESET_ALL}")
            
            # Method 2: Clipboard paste (more reliable for contenteditable)
            if not injection_success:
                print(f"{Fore.CYAN}    Trying clipboard paste method...{Style.RESET_ALL}")
                try:
                    import pyperclip
                    pyperclip.copy(formatting_request)
                    prompt_box.click()
                    time.sleep(0.5)
                    from selenium.webdriver.common.keys import Keys
                    prompt_box.send_keys(Keys.CONTROL, 'a')
                    time.sleep(0.3)
                    prompt_box.send_keys(Keys.CONTROL, 'v')
                    time.sleep(2)
                    
                    injected_text = self.driver.execute_script(
                        "return arguments[0].innerText || arguments[0].textContent || '';",
                        prompt_box
                    )
                    if len(injected_text.strip()) > 100:
                        injection_success = True
                        print(f"{Fore.GREEN}    [+] Clipboard paste verified ({len(injected_text)} chars){Style.RESET_ALL}")
                    else:
                        print(f"{Fore.YELLOW}    [!] Clipboard paste may have failed{Style.RESET_ALL}")
                except Exception as clip_e:
                    print(f"{Fore.YELLOW}    [!] Clipboard method failed: {clip_e}{Style.RESET_ALL}")
            
            # Method 3: Direct send_keys (slowest but most compatible)
            if not injection_success:
                print(f"{Fore.CYAN}    Trying direct keyboard input (slow)...{Style.RESET_ALL}")
                try:
                    prompt_box.click()
                    time.sleep(0.5)
                    # Send in chunks to avoid timeout
                    chunk_size = 500
                    for i in range(0, len(formatting_request), chunk_size):
                        chunk = formatting_request[i:i+chunk_size]
                        prompt_box.send_keys(chunk)
                        time.sleep(0.1)
                    time.sleep(2)
                    injection_success = True
                    print(f"{Fore.GREEN}    [+] Keyboard input completed{Style.RESET_ALL}")
                except Exception as key_e:
                    print(f"{Fore.RED}    [!] All input methods failed: {key_e}{Style.RESET_ALL}")
            
            # Save FULL pre-submission text for text-subtraction extraction later
            pre_submit_text = self._extract_using_selectors()
            pre_submit_responses = self._count_bs_responses()
            pre_submit_length = len(pre_submit_text)
            print(f"{Fore.CYAN}    Response containers before submit: {pre_submit_responses}{Style.RESET_ALL}")
            print(f"{Fore.CYAN}    Total response text before submit: {pre_submit_length} chars{Style.RESET_ALL}")
            
            time.sleep(1)
            
            # Click send button
            send_clicked = False
            send_button_selectors = [
                "button[aria-label*='Send']",
                "button[class*='send-button']", 
                "button[data-test-id='send-button']",
                ".send-button",
                "button[type='submit']"
            ]
            
            for selector in send_button_selectors:
                try:
                    btns = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for btn in btns:
                        if btn.is_displayed():
                            self.driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                            time.sleep(0.5)
                            self.driver.execute_script("arguments[0].click();", btn)
                            print(f"{Fore.GREEN}    [+] Clicked send button{Style.RESET_ALL}")
                            send_clicked = True
                            break
                    if send_clicked:
                        break
                except:
                    continue
            
            if not send_clicked:
                print(f"{Fore.YELLOW}    Send button not found, trying Enter key...{Style.RESET_ALL}")
                prompt_box.send_keys(Keys.RETURN)
                time.sleep(1)
                # Try clicking send again after Enter
                for selector in send_button_selectors:
                    try:
                        btns = self.driver.find_elements(By.CSS_SELECTOR, selector)
                        for btn in btns:
                            if btn.is_displayed():
                                self.driver.execute_script("arguments[0].click();", btn)
                                send_clicked = True
                                break
                        if send_clicked:
                            break
                    except:
                        continue
            
            # PHASE 1: Wait for NEW response to appear
            # Detection uses MULTIPLE signals: container count, text length, AND generating indicator
            print(f"{Fore.YELLOW}[...] Waiting for Gemini to start formatting...{Style.RESET_ALL}")
            new_response_appeared = False
            for wait_check in range(30):  # Up to 150 seconds
                time.sleep(5)
                current_responses = self._count_bs_responses()
                current_length = self._get_response_length()
                new_text = current_length - pre_submit_length
                is_generating = self._is_gemini_generating()
                print(f"{Fore.CYAN}    Waiting... containers: {current_responses} (was {pre_submit_responses}), new text: {new_text} chars, generating: {is_generating}{Style.RESET_ALL}")
                
                # Any of these signals means response started:
                if current_responses > pre_submit_responses or new_text > 50 or is_generating:
                    new_response_appeared = True
                    print(f"{Fore.GREEN}    [+] New response detected!{Style.RESET_ALL}")
                    break
            
            if not new_response_appeared:
                print(f"{Fore.YELLOW}[!] No new response detected after waiting — will attempt extraction anyway{Style.RESET_ALL}")
                # DON'T return — try extraction anyway, Gemini may have responded
                # but our detection missed it
            
            # PHASE 2: Wait for response to finish (text stops growing)
            print(f"{Fore.YELLOW}[...] Waiting for formatting to complete...{Style.RESET_ALL}")
            max_checks = 40  # Up to 200 seconds
            stable_count = 0
            stable_needed = 3  # Need 3 stable checks for confidence
            last_new_len = 0
            
            for check_num in range(max_checks):
                time.sleep(5)
                
                is_generating = self._is_gemini_generating()
                current_length = self._get_response_length()
                new_text_len = current_length - pre_submit_length
                
                print(f"{Fore.CYAN}    Check {check_num + 1}/{max_checks}: New text={new_text_len}, Generating={is_generating}{Style.RESET_ALL}")
                
                # Text stable AND not generating → complete
                if new_text_len == last_new_len and new_text_len > 200:
                    if not is_generating:
                        stable_count += 1
                        if stable_count >= stable_needed:
                            print(f"{Fore.GREEN}[+] Formatting complete! ({new_text_len} chars of new content){Style.RESET_ALL}")
                            break
                    else:
                        # Text stable but still generating — could be "thinking"
                        stable_count = 0
                else:
                    stable_count = 0
                
                last_new_len = new_text_len
            else:
                # Loop ended without breaking — wait more if still generating
                if self._is_gemini_generating():
                    print(f"{Fore.YELLOW}    Still generating, waiting additional time...{Style.RESET_ALL}")
                    for extra in range(24):  # Additional 120 seconds
                        time.sleep(5)
                        if not self._is_gemini_generating():
                            print(f"{Fore.GREEN}    [+] Generation finished!{Style.RESET_ALL}")
                            time.sleep(3)  # Brief pause for DOM to settle
                            break
            
            # ============================================================
            # EXTRACT the formatted response using TEXT SUBTRACTION
            # ============================================================
            # Get ALL text from page, subtract pre-submission text to get only Stage 2
            # This captures the FULL response even when it spans multiple containers
            
            time.sleep(2)  # Let DOM settle
            
            all_text = self._extract_using_selectors()
            print(f"{Fore.CYAN}    Total page text: {len(all_text)} chars (pre-submit was {pre_submit_length}){Style.RESET_ALL}")
            
            formatted_response = ""
            
            # Method 1: Text subtraction (most reliable for full content)
            if len(all_text) > pre_submit_length + 100:
                # Try to find where pre-submit text ends and new text begins
                # The pre-submit text should be the prefix of all_text
                if all_text.startswith(pre_submit_text[:200]):
                    formatted_response = all_text[pre_submit_length:].strip()
                    print(f"{Fore.GREEN}    [+] Text subtraction: {len(formatted_response)} chars of new content{Style.RESET_ALL}")
                else:
                    # Pre-text may have shifted slightly — find the split
                    # Look for the first "--- SLIDE" marker in the new portion
                    slide_marker_pos = all_text.find('--- SLIDE 1 ---', pre_submit_length // 2)
                    if slide_marker_pos > 0:
                        formatted_response = all_text[slide_marker_pos:].strip()
                        print(f"{Fore.GREEN}    [+] Marker-based extraction: {len(formatted_response)} chars{Style.RESET_ALL}")
                    else:
                        # Just take everything after pre_submit_length
                        formatted_response = all_text[pre_submit_length:].strip()
                        print(f"{Fore.YELLOW}    [!] Approximate subtraction: {len(formatted_response)} chars{Style.RESET_ALL}")
            
            # Method 2: JS innerText on ALL containers after pre-submission ones (fallback)
            if len(formatted_response) < 500:
                print(f"{Fore.CYAN}    Trying JS container extraction...{Style.RESET_ALL}")
                try:
                    js_result = self.driver.execute_script("""
                        var selectors = [
                            'div[role="article"]',
                            'article',
                            'div.model-response-text',
                            'div.response-container'
                        ];
                        
                        var bestContainers = [];
                        for (var i = 0; i < selectors.length; i++) {
                            var elements = document.querySelectorAll(selectors[i]);
                            var real = [];
                            for (var j = 0; j < elements.length; j++) {
                                var txt = elements[j].innerText || '';
                                if (txt.trim().length > 50) {
                                    real.push(elements[j]);
                                }
                            }
                            if (real.length > bestContainers.length) {
                                bestContainers = real;
                            }
                        }
                        
                        // Get ALL containers after the known pre-submit count
                        var preCount = arguments[0];
                        var newText = '';
                        for (var i = preCount; i < bestContainers.length; i++) {
                            newText += bestContainers[i].innerText + '\\n\\n';
                        }
                        
                        // If no new containers, get the last one (may contain appended text)
                        if (!newText && bestContainers.length > 0) {
                            newText = bestContainers[bestContainers.length - 1].innerText;
                        }
                        
                        return newText;
                    """, pre_submit_responses) or ""
                    
                    if len(js_result.strip()) > len(formatted_response):
                        formatted_response = js_result.strip()
                        print(f"{Fore.GREEN}    [+] JS container extraction: {len(formatted_response)} chars{Style.RESET_ALL}")
                except Exception as js_e:
                    print(f"{Fore.YELLOW}    JS extraction failed: {js_e}{Style.RESET_ALL}")
            
            # Method 3: BeautifulSoup last response (last resort)
            if len(formatted_response) < 500:
                print(f"{Fore.CYAN}    Trying BS extraction fallback...{Style.RESET_ALL}")
                bs_response = self._extract_last_response_bs()
                if len(bs_response) > len(formatted_response):
                    formatted_response = bs_response
                    print(f"{Fore.CYAN}    BS extraction: {len(formatted_response)} chars{Style.RESET_ALL}")
            
            # Clean up: remove "Show thinking" / "Gemini said" prefixes
            for prefix in ['Show thinking\n', 'Gemini said\n', 'Show thinking ', 'Gemini said ']:
                if formatted_response.startswith(prefix):
                    formatted_response = formatted_response[len(prefix):].strip()
            
            print(f"{Fore.CYAN}    Final extracted: {len(formatted_response)} chars{Style.RESET_ALL}")
            if formatted_response:
                # Show first 300 chars to verify format
                preview = formatted_response[:300].replace('\n', '\\n')
                print(f"{Fore.CYAN}    Preview: {preview}{Style.RESET_ALL}")
            
            if formatted_response and len(formatted_response) > 100:
                research_data['formatted_slides'] = formatted_response
                print(f"{Fore.GREEN}[+] Got {len(formatted_response)} chars of formatted slides{Style.RESET_ALL}")
                
                # Save Stage 2 extracted output to log file for debugging
                try:
                    output_dir = research_data.get('output_dir', '')
                    if not output_dir:
                        # Fallback: save next to the script
                        output_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'output')
                    os.makedirs(output_dir, exist_ok=True)
                    log_path = os.path.join(output_dir, 'second_prompt_output.txt')
                    with open(log_path, 'w', encoding='utf-8') as f:
                        f.write(formatted_response)
                    print(f"{Fore.GREEN}[+] Stage 2 output saved to: {log_path}{Style.RESET_ALL}")
                except Exception as log_e:
                    print(f"{Fore.YELLOW}[!] Could not save Stage 2 log: {log_e}{Style.RESET_ALL}")
            else:
                print(f"{Fore.YELLOW}[!] Formatting response too short ({len(formatted_response)} chars){Style.RESET_ALL}")
                research_data['formatted_slides'] = ''
            
        except Exception as e:
            print(f"{Fore.RED}[!] Slide formatting failed: {e}{Style.RESET_ALL}")
            import traceback
            traceback.print_exc()
            research_data['formatted_slides'] = ''
        
        return research_data
    
    def _is_gemini_generating(self):
        """Check if Gemini is still generating a response.
        Uses JavaScript to check for VISIBLE stop/cancel buttons and spinners,
        NOT string matching (which falsely matches CSS class names)."""
        try:
            return self.driver.execute_script("""
                // Check for visible "Stop generating" or "Stop" buttons
                var buttons = document.querySelectorAll('button');
                for (var i = 0; i < buttons.length; i++) {
                    var btn = buttons[i];
                    if (!btn.offsetParent && btn.offsetWidth === 0) continue; // skip hidden
                    var label = (btn.getAttribute('aria-label') || '').toLowerCase();
                    var text = (btn.innerText || '').toLowerCase();
                    if (label.indexOf('stop') >= 0 || text.indexOf('stop') >= 0) {
                        return true;
                    }
                }
                
                // Check for loading spinners/progress indicators
                var spinners = document.querySelectorAll('mat-spinner, .loading-spinner, [role="progressbar"]');
                for (var i = 0; i < spinners.length; i++) {
                    if (spinners[i].offsetParent || spinners[i].offsetWidth > 0) {
                        return true;
                    }
                }
                
                // Check for typing indicators
                var indicators = document.querySelectorAll('.typing-indicator, [data-is-generating="true"]');
                if (indicators.length > 0) return true;
                
                return false;
            """)
        except:
            return False
    
    def _get_response_length(self):
        """Get the current length of Gemini's response"""
        try:
            research_text = self._extract_using_selectors()
            return len(research_text)
        except:
            return 0
    
    def _count_bs_responses(self):
        """Count response containers using BeautifulSoup.
        Used to detect when a NEW response appears after submitting Stage 2."""
        try:
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')
            
            # Try multiple selectors for response containers
            selectors = [
                {'name': 'div', 'attrs': {'role': 'article'}},
                {'name': 'article', 'attrs': {}},
                {'name': 'div', 'attrs': {'class': re.compile(r'model-response|response-container', re.I)}},
                {'name': 'div', 'attrs': {'data-test-id': re.compile(r'conversation|response', re.I)}},
            ]
            
            max_count = 0
            for selector in selectors:
                containers = soup.find_all(selector['name'], selector['attrs'])
                # Only count containers with actual content
                real = [c for c in containers if len(c.get_text(strip=True)) > 50]
                if len(real) > max_count:
                    max_count = len(real)
            
            return max_count
        except:
            return 0
    
    def _extract_last_response_bs(self):
        """Extract the LAST response element using BeautifulSoup.
        Uses get_text(separator='\\n') to PRESERVE NEWLINES between bullets.
        This is critical for the slide parser to work correctly."""
        try:
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')
            
            # Try multiple selectors
            selectors = [
                {'name': 'div', 'attrs': {'role': 'article'}},
                {'name': 'article', 'attrs': {}},
                {'name': 'div', 'attrs': {'class': re.compile(r'model-response|response-container', re.I)}},
                {'name': 'div', 'attrs': {'data-test-id': re.compile(r'conversation|response', re.I)}},
            ]
            
            best_text = ""
            for selector in selectors:
                containers = soup.find_all(selector['name'], selector['attrs'])
                # Get containers with actual content
                real = [c for c in containers if len(c.get_text(strip=True)) > 50]
                
                if real:
                    # Get the LAST one (= Stage 2 response)
                    last = real[-1]
                    text = last.get_text(separator='\n', strip=True)
                    
                    if len(text) > len(best_text):
                        best_text = text
            
            return best_text
        except Exception as e:
            print(f"{Fore.RED}[!] BS extraction failed: {e}{Style.RESET_ALL}")
            return ""
    
    def _count_response_elements(self):
        """Count how many response/message elements exist on the page.
        Used to detect when a NEW response appears after submitting a second prompt."""
        try:
            script = """
            var selectors = [
                'div[role="article"]', 'article',
                '[class*="model-response"]', '[class*="response-container"]',
                '[data-test-id*="response"]'
            ];
            var count = 0;
            for (var i = 0; i < selectors.length; i++) {
                var els = document.querySelectorAll(selectors[i]);
                if (els.length > count) count = els.length;
            }
            return count;
            """
            return self.driver.execute_script(script) or 0
        except:
            return 0
    
    def extract_research_data(self):
        """Extract structured research data from Gemini's response"""
        try:
            # Wait a moment for any final rendering
            time.sleep(2)
            
            # Method 1: Try to find response using updated selectors
            research_text = self._extract_using_selectors()
            
            # Method 2: If that fails, get all meaningful text from page
            if len(research_text) < 200:
                research_text = self._extract_all_text()
            
            # Method 3: Last resort - try JavaScript to get visible text
            if len(research_text) < 200:
                research_text = self._extract_with_javascript()
            
            if len(research_text) < 100:
                print(f"{Fore.RED}[!] Could not extract sufficient content{Style.RESET_ALL}")
                return {"raw_text": research_text, "sections": []}
            
            # Parse and structure the data
            structured_data = self.parse_research_text(research_text)
            
            print(f"{Fore.GREEN}[+] Research completed ({len(research_text)} characters extracted){Style.RESET_ALL}")
            
            return structured_data
            
        except Exception as e:
            print(f"{Fore.RED}[!] Error extracting research data: {e}{Style.RESET_ALL}")
            return {"raw_text": "", "sections": []}
    
    def _extract_using_selectors(self):
        """Try to extract using CSS selectors for response containers"""
        try:
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')
            
            # Updated selectors for Gemini's current UI
            # Try multiple strategies
            selectors = [
                {'name': 'div', 'attrs': {'class': re.compile(r'model-response|response-container', re.I)}},
                {'name': 'div', 'attrs': {'data-test-id': re.compile(r'conversation|response', re.I)}},
                {'name': 'article', 'attrs': {}},
                {'name': 'div', 'attrs': {'role': 'article'}},
            ]
            
            research_text = ""
            for selector in selectors:
                containers = soup.find_all(selector['name'], selector['attrs'])
                for container in containers:
                    text = container.get_text(separator='\n', strip=True)
                    if len(text) > 100:
                        research_text += text + "\n\n"
                
                if len(research_text) > 500:
                    break
            
            return research_text.strip()
        except Exception as e:
            print(f"{Fore.YELLOW}[!] Selector extraction failed: {e}{Style.RESET_ALL}")
            return ""
    
    def _extract_all_text(self):
        """Extract all meaningful text from the page"""
        try:
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')
            
            # Remove script and style elements
            for script in soup(['script', 'style', 'nav', 'header', 'footer']):
                script.decompose()
            
            # Get all text
            text = soup.get_text(separator='\n', strip=True)
            
            # Clean up: remove lines that are likely UI elements
            lines = text.split('\n')
            cleaned_lines = []
            for line in lines:
                line = line.strip()
                # Skip short lines or UI-related text
                if len(line) > 20 and not any(skip in line.lower() for skip in ['sign in', 'menu', 'settings', 'new chat']):
                    cleaned_lines.append(line)
            
            return '\n'.join(cleaned_lines)
        except Exception as e:
            print(f"{Fore.YELLOW}[!] Text extraction failed: {e}{Style.RESET_ALL}")
            return ""
    
    def _extract_with_javascript(self):
        """Use JavaScript to extract visible text"""
        try:
            # Get all visible text using JavaScript
            script = """
            var elements = document.querySelectorAll('div[role="article"], article, .response, [class*="model"], [class*="message"]');
            var text = '';
            for (var i = 0; i < elements.length; i++) {
                text += elements[i].innerText + '\\n\\n';
            }
            return text;
            """
            text = self.driver.execute_script(script)
            return text if text else ""
        except Exception as e:
            print(f"{Fore.YELLOW}[!] JavaScript extraction failed: {e}{Style.RESET_ALL}")
            return ""

    
    
    def parse_research_text(self, text):
        """Parse research text into structured sections with improved filtering"""
        sections = []
        current_section = {"title": "Introduction", "content": []}
        
        lines = text.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Skip very short lines that are likely abbreviations or UI artifacts
            if len(line) < 4 and not line[0].isdigit():
                continue
            
            # Skip common UI text (exact matches only to avoid false positives)
            skip_patterns = ['sign in', 'menu', 'settings', 'share', 'copy link', 'edit', 'new chat']
            if line.lower() in skip_patterns:
                continue
            
            # Detect section headers
            is_header = False
            
            # Pattern 1: Numbered headers (1. Title, 1) Title, etc.)
            if line.startswith(('1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.',
                               '1)', '2)', '3)', '4)', '5)', '6)', '7)', '8)', '9)')):
                is_header = True
            
            # Pattern 2: Headers with ## or ###
            elif line.startswith('#'):
                is_header = True
            
            # Pattern 3: Short all-caps lines (likely headers)
            elif len(line) < 80 and line.isupper() and len(line) > 5:
                is_header = True
            
            # Pattern 4: Lines ending with colon (might be headers)
            elif line.endswith(':') and len(line) < 100 and len(line) > 10:
                is_header = True
            
            if is_header:
                # Save previous section if it has content
                if current_section["content"]:
                    sections.append(current_section)
                
                # Start new section
                title = line.lstrip('#').lstrip('0123456789.)').strip().rstrip(':')
                current_section = {"title": title, "content": []}
            else:
                # IMPROVED: More lenient content filtering (reduced from 15 to 10 chars)
                # Accept substantial content or complete sentences
                if len(line) > 10 or (len(line) > 5 and '.' in line):
                    # Clean up the line
                    line = line.strip('*---')
                    if line:
                        current_section["content"].append(line)
        
        # Add the last section
        if current_section["content"]:
            sections.append(current_section)
        
        # If no sections were created, create one big section from all text
        if not sections and text:
            # Split into sentences
            sentences = []
            for line in text.split('\n'):
                line = line.strip()
                if len(line) > 20:  # Only substantial lines
                    sentences.append(line)
            
            if sentences:
                # Group sentences into chunks
                chunk_size = 8
                for i in range(0, len(sentences), chunk_size):
                    chunk = sentences[i:i+chunk_size]
                    if chunk:
                        sections.append({
                            "title": f"Topic {len(sections) + 1}",
                            "content": chunk
                        })
        
        return {
            "raw_text": text,
            "sections": sections,
            "total_sections": len(sections)
        }
    
    def ask_gemini(self, prompt):
        """General function to ask Gemini a specific question"""
        print(f"{Fore.CYAN}[*] Asking Gemini: {prompt[:50]}...{Style.RESET_ALL}")
        
        try:
            # Find and clear the input box
            input_selectors = [
                "rich-textarea[placeholder*='Enter']",
                "div[contenteditable='true']",
                "textarea",
                ".ql-editor"
            ]
            
            prompt_box = None
            for selector in input_selectors:
                try:
                    prompt_box = self.driver.find_element(By.CSS_SELECTOR, selector)
                    if prompt_box:
                        break
                except:
                    continue
            
            if not prompt_box:
                return None
            
            # Clear previous content if any
            prompt_box.clear()
            time.sleep(0.5)
            
            # Type and submit new prompt
            prompt_box.click()
            prompt_box.send_keys(prompt)
            time.sleep(0.5)
            prompt_box.send_keys(Keys.RETURN)
            
            # Wait for response (timeout for shorter queries)
            print(f"{Fore.YELLOW}[...] Waiting for response...{Style.RESET_ALL}")
            time.sleep(20)  # Increased from 10 to 20 seconds
            
            # Extract response using improved method
            return self.extract_last_response()
            
        except Exception as e:
            print(f"{Fore.RED}[!] Error asking Gemini: {e}{Style.RESET_ALL}")
            return None
    
    def extract_last_response(self):
        """Extract the most recent Gemini response"""
        try:
            # Use JavaScript to get the last response
            script = """
            var elements = document.querySelectorAll('div[role="article"], article, [class*="model"], [class*="message"]');
            if (elements.length > 0) {
                return elements[elements.length - 1].innerText;
            }
            return '';
            """
            text = self.driver.execute_script(script)
            
            if text and len(text) > 50:
                return text
            
            # Fallback to BeautifulSoup
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')
            
            # Find all response elements
            responses = soup.find_all(['div', 'article'], 
                                     attrs={'role': 'article'})
            
            if not responses:
               responses = soup.find_all(['div', 'article'], 
                                        class_=re.compile(r'response|message|model', re.I))
            
            if responses:
                # Get the last response
                last_response = responses[-1].get_text(separator='\n', strip=True)
                return last_response
            
            return ""
            
        except Exception as e:
            print(f"{Fore.RED}[!] Error extracting response: {e}{Style.RESET_ALL}")
            return ""
    
    def close_browser(self):
        """Clean up and close browser"""
        if self.driver:
            print(f"{Fore.CYAN}[*] Closing browser...{Style.RESET_ALL}")
            self.driver.quit()
            print(f"{Fore.GREEN}[+] Browser closed{Style.RESET_ALL}")


def gather_research(topic, research_level=3, super_research=False, progress_callback=None, output_folder=None):
    """Main function to gather research on a topic using Gemini
    
    Args:
        topic: The presentation topic
        research_level: 1 (Basic), 2 (Professional), or 3 (Master Thesis)
        super_research: Boolean - Enable academic API integration
        progress_callback: Optional callback for progress updates
        output_folder: Output folder path for super_research data
    """
    researcher = GeminiResearcher()
    
    try:
        # Initialize browser
        researcher.init_browser()
        
        # Navigate to Gemini
        if not researcher.navigate_to_gemini():
            return None
        
        # Perform deep research with progress callback (super_research parameter already received from function args)
        research_data = researcher.perform_deep_research(topic, research_level, super_research, progress_callback, output_folder)
        
        return research_data, researcher  # Return researcher to keep browser open for more queries
        
    except Exception as e:
        import traceback
        print(f"{Fore.RED}[!] ERROR: Research failed{Style.RESET_ALL}")
        print(f"{Fore.YELLOW}Details: {str(e)}{Style.RESET_ALL}")
        print(f"{Fore.YELLOW}Traceback:{Style.RESET_ALL}")
        traceback.print_exc()
        researcher.close_browser()
        raise  # Re-raise the exception so we can see it in the GUI error dialog

