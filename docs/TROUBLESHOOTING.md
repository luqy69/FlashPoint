# Troubleshooting Guide

## Common Issues and Solutions

### 1. "Timeout waiting for response" Error

**Symptoms:**
- Gemini provides a detailed response
- Script shows "⏳ Waiting for Gemini's response..."
- After 2 minutes: "✗ Timeout waiting for response"
- You can see the response in browser, but script can't extract it

**Root Cause:**
The script couldn't detect when Gemini finished responding or couldn't extract the response content.

**Solution (FIXED in latest version):**
The research module has been updated with:
- Fixed 60-second wait time (instead of trying to detect completion)
- Multiple extraction methods (CSS selectors, JavaScript, full page text)
- 3 retry attempts if initial extraction fails
- Better handling of Gemini's dynamic UI

**What Changed:**
```python
# OLD: Looked for 'generating' text (unreliable)
if 'generating' not in page_source.lower():
    return self.extract_research_data()

# NEW: Fixed wait time + multiple extraction attempts
time.sleep(60)  # Wait for Gemini to finish
for attempt in range(3):
    data = self.extract_research_data()  # Try extracting
    if sufficient data: return data
    time.sleep(15)  # Wait more if needed
```

**To Fix:**
Make sure you're using the updated `modules/research.py` file. The fixes include:
1. 60-second initial wait
2. 3 extraction attempts with 15-second intervals
3. Multiple extraction methods (selectors, JavaScript, full-text)

---

### 2. "Could not find Gemini input box" Error

**Symptoms:**
- Browser opens correctly
- You log in successfully
- Error: "Could not find Gemini input box"

**Cause:**
Gemini's UI has changed and the CSS selectors don't match.

**Solution:**
1. Open browser developer tools (F12)
2. Inspect the text input area where you type prompts
3. Find the element's class, id, or attributes
4. Update `modules/research.py` line 86-91 with new selectors:

```python
input_selectors = [
    "rich-textarea[placeholder*='Enter']",
    "div[contenteditable='true']",
    "textarea",
    ".ql-editor",
    # Add new selectors here based on what you found
    "YOUR_NEW_SELECTOR_HERE"
]
```

---

### 3. Response Extraction Returns Empty or Garbage Text

**Symptoms:**
- Script completes without errors
- But extracted text is empty, very short, or full of UI elements

**Solution:**
The updated version uses 3 extraction methods in order:

**Method 1: CSS Selectors**
```python
# Looks for: div[role="article"], article tags, response containers
```

**Method 2: Full Text Extraction**
```python
# Gets all page text, filters out UI elements
# Removes: "sign in", "menu", "settings", etc.
```

**Method 3: JavaScript**
```python
# Uses JavaScript to get innerText of response elements
```

If still having issues:
1. Keep browser window visible (don't minimize)
2. Wait for the full 60 seconds
3. Check if Gemini is actually showing a response
4. Try a different/simpler topic

---

### 4. Slow or No Response from Gemini

**Symptoms:**
- Script waits 60+ seconds
- Gemini page is blank or stuck loading
- No response appears in browser

**Possible Causes:**
- Internet connection issues
- Gemini service temporarily down
- Rate limiting
- Topic was flagged/blocked

**Solutions:**
1. **Check Internet**: Verify connection is stable
2. **Try Manual Query**: Type the same prompt manually in Gemini
3. **Wait Longer**: Sometimes Gemini takes 2-3 minutes for complex topics
4. **Simplify Topic**: Try a simpler, more general topic
5. **Check Gemini Status**: Visit gemini.google.com manually
6. **New Session**: Close browser, restart script

---

### 5. Browser Closes Immediately

**Symptoms:**
- Chrome opens
- Immediately closes
- Error in terminal

**Solutions:**
- Check ChromeDriver compatibility
- Update Selenium: `pip install --upgrade selenium`
- Update webdriver-manager: `pip install --upgrade webdriver-manager`
- Run as administrator
- Check antivirus isn't blocking

---

### 6. Login Loop - Keeps Asking to Log In

**Symptoms:**
- You log in successfully
- Press Enter
- Script asks to log in again

**Solution:**
The script checks for "sign in" text on the page. If your UI has this text even when logged in:

1. Edit `modules/research.py` line 70-77
2. Improve the login detection:

```python
def check_login_required(self):
    """Check if user needs to log in"""
    try:
        page_source = self.driver.page_source.lower()
        # More specific check
        return ('sign in' in page_source and 
                'chat' not in page_source)
    except:
        return False
```

---

### 7. Content Generation Fails / No Slides Created

**Symptoms:**
- Research completes successfully
- Error during content generation
- PowerPoint has minimal slides

**Cause:**
- Gemini's response wasn't properly structured
- follow-up queries failed

**Solution:**
The system has fallbacks:
1. If Gemini outline fails → uses basic outline from research
2. If visualizations fail → creates sample charts
3. Ensures minimum 20 slides

To improve:
- Make sure research data is substantial (check terminal output)
- Keep browser open during entire process
- Try topics with more searchable content

---

### 8. Charts/Visualizations Not Generated

**Symptoms:**
- Presentation created
- No charts folder or empty
- Slides don't have visualizations

**Cause:**
- Gemini couldn't suggest data for charts
- Matplotlib error
- Fallback to samples failed

**What Happens:**
- System asks Gemini for chart suggestions
- If fails → creates 3 sample charts automatically
- Sample charts always work as backup

**Check:**
- Look in `output/[topic]/charts/` folder
- Should have at least 3 PNG files
- If missing → check matplotlib installation: `pip install matplotlib`

---

### 9. "ModuleNotFoundError" Errors

**Fix:**
```bash
# Reinstall dependencies
cd "path\to\FlashPoint"
pip install --upgrade -r requirements.txt
```

---

### 10. PowerPoint Won't Open

**Symptoms:**
- .pptx file created
- Error when opening in PowerPoint
- File corrupted message

**Solutions:**
1. **Try Different Viewer**: Google Slides, LibreOffice
2. **Check File Size**: Should be > 100KB
3. **Reinstall python-pptx**: `pip install --upgrade python-pptx`
4. **Check Output Path**: Ensure no special characters

---

## Debug Mode

To see more details while running:

1. **Keep Browser Visible**: Don't minimize, watch what's happening
2. **Check Terminal**: Read all colored messages
3. **Inspect Manually**: After script submits prompt, check browser manually
4. **Take Screenshot**: If stuck, screenshot browser and terminal

---

## Manual Verification Steps

If automatic extraction fails, you can manually verify:

1. **Let Script Run**: Wait full 60 seconds
2. **Check Browser**: Is there a Gemini response visible?
3. **Copy Response**: If visible, select all text and copy
4. **Save to File**: Paste into a text file
5. **Re-run Script**: (Future feature: paste in manual content)

---

## Getting Help

If issues persist:

1. **Check Version**: Make sure you have latest `modules/research.py`
2. **Note Your Setup**:
   - Python version: `python --version`
   - OS: Windows version
   - Chrome version
   - Error message (full text)
3. **Try Different Topic**: See if issue is topic-specific
4. **Clear Browser Data**: Cache/cookies may interfere

---

## Quick Fixes Checklist

- [ ] Updated all files to latest version
- [ ] Reinstalled dependencies: `pip install -r requirements.txt`
- [ ] Chrome is up to date
- [ ] Logged into Google/Gemini
- [ ] Internet connection is stable
- [ ] Antivirus not blocking Python/Chrome
- [ ] Running from correct directory
- [ ] Tried a simple topic (e.g., "Renewable Energy")

---

## Advanced: Increasing Wait Times

If your internet is slow or Gemini takes longer:

Edit `modules/research.py`:

```python
# Line ~135: Initial wait time
time.sleep(90)  # Increase from 60 to 90 seconds

# Line ~145: Additional wait per attempt  
time.sleep(20)  # Increase from 15 to 20 seconds
```

---

## Latest Improvements (v1.1)

✅ Fixed: Timeout detection - now uses fixed wait time  
✅ Fixed: Response extraction - 3 different methods
✅ Fixed: Retry logic - 3 attempts with delays  
✅ Fixed: JavaScript extraction - more reliable  
✅ Improved: Better logging and error messages  
✅ Added: Fallback content if extraction fails  

The system is now much more robust and should successfully extract Gemini responses even when the UI changes.
