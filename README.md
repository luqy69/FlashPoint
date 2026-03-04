# 🎯 AI PowerPoint Presentation Generator

An intelligent Python application that automatically generates professional PowerPoint presentations on any topic using Google Gemini's deep research capabilities and browser automation.

## ✨ Features

- 🔍 **Automated Research**: Uses Selenium to automate Chrome and leverage Google Gemini's deep research
- 🎨 **Professional Design**: Clean, modern slide layouts with consistent formatting
- 📊 **Data Visualizations**: Automatically generates charts, graphs, and diagrams using matplotlib
- 📝 **Speaker Notes**: Includes detailed speaker notes for each slide
- 🎯 **Smart Content**: AI-powered content organization and structuring
- 🔄 **Customizable**: Adjust slide count and presentation focus
- 💾 **Citations**: Proper attribution of research sources

## 📋 Requirements

- Python 3.8 or higher
- Google Chrome browser installed
- Internet connection
- Access to gemini.google.com (you may need to log in)

## 🚀 Installation

### Step 1: Install Python Dependencies

```bash
cd "path\to\FlashPoint"
pip install --upgrade -r requirements.txt
```

The `requirements.txt` includes:
- `python-pptx` - PowerPoint creation
- `selenium` - Browser automation
- `webdriver-manager` - Automatic ChromeDriver setup
- `beautifulsoup4` - HTML parsing
- `matplotlib` - Data visualization
- `pillow` - Image processing
- `tqdm` - Progress bars
- `colorama` - Colored terminal output

### Step 2: Verify Chrome Installation

Make sure Google Chrome is installed on your system. The script will automatically download and configure the appropriate ChromeDriver.

## 📖 Usage

### Option 1: Run the Executable (Easiest — No Python needed)

1. Download the `package/` folder or the **FlashPoint_Deploy.zip** from this repository
2. Extract the zip (if downloaded)
3. Double-click **`FlashPoint_AI.exe`** to launch the app
4. Enter your topic, select theme, and let the AI generate your presentation!

> The `package/` folder contains everything you need: the executable, prompt configs, and an output folder.

### Option 2: Run from Source (For developers)

1. Install dependencies:
   ```bash
   pip install --upgrade -r requirements.txt
   ```
2. Launch the GUI:
   ```bash
   python ppt_wizard.py
   ```
   Or the command-line version:
   ```bash
   python ppt_generator.py
   ```

### What happens next

1. **Enter your topic** — e.g. "Artificial Intelligence in Healthcare"
2. **Choose slide count & research depth** — Basic, Professional, or Master Thesis
3. **Select a theme** — 10+ built-in professional themes
4. **Chrome opens automatically** — logs into Gemini and researches your topic
5. **Presentation is generated** — saved in the `output/` folder as `.pptx`

## 📁 Project Structure

```
ppt/
├── ppt_generator.py          # Command-line application
├── ppt_wizard.py             # GUI application
├── run.bat                   # Quick start script
├── requirements.txt          # Python dependencies
├── README.md                 # This file
├── config/                   # Prompt templates (Level 1-3)
├── docs/                     # Documentation and guides
├── assets/                   # Icons and logos
├── scripts/                  # Build and utility scripts
├── tests/                    # Testing scripts
├── modules/
│   ├── __init__.py
│   ├── research.py           # Gemini browser automation
│   ├── content_organizer.py  # Content structuring
│   ├── visualizations.py     # Chart generation
│   └── ppt_builder.py        # PowerPoint assembly
└── output/                   # Generated presentations
    └── [Topic]_[Timestamp]/
        ├── [Topic]_presentation.pptx
        └── charts/
            ├── chart_1.png
            ├── chart_2.png
            └── ...
```

## 🎨 Presentation Structure

Each generated presentation includes:

1. **Title Slide**: Topic name and subtitle
2. **Table of Contents**: Overview of sections
3. **Introduction** (2-3 slides): Context and background
4. **Main Content** (15-20 slides): 
   - Organized into logical sections
   - Bullet points with key information
   - Data visualizations integrated throughout
5. **Conclusion** (2-3 slides): Summary and key takeaways
6. **References**: Research sources and citations
7. **Thank You Slide**: Closing slide


## ⚙️ How It Works

### 1. Research Phase
- Opens Chrome browser with Selenium
- Navigates to gemini.google.com
- Submits comprehensive research prompt
- Waits for Gemini's deep research to complete
- Extracts and parses structured content

### 2. Content Generation
- Analyzes research data
- Asks Gemini to create presentation outline
- Organizes content into 20-30 slides
- Generates speaker notes for key slides
- Identifies opportunities for visualizations

### 3. Visualization Creation
- Asks Gemini for relevant data and statistics
- Generates chart suggestions
- Creates professional visualizations with matplotlib
- Saves as high-resolution PNG images

### 4. PowerPoint Assembly
- Creates presentation using python-pptx
- Applies professional theme and formatting
- Adds title, content, and image slides
- Inserts speaker notes
- Saves final PPTX file

## 🔧 Customization

### Changing Colors

Edit `modules/visualizations.py` and `modules/ppt_builder.py`:

```python
COLORS = {
    'primary': '#2E5090',    # Change to your color
    'secondary': '#4A90E2',
    'accent': '#E94B3C',
    # ...
}
```

### Adjusting Slide Count

When running the script, enter your desired slide count, or modify the default in `ppt_generator.py`:

```python
slide_count = int(slide_count_input) if slide_count_input else 30  # Change 22 to 30
```

### Modifying Prompts

Edit the research prompt in `modules/research.py`:

```python
research_prompt = f"Your custom prompt for '{topic}'..."
```

## 🐛 Troubleshooting

### Issue: "Could not find Gemini input box"

**Solution**: Gemini's UI may have changed. Update the CSS selectors in `modules/research.py`:

```python
input_selectors = [
    "rich-textarea[placeholder*='Enter']",  # Try different selectors
    "div[contenteditable='true']",
    # Add new selectors based on Gemini's current UI
]
```

### Issue: Browser doesn't open

**Solution**: 
- Verify Chrome is installed
- Run: `pip install --upgrade selenium webdriver-manager`
- Manually download ChromeDriver if needed

### Issue: "Login required" but already logged in

**Solution**: 
- Clear browser cache
- Try logging out and back in
- Press Enter in terminal to continue after manual login

### Issue: No visualizations generated

**Solution**: 
- The script will create sample visualizations as fallback
- Check the `charts/` directory in output folder
- Review terminal output for specific errors

### Issue: Incomplete research data

**Solution**:
- Increase wait time in `modules/research.py`:
  ```python
  max_wait = 180  # Increase from 120 to 180 seconds
  ```
- Check internet connection
- Verify Gemini access

## 🎯 Tips for Best Results

1. **Be Specific**: Use clear, specific topics (e.g., "Machine Learning in Healthcare" vs. "AI")
2. **Allow Time**: Deep research takes 2-3 minutes - don't interrupt
3. **Stay Logged In**: Keep your Google account logged in to Gemini
4. **Review Output**: Always review and customize the generated presentation
5. **Edit in PowerPoint**: Fine-tune formatting and content in Microsoft PowerPoint or Google Slides

## 🔐 Privacy & Security

- The script only accesses gemini.google.com
- No data is stored or transmitted except to Gemini
- Browser automation runs locally on your machine
- Generated presentations are saved only to your local output folder

## 📝 License

This project is provided as-is for educational and personal use.

## 🤝 Contributing

Feel free to:
- Report bugs
- Suggest new features
- Improve documentation
- Submit pull requests

## 📧 Support

For issues or questions:
1. Check the Troubleshooting section above
2. Review terminal output for error messages
3. Verify all dependencies are installed correctly

## 🎓 Examples

### Example 1: Technology Topic
```
Topic: "Quantum Computing"
Output: 25 slides covering quantum mechanics basics, qubits, algorithms, 
        applications, and future prospects with 5 visualizations
```

### Example 2: Business Topic
```
Topic: "Digital Marketing Strategies"
Output: 22 slides on SEO, social media, content marketing, analytics,
        with charts showing trends and statistics
```

### Example 3: Science Topic
```
Topic: "Renewable Energy Sources"
Output: 28 slides covering solar, wind, hydro, geothermal energy
        with comparative charts and growth trends
```

## 🚀 Advanced Usage

### Running in Headless Mode

Edit `modules/research.py`:

```python
researcher.init_browser(headless=True)  # Browser runs in background
```

Note: Headless mode may have issues with Gemini's interface. Use with caution.

### Batch Processing

Create a script to generate multiple presentations:

```python
topics = [
    "Artificial Intelligence",
    "Climate Change", 
    "Blockchain Technology"
]

for topic in topics:
    # Run ppt_generator with topic
    # ...
```

## ⚡ Performance

- **Research**: 2-3 minutes (depends on Gemini response time)
- **Content Generation**: 1-2 minutes (multiple Gemini queries)
- **Visualization**: 30 seconds (local processing)
- **PowerPoint**: 30 seconds (local processing)
- **Total**: ~4-6 minutes per presentation

## 🌟 Features Coming Soon

- Multiple theme options
- Export to PDF
- More chart types
- Custom templates
- Image search integration
- Multi-language support

---

**Made with ❤️ using Python, Selenium, and Google Gemini**

**Happy Presenting! 🎉**
