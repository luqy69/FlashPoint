# Quick Start Examples

## Example 1: Basic Usage

```bash
python ppt_generator.py
```

**Input:**
```
📌 Enter presentation topic: Artificial Intelligence
📊 Target slide count (default: 20-25): [Press Enter]
```

**Output:**
- Presentation: `output/Artificial_Intelligence_[timestamp]/Artificial_Intelligence_presentation.pptx`
- 25 slides covering AI fundamentals, applications, ethics, and future trends
- 5 charts showing AI adoption, market growth, and comparisons

---

## Example 2: Custom Slide Count

```bash
python ppt_generator.py
```

**Input:**
```
📌 Enter presentation topic: Renewable Energy
📊 Target slide count (default: 20-25): 30
```

**Output:**
- More detailed presentation with 30 slides
- In-depth coverage of solar, wind, hydro, geothermal energy
- Additional sections on policy, economics, and case studies

---

## Example 3: Business Topic

```bash
python ppt_generator.py
```

**Input:**
```
📌 Enter presentation topic: Digital Marketing Strategies 2024
```

**Expected Slides:**
1. Title: "Digital Marketing Strategies 2024"
2. Table of Contents
3-5. Introduction to Digital Marketing
6-10. SEO and Content Marketing
11-15. Social Media Marketing
16-20. Analytics and Metrics
21. Data Visualization (conversion rates, ROI)
22. Data Visualization (channel comparison)
23-24. Future Trends
25. Conclusion
26. References

---

## Example 4: Science Topic

```bash
python ppt_generator.py
```

**Input:**
```
📌 Enter presentation topic: Quantum Computing
```

**Expected Content:**
- Quantum mechanics basics
- Qubits and superposition
- Quantum algorithms
- Current applications
- Industry adoption charts
- Timeline visualization
- Future potential

---

## Example 5: Educational Topic

```bash
python ppt_generator.py
```

**Input:**
```
📌 Enter presentation topic: World War II History
```

**Expected Slides:**
- Historical timeline
- Key events and battles
- Major figures
- Impact and consequences
- Statistical charts (casualties, economics)
- Maps and diagrams
- Legacy and lessons

---

## Tips for Best Results

### 🎯 Topic Selection

**Good Topics:**
- Specific and focused: "Machine Learning in Healthcare"
- Current and relevant: "Electric Vehicles 2024"
- Well-researched: "Climate Change Solutions"

**Avoid:**
- Too broad: "Technology" (be more specific)
- Too narrow: "A specific bug in a specific library"
- Controversial without context: Provide clear angle

### ⏱️ Expected Timing

| Phase | Duration |
|-------|----------|
| Research | 2-3 min |
| Content Generation | 1-2 min |
| Visualizations | 30 sec |
| PowerPoint | 30 sec |
| **Total** | **4-6 min** |

### 🔍 First Run

On your first run:
1. Browser will open
2. You may need to log into Google
3. Accept any Gemini terms if prompted
4. Press Enter in terminal after logging in
5. Script continues automatically

### 📊 Customization

After generation, you can:
- Edit slides in PowerPoint
- Change colors and fonts
- Add your company logo
- Reorder sections
- Add additional images
- Customize speaker notes

---

## Troubleshooting Examples

### Issue: Browser Opens But Nothing Happens

**Solution:**
```
1. Check if you're on gemini.google.com
2. Look for login prompt
3. Log in if needed
4. Press Enter in terminal
```

### Issue: "Could not find Gemini input box"

**Cause:** Gemini UI may have changed

**Solution:** Edit `modules/research.py` line 85:
```python
# Try these selectors one by one
input_selectors = [
    "rich-textarea",
    "textarea",
    "div[contenteditable='true']",
    "[data-placeholder='Enter']",  # Add new ones
]
```

### Issue: Not Enough Slides Generated

**Solution:**
- Wait longer for Gemini research (2-3 minutes)
- Use more specific topic
- Increase slide count when prompted
- Check internet connection

---

## Sample Topics by Category

### Technology
- "Blockchain Technology and Cryptocurrencies"
- "5G Networks and Their Impact"
- "Cybersecurity Best Practices 2024"
- "Cloud Computing Architecture"

### Business
- "Startup Funding Strategies"
- "E-Commerce Trends 2024"
- "Leadership and Management Skills"
- "Financial Planning Basics"

### Science
- "CRISPR and Gene Editing"
- "Space Exploration Technologies"
- "Neuroscience and Brain Research"
- "Sustainable Agriculture"

### Education
- "Online Learning Platforms"
- "Educational Psychology"
- "STEM Education Importance"
- "Study Techniques and Methods"

### Health
- "Mental Health Awareness"
- "Nutrition and Healthy Eating"
- "Exercise Science"
- "Medical Technology Advances"

---

## Output Structure

Every run creates:

```
output/
└── [Topic]_[Timestamp]/
    ├── [Topic]_presentation.pptx    ← Main file
    └── charts/
        ├── chart_1.png
        ├── chart_2.png
        ├── chart_3.png
        └── ...
```

**Presentation Structure:**
1. Title Slide
2. Table of Contents
3-5. Introduction Section
6-20. Main Content (organized into logical sections)
21-23. Visualizations integrated throughout
24-25. Conclusion
26. References
27. Thank You

---

## Advanced Usage

### Batch Processing (Future)

Create `batch_generate.py`:

```python
topics = [
    "Artificial Intelligence",
    "Renewable Energy",
    "Digital Marketing"
]

for topic in topics:
    # Run generator for each topic
    # Save to separate folders
```

### Custom Themes (Future)

Modify `modules/ppt_builder.py`:

```python
THEME_COLORS = {
    'primary': RGBColor(YOUR_R, YOUR_G, YOUR_B),
    # ... customize colors
}
```

---

## Validation Checklist

After generation, verify:

- [ ] Presentation file exists and opens
- [ ] All slides have titles
- [ ] Content is relevant to topic
- [ ] Charts are visible and clear
- [ ] Speaker notes are present
- [ ] No duplicate slides
- [ ] Proper formatting throughout
- [ ] References included

---

**Ready to start? Run:**

```bash
python ppt_generator.py
```

**Happy Presenting! 🎉**
