# Advanced Options Feature - Complete! 🎉

The AI PowerPoint Generator now includes a powerful **Advanced Options** step that gives you full control over your presentations.

## What's New

### 1. **Wider Window** ✅
- Increased from 700px to 850px
- No more text cutoff!

### 2. **New Advanced Options Step** ✅
Inserted as Step 2 (after Topic Input), includes:

#### Slide Count Selection
- Choose from: 10, 15, 20, 25, or 30 slides
- Default: 25 slides
- Dropdown (Combobox) for easy selection

#### Research Depth Selection (3 Levels)
Choose how detailed the AI research should be:

- **Level 1: Basic** 
  - For general audiences
  - Simple explanations with analogies
  - Fun facts and real-world examples
  - 5-6 bullet points per section
  
- **Level 2: Professional**
  - For technical presentations
  - Data-driven with 100-150 word paragraphs
  - Includes 2024-2026 studies and statistics
  - Comparative analysis and protocols

- **Level 3: Master Thesis** (Default)
  - Maximum detail (5000-word depth)
  - 250-300 word detailed paragraphs per point
  - 8-module structure with mechanisms and evidence
  - Expert-level content with clinical data

### 3. **Research Prompt Templates** ✅
Created three separate prompt files:
- `prompt_lvl_1` - Basic educational content
- `prompt_lvl_2` - Professional technical content
- `prompt_lvl_3` - Master thesis-level depth

These are loaded at runtime and dynamically applied based on user selection.

## Technical Implementation

### Files Modified
1. **ppt_wizard.py**
   - Widened window to 850px
   - Added `slide_count` and `research_level` to config (defaults: 25, Level 3)
   - Created `show_advanced_options()` method with Combobox and radio buttons
   - Updated step navigation (now 7 steps total)

2. **modules/research.py**
   - Added `load_prompt_templates()` function at module init
   - Updated `perform_deep_research()` to accept `research_level` parameter
   - Updated `gather_research()` to accept and pass `research_level`
   - Dynamically loads appropriate template and replaces `[Insert Topic Here]` with topic

3. **Prompt Template Files**
   - Created `prompt_lvl_1`, `prompt_lvl_2`, `prompt_lvl_3` in root directory
   - Each contains complete research instructions for that level

## User Workflow

1. **Step 1**: Enter topic
2. **Step 2 (NEW)**: Advanced Options
   - Select slide count (10-30)
   - Choose research depth (1-3)
3. **Step 3**: Choose theme
4. **Step 4**: Select exports
5. **Step 5**: Generation
6. **Step 6**: Complete

## Testing Notes

- Defaults maintain current behavior (25 slides, Level 3)
- All 3 research levels tested and working
- Window width sufficient for all content
- Config values properly saved and passed through to research

---

**Status**: ✅ **COMPLETE** - All features implemented and tested!
