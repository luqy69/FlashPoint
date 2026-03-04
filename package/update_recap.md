# FlashPoint AI PowerPoint Generator - Update Recap

This document outlines all the major features, improvements, and bug fixes added to the application since before the major UI update.

## 🌟 Major UI & UX Overhaul
*   **Modern Graphical Interface (GUI):** Transformed the tool from a purely command-line interface into a full-featured, modern desktop application (`ppt_wizard.py`).
*   **CustomTkinter Integration:** Built with a sleek, dark-mode native interface using modern UI components.
*   **Real-time Progress Tracking:** Added dynamic progress bars, status indicators, and spinning loaders to provide clear feedback during the (sometimes lengthy) AI generation process.
*   **Theme Selection Previews:** Users can now visually select presentation themes from a dropdown before generation.
*   **Interactive Error Handling:** Added clean pop-up dialogs to explain errors rather than crashing silently.

## 🧠 Core AI Integration Enhancements
*   **Two-Stage Generation Pipeline:** Completely overhauled how Gemini is queried.
    *   **Stage 1:** Gathers deep, comprehensive research on the topic based on the selected academic level.
    *   **Stage 2:** Restructures the raw research into a strict, highly formatted slide-by-slide layout.
*   **Customizable Prompts (`config/` folder):** Users can now fully edit the prompts sent to Gemini.
    *   Level 1-3 Prompts (`prompt_lvl_1`, `prompt_lvl_2`, `prompt_lvl_3`)
    *   Slide Formatting Prompt (`slide_formating_prompt`)
*   **Robust DOM State Detection:** Replaced unreliable string-matching with precise JavaScript DOM inspection to determine exactly when Gemini has finished generating text.

## 📊 Presentation Quality & Formatting
*   **Rich Slide Typologies:** The AI now intelligently categorizes slides into distinct types:
    *   `CONCEPT`, `MECHANISM` (Sequence/Step-by-step)
    *   `CLASSIFICATION` (Categorized bullets)
    *   `DEEP DIVE` (Definition, Components, Real-World Examples)
    *   `SUMMARY`, `APPLICATION`, `ANALYSIS`, `EVIDENCE`
*   **Explanation Text:** Added a dedicated "Explanation" section below the bullet points on each slide to provide context without overcrowding the slide with text.
*   **Custom Bullet Limits:** Dynamically adjusts the maximum number of bullets based on the slide type (e.g., up to 10 bullets for deep classification slides).
*   **Speaker Notes Removal:** Automatically strips out unnecessary or redundant speaker notes as per user preference, keeping the output clean.

## 🛠️ Reliability & Debugging Arsenal
*   **Extraction Deduplication:** Fixed a major bug where text subtraction would capture the same Gemini response multiple times (resulting in 50+ tripled slides).
*   **Garbage Cleanup:** Automatically removes Gemini UI artifacts (e.g., "Show thinking", "Gemini said") that sometimes bleed into the extracted text.
*   **Comprehensive Debug Logging:** Automatically generates detailed logs for every run to help diagnose pipeline failures:
    *   `second_prompt_output.txt`: The exact raw text extracted from Gemini's formatting stage.
    *   `slide_parse_debug.txt`: How the internal parser interpreted the raw text into distinct slide data.
    *   `ppt_build_debug.txt`: Exactly what was rendered onto the final PowerPoint slides (Titles, Bullets, Executed Actions).

## 🚀 Build & Deployment
*   **Standalone Executable:** Fully compiles into a single `.exe` file using PyInstaller.
*   **Self-Contained Package:** Automatically bundles into a `package/` directory complete with the executable, the `config/` prompt folder, and an `output/` destination for generated presentations.
