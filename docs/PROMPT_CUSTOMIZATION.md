# RESEARCH PROMPT CUSTOMIZATION GUIDE

## What Are These Files?

The `prompt_lvl_1`, `prompt_lvl_2`, and `prompt_lvl_3` files control how deeply FlashPoint AI researches your topics.

- **prompt_lvl_1**: Basic/Educational level (General audiences)
- **prompt_lvl_2**: Professional level (Technical presentations)
- **prompt_lvl_3**: Master Thesis level (Academic/Expert content)

## How to Customize

1. **Open** any of the three files in a text editor (Notepad, VS Code, etc.)
2. **Edit** the instructions to change how AI researches
3. **Save** the file
4. **Run** FlashPoint_AI.exe - it will automatically use your custom prompts!

## Important Notes

✅ **You can edit these files anytime** - No need to rebuild the application!  
✅ The placeholder `[Insert Topic Here]` will be automatically replaced with your topic  
✅ Changes take effect immediately on next run  

⚠️ **Backup your prompts** before making major changes  
⚠️ If FlashPoint can't find these files, it uses built-in defaults

## Example Customization

Want more fun facts in Level 1? Add this line:
```
- Include 3-5 interesting "Did You Know?" facts per section
- Add relevant analogies and real-world comparisons
```

Want specific data in Level 3? Add this requirement:
```
- Cite at least 10 peer-reviewed studies from 2024-2026
- Include molecular/chemical formulas where applicable
```

## Troubleshooting

**Problem**: Prompts not loading  
**Solution**: Make sure files are named exactly `prompt_lvl_1`, `prompt_lvl_2`, `prompt_lvl_3` (no .txt extension!)

**Problem**: Research output hasn't changed  
**Solution**: Double-check you saved the file and selected the correct research level in the wizard

---

Happy customizing! 🎉
