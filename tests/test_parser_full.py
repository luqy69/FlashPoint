"""Test parser with the user's ACTUAL second_prompt_output.txt content"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

# Read the actual extracted output
with open(r'g:\desktop\code\ppt\package\output\kidney_pathology_20260301_101343\second_prompt_output.txt', 'r', encoding='utf-8') as f:
    test_data = f.read()

from modules.content_organizer import parse_slides

slides = parse_slides(test_data)

print(f"Total slides parsed: {len(slides)}")
print()

for i, s in enumerate(slides):
    slide_type = s.get('type', '?')
    title = s.get('title', '?')
    subtitle = s.get('subtitle', '')
    bullets = s.get('bullets', [])
    explanation = s.get('explanation', '')
    notes = s.get('notes', '')
    
    print(f"Slide {i+1}: [{slide_type}] {title}")
    if subtitle:
        print(f"  Subtitle: {subtitle}")
    print(f"  Bullets ({len(bullets)}):")
    for j, b in enumerate(bullets):
        text = b[:90]
        print(f"    [{j+1}] {text}{'...' if len(b) > 90 else ''}")
    if explanation:
        print(f"  Explanation: {explanation[:100]}{'...' if len(explanation) > 100 else ''}")
    if notes:
        print(f"  Notes: {notes[:80]}{'...' if len(notes) > 80 else ''}")
    else:
        print(f"  Notes: (empty - good!)")
    print()
