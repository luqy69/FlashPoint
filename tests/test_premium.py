"""Quick test to verify all new modules work end-to-end"""
import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from modules.themes import get_theme, PREMIUM_THEMES
from modules.ppt_builder import build_presentation
from modules.content_organizer import organize_content

# Simulate research data
research_data = {
    'raw_text': """Introduction to Artificial Intelligence
Artificial Intelligence (AI) is a branch of computer science that aims to create intelligent machines.
Machine Learning is a subset of AI that enables systems to learn from experience.
Deep Learning uses neural networks with multiple layers to analyze data.
Applications include NLP, computer vision, robotics, and autonomous vehicles.
Key Concepts in Machine Learning
Supervised learning trains a model on labeled data to make predictions.
Unsupervised learning finds hidden patterns without labeled examples.
Reinforcement learning trains agents by maximizing a reward signal.
Neural Networks and Deep Learning
Artificial neural networks are inspired by biological neural networks.
CNNs are specialized for processing images and grid-like data.
RNNs are designed for sequential data and time series analysis.
Transformers revolutionized NLP with self-attention mechanisms.
Impact on Society
AI is transforming healthcare, education, and transportation.
Ethical considerations include bias, job displacement, and privacy.
The future promises tremendous opportunities and challenges.""",
    'sections': [
        {'title': 'Introduction to AI', 'content': [
            'AI is a branch of computer science creating intelligent machines',
            'Machine Learning enables systems to learn from experience',
            'Deep Learning uses multi-layer neural networks',
            'Applications: NLP, computer vision, robotics, autonomous vehicles'
        ]},
        {'title': 'Key Concepts in Machine Learning', 'content': [
            'Supervised learning: training on labeled data for predictions',
            'Unsupervised learning: finding patterns without labels',
            'Reinforcement learning: maximizing reward signals'
        ]},
        {'title': 'Neural Networks and Deep Learning', 'content': [
            'Inspired by biological neural networks in the brain',
            'CNNs specialized for image and grid-like data processing',
            'RNNs designed for sequential and time series data',
            'Transformers use self-attention for NLP breakthroughs'
        ]},
        {'title': 'Impact on Society', 'content': [
            'Transforming healthcare, education, and transportation',
            'Key ethical concerns: algorithmic bias and job displacement',
            'Privacy implications of AI data collection',
            'Future promises both opportunities and challenges'
        ]},
    ],
    'total_sections': 4
}

# Test content organizer
print("=" * 60)
print("TESTING CONTENT ORGANIZER")
print("=" * 60)
slides = organize_content(research_data, 'Artificial Intelligence', target_slides=15)
print(f"\nOrganized into {len(slides)} slides:")
for s in slides:
    title = s.get('title', 'N/A')
    count = len(s.get('bullets', []))
    print(f"  - {title}: {count} bullets")

# Test building with each theme
print("\n" + "=" * 60)
print("TESTING PPT BUILDER WITH ALL 6 THEMES")
print("=" * 60)

os.makedirs('output/test_themes', exist_ok=True)

for theme_id in ['1', '2', '3', '4', '5', '6']:
    theme = get_theme(theme_id)
    config = {'theme_choice': theme_id}
    prs = build_presentation(
        'Artificial Intelligence', slides, [],
        f'output/test_themes/test.pptx', config
    )
    output_path = f'output/test_themes/test_{theme["name"]}.pptx'
    prs.save(output_path)
    print(f"[OK] {theme['name']} theme -> {len(prs.slides)} slides saved to {output_path}")

print("\n" + "=" * 60)
print("ALL TESTS PASSED!")
print("=" * 60)
