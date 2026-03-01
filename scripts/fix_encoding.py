import os

file_path = r"g:\desktop\code\ppt\modules\research.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replacements
replacements = {
    "✓": "[+]",
    "✗": "[!]",
    "⚠": "[!]",
    "📄": "[*]",
    "📚": "[*]",
    "📁": "[*]",
    "📂": "[+]",
    "💾": "[*]",
    "💣": "[!]"
}

new_content = content
for emoji, ascii_char in replacements.items():
    new_content = new_content.replace(emoji, ascii_char)

if new_content != content:
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    print("Successfully replaced emojis.")
else:
    print("No emojis found to replace.")
