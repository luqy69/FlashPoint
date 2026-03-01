"""
Second pass: Fix ALL remaining non-ASCII characters that could cause UnicodeEncodeError
"""
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Extended map covering ALL remaining chars from first scan
EMOJI_MAP = {
    # From warnings in first scan
    '\U0001f4cb': '[*]',  # 📋
    '\u2022': '-',         # • (bullet)
    '\u00b7': '-',         # · (middle dot)
    '\U0001f4d5': '[*]',  # 📕
    '\u2139': '[i]',       # ℹ
    '\U0001f3a8': '[*]',  # 🎨
    '\U0001f310': '[*]',  # 🌐
    '\U0001f4f1': '[*]',  # 📱
    '\u23f3': '[...]',     # ⏳
    '\U0001f4ac': '[*]',  # 💬
    '\U0001f4d6': '[*]',  # 📖
    '\u2192': '->',        # →
    '\U0001f4d0': '[*]',  # 📐
}

def fix_file(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"  ERROR reading {filepath}: {e}")
        return False
    
    original = content
    changes = []
    
    for emoji, replacement in EMOJI_MAP.items():
        if emoji in content:
            count = content.count(emoji)
            content = content.replace(emoji, replacement)
            changes.append(f"  Replaced {repr(emoji)} x{count} -> {replacement}")
    
    if content != original:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"FIXED: {filepath}")
        for change in changes:
            print(change)
        return True
    return False

# Fix files
modules_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'modules')
fixed = 0

print("=" * 60)
print("PASS 2: Fixing remaining non-ASCII characters")
print("=" * 60)

for filename in os.listdir(modules_dir):
    if filename.endswith('.py'):
        filepath = os.path.join(modules_dir, filename)
        if fix_file(filepath):
            fixed += 1

wizard = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ppt_wizard.py')
if os.path.exists(wizard):
    if fix_file(wizard):
        fixed += 1

print(f"\nFixed {fixed} file(s) in pass 2")

# Final verification - check for ANY remaining non-ASCII
print("\n" + "=" * 60)
print("FINAL VERIFICATION")
print("=" * 60)

all_files = [os.path.join(modules_dir, f) for f in os.listdir(modules_dir) if f.endswith('.py')]
all_files.append(wizard)

clean = True
for filepath in all_files:
    with open(filepath, 'r', encoding='utf-8') as f:
        for line_num, line in enumerate(f, 1):
            for char in line:
                if ord(char) > 127 and char not in '\r\n\t':
                    print(f"  REMAINING: {os.path.basename(filepath)} line {line_num}: {repr(char)} (U+{ord(char):04X})")
                    clean = False

if clean:
    print("ALL CLEAN! No non-ASCII characters remaining.")
else:
    print("\nSome non-ASCII chars remain (may be in string literals/comments).")
