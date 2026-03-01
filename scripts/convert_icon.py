"""
Convert user's final_icon.png to icon.ico for executable
"""
from PIL import Image

# Load the user's icon
img = Image.open('final_icon.png')

# Ensure it's in RGBA mode
img = img.convert('RGBA')

# Create multiple sizes for Windows icon
icon_sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
images = []

for size in icon_sizes:
    resized = img.resize(size, Image.Resampling.LANCZOS)
    images.append(resized)

# Save as ICO
ico_path = 'icon.ico'
images[0].save(ico_path, format='ICO', sizes=[(img.width, img.height) for img in images])

print("[OK] Custom icon converted successfully!")
print(f"   Input: final_icon.png")
print(f"   Output: icon.ico")
print(f"   Sizes: {', '.join([f'{s[0]}x{s[1]}' for s in icon_sizes])}")
