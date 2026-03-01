"""
Generate FlashPoint AI Icon - IMPROVED VERSION
Elegant quill/feather pen with navy background and gold accents
"""
from PIL import Image, ImageDraw
import os

# Create high-res image for icon (256x256)
size = 256
img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

# Colors matching minimalist UI and user's specifications
navy = (15, 23, 42)  # #0F172A
gold = (218, 165, 32)  # Goldenrod
white = (255, 255, 255)
light_gray = (240, 240, 245)

# Draw navy circle background
circle_center = (size//2, size//2)
circle_radius = 115
draw.ellipse([circle_center[0]-circle_radius, circle_center[1]-circle_radius,
              circle_center[0]+circle_radius, circle_center[1]+circle_radius],
             fill=navy)

# Draw gold border ring
border_width = 8
draw.ellipse([circle_center[0]-circle_radius, circle_center[1]-circle_radius,
              circle_center[0]+circle_radius, circle_center[1]+circle_radius],
             outline=gold, width=border_width)

# Draw elegant flowing feather (white/silver)
# Main feather shaft - slightly curved
shaft_points = []
for i in range(30):
    y = 50 + i * 4
    x_offset = int((i / 30) ** 2 * 5)  # Slight curve
    shaft_points.append((size//2 + x_offset, y))

# Draw thick shaft line
for i in range(len(shaft_points)-1):
    draw.line([shaft_points[i], shaft_points[i+1]], fill=white, width=7)

# Feather barbs - elegant flowing left side
for i in range(10):
    y_pos = 65 + i * 11
    length = 20 + i * 2
    angle_offset = i * 1.5
    start_x = size//2 + int((i / 30) ** 2 * 5)
    draw.line([(start_x, y_pos), (start_x - length, y_pos + angle_offset)], 
              fill=light_gray, width=2)

# Feather barbs - elegant flowing right side  
for i in range(10):
    y_pos = 65 + i * 11
    length = 20 + i * 2
    angle_offset = i * 1.5
    start_x = size//2 + int((i / 30) ** 2 * 5)
    draw.line([(start_x, y_pos), (start_x + length, y_pos + angle_offset)], 
              fill=light_gray, width=2)

# Gold nib at bottom - professional calligraphy style
nib_center_x = size//2 + int((25 / 30) ** 2 * 5)
nib_y = 165

# Nib body
nib_points = [
    (nib_center_x - 18, nib_y),      # Left top
    (nib_center_x, nib_y + 30),      # Bottom point
    (nib_center_x + 18, nib_y)       # Right top
]
draw.polygon(nib_points, fill=gold)

# Nib detail lines (make it look like real pen nib)
draw.line([(nib_center_x, nib_y), (nib_center_x, nib_y + 25)], fill=navy, width=2)
draw.line([(nib_center_x - 10, nib_y + 5), (nib_center_x + 10, nib_y + 5)], fill=navy, width=2)

# Save as PNG first
png_path = 'icon_new.png'
img.save(png_path, 'PNG')

# Convert to ICO with multiple sizes
icon_sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
images = []
for icon_size in icon_sizes:
    images.append(img.resize(icon_size, Image.Resampling.LANCZOS))

# Save as ICO
ico_path = 'icon.ico'
images[0].save(ico_path, format='ICO', sizes=[(img.width, img.height) for img in images])

print("[OK] Improved icon created successfully!")
print(f"   PNG: {png_path}")
print(f"   ICO: {ico_path}")
print(f"   Design: Elegant feather quill with gold nib on navy background")
print(f"   Sizes: {', '.join([f'{s[0]}x{s[1]}' for s in icon_sizes])}")
