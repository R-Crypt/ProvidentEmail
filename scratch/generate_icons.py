import os
import sys

# Ensure pillow is installed
try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("Pillow not found. Installing it...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])
    from PIL import Image, ImageDraw, ImageFont

def generate_icon(size, filename):
    # Create image with Provident dark blue theme (#1b2a47)
    img = Image.new('RGB', (size, size), color='#1b2a47')
    draw = ImageDraw.Draw(img)
    
    # Draw gold border (#b39257)
    border_width = max(1, int(size * 0.05))
    draw.rectangle([0, 0, size - 1, size - 1], outline='#b39257', width=border_width)
    
    # Draw text "P" (or "PP") in white/gold
    text = "P" if size < 32 else "PP"
    # Choose font size proportional to image size
    font_size = max(8, int(size * 0.5))
    
    try:
        # Try to load Segoe UI or Arial, fallback to default font
        font = ImageFont.truetype("Arial", font_size)
    except IOError:
        font = ImageFont.load_default()
        
    # Get text dimensions to center it
    try:
        # Pillow >= 10.0.0
        left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
        text_width = right - left
        text_height = bottom - top
    except AttributeError:
        # Older Pillow versions
        text_width, text_height = draw.textsize(text, font=font)
        
    x = (size - text_width) / 2
    y = (size - text_height) / 2 - (size * 0.05 if size > 16 else 0)
    
    # Draw the text in Provident gold
    draw.text((x, y), text, fill='#b39257', font=font)
    
    # Ensure folder exists
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    img.save(filename, 'PNG')
    print(f"Generated {filename} ({size}x{size})")

if __name__ == "__main__":
    assets_dir = "/Users/rayaankhan/Desktop/ProvidentClassifier/provident-email-classifier/outlook-addin/assets"
    sizes = [16, 32, 64, 80, 128]
    for size in sizes:
        generate_icon(size, os.path.join(assets_dir, f"icon-{size}.png"))
    print("All icons generated successfully.")
