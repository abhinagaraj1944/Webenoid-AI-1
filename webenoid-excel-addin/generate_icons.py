from PIL import Image, ImageDraw, ImageFont
import os

def create_w_icon(size, output_path):
    # Create white background image
    image = Image.new('RGBA', (size, size), (255, 255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    # Try to load a nice font, or fallback to default
    try:
        from PIL import ImageFont
        try:
            # Try a standard windows font
            font = ImageFont.truetype("arialbd.ttf", int(size * 0.7))
        except OSError:
            font = ImageFont.load_default()
    except ImportError:
        font = None
    
    # Text to draw
    text = "W"
    
    # Calculate text bounding box
    if font and hasattr(font, 'getbbox'):
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
    else:
        text_width = size * 0.5
        text_height = size * 0.5
        
    x = (size - text_width) / 2
    # Adjust y to roughly center
    y = (size - text_height) / 2 - (size * 0.1)
    
    draw.text((x, y), text, fill="black", font=font)
    
    # Save the image
    image.save(output_path, "PNG")

if __name__ == "__main__":
    assets_dir = r"c:\Users\Abhishek N\Desktop\webenoid-ai\webenoid-excel-addin\src\taskpane\assets"
    os.makedirs(assets_dir, exist_ok=True)
    
    sizes = [16, 32, 64, 80, 128]
    for size in sizes:
        output_path = os.path.join(assets_dir, f"icon-{size}.png")
        create_w_icon(size, output_path)
        print(f"Created {output_path}")
