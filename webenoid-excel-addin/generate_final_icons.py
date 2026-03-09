import os
import glob
from PIL import Image

def get_latest_media():
    brain_dir = r"C:\Users\Abhishek N\.gemini\antigravity\brain\64dea96f-7d70-4d06-a78b-5afb39d58360"
    files = glob.glob(os.path.join(brain_dir, "*.png"))
    if not files:
        return None
    files.sort(key=os.path.getmtime, reverse=True)
    return files[0]

def generate_icons(source_path, dest_dir, sizes):
    os.makedirs(dest_dir, exist_ok=True)
    try:
        img = Image.open(source_path)
    except Exception as e:
        print(f"Error opening image: {e}")
        return
        
    if img.mode != 'RGBA':
        img = img.convert('RGBA')

    for size in sizes:
        resized_img = img.resize((size, size), Image.Resampling.LANCZOS)
        # Use BOTH logo- and icon- to make sure it covers whichever the user wants 
        out_path_logo = os.path.join(dest_dir, f"logo-{size}.png")
        out_path_icon = os.path.join(dest_dir, f"icon-{size}.png")
        resized_img.save(out_path_logo, format="PNG")
        resized_img.save(out_path_icon, format="PNG")
        print(f"Saved: {out_path_logo}")
        print(f"Saved: {out_path_icon}")

if __name__ == "__main__":
    latest_img = get_latest_media()
    if not latest_img:
        print("No image found!")
    else:
        print(f"Processing image: {latest_img}")
        assets_dir = "C:/Users/Abhishek N/Desktop/webenoid-ai/webenoid-excel-addin/src/taskpane/assets"
        sizes = [16, 32, 64, 80, 128, 300]
        generate_icons(latest_img, assets_dir, sizes)
