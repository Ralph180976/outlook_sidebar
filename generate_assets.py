import os
import sys
from PIL import Image

def generate_assets(source_path):
    if not os.path.exists(source_path):
        print("Error: Source image not found at {}".format(source_path))
        return

    # Ensure directories exist
    if not os.path.exists("icons"):
        os.makedirs("icons")
    if not os.path.exists("Assets"):
        os.makedirs("Assets")

    try:
        img = Image.open(source_path)
        
        # Generate ICO (Standard Windows Icon)
        # Sizes: 16, 32, 48, 64, 128, 256
        print("Generating icons/app.ico...")
        img.save("icons/app.ico", format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
        
        # Generate PNG (High Res for General Use)
        print("Generating icons/app.png...")
        img.resize((256, 256), Image.ANTIALIAS).save("icons/app.png")

        # Generate AppxManifest Assets
        assets = {
            "StoreLogo.png": (50, 50),
            "Square150x150Logo.png": (150, 150),
            "Square44x44Logo.png": (44, 44)
        }

        for filename, size in assets.items():
            print("Generating Assets/{} ({})...".format(filename, size))
            # Resize and save
            resized_img = img.resize(size, Image.ANTIALIAS)
            resized_img.save(os.path.join("Assets", filename))

        print("Asset generation complete!")

    except Exception as e:
        print("Failed to generate assets: {}".format(e))

if __name__ == "__main__":
    # Source path from the artifacts directory (User provided image)
    source_image = r"C:\Users\Ralph.SOUTHERNC\.gemini\antigravity\brain\9be3a381-259f-4531-b015-0978c7c6088b\uploaded_media_1770153728541.png"
    generate_assets(source_image)
