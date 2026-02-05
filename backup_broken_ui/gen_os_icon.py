from PIL import Image, ImageDraw, ImageFont
import os

def create_icon():
    try:
        size = (256, 256)
        # Create transparent image
        img = Image.new('RGBA', size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Draw Blue Circle (Outlookish Blue)
        blue = (0, 120, 212) 
        margin = 10
        draw.ellipse([(margin, margin), (size[0]-margin, size[1]-margin)], fill=blue)

        # Load Font
        try:
            # Try loading Arial
            font = ImageFont.truetype("arial.ttf", 100)
        except IOError:
            try:
                 font = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 100)
            except:
                 font = ImageFont.load_default()
                 print("Using default font")

        # Draw "OS" text
        text = "OS"
        
        # Calculate text size for centering
        # Using textbbox if available (Pillow 8+), else textsize
        if hasattr(draw, "textbbox"):
            left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
            w = right - left
            h = bottom - top
        else:
            w, h = draw.textsize(text, font=font)

        # Center
        x = (size[0] - w) / 2
        y = (size[1] - h) / 2
        
        # Adjust for baseline if needed, but centering bbox is usually enough
        # Move up slightly to visually center in circle
        y -= 10 

        draw.text((x, y), text, fill="white", font=font)

        # Ensure dir exists
        if not os.path.exists("icons"):
            os.makedirs("icons")

        # Save
        icon_path = "icons/os_sidebar.ico"
        img.save(icon_path, format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32)])
        print("Successfully created " + icon_path)
        
    except Exception as e:
        print("Error creating icon: " + str(e))

if __name__ == "__main__":
    create_icon()
