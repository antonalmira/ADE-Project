import os
import shutil
from PIL import Image
from utils import log_message

def crop_and_save(image_path, crop_left=0, crop_top=0, crop_right=0, crop_bottom=0, temp_dir="temp_cropped_images"):
    """
    Centralized cropping function.
    left/top/right/bottom are pixels to REMOVE from each respective edge.
    """
    if not os.path.exists(image_path):
        log_message(f"Image not found: {image_path}")
        return None
        
    try:
        # Fix: Ensure the destination directory exists before saving
        os.makedirs(temp_dir, exist_ok=True)

        with Image.open(image_path) as img:
            width, height = img.size
            
            # Fix: Calculate boundaries ensuring we don't go out of bounds or create negative sizes
            left = min(max(0, int(crop_left)), width - 1)
            top = min(max(0, int(crop_top)), height - 1)
            right = max(left + 1, width - int(crop_right))
            bottom = max(top + 1, height - int(crop_bottom))
            
            cropped_img = img.crop((left, top, right, bottom))
            
            # Handle PNG transparency (convert to RGB with white background for Word)
            if cropped_img.mode in ("RGBA", "P"):
                background = Image.new("RGB", cropped_img.size, (255, 255, 255))
                # Paste using alpha channel as mask
                background.paste(cropped_img, mask=cropped_img.split()[3] if cropped_img.mode == "RGBA" else None)
                cropped_img = background
            
            cropped_filename = os.path.basename(image_path)
            cropped_path = os.path.join(temp_dir, cropped_filename)
            cropped_img.save(cropped_path, "PNG")
            return cropped_path
            
    except Exception as e:
        log_message(f"Cropping Error on {image_path}: {str(e)}")
        return None