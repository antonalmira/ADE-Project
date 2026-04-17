import os
from PIL import Image
from utils import log_message

def crop_and_save(image_path, crop_left=0, crop_top=0, crop_right=0, crop_bottom=0, temp_dir="temp_cropped_images"):
    if not os.path.exists(image_path):
        log_message(f"Image not found for cropping: {image_path}")
        return None
    try:
        with Image.open(image_path) as img:
            width, height = img.size
            if width > crop_left + crop_right and height > crop_top + crop_bottom:
                left = crop_left
                top = crop_top
                right = width - crop_right
                bottom = height - crop_bottom
                cropped_img = img.crop((left, top, right, bottom))
                # Use original filename for matching
                cropped_filename = os.path.basename(image_path)
                cropped_path = os.path.join(temp_dir, cropped_filename)
                cropped_img.save(cropped_path)
                log_message(f"Cropped and saved image: {cropped_path}")
                return cropped_path
            else:
                log_message(f"Skipping cropping for {image_path}: Dimensions too small")
                return image_path
    except Exception as e:
        log_message(f"Error cropping image {image_path}: {str(e)}")
        return None