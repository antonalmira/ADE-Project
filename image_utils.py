import os
from PIL import Image, ImageFile
from utils import log_message

# Allow PIL to load truncated/partially-broken image files instead of raising
# "broken data stream" immediately. The image may render with minor artifacts
# near the bottom, but it will not crash the pipeline.
ImageFile.LOAD_TRUNCATED_IMAGES = True


def crop_and_save(image_path, crop_left=0, crop_top=0, crop_right=0, crop_bottom=0,
                  temp_dir="temp_cropped_images"):
    """
    Crop an image and save the result to temp_dir.
    Returns the saved path on success, or None if the file is missing/unreadable.

    Fixes applied:
      - ImageFile.LOAD_TRUNCATED_IMAGES = True  → tolerates broken JPG data streams
      - img.load() called explicitly             → forces full decode inside try/except
                                                   so a bad file fails fast and cleanly
      - 'with Image.open()' context manager     → guarantees the file handle is released
                                                   before we return, preventing WinError 32
                                                   (file locked by another process)
      - P-mode (palette/GIF) images converted   → correct alpha compositing on white bg
      - Crop bounds clamped                      → prevents zero/negative-size crop crash
    """
    if not os.path.exists(image_path):
        log_message(f"Image not found: {image_path}")
        return None

    try:
        os.makedirs(temp_dir, exist_ok=True)

        with Image.open(image_path) as img:
            # Force full decode NOW, inside the try block, so a broken/truncated
            # file raises immediately rather than failing later during paste/resize.
            img.load()

            width, height = img.size

            # Clamp crop values so we never produce a zero-size or inverted region
            left   = min(max(0, int(crop_left)),  width  - 1)
            top    = min(max(0, int(crop_top)),   height - 1)
            right  = max(left + 1, width  - int(crop_right))
            bottom = max(top  + 1, height - int(crop_bottom))

            cropped_img = img.crop((left, top, right, bottom))

            # Composite transparency onto a white background for Word compatibility.
            # Palette ("P") images must be converted to RGBA first so that
            # split()[3] is actually the alpha channel (split("P") yields only 1 band).
            if cropped_img.mode == "RGBA":
                background = Image.new("RGB", cropped_img.size, (255, 255, 255))
                background.paste(cropped_img, mask=cropped_img.split()[3])
                cropped_img = background
            elif cropped_img.mode == "P":
                rgba = cropped_img.convert("RGBA")
                background = Image.new("RGB", rgba.size, (255, 255, 255))
                background.paste(rgba, mask=rgba.split()[3])
                cropped_img = background
            elif cropped_img.mode not in ("RGB", "L"):
                cropped_img = cropped_img.convert("RGB")

            cropped_filename = os.path.basename(image_path)
            cropped_path = os.path.join(temp_dir, cropped_filename)

            # Save while cropped_img is still in scope (the 'with' block for the
            # source file has already exited, so the original handle is closed).
            cropped_img.save(cropped_path, "PNG")

        # cropped_img is a new in-memory image — it holds no file handle.
        # Returning here is safe; the source file lock is fully released.
        return cropped_path

    except Exception as e:
        log_message(f"Cropping Error on {image_path}: {str(e)}")
        return None