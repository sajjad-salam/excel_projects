from PIL import Image
import pytesseract
import re


def extract_text_from_image(image_path):
    """
    Extract text from an image file.

    Args:
        image_path (str): Path to the image file

    Returns:
        str: Extracted text from the image
    """
    try:
        # Open the image file
        img = Image.open(image_path)

        # Convert to grayscale for better OCR
        img = img.convert('L')

        # Try Arabic first
        text_ar = pytesseract.image_to_string(
            img, lang='ara', config='--psm 6')

        # Try English
        text_en = pytesseract.image_to_string(
            img, lang='eng', config='--psm 6')

        # Combine results
        combined_text = ""
        if text_ar.strip():
            combined_text += text_ar.strip() + "\n"
        if text_en.strip():
            combined_text += text_en.strip()

        # Clean the text
        # Remove empty lines
        cleaned_text = re.sub(r'\n\s*\n', '\n', combined_text)
        # Remove extra spaces
        cleaned_text = re.sub(r'\s+', ' ', cleaned_text)

        return cleaned_text.strip()

    except Exception as e:
        return f"Error processing image: {str(e)}"


# Example usage:
text = extract_text_from_image("image.png")
print(text)
