from PIL import Image

def process_and_crop():
    input_path = r"C:\Users\Administrator\.gemini\antigravity\brain\b4338db8-902f-47af-b3f5-9efa8d7d1f5f\lark_app_icon_transparent_1772597154251.png"
    output_path = r"c:\Users\Administrator\Desktop\docx-formatter\src\ui\icons\app_icon.png"
    
    img = Image.open(input_path).convert("RGBA")
    
    # Optional: ensure white is explicitly transparent if the AI didn't perfectly clear the background
    data = img.getdata()
    newData = []
    threshold = 240
    for item in data:
        # If it's mostly white and solid, make it transparent
        if item[0] > threshold and item[1] > threshold and item[2] > threshold:
            newData.append((255, 255, 255, 0))
        else:
            newData.append(item)
    img.putdata(newData)
    
    # Get bounding box of non-zero alpha pixels
    bbox = img.getbbox()
    if bbox:
        # Pad slightly so it breathes in the window header
        margin = 10 
        new_bbox = (
            max(0, bbox[0] - margin),
            max(0, bbox[1] - margin),
            min(img.width, bbox[2] + margin),
            min(img.height, bbox[3] + margin)
        )
        img = img.crop(new_bbox)
        
    img.save(output_path, "PNG")
    print("Transparent icon cropped and saved!")

if __name__ == "__main__":
    process_and_crop()
