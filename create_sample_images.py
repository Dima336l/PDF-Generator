from PIL import Image, ImageDraw, ImageFont
import os

def create_sample_images():
    """Create sample property images for testing"""
    
    # Create sample_images directory if it doesn't exist
    if not os.path.exists('sample_images'):
        os.makedirs('sample_images')
    
    # Sample image data
    images = [
        {
            'filename': 'exterior_front.jpg',
            'title': 'Front Exterior',
            'description': 'Beautiful front view of the property',
            'color': (34, 139, 34)  # Forest green
        },
        {
            'filename': 'living_room.jpg', 
            'title': 'Living Room',
            'description': 'Spacious and modern living area',
            'color': (255, 140, 0)  # Dark orange
        },
        {
            'filename': 'kitchen.jpg',
            'title': 'Modern Kitchen',
            'description': 'Fully equipped modern kitchen',
            'color': (70, 130, 180)  # Steel blue
        },
        {
            'filename': 'bedroom.jpg',
            'title': 'Master Bedroom',
            'description': 'Comfortable master bedroom',
            'color': (147, 112, 219)  # Medium purple
        },
        {
            'filename': 'bathroom.jpg',
            'title': 'Bathroom',
            'description': 'Clean and modern bathroom',
            'color': (0, 191, 255)  # Deep sky blue
        },
        {
            'filename': 'garden.jpg',
            'title': 'Garden',
            'description': 'Well-maintained rear garden',
            'color': (50, 205, 50)  # Lime green
        }
    ]
    
    for img_data in images:
        # Create a 800x600 image
        img = Image.new('RGB', (800, 600), img_data['color'])
        draw = ImageDraw.Draw(img)
        
        # Try to use a default font, fallback to basic if not available
        try:
            font_large = ImageFont.truetype("arial.ttf", 48)
            font_medium = ImageFont.truetype("arial.ttf", 24)
            font_small = ImageFont.truetype("arial.ttf", 18)
        except:
            font_large = ImageFont.load_default()
            font_medium = ImageFont.load_default()
            font_small = ImageFont.load_default()
        
        # Add title
        title_bbox = draw.textbbox((0, 0), img_data['title'], font=font_large)
        title_width = title_bbox[2] - title_bbox[0]
        title_x = (800 - title_width) // 2
        draw.text((title_x, 200), img_data['title'], fill='white', font=font_large)
        
        # Add description
        desc_bbox = draw.textbbox((0, 0), img_data['description'], font=font_medium)
        desc_width = desc_bbox[2] - desc_bbox[0]
        desc_x = (800 - desc_width) // 2
        draw.text((desc_x, 280), img_data['description'], fill='white', font=font_medium)
        
        # Add filename
        filename_text = f"File: {img_data['filename']}"
        filename_bbox = draw.textbbox((0, 0), filename_text, font=font_small)
        filename_width = filename_bbox[2] - filename_bbox[0]
        filename_x = (800 - filename_width) // 2
        draw.text((filename_x, 350), filename_text, fill='white', font=font_small)
        
        # Add a decorative border
        draw.rectangle([50, 50, 750, 550], outline='white', width=3)
        
        # Save the image
        filepath = os.path.join('sample_images', img_data['filename'])
        img.save(filepath, 'JPEG', quality=95)
        print(f"Created: {filepath}")
    
    print(f"\n✅ Created {len(images)} sample images in 'sample_images' folder")
    print("You can now run the app and it will automatically load these images!")

if __name__ == "__main__":
    create_sample_images()
