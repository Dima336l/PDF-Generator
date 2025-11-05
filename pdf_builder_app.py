import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import os
import sys
from reportlab.lib.pagesizes import A4, letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics import renderPDF
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.textlabels import Label
from reportlab.lib.colors import HexColor
from reportlab.platypus import KeepTogether, Frame
from reportlab.platypus.flowables import Flowable
from reportlab.lib.utils import ImageReader
import datetime

def format_date_with_ordinal(date_obj):
    """Format date with ordinal suffix (1st, 2nd, 3rd, 4th, etc.)"""
    day = date_obj.day
    if 10 <= day % 100 <= 20:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    return date_obj.strftime(f"{day}{suffix} %B %Y")

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def create_placeholder_drawing(width, height):
    """Create a gray placeholder Drawing for use in PDF"""
    placeholder = Drawing(width, height)
    placeholder.add(Rect(0, 0, width, height, fillColor=HexColor('#CCCCCC'), strokeColor=HexColor('#CCCCCC')))
    return placeholder

class CoverPageFlowable(Flowable):
    """Custom flowable for cover page with absolute positioning"""
    def __init__(self, data, images, accent_gold, primary_blue):
        Flowable.__init__(self)
        self.data = data
        self.images = images
        self.accent_gold = accent_gold
        self.primary_blue = primary_blue
        # Account for margins - use exact frame size from reportlab
        # Error showed frame size: 475.28 x 721.89 points
        # A4 is 595.28 x 841.89 points
        # So available space is: 475.28 x 721.89 points
        # Convert to inches: 475.28/72 = 6.6 inches, 721.89/72 = 10.0 inches
        # Use conservative margin: 1.8 inches total (0.9 per side) to ensure fit
        margin_total = 1.8*inch
        self.width = A4[0] - margin_total
        self.height = A4[1] - margin_total
        
    def draw(self):
        """Draw the cover page - note: reportlab uses bottom-left as origin"""
        canvas = self.canv
        # Use the declared width/height
        width = self.width
        height = self.height
        
        # Logo at top left - adjust Y for bottom-left origin
        logo_path = get_resource_path("logo.png")
        logo_bottom = None
        if os.path.exists(logo_path):
            try:
                logo_img = Image.open(logo_path)
                logo_width = 1.5*inch
                logo_height = logo_width * (logo_img.height / logo_img.width)
                logo_y = height - logo_height - 0.2*inch
                logo_bottom = logo_y
                canvas.drawImage(logo_path, 0, logo_y, 
                               width=logo_width, height=logo_height, preserveAspectRatio=True)
            except Exception as e:
                print(f"Error loading logo: {e}")
        
        # "Elevating your property experience" tagline below logo - smaller, italic style
        if logo_bottom is not None:
            tagline_y = logo_bottom - 0.25*inch
        else:
            tagline_y = height - 1.5*inch
        canvas.setFont("Helvetica-Oblique", 8)  # Smaller, italic font
        canvas.setFillColor(HexColor('#666666'))  # Gray color for subtlety
        canvas.drawString(0, tagline_y, "Elevating your property experience")
        
        # "Property Report" text below tagline - smaller, regular style
        canvas.setFont("Helvetica", 10)  # Smaller font
        canvas.setFillColor(colors.black)
        report_text_y = tagline_y - 0.25*inch
        canvas.drawString(0, report_text_y, "Property Report")
        
        # Property Address (large, bold) - handle text wrapping
        address = f"{self.data.get('address', '')}, {self.data.get('postal_code', '')}"
        canvas.setFont("Helvetica-Bold", 24)
        canvas.setFillColor(colors.black)
        # Split address if too long (simple approach)
        address_lines = []
        words = address.split()
        current_line = ""
        for word in words:
            test_line = current_line + (" " if current_line else "") + word
            if canvas.stringWidth(test_line, "Helvetica-Bold", 24) <= width:
                current_line = test_line
            else:
                if current_line:
                    address_lines.append(current_line)
                current_line = word
        if current_line:
            address_lines.append(current_line)
        
        # Move the address block up even more (starting from "5, Ridley Road")
        address_y = height - 2.0*inch  # Moved up from 2.5 to 2.0
        for i, line in enumerate(address_lines):
            canvas.drawString(0, address_y - (i * 0.35*inch), line)
        
        # Calculate where the address text area ends
        # Last line baseline - font height (24pt = ~0.33 inch)
        last_line_baseline = address_y - (len(address_lines) - 1) * 0.35*inch
        font_height = 24 / 72.0 * inch  # Convert 24pt to inches
        address_bottom = last_line_baseline - font_height
        
        # Add spacing between address and main image - reduced to push images up
        spacing_between_text_and_images = 0.1*inch  # Reduced spacing between address and main image
        
        # Main image - find exterior front image, or use first image as fallback, or use placeholder
        # Position main image below the address with spacing
        main_image_height = 4.5*inch
        main_image_bottom = address_bottom - spacing_between_text_and_images - main_image_height
        main_image_width = width
        
        # Find exterior front image (case-insensitive search)
        main_img_path = None
        main_img_index = -1
        if self.images and len(self.images) > 0:
            for idx, img_path in enumerate(self.images):
                filename = os.path.basename(img_path).lower()
                if 'exterior' in filename and 'front' in filename:
                    main_img_path = img_path
                    main_img_index = idx
                    break
            
            # Fallback to first image if no exterior front found
            if main_img_path is None:
                main_img_path = self.images[0]
                main_img_index = 0
        
        # Always draw main image (use placeholder if no image available)
        try:
            if main_img_path and os.path.exists(main_img_path):
                main_img = Image.open(main_img_path)
                # Calculate dimensions to fit while maintaining aspect ratio
                img_ratio = main_img.width / main_img.height
                target_ratio = main_image_width / main_image_height
                
                if img_ratio > target_ratio:
                    # Image is wider, fit to width
                    draw_width = main_image_width
                    draw_height = main_image_width / img_ratio
                else:
                    # Image is taller, fit to height
                    draw_height = main_image_height
                    draw_width = main_image_height * img_ratio
                
                # Center the image
                x_offset = (main_image_width - draw_width) / 2
                canvas.drawImage(main_img_path, x_offset, main_image_bottom, 
                               width=draw_width, height=draw_height, preserveAspectRatio=True)
            else:
                # Draw gray placeholder rectangle
                canvas.setFillColor(HexColor('#CCCCCC'))
                canvas.rect(0, main_image_bottom, main_image_width, main_image_height, fill=1, stroke=0)
        except Exception as e:
            print(f"Error adding main image: {e}")
            # Draw gray placeholder rectangle on error
            canvas.setFillColor(HexColor('#CCCCCC'))
            canvas.rect(0, main_image_bottom, main_image_width, main_image_height, fill=1, stroke=0)
        
        # Three thumbnail images below main image (exclude the main image)
        thumbnail_bottom = main_image_bottom - 2*inch
        thumbnail_height = 1.5*inch
        thumbnail_width = (width - 0.4*inch) / 3  # 3 thumbnails with spacing
        
        # Get list of images excluding the main image
        thumbnail_images = [img for idx, img in enumerate(self.images) if idx != main_img_index]
        
        # Track the lowest point of thumbnails to add padding above footer
        lowest_thumbnail_bottom = thumbnail_bottom
        
        # Always show 3 thumbnails (use placeholders if not enough images)
        for i in range(3):
            thumb_x = i * (thumbnail_width + 0.2*inch)
            thumb_y_position = thumbnail_bottom
            
            if i < len(thumbnail_images):
                # Use actual image
                try:
                    thumb_img_path = thumbnail_images[i]
                    if os.path.exists(thumb_img_path):
                        canvas.drawImage(thumb_img_path, thumb_x, thumb_y_position,
                                       width=thumbnail_width, height=thumbnail_height, preserveAspectRatio=False)
                    else:
                        # Draw gray placeholder
                        canvas.setFillColor(HexColor('#CCCCCC'))
                        canvas.rect(thumb_x, thumb_y_position, thumbnail_width, thumbnail_height, fill=1, stroke=0)
                except Exception as e:
                    print(f"Error adding thumbnail {i+1}: {e}")
                    # Draw gray placeholder on error
                    canvas.setFillColor(HexColor('#CCCCCC'))
                    canvas.rect(thumb_x, thumb_y_position, thumbnail_width, thumbnail_height, fill=1, stroke=0)
            else:
                # Draw gray placeholder for missing thumbnails
                canvas.setFillColor(HexColor('#CCCCCC'))
                canvas.rect(thumb_x, thumb_y_position, thumbnail_width, thumbnail_height, fill=1, stroke=0)
            
            # All thumbnails have the same bottom position
            lowest_thumbnail_bottom = thumbnail_bottom
        
        # Footer bar with gold background - fixed at the bottom of the page
        footer_height = 0.4*inch
        # Always position footer at the bottom of the page (accounting for margins)
        footer_y = 0  # Bottom of the available canvas area
        
        canvas.setFillColor(self.accent_gold)
        canvas.rect(0, footer_y, width, footer_height, fill=1, stroke=0)
        
        # Footer text (centered, white)
        report_date = format_date_with_ordinal(datetime.datetime.now())
        footer_text = f"Report created on {report_date}"
        canvas.setFont("Helvetica", 11)
        canvas.setFillColor(colors.white)
        text_width = canvas.stringWidth(footer_text, "Helvetica", 11)
        # Center text vertically in footer
        text_y = footer_y + (footer_height / 2) - 0.1*inch  # Adjust for font baseline
        canvas.drawString((width - text_width) / 2, text_y, footer_text)

class PDFBuilderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Property Investment Report Builder")
        self.root.geometry("1000x700")
        
        # Data storage
        self.images = []
        self.property_data = {}
        
        # Create main frame
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        
        self.create_widgets()
        self.load_sample_data()
        
    def create_widgets(self):
        """Create all the GUI widgets with tabs"""
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Create tabs
        self.create_property_tab()
        self.create_investment_tab()
        self.create_epc_tab()
        self.create_location_tab()
        self.create_images_tab()
        
        # Main buttons
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=1, column=0, pady=10)
        
        ttk.Button(button_frame, text="Generate Investment Report PDF", command=self.generate_pdf).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Clear All", command=self.clear_all).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Load Sample Data", command=self.load_sample_data).pack(side=tk.LEFT)
        
    def create_property_tab(self):
        """Create Property Information tab"""
        self.property_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.property_frame, text="Property Info")
        
        # Property Information Section
        info_frame = ttk.LabelFrame(self.property_frame, text="Property Information", padding="10")
        info_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Property fields
        fields = [
            ("Property Address:", "address"),
            ("Postal Code:", "postal_code"),
            ("Property Type:", "property_type"),
            ("Bedrooms:", "bedrooms"),
            ("Bathrooms:", "bathrooms"),
            ("Size (sqm):", "size_sqm"),
            ("Asking Price:", "asking_price"),
            ("On Market For (days):", "days_on_market"),
            ("Key Features:", "key_features"),
            ("Description:", "description")
        ]
        
        self.entry_widgets = {}
        
        for i, (label_text, field_name) in enumerate(fields):
            row_frame = ttk.Frame(info_frame)
            row_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(row_frame, text=label_text, width=20).pack(side=tk.LEFT, padx=(0, 10))
            
            if field_name in ["key_features", "description"]:
                widget = tk.Text(row_frame, height=3, width=60)
                widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            else:
                widget = ttk.Entry(row_frame, width=60)
                widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            self.entry_widgets[field_name] = widget
            
    def create_investment_tab(self):
        """Create Investment Analysis tab"""
        self.investment_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.investment_frame, text="Investment Analysis")
        
        # Investment Analysis Section
        investment_frame = ttk.LabelFrame(self.investment_frame, text="Investment Analysis", padding="10")
        investment_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Investment fields
        investment_fields = [
            ("Purchase Price:", "purchase_price"),
            ("Deposit (%):", "deposit_percent"),
            ("Estimated Monthly Rent:", "monthly_rent"),
            ("Mortgage Rate (%):", "mortgage_rate"),
            ("Council Tax (annual):", "council_tax"),
            ("Repairs/Maintenance (annual):", "repairs_maintenance"),
            ("Electric/Gas (annual):", "utilities"),
            ("Water (annual):", "water"),
            ("Broadband/TV (annual):", "broadband_tv"),
            ("Insurance (annual):", "insurance"),
            ("Stamp Duty:", "stamp_duty"),
            ("Survey Cost:", "survey_cost"),
            ("Legal Fees:", "legal_fees"),
            ("Loan Set-up:", "loan_setup")
        ]
        
        for i, (label_text, field_name) in enumerate(investment_fields):
            row_frame = ttk.Frame(investment_frame)
            row_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(row_frame, text=label_text, width=25).pack(side=tk.LEFT, padx=(0, 10))
            
            widget = ttk.Entry(row_frame, width=30)
            widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            self.entry_widgets[field_name] = widget
            
    def create_epc_tab(self):
        """Create EPC & Details tab"""
        self.epc_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.epc_frame, text="EPC & Details")
        
        # EPC Section
        epc_frame = ttk.LabelFrame(self.epc_frame, text="Energy Performance Certificate", padding="10")
        epc_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # EPC fields
        epc_fields = [
            ("EPC Grade:", "epc_grade"),
            ("Current Rating:", "current_rating"),
            ("Potential Rating:", "potential_rating"),
            ("Latest Inspection Date:", "inspection_date"),
            ("Window Glazing:", "window_glazing"),
            ("Building Construction Age:", "building_age"),
            ("Broadband Available:", "broadband_available"),
            ("Highest Download Speed:", "download_speed"),
            ("Highest Upload Speed:", "upload_speed")
        ]
        
        for i, (label_text, field_name) in enumerate(epc_fields):
            row_frame = ttk.Frame(epc_frame)
            row_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(row_frame, text=label_text, width=25).pack(side=tk.LEFT, padx=(0, 10))
            
            widget = ttk.Entry(row_frame, width=30)
            widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            self.entry_widgets[field_name] = widget
            
    def create_location_tab(self):
        """Create Location & Transport tab"""
        self.location_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.location_frame, text="Location & Transport")
        
        # Location Section
        location_frame = ttk.LabelFrame(self.location_frame, text="Location Information", padding="10")
        location_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Location fields
        location_fields = [
            ("City:", "city"),
            ("Population:", "population"),
            ("Distance to City Centre (miles):", "distance_city_centre"),
            ("Time to City Centre by Car (minutes):", "time_car"),
            ("Time to City Centre by Public Transport (minutes):", "time_public_transport"),
            ("Walk to Station (minutes):", "walk_to_station"),
            ("Station Distance (miles):", "station_distance"),
            ("Bus Routes:", "bus_routes"),
            ("Bus Frequency:", "bus_frequency"),
            ("About the City:", "about_city")
        ]
        
        for i, (label_text, field_name) in enumerate(location_fields):
            row_frame = ttk.Frame(location_frame)
            row_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(row_frame, text=label_text, width=30).pack(side=tk.LEFT, padx=(0, 10))
            
            if field_name == "about_city":
                widget = tk.Text(row_frame, height=4, width=50)
                widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            else:
                widget = ttk.Entry(row_frame, width=30)
                widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            self.entry_widgets[field_name] = widget
            
    def create_images_tab(self):
        """Create Images tab"""
        self.images_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.images_frame, text="Images")
        
        # Instructions Section
        instructions_frame = ttk.LabelFrame(self.images_frame, text="📋 Image Naming Guide", padding="10")
        instructions_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        instructions_text = """Images are automatically placed based on their filename. Include these keywords in your image filenames:

• Cover Page Main Image: Include "exterior" and "front" (e.g., "exterior_front.jpg")
• Floor Plans: Include "floorplan" or "plan" (e.g., "floorplan1.png")
• Directions Map: Include "directions", "map", or "city" (e.g., "directions.png")
• Liverpool Pictures: Include "liverpool" (e.g., "liverpool1.jpg")
• Property Images: Any other images will appear in the Property Images section

Tip: You can rename images before adding them, or use the filename as-is if it already contains these keywords."""
        
        instructions_label = tk.Label(instructions_frame, text=instructions_text, 
                                     justify=tk.LEFT, wraplength=800, font=("TkDefaultFont", 9))
        instructions_label.pack(anchor=tk.W)
        
        # Images Section
        images_frame = ttk.LabelFrame(self.images_frame, text="Property Images", padding="10")
        images_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Image listbox with scrollbar
        listbox_frame = ttk.Frame(images_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.image_listbox = tk.Listbox(listbox_frame, height=15)
        self.image_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.image_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.image_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Image buttons
        button_frame = ttk.Frame(images_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="Add Images", command=self.add_images).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Remove Selected", command=self.remove_image).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Move Up", command=self.move_image_up).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Move Down", command=self.move_image_down).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Load Sample Images", command=self.load_sample_images).pack(side=tk.LEFT, padx=(10, 0))
        
    def get_image_placement(self, filename):
        """Determine where an image will be placed based on its filename"""
        filename_lower = filename.lower()
        
        if 'exterior' in filename_lower and 'front' in filename_lower:
            return "📸 Cover Page (Main Image)"
        elif 'floor' in filename_lower or 'plan' in filename_lower:
            return "📐 Floor Plans"
        elif 'direction' in filename_lower or 'map' in filename_lower or 'city' in filename_lower:
            return "🗺️ Directions Map"
        elif 'liverpool' in filename_lower:
            return "🏙️ Liverpool Pictures"
        else:
            return "🏠 Property Images"
        
    def add_images(self):
        """Add images to the list"""
        file_paths = filedialog.askopenfilenames(
            title="Select Property Images",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.bmp *.tiff *.webp")]
        )
        
        for file_path in file_paths:
            self.images.append(file_path)
            filename = os.path.basename(file_path)
            placement = self.get_image_placement(filename)
            # Display filename with placement info
            display_text = f"{filename} [{placement}]"
            self.image_listbox.insert(tk.END, display_text)
            
    def remove_image(self):
        """Remove selected image from the list"""
        selection = self.image_listbox.curselection()
        if selection:
            index = selection[0]
            self.image_listbox.delete(index)
            del self.images[index]
            
    def move_image_up(self):
        """Move selected image up in the list"""
        selection = self.image_listbox.curselection()
        if selection and selection[0] > 0:
            index = selection[0]
            # Swap items
            self.images[index], self.images[index-1] = self.images[index-1], self.images[index]
            # Update listbox
            item1 = self.image_listbox.get(index)
            item2 = self.image_listbox.get(index-1)
            self.image_listbox.delete(index-1, index)
            self.image_listbox.insert(index-1, item1)
            self.image_listbox.insert(index, item2)
            self.image_listbox.selection_set(index-1)
            
    def move_image_down(self):
        """Move selected image down in the list"""
        selection = self.image_listbox.curselection()
        if selection and selection[0] < len(self.images) - 1:
            index = selection[0]
            # Swap items
            self.images[index], self.images[index+1] = self.images[index+1], self.images[index]
            # Update listbox
            item1 = self.image_listbox.get(index)
            item2 = self.image_listbox.get(index+1)
            self.image_listbox.delete(index, index+1)
            self.image_listbox.insert(index, item2)
            self.image_listbox.insert(index+1, item1)
            self.image_listbox.selection_set(index+1)
            
    def load_sample_images(self):
        """Load sample images from the sample_images folder"""
        sample_folder = get_resource_path("sample_images")
        if os.path.exists(sample_folder):
            self.load_images_from_folder(sample_folder)
        else:
            messagebox.showwarning("Warning", "Sample images folder not found. Please create sample images first.")
            
    def load_images_from_folder(self, folder_path):
        """Load all images from a specified folder"""
        image_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp')
        
        try:
            print(f"Loading images from folder: {folder_path}")
            files = os.listdir(folder_path)
            print(f"Found {len(files)} files in folder")
            
            for filename in files:
                if filename.lower().endswith(image_extensions):
                    image_path = os.path.join(folder_path, filename)
                    self.images.append(image_path)
                    placement = self.get_image_placement(filename)
                    display_text = f"{filename} [{placement}]"
                    self.image_listbox.insert(tk.END, display_text)
                    print(f"Loaded image: {filename}")
                    
            print(f"Total images loaded: {len(self.images)}")
        except Exception as e:
            print(f"Error loading images from folder: {e}")
            messagebox.showerror("Error", f"Error loading images from folder: {e}")
            
    def clear_all(self):
        """Clear all form data"""
        for widget in self.entry_widgets.values():
            if isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
            else:
                widget.delete(0, tk.END)
        
        self.images.clear()
        self.image_listbox.delete(0, tk.END)
        
    def load_sample_data(self):
        """Load sample data for testing"""
        # Sample property data
        sample_data = {
            'address': '5, Ridley Road',
            'postal_code': 'L6 6DN',
            'property_type': 'Semi-Detached House',
            'bedrooms': '5',
            'bathrooms': '5',
            'size_sqm': '116',
            'asking_price': '£290,000',
            'days_on_market': '6',
            'key_features': 'Spacious Three Storey HMO Property\nFive Spacious En-Suite Double Bedrooms\nFantastic Investment Opportunity\nContemporary Fitted Kitchen\nCommunal Lounge\nSunny Rear Courtyard\nYield of 10.31%\nClose To Great Local Amenities, Train Station And Road Links\nClose To City Centre\nEPC GRADE = C',
            'description': 'Beautiful semi-detached family home in excellent condition. Features include modern kitchen, spacious living areas, and a well-maintained garden. Perfect for families looking for comfort and convenience. Located in a quiet residential area with excellent transport links.',
            
            # Investment data
            'purchase_price': '£290,000',
            'deposit_percent': '20',
            'monthly_rent': '£2,750',
            'mortgage_rate': '5.8',
            'council_tax': '£1,670',
            'repairs_maintenance': '£660',
            'utilities': '£1,080',
            'water': '£300',
            'broadband_tv': '£480',
            'insurance': '£480',
            'stamp_duty': '£19,000',
            'survey_cost': '£800',
            'legal_fees': '£2,400',
            'loan_setup': '£4,640',
            
            # EPC data
            'epc_grade': 'C',
            'current_rating': '84',
            'potential_rating': '72',
            'inspection_date': '30th January 2019',
            'window_glazing': 'Double glazing installed during or after 2002',
            'building_age': 'before 1900',
            'broadband_available': 'Broadband available',
            'download_speed': '1,800 Mbps',
            'upload_speed': '220 Mbps',
            
            # Location data
            'city': 'Liverpool',
            'population': '508,986',
            'distance_city_centre': '1.8',
            'time_car': '6',
            'time_public_transport': '18',
            'walk_to_station': '11',
            'station_distance': '0.5',
            'bus_routes': '10A / 9',
            'bus_frequency': 'Every 8 minutes',
            'about_city': 'Liverpool is a port city and metropolitan borough in Merseyside, England. It is situated on the eastern side of the Mersey Estuary, near the Irish Sea, 178 miles (286 km) north-west of London. With a population of 496,770, Liverpool is the administrative, cultural and economic centre of the Liverpool City Region, a combined authority area with a population of over 1.5 million.'
        }
        
        # Fill in the form fields
        for field_name, value in sample_data.items():
            if field_name in self.entry_widgets:
                widget = self.entry_widgets[field_name]
                if isinstance(widget, tk.Text):
                    widget.delete("1.0", tk.END)
                    widget.insert("1.0", value)
                else:
                    widget.delete(0, tk.END)
                    widget.insert(0, value)
        
        # Load sample images
        self.load_sample_images()
        
    def generate_pdf(self):
        """Generate the comprehensive investment report PDF with professional styling"""
        try:
            # Get all form data
            data = {}
            for field_name, widget in self.entry_widgets.items():
                if isinstance(widget, tk.Text):
                    data[field_name] = widget.get("1.0", tk.END).strip()
                else:
                    data[field_name] = widget.get().strip()
            
            # Check if we have at least an address
            if not data.get('address'):
                messagebox.showerror("Error", "Please enter at least the property address.")
                return
            
            # Ask for output file location
            default_filename = f"{data.get('address', 'Property')} - Investment Report.pdf"
            file_path = filedialog.asksaveasfilename(
                title="Save Investment Report As",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialfile=default_filename
            )
            
            if not file_path:
                return
            
            print(f"Generating PDF: {file_path}")
            
            # Create PDF document with custom margins
            doc = SimpleDocTemplate(file_path, pagesize=A4, 
                                  leftMargin=0.75*inch, rightMargin=0.75*inch,
                                  topMargin=0.75*inch, bottomMargin=0.75*inch)
            story = []
            
            # Define professional color scheme
            primary_blue = HexColor('#1e3a8a')  # Dark blue
            accent_gold = HexColor('#f59e0b')   # Gold
            light_grey = HexColor('#f8fafc')   # Light grey
            dark_grey = HexColor('#374151')    # Dark grey
            success_green = HexColor('#10b981') # Green
            warning_orange = HexColor('#f59e0b') # Orange
            
            # Get styles
            styles = getSampleStyleSheet()
            
            # Custom styles
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=28,
                spaceAfter=20,
                alignment=TA_CENTER,
                textColor=primary_blue,
                fontName='Helvetica-Bold'
            )
            
            header_style = ParagraphStyle(
                'CustomHeader',
                parent=styles['Heading2'],
                fontSize=18,
                spaceAfter=15,
                spaceBefore=25,
                textColor=primary_blue,
                fontName='Helvetica-Bold'
            )
            
            subheader_style = ParagraphStyle(
                'CustomSubHeader',
                parent=styles['Heading3'],
                fontSize=14,
                spaceAfter=10,
                spaceBefore=15,
                textColor=dark_grey,
                fontName='Helvetica-Bold'
            )
            
            body_style = ParagraphStyle(
                'CustomBody',
                parent=styles['Normal'],
                fontSize=11,
                spaceAfter=8,
                textColor=dark_grey,
                fontName='Helvetica'
            )
            
            highlight_style = ParagraphStyle(
                'CustomHighlight',
                parent=styles['Normal'],
                fontSize=12,
                spaceAfter=8,
                textColor=primary_blue,
                fontName='Helvetica-Bold'
            )
            
            # Create cover page (first page)
            cover_page = self.create_cover_page(data, accent_gold, primary_blue)
            if cover_page:
                story.append(cover_page)
                story.append(PageBreak())
            
            # Investment Opportunity Section - Second Page Design
            # Logo on left with title next to it
            logo_path = get_resource_path("logo.png")
            if os.path.exists(logo_path):
                try:
                    # Calculate logo dimensions same as first page
                    logo_img_pil = Image.open(logo_path)
                    logo_width = 1.5*inch
                    logo_height = logo_width * (logo_img_pil.height / logo_img_pil.width)
                    logo_img = RLImage(logo_path, width=logo_width, height=logo_height)
                    
                    # Create header table with logo on left and title on right
                    title_para = Paragraph("Investment Opportunity", 
                        ParagraphStyle('InvestmentTitle', parent=styles['Heading1'], fontSize=24, 
                                      textColor=colors.black, fontName='Helvetica-Bold'))
                    
                    header_table = Table([[logo_img, title_para]], colWidths=[2*inch, 5*inch])
                    header_table.setStyle(TableStyle([
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                        ('ALIGN', (1, 0), (1, 0), 'LEFT'),
                        ('LEFTPADDING', (0, 0), (0, 0), 0),
                        ('RIGHTPADDING', (0, 0), (0, 0), 10),
                        ('LEFTPADDING', (1, 0), (1, 0), 10),
                        ('TOPPADDING', (0, 0), (-1, -1), 5),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                    ]))
                    story.append(header_table)
                except Exception as e:
                    print(f"Error loading logo: {e}")
                    # Fallback to just title
                    investment_title_style = ParagraphStyle(
                        'InvestmentTitle',
                        parent=styles['Heading1'],
                        fontSize=24,
                        spaceAfter=25,
                        textColor=colors.black,
                        fontName='Helvetica-Bold'
                    )
                    story.append(Paragraph("Investment Opportunity", investment_title_style))
            else:
                # Just title if no logo
                investment_title_style = ParagraphStyle(
                    'InvestmentTitle',
                    parent=styles['Heading1'],
                    fontSize=24,
                    spaceAfter=25,
                    textColor=colors.black,
                    fontName='Helvetica-Bold'
                )
                story.append(Paragraph("Investment Opportunity", investment_title_style))
            
            story.append(Spacer(1, 20))
            
            # Calculate investment metrics
            try:
                purchase_price = float(data.get('purchase_price', '0').replace('£', '').replace(',', ''))
                deposit_percent = float(data.get('deposit_percent', '20'))
                monthly_rent = float(data.get('monthly_rent', '0').replace('£', '').replace(',', ''))
                mortgage_rate = float(data.get('mortgage_rate', '5.8'))
                
                deposit_amount = purchase_price * (deposit_percent / 100)
                annual_rent = monthly_rent * 12
                rental_yield = (annual_rent / purchase_price) * 100
                
                # Calculate costs
                stamp_duty = float(data.get('stamp_duty', '0').replace('£', '').replace(',', ''))
                survey_cost = float(data.get('survey_cost', '0').replace('£', '').replace(',', ''))
                legal_fees = float(data.get('legal_fees', '0').replace('£', '').replace(',', ''))
                loan_setup = float(data.get('loan_setup', '0').replace('£', '').replace(',', ''))
                
                total_purchase_costs = stamp_duty + survey_cost + legal_fees + loan_setup
                total_investment = deposit_amount + total_purchase_costs
                
                # Calculate expenses
                mortgage_amount = purchase_price - deposit_amount
                annual_mortgage_interest = mortgage_amount * (mortgage_rate / 100)
                
                council_tax = float(data.get('council_tax', '0').replace('£', '').replace(',', ''))
                repairs = float(data.get('repairs_maintenance', '0').replace('£', '').replace(',', ''))
                utilities = float(data.get('utilities', '0').replace('£', '').replace(',', ''))
                water = float(data.get('water', '0').replace('£', '').replace(',', ''))
                broadband = float(data.get('broadband_tv', '0').replace('£', '').replace(',', ''))
                insurance = float(data.get('insurance', '0').replace('£', '').replace(',', ''))
                
                total_annual_expenses = annual_mortgage_interest + council_tax + repairs + utilities + water + broadband + insurance
                annual_profit = annual_rent - total_annual_expenses
                monthly_profit = annual_profit / 12
                roi = (annual_profit / total_investment) * 100
                
            except (ValueError, ZeroDivisionError):
                # Default values if calculation fails
                purchase_price = 0
                deposit_amount = 0
                monthly_rent = 0
                rental_yield = 0
                total_investment = 0
                total_annual_expenses = 0
                annual_profit = 0
                monthly_profit = 0
                roi = 0
            
            # Three Key Metrics Boxes (horizontal gold boxes) - individual boxes, bigger
            box_width = 2.4*inch
            box_spacing = 0.15*inch  # Space between boxes
            
            # Create three individual boxes side by side with spacing
            metrics_table = Table([
                ['Purchase Price', '', 'Estimated Monthly Rent', '', 'Rental Yield'],
                [f"£{purchase_price:,.0f}", '', f"£{monthly_rent:,.0f}pcm", '', f"{rental_yield:.1f}%"]
            ], colWidths=[box_width, box_spacing, box_width, box_spacing, box_width])
            metrics_table.setStyle(TableStyle([
                # First box (Purchase Price)
                ('BACKGROUND', (0, 0), (0, 1), accent_gold),
                ('TEXTCOLOR', (0, 0), (0, 0), colors.black),
                ('TEXTCOLOR', (0, 1), (0, 1), colors.white),
                # Second box (Monthly Rent)
                ('BACKGROUND', (2, 0), (2, 1), accent_gold),
                ('TEXTCOLOR', (2, 0), (2, 0), colors.black),
                ('TEXTCOLOR', (2, 1), (2, 1), colors.white),
                # Third box (Rental Yield)
                ('BACKGROUND', (4, 0), (4, 1), accent_gold),
                ('TEXTCOLOR', (4, 0), (4, 0), colors.black),
                ('TEXTCOLOR', (4, 1), (4, 1), colors.white),
                # Spacing columns (white background)
                ('BACKGROUND', (1, 0), (1, 1), colors.white),
                ('BACKGROUND', (3, 0), (3, 1), colors.white),
                # Alignment
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                # Fonts
                ('FONTNAME', (0, 0), (4, 0), 'Helvetica'),
                ('FONTSIZE', (0, 0), (4, 0), 12),
                ('FONTNAME', (0, 1), (4, 1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 1), (4, 1), 24),
                # Padding - bigger boxes
                ('TOPPADDING', (0, 0), (4, 0), 15),
                ('BOTTOMPADDING', (0, 0), (4, 0), 5),
                ('TOPPADDING', (0, 1), (4, 1), 5),
                ('BOTTOMPADDING', (0, 1), (4, 1), 15),
                ('LEFTPADDING', (0, 0), (4, -1), 12),
                ('RIGHTPADDING', (0, 0), (4, -1), 12),
            ]))
            story.append(metrics_table)
            story.append(Spacer(1, 30))
            
            # Two Column Layout for Costs and Expenses
            # Left Column: Total Purchase Costs
            purchase_costs_data = [
                ['Total Purchase Costs', ''],
                ['Deposit(20%)', f"£{deposit_amount:,.0f}"],
                ['Stamp Duty', f"£{stamp_duty:,.0f}"],
                ['Survey', f"£{survey_cost:,.0f}"],
                ['Legal Fees', f"£{legal_fees:,.0f}"],
                ['Loan Set-up', f"£{loan_setup:,.0f}"],
                ['Total Investment Required', f"£{total_investment:,.0f}"]
            ]
            
            # Right Column: Total Annual Expenses
            expenses_data = [
                ['Total Annual Expenses', ''],
                [f'Mortgage @ {mortgage_rate}% (Interest Only)', f"£{annual_mortgage_interest:,.0f}"],
                ['Council Tax', f"£{council_tax:,.0f}"],
                ['Repairs / Maintenance', f"£{repairs:,.0f}"],
                ['Electric / Gas', f"£{utilities:,.0f}"],
                ['Water', f"£{water:,.0f}"],
                ['Broadband / TV', f"£{broadband:,.0f}"],
                ['Insurance', f"£{insurance:,.0f}"],
                ['Total', f"£{total_annual_expenses:,.0f}"]
            ]
            
            # Create two-column table
            two_col_data = []
            max_rows = max(len(purchase_costs_data), len(expenses_data))
            for i in range(max_rows):
                left_col = purchase_costs_data[i] if i < len(purchase_costs_data) else ['', '']
                right_col = expenses_data[i] if i < len(expenses_data) else ['', '']
                two_col_data.append([left_col[0], left_col[1], right_col[0], right_col[1]])
            
            two_col_table = Table(two_col_data, colWidths=[2.2*inch, 1.3*inch, 2.2*inch, 1.3*inch])
            
            # Add horizontal lines between all rows
            table_style_commands = [
                # Header rows
                ('BACKGROUND', (0, 0), (1, 0), colors.white),
                ('BACKGROUND', (2, 0), (3, 0), colors.white),
                ('FONTNAME', (0, 0), (1, 0), 'Helvetica-Bold'),
                ('FONTNAME', (2, 0), (3, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('ALIGN', (3, 0), (3, -1), 'RIGHT'),
                # Data rows
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 11),
                # Total row styling (bold)
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('LEFTPADDING', (0, 0), (0, -1), 5),
                ('RIGHTPADDING', (1, 0), (1, -1), 5),
                ('LEFTPADDING', (2, 0), (2, -1), 15),
                ('RIGHTPADDING', (3, 0), (3, -1), 5),
            ]
            
            # Add horizontal lines between all rows (light grey lines)
            for row_idx in range(1, len(two_col_data)):
                # Left column line - spans full width of left column
                table_style_commands.append(('LINEBELOW', (0, row_idx), (1, row_idx), 0.5, HexColor('#E0E0E0')))
                # Right column line - spans full width of right column
                table_style_commands.append(('LINEBELOW', (2, row_idx), (3, row_idx), 0.5, HexColor('#E0E0E0')))
            
            two_col_table.setStyle(TableStyle(table_style_commands))
            
            # Three Vertical Boxes for Profit/ROI (positioned at bottom right, bigger)
            # Create table with left column empty and boxes on the right
            profit_table = Table([
                ['', 'Monthly Profit', f"£{monthly_profit:,.0f}"],
                ['', 'Annual Profit', f"£{annual_profit:,.0f}"],
                ['', 'ROI', f"{roi:.1f}%"]
            ], colWidths=[3.5*inch, 2*inch, 2*inch])
            profit_table.setStyle(TableStyle([
                ('BACKGROUND', (1, 0), (1, -1), accent_gold),
                ('BACKGROUND', (2, 0), (2, -1), accent_gold),
                ('TEXTCOLOR', (1, 0), (1, -1), colors.black),  # Labels in black
                ('TEXTCOLOR', (2, 0), (2, -1), colors.white),  # Values in white bold
                ('ALIGN', (1, 0), (1, -1), 'LEFT'),
                ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
                ('VALIGN', (1, 0), (2, -1), 'MIDDLE'),
                ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
                ('FONTSIZE', (1, 0), (1, -1), 13),  # Bigger labels
                ('FONTNAME', (2, 0), (2, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (2, 0), (2, -1), 24),  # Bigger values
                ('TOPPADDING', (1, 0), (2, -1), 16),  # Bigger boxes
                ('BOTTOMPADDING', (1, 0), (2, -1), 16),
                ('LEFTPADDING', (1, 0), (1, -1), 18),
                ('RIGHTPADDING', (1, 0), (1, -1), 12),
                ('LEFTPADDING', (2, 0), (2, -1), 12),
                ('RIGHTPADDING', (2, 0), (2, -1), 18),
            ]))
            
            # Use KeepTogether to ensure profit boxes stay on same page as costs table
            # Include spacer to push profit boxes toward bottom
            story.append(KeepTogether([two_col_table, Spacer(1, 2*inch), profit_table]))
            
            # Page break before Key Information section
            story.append(PageBreak())
            
            # Key Information Section - Third Page Design (all content on one page)
            # Collect all content first, then wrap in KeepTogether
            key_info_content = []
            
            # Logo on left with "Key Information" title
            logo_path = get_resource_path("logo.png")
            if os.path.exists(logo_path):
                try:
                    # Calculate logo dimensions same as first page
                    logo_img_pil = Image.open(logo_path)
                    logo_width = 1.5*inch
                    logo_height = logo_width * (logo_img_pil.height / logo_img_pil.width)
                    logo_img = RLImage(logo_path, width=logo_width, height=logo_height)
                    
                    # Create header table with logo on left and title on right
                    key_info_title_para = Paragraph("Key Information", 
                        ParagraphStyle('KeyInfoTitle', parent=styles['Heading1'], fontSize=24, 
                                      textColor=colors.black, fontName='Helvetica-Bold'))
                    
                    header_table = Table([[logo_img, key_info_title_para]], colWidths=[2*inch, 5*inch])
                    header_table.setStyle(TableStyle([
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                        ('ALIGN', (1, 0), (1, 0), 'LEFT'),
                        ('LEFTPADDING', (0, 0), (0, 0), 0),
                        ('RIGHTPADDING', (0, 0), (0, 0), 10),
                        ('LEFTPADDING', (1, 0), (1, 0), 10),
                        ('TOPPADDING', (0, 0), (-1, -1), 5),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                    ]))
                    key_info_content.append(header_table)
                except Exception as e:
                    print(f"Error loading logo: {e}")
                    key_info_content.append(Paragraph("Key Information", header_style))
            else:
                key_info_content.append(Paragraph("Key Information", header_style))
            
            key_info_content.append(Spacer(1, 15))
            
            # Property image (reduced size to fit on page) - always show, use placeholder if missing
            img_width = 6.5*inch
            img_height = 3.5*inch
            main_img_path = None
            
            if self.images:
                try:
                    # Find exterior front image or use first image
                    for img_path in self.images:
                        filename = os.path.basename(img_path).lower()
                        if 'exterior' in filename and 'front' in filename:
                            main_img_path = img_path
                            break
                    if not main_img_path:
                        main_img_path = self.images[0]
                    
                    if main_img_path and os.path.exists(main_img_path):
                        # Display image (reduced size to fit on one page)
                        property_img_pil = Image.open(main_img_path)
                        img_aspect = property_img_pil.width / property_img_pil.height
                        img_width = 6.5*inch  # Slightly smaller to fit
                        img_height = img_width / img_aspect
                        # Limit height to ensure it fits
                        if img_height > 3.5*inch:
                            img_height = 3.5*inch
                            img_width = img_height * img_aspect
                        property_img = RLImage(main_img_path, width=img_width, height=img_height)
                        key_info_content.append(property_img)
                    else:
                        # Use placeholder
                        placeholder = create_placeholder_drawing(img_width, img_height)
                        key_info_content.append(placeholder)
                    key_info_content.append(Spacer(1, 15))
                except Exception as e:
                    print(f"Error loading property image: {e}")
                    # Use placeholder on error
                    placeholder = create_placeholder_drawing(img_width, img_height)
                    key_info_content.append(placeholder)
                    key_info_content.append(Spacer(1, 15))
            else:
                # No images - use placeholder
                placeholder = create_placeholder_drawing(img_width, img_height)
                key_info_content.append(placeholder)
                key_info_content.append(Spacer(1, 15))
            
            # Property Metrics - Four metrics displayed horizontally (label above value)
            asking_price = data.get('asking_price', 'N/A')
            if asking_price.startswith('£'):
                asking_price_value = asking_price
            else:
                asking_price_value = f"£{asking_price.replace('£', '').strip()}"
            
            metrics_data = [
                ['Asking price', 'Bedrooms', 'Size', 'On the market for'],
                [asking_price_value, data.get('bedrooms', 'N/A'), 
                 f"{data.get('size_sqm', 'N/A')} sqm", 
                 f"{data.get('days_on_market', 'N/A')} days"]
            ]
            
            metrics_table = Table(metrics_data, colWidths=[1.8*inch, 1.8*inch, 1.8*inch, 1.8*inch])
            metrics_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, 0), 11),
                ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 1), (-1, 1), 16),
                ('TOPPADDING', (0, 0), (-1, 0), 8),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
                ('TOPPADDING', (0, 1), (-1, 1), 4),
                ('BOTTOMPADDING', (0, 1), (-1, 1), 12),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ]))
            key_info_content.append(metrics_table)
            key_info_content.append(Spacer(1, 20))
            
            # Key Features - Bulleted list
            if data.get('key_features'):
                # Split key features by newline and create bulleted list
                features_text = data.get('key_features', '')
                features_list = [f.strip() for f in features_text.split('\n') if f.strip()]
                
                key_info_content.append(Paragraph("Key Features", header_style))
                key_info_content.append(Spacer(1, 8))
                
                # Create bulleted list with reduced spacing to fit on page
                for feature in features_list:
                    bullet_para = Paragraph(f"• {feature}", 
                        ParagraphStyle('KeyFeature', parent=styles['Normal'], 
                                      fontSize=11, textColor=colors.black,
                                      leftIndent=20, spaceAfter=5))  # Further reduced spacing
                    key_info_content.append(bullet_para)
            
            # Use KeepTogether to ensure entire Key Information section stays on same page
            story.append(KeepTogether(key_info_content))
            story.append(Spacer(1, 20))
            
            # Page break after Key Information page
            story.append(PageBreak())
            
            # Other Key Information Page Header
            logo_path = get_resource_path("logo.png")
            if os.path.exists(logo_path):
                try:
                    logo_img_pil = Image.open(logo_path)
                    logo_width = 1.5*inch
                    logo_height = logo_width * (logo_img_pil.height / logo_img_pil.width)
                    logo_img = RLImage(logo_path, width=logo_width, height=logo_height)
                    
                    # Create header table with logo on left and title centered
                    other_key_title_para = Paragraph("Other Key Information", 
                        ParagraphStyle('OtherKeyTitle', parent=styles['Heading1'], fontSize=24, 
                                      textColor=colors.black, fontName='Helvetica-Bold',
                                      alignment=1))  # 1 = TA_CENTER
                    
                    header_table = Table([[logo_img, other_key_title_para]], colWidths=[2*inch, 5*inch])
                    header_table.setStyle(TableStyle([
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                        ('LEFTPADDING', (0, 0), (0, 0), 0),
                        ('RIGHTPADDING', (0, 0), (0, 0), 10),
                        ('LEFTPADDING', (1, 0), (1, 0), 10),
                        ('TOPPADDING', (0, 0), (-1, -1), 5),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                    ]))
                    story.append(header_table)
                except Exception as e:
                    print(f"Error loading logo: {e}")
                    story.append(Paragraph("Other Key Information", header_style))
            else:
                story.append(Paragraph("Other Key Information", header_style))
            
            story.append(Spacer(1, 20))
            
            # Large Property Image below header - always show, use placeholder if missing
            img_width = 7*inch
            img_height = 3.5*inch
            main_img_path = None
            
            if self.images:
                try:
                    # Find exterior front image or use first image
                    for img_path in self.images:
                        filename = os.path.basename(img_path).lower()
                        if 'exterior' in filename and 'front' in filename:
                            main_img_path = img_path
                            break
                    if not main_img_path:
                        main_img_path = self.images[0]
                    
                    if main_img_path and os.path.exists(main_img_path):
                        # Display large image
                        property_img_pil = Image.open(main_img_path)
                        img_aspect = property_img_pil.width / property_img_pil.height
                        img_width = 7*inch  # Full width
                        img_height = img_width / img_aspect
                        # Limit height to ensure it fits nicely
                        if img_height > 3.5*inch:
                            img_height = 3.5*inch
                            img_width = img_height * img_aspect
                        property_img = RLImage(main_img_path, width=img_width, height=img_height)
                        story.append(property_img)
                    else:
                        # Use placeholder
                        placeholder = create_placeholder_drawing(img_width, img_height)
                        story.append(placeholder)
                    story.append(Spacer(1, 20))
                except Exception as e:
                    print(f"Error loading property image: {e}")
                    # Use placeholder on error
                    placeholder = create_placeholder_drawing(img_width, img_height)
                    story.append(placeholder)
                    story.append(Spacer(1, 20))
            else:
                # No images - use placeholder
                placeholder = create_placeholder_drawing(img_width, img_height)
                story.append(placeholder)
                story.append(Spacer(1, 20))
            
            # EPC Section - Horizontal Layout: Title (left) | Chart (middle) | Details (right)
            epc_title_para = Paragraph("Energy Performance Certificate", 
                ParagraphStyle('EPCTitle', parent=styles['Heading2'], fontSize=14, 
                              textColor=colors.black, fontName='Helvetica-Bold'))
            
            # Create EPC visual chart (vertical bars)
            epc_chart = self.create_epc_chart(data, primary_blue, accent_gold, success_green)
            
            # EPC Details (right side)
            epc_details_para = Paragraph(
                f"<b>Latest available inspection date</b><br/>"
                f"{data.get('inspection_date', 'N/A')}<br/><br/>"
                f"<b>Window glazing</b><br/>"
                f"{data.get('window_glazing', 'N/A')}<br/><br/>"
                f"<b>Building construction age band</b><br/>"
                f"{data.get('building_age', 'N/A')}",
                ParagraphStyle('EPCDetails', parent=styles['Normal'], 
                              fontSize=11, textColor=colors.black,
                              leftIndent=0))
            
            # Create horizontal table for EPC section
            epc_table = Table([[epc_title_para, epc_chart, epc_details_para]], 
                             colWidths=[2*inch, 3.5*inch, 2.5*inch])
            epc_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                ('ALIGN', (2, 0), (2, 0), 'LEFT'),
                ('TOPPADDING', (0, 0), (0, 0), 20),  # More padding for title column
                ('TOPPADDING', (1, 0), (2, 0), 10),  # Regular padding for other columns
                ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ]))
            story.append(epc_table)
            story.append(Spacer(1, 10))
            
            # Disclaimer text in grey
            disclaimer_para = Paragraph(
                "This EPC data is accurate up to 6 months ago. If a more recent EPC assessment was done within this period, it will not be displayed here.",
                ParagraphStyle('EPCDisclaimer', parent=styles['Normal'], 
                              fontSize=9, textColor=HexColor('#666666')))
            story.append(disclaimer_para)
            story.append(Spacer(1, 20))
            
            # Internet / Broadband Availability - Horizontal Layout
            if data.get('broadband_available'):
                broadband_title_para = Paragraph("Internet / Broadband Availability", 
                    ParagraphStyle('BroadbandTitle', parent=styles['Heading2'], fontSize=14, 
                                  textColor=colors.black, fontName='Helvetica-Bold'))
                
                # Create three horizontal items (values are bold, not labels)
                broadband_item1 = Paragraph(
                    f"Broadband available<br/><b>{data.get('broadband_available', 'N/A')}</b>",
                    ParagraphStyle('BroadbandItem', parent=styles['Normal'], 
                                  fontSize=11, textColor=colors.black))
                
                broadband_item2 = Paragraph(
                    f"Highest available download speed<br/><b>{data.get('download_speed', 'N/A')}</b>",
                    ParagraphStyle('BroadbandItem', parent=styles['Normal'], 
                                  fontSize=11, textColor=colors.black))
                
                broadband_item3 = Paragraph(
                    f"Highest available upload speed<br/><b>{data.get('upload_speed', 'N/A')}</b>",
                    ParagraphStyle('BroadbandItem', parent=styles['Normal'], 
                                  fontSize=11, textColor=colors.black))
                
                # Create horizontal table for broadband info (title row, then three items)
                broadband_table = Table([
                    [broadband_title_para, '', ''],
                    [broadband_item1, broadband_item2, broadband_item3]
                ], colWidths=[2.3*inch, 2.3*inch, 2.4*inch])
                broadband_table.setStyle(TableStyle([
                    ('SPAN', (0, 0), (-1, 0)),  # Title spans all columns
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                    ('ALIGN', (0, 1), (-1, 1), 'LEFT'),
                    ('TOPPADDING', (0, 0), (-1, 0), 5),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                    ('TOPPADDING', (0, 1), (-1, 1), 5),
                    ('BOTTOMPADDING', (0, 1), (-1, 1), 5),
                ]))
                story.append(broadband_table)
                story.append(Spacer(1, 20))
            
            # Floor Plans Section - always show at least one placeholder if no floor plans
            floor_plan_images = []
            if self.images:
                # Filter floor plan images
                for img_path in self.images:
                    filename = os.path.basename(img_path).lower()
                    if 'floor' in filename or 'plan' in filename:
                        floor_plan_images.append(img_path)
            
            # Always show at least one floor plan (placeholder if none available)
            if not floor_plan_images:
                floor_plan_images = [None]  # Use None as placeholder marker
            
            if floor_plan_images:
                # Add Floor Plans header with logo on first page only
                story.append(PageBreak())
                
                # Logo on left with "Floor Plans" title
                logo_path = get_resource_path("logo.png")
                if os.path.exists(logo_path):
                    try:
                        # Calculate logo dimensions same as first page
                        logo_img_pil = Image.open(logo_path)
                        logo_width = 1.5*inch
                        logo_height = logo_width * (logo_img_pil.height / logo_img_pil.width)
                        logo_img = RLImage(logo_path, width=logo_width, height=logo_height)
                        
                        # Create header table with logo on left and title on right
                        floor_plans_title_para = Paragraph("Floor Plans", 
                            ParagraphStyle('FloorPlansTitle', parent=styles['Heading1'], fontSize=24, 
                                          textColor=colors.black, fontName='Helvetica-Bold'))
                        
                        header_table = Table([[logo_img, floor_plans_title_para]], colWidths=[2*inch, 5*inch])
                        header_table.setStyle(TableStyle([
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                            ('ALIGN', (1, 0), (1, 0), 'LEFT'),
                            ('LEFTPADDING', (0, 0), (0, 0), 0),
                            ('RIGHTPADDING', (0, 0), (0, 0), 10),
                            ('LEFTPADDING', (1, 0), (1, 0), 10),
                            ('TOPPADDING', (0, 0), (-1, -1), 5),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                        ]))
                        story.append(header_table)
                        story.append(Spacer(1, 20))
                    except Exception as e:
                        print(f"Error loading logo: {e}")
                        story.append(Paragraph("Floor Plans", header_style))
                        story.append(Spacer(1, 20))
                else:
                    story.append(Paragraph("Floor Plans", header_style))
                    story.append(Spacer(1, 20))
                
                for i, image_path in enumerate(floor_plan_images):
                    # For subsequent floor plans, add page break (but not for the first one, since we already added it)
                    if i > 0:
                        story.append(PageBreak())
                    
                    if image_path is None:
                        # Use placeholder
                        img_width = 6.5*inch
                        img_height = 4.5*inch
                        placeholder = create_placeholder_drawing(img_width, img_height)
                        
                        # Center vertically
                        available_height = 10*inch
                        remaining_height = available_height - img_height
                        top_spacer = remaining_height / 2
                        if top_spacer > 0:
                            story.append(Spacer(1, top_spacer))
                        
                        # Center horizontally
                        center_table = Table([[placeholder]], colWidths=[7.5*inch])
                        center_table.setStyle(TableStyle([
                            ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                            ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                        ]))
                        story.append(center_table)
                    else:
                        try:
                            # Load image to get dimensions
                            if os.path.exists(image_path):
                                img_pil = Image.open(image_path)
                                img_aspect = img_pil.width / img_pil.height
                                
                                # Display floor plan images - larger size for better visibility
                                img_width = 6.5*inch
                                img_height = img_width / img_aspect
                                
                                # Limit height to ensure it fits on page
                                if img_height > 4.5*inch:
                                    img_height = 4.5*inch
                                    img_width = img_height * img_aspect
                                
                                img = RLImage(image_path, width=img_width, height=img_height)
                            else:
                                # Use placeholder if file doesn't exist
                                img_width = 6.5*inch
                                img_height = 4.5*inch
                                img = create_placeholder_drawing(img_width, img_height)
                            
                            # Center the image on the page
                            available_height = 10*inch
                            remaining_height = available_height - img_height
                            top_spacer = remaining_height / 2
                            
                            # Add spacer to center vertically (before image)
                            if top_spacer > 0:
                                story.append(Spacer(1, top_spacer))
                            
                            # Center horizontally using Table
                            center_table = Table([[img]], colWidths=[7.5*inch])
                            center_table.setStyle(TableStyle([
                                ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                                ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                            ]))
                            story.append(center_table)
                            
                        except Exception as e:
                            print(f"Error adding floor plan image {image_path}: {e}")
                            # Use placeholder on error
                            img_width = 6.5*inch
                            img_height = 4.5*inch
                            placeholder = create_placeholder_drawing(img_width, img_height)
                            available_height = 10*inch
                            remaining_height = available_height - img_height
                            top_spacer = remaining_height / 2
                            if top_spacer > 0:
                                story.append(Spacer(1, top_spacer))
                            center_table = Table([[placeholder]], colWidths=[7.5*inch])
                            center_table.setStyle(TableStyle([
                                ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                                ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                            ]))
                            story.append(center_table)
            
            # Property Images (excluding floor plans, directions, and Liverpool images)
            # Always show at least one placeholder if no images
            regular_images = []
            if self.images:
                # Filter out floor plan images, directions image, and Liverpool images (they're already shown in their sections)
                for img_path in self.images:
                    filename = os.path.basename(img_path).lower()
                    if ('floor' not in filename and 'plan' not in filename and 
                        'direction' not in filename and 'map' not in filename and 'city' not in filename and
                        'liverpool' not in filename):
                        regular_images.append(img_path)
            
            # Always show property images section (with placeholder if empty)
            if not regular_images:
                regular_images = [None]  # Use None as placeholder marker
            
            if regular_images:
                story.append(PageBreak())
                story.append(Paragraph("Property Images", header_style))
                
                for i, image_path in enumerate(regular_images):
                    if image_path is None:
                        # Use placeholder
                        placeholder = create_placeholder_drawing(6*inch, 4*inch)
                        story.append(placeholder)
                        story.append(Spacer(1, 12))
                        story.append(Paragraph(f"Image {i+1}: Placeholder", body_style))
                        story.append(Spacer(1, 20))
                    else:
                        try:
                            # Add image with caption
                            if os.path.exists(image_path):
                                img = RLImage(image_path, width=6*inch, height=4*inch)
                            else:
                                # Use placeholder if file doesn't exist
                                img = create_placeholder_drawing(6*inch, 4*inch)
                            story.append(img)
                            story.append(Spacer(1, 12))
                            
                            # Add caption
                            filename = os.path.basename(image_path)
                            story.append(Paragraph(f"Image {i+1}: {filename}", body_style))
                            story.append(Spacer(1, 20))
                            
                        except Exception as e:
                            print(f"Error adding image {image_path}: {e}")
                            # Use placeholder on error
                            placeholder = create_placeholder_drawing(6*inch, 4*inch)
                            story.append(placeholder)
                            story.append(Spacer(1, 12))
                            story.append(Paragraph(f"Image {i+1}: Error loading image", body_style))
                            story.append(Spacer(1, 20))
            
            # Getting To The City Centre and About the City (Last Page)
            story.append(PageBreak())
            
            # Location Information - Logo, Title, and Directions Image
            logo_path = get_resource_path("logo.png")
            if os.path.exists(logo_path):
                try:
                    logo_img_pil = Image.open(logo_path)
                    logo_width = 1.5*inch
                    logo_height = logo_width * (logo_img_pil.height / logo_img_pil.width)
                    logo_img = RLImage(logo_path, width=logo_width, height=logo_height)
                    
                    # Create header table with logo on left and title centered
                    city_centre_title_para = Paragraph("Getting To The City Centre", 
                        ParagraphStyle('CityCentreTitle', parent=styles['Heading1'], fontSize=24, 
                                      textColor=colors.black, fontName='Helvetica-Bold',
                                      alignment=1))  # 1 = TA_CENTER
                    
                    header_table = Table([[logo_img, city_centre_title_para]], colWidths=[2*inch, 5*inch])
                    header_table.setStyle(TableStyle([
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                        ('LEFTPADDING', (0, 0), (0, 0), 0),
                        ('RIGHTPADDING', (0, 0), (0, 0), 10),
                        ('LEFTPADDING', (1, 0), (1, 0), 10),
                        ('TOPPADDING', (0, 0), (-1, -1), 5),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                    ]))
                    story.append(header_table)
                except Exception as e:
                    print(f"Error loading logo: {e}")
                    story.append(Paragraph("Getting To The City Centre", header_style))
            else:
                story.append(Paragraph("Getting To The City Centre", header_style))
            
            story.append(Spacer(1, 20))
            
            # Directions Image
            directions_image_path = None
            # Check if directions image exists in images list
            for img_path in self.images:
                filename = os.path.basename(img_path).lower()
                if 'direction' in filename or 'map' in filename or 'city' in filename:
                    directions_image_path = img_path
                    break
            
            # If not found in images list, check sample_images folder
            if not directions_image_path:
                sample_paths = [
                    get_resource_path("sample_images/directions.png"),
                    get_resource_path("directions.png"),
                    get_resource_path("sample_images/directions.jpg"),
                    get_resource_path("directions.jpg")
                ]
                for path in sample_paths:
                    if os.path.exists(path):
                        directions_image_path = path
                        break
            
            # Always show directions image (placeholder if not found)
            if directions_image_path and os.path.exists(directions_image_path):
                try:
                    # Load image to get dimensions
                    img_pil = Image.open(directions_image_path)
                    img_aspect = img_pil.width / img_pil.height
                    
                    # Display directions image - full width
                    img_width = 7*inch
                    img_height = img_width / img_aspect
                    
                    # Limit height to ensure it fits on page
                    if img_height > 8*inch:
                        img_height = 8*inch
                        img_width = img_height * img_aspect
                    
                    directions_img = RLImage(directions_image_path, width=img_width, height=img_height)
                    story.append(directions_img)
                except Exception as e:
                    print(f"Error loading directions image: {e}")
                    # Use placeholder on error
                    placeholder = create_placeholder_drawing(7*inch, 4.5*inch)
                    story.append(placeholder)
            else:
                # Use placeholder if not found
                placeholder = create_placeholder_drawing(7*inch, 4.5*inch)
                story.append(placeholder)
            
            story.append(Spacer(1, 15))
            
            # About the City and Liverpool Pictures - keep together on same page
            about_city_content = []
            
            # About the City
            if data.get('about_city'):
                about_city_content.append(Paragraph("About the City", header_style))
                about_city_content.append(Paragraph(f"<b>{data.get('city', 'N/A')}</b>", highlight_style))
                about_city_content.append(Paragraph(data.get('about_city'), body_style))
                about_city_content.append(Paragraph(f"<b>Population:</b> {data.get('population', 'N/A')}", body_style))
                about_city_content.append(Spacer(1, 10))
            
            # Liverpool Pictures Section
            liverpool_images = []
            # Check if Liverpool images exist in images list
            for img_path in self.images:
                filename = os.path.basename(img_path).lower()
                if 'liverpool' in filename:
                    liverpool_images.append(img_path)
            
            # If not found in images list, check sample_images folder
            if not liverpool_images:
                sample_paths = [
                    get_resource_path("sample_images/liverpool1.jpg"),
                    get_resource_path("sample_images/liverpool2.jpg"),
                    get_resource_path("sample_images/liverpool3.jpg")
                ]
                for path in sample_paths:
                    if os.path.exists(path):
                        liverpool_images.append(path)
            
            # Always show Liverpool images (placeholders if not found)
            fixed_width = 2.3*inch
            fixed_height = 2.3*inch
            
            # Always show 3 Liverpool images (placeholders if missing)
            while len(liverpool_images) < 3:
                liverpool_images.append(None)
            
            # Display 3 images horizontally
            img_cells = []
            for img_path in liverpool_images[:3]:
                if img_path is None:
                    # Use placeholder
                    placeholder = create_placeholder_drawing(fixed_width, fixed_height)
                    img_cells.append(placeholder)
                else:
                    try:
                        if os.path.exists(img_path):
                            # Use fixed dimensions to make them all the same size (may crop/distort to fit)
                            img = RLImage(img_path, width=fixed_width, height=fixed_height)
                            img_cells.append(img)
                        else:
                            # Use placeholder if file doesn't exist
                            placeholder = create_placeholder_drawing(fixed_width, fixed_height)
                            img_cells.append(placeholder)
                    except Exception as e:
                        print(f"Error loading Liverpool image {img_path}: {e}")
                        # Use placeholder on error
                        placeholder = create_placeholder_drawing(fixed_width, fixed_height)
                        img_cells.append(placeholder)
            
            if len(img_cells) == 3:
                liverpool_table = Table([img_cells], colWidths=[fixed_width, fixed_width, fixed_width])
                liverpool_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('TOPPADDING', (0, 0), (-1, -1), 5),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ]))
                about_city_content.append(liverpool_table)
            
            about_city_content.append(Spacer(1, 15))
            
            # Wrap in KeepTogether to ensure they stay on same page
            story.append(KeepTogether(about_city_content))
            
            # Build PDF
            doc.build(story)
            
        except Exception as e:
            print(f"Error generating PDF: {e}")
            messagebox.showerror("Error", f"Error generating PDF: {str(e)}")
            
    def create_cover_page(self, data, accent_gold, primary_blue):
        """Create the cover page with logo, property address, main image, thumbnails, and footer"""
        # Always create cover page, even if no images (will use placeholders)
        images = self.images if self.images else []
        return CoverPageFlowable(data, images, accent_gold, primary_blue)
            
    def create_header(self, data, accent_gold, primary_blue):
        """Create a professional header with branding"""
        header_drawing = Drawing(7.5*inch, 1.5*inch)
        
        # Background rectangle
        header_drawing.add(Rect(0, 0, 7.5*inch, 1.5*inch, fillColor=accent_gold, strokeColor=accent_gold))
        
        # Logo area (simplified)
        header_drawing.add(String(0.5*inch, 1*inch, "Property Report", 
                                 fontName="Helvetica-Bold", fontSize=24, fillColor=colors.white))
        
        # Tagline
        header_drawing.add(String(0.5*inch, 0.7*inch, "Professional Investment Analysis", 
                                 fontName="Helvetica", fontSize=12, fillColor=colors.white))
        
        return header_drawing
        
    def create_epc_chart(self, data, primary_blue, accent_gold, success_green):
        """Create a visual EPC rating chart with vertical bars"""
        # Create a vertical bar chart - narrower width for middle column
        drawing = Drawing(3.5*inch, 2.5*inch)
        
        # EPC bands with colors (A at top, G at bottom)
        epc_bands = [
            ('A', 92, 100, HexColor('#008450')),  # Dark green
            ('B', 81, 91, HexColor('#22c55e')),   # Green
            ('C', 69, 80, HexColor('#84cc16')),   # Light green
            ('D', 55, 68, HexColor('#eab308')),   # Yellow
            ('E', 39, 54, HexColor('#f59e0b')),   # Orange
            ('F', 21, 38, HexColor('#ef4444')),   # Red
            ('G', 1, 20, HexColor('#dc2626'))     # Dark red
        ]
        
        # Create vertical bars for EPC chart (stacked vertically as continuous scale)
        bar_width = 0.5*inch
        total_bar_height = 1.9*inch
        bar_height = total_bar_height / len(epc_bands)
        start_x = 0.8*inch  # Position for the bar
        start_y = 0.2*inch
        
        # Draw bars from top (A) to bottom (G) as continuous stacked rectangles
        for i, (grade, min_score, max_score, color) in enumerate(epc_bands):
            y = start_y + (len(epc_bands) - 1 - i) * bar_height
            
            # Bar background - continuous vertical scale
            drawing.add(Rect(start_x, y, bar_width, bar_height, 
                           fillColor=color, strokeColor=colors.black, strokeWidth=0.5))
            
            # Grade label on the right side of the bar
            label_x = start_x + bar_width + 0.15*inch
            label_y = y + bar_height / 2
            drawing.add(String(label_x, label_y, grade, 
                             fontName="Helvetica-Bold", fontSize=13, 
                             fillColor=colors.black, textAnchor="start"))
            
            # Score range next to grade
            range_x = label_x + 0.35*inch
            if min_score == 92:
                range_text = f"{min_score}+"
            else:
                range_text = f"{min_score}-{max_score}"
            drawing.add(String(range_x, label_y, range_text, 
                             fontName="Helvetica", fontSize=10, 
                             fillColor=HexColor('#666666'), textAnchor="start"))
        
        # Add current and potential scores
        current_score = int(data.get('current_rating', '72'))
        potential_score = int(data.get('potential_rating', '84'))
        
        # Find which band the scores fall into and calculate position
        def get_y_position(score):
            for i, (_, min_score, max_score, _) in enumerate(epc_bands):
                if min_score <= score <= max_score:
                    # Calculate position within the band
                    band_index = len(epc_bands) - 1 - i
                    y_base = start_y + band_index * bar_height
                    # Position in middle of band
                    return y_base + bar_height / 2
            return start_y + len(epc_bands) * bar_height / 2
        
        current_y = get_y_position(current_score)
        potential_y = get_y_position(potential_score)
        
        # Add "Current" label and score on the left
        current_x = start_x - 0.4*inch
        drawing.add(String(current_x, current_y + 0.08*inch, "Current", 
                         fontName="Helvetica", fontSize=9, 
                         fillColor=colors.black, textAnchor="end"))
        drawing.add(String(current_x, current_y - 0.08*inch, f"{current_score}", 
                         fontName="Helvetica-Bold", fontSize=12, 
                         fillColor=colors.black, textAnchor="end"))
        
        # Add "Potential" label and score on the left
        drawing.add(String(current_x, potential_y + 0.08*inch, "Potential", 
                         fontName="Helvetica", fontSize=9, 
                         fillColor=colors.black, textAnchor="end"))
        drawing.add(String(current_x, potential_y - 0.08*inch, f"{potential_score}", 
                         fontName="Helvetica-Bold", fontSize=12, 
                         fillColor=colors.black, textAnchor="end"))
        
        # Add text above chart
        drawing.add(String(start_x + bar_width/2, start_y + total_bar_height + 0.18*inch, 
                         "Very energy efficient - lower running costs", 
                         fontName="Helvetica", fontSize=8, 
                         fillColor=HexColor('#666666'), textAnchor="middle"))
        
        # Add text below chart
        drawing.add(String(start_x + bar_width/2, start_y - 0.18*inch, 
                         "Not energy efficient - higher running costs", 
                         fontName="Helvetica", fontSize=8, 
                         fillColor=HexColor('#666666'), textAnchor="middle"))
        
        return drawing

def main():
    root = tk.Tk()
    app = PDFBuilderApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()