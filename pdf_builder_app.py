import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import os
import sys
import math
import datetime
try:
    import requests  # type: ignore[import]
except ImportError:  # pragma: no cover - optional dependency handled at runtime
    requests = None
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
    def __init__(self, data, images_by_section, accent_gold, primary_blue):
        Flowable.__init__(self)
        self.data = data
        self.images_by_section = images_by_section
        self.accent_gold = accent_gold
        self.primary_blue = primary_blue
        # Account for margins - use exact frame size from reportlab
        # Error showed frame size: 475.28 x 721.89 points
        # A4 is 595.28 x 841.89 points
        # So available space is: 475.28 x 721.89 points
        # Convert to inches: 475.28/72 = 6.6 inches, 721.89/72 = 10.0 inches
        # Increase internal margins so the flowable fits within the document frame (matching new top margin)
        margin_total = 2.6*inch
        self.width = A4[0] - margin_total
        self.height = A4[1] - margin_total
        
    def draw(self):
        """Draw the cover page - note: reportlab uses bottom-left as origin"""
        canvas = self.canv
        # Use the declared width/height for content
        width = self.width
        height = self.height
        
        # Logo and tagline are drawn in _draw_cover_header callback
        # Calculate tagline_y position to match _draw_cover_header logic
        # Logo is positioned HEADER_TOP_OFFSET from top, with tagline 12 points below logo bottom
        page_height = A4[1]  # A4 page height in points
        HEADER_TOP_OFFSET = 0.45 * inch  # Match PDFBuilderApp.HEADER_TOP_OFFSET
        logo_width = 1.4 * inch  # Match standard header
        # Estimate logo height (assuming roughly square logo, adjust if needed)
        # In _draw_cover_header, logo_height = logo_width * (img_height / img_width)
        # Using a reasonable default aspect ratio of 1:1 (square logo)
        logo_height = logo_width  # Default to square, will be adjusted if logo exists
        header_top = page_height - HEADER_TOP_OFFSET
        logo_y = header_top - logo_height
        TAGLINE_SPACING = 12  # Points between logo bottom and tagline baseline
        tagline_y = logo_y - TAGLINE_SPACING  # Match _draw_cover_header
        
        # Start content below the header area
        # "Property Report" text - smaller, regular style
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
        
        # Main image - select from configured sections
        # Position main image below the address with spacing
        main_image_height = 4.5*inch
        main_image_bottom = address_bottom - spacing_between_text_and_images - main_image_height
        main_image_width = width
        
        cover_images = self.images_by_section.get('cover', [])
        gallery_images = self.images_by_section.get('property', [])
        fallback_sections = ['floor_plans', 'directions', 'city']
        
        main_img_path = None
        candidate_groups = [cover_images, gallery_images] + [
            self.images_by_section.get(section_key, []) for section_key in fallback_sections
        ]
        for group in candidate_groups:
            if group:
                main_img_path = group[0]
                break
        
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
        
        thumbnail_images = []
        thumbnail_candidates = []
        if len(cover_images) > 1:
            thumbnail_candidates.extend(cover_images[1:])
        thumbnail_candidates.extend(img for img in gallery_images if img != main_img_path)
        
        for img_path in thumbnail_candidates:
            if len(thumbnail_images) >= 3:
                break
            thumbnail_images.append(img_path)
        
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
        self.root.configure(background="#f3f4f6")
        
        # Data storage
        self.property_data = {}
        self.logo_path = get_resource_path("logo.png")
        self.header_logo_image = None
        
        # Header positioning constant - consistent across all pages
        self.HEADER_TOP_OFFSET = 0.45 * inch  # Distance from top of page to top of logo
        if requests is not None:
            self.http_session = requests.Session()
            self.http_session.headers.update({
                "User-Agent": "PropertyPDFBuilder/1.0 (+https://property-pdf.local)"
            })
        else:
            self.http_session = None

        # Styling
        self.style = ttk.Style()
        for theme in ("vista", "xpnative", "clam", "default"):
            if theme in self.style.theme_names():
                self.style.theme_use(theme)
                break
        self.style.configure("TFrame", background="#f3f4f6")
        self.style.configure("Content.TFrame", background="#f3f4f6")
        self.style.configure("Section.TFrame", background="#ffffff")
        self.style.configure("Section.TLabelframe", background="#ffffff", borderwidth=0, padding=12)
        self.style.configure("Section.TLabelframe.Label", font=("Segoe UI Semibold", 11))
        self.style.configure("TLabel", background="#f3f4f6", font=("Segoe UI", 10))
        self.style.configure("Section.TLabel", background="#ffffff", font=("Segoe UI", 10))
        self.style.configure("TEntry", foreground="#111827", fieldbackground="#ffffff")
        self.style.configure(
            "Accent.TButton",
            padding=8,
            font=("Segoe UI", 10),
            foreground="#ffffff",
            background="#1e3a8a"
        )
        self.style.map(
            "Accent.TButton",
            foreground=[("disabled", "#d4d4d8"), ("!disabled", "#ffffff")],
            background=[("active", "#1d4ed8"), ("pressed", "#1e40af"), ("!disabled", "#1e3a8a")]
        )
        self.image_sections_config = {
            'cover': {
                'title': 'Cover Page Images',
                'description': 'First image appears as the large hero photo. Additional images become cover thumbnails.',
                'allow_reorder': True,
                'max_items': None,
                'list_height': 5
            },
            'property': {
                'title': 'Property Gallery',
                'description': 'General property photos used throughout the report and cover thumbnails.',
                'allow_reorder': True,
                'max_items': None,
                'list_height': 6
            },
            'floor_plans': {
                'title': 'Floor Plans',
                'description': 'Uploaded plans appear on dedicated floor plan pages.',
                'allow_reorder': True,
                'max_items': None,
                'list_height': 5
            },
            'directions': {
                'title': 'Directions / Map',
                'description': 'Used on the “Getting To The City Centre” page. Only the first image is used.',
                'allow_reorder': False,
                'max_items': 1,
                'list_height': 3
            },
            'city': {
                'title': 'City Images',
                'description': 'Urban lifestyle shots shown together near the end of the report.',
                'allow_reorder': True,
                'max_items': 3,
                'list_height': 4
            }
        }
        self.image_sections = {key: [] for key in self.image_sections_config.keys()}
        self.image_listboxes = {}
        
        # Create main frame
        self.main_frame = ttk.Frame(root, padding="20", style="Content.TFrame")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        
        self.create_widgets()
        self.load_mock_data_defaults()
        self.load_default_images()
        
    def create_widgets(self):
        """Create all the GUI widgets with tabs"""
        
        # Header with logo and title
        header_frame = ttk.Frame(self.main_frame, style="Content.TFrame")
        header_frame.grid(row=0, column=0, sticky=tk.W, pady=(0, 12))
        if os.path.exists(self.logo_path):
            try:
                logo_image = Image.open(self.logo_path)
                max_width = 220
                aspect_ratio = logo_image.height / logo_image.width if logo_image.width else 1
                resized_height = int(max_width * aspect_ratio)
                resized_logo = logo_image.resize((max_width, resized_height), Image.LANCZOS)
                self.header_logo_image = ImageTk.PhotoImage(resized_logo)
                ttk.Label(header_frame, image=self.header_logo_image, style="Content.TFrame").pack(side=tk.LEFT)
            except Exception as exc:
                print(f"Unable to load logo: {exc}")
        ttk.Label(
            header_frame,
            text="Property Investment Report Builder",
            font=("Segoe UI Semibold", 14),
            background="#f3f4f6"
        ).pack(side=tk.LEFT, padx=(12, 0))
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 16))
        self.main_frame.rowconfigure(1, weight=1)
        
        # Create tabs
        self.create_property_tab()
        self.create_investment_tab()
        self.create_epc_tab()
        self.create_location_tab()
        self.create_images_tab()
        
        # Main buttons
        button_frame = ttk.Frame(self.main_frame, style="Content.TFrame")
        button_frame.grid(row=2, column=0, pady=(8, 0))
        
        ttk.Button(button_frame, text="Generate Investment Report PDF",
                   command=self.generate_pdf).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Button(button_frame, text="Clear All", command=self.clear_all).pack(side=tk.LEFT, padx=(0, 12))
        
    def create_property_tab(self):
        """Create Property Information tab"""
        self.property_frame = ttk.Frame(self.notebook)
        self.property_frame.pack_propagate(False)
        self.notebook.add(self.property_frame, text="Property Info")
        
        # Property Information Section
        info_frame = ttk.LabelFrame(self.property_frame, text="Property Information", style="Section.TLabelframe")
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
            row_frame = ttk.Frame(info_frame, style="Section.TFrame")
            row_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(row_frame, text=label_text, width=22, style="Section.TLabel").pack(side=tk.LEFT, padx=(0, 12))
            
            if field_name in ["key_features", "description"]:
                widget = tk.Text(row_frame, height=3, width=60, bg="#ffffff", fg="#111827", wrap="word", insertbackground="#1e3a8a")
                widget.configure(font=("Segoe UI", 10))
                widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            else:
                widget = ttk.Entry(row_frame, width=60)
                widget.configure(font=("Segoe UI", 10))
                widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            self.entry_widgets[field_name] = widget
            
    def create_investment_tab(self):
        """Create Investment Analysis tab"""
        self.investment_frame = ttk.Frame(self.notebook)
        self.investment_frame.pack_propagate(False)
        self.notebook.add(self.investment_frame, text="Investment Analysis")
        
        # Investment Analysis Section
        investment_frame = ttk.LabelFrame(self.investment_frame, text="Investment Analysis", style="Section.TLabelframe")
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
            row_frame = ttk.Frame(investment_frame, style="Section.TFrame")
            row_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(row_frame, text=label_text, width=26, style="Section.TLabel").pack(side=tk.LEFT, padx=(0, 12))
            
            widget = ttk.Entry(row_frame, width=30)
            widget.configure(font=("Segoe UI", 10))
            widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            self.entry_widgets[field_name] = widget
            
    def create_epc_tab(self):
        """Create EPC & Details tab"""
        self.epc_frame = ttk.Frame(self.notebook)
        self.epc_frame.pack_propagate(False)
        self.notebook.add(self.epc_frame, text="EPC & Details")
        
        # EPC Section
        epc_frame = ttk.LabelFrame(self.epc_frame, text="Energy Performance Certificate", style="Section.TLabelframe")
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
            row_frame = ttk.Frame(epc_frame, style="Section.TFrame")
            row_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(row_frame, text=label_text, width=26, style="Section.TLabel").pack(side=tk.LEFT, padx=(0, 12))
            
            widget = ttk.Entry(row_frame, width=30)
            widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            self.entry_widgets[field_name] = widget
            
    def create_location_tab(self):
        """Create Location & Transport tab"""
        self.location_frame = ttk.Frame(self.notebook)
        self.location_frame.pack_propagate(False)
        self.notebook.add(self.location_frame, text="Location & Transport")
        
        # Location Section
        location_frame = ttk.LabelFrame(self.location_frame, text="Location Information", style="Section.TLabelframe")
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
            row_frame = ttk.Frame(location_frame, style="Section.TFrame")
            row_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(
                row_frame,
                text=label_text,
                width=34,
                anchor=tk.W,
                wraplength=260,
                style="Section.TLabel"
            ).pack(side=tk.LEFT, padx=(0, 12))
            
            if field_name == "about_city":
                widget = tk.Text(row_frame, height=4, width=50, bg="#ffffff", fg="#111827", wrap="word", insertbackground="#1e3a8a")
                widget.configure(font=("Segoe UI", 10))
                widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            else:
                widget = ttk.Entry(row_frame, width=30)
                widget.configure(font=("Segoe UI", 10))
                widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            self.entry_widgets[field_name] = widget
        
        ttk.Button(
            location_frame,
            text="Auto-Fill Location From Web",
            command=self.auto_fill_location_from_web
        ).pack(anchor=tk.W, pady=(12, 0))
            
    def create_images_tab(self):
        """Create Images tab"""
        self.images_frame = ttk.Frame(self.notebook)
        self.images_frame.pack_propagate(False)
        self.notebook.add(self.images_frame, text="Images")
        
        # Instructions Section
        instructions_frame = ttk.LabelFrame(self.images_frame, text="Image Upload Guide", style="Section.TLabelframe")
        instructions_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        ttk.Label(
            instructions_frame,
            text="Organise your images by section so it’s clear where each one appears in the report.",
            style="Section.TLabel",
            wraplength=820,
            justify=tk.LEFT
        ).pack(anchor=tk.W)
        
        # Scrollable container for image sections
        canvas_frame = ttk.Frame(self.images_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        canvas = tk.Canvas(canvas_frame, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda event: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        def _sync_frame_width(event):
            canvas.itemconfigure(scrollable_window, width=event.width)
        
        canvas.bind("<Configure>", _sync_frame_width)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        def _bind_mousewheel(widget):
            widget.bind("<Enter>", lambda _: canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units")))
            widget.bind("<Leave>", lambda _: canvas.unbind_all("<MouseWheel>"))
        
        _bind_mousewheel(scrollable_frame)
        
        sections_container = scrollable_frame
        
        for section_key, config in self.image_sections_config.items():
            self._create_image_section_ui(sections_container, section_key, config)
        
        # Global image actions
        global_actions = ttk.Frame(self.images_frame, style="Content.TFrame")
        global_actions.pack(fill=tk.X, padx=10, pady=(0, 10))
        ttk.Button(global_actions, text="Clear All Images", command=self.clear_image_sections).pack(side=tk.LEFT)
        
    def _create_image_section_ui(self, parent, section_key, config):
        section_frame = ttk.LabelFrame(parent, text=config['title'], style="Section.TLabelframe")
        section_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        if config.get('description'):
            ttk.Label(
                section_frame,
                text=config['description'],
                style="Section.TLabel",
                justify=tk.LEFT,
                wraplength=780
            ).pack(anchor=tk.W, pady=(0, 8))
        
        listbox_frame = ttk.Frame(section_frame, style="Section.TFrame")
        listbox_frame.pack(fill=tk.BOTH, expand=True)
        
        listbox = tk.Listbox(listbox_frame, height=config.get('list_height', 5))
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.image_listboxes[section_key] = listbox
        
        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))
        listbox.configure(yscrollcommand=scrollbar.set)
        
        button_frame = ttk.Frame(listbox_frame, style="Section.TFrame")
        button_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        
        add_label = "Add Image" if config.get('max_items') == 1 else "Add Images"
        ttk.Button(
            button_frame,
            text=add_label,
            command=lambda key=section_key: self.add_images_to_section(key)
        ).pack(fill=tk.X, pady=(0, 5))
        
        ttk.Button(
            button_frame,
            text="Remove Selected",
            command=lambda key=section_key: self.remove_selected_image(key)
        ).pack(fill=tk.X)
        
        if config.get('allow_reorder'):
            ttk.Button(
                button_frame,
                text="Move Up",
                command=lambda key=section_key: self.move_image_up_in_section(key)
            ).pack(fill=tk.X, pady=(10, 2))
            
            ttk.Button(
                button_frame,
                text="Move Down",
                command=lambda key=section_key: self.move_image_down_in_section(key)
            ).pack(fill=tk.X)
    
    def add_images_to_section(self, section_key):
        """Prompt user to add images to a specific section"""
        file_paths = filedialog.askopenfilenames(
            title="Select Images",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.bmp *.tiff *.webp")]
        )
        if not file_paths:
            return
        
        config = self.image_sections_config.get(section_key, {})
        for file_path in file_paths:
            max_items = config.get('max_items')
            if max_items and len(self.image_sections[section_key]) >= max_items:
                messagebox.showwarning(
                    "Image limit reached",
                    f"{config['title']} allows up to {max_items} image(s). Remove an image before adding a new one."
                )
                break
            self._add_image_path(section_key, file_path, skip_duplicates=True)
    
    def remove_selected_image(self, section_key):
        """Remove the selected image from a section"""
        listbox = self.image_listboxes.get(section_key)
        if not listbox:
            return
        selection = listbox.curselection()
        if selection:
            index = selection[0]
            listbox.delete(index)
            del self.image_sections[section_key][index]
    
    def move_image_up_in_section(self, section_key):
        """Move the selected image up within a section"""
        listbox = self.image_listboxes.get(section_key)
        if not listbox:
            return
        selection = listbox.curselection()
        if selection and selection[0] > 0:
            index = selection[0]
            images = self.image_sections[section_key]
            images[index], images[index - 1] = images[index - 1], images[index]
            
            current_text = listbox.get(index)
            above_text = listbox.get(index - 1)
            listbox.delete(index - 1, index)
            listbox.insert(index - 1, current_text)
            listbox.insert(index, above_text)
            listbox.selection_set(index - 1)
    
    def move_image_down_in_section(self, section_key):
        """Move the selected image down within a section"""
        listbox = self.image_listboxes.get(section_key)
        if not listbox:
            return
        selection = listbox.curselection()
        images = self.image_sections.get(section_key, [])
        if selection and selection[0] < len(images) - 1:
            index = selection[0]
            images[index], images[index + 1] = images[index + 1], images[index]
            
            current_text = listbox.get(index)
            below_text = listbox.get(index + 1)
            listbox.delete(index, index + 1)
            listbox.insert(index, below_text)
            listbox.insert(index + 1, current_text)
            listbox.selection_set(index + 1)
    
    def auto_fill_location_from_web(self):
        """Attempt to populate the Location & Transport fields using online open-data sources."""
        if self.http_session is None or requests is None:
            messagebox.showerror(
                "Dependency Missing",
                "The 'requests' package is not available. Install it with 'pip install requests' to enable auto-fill."
            )
            return
        address = self._get_widget_value('address')
        postal_code = self._get_widget_value('postal_code')
        city = self._get_widget_value('city')
        if not city:
            city = self._infer_city_from_address(address)
        if not city:
            messagebox.showwarning("Missing City", "Please enter a city name before using auto-fill.")
            return
        self.root.config(cursor="watch")
        self.root.update_idletasks()
        try:
            location_data = self._gather_location_data(address, postal_code, city)
            if not location_data:
                messagebox.showinfo("No Data Found", f"Could not locate enough data for {city}.")
                return
            self._apply_location_data(location_data)
            messagebox.showinfo("Location Updated", "Location & transport fields were filled using open web data.")
        except requests.exceptions.RequestException as req_err:
            messagebox.showerror("Network Error", f"Unable to reach the data services:\n{req_err}")
        except Exception as exc:
            messagebox.showerror("Auto-Fill Error", f"Unable to auto-fill location data:\n{exc}")
        finally:
            self.root.config(cursor="")
    
    def _gather_location_data(self, address, postal_code, city):
        combined_address = address.strip()
        if postal_code and postal_code not in combined_address:
            combined_address = f"{combined_address}, {postal_code}" if combined_address else postal_code
        
        property_coords = self._geocode(combined_address) if combined_address else None
        city_coords = self._geocode(city)
        if property_coords is None and city_coords:
            property_coords = city_coords
        
        data = {'city': city}
        
        # About city and population from Wikipedia/Wikidata
        population = self._fetch_population(city)
        if population:
            data['population'] = f"{population:,}"
        about_city = self._fetch_city_summary(city)
        if about_city:
            data['about_city'] = about_city
        
        # Distance and travel time estimates
        if property_coords and city_coords:
            distance_miles = self._haversine_miles(property_coords['lat'], property_coords['lon'],
                                                   city_coords['lat'], city_coords['lon'])
            if distance_miles is not None:
                data['distance_city_centre'] = f"{distance_miles:.1f}"
                car_minutes = max(5, int(round(distance_miles / 18 * 60)))
                public_minutes = max(car_minutes + 5, int(round(car_minutes * 1.3)))
                data['time_car'] = str(car_minutes)
                data['time_public_transport'] = str(public_minutes)
        elif city_coords:
            data['distance_city_centre'] = "0"
            data['time_car'] = "10"
            data['time_public_transport'] = "15"
        
        # Nearest rail station
        station_info = self._find_nearest_station(property_coords)
        if station_info:
            if station_info.get('distance_miles') is not None:
                distance_miles = station_info['distance_miles']
                data['station_distance'] = f"{distance_miles:.2f}"
                walk_minutes = max(3, int(round(distance_miles / 3.0 * 60)))
                data['walk_to_station'] = str(walk_minutes)
            if station_info.get('name'):
                data.setdefault('bus_routes', f"Nearest station: {station_info['name']}")
        
        # Bus route hints
        bus_info = self._fetch_bus_information(property_coords)
        if bus_info:
            routes = bus_info.get('routes')
            if routes:
                data['bus_routes'] = ", ".join(routes)
            frequency = bus_info.get('frequency')
            if frequency:
                data['bus_frequency'] = frequency
        
        defaults = {
            'population': "N/A",
            'distance_city_centre': "N/A",
            'time_car': "N/A",
            'time_public_transport': "N/A",
            'walk_to_station': "N/A",
            'station_distance': "N/A",
            'bus_routes': "N/A",
            'bus_frequency': "N/A",
            'about_city': "Information currently unavailable."
        }
        for key, fallback in defaults.items():
            data.setdefault(key, fallback)
        
        return {k: v for k, v in data.items() if v}
    
    def _apply_location_data(self, data):
        for field, value in data.items():
            if field not in self.entry_widgets or value is None:
                continue
            self._set_widget_value(field, value)
    
    def _get_widget_value(self, field_name):
        widget = self.entry_widgets.get(field_name)
        if not widget:
            return ""
        if isinstance(widget, tk.Text):
            return widget.get("1.0", tk.END).strip()
        return widget.get().strip()
    
    def _set_widget_value(self, field_name, value):
        widget = self.entry_widgets.get(field_name)
        if not widget:
            return
        if isinstance(widget, tk.Text):
            widget.delete("1.0", tk.END)
            widget.insert("1.0", value)
        else:
            widget.delete(0, tk.END)
            widget.insert(0, value)
    
    def _infer_city_from_address(self, address):
        if not address:
            return ""
        parts = [part.strip() for part in address.split(',') if part.strip()]
        return parts[-1] if parts else ""
    
    def _geocode(self, query):
        if not query:
            return None
        if self.http_session is None:
            return None
        response = self.http_session.get(
            "https://nominatim.openstreetmap.org/search",
            params={"q": query, "format": "json", "limit": 1, "addressdetails": 0},
            timeout=15
        )
        response.raise_for_status()
        results = response.json()
        if not results:
            return None
        result = results[0]
        try:
            return {
                "lat": float(result['lat']),
                "lon": float(result['lon']),
                "display_name": result.get('display_name', query)
            }
        except (KeyError, ValueError):
            return None
    
    def _haversine_miles(self, lat1, lon1, lat2, lon2):
        if None in (lat1, lon1, lat2, lon2):
            return None
        r = 3958.8  # Earth radius in miles
        phi1 = math.radians(lat1)
        phi2 = math.radians(lat2)
        delta_phi = math.radians(lat2 - lat1)
        delta_lambda = math.radians(lon2 - lon1)
        a = math.sin(delta_phi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(delta_lambda / 2) ** 2
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
        return r * c
    
    def _find_nearest_station(self, coords):
        if not coords:
            return None
        query = f"""
        [out:json][timeout:25];
        node["railway"="station"](around:10000,{coords['lat']},{coords['lon']});
        out body;
        """
        data = self._run_overpass_query(query)
        best = None
        for element in data.get('elements', []):
            lat = element.get('lat')
            lon = element.get('lon')
            if lat is None or lon is None:
                continue
            distance = self._haversine_miles(coords['lat'], coords['lon'], lat, lon)
            if distance is None:
                continue
            if not best or distance < best['distance_miles']:
                best = {
                    "name": element.get('tags', {}).get('name', 'Unnamed Station'),
                    "distance_miles": distance
                }
        return best
    
    def _fetch_bus_information(self, coords):
        if not coords:
            return None
        query = f"""
        [out:json][timeout:25];
        node["highway"="bus_stop"](around:1200,{coords['lat']},{coords['lon']});
        out body;
        """
        data = self._run_overpass_query(query)
        elements = data.get('elements', [])
        if not elements:
            return None
        routes = set()
        nearest_stop = None
        nearest_distance = None
        for element in elements:
            tags = element.get('tags', {})
            for key in ('route_ref', 'ref', 'bus_routes'):
                value = tags.get(key)
                if value:
                    for part in value.replace('/', ';').split(';'):
                        part = part.strip()
                        if part:
                            routes.add(part)
            lat = element.get('lat')
            lon = element.get('lon')
            if lat is None or lon is None:
                continue
            distance = self._haversine_miles(coords['lat'], coords['lon'], lat, lon)
            if distance is None:
                continue
            if nearest_distance is None or distance < nearest_distance:
                nearest_distance = distance
                nearest_stop = tags.get('name') or tags.get('ref') or "Nearby bus stop"
        frequency_text = None
        if nearest_distance is not None:
            walk_minutes = max(2, int(round(nearest_distance / 3.0 * 60)))
            frequency_text = f"{nearest_stop} (~{walk_minutes} min walk). Typical city routes run every 10–15 minutes."
        if not routes and nearest_stop:
            routes.add(nearest_stop)
        if not routes:
            return {"frequency": frequency_text} if frequency_text else None
        return {
            "routes": sorted(routes)[:5],
            "frequency": frequency_text
        }
    
    def _run_overpass_query(self, query):
        if self.http_session is None:
            return {"elements": []}
        response = self.http_session.get(
            "https://overpass-api.de/api/interpreter",
            params={"data": query},
            timeout=30
        )
        response.raise_for_status()
        return response.json()
    
    def _fetch_city_summary(self, city):
        if self.http_session is None:
            return ""
        try:
            response = self.http_session.get(
                f"https://en.wikipedia.org/api/rest_v1/page/summary/{city}",
                timeout=20
            )
            if response.status_code != 200:
                return ""
            data = response.json()
            extract = data.get('extract', '')
            if not extract:
                return ""
            return extract if len(extract) <= 900 else extract[:900].rsplit(' ', 1)[0] + "..."
        except requests.exceptions.RequestException:
            return ""
    
    def _fetch_population(self, city):
        if self.http_session is None:
            return None
        try:
            page_resp = self.http_session.get(
                "https://en.wikipedia.org/w/api.php",
                params={
                    "action": "query",
                    "format": "json",
                    "titles": city,
                    "prop": "pageprops"
                },
                timeout=20
            )
            page_resp.raise_for_status()
            page_data = page_resp.json()
            pages = page_data.get('query', {}).get('pages', {})
            if not pages:
                return None
            page = next(iter(pages.values()))
            wikidata_id = page.get('pageprops', {}).get('wikibase_item')
            if not wikidata_id:
                return None
            wd_resp = self.http_session.get(
                f"https://www.wikidata.org/wiki/Special:EntityData/{wikidata_id}.json",
                timeout=20
            )
            wd_resp.raise_for_status()
            entity = wd_resp.json().get('entities', {}).get(wikidata_id, {})
            claims = entity.get('claims', {})
            population_claims = claims.get('P1082', [])
            best_value = None
            best_time = None
            for claim in population_claims:
                mainsnak = claim.get('mainsnak', {})
                datavalue = mainsnak.get('datavalue', {})
                value = datavalue.get('value')
                if not isinstance(value, dict):
                    continue
                amount = value.get('amount')
                if amount is None:
                    continue
                try:
                    amount_int = int(float(amount))
                except ValueError:
                    continue
                time_qualifiers = claim.get('qualifiers', {}).get('P585', [])
                qualifier_time = None
                if time_qualifiers:
                    time_val = time_qualifiers[0].get('datavalue', {}).get('value', {}).get('time')
                    qualifier_time = time_val
                if best_value is None or (qualifier_time and qualifier_time > (best_time or "")):
                    best_value = amount_int
                    best_time = qualifier_time
            return best_value
        except requests.exceptions.RequestException:
            return None
    
    def clear_image_sections(self):
        """Remove all images from every section"""
        for section_key in self.image_sections.keys():
            self.image_sections[section_key].clear()
            listbox = self.image_listboxes.get(section_key)
            if listbox:
                listbox.delete(0, tk.END)
    
    def _add_image_path(self, section_key, file_path, skip_duplicates=False, skip_if_full=False):
        """Internal helper to add an image path to a section and update UI"""
        images = self.image_sections.setdefault(section_key, [])
        config = self.image_sections_config.get(section_key, {})
        max_items = config.get('max_items')
        
        if skip_duplicates and file_path in images:
            return
        
        if max_items and len(images) >= max_items:
            if skip_if_full:
                return
            messagebox.showwarning(
                "Image limit reached",
                f"{config.get('title', 'This section')} allows up to {max_items} image(s)."
            )
            return
        
        images.append(file_path)
        listbox = self.image_listboxes.get(section_key)
        if listbox:
            listbox.insert(tk.END, os.path.basename(file_path))
    
    def load_mock_data_defaults(self):
        """Populate the form with bundled mock data for quicker previews."""
        mock_data = {
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
            'epc_grade': 'C',
            'current_rating': '84',
            'potential_rating': '72',
            'inspection_date': '30th January 2019',
            'window_glazing': 'Double glazing installed during or after 2002',
            'building_age': 'before 1900',
            'broadband_available': 'Broadband available',
            'download_speed': '1,800 Mbps',
            'upload_speed': '220 Mbps',
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
        for field_name, value in mock_data.items():
            widget = self.entry_widgets.get(field_name)
            if not widget:
                continue
            if isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
                widget.insert("1.0", value)
            else:
                widget.delete(0, tk.END)
                widget.insert(0, value)
    
    def load_default_images(self):
        """Load bundled sample images into their sections as placeholders."""
        sample_folder = get_resource_path("sample_images")
        if not os.path.exists(sample_folder):
            return
        self.clear_image_sections()
        image_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp')
        for filename in sorted(os.listdir(sample_folder)):
            if not filename.lower().endswith(image_extensions):
                continue
            image_path = os.path.join(sample_folder, filename)
            section_key = self.get_image_section(filename)
            self._add_image_path(section_key, image_path, skip_duplicates=True, skip_if_full=True)
    
    def _draw_cover_header(self, canvas, doc):
        """Cover page - draw logo and tagline consistent with standard header."""
        if not os.path.exists(self.logo_path):
            return
        try:
            logo_reader = ImageReader(self.logo_path)
            canvas.saveState()
            img_width, img_height = logo_reader.getSize()
            if not img_width or not img_height:
                canvas.restoreState()
                return
            logo_width = 1.4 * inch  # Match standard header
            logo_height = logo_width * (img_height / img_width)
            page_width, page_height = doc.pagesize
            # Use consistent top offset - same as standard header
            # Logo top is positioned HEADER_TOP_OFFSET from the top of the page
            header_top = page_height - self.HEADER_TOP_OFFSET
            logo_x = doc.leftMargin  # Match standard header positioning
            logo_y = header_top - logo_height
            canvas.drawImage(
                logo_reader,
                logo_x,
                logo_y,
                width=logo_width,
                height=logo_height,
                mask='auto',
                preserveAspectRatio=True
            )
            # Tagline positioned consistently below logo - match standard header
            TAGLINE_FONT_SIZE = 9
            TAGLINE_SPACING = 12  # Points between logo bottom and tagline baseline
            canvas.setFont("Helvetica", TAGLINE_FONT_SIZE)  # Match standard header
            canvas.setFillColor(HexColor('#334155'))  # Match standard header
            tagline_y = logo_y - TAGLINE_SPACING
            canvas.drawString(
                logo_x,
                tagline_y,
                "Elevating Your Property Experience"  # Match standard header text
            )
            canvas.restoreState()
        except Exception as exc:
            print(f"Error drawing cover page logo: {exc}")
    
    def _draw_standard_header(self, canvas, doc):
        """Draw consistent logo and tagline at the top of every PDF page."""
        if not os.path.exists(self.logo_path):
            return
        try:
            logo_reader = ImageReader(self.logo_path)
            canvas.saveState()
            img_width, img_height = logo_reader.getSize()
            if not img_width or not img_height:
                canvas.restoreState()
                return
            logo_width = 1.4 * inch
            logo_height = logo_width * (img_height / img_width)
            page_width, page_height = doc.pagesize
            # Use consistent top offset - same across all pages
            # Logo top is positioned HEADER_TOP_OFFSET from the top of the page
            header_top = page_height - self.HEADER_TOP_OFFSET
            logo_x = doc.leftMargin
            logo_y = header_top - logo_height
            canvas.drawImage(
                logo_reader,
                logo_x,
                logo_y,
                width=logo_width,
                height=logo_height,
                mask='auto',
                preserveAspectRatio=True
            )
            # Tagline positioned consistently below logo
            TAGLINE_FONT_SIZE = 9
            TAGLINE_SPACING = 12  # Points between logo bottom and tagline baseline
            canvas.setFont("Helvetica", TAGLINE_FONT_SIZE)
            canvas.setFillColor(HexColor('#334155'))
            tagline_y = logo_y - TAGLINE_SPACING
            canvas.drawString(
                logo_x,
                tagline_y,
                "Elevating Your Property Experience"
            )
            canvas.restoreState()
        except Exception as exc:
            print(f"Error drawing PDF header: {exc}")
    
    def get_image_section(self, filename):
        """Infer the most suitable section for an image based on its filename"""
        filename_lower = filename.lower()
        if 'exterior' in filename_lower and 'front' in filename_lower:
            return 'cover'
        if 'floor' in filename_lower or 'plan' in filename_lower:
            return 'floor_plans'
        if 'direction' in filename_lower or 'map' in filename_lower or 'city_centre' in filename_lower:
            return 'directions'
        if 'city' in filename_lower or 'liverpool' in filename_lower or 'urban' in filename_lower:
            return 'city'
        return 'property'
            
    def clear_all(self):
        """Clear all form data"""
        for widget in self.entry_widgets.values():
            if isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
            else:
                widget.delete(0, tk.END)
        
        self.clear_image_sections()
        
        
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
            
            # Calculate exact header height for consistent spacing (before creating doc)
            # Use the same HEADER_TOP_OFFSET constant for consistency
            # Header positioning:
            #   - Logo top: HEADER_TOP_OFFSET from page top (0.45")
            #   - Logo height: varies by aspect ratio, typically ~0.5-0.7"
            #   - Tagline spacing: 12 points below logo bottom
            #   - Tagline font: 9pt, with ~2pt margin for text height
            #   - Content spacing: 0.3" below tagline bottom
            LOGO_AVERAGE_HEIGHT = 0.6*inch  # Average logo height (will vary slightly)
            TAGLINE_SPACING_POINTS = 12  # Points between logo bottom and tagline baseline
            TAGLINE_FONT_SIZE = 9  # Tagline font size in points
            TAGLINE_TEXT_MARGIN = 2  # Additional points for text height below baseline
            TAGLINE_TOTAL_HEIGHT = (TAGLINE_SPACING_POINTS + TAGLINE_FONT_SIZE + TAGLINE_TEXT_MARGIN) / 72.0 * inch  # Convert to inches
            HEADER_TO_CONTENT_SPACING = 0.3*inch  # Space from tagline bottom to first content
            # Total: top offset + logo + tagline + content spacing
            REQUIRED_TOP_MARGIN = self.HEADER_TOP_OFFSET + LOGO_AVERAGE_HEIGHT + TAGLINE_TOTAL_HEIGHT + HEADER_TO_CONTENT_SPACING
            
            # Create PDF document with custom margins
            doc = SimpleDocTemplate(
                file_path,
                pagesize=A4,
                leftMargin=0.75*inch,
                rightMargin=0.75*inch,
                topMargin=REQUIRED_TOP_MARGIN,
                bottomMargin=0.75*inch
            )
            story = []
            
            cover_images = list(self.image_sections.get('cover', []))
            property_gallery_images = list(self.image_sections.get('property', []))
            floor_plan_images = list(self.image_sections.get('floor_plans', []))
            directions_images = list(self.image_sections.get('directions', []))
            city_images = list(self.image_sections.get('city', []))
            
            # Define professional color scheme
            primary_blue = HexColor('#1e3a8a')  # Dark blue
            accent_gold = HexColor('#f59e0b')   # Gold
            light_grey = HexColor('#f8fafc')   # Light grey
            dark_grey = HexColor('#374151')    # Dark grey
            success_green = HexColor('#10b981') # Green
            warning_orange = HexColor('#f59e0b') # Orange
            
            # Standard spacing constants for consistency
            # Use the same calculation as topMargin for perfect alignment
            STANDARD_PAGE_BREAK_SPACING = REQUIRED_TOP_MARGIN  # Exact space after page break (matches topMargin)
            STANDARD_SECTION_SPACING = 20  # Space after section title (points)
            STANDARD_IMAGE_SPACING = 20  # Space after images (points)
            STANDARD_TABLE_SPACING = 20  # Space after tables (points)
            STANDARD_CONTENT_SPACING = 12  # Space between content elements (points)
            STANDARD_SUBSECTION_SPACING = 15  # Space after subsection headers (points)
            
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
                spaceAfter=STANDARD_SUBSECTION_SPACING,
                spaceBefore=0,
                textColor=primary_blue,
                fontName='Helvetica-Bold'
            )
            
            subheader_style = ParagraphStyle(
                'CustomSubHeader',
                parent=styles['Heading3'],
                fontSize=14,
                spaceAfter=STANDARD_CONTENT_SPACING,
                spaceBefore=0,
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

            section_title_style = ParagraphStyle(
                'SectionTitle',
                parent=styles['Heading1'],
                fontSize=24,
                spaceAfter=0,  # Use explicit Spacer for consistency
                spaceBefore=0,
                textColor=colors.black,
                fontName='Helvetica-Bold',
                alignment=TA_LEFT
            )
            
            # Create cover page (first page)
            cover_page = self.create_cover_page(data, accent_gold, primary_blue)
            if cover_page:
                story.append(cover_page)
                story.append(PageBreak())
                story.append(Spacer(1, STANDARD_PAGE_BREAK_SPACING))
            
            # Investment Opportunity Section - Second Page Design
            story.append(Paragraph("Investment Opportunity", section_title_style))
            story.append(Spacer(1, STANDARD_SECTION_SPACING))
            
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
            story.append(Spacer(1, STANDARD_TABLE_SPACING))
            
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
            story.append(KeepTogether([two_col_table, Spacer(1, 0.5*inch), profit_table]))
            
            # Page break before Key Information section
            story.append(PageBreak())
            story.append(Spacer(1, STANDARD_PAGE_BREAK_SPACING))
            
            # Key Information Section - Third Page Design (all content on one page)
            # Collect all content first, then wrap in KeepTogether
            key_info_content = []
            
            key_info_content.append(Paragraph("Key Information", section_title_style))
            key_info_content.append(Spacer(1, STANDARD_SECTION_SPACING))
            
            # Property image (reduced size to fit on page) - always show, use placeholder if missing
            img_width = 6.5*inch
            img_height = 3.5*inch
            main_img_path = None
            
            candidate_images = cover_images + property_gallery_images
            if candidate_images:
                try:
                    for img_path in candidate_images:
                        filename = os.path.basename(img_path).lower()
                        if 'exterior' in filename and 'front' in filename:
                            main_img_path = img_path
                            break
                    if not main_img_path:
                        main_img_path = candidate_images[0]
                    
                    if main_img_path and os.path.exists(main_img_path):
                        property_img_pil = Image.open(main_img_path)
                        img_aspect = property_img_pil.width / property_img_pil.height
                        img_width = 6.5*inch  # Slightly smaller to fit
                        img_height = img_width / img_aspect
                        if img_height > 3.5*inch:
                            img_height = 3.5*inch
                            img_width = img_height * img_aspect
                        property_img = RLImage(main_img_path, width=img_width, height=img_height)
                        key_info_content.append(property_img)
                    else:
                        placeholder = create_placeholder_drawing(img_width, img_height)
                        key_info_content.append(placeholder)
                    key_info_content.append(Spacer(1, STANDARD_IMAGE_SPACING))
                except Exception as e:
                    print(f"Error loading property image: {e}")
                    placeholder = create_placeholder_drawing(img_width, img_height)
                    key_info_content.append(placeholder)
                    key_info_content.append(Spacer(1, STANDARD_IMAGE_SPACING))
            else:
                placeholder = create_placeholder_drawing(img_width, img_height)
                key_info_content.append(placeholder)
                key_info_content.append(Spacer(1, STANDARD_IMAGE_SPACING))
            
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
            key_info_content.append(Spacer(1, STANDARD_TABLE_SPACING))
            
            # Key Features - Bulleted list
            if data.get('key_features'):
                # Split key features by newline and create bulleted list
                features_text = data.get('key_features', '')
                features_list = [f.strip() for f in features_text.split('\n') if f.strip()]
                
                key_info_content.append(Paragraph("Key Features", header_style))
                key_info_content.append(Spacer(1, STANDARD_CONTENT_SPACING))
                
                # Create bulleted list with reduced spacing to fit on page
                for feature in features_list:
                    bullet_para = Paragraph(f"• {feature}", 
                        ParagraphStyle('KeyFeature', parent=styles['Normal'], 
                                      fontSize=11, textColor=colors.black,
                                      leftIndent=20, spaceAfter=5))  # Further reduced spacing
                    key_info_content.append(bullet_para)
            
            # Use KeepTogether to ensure entire Key Information section stays on same page
            story.append(KeepTogether(key_info_content))
            
            # Page break after Key Information page
            story.append(PageBreak())
            story.append(Spacer(1, STANDARD_PAGE_BREAK_SPACING))
            
            # Other Key Information Page Header
            other_key_content = []
            other_key_content.append(Paragraph("Other Key Information", section_title_style))
            other_key_content.append(Spacer(1, STANDARD_SECTION_SPACING))
            
            # Large Property Image below header - always show, use placeholder if missing
            img_width = 7*inch
            img_height = 3.5*inch
            main_img_path = None
            
            candidate_images = cover_images + property_gallery_images
            if candidate_images:
                try:
                    for img_path in candidate_images:
                        filename = os.path.basename(img_path).lower()
                        if 'exterior' in filename and 'front' in filename:
                            main_img_path = img_path
                            break
                    if not main_img_path:
                        main_img_path = candidate_images[0]
                    
                    if main_img_path and os.path.exists(main_img_path):
                        property_img_pil = Image.open(main_img_path)
                        img_aspect = property_img_pil.width / property_img_pil.height
                        img_width = 7*inch  # Full width
                        img_height = img_width / img_aspect
                        if img_height > 3.3*inch:
                            img_height = 3.3*inch
                            img_width = img_height * img_aspect
                        property_img = RLImage(main_img_path, width=img_width, height=img_height)
                        other_key_content.append(property_img)
                    else:
                        placeholder = create_placeholder_drawing(img_width, img_height)
                        other_key_content.append(placeholder)
                except Exception as e:
                    print(f"Error loading property image: {e}")
                    placeholder = create_placeholder_drawing(img_width, img_height)
                    other_key_content.append(placeholder)
            else:
                placeholder = create_placeholder_drawing(img_width, img_height)
                other_key_content.append(placeholder)
            
            other_key_content.append(Spacer(1, STANDARD_IMAGE_SPACING))
            
            # EPC Section - Horizontal Layout: Title (left) | Chart (middle) | Details (right)
            epc_title_para = Paragraph("Energy Performance Certificate", 
                ParagraphStyle('EPCTitle', parent=styles['Heading2'], fontSize=14, 
                              textColor=colors.black, fontName='Helvetica-Bold'))
            
            epc_chart = self.create_epc_chart(data, primary_blue, accent_gold, success_green)
            
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
            
            epc_table = Table([[epc_title_para, epc_chart, epc_details_para]], 
                             colWidths=[2*inch, 3.3*inch, 2.3*inch])
            epc_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                ('ALIGN', (2, 0), (2, 0), 'LEFT'),
                ('TOPPADDING', (0, 0), (0, 0), 14),
                ('TOPPADDING', (1, 0), (2, 0), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ]))
            other_key_content.append(epc_table)
            other_key_content.append(Spacer(1, STANDARD_CONTENT_SPACING))
            
            disclaimer_para = Paragraph(
                "This EPC data is accurate up to 6 months ago. If a more recent EPC assessment was done within this period, it will not be displayed here.",
                ParagraphStyle('EPCDisclaimer', parent=styles['Normal'], 
                              fontSize=9, textColor=HexColor('#666666')))
            other_key_content.append(disclaimer_para)
            other_key_content.append(Spacer(1, STANDARD_CONTENT_SPACING))
            
            if data.get('broadband_available'):
                broadband_title_para = Paragraph("Internet / Broadband Availability", 
                    ParagraphStyle('BroadbandTitle', parent=styles['Heading2'], fontSize=14, 
                                  textColor=colors.black, fontName='Helvetica-Bold'))
                
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
                
                broadband_table = Table([
                    [broadband_title_para, '', ''],
                    [broadband_item1, broadband_item2, broadband_item3]
                ], colWidths=[2.3*inch, 2.3*inch, 2.4*inch])
                broadband_table.setStyle(TableStyle([
                    ('SPAN', (0, 0), (-1, 0)),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                    ('ALIGN', (0, 1), (-1, 1), 'LEFT'),
                    ('TOPPADDING', (0, 0), (-1, 0), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('TOPPADDING', (0, 1), (-1, 1), 4),
                    ('BOTTOMPADDING', (0, 1), (-1, 1), 6),
                ]))
                other_key_content.append(broadband_table)
            
            story.append(KeepTogether(other_key_content))
            
            # Floor Plans Section - always show at least one placeholder if no floor plans
            floor_plan_queue = floor_plan_images if floor_plan_images else [None]
            
            if floor_plan_queue:
                story.append(PageBreak())
                story.append(Spacer(1, STANDARD_PAGE_BREAK_SPACING))
                
                story.append(Paragraph("Floor Plans", section_title_style))
                story.append(Spacer(1, STANDARD_SECTION_SPACING))
                
                for i, image_path in enumerate(floor_plan_queue):
                    # For subsequent floor plans, add page break (but not for the first one, since we already added it)
                    if i > 0:
                        story.append(PageBreak())
                        story.append(Spacer(1, STANDARD_PAGE_BREAK_SPACING))
                    
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
            
            # Property Images Gallery - defaults to cover images if gallery is empty
            regular_images = list(property_gallery_images)
            if not regular_images:
                regular_images = list(cover_images)
            if not regular_images:
                regular_images = [None]  # Use None as placeholder marker
            
            if regular_images:
                story.append(PageBreak())
                story.append(Spacer(1, STANDARD_PAGE_BREAK_SPACING))
                story.append(Paragraph("Property Images", section_title_style))
                story.append(Spacer(1, STANDARD_SECTION_SPACING))
                
                image_blocks = []
                for i, image_path in enumerate(regular_images):
                    caption_text = f"Image {i+1}: Placeholder"
                    
                    if image_path is None:
                        img_flowable = create_placeholder_drawing(6*inch, 3.5*inch)
                    else:
                        try:
                            if os.path.exists(image_path):
                                img_flowable = RLImage(image_path, width=6*inch, height=3.5*inch)
                                caption_text = f"Image {i+1}: {os.path.basename(image_path)}"
                            else:
                                img_flowable = create_placeholder_drawing(6*inch, 3.5*inch)
                                caption_text = f"Image {i+1}: File not found"
                        except Exception as e:
                            print(f"Error adding image {image_path}: {e}")
                            img_flowable = create_placeholder_drawing(6*inch, 3.5*inch)
                            caption_text = f"Image {i+1}: Error loading image"
                    
                    block_elements = [
                        Spacer(1, STANDARD_IMAGE_SPACING if i == 0 else STANDARD_CONTENT_SPACING),
                        img_flowable,
                        Spacer(1, STANDARD_CONTENT_SPACING),
                        Paragraph(caption_text, body_style),
                        Spacer(1, STANDARD_CONTENT_SPACING),
                    ]
                    image_blocks.append(block_elements)
                
                for block in image_blocks:
                    story.append(KeepTogether(block))
            
            # Getting To The City Centre and About the City (Last Page)
            story.append(PageBreak())
            story.append(Spacer(1, STANDARD_PAGE_BREAK_SPACING))
            
            # Location Information - Logo, Title, and Directions Image
            story.append(Paragraph("Getting To The City Centre", section_title_style))
            story.append(Spacer(1, STANDARD_SECTION_SPACING))
            
            # Directions Image
            directions_image_path = next(iter(directions_images), None)
            
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
            
            story.append(Spacer(1, STANDARD_IMAGE_SPACING))
            
            # About the City and City Images - keep together on same page
            about_city_content = []
            
            # About the City
            if data.get('about_city'):
                about_city_content.append(Paragraph("About the City", header_style))
                about_city_content.append(Paragraph(f"<b>{data.get('city', 'N/A')}</b>", highlight_style))
                about_city_content.append(Paragraph(data.get('about_city'), body_style))
                about_city_content.append(Paragraph(f"<b>Population:</b> {data.get('population', 'N/A')}", body_style))
                about_city_content.append(Spacer(1, STANDARD_CONTENT_SPACING))
            
            # City Images Section
            city_queue = list(city_images)
            
            # If not provided, fall back to bundled sample images
            if not city_queue:
                sample_paths = [
                    get_resource_path("sample_images/liverpool1.jpg"),
                    get_resource_path("sample_images/liverpool2.jpg"),
                    get_resource_path("sample_images/liverpool3.jpg")
                ]
                for path in sample_paths:
                    if os.path.exists(path):
                        city_queue.append(path)
            
            # Always show City images (placeholders if not found)
            fixed_width = 2.3*inch
            fixed_height = 2.3*inch
            
            # Always show 3 City images (placeholders if missing)
            while len(city_queue) < 3:
                city_queue.append(None)
            
            # Display 3 images horizontally
            img_cells = []
            for img_path in city_queue[:3]:
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
                        print(f"Error loading city image {img_path}: {e}")
                        # Use placeholder on error
                        placeholder = create_placeholder_drawing(fixed_width, fixed_height)
                        img_cells.append(placeholder)
            
            if len(img_cells) == 3:
                city_table = Table([img_cells], colWidths=[fixed_width, fixed_width, fixed_width])
                city_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('TOPPADDING', (0, 0), (-1, -1), 5),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ]))
                about_city_content.append(city_table)
            
            about_city_content.append(Spacer(1, STANDARD_IMAGE_SPACING))
            
            # Wrap in KeepTogether to ensure they stay on same page
            story.append(KeepTogether(about_city_content))
            
            # Build PDF
            doc.build(
                story,
                onFirstPage=self._draw_cover_header,
                onLaterPages=self._draw_standard_header
            )
            
        except Exception as e:
            print(f"Error generating PDF: {e}")
            messagebox.showerror("Error", f"Error generating PDF: {str(e)}")
            
    def create_cover_page(self, data, accent_gold, primary_blue):
        """Create the cover page with logo, property address, main image, thumbnails, and footer"""
        # Always create cover page, even if no images (will use placeholders)
        images_by_section = {key: list(paths) for key, paths in self.image_sections.items()}
        return CoverPageFlowable(data, images_by_section, accent_gold, primary_blue)
            
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