import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import os
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
import datetime

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
        
        # Images Section
        images_frame = ttk.LabelFrame(self.images_frame, text="Property Images", padding="10")
        images_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
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
        
    def add_images(self):
        """Add images to the list"""
        file_paths = filedialog.askopenfilenames(
            title="Select Property Images",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.bmp *.tiff")]
        )
        
        for file_path in file_paths:
            self.images.append(file_path)
            filename = os.path.basename(file_path)
            self.image_listbox.insert(tk.END, filename)
            
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
        sample_folder = "sample_images"
        if os.path.exists(sample_folder):
            self.load_images_from_folder(sample_folder)
        else:
            messagebox.showwarning("Warning", "Sample images folder not found. Please create sample images first.")
            
    def load_images_from_folder(self, folder_path):
        """Load all images from a specified folder"""
        image_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff')
        
        try:
            print(f"Loading images from folder: {folder_path}")
            files = os.listdir(folder_path)
            print(f"Found {len(files)} files in folder")
            
            for filename in files:
                if filename.lower().endswith(image_extensions):
                    image_path = os.path.join(folder_path, filename)
                    self.images.append(image_path)
                    self.image_listbox.insert(tk.END, filename)
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
            
            # Create header with branding
            story.append(self.create_header(data, accent_gold, primary_blue))
            story.append(Spacer(1, 20))
            
            # Property Header Section
            property_header = f"{data.get('address', '')}, {data.get('postal_code', '')}"
            story.append(Paragraph(property_header, header_style))
            
            # Asking Price in highlighted box
            if data.get('asking_price'):
                asking_price_text = f"<b>Asking Price:</b> {data.get('asking_price')}"
                story.append(Paragraph(asking_price_text, highlight_style))
            
            # Report Date
            report_date = datetime.datetime.now().strftime("%dth %B %Y")
            story.append(Paragraph(f"Report created on {report_date}", body_style))
            story.append(Spacer(1, 25))
            
            # Investment Opportunity Section
            story.append(Paragraph("Investment Opportunity", header_style))
            
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
            
            # Investment Summary Table with professional styling
            investment_data = [
                ['Purchase Price', f"£{purchase_price:,.0f}"],
                ['Total Purchase Costs', ''],
                ['Deposit (20%)', f"£{deposit_amount:,.0f}"],
                ['Stamp Duty', f"£{stamp_duty:,.0f}"],
                ['Survey', f"£{survey_cost:,.0f}"],
                ['Legal Fees', f"£{legal_fees:,.0f}"],
                ['Loan Set-up', f"£{loan_setup:,.0f}"],
                ['Total Investment Required', f"£{total_investment:,.0f}"],
                ['Estimated Monthly Rent', f"£{monthly_rent:,.0f}pcm"],
                ['Rental Yield', f"{rental_yield:.1f}%"]
            ]
            
            investment_table = Table(investment_data, colWidths=[3*inch, 2*inch])
            investment_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), accent_gold),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('TOPPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), light_grey),
                ('GRID', (0, 0), (-1, -1), 1, primary_blue),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 11),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [light_grey, colors.white])
            ]))
            
            story.append(investment_table)
            story.append(Spacer(1, 20))
            
            # Annual Expenses Table
            expenses_data = [
                ['Total Annual Expenses', ''],
                ['Mortgage @ 5.8% (Interest Only)', f"£{annual_mortgage_interest:,.0f}"],
                ['Council Tax', f"£{council_tax:,.0f}"],
                ['Repairs / Maintenance', f"£{repairs:,.0f}"],
                ['Electric / Gas', f"£{utilities:,.0f}"],
                ['Water', f"£{water:,.0f}"],
                ['Broadband / TV', f"£{broadband:,.0f}"],
                ['Insurance', f"£{insurance:,.0f}"],
                ['Total', f"£{total_annual_expenses:,.0f}"],
                ['Monthly Profit', f"£{monthly_profit:,.0f}"],
                ['Annual Profit', f"£{annual_profit:,.0f}"],
                ['ROI', f"{roi:.1f}%"]
            ]
            
            expenses_table = Table(expenses_data, colWidths=[3*inch, 2*inch])
            expenses_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), success_green),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('TOPPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), light_grey),
                ('GRID', (0, 0), (-1, -1), 1, primary_blue),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 11),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [light_grey, colors.white])
            ]))
            
            story.append(expenses_table)
            story.append(Spacer(1, 25))
            
            # Key Information Section
            story.append(Paragraph("Key Information", header_style))
            
            key_info_data = [
                ['Asking Price', data.get('asking_price', 'N/A')],
                ['Bedrooms', data.get('bedrooms', 'N/A')],
                ['Size', f"{data.get('size_sqm', 'N/A')} sqm"],
                ['On the market for', f"{data.get('days_on_market', 'N/A')} days"]
            ]
            
            key_info_table = Table(key_info_data, colWidths=[2*inch, 3*inch])
            key_info_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), primary_blue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('TOPPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), light_grey),
                ('GRID', (0, 0), (-1, -1), 1, primary_blue),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 11),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [light_grey, colors.white])
            ]))
            
            story.append(key_info_table)
            story.append(Spacer(1, 15))
            
            # Key Features
            if data.get('key_features'):
                story.append(Paragraph("Key Features", subheader_style))
                story.append(Paragraph(data.get('key_features'), body_style))
                story.append(Spacer(1, 15))
            
            # EPC Information with visual chart
            story.append(Paragraph("Energy Performance Certificate", header_style))
            
            # Create EPC visual chart
            epc_chart = self.create_epc_chart(data, primary_blue, accent_gold, success_green)
            story.append(epc_chart)
            story.append(Spacer(1, 15))
            
            # EPC Details Table
            epc_data = [
                ['Latest Inspection Date', data.get('inspection_date', 'N/A')],
                ['Window Glazing', data.get('window_glazing', 'N/A')],
                ['Building Construction Age', data.get('building_age', 'N/A')]
            ]
            
            epc_table = Table(epc_data, colWidths=[2*inch, 3*inch])
            epc_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), accent_gold),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('TOPPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), light_grey),
                ('GRID', (0, 0), (-1, -1), 1, primary_blue),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 11),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [light_grey, colors.white])
            ]))
            
            story.append(epc_table)
            story.append(Spacer(1, 15))
            
            # Internet/Broadband
            if data.get('broadband_available'):
                story.append(Paragraph("Internet / Broadband Availability", subheader_style))
                story.append(Paragraph(f"<b>Broadband available:</b> {data.get('broadband_available')}", body_style))
                story.append(Paragraph(f"<b>Highest available download speed:</b> {data.get('download_speed', 'N/A')}", body_style))
                story.append(Paragraph(f"<b>Highest available upload speed:</b> {data.get('upload_speed', 'N/A')}", body_style))
                story.append(Spacer(1, 15))
            
            # Location Information
            story.append(Paragraph("Getting To The City Centre", header_style))
            
            location_data = [
                ['By Car', f"{data.get('time_car', 'N/A')} minutes", f"{data.get('distance_city_centre', 'N/A')} miles"],
                ['By Public Transport', f"{data.get('time_public_transport', 'N/A')} minutes", f"{data.get('walk_to_station', 'N/A')} mins walk ({data.get('station_distance', 'N/A')} mi)"],
                ['Bus Routes', data.get('bus_routes', 'N/A'), f"Every {data.get('bus_frequency', 'N/A')}"]
            ]
            
            location_table = Table(location_data, colWidths=[2*inch, 1.5*inch, 2.5*inch])
            location_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), success_green),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('TOPPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), light_grey),
                ('GRID', (0, 0), (-1, -1), 1, primary_blue),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 11),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [light_grey, colors.white])
            ]))
            
            story.append(location_table)
            story.append(Spacer(1, 15))
            
            # About the City
            if data.get('about_city'):
                story.append(Paragraph("About the City", header_style))
                story.append(Paragraph(f"<b>{data.get('city', 'N/A')}</b>", highlight_style))
                story.append(Paragraph(data.get('about_city'), body_style))
                story.append(Paragraph(f"<b>Population:</b> {data.get('population', 'N/A')}", body_style))
                story.append(Spacer(1, 15))
            
            # Property Images
            if self.images:
                story.append(PageBreak())
                story.append(Paragraph("Property Images", header_style))
                
                for i, image_path in enumerate(self.images):
                    try:
                        # Add image with caption
                        img = RLImage(image_path, width=6*inch, height=4*inch)
                        story.append(img)
                        story.append(Spacer(1, 12))
                        
                        # Add caption
                        filename = os.path.basename(image_path)
                        story.append(Paragraph(f"Image {i+1}: {filename}", body_style))
                        story.append(Spacer(1, 20))
                        
                    except Exception as e:
                        print(f"Error adding image {image_path}: {e}")
                        story.append(Paragraph(f"Image {i+1}: Error loading image", body_style))
                        story.append(Spacer(1, 20))
            
            # Build PDF
            doc.build(story)
            
            messagebox.showinfo("Success", f"Professional investment report PDF generated successfully!\nSaved as: {file_path}")
            
        except Exception as e:
            print(f"Error generating PDF: {e}")
            messagebox.showerror("Error", f"Error generating PDF: {str(e)}")
            
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
        """Create a visual EPC rating chart"""
        drawing = Drawing(7*inch, 3*inch)
        
        # EPC bands with colors
        epc_bands = [
            ('A', 92, 100, success_green),
            ('B', 81, 91, HexColor('#22c55e')),
            ('C', 69, 80, HexColor('#84cc16')),
            ('D', 55, 68, HexColor('#eab308')),
            ('E', 39, 54, HexColor('#f59e0b')),
            ('F', 21, 38, HexColor('#ef4444')),
            ('G', 1, 20, HexColor('#dc2626'))
        ]
        
        # Create vertical bars for EPC chart
        bar_width = 0.8*inch
        bar_height = 2*inch
        start_x = 1*inch
        start_y = 0.5*inch
        
        for i, (grade, min_score, max_score, color) in enumerate(epc_bands):
            x = start_x + (i * bar_width)
            
            # Bar background
            drawing.add(Rect(x, start_y, bar_width, bar_height, 
                           fillColor=color, strokeColor=primary_blue, strokeWidth=1))
            
            # Grade label
            drawing.add(String(x + bar_width/2, start_y - 0.2*inch, grade, 
                             fontName="Helvetica-Bold", fontSize=14, 
                             fillColor=primary_blue, textAnchor="middle"))
            
            # Score range
            drawing.add(String(x + bar_width/2, start_y - 0.4*inch, f"({min_score}+)", 
                             fontName="Helvetica", fontSize=8, 
                             fillColor=primary_blue, textAnchor="middle"))
        
        # Add current and potential scores
        current_score = int(data.get('current_rating', '72'))
        potential_score = int(data.get('potential_rating', '84'))
        
        # Find which band the scores fall into
        current_band = next((i for i, (_, min_score, max_score, _) in enumerate(epc_bands) 
                           if min_score <= current_score <= max_score), 2)
        potential_band = next((i for i, (_, min_score, max_score, _) in enumerate(epc_bands) 
                             if min_score <= potential_score <= max_score), 1)
        
        # Highlight current score
        current_x = start_x + (current_band * bar_width)
        drawing.add(String(current_x + bar_width/2, start_y + bar_height + 0.1*inch, 
                         f"Current: {current_score}", 
                         fontName="Helvetica-Bold", fontSize=12, 
                         fillColor=primary_blue, textAnchor="middle"))
        
        # Highlight potential score
        potential_x = start_x + (potential_band * bar_width)
        drawing.add(String(potential_x + bar_width/2, start_y + bar_height + 0.3*inch, 
                         f"Potential: {potential_score}", 
                         fontName="Helvetica-Bold", fontSize=12, 
                         fillColor=success_green, textAnchor="middle"))
        
        return drawing

def main():
    root = tk.Tk()
    app = PDFBuilderApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()