# Property Investment Report Builder - Frontend

A modern web application for generating professional property investment reports with detailed calculations, location data, and visual content.

## Features

- **Multi-Calculator Support**: Select and configure multiple investment calculators:
  - Standard Buy to Let
  - Buy Refurbish Refinance (BRR)
  - Flip
  - Holiday Let
  - Rent to HMO
  - Rent to Serviced Accommodation
  - Purchase

- **Automatic Location Data**: Fetch location information with one click:
  - City information and population
  - Nearest train station and distance
  - Nearest school and distance
  - Local amenities
  - City map generation
  - City-specific images

- **Image Management**: Organize property images across multiple categories:
  - Cover page images
  - Property gallery
  - Floor plans
  - City map
  - City lifestyle images

- **Property Information Forms**: Comprehensive input for:
  - Property details (address, type, bedrooms, bathrooms, size, price)
  - Investment calculations (multiple calculator types)
  - EPC ratings and energy costs
  - Broadband information

- **Professional PDF Generation**: Generate detailed investment reports with:
  - Property information pages
  - Investment opportunity calculations with charts
  - Key information and images
  - EPC and broadband details
  - City map and lifestyle images

## Quick Start

### Prerequisites

- Node.js (v14 or higher)
- npm (comes with Node.js)

### Local Development

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd Property-PDF
   ```

2. **Install dependencies**:
   ```bash
   npm install
   ```

3. **Start the development server**:
   ```bash
   npm start
   ```
   
   Or use the batch script (Windows):
   ```cmd
   start-local-dev.bat
   ```

4. **Open your browser**:
   Navigate to `http://localhost:3000`

### Running with Backend

The frontend requires a backend API for PDF generation. The backend should be running on `http://localhost:8080` (or configured via environment variables).

To run both frontend and backend together:
```bash
npm run start:dev
```

This will start:
- Frontend web server on port 3000
- Backend API server on port 8080

## Project Structure

```
Property-PDF/
├── index.html              # Main HTML file
├── renderer.js             # Frontend JavaScript logic
├── styles.css              # Application styles
├── calculator-logic.js     # Calculator logic (for reference)
├── web-server.js           # Express web server
├── logo.png                # Application logo
├── sample_images/          # Sample property images
├── package.json            # Dependencies and scripts
├── start-local-dev.bat     # Windows development script
└── .github/
    └── workflows/
        └── pages.yml       # GitHub Pages deployment
```

## Available Scripts

- `npm start` - Start the web server (port 3000)
- `npm run start:dev` - Start both frontend and backend (requires backend folder)
- `npm run start:electron` - Start Electron app (if needed)

## Features in Detail

### Calculator Selection

Select one or more investment calculators from the Investment tab. Each calculator has its own set of input fields that are dynamically generated based on your selection.

### Location Data Fetching

Click "Fetch Location Data" in the Location tab to automatically populate:
- City name
- Distance to city centre
- Nearest train station and distance
- Nearest school and distance
- Local amenities
- City description and population
- City map (generated automatically)
- City-specific images (fetched from Pexels API)

### Image Management

Add images to different sections:
- **Cover Page**: First image appears as hero photo, others as thumbnails
- **Property Gallery**: General property photos
- **Floor Plans**: Dedicated floor plan pages
- **City Map**: Map of the capital city
- **City Images**: Urban lifestyle shots

Images can be reordered using "Move Up" and "Move Down" buttons.

### PDF Generation

Click "Generate Investment Report PDF" to create a comprehensive PDF report including:
- Cover page with property images
- Property information
- Investment opportunity calculations (for each selected calculator)
- Key information pages
- EPC and broadband details
- City map and lifestyle images

## Configuration

### Backend URL

The frontend automatically detects the backend URL:
- **Local development**: Uses `http://localhost:8080` when running locally
- **Production**: Uses the production backend URL (configured in `renderer.js`)

To change the backend URL, edit the `BACKEND_URL` constant in `renderer.js`.

### Environment Variables

For local development, you can set:
- `PORT` - Web server port (default: 3000)
- `BACKEND_URL` - Backend API URL (default: auto-detected)

## Deployment

### GitHub Pages

The application is automatically deployed to GitHub Pages on push to the `main` branch via GitHub Actions.

The workflow:
1. Builds the static site
2. Copies necessary files to `public/` directory
3. Deploys to GitHub Pages

### Manual Deployment

1. Build the static files:
   ```bash
   npm run build
   ```

2. Deploy the `public/` directory to your hosting service

## Technologies Used

- **Frontend**: HTML5, CSS3, Vanilla JavaScript
- **Web Server**: Express.js
- **PDF Generation**: Backend API (separate repository)
- **Image Processing**: Browser APIs, Backend proxy
- **Location Services**: 
  - OpenStreetMap Nominatim (geocoding)
  - Overpass API (stations, schools, amenities)
  - Wikipedia API (city information)
  - Pexels API (city images)

## Browser Support

- Chrome (latest)
- Firefox (latest)
- Safari (latest)
- Edge (latest)

## Troubleshooting

### Port Already in Use

If you see `EADDRINUSE` errors:
- Windows: The `start-local-dev.bat` script automatically kills processes on ports 3000 and 8080
- Manual: Kill the process using the port:
  ```bash
  # Windows
  netstat -ano | findstr :3000
  taskkill /PID <PID> /F
  ```

### Images Not Loading

- Check browser console for CORS errors
- Ensure backend proxy is running for external images
- Verify image URLs are valid

### Calculator Fields Not Showing

- Ensure JavaScript is enabled
- Check browser console for errors
- Verify calculator selection checkboxes are checked

### PDF Generation Fails

- Ensure backend API is running and accessible
- Check backend URL configuration
- Verify all required fields are filled
- Check browser console and network tab for errors

## Development

### Adding New Calculators

1. Add calculator configuration to `calculatorConfigs` in `renderer.js`
2. Add calculation logic to backend `calculator-logic.js`
3. Update PDF generator to handle new calculator type

### Styling

Styles are in `styles.css`. The application uses:
- Flexbox for layouts
- CSS Grid for form fields
- Responsive design for mobile devices

## License

This project is open source and available under the MIT License.

## Related Repositories

- **Backend**: [PDF-Generator-Backend](https://github.com/Dima336l/PDF-Generator-Backend) - Node.js/Express API for PDF generation

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## Support

For issues and questions, please open an issue on GitHub.
