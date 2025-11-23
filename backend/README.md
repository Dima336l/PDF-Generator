# PDF-Generator-Backend

Minimal Node/Express API that wraps the existing `pdf-generator.js` to produce the exact same PDF as the desktop/Electron app.

## Endpoints

- POST `/generate`  
  Body (JSON):
  ```json
  {
    "data": { "... form fields ..." },
    "images": {
      "cover": ["data:image/png;base64,..."],
      "property": ["data:image/jpeg;base64,..."],
      "floor_plans": [],
      "directions": [],
      "city": []
    },
    "logo_base64": "data:image/png;base64,..."
  }
  ```
  Returns: `application/pdf` stream (attachment).

Notes:
- Image arrays accept base64 data URLs. They are written to temp files and passed to the existing generator.
- CORS is enabled for all origins by default; scope it as needed.

## Local run

```bash
cd backend
npm install
npm start
# POST http://localhost:8080/generate
```

## Deploy (Render)

- Create a new Web Service, Node environment.
- Build command: `npm install`
- Start command: `npm start`
- Expose port `8080`.
- After deploy, POST to `https://<your-service>.onrender.com/generate`.


