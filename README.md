# SharePoint Flask App

This project contains the uploaded Flask application `app.py`.

## Files
- `app.py` - Flask app with Excel upload, processing, and optional email sending.
- `requirements.txt` - Python dependencies for the app.
- `render.yaml` - Render deployment configuration.

## Setup
1. Open PowerShell and navigate to the project folder:
   ```powershell
   cd "C:\Users\משתמש\Desktop\sharepoint-flask-app"
   ```
2. Create a virtual environment:
   ```powershell
   python -m venv venv
   .\venv\Scripts\Activate.ps1
   ```
3. Install dependencies:
   ```powershell
   pip install -r requirements.txt
   ```
4. Set Gmail credentials in environment variables:
   ```powershell
   $env:GMAIL_USER = "your@gmail.com"
   $env:GMAIL_PASS = "your-app-password"
   ```
5. Run the app:
   ```powershell
   python app.py
   ```
6. Open http://localhost:5000 in your browser.

## Render deployment
1. Push this folder to a GitHub repository.
2. Create a service in Render and connect it to that repository.
3. Ensure `render.yaml` exists at the repo root.
4. Assign the Render Env Group that contains `GMAIL_USER` and `GMAIL_PASS`.

## Notes
- The app loads an Excel sheet named `Export` from the uploaded file.
- It generates XLSX files and can send them by email if addresses are provided.
- For deployment to cloud, set `PORT` environment variable and use a hosted Python service.
- If Git is not installed locally, install Git first and then run:
  ```powershell
  git init
  git add .
  git commit -m "Initial deploy-ready Flask app"
  ```
