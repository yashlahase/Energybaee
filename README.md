# Solar Load Calculator Automation (ENERGYBAE)

An end-to-end AI automation tool that processes electricity bills and generates a structured Excel output for solar load calculation.

## Features
- **AI-Powered Extraction**: Uses Gemini Vision to extract structured data directly from bill images.
- **Excel Automation**: Populates a predefined template while **preserving all formulas**.
- **Modern UI**: A premium, responsive web interface built with Flask and Vanilla CSS (Glassmorphism).
- **Deployment Ready**: Fully optimized for Vercel deployment.

## Tech Stack
- **Backend**: Python (Flask)
- **AI**: Google Generative AI (Gemini 1.5 Flash)
- **Excel Handling**: openpyxl
- **Styling**: Vanilla CSS with modern aesthetics

## Setup & Local Running

1. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Environment Variables**:
   Create a `.env` file in the root directory:
   ```env
   GEMINI_API_KEY=your_key_here
   FLASK_SECRET_KEY=your_random_secret
   ```

3. **Initialize Template**:
   ```bash
   python3 create_template.py
   ```

4. **Run the App**:
   ```bash
   python3 app.py
   ```
   Open `http://localhost:5001` in your browser.

## Vercel Deployment
The project includes a `vercel.json` for easy deployment.
1. Run `vercel`.
2. Add your `GEMINI_API_KEY` and `FLASK_SECRET_KEY` to the Vercel dashboard.
3. Run `vercel --prod`.

## Directory Structure
- `app.py`: Main Flask server.
- `extractor.py`: Gemini Vision extraction logic.
- `excel_handler.py`: Logic to fill the Excel template.
- `assets/`: Contains the Excel template.
- `static/`: CSS and temporary files.
- `templates/`: Main UI (index.html).
