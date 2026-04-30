import os
import json
import google.generativeai as genai
from dotenv import load_dotenv
import PIL.Image
import io

load_dotenv()

# Configure Gemini
api_key = os.getenv("GEMINI_API_KEY")
if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-flash-latest')
else:
    model = None

class BillExtractor:
    def get_structured_data(self, image_bytes):
        """Use Gemini Vision to extract data directly from image bytes"""
        if not model:
            return {"error": "API Key missing"}

        # Convert bytes to PIL Image
        img = PIL.Image.open(io.BytesIO(image_bytes))

        prompt = """
        Extract the following information from this electricity bill image. 
        Return the data strictly in JSON format.
        
        Fields:
        - Consumer Name
        - Consumer Number
        - Billing Date (YYYY-MM-DD)
        - Billing Period
        - Units Consumed (kWh) (Number only)
        - Sanctioned Load (kW) (Number only)
        - Connected Load (kW) (Number only, null if not available)
        - Tariff Category
        - Total Bill Amount (Number only)
        - Due Date (YYYY-MM-DD)
        - Meter Number (null if not available)

        Rules:
        1. Dates must be in YYYY-MM-DD format.
        2. Numbers must be clean (no symbols like ₹, $, commas).
        3. Use null for missing values.
        4. Do NOT guess values.
        5. Return ONLY a valid JSON object.
        """

        try:
            response = model.generate_content([prompt, img])
            content = response.text
            # Clean up markdown
            if "```json" in content:
                content = content.split("```json")[1].split("```")[0]
            elif "```" in content:
                content = content.split("```")[1].split("```")[0]
            
            return json.loads(content.strip())
        except Exception as e:
            return {"error": f"Extraction failed: {str(e)}"}

def process_bill(file_path):
    extractor = BillExtractor()
    with open(file_path, "rb") as f:
        image_bytes = f.read()
    return extractor.get_structured_data(image_bytes)
