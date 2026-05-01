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
    model_name = 'gemini-flash-latest'
    model = genai.GenerativeModel(model_name)
    print(f"--- Extractor Initialized ---")
    print(f"Model: {model_name}")
    print(f"API Key (masked): {api_key[:5]}...{api_key[-4:]}")
else:
    model = None
    print("--- Extractor Warning: No API Key found ---")

class BillExtractor:
    def get_structured_data(self, file_bytes, mime_type):
        """Use Gemini Vision to extract data directly from file bytes"""
        if not model:
            return {"error": "API Key missing"}

        # Validate mime_type
        valid_mimes = ["application/pdf", "image/jpeg", "image/png", "image/webp", "image/heic", "image/heif"]
        if mime_type not in valid_mimes:
            # Fallback based on common extensions if possible, or just try
            if "pdf" in mime_type.lower(): mime_type = "application/pdf"
            elif "jpg" in mime_type.lower() or "jpeg" in mime_type.lower(): mime_type = "image/jpeg"
            elif "png" in mime_type.lower(): mime_type = "image/png"
            else: mime_type = "image/jpeg" # Default to image

        prompt = """
        Extract the following information from this electricity bill. 
        Focus on extracting the consumption history for the last 12-14 months if available.
        Return ONLY a valid JSON object.
        
        JSON Structure:
        {
            "consumer_name": "Full name",
            "consumer_number": "Consumer number",
            "fixed_charges": 130.0,
            "sanctioned_load": "3.3 kW",
            "connection_type": "LT I Res 1-Phase",
            "contract_demand": "null",
            "consumption_history": [
                {"month": "January 2026", "units": 100, "bill_amount": 500, "unit_cost": 5.0},
                ...
            ]
        }

        Rules:
        1. Consumption history should be ordered from oldest to newest.
        2. Numbers must be numbers (no symbols).
        3. Use null for missing values.
        """

        try:
            # Using the recommended list format for multi-modal
            response = model.generate_content([
                {'mime_type': mime_type, 'data': file_bytes},
                prompt
            ])
            
            if not response.text:
                print(f"Empty response from Gemini. Safety ratings: {response.candidates[0].safety_ratings}")
                return {"error": "Gemini returned an empty response (possibly blocked by safety filters)"}

            content = response.text
            print(f"Raw Gemini Response: {content[:200]}...") # Log first 200 chars
            
            # Extract JSON from potential markdown
            json_str = content
            if "```json" in content:
                json_str = content.split("```json")[1].split("```")[0]
            elif "```" in content:
                json_str = content.split("```")[1].split("```")[0]
            
            return json.loads(json_str.strip())
        except Exception as e:
            print(f"Extraction Error: {str(e)}")
            return {"error": f"Extraction failed: {str(e)}"}

def process_bill(file_path, mime_type):
    extractor = BillExtractor()
    with open(file_path, "rb") as f:
        file_bytes = f.read()
    return extractor.get_structured_data(file_bytes, mime_type)
