import os
import requests
import pandas as pd
from pathlib import Path
import json
import mimetypes
import re

class InvoiceUploader:
    def __init__(self, api_url, folder_path, excel_output_path, access_token, extractor_id):
        self.api_url = api_url
        self.folder_path = Path(folder_path)
        self.excel_output_path = excel_output_path
        self.access_token = access_token
        self.extractor_id = extractor_id
        self.results = []
    
    def upload_file(self, file_path):
        """Upload a single file to the API and return JSON response"""
        try:
            content_type, _ = mimetypes.guess_type(file_path)
            if not content_type:
                content_type = 'application/octet-stream'
            
            with open(file_path, 'rb') as file:
                headers = {
                    'X-WORKER-TOKEN': self.access_token,
                    'X-WORKER-EXTRACTOR-ID': self.extractor_id,
                    'Content-Type': content_type
                }
                response = requests.post(self.api_url, headers=headers, data=file)
                response.raise_for_status()
                return response.json()
        except Exception as e:
            return {"error": str(e), "file": str(file_path)}
    
    def process_folder(self):
        """Process only files with INVOICE in filename"""
        for file_path in self.folder_path.iterdir():
            if file_path.is_file() and 'invoice' in file_path.name.lower():
                print(f"Uploading: {file_path.name}")
                json_response = self.upload_file(file_path)
                json_response['filename'] = file_path.name
                self.results.append(json_response)
    
    def write_to_excel(self):
        """Write all results to Excel sheet with invoice columns"""
        processed_data = []
        
        for result in self.results:
            # Extract data from correct nested structure
            data = {}
            if 'documents' in result and len(result['documents']) > 0:
                data = result['documents'][0].get('data', {})
            
            row = {
                'filename': result.get('filename', ''),
                'B/L Number': data.get('bl_number', ''),
                'Invoice Number': data.get('invoice_number', ''),
                'Freight Charges Item/Charge Description': data.get('freight_charges_description', ''),
                'Currency': data.get('currency', ''),
                'Amount': data.get('amount', ''),
                'Exchange Rate': data.get('exchange_rate', ''),
                'Total Payment Amount': data.get('total_payment_amount', ''),
                'Error': result.get('error', '')
            }
            
            processed_data.append(row)
        
        df = pd.DataFrame(processed_data)
        df.to_excel(self.excel_output_path, index=False)
        print(f"Invoice results written to: {self.excel_output_path}")

def main():
    # Configuration - UPDATE WITH NEW CREDENTIALS
    API_URL = "https://worker.formextractorai.com/v2/extract"
    FOLDER_PATH = "files_to_upload"
    EXCEL_OUTPUT = "invoice_data.xlsx"
    ACCESS_TOKEN = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJyZXNvdXJjZV9vd25lcl9pZCI6IjJkMWNlMzZkLWM3NTAtNDc4OS1hOTVjLTYxOWQwMTdhYTZlMSIsIndvcmtlcl90b2tlbl9pZCI6IjhlZTZkNzJjLWRlYzItNDAwNi05MDQyLTg5ZTRjNzRjMTQyNiJ9.qia0zCgdK-fx1YadnvGcnt-6813oPTt3z4_C8M1UZG8"
    EXTRACTOR_ID = "1f3d0e72-6407-4b12-aeb8-3f3eb42ab6e7"
    
    uploader = InvoiceUploader(API_URL, FOLDER_PATH, EXCEL_OUTPUT, ACCESS_TOKEN, EXTRACTOR_ID)
    uploader.process_folder()
    uploader.write_to_excel()

if __name__ == "__main__":
    main()