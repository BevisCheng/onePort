import os
import requests
import pandas as pd
from pathlib import Path
import json
import mimetypes
import re

class FileUploader:
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
            # Determine content type based on file extension
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
        """Process all files in the folder"""
        for file_path in self.folder_path.iterdir():
            if file_path.is_file():
                print(f"Uploading: {file_path.name}")
                json_response = self.upload_file(file_path)
                
                # Add filename to response for tracking
                json_response['filename'] = file_path.name
                self.results.append(json_response)
    
    def write_to_excel(self):
        """Write all results to Excel sheet with shipping document columns"""
        processed_data = []
        
        for result in self.results:
            # Clean filename by removing invoice patterns
            filename = result.get('filename', '')
            cleaned_filename = re.sub(r'\s*\([^)]*INVOICE[^)]*\)', '', filename, flags=re.IGNORECASE)
            cleaned_filename = re.sub(r'\s*（[^）]*INVOICE[^）]*）', '', cleaned_filename, flags=re.IGNORECASE)
            cleaned_filename = cleaned_filename.strip()
            
            row = {'filename': cleaned_filename}
            
            # Extract data from correct nested structure
            data = {}
            if 'documents' in result and len(result['documents']) > 0:
                data = result['documents'][0].get('data', {})
            
            # Map API fields to Excel columns
            row['Shipper Company Name'] = data.get('shipper_company_name', '')
            row['Shipper Details'] = data.get('shipper_details', '')
            row['Consignee Company Name'] = data.get('consignee_company_name', '')
            row['Consignee Details'] = data.get('consignee_details', '')
            row['Consignee Country'] = data.get('consignee_country', '')
            row['Notify Party Company Name'] = data.get('notify_party_company_name', '')
            row['Notify Party Details'] = data.get('notify_party_details', '')
            row['Also Notify Company Name'] = data.get('also_notify_company_name', '')
            row['Also Notify Details'] = data.get('also_notify_details', '')
            row['B/L Type'] = data.get('bl_type', '')
            row['B/L Number'] = data.get('bl_number', '')
            row['Vessel Name'] = data.get('vessel_name', '')
            row['Voyage Number'] = data.get('voyage_no', '')
            row['Port of Loading'] = data.get('port_of_loading_raw_text', '')
            row['Port of Discharge'] = data.get('port_of_discharge_raw_text', '')
            row['Place of Receipt'] = data.get('place_of_receipt', '')
            row['Place of Delivery'] = data.get('place_of_delivery', '')
            
            # Extract cargo details from marks_and_descriptions_table
            cargo_table = data.get('marks_and_descriptions_table', [])
            if cargo_table and len(cargo_table) > 0:
                cargo = cargo_table[0]  # Take first cargo entry
                row['Cargo Description'] = cargo.get('description', '')
                row['Package Count'] = cargo.get('no_of_package', '')
                row['Package Unit'] = cargo.get('package_unit', '')
                row['Package Raw Text'] = cargo.get('package_raw_text', '')
                row['Gross Weight'] = cargo.get('gross_weight', '')
                row['Gross Weight Unit'] = cargo.get('gross_weight_unit', '')
                row['Gross Weight Raw'] = cargo.get('gross_weight_raw_text', '')
                row['Measurement'] = cargo.get('measurement', '')
                row['Measurement Unit'] = cargo.get('measurement_unit', '')
                row['Measurement Raw'] = cargo.get('measurement_raw_text', '')
            else:
                # Empty cargo fields if no data
                cargo_fields = ['Cargo Description', 'Package Count', 'Package Unit', 'Package Raw Text',
                              'Gross Weight', 'Gross Weight Unit', 'Gross Weight Raw',
                              'Measurement', 'Measurement Unit', 'Measurement Raw']
                for field in cargo_fields:
                    row[field] = ''
            
            # Add error field if present
            if 'error' in result:
                row['Error'] = result['error']
            else:
                row['Error'] = ''
            
            processed_data.append(row)
        
        # Create DataFrame and save to Excel
        df = pd.DataFrame(processed_data)
        df.to_excel(self.excel_output_path, index=False)
        print(f"Results written to: {self.excel_output_path}")

def main():
    # Configuration
    API_URL = "https://worker.formextractorai.com/v2/extract"
    FOLDER_PATH = "files_to_upload"  # Replace with your folder path
    EXCEL_OUTPUT = "shipping_data_cleaned.xlsx"
    ACCESS_TOKEN = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJyZXNvdXJjZV9vd25lcl9pZCI6IjJkMWNlMzZkLWM3NTAtNDc4OS1hOTVjLTYxOWQwMTdhYTZlMSIsIndvcmtlcl90b2tlbl9pZCI6IjhlZTZkNzJjLWRlYzItNDAwNi05MDQyLTg5ZTRjNzRjMTQyNiJ9.qia0zCgdK-fx1YadnvGcnt-6813oPTt3z4_C8M1UZG8"
    EXTRACTOR_ID = "023fe11d-ef3e-4ae5-b75a-2efab8022c2d"
    
    uploader = FileUploader(API_URL, FOLDER_PATH, EXCEL_OUTPUT, ACCESS_TOKEN, EXTRACTOR_ID)
    uploader.process_folder()
    uploader.write_to_excel()

if __name__ == "__main__":
    main()