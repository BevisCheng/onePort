import os
import requests
import pandas as pd
from pathlib import Path
import json

class FileUploader:
    def __init__(self, api_url, folder_path, excel_output_path):
        self.api_url = api_url
        self.folder_path = Path(folder_path)
        self.excel_output_path = excel_output_path
        self.results = []
    
    def upload_file(self, file_path):
        """Upload a single file to the API and return JSON response"""
        try:
            with open(file_path, 'rb') as file:
                files = {'file': file}
                response = requests.post(self.api_url, files=files)
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
        """Write all results to Excel sheet"""
        df = pd.json_normalize(self.results)
        df.to_excel(self.excel_output_path, index=False)
        print(f"Results written to: {self.excel_output_path}")

def main():
    # Configuration
    API_URL = "https://your-api-endpoint.com/upload"  # Replace with your API URL
    FOLDER_PATH = "files_to_upload"  # Replace with your folder path
    EXCEL_OUTPUT = "upload_results.xlsx"
    
    uploader = FileUploader(API_URL, FOLDER_PATH, EXCEL_OUTPUT)
    uploader.process_folder()
    uploader.write_to_excel()

if __name__ == "__main__":
    main()