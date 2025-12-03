# File Uploader to API and Excel

A Python program that reads files from a folder, uploads them one-by-one to an API endpoint, and writes the JSON responses to an Excel sheet.

## Features

- Upload files sequentially to any API endpoint
- Handle API responses and errors gracefully
- Export results to Excel with flattened JSON structure
- Track filenames in results

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```python
from file_uploader import FileUploader

uploader = FileUploader(
    api_url="https://worker.formextractorai.com/v2/extract",
    folder_path="./files",
    excel_output_path="results.xlsx",
    access_token="YOUR_ACCESS_TOKEN",
    extractor_id="YOUR_EXTRACTOR_ID"
)
uploader.process_folder()
uploader.write_to_excel()
```

## Configuration

Update the configuration in `file_uploader.py`:
- `ACCESS_TOKEN` - your FormExtractorAI access token
- `EXTRACTOR_ID` - your extractor ID
- `FOLDER_PATH` - folder containing files to upload
- `EXCEL_OUTPUT` - output Excel file path