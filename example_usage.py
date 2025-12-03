from file_uploader import FileUploader

# Example usage with custom configuration
def upload_images_example():
    uploader = FileUploader(
        api_url="https://api.example.com/process-image",
        folder_path="./images",
        excel_output_path="image_results.xlsx"
    )
    uploader.process_folder()
    uploader.write_to_excel()

# Example with error handling and progress tracking
def upload_with_progress():
    uploader = FileUploader(
        api_url="https://api.example.com/upload",
        folder_path="./documents", 
        excel_output_path="results.xlsx"
    )
    
    total_files = len(list(uploader.folder_path.iterdir()))
    print(f"Found {total_files} files to process")
    
    uploader.process_folder()
    uploader.write_to_excel()
    
    print(f"Processed {len(uploader.results)} files")

if __name__ == "__main__":
    upload_with_progress()