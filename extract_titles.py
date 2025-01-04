import os
from openpyxl import Workbook

def extract_video_titles(directory, output_excel="video_titles.xlsx"):
    if not os.path.isdir(directory):
        print(f"The provided directory '{directory}' does not exist.")
        return
    
    # Video file extensions to filter
    video_extensions = ('.mp4', '.avi', '.mkv', '.mov', '.flv', '.wmv', '.webm')
    
    # Extract video titles
    video_titles = []
    for file in os.listdir(directory):
        if file.lower().endswith(video_extensions):
            video_titles.append(file)
    
    if not video_titles:
        print("No video files found in the directory.")
        return
    
    # Save to Excel
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Video Titles"
    
    # Write headers
    sheet.append(["S.No", "Video Title"])
    
    # Write video titles
    for index, title in enumerate(video_titles, start=1):
        sheet.append([index, title])
    
    # Save the Excel file
    workbook.save(output_excel)
    print(f"Video titles have been saved to '{output_excel}'.")

directory_path = input("Enter the full path to the directory: ")
extract_video_titles(directory_path)
