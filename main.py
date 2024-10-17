import pandas as pd
import requests
import os
from urllib.parse import urlparse, unquote

# Set the path to your Excel file
file_path = "C:/Users/hank.aungkyaw/Documents/Product_Photos.xlsx"

# Read the Excel file into a DataFrame
df = pd.read_excel(file_path)

# Create a folder to save the downloaded images
output_folder = 'downloaded_images'
os.makedirs(output_folder, exist_ok=True)

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    sku = row['Handle']
    url = row['Image Src']

    # Remove query parameters from the URL and get the file extension
    parsed_url = urlparse(url)
    base_url = parsed_url.path
    file_extension = os.path.splitext(base_url)[1]  # Get the file extension (e.g., .png, .jpg)

    # Decode any URL-encoded characters in SKU and remove invalid characters
    valid_sku = ''.join(c for c in unquote(sku) if c.isalnum() or c in ('-', '_'))

    # Set the image file name using the cleaned SKU
    file_name = f"{valid_sku}{file_extension}"
    file_path = os.path.join(output_folder, file_name)

    # Download the image
    try:
        response = requests.get(url, stream=True)
        if response.status_code == 200:
            with open(file_path, 'wb') as file:
                for chunk in response.iter_content(1024):
                    file.write(chunk)
            print(f"Downloaded: {file_name}")
        else:
            print(f"Failed to download {sku} - Status code: {response.status_code}")
    except Exception as e:
        print(f"Error downloading {sku}: {e}")

print("Download completed.")
