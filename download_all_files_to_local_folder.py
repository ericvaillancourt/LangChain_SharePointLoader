import traceback
from sharepoint_api import SharePointClient
from dotenv import load_dotenv
import os


# Load environment variables
load_dotenv()

# Configuration of environment keys
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
tenant_id = os.getenv('TENANT_ID')
client_id = os.getenv('CLIENT_ID')
client_secret = os.getenv('CLIENT_SECRET')
site_url = os.getenv('SITE_URL')
resource = os.getenv('RESOURCE')

# Initialize SharePoint client
client = SharePointClient(tenant_id, client_id, client_secret, resource)

# Get site ID
try:
    site_id = client.get_site_id(site_url)
except Exception as e:
    print(f"Error getting site ID: {e}")
    exit(1)

# Get drive information
try:
    drive_info = client.get_drive_id(site_id)
    drive_id = drive_info[0]['id']  # Assume the first drive ID
    print(f"Drive ID: {drive_id}")
    
except Exception as e:
    print(f"Error getting drive ID: {e}")
    exit(1)




# Get the content of the specific folder
try:
    # to get the list of files from the root folder
    print(f"Getting folder content for {site_url}")
    folder_content = client.list_folder_contents(site_id, drive_id) 
    # print(len(folder_content))
    # to get the list of files from the a specific folder
    # folder_id = client.get_folder_id(site_id, drive_id, 'BASIC STUDIES') # path from the root folder folder/subfolder/...
    # folder_content = client.list_folder_contents(site_id, drive_id, folder_id)
except Exception as e:
    print(f"Error getting folder content: {e}")
    exit(1)

local_save_path = "data"

processed_files = set()

def process_element(element, site_id, drive_id, local_save_path):
    element_id = element['id']
    element_type = element['type']

    if element_type == 'folder':
        print(f"Processing folder: {element['name']} ({element_id})")

        try:
            contents = client.list_folder_contents(site_id, drive_id, element_id)
            for content in contents:
                # Pass the updated local save path to the recursive call
                process_element(content, site_id, drive_id, local_save_path)
        except Exception as e:
            print(f"Error processing folder: {e}")

    elif element_type == 'file':
        if element_id in processed_files:
            print(f"Skipping already processed file: {element['name']} ({element_id})")
        else:
            print(f"Processing file: {element['name']} ({element_id})")
            try:
                file_id = element_id
                is_err = client.download_file_contents(site_id, drive_id, file_id, local_save_path)
                processed_files.add(file_id)
            except Exception as e:
                traceback_details = traceback.format_exc()  # Get the full traceback
                print(f"Error type: {type(e).__name__}")  # Print the type of the exception
                print(f"Error processing file: {e} path {element['fullpath']}")  # Print the error with path
                print("Traceback details:")
                print(traceback_details)  # Print the full traceback



# Process only the content of this specific folder
for element in folder_content:
    process_element(element, site_id, drive_id, local_save_path)
