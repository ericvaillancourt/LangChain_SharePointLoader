import os
from dotenv import load_dotenv
from langchain_text_splitters import CharacterTextSplitter
from sharepoint_api import SharePointClient

# Load environment variables from .env file
load_dotenv()

# Retrieval of variables from environment variables
tenant_id = os.getenv('TENANT_ID')
client_id = os.getenv('CLIENT_ID')
client_secret = os.getenv('CLIENT_SECRET')
site_url = os.getenv('SITE_URL')
resource = os.getenv('RESOURCE')

client = SharePointClient(tenant_id, client_id, client_secret, resource)
site_id = client.get_site_id(site_url)
print("Site ID:", site_id)

drive_info = client.get_drive_id(site_id)
print("Root folder:", drive_info)

drive_id = drive_info[0]['id']  # Assume the first drive ID
folder_content = client.get_folder_content(site_id, drive_id)
print("Root Content:", folder_content)

# folder_id = folder_content[0]['id']

# contents = client.list_folder_contents(site_id, drive_id, folder_id)