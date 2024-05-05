import traceback
from sharepoint_api import SharePointClient
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()

# Configuration of environment keys
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

local_save_path = "data"
client.download_all_files(site_id, drive_id, local_save_path)