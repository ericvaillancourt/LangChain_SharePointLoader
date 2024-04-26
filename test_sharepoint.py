import os
from dotenv import load_dotenv
from langchain_text_splitters import CharacterTextSplitter
from sharepoint_api import SharePointClient

# Charger les variables d'environnement du fichier .env
load_dotenv()

# Récupération des variables depuis les variables d'environnement
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

folder_id = folder_content[0]['id']

contents = client.list_folder_contents(site_id, drive_id, folder_id)

for content in contents:
    print(f"Name: {content['name']}, Type: {content['type']}, MimeType: {content.get('mimeType', 'N/A')}, Path: {content['path']}")

local_save_path = "data"

loader = client.load_sharepoint_document(site_id, drive_id, file_id, file_name, file_type)

text_splitter = CharacterTextSplitter(
    separator="\n",
    chunk_size=200,
    chunk_overlap=0
)
docs = loader.load_and_split()
print(docs)
print(len(docs))
