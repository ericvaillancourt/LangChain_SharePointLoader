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
# print("Site ID:", site_id)

drive_info = client.get_drive_id(site_id)
# print("Root folder:", drive_info)

drive_id = drive_info[0]['id']  # Assume the first drive ID
folder_content = client.list_folder_contents(site_id, drive_id)
print("Root Content:", folder_content)

folder_id = folder_content[4]['id']

contents = client.list_folder_contents(site_id, drive_id, folder_id)

# for content in contents:
#     print(f"Name: {content['name']}, Type: {content['type']}, MimeType: {content.get('mimeType', 'N/A')}, Path: {content['path']}")

# local_save_path = "data"
# download = client.download_folder_contents(site_id, drive_id, folder_id, local_save_path)

text_splitter = CharacterTextSplitter(separator="\n", 
                                      chunk_size=50, 
                                      chunk_overlap=0)

for content in contents:
    print(f"Processing file: {content['name']}")

    file_id = content['id']
    file_type = content.get('mimeType', 'N/A')
    file_name = content['path']

    loader = client.load_sharepoint_document(site_id, drive_id, file_id, file_name, file_type)
    # docs = loader.load_and_split()  # Use text_splitter if needed, as commented out below
    docs = loader.load_and_split(text_splitter=text_splitter)  # Uncomment this line if you want to use the specific text splitter.

    print(f"Document: {file_name}")
    print(docs)
    print("Number of chunks:", len(docs))
    # print("Length of second chunk's content:", len(docs[1].page_content))  # Uncomment if the Document model includes `page_content`
