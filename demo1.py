from dotenv import load_dotenv
import os
from langchain_community.document_loaders.sharepoint import SharePointLoader

load_dotenv()

O365_CLIENT_ID = os.environ.get('O365_CLIENT_ID')
O365_CLIENT_SECRET = os.environ.get('O365_CLIENT_SECRET')
DOCUMENT_LIBRARY_ID = os.environ.get('DOCUMENT_LIBRARY_ID')

loader = SharePointLoader(
    document_library_id=DOCUMENT_LIBRARY_ID, 
    folder_path="/txt", 
    auth_with_token=False)
#
documents = loader.load()
print(documents)
