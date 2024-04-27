from dotenv import load_dotenv
import os
from langchain_community.document_loaders.sharepoint import SharePointLoader

load_dotenv()

O365_CLIENT_ID = os.environ.get('O365_CLIENT_ID')
O365_CLIENT_SECRET = os.environ.get('O365_CLIENT_SECRET')

loader = SharePointLoader(
    document_library_id="b!C7H6KUXEp0yjqg1wd184anbhRPIrEYpPlKEpYFiANAfXDSAd-6JAQqfzbtvkeV5_", 
    folder_path="/docs", 
    auth_with_token=False)
#
documents = loader.load()
print(documents)
