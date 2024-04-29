from typing import List, Optional, Union
import pandas as pd
import requests
import os
import io
from io import BytesIO
from langchain_core.document_loaders import Blob
from langchain_core.documents.base import Document
from langchain_community.document_loaders.parsers.pdf import PyPDFParser
from langchain_core.document_loaders.base import BaseLoader
from docx import Document as DocxDocument
from pptx import Presentation


class SharePointClient:
    def __init__(self, tenant_id, client_id, client_secret, resource_url):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.resource_url = resource_url
        self.base_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        self.headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        self.access_token = self.get_access_token()  # Initialize and store the access token upon instantiation

    def get_access_token(self):
        # Body for the access token request
        body = {
            'grant_type': 'client_credentials',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'scope': self.resource_url + '.default'
        }
        response = requests.post(self.base_url, headers=self.headers, data=body)
        return response.json().get('access_token')  # Extract access token from the response

    def get_site_id(self, site_url):
        # Build URL to request site ID
        full_url = f'https://graph.microsoft.com/v1.0/sites/{site_url}'
        response = requests.get(full_url, headers={'Authorization': f'Bearer {self.access_token}'})
        return response.json().get('id')  # Return the site ID

    def get_drive_id(self, site_id):
        # Retrieve drive IDs and names associated with a site
        drives_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives'
        response = requests.get(drives_url, headers={'Authorization': f'Bearer {self.access_token}'})
        drives = response.json().get('value', [])
        return [({'id': drive['id'], 'name': drive['name']}) for drive in drives]

    def get_folder_content(self, site_id, drive_id, folder_path='root'):
        # Build the URL to access the contents of the specified folder
        folder_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/{folder_path}/children'
        response = requests.get(folder_url, headers={'Authorization': f'Bearer {self.access_token}'})
        items_data = response.json()
        rootdir = []
        
        if 'value' in items_data:
            for item in items_data['value']:
                if 'folder' in item:
                    item_type = 'folder'
                    mime_type = None  # No MIME type for folders
                elif 'file' in item:
                    item_type = 'file'
                    mime_type = item['file'].get('mimeType', 'unknown')  # Retrieve MIME type if it's a file
                else:
                    item_type = 'unknown'
                    mime_type = None

                rootdir.append({
                    'id': item['id'],
                    'name': item['name'],
                    'type': item_type,
                    'mimeType': mime_type  # Add MIME type information
                })
        return rootdir


    # Recursive function to browse folders
    def list_folder_contents(self, site_id, drive_id, folder_id, level=0):
        # Get the contents of a specific folder
        folder_contents_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_id}/children'
        contents_headers = {'Authorization': f'Bearer {self.access_token}'}
        contents_response = requests.get(folder_contents_url, headers=contents_headers)
        folder_contents = contents_response.json()

        items_list = []  # List to store information

        if 'value' in folder_contents:
            for item in folder_contents['value']:
                # Split the path on 'root:' and take the second part, if it exists
                path_parts = item['parentReference']['path'].split('root:')
                path = path_parts[1] if len(path_parts) > 1 else ''
                full_path = f"{path}/{item['name']}" if path else item['name']

                if 'folder' in item:
                    # Add folder to list
                    items_list.append(
                        {'id': item['id'], 'name': item['name'], 'type': 'Folder', 'mimeType': None, 'uri': None,
                         'path': full_path})
                    # Recursive call for subfolders
                    items_list.extend(self.list_folder_contents(site_id, drive_id, item['id'], level + 1))
                elif 'file' in item:
                    # Add file to the list with its mimeType and uri
                    items_list.append(
                        {'id': item['id'], 'name': item['name'], 'type': 'File', 'mimeType': item['file']['mimeType'],
                         'uri': item['@microsoft.graph.downloadUrl'], 'path': full_path})

        return items_list

    def download_file(self, download_url, local_path, file_name):
        headers = {'Authorization': f'Bearer {self.access_token}'}
        response = requests.get(download_url, headers=headers)
        if response.status_code == 200:
            full_path = os.path.join(local_path, file_name)
            with open(full_path, 'wb') as file:
                file.write(response.content)
            print(f"File downloaded: {full_path}")
        else:
            print(f"Failed to download {file_name}: {response.status_code} - {response.reason}")

    def download_folder_contents(self, site_id, drive_id, folder_id, local_folder_path, level=0):
        # Recursively download all contents from a folder
        folder_contents_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_id}/children'
        contents_headers = {'Authorization': f'Bearer {self.access_token}'}
        contents_response = requests.get(folder_contents_url, headers=contents_headers)
        folder_contents = contents_response.json()

        if 'value' in folder_contents:
            for item in folder_contents['value']:
                if 'folder' in item:
                    new_path = os.path.join(local_folder_path, item['name'])
                    if not os.path.exists(new_path):
                        os.makedirs(new_path)
                    self.download_folder_contents(site_id, drive_id, item['id'], new_path,
                                                  level + 1)  # Recursive call for subfolders
                elif 'file' in item:
                    file_name = item['name']
                    file_download_url = f"{self.resource_url}/v1.0/sites/{site_id}/drives/{drive_id}/items/{item['id']}/content"
                    self.download_file(file_download_url, local_folder_path, file_name)

    def download_file_contents(self, site_id, drive_id, file_id, local_save_path):
        # Get the file details
        file_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_id}'
        headers = {'Authorization': f'Bearer {self.access_token}'}
        response = requests.get(file_url, headers=headers)
        file_data = response.json()

        # Get the download URL and file name
        download_url = file_data['@microsoft.graph.downloadUrl']
        file_name = file_data['name']

        # Download the file
        self.download_file(download_url, local_save_path, file_name)

    def load_sharepoint_document(self, site_id, drive_id, file_id, file_name, file_type):
        # Get the download URL and the file name by querying the Microsoft Graph API
        file_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_id}'
        headers = {'Authorization': f'Bearer {self.access_token}'}  # Use the stored access token for authorization
        response = requests.get(file_url, headers=headers)  # Make the HTTP request to get file details
        file_data = response.json()  # Parse the JSON response to get file data
        download_url = file_data['@microsoft.graph.downloadUrl']  # Extract the direct download URL from the response

        # Get the file content from the download URL
        response = requests.get(download_url, headers=headers)  # Make the HTTP request to download the file

        # Create a BytesIO object from the response content, which allows for reading and writing bytes in memory
        stream = io.BytesIO(response.content)  # This is useful for handling binary data like files without saving to disk

        # Check the file type and use the appropriate custom loader to handle the file content
        if file_type == 'application/pdf':
            # Use CustomPDFLoader to handle PDF files; it initializes with the stream and file name
            loader = CustomPDFLoader(stream, file_name)
            return loader
        elif file_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            # Use CustomWordLoader for Word documents to handle and potentially split the document's content
            loader = CustomWordLoader(stream, file_name)
            return loader
        elif file_type == 'application/vnd.openxmlformats-officedocument.presentationml.presentation':
            # Use CustomPPTLoader for PowerPoint presentations to read and split the presentation into slides
            loader = CustomPPTLoader(stream, file_name)
            return loader
        elif file_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            # Use CustomExcelLoader for Excel spreadsheets to read and possibly split the sheets into separate parts
            loader = CustomExcelLoader(stream, file_name)
            return loader
        elif file_type in ['text/csv', 'text/plain']:
            # Use CustomTextLoader for plain text or CSV files to handle and split text as needed
            loader = CustomTextLoader(stream, file_name)
            return loader
        else:
            pass  # Placeholder for additional file types that may need to be implemented in the future



class CustomPDFLoader(BaseLoader):
    def __init__(self, stream: BytesIO, filename: str, password: Optional[Union[str, bytes]] = None,
                 extract_images: bool = False):
        # Initialize with a binary stream, file name, optional password, and an image extraction flag
        self.stream = stream
        self.filename = filename
        # Initialize a PDF parser with optional password protection and image extraction settings
        self.parser = PyPDFParser(password=password, extract_images=extract_images)

    def load(self) -> List[Document]:
        # Convert the binary stream into a Blob object which is required by the parser
        blob = Blob.from_data(self.stream.getvalue())
        # Parse the PDF and convert each page or segment into a separate document object
        documents = list(self.parser.parse(blob))

        # Add the filename as metadata to each document for identification
        for doc in documents:
            doc.metadata.update({'source': self.filename})

        return documents

class CustomWordLoader(BaseLoader):
    def __init__(self, stream, filename: str):
        # Initialize with a binary stream and filename
        self.stream = stream
        self.filename = filename

    def load_and_split(self, text_splitter=None):
        # Use python-docx to parse the Word document from the binary stream
        doc = DocxDocument(self.stream)
        # Extract and concatenate all paragraph texts into a single string
        text = "\n".join([p.text for p in doc.paragraphs])

        # Check if a text splitter utility is provided
        if text_splitter is not None:
            # Use the provided splitter to divide the text into manageable documents
            split_text = text_splitter.create_documents([text])
        else:
            # Without a splitter, treat the entire text as one document
            split_text = [{'text': text, 'metadata': {'source': self.filename}}]

        # Add source metadata to each resulting document
        for doc in split_text:
            if isinstance(doc, dict):
                doc['metadata'] = {**doc.get('metadata', {}), 'source': self.filename}
            else:
                doc.metadata = {**doc.metadata, 'source': self.filename}

        return split_text

class CustomExcelLoader(BaseLoader):
    def __init__(self, stream, filename: str):
        # Initialize with a binary stream and filename
        self.stream = stream
        self.filename = filename

    def load_and_split(self, text_splitter=None):
        # Use pandas to load the Excel file from the binary stream
        xls = pd.ExcelFile(self.stream, engine='openpyxl')
        # Get the list of all sheet names in the workbook
        sheet_names = xls.sheet_names

        split_sheets = []
        for sheet in sheet_names:
            # Parse each sheet into a DataFrame
            df = xls.parse(sheet)
            # Convert the DataFrame to a single string with each cell value separated by new lines
            text = '\n'.join(df.values.astype(str).flatten().tolist())

            # Check if a text splitter is provided to further divide the sheet content
            if text_splitter is not None:
                # Use the splitter to create documents from the text
                split_text = text_splitter.create_documents([text])
                # Add metadata to each document
                for doc in split_text:
                    doc.metadata = {'source': self.filename, 'page': sheet}
                split_sheets.extend(split_text)
            else:
                # Without a splitter, treat the entire sheet text as one document
                doc = Document(text, metadata={'source': self.filename, 'page': sheet})
                split_sheets.append(doc)

        return split_sheets

class CustomPPTLoader(BaseLoader):
    def __init__(self, stream, filename):
        # Initialize with a binary stream and filename
        self.stream = stream
        self.filename = filename

    def load_and_split(self, text_splitter=None):
        # Use python-pptx to parse the PowerPoint file from the binary stream
        prs = Presentation(self.stream)
        # Prepare to collect documents
        documents = []

        # Iterate over each slide in the presentation
        for i, slide in enumerate(prs.slides):
            # Extract all text content from each slide
            slide_text = "\n".join([paragraph.text for shape in slide.shapes if shape.has_text_frame for paragraph in shape.text_frame.paragraphs])

            # Check if a text splitter is provided
            if text_splitter is None:
                # Treat each slide's text as a single document
                doc = {'text': slide_text, 'metadata': {'source': self.filename, 'page': i + 1}}
                documents.append(doc)
            else:
                # Use the splitter to divide the slide text into smaller documents
                split_text = text_splitter.create_documents([slide_text])
                # Add metadata and collect each document
                for doc in split_text:
                    doc.metadata = {'source': self.filename, 'page': i + 1}
                documents.extend(split_text)

        return documents



import chardet

class CustomTextLoader(BaseLoader):
    def __init__(self, stream, filename: str):
        self.stream = stream
        self.filename = filename

    def load_and_split(self, text_splitter=None):
        # Use chardet to detect the encoding of the stream
        rawdata = self.stream.read()
        result = chardet.detect(rawdata)
        text = rawdata.decode(result['encoding'])

        if text_splitter is not None:
            split_text = text_splitter.create_documents([text])
        else:
            split_text = [{'text': text, 'metadata': {'source': self.filename}}]

        for doc in split_text:
            if isinstance(doc, dict):
                doc['metadata'] = {**doc.get('metadata', {}), 'source': self.filename}
            else:
                doc.metadata = {**doc.metadata, 'source': self.filename}

        return split_text
