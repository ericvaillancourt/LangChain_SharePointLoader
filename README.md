# SharePoint Document Management with Python

This repository provides a set of tools and scripts designed to interact with Microsoft SharePoint to manage and process documents. Utilizing Python, this project simplifies tasks such as retrieving, managing, and processing documents stored in SharePoint.

## Features

- Retrieve site and drive IDs from SharePoint
- List contents of a folder in SharePoint
- Download files and folders from SharePoint
- Process various types of documents (PDF, DOCX, PPTX, CSV, TXT)
- Use custom loaders for handling different file formats
- Integration with Python's `dotenv` package for environment management

## Prerequisites

Before you begin, ensure you have the following:

- Python 3.11 or higher
- Access to a Microsoft SharePoint instance
- Client credentials for accessing Microsoft Graph API

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/ericvaillancourt/LangChain_SharePointLoader
   ```
2. Navigate to the cloned directory:
   ```bash
   cd LangChain_SharePointLoader
   ```
3. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Setup

1. Create a `.env` file in the root directory of your project.
2. Add the following environment variables:
   ```
   TENANT_ID=your_tenant_id
   CLIENT_ID=your_client_id
   CLIENT_SECRET=your_client_secret
   SITE_URL=your_site_url
   RESOURCE=your_resource_url
   ```
3. Replace `your_tenant_id`, `your_client_id`, `your_client_secret`, `your_site_url`, and `your_resource_url` with your actual SharePoint and Microsoft Graph credentials.

## Usage

To start using this tool with SharePoint:

1. Run the main script to interact with SharePoint:
   ```bash
   python main.py
   ```
   This will perform operations defined in the script, such as listing folder contents and downloading files.

2. Check the output in the console for results of operations such as retrieving site ID, drive information, and document processing.

## Contributing

Contributions are welcome! If you'd like to contribute to this project, please follow these steps:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Make your changes and commit them (`git commit -am 'Add some feature'`).
4. Push to the branch (`git push origin feature-branch`).
5. Create a new Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
