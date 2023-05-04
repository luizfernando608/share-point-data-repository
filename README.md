# SharePoint Data Repository

This project aims to provide a simple and efficient way to use SharePoint as a data repository for data science projects. 

SharePoint is a web-based collaborative platform that integrates with Microsoft Office and offers cloud storage, document management, and collaboration features.

With this code, you can access, upload, download, and delete files and folders from SharePoint using Python. This can help you leverage the benefits of SharePoint, such as security, version control, and metadata, without relying on S3 or other cloud storage services.

To use this code, you need a Microsoft account and a SharePoint site. You can sign up for a free Microsoft account [here](https://signup.live.com/), and create a SharePoint site [here](https://www.microsoft.com/en-us/microsoft-365/sharepoint/online-create-site). 

The main module of this project is `sharepoint_connection.py`, which contains the `SharePointConnection` class that handles the communication with the SharePoint API. You can import this module and create an instance of the `SharePointConnection` class by passing your username, password, and site URL as arguments.
