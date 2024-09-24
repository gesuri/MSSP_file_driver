# MSSP_file_driver
Driver to upload and download files from Microsoft SharePopint

---



## Overview
"""
This script provides a wrapper class `SharePoint` to interact with SharePoint resources 
using the Office365-REST-Python-Client library. The class supports authentication via 
username/password or client credentials, file and folder operations (list, download, upload), 
as well as some auxiliary functionalities like logging, progress bars, and chunk-based uploads.

Libraries:
- `office365.sharepoint.client_context.ClientContext`: For connecting to the SharePoint site.
- `office365.runtime.auth.user_credential.UserCredential`: For user credential-based authentication.
- `office365.runtime.auth.client_credential.ClientCredential`: For client ID/secret-based authentication.
- `office365.sharepoint.files.file.File`: For downloading and uploading files.
- `tqdm`: For displaying a progress bar during file download/upload.
- `environ`: For handling environment variables.

Usage Example:
```
sp1 = SharePoint(sharepoint_site='https://example.sharepoint.com/sites/my_site', sharepoint_site_name='my_site', sharepoint_doc='Documents')
print(sp1.get_files_list())
print(sp1.get_folder_list())
sp1.get_files_list('Path/To/Files')
```

Dependencies:
```
- `pip install Office365-REST-Python-Client`
- `pip install python-environ`
- `pip install tqdm`
"""
```

## Requiered Environment Variables
Ensure that the following environment variables are set in your .env file:

```css
sharepoint_email=[your_email]
sharepoint_password=[your_password]
sharepoint_client_id=[your_client_id]
sharepoint_client_secret=[your_client_secret]
sharepoint_url_site=[your_sharepoint_url]
sharepoint_site_name=[your_site_name]
sharepoint_doc_library=[your_document_library]
```

### Class: `SharePoint`
This class encapsulates functionality to interact with a SharePoint site. It supports authentication using either user credentials (username/password) or client credentials (client ID/secret).
https://github.com/vgrem/Office365-REST-Python-Client

#### `__init__(self, username=None, password=None, client_id=None, client_secret=None, sharepoint_site=None, sharepoint_site_name=None, sharepoint_doc=None, log=None)`
- **Parameters**:
  - `username`: The username to authenticate with SharePoint. If not provided, it falls back to the environment variable `sharepoint_email`.
  - `password`: The password to authenticate with SharePoint. Defaults to `sharepoint_password` from environment variables.
  - `client_id`: Optional client ID for client credential authentication.
  - `client_secret`: Optional client secret for client credential authentication.
  - `sharepoint_site`: The SharePoint site URL to connect to.
  - `sharepoint_site_name`: The name of the SharePoint site.
  - `sharepoint_doc`: The document library where operations will occur.
  - `log`: A custom logging object. If not provided, a default logger will be created.

#### `getConnection(self, renew=False)`
- **Description**: Establishes a connection to SharePoint, using either client credentials or user credentials, based on the available data. If the connection already exists, it reuses it unless `renew` is set to True.

#### `get_files_list(self, folder_name=None)`
- **Description**: Retrieves a list of files from the specified folder in SharePoint.
- **Parameters**:
  - `folder_name`: The relative path of the folder to list files from. Defaults to the root of the document library.
- **Returns**: A list of files in the folder.

#### `get_folder_list(self, folder_name=None)`
- **Description**: Retrieves a list of subfolders from the specified folder.
- **Parameters**:
  - `folder_name`: The relative path of the folder to list subfolders from.
- **Returns**: A list of subfolders in the folder.

#### `download_file(self, file_name, folder_name)`
- **Description**: Downloads a file from SharePoint.
- **Parameters**:
  - `file_name`: The name of the file to download.
  - `folder_name`: The folder path where the file is located.
- **Returns**: The content of the file.

#### `download_large_file(self, file_name, folder_name, local_path_name)`
- **Description**: Downloads large files from SharePoint in chunks, updating a progress bar.
- **Parameters**:
  - `file_name`: The name of the file to download.
  - `folder_name`: The folder where the file is located.
  - `local_path_name`: The local path where the file will be saved.

#### `upload_large_file(self, local_file_path, target_file_url, chunk_size=CHUNK_SIZE, _retry=-1)`
- **Description**: Uploads a large file to SharePoint in chunks.
- **Parameters**:
  - `local_file_path`: Path to the local file to be uploaded.
  - `target_file_url`: SharePoint target path where the file should be uploaded.
  - `chunk_size`: Size of each chunk for the upload. Defaults to 10 MB.
  - `_retry`: Number of retries if the upload fails.

#### `rename_file(self, url_src_path_file, url_dst_path_file, _retry=-1)`
- **Description**: Renames or moves a file in SharePoint.
- **Parameters**:
  - `url_src_path_file`: The current URL path of the file.
  - `url_dst_path_file`: The new URL path for the file.

#### `get_file_properties_from_folder(self, folder_name)`
- **Description**: Retrieves properties of all files in the specified folder.
- **Parameters**:
  - `folder_name`: The folder to retrieve file properties from.
- **Returns**: A list of dictionaries containing file properties such as name, size, and timestamps.

#### `bar_download_progress(self, offset)`
- **Description**: Updates the progress bar during a file download.

#### `bar_upload_progress(self, offset)`
- **Description**: Displays the progress of file uploads in megabytes.

### Notes:
- **Environment Variables**: If credentials and other necessary details are not passed as parameters, the code attempts to read them from the environment. Ensure that variables like `sharepoint_email`, `sharepoint_password`, `sharepoint_client_id`, `sharepoint_client_secret`, etc., are set in the environment.
- **Logging**: The logging mechanism is either a custom `Log` object or a default logger that prints to the console.
- **Error Handling**: For every SharePoint-related operation, errors are caught and logged. The retry mechanism is in place for critical file operations like uploads and renaming.

