# https://www.youtube.com/watch?v=dkxVTX5Hs_Q&list=PL2siCn4iJewMoQw-UF56Aximqflbo2q8q&index=13
# from https://www.youtube.com/watch?v=w0pBFo9zpiU&list=PL2siCn4iJewMoQw-UF56Aximqflbo2q8q&index=21
# main library repository: https://github.com/vgrem/Office365-REST-Python-Client.git#egg=Office365-REST-Python-Client
# examples: https://github.com/iamlu-coding/python-sharepoint-office365-api

# example of usage:
"""
import office365_api
sp1 = office365_api.SharePoint(sharepoint_site='https://minersutep.sharepoint.com/sites/CZO_data',
    sharepoint_site_name='CZO_data', sharepoint_doc='data')
print(sp1.get_files_list())
print(sp1.get_folder_list())
sp1.get_files_list('Bahada/Tower/ts_data_2/2024/Raw_Data/ASCII')
"""

# to install:
# pip install Office365-REST-Python-Client
# pip install python-environ
# pip install tqdm

'''
# SharePoint API Integration with Office365 Python Client

## Overview

This project integrates with the SharePoint API using the `Office365-REST-Python-Client` library. The code provides a 
wrapper around the API, enabling functionalities such as file upload, download, and folder management on SharePoint. It 
also supports handling large files by breaking them into chunks for upload or download.

## Requirements

- Python 3.x
- Office365-REST-Python-Client
- python-environ
- tqdm

### Installation

To install the necessary libraries, run the following commands:

```bash
pip install Office365-REST-Python-Client
pip install python-environ
pip install tqdm
```

### Example Usage

```python
import office365_api

# Initialize the SharePoint class with the required site and document library
sp1 = office365_api.SharePoint(sharepoint_site='https://minersutep.sharepoint.com/sites/CZO_data',
    sharepoint_site_name='CZO_data', sharepoint_doc='data')

# Get the list of files in the document library
print(sp1.get_files_list())

# Get the list of folders in the document library
print(sp1.get_folder_list())

# Get files from a specific folder path
sp1.get_files_list('Bahada/Tower/ts_data_2/2024/Raw_Data/ASCII')
```

## Authentication Setup

The `SharePoint` class supports two types of authentication:

1. **Username and Password**
2. **Client ID and Secret**

Both can be set via environment variables using `python-environ` or directly passed into the class.

### Required Environment Variables

Ensure that the following environment variables are set in your `.env` file:

```
sharepoint_email=[your_email]
sharepoint_password=[your_password]
sharepoint_client_id=[your_client_id]
sharepoint_client_secret=[your_client_secret]
sharepoint_url_site=[your_sharepoint_url]
sharepoint_site_name=[your_site_name]
sharepoint_doc_library=[your_document_library]
```

Alternatively, you can provide these values directly when initializing the class.

## Features

### 1. **File and Folder Operations**

- **`get_files_list(folder_name=None)`**: Fetches the list of files in a specified folder. If no folder is specified, it
 fetches the list from the root of the document library.
  
- **`get_folder_list(folder_name=None)`**: Fetches the list of folders in the specified folder. If no folder is 
specified, it lists the root folder.

- **`download_file(file_name, folder_name)`**: Downloads a specified file from a given folder.

- **`download_large_file(file_name, folder_name, local_path_name)`**: Downloads a large file in chunks and saves it 
locally.

- **`upload_large_file(local_file_path, target_file_url, chunk_size=CHUNK_SIZE)`**: Uploads a large file to SharePoint 
in chunks.

- **`upload_file(file_name, folder_name, content)`**: Uploads a small file to the specified folder.

### 2. **Progress Tracking**

The script uses `tqdm` for progress tracking when downloading large files. A progress bar will be displayed to show the 
upload/download progress.

### 3. **Authentication**

The script supports both user-based and client-based authentication. It will attempt to authenticate using the client ID
 and secret first. If those are not provided, it will use the username and password for authentication.

### 4. **Logging**

This integration uses a custom logging mechanism via the `Log` class. Logs are displayed in the console and can 
optionally be saved to a file.

### 5. **Retries**

The script has built-in retry mechanisms for failed file uploads, ensuring that uploads are retried a certain number of 
times before failing completely.

## Example Configuration for `.env`

Here's an example of what the `.env` file might look like:

```
sharepoint_email=user@example.com
sharepoint_password=yourpassword
sharepoint_client_id=your_client_id
sharepoint_client_secret=your_client_secret
sharepoint_url_site=https://minersutep.sharepoint.com/sites/CZO_data
sharepoint_site_name=CZO_data
sharepoint_doc_library=data
```

## References

- [Library Repository](https://github.com/vgrem/Office365-REST-Python-Client.git)
- [YouTube Tutorial - Part 1](https://www.youtube.com/watch?v=dkxVTX5Hs_Q&list=PL2siCn4iJewMoQw-UF56Aximqflbo2q8q&index=13)
- [YouTube Tutorial - Part 2](https://www.youtube.com/watch?v=w0pBFo9zpiU&list=PL2siCn4iJewMoQw-UF56Aximqflbo2q8q&index=21)
- [Example Usage and API](https://github.com/iamlu-coding/python-sharepoint-office365-api)
```
'''

import os
import environ
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File
# from office365.runtime.client_request_exception import ClientRequestException
import datetime
from time import sleep
from tqdm import tqdm
import sys
from pathlib import Path
import Log
import ElapsedTime


CHUNK_SIZE = 20 * 1000000  # 20Mb

env = environ.Env()
environ.Env.read_env()


class SharePoint:
    """
    SharePoint class for interacting with SharePoint's API.

    It supports file and folder operations such as listing, uploading, downloading, and
    managing large files. The class handles authentication via either user credentials or
    client credentials and uses environment variables to store sensitive information.

    Attributes:
        ctx: ClientContext object for handling the SharePoint connection.
        pbar: Progress bar instance for file download and upload tracking.
        log: Log object for capturing events and errors.
        __total_size_: Internal tracking for file size during uploads.
    """
    pbar = None
    __total_size_ = 0

    def __init__(self, username=None, password=None, client_id=None, client_secret=None, sharepoint_site=None,
                 sharepoint_site_name=None, sharepoint_doc=None, log=None):
        """
        Initializes the SharePoint class and authenticates using either user or client credentials.

        :param username: SharePoint username (email).
        :param password: SharePoint password.
        :param client_id: Client ID for SharePoint app authentication.
        :param client_secret: Client secret for SharePoint app authentication.
        :param sharepoint_site: SharePoint site URL.
        :param sharepoint_site_name: SharePoint site name.
        :param sharepoint_doc: SharePoint document library name.
        :param log: Log object to handle logging, defaults to internal Log class.
        """
        self.ctx = None
        if username is None:
            self.__username_ = env('sharepoint_email')
        else:
            self.__username_ = username
        if password is None:
            self.__password_ = env('sharepoint_password')
        else:
            self.__password_ = password
        if client_id is None:
            self.__client_id_ = env('sharepoint_client_id')
        else:
            self.__client_id_ = client_id
        if client_secret is None:
            self.__client_secret_ = env('sharepoint_client_secret')
        else:
            self.__client_secret_ = client_secret
        if sharepoint_site is None:
            self.__sharepoint_site_ = env('sharepoint_url_site')
        else:
            self.__sharepoint_site_ = sharepoint_site
        if sharepoint_site_name is None:
            self.__sharepoint_site_name_ = env('sharepoint_site_name')
        else:
            self.__sharepoint_site_name_ = sharepoint_site_name
        if sharepoint_doc is None:
            self.__sharepoint_doc_ = env('sharepoint_doc_library')
        else:
            self.__sharepoint_doc_ = sharepoint_doc
        if log is not None and isinstance(log, str):
            self.log = Log.Log(log)
        elif isinstance(log, Log.Log):
            self.log = log
        else:
            self.log = Log.Log(fprint=False, sprint=True)
        self.getConnection()

    def getConnection(self, renew=False):
        """
        Authenticates with SharePoint and initializes the connection context.

        :param renew: If True, re-authenticates even if there is an existing connection.
        """
        if renew:
            self.ctx = None
            self.log.live('Connection going to renew...')
        if self.ctx is not None:
            self.log.live('Connection already exists')
            return
        if self.__client_id_ is not None and len(self.__client_id_) > 0 and self.__client_secret_ is not None and len(
                self.__client_secret_) > 0:
            self.log.live('Authenticating with client...')
            self._auth_with_client()
        elif self.__username_ is not None and len(self.__username_) > 0 and self.__password_ is not None and len(
                self.__password_) > 0:
            self.log.live('Authenticating with user...')
            self._auth_with_user()
        else:
            self.log.error('No credentials provided.')
            self.ctx = None

    def print_all_vars(self):
        """
        Prints all internal variables (credentials, site information) for debugging.
        """
        print(f'username: {self.__username_}')
        print(f'sharepoint_site: {self.__sharepoint_site_}')
        print(f'sharepoint_site_name: {self.__sharepoint_site_name_}')
        print(f'sharepoint_doc: {self.__sharepoint_doc_}')
        print(f'client_id: {self.__client_id_}')
        print(f'client_secret: {self.__client_secret_}')

    def _auth_with_user(self):
        """
        Authenticates with SharePoint using username and password credentials.
        """
        try:
            self.ctx = ClientContext(self.__sharepoint_site_).with_credentials(
                UserCredential(self.__username_, self.__password_))
        except Exception as e:
            self.log.error(f'Not possible to authenticate.')
            self.log.error(f'Error: {e}')
            return None
        return self.ctx

    def _auth_with_client(self):
        """
        Authenticates with SharePoint using client ID and secret credentials.
        """
        try:
            client_credentials = ClientCredential(self.__client_id_, self.__client_secret_)
            self.ctx = ClientContext(self.__sharepoint_site_).with_credentials(client_credentials)
        except Exception as e:
            self.log.error(f'Not possible to authenticate.')
            self.log.error(f'Error: {e}')
            return None
        return self.ctx

    def get_files_list(self, folder_name=None):
        """
        Retrieves the list of files from the specified folder in the document library.

        :param folder_name: Name of the folder within the document library to list files from.
        :return: List of files in the folder, or None if the folder is not accessible.
        """
        if self.ctx is None:
            self.getConnection()
        if folder_name is None:
            folder_name = ''
        target_folder_url = f'{self.__sharepoint_doc_}/{folder_name}'
        try:
            root_folder = self.ctx.web.get_folder_by_server_relative_url(target_folder_url)
            root_folder.expand(["Files", "Folders"]).get().execute_query()
        except Exception as e:
            self.log.error(f'Not possible to get files list.')
            self.log.error(f'Error: {e}')
            return None
        return root_folder.files

    def get_folder_list(self, folder_name=None):
        """
        Retrieves the list of subfolders from the specified folder in the document library.

        :param folder_name: Name of the folder to list subfolders from.
        :return: List of subfolders, or None if the folder is not accessible.
        """
        if self.ctx is None:
            self.getConnection()
        if folder_name is None:
            folder_name = ''
        target_folder_url = f'{self.__sharepoint_doc_}/{folder_name}'
        try:
            root_folder = self.ctx.web.get_folder_by_server_relative_url(target_folder_url)
            root_folder.expand(["Folders"]).get().execute_query()
        except Exception as e:
            self.log.error(f'Not possible to get folder list.')
            self.log.error(f'Error: {e}')
            return None
        return root_folder.folders

    def download_file(self, file_name, folder_name):
        """
        Downloads a file from the specified folder in the SharePoint document library.

        :param file_name: Name of the file to download.
        :param folder_name: Name of the folder containing the file.
        :return: The content of the file, or None if the download fails.
        """
        if self.ctx is None:
            self.getConnection()
        file_url = f'/sites/{self.__sharepoint_site_name_}/{self.__sharepoint_doc_}/{folder_name}/{file_name}'
        try:
            file = File.open_binary(self.ctx, file_url)
        except Exception as e:
            self.log.error(f'Not possible to download file.')
            self.log.error(f'Error: {e}')
            return None
        return file.content

    def download_large_file(self, file_name, folder_name, local_path_name):
        """
        Downloads a large file in chunks from SharePoint and saves it locally.

        :param file_name: Name of the file to download.
        :param folder_name: Folder containing the file.
        :param local_path_name: Local path where the downloaded file should be saved.
        :return: True if download succeeds, False otherwise.
        """
        if self.ctx is None:
            self.getConnection()
        file_url = f'/sites/{self.__sharepoint_site_name_}/{self.__sharepoint_doc_}/{folder_name}/{file_name}'
        elapsed_time = ElapsedTime.ElapsedTime()
        try:
            source_file = self.ctx.web.get_file_by_server_relative_path(file_url)
            # Get the file size for the progress bar
            file_info = source_file.get().execute_query()
            total_size = int(file_info.length)
            # Initialize the progress bar
            self.pbar = tqdm(total=total_size, unit='B', unit_scale=True, desc="Downloading", ascii=True)
            # download the file
            with open(local_path_name, 'wb') as local_file:
                source_file.download_session(local_file, self.bar_download_progress).execute_query()
            self.pbar.close()
            self.log.info(f'File {file_name} downloaded successfully in {elapsed_time.elapsed()}')
        except Exception as e:
            self.log.error(f'Not possible to download file.')
            self.log.error(f'Error: {e}')
            return False
        self.log.info(f'File {file_name} downloaded successfully.')
        return True

    def upload_large_file(self, local_file_path, target_file_url, chunk_size=CHUNK_SIZE, _retry=-1):
        """
        Uploads a large file to SharePoint in chunks.

        :param local_file_path: Path to the local file to be uploaded.
        :param target_file_url: Target URL where the file should be uploaded.
        :param chunk_size: Size of each chunk (default: 10MB).
        :param _retry: Number of retries in case of failure (default: -1 for infinite retries).
        :return: True if upload succeeds, False otherwise.
        """
        if self.ctx is None:
            self.getConnection()
        local_file_path = Path(local_file_path)
        target_file_url = Path(target_file_url)
        # make sure the folder exists on SharePoint, if not, it is created
        target_folder_url = f'/sites/{self.__sharepoint_site_name_}/{self.__sharepoint_doc_}/{target_file_url.parent.as_posix()}'
        try:
            self.ctx.web.ensure_folder_path(target_folder_url).execute_query()
        except Exception as e:
            self.log.error(f'Not possible to upload file. When try to create folder {target_folder_url} for file {target_file_url.name}.')
            self.log.error(f'Error: {e}')
            if _retry == -1:
                self.log.info(f'Trying again...')
                self.upload_large_file(local_file_path, target_file_url, _retry=5)
            elif _retry > 0:
                self.log.info(f'And trying again...')
                self.upload_large_file(local_file_path, target_file_url, _retry=_retry - 1)
            else:
                self.log.fatal(f'Not possible to upload {local_file_path} to {target_file_url}!!!')
            return False
        targ_file_url = f'/sites/{self.__sharepoint_site_name_}/{self.__sharepoint_doc_}/{target_file_url.as_posix()}'
        self.log.info(f'Uploading file {local_file_path} to {targ_file_url}...')
        elapsed_time = ElapsedTime.ElapsedTime()
        try:
            self.__total_size_ = os.path.getsize(local_file_path)
            with open(local_file_path, 'rb') as local_file:
                file_name = os.path.basename(targ_file_url)
                folder_url = os.path.dirname(targ_file_url)
                folder = self.ctx.web.get_folder_by_server_relative_url(folder_url)
                upload_session = folder.files.create_upload_session(
                    file_name=file_name,
                    file=local_file,
                    chunk_size=chunk_size,
                    chunk_uploaded=self.bar_upload_progress
                )
                upload_session.execute_query()
            print()
            self.log.info(f'Upload completed in {elapsed_time.elapsed()}')
        except Exception as e:
            self.log.error(f'Not possible to upload file {file_name}.')
            self.log.error(f'Error: {e}')
            if _retry == -1:
                self.log.info(f'Trying again...')
                self.upload_large_file(local_file_path, target_file_url, _retry=5)
            elif _retry > 0:
                self.log.info(f'And trying again...')
                self.upload_large_file(local_file_path, target_file_url, _retry=_retry - 1)
            else:
                self.log.fatal(f'Not possible to upload {local_file_path} to {targ_file_url}!!!')
                return False
        file_properties = self.get_file_properties(file_name, target_file_url.parent.as_posix())
        if file_properties is None:
            file_size_sp = 0
        else:
            file_size_sp = file_properties['file_size']
        if file_size_sp != self.__total_size_:  # check if the file was uploaded correctly
            self.log.error(f'File {file_name} uploaded incorrectly. {file_size_sp} != {self.__total_size_}')
            if _retry == -1:
                self.log.info(f'Trying again...')
                self.upload_large_file(local_file_path, target_file_url, _retry=5)
            elif _retry > 0:
                self.log.info(f'And trying again...')
                self.upload_large_file(local_file_path, target_file_url, _retry=_retry - 1)
            else:
                self.log.fatal(f'Not possible to upload {local_file_path} to {targ_file_url}!!!')
                return False
        self.log.info(f'File {file_name} uploaded successfully.')
        return True

    def download_latest_file(self, folder_name):
        """
        Downloads the most recently modified file from a specified folder in SharePoint.

        :param folder_name: Name of the folder to retrieve the latest file from.
        :return: Tuple of latest file name and its content, or None if the download fails.
        """
        date_format = "%Y-%m-%dT%H:%M:%SZ"
        files_list = self.get_files_list(folder_name)
        if files_list is None:
            return None
        file_dict = {}
        for file in files_list:
            dt_obj = datetime.datetime.strptime(file.time_last_modified, date_format)
            file_dict[file.name] = dt_obj
        # sort dict object to get the latest file
        file_dict_sorted = {key: value for key, value in
                            sorted(file_dict.items(), key=lambda item: item[1], reverse=True)}
        latest_file_name = next(iter(file_dict_sorted))
        content = self.download_file(latest_file_name, folder_name)
        return latest_file_name, content

    def upload_file(self, file_name, folder_name, content):
        """
        Uploads a small file to the specified folder in SharePoint.

        :param file_name: Name of the file to upload.
        :param folder_name: Folder in which to upload the file.
        :param content: Content of the file to upload.
        :return: Response from SharePoint, or None if the upload fails.
        """
        if self.ctx is None:
            self.getConnection()
        target_folder_url = f'/sites/{self.__sharepoint_site_name_}/{self.__sharepoint_doc_}/{folder_name}'
        try:
            target_folder = self.ctx.web.get_folder_by_server_relative_path(target_folder_url)
            return target_folder.upload_file(file_name, content).execute_query()
        except Exception as e:
            self.log.error(f'Not possible to upload file.')
            self.log.error(f'Error: {e}')
            return None

    def upload_file_in_chunks(self, file_path, folder_name, chunk_size, chunk_uploaded=None, **kwargs):
        """
        Uploads a file to SharePoint in chunks to handle large files.

        :param file_path: Local path of the file to be uploaded.
        :param folder_name: Folder in which to upload the file.
        :param chunk_size: Size of each chunk for uploading the file.
        :param chunk_uploaded: Callback function to track progress during upload.
        :param kwargs: Additional arguments for file upload.
        :return: Response from SharePoint, or None if the upload fails.
        """
        if self.ctx is None:
            self.getConnection()
        target_folder_url = f'/sites/{self.__sharepoint_site_name_}/{self.__sharepoint_doc_}/{folder_name}'
        try:
            target_folder = self.ctx.web.get_folder_by_server_relative_path(target_folder_url)
            return target_folder.files.create_upload_session(
                source_path=file_path,
                chunk_size=chunk_size,
                chunk_uploaded=chunk_uploaded,
                **kwargs
            ).execute_query()
        except Exception as e:
            self.log.error(f'Not possible to upload file.')
            self.log.error(f'Error: {e}')
            return None

    def rename_file(self, url_src_path_file, url_dst_path_file, _retry=-1):
        """
        Renames or moves a file in SharePoint.

        :param url_src_path_file: Source file path in SharePoint.
        :param url_dst_path_file: Destination file path in SharePoint.
        :param _retry: Number of retries in case of failure.
        :return: True if the file was successfully renamed, False otherwise.
        """
        if self.ctx is None:
            self.getConnection()
        src = f'/sites/{self.__sharepoint_site_name_}/{self.__sharepoint_doc_}/{Path(url_src_path_file).as_posix()}'
        dst = Path(url_dst_path_file).name
        try:
            # get the file to move
            src_file = self.ctx.web.get_file_by_server_relative_url(src)
            # rename the file
            src_file.rename(dst).execute_query()
        except Exception as e:
            self.log.error(f'Not possible to move file. {src} -> {dst}.')
            self.log.error(f'Error: {e}')
            if _retry == -1:
                self.log.info(f'Trying again...')
                sleep(5)
                self.rename_file(url_src_path_file, url_dst_path_file, _retry=5)
            elif _retry > 0:
                self.log.info(f'And trying again...')
                sleep(5)
                self.rename_file(url_src_path_file, url_dst_path_file, _retry=_retry - 1)
            else:
                self.log.fatal(f'Not possible to move {url_src_path_file} to {url_dst_path_file}!!!')
            return False
        return True

    def get_list(self, list_name):  # this is for lists and NOT files NOR folders
        """
        Retrieves items from a specified SharePoint list.

        :param list_name: Name of the SharePoint list to retrieve items from.
        :return: List of items from the specified SharePoint list.
        """
        if self.ctx is None:
            self.getConnection()
        target_list = self.ctx.web.lists.get_by_title(list_name)
        items = target_list.items.get().execute_query()
        return items

    def get_file_properties_from_folder(self, folder_name):
        """
        Retrieves properties of all files in the specified folder.

        :param folder_name: Name of the folder to retrieve file properties from.
        :return: List of dictionaries containing file properties.
        """
        files_list = self.get_files_list(folder_name)
        if files_list is None:
            print('Waiting a few seconds...')
            sleep(5)
            files_list = self.get_files_list(folder_name)
            if files_list is None:
                return []
        properties_list = []
        for file in files_list:
            file_dict = {
                'file_id': file.unique_id,
                'file_name': file.name,
                'major_version': file.major_version,
                'minor_version': file.minor_version,
                'file_size': file.length,
                'time_created': file.time_created,
                'time_last_modified': file.time_last_modified
            }
            properties_list.append(file_dict)
            # file_dict = {}
        return properties_list

    def get_file_properties(self, file_name, folder_name):
        """
        Retrieves properties of a specific file in a given folder.

        :param file_name: Name of the file to retrieve properties for.
        :param folder_name: Folder containing the file.
        :return: Dictionary containing the file properties, or None if not found.
        """
        file_properties_list = self.get_file_properties_from_folder(folder_name)
        for file in file_properties_list:
            if file['file_name'] == file_name:
                return file
        return None

    def ensure_folder_exists(self, folder_url):
        """
        Ensures that a folder exists in SharePoint. If the folder doesn't exist, it creates it.

        :param folder_url: The URL of the folder to ensure.
        :return: True if the folder exists or was successfully created, False otherwise.
        """
        if self.ctx is None:
            self.getConnection()
        folder_full_url = Path(f'/sites/{self.__sharepoint_site_name_}/{self.__sharepoint_doc_}/{folder_url}')
        try:
            self.ctx.web.ensure_folder_path(folder_full_url.as_posix()).execute_query()
            return True
        except Exception as e:
            self.log.error(f'Problem creating or checking {folder_full_url}.')
            self.log.error(f'Error: {e}')
            return False

    def set_username(self, username):
        self.__username_ = username

    def set_password(self, password):
        self.__password_ = password

    def set_sharepoint_site(self, site):
        self.__sharepoint_site_ = site

    def set_sharepoint_site_name(self, site_name):
        self.__sharepoint_site_name_ = site_name

    def set_sharepoint_doc(self, doc):
        self.__sharepoint_doc_ = doc

    def get_username(self):
        return self.__username_

    def get_sharepoint_site(self):
        return self.__sharepoint_site_

    def get_sharepoint_site_name(self):
        return self.__sharepoint_site_name_

    def get_sharepoint_doc(self):
        return self.__sharepoint_doc_

    def bar_download_progress(self, offset):
        """
        Progress bar handler for file downloads.

        :param offset: Current position in the upload (in bytes).
        """
        self.pbar.update(offset - self.pbar.n)

    def bar_upload_progress(self, offset):
        """
        Progress bar handler for file uploads.

        :param offset: Current position in the upload (in bytes).
        """
        mb = 1024 * 1024
        sys.stdout.write(f"\rUploaded {round(offset/mb,2)} MB of {round(self.__total_size_ / mb, 2)} MB ...[{round(offset / self.__total_size_ * 100, 2)}%]")
        sys.stdout.flush()
