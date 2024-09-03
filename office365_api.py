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


CHUNK_SIZE = 20 * 1000000  # 10MB

env = environ.Env()
environ.Env.read_env()


class SharePoint:
    pbar = None
    __total_size_ = 0

    def __init__(self, username=None, password=None, client_id=None, client_secret=None, sharepoint_site=None,
                 sharepoint_site_name=None, sharepoint_doc=None, log=None):
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
        #self.log = log
        if log is not None and isinstance(log, str):
            self.log = Log.Log(log)
            print('Log is a string')
        elif isinstance(log, Log.Log):
            self.log = log
            print('Log is a log object')
        else:
            self.log = Log.Log(fprint=False, sprint=True)
            print('Log is a default log object with not print into a file')
        self.getConnection()

    def getConnection(self, renew=False):
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
        print(f'username: {self.__username_}')
        print(f'password: {self.__password_}')
        print(f'sharepoint_site: {self.__sharepoint_site_}')
        print(f'sharepoint_site_name: {self.__sharepoint_site_name_}')
        print(f'sharepoint_doc: {self.__sharepoint_doc_}')
        print(f'client_id: {self.__client_id_}')
        print(f'client_secret: {self.__client_secret_}')

    def _auth_with_user(self):  # With username and password
        try:
            self.ctx = ClientContext(self.__sharepoint_site_).with_credentials(
                UserCredential(self.__username_, self.__password_))
            # print(self.get_files_list('Bahada/Tower/ts_data_2/2024/Raw_Data/ASCII'))
        except Exception as e:
            self.log.error(f'Not possible to authenticate.')
            self.log.error(f'Error: {e}')
            return None
        return self.ctx

    def _auth_with_client(self):  # With Client id and secret
        try:
            client_credentials = ClientCredential(self.__client_id_, self.__client_secret_)
            self.ctx = ClientContext(self.__sharepoint_site_).with_credentials(client_credentials)
            # print(self.get_files_list('Bahada/Tower/ts_data_2/2024/Raw_Data/ASCII'))
        except Exception as e:
            self.log.error(f'Not possible to authenticate.')
            self.log.error(f'Error: {e}')
            return None
        return self.ctx

    def get_files_list(self, folder_name=None):
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

    def get_list(self, list_name):  # this is for lists and NOT files NOR folders
        if self.ctx is None:
            self.getConnection()
        target_list = self.ctx.web.lists.get_by_title(list_name)
        items = target_list.items.get().execute_query()
        return items

    def get_file_properties_from_folder(self, folder_name):
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
        file_properties_list = self.get_file_properties_from_folder(folder_name)
        for file in file_properties_list:
            if file['file_name'] == file_name:
                return file
        return None

    def ensure_folder_exists(self, folder_url):
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
        self.pbar.update(offset - self.pbar.n)

    def bar_upload_progress(self, offset):
        mb = 1024 * 1024
        sys.stdout.write(f"\rUploaded {round(offset/mb,2)} MB of {round(self.__total_size_ / mb, 2)} MB ...[{round(offset / self.__total_size_ * 100, 2)}%]")
        sys.stdout.flush()
