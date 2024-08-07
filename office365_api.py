## https://www.youtube.com/watch?v=dkxVTX5Hs_Q&list=PL2siCn4iJewMoQw-UF56Aximqflbo2q8q&index=13
## from https://www.youtube.com/watch?v=w0pBFo9zpiU&list=PL2siCn4iJewMoQw-UF56Aximqflbo2q8q&index=21
# main library repository: https://github.com/vgrem/Office365-REST-Python-Client.git#egg=Office365-REST-Python-Client
# examples: https://github.com/iamlu-coding/python-sharepoint-office365-api

# example of usage:
'''
import office365_api
sp1 = office365_api.SharePoint(sharepoint_site='https://minersutep.sharepoint.com/sites/CZO_data',
    sharepoint_site_name='CZO_data', sharepoint_doc='data')
print(sp1.get_files_list())
print(sp1.get_folder_list())
sp1.get_files_list('Bahada/Tower/ts_data_2/2024/Raw_Data/ASCII')
'''

# to install:
# pip install Office365-REST-Python-Client
# pip install python-environ

from urllib import response
import environ
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File
import datetime

from itertools import permutations

env = environ.Env()
environ.Env.read_env()

values = ['2ebf97c4-d59a-4cc2-a690-f9b2c288bdd6',
          'cab2a1ef-78ec-43d6-96f2-57f4aefe64cf',
          '857c21d2-1a16-43a4-90cf-d57f3fab9d2f',
          'Twx8Q~vSfk4C8qcxV521cF3K2~aOcMwgYzPBLbnF',
          'd53d3165-e13d-486e-872f-a7ea0c003fc5']
def generate_username_password_combinations(values):
    if len(values) != 5:
        raise ValueError("The input list must contain exactly 5 values.")
    # Generate all permutations of the 5 values taken 2 at a time
    return list(permutations(values, 2))


class SharePoint:
    def __init__(self, username=None, password=None, sharepoint_site=None, sharepoint_site_name=None, sharepoint_doc=None):
        if username is None:
            self.__username_ = env('sharepoint_email')
        else:
            self.__username_ = username
        if password is None:
            self.__password_ = env('sharepoint_password')
        else:
            self.__password_ = password
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
        self.conn = None

    def print_all_vars(self):
        print(f'username: {self.__username_}')
        print(f'password: {self.__password_}')
        print(f'sharepoint_site: {self.__sharepoint_site_}')
        print(f'sharepoint_site_name: {self.__sharepoint_site_name_}')
        print(f'sharepoint_doc: {self.__sharepoint_doc_}')

    def _auth(self):  # With username and password
        try:
            self.conn = ClientContext(self.__sharepoint_site_).with_credentials(
                UserCredential(self.__username_, self.__password_))
            #print(self.get_files_list('Bahada/Tower/ts_data_2/2024/Raw_Data/ASCII'))
        except Exception as e:
            print(f'Not possible to authenticate.\nError: {e}')
            return None
        return self.conn

    def _auth_(self):  # With Client id and secret
        CLIENT_ID = 'kF-8Q~zeGV3HjcZlPCtgAPhZnAMVaio73LmG9c5C'
        CLIENT_SECRET = 'c9f13109-0ac9-4a5e-81a4-23d1a00cb8b2'
        try:
            client_credentials = ClientCredential(CLIENT_ID, CLIENT_SECRET)
            self.conn = ClientContext(self.__sharepoint_site_).with_credentials(client_credentials)
            print(self.get_files_list('Bahada/Tower/ts_data_2/2024/Raw_Data/ASCII'))
        except Exception as e:
            print(f'Not possible to authenticate.\nError: {e}')
            return None
        return self.conn

    def get_files_list(self, folder_name=None):
        if self.conn is None:
            self.conn = self._auth()
        if folder_name is None:
            folder_name = ''
        target_folder_url = f'{self.__sharepoint_doc_}/{folder_name}'
        print(f'target_folder_url: {target_folder_url}')
        try:
            root_folder = self.conn.web.get_folder_by_server_relative_url(target_folder_url)
            root_folder.expand(["Files", "Folders"]).get().execute_query()
        except Exception as e:
            print(f'Not possible to get files list.\nError: {e}')
            return None
        return root_folder.files
    
    def get_folder_list(self, folder_name=None):
        if self.conn is None:
            self.conn = self._auth()
        if folder_name is None:
            folder_name = ''
        target_folder_url = f'{self.__sharepoint_doc_}/{folder_name}'
        try:
            root_folder = self.conn.web.get_folder_by_server_relative_url(target_folder_url)
            root_folder.expand(["Folders"]).get().execute_query()
        except Exception as e:
            print(f'Not possible to get folder list.\nError: {e}')
            return None
        return root_folder.folders

    def download_file(self, file_name, folder_name):
        if self.conn is None:
            self.conn = self._auth()
        file_url = f'/sites/{self.__sharepoint_site_name_}/{self.__sharepoint_doc_}/{folder_name}/{file_name}'
        try:
            file = File.open_binary(self.conn, file_url)
        except Exception as e:
            print(f'Not possible to download file.\nError: {e}')
            return None
        return file.content
    
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
        file_dict_sorted = {key:value for key, value in sorted(file_dict.items(), key=lambda item:item[1], reverse=True)}    
        latest_file_name = next(iter(file_dict_sorted))
        content = self.download_file(latest_file_name, folder_name)
        return latest_file_name, content
        
    def upload_file(self, file_name, folder_name, content):
        if self.conn is None:
            self.conn = self._auth()
        target_folder_url = f'/sites/{self.__sharepoint_site_name_}/{self.__sharepoint_doc_}/{folder_name}'
        try:
            target_folder = self.conn.web.get_folder_by_server_relative_path(target_folder_url)
            return target_folder.upload_file(file_name, content).execute_query()
        except Exception as e:
            print(f'Not possible to upload file.\nError: {e}')
            return None

    def upload_file_in_chunks(self, file_path, folder_name, chunk_size, chunk_uploaded=None, **kwargs):
        if self.conn is None:
            self.conn = self._auth()
        target_folder_url = f'/sites/{self.__sharepoint_site_name_}/{self.__sharepoint_doc_}/{folder_name}'
        try:
            target_folder = self.conn.web.get_folder_by_server_relative_path(target_folder_url)
            return target_folder.files.create_upload_session(
                source_path=file_path,
                chunk_size=chunk_size,
                chunk_uploaded=chunk_uploaded,
                **kwargs
            ).execute_query()
        except Exception as e:
            print(f'Not possible to upload file.\nError: {e}')
            return None
    
    def get_list(self, list_name):  # this is for lists and NOT files NOR folders
        if self.conn is None:
            self.conn = self._auth()
        target_list = self.conn.web.lists.get_by_title(list_name)
        items = target_list.items.get().execute_query()
        return items
        
    def get_file_properties_from_folder(self, folder_name):
        files_list = self.get_files_list(folder_name)
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
            file_dict = {}
        return properties_list

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

    def get_password(self):
        return self.__password_

    def get_sharepoint_site(self):
        return self.__sharepoint_site_

    def get_sharepoint_site_name(self):
        return self.__sharepoint_site_name_

    def get_sharepoint_doc(self):
        return self.__sharepoint_doc_