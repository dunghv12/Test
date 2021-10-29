from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.client_request_exception import ClientRequestException

import pandas as pd
import os 
import json
import datetime
import io
import platform
import chardet
from io import StringIO, BytesIO


ROOT_PATH = os.getcwd()
config_path = os.path.join(ROOT_PATH, 'config.json')


# read json config file
def read_config_json(path, header):
    with open(path) as config_file:
        config = json.load(config_file)
        config = config[header]
    return config

class SharePoint:
    def __init__(self, config):
        self.context_auth = AuthenticationContext(url=config['url'])
        self.context_auth.acquire_token_for_app(client_id=config['client_id'], 
                                           client_secret=config['client_secret'])
        self.ctx = ClientContext(config['url'], self.context_auth)

    def check_connect(self):
        self.web = self.ctx.web
        self.ctx.load(self.web)
        self.ctx.execute_query()
        print("Web site title: {0} _ Successful Connection!".format(self.web.properties['Title']))
        
    def object_content_url(self, relative_url):
        libraryRoot = self.ctx.web.get_folder_by_server_relative_path(relative_url)
        self.ctx.load(libraryRoot)
        return libraryRoot
        
    def get_content_url(self, relative_url):
        self.libraryRoot = self.object_content_url(relative_url)
#         self.ctx.execute_query()

        # list all folders
        self.folders = self.libraryRoot.folders
        self.ctx.load(self.folders)
        self.ctx.execute_query()
        for myfolder in self.folders:
            print("Folder name: {0}".format(myfolder.properties["ServerRelativeUrl"]))
            
        #list all files
        self.files = self.libraryRoot.files
        self.ctx.load(self.files)
        self.ctx.execute_query()
        my_dict = {'Name':[],'ServerRelativeUrl':[], 'TimeLastModified':[], 'ModTime':[], 'Modified_by_ID':[]}
        for myfile in self.files:
            print("Files name: {0}".format(myfile.properties["ServerRelativeUrl"]))
            #use mod_time to get in better date format
            meta_data = myfile.expand(["modified_by","listItemAllFields"]).get().execute_query()
            mod_time = datetime.datetime.strptime(myfile.properties['TimeLastModified'], '%Y-%m-%dT%H:%M:%SZ')  
            #create a dict of all of the info to add into dataframe and then append to dataframe
            my_dict["Name"].append(myfile.properties['Name'])
            my_dict["ServerRelativeUrl"].append(myfile.properties['ServerRelativeUrl'])
            my_dict["TimeLastModified"].append(myfile.properties['TimeLastModified'])
            my_dict["ModTime"].append(mod_time)
            my_dict["Modified_by_ID"].append(meta_data.listItemAllFields.get_property("EditorId"))
        df_summary_files = pd.DataFrame(my_dict)
            
        return df_summary_files
           
    
    def get_file(self, file_url):
        self.file_name = file_url.split('/')[-1]
        response= File.open_binary(self.ctx, file_url)
            # save data to BytesIO stream
        self.bytes_file_obj = io.BytesIO()
        self.bytes_file_obj.write(response.content)
        self.bytes_file_obj.seek(0)  # set file object to start
            # load Excel file from BytesIO stream
        df = False
        if self.file_name.split('.')[-1] == 'xlsx':
            try:
                df = pd.read_excel(self.bytes_file_obj, header=0)
            except:
                try:
                    self.bytes_file_obj.seek(0)
                    df = pd.read_excel(self.bytes_file_obj, header=0, sep=';')
                except:
                    try:
                        self.bytes_file_obj.seek(0)
                        df = pd.read_excel(self.bytes_file_obj, encoding='cp1252', header=0)
                        print("Read xlsx with Encoding = cp1252 ...")
                    except:
                        print("Can not read {}".format(self.file_name))

        elif self.file_name.split('.')[-1] == 'csv':
            try:
                df = pd.read_csv(self.bytes_file_obj, header=0)
            except:
                try:
                    self.bytes_file_obj.seek(0)
                    df = pd.read_csv(self.bytes_file_obj, header=0, sep=';')
                except:
                    try:
                        self.bytes_file_obj.seek(0)
                        df = pd.read_csv(self.bytes_file_obj, encoding='cp1252', header=0)
                        print("Read csv with Encoding = cp1252 ...")
                    except:
                        print("Can not read {}".format(self.file_name))

        else:
            pass
        
        return df
    
    def upload_file(self, file_url, file_path):
        if platform.system() == 'window':
            file_path = r"{}".format(file_path)
            self.head, self.file_name = os.path.split(file_path)
        else:
            self.file_name = file_path.split('/')[-1]
        print(self.file_name)  
        if os.path.isfile(file_path):
            with open(file_path, 'rb') as content_file:
                self.file_content = content_file.read()
            print(file_path)
            file = self.ctx.web.get_folder_by_server_relative_url(file_url).upload_file(self.file_name, self.file_content).execute_query()
        else:
            print("Check file path again!")

    def upload_dataframe(self, file_url, dataframe, filename):

        buffer = BytesIO()               # Create a buffer object
        dataframe.to_csv(buffer) # Write the dataframe to the buffer
        buffer.seek(0)
        self.file_content = buffer.read()
        # self.file_content = self.buffer.read()
        file = self.ctx.web.get_folder_by_server_relative_url(file_url).upload_file(filename, self.file_content).execute_query()
            
        

