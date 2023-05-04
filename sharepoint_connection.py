#%%
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File
import pandas as pd
from io import BytesIO
# import display 
from IPython.display import display


#%%

def print_success(*args, **kwargs):
    print("\033[92m", *args, "\033[0m", **kwargs)


#%%
class SharePointConnection:
    def __init__(self, username:str, password:str, sharepoint_site, sharepoint_site_name, sharepoint_doc):
        """Insira o email e senha da FGV para acessar o SharePoint

        Parameters
        ----------
        username : str
            Email da FGV no formato @fgv.edu.br
        password : str
            Senha da FGV.
        """
        self.username = username
        self.password = password
        self.sharepoint_site = sharepoint_site
        self.sharepoint_site_name = sharepoint_site_name
        self.sharepoint_doc = sharepoint_doc
        # verify credentials 
        conn = self._auth()
        try:
            conn.web.get().execute_query()
            print("\033[92m Conexão com o SharePoint realizada com sucesso \033[0m")
        except ValueError:
            raise ValueError("Credenciais inválidas")
        finally:
            # print("Conexão com o SharePoint realizada com sucesso")
            # print green
            pass

    def _auth(self):
        conn = ClientContext(self.sharepoint_site).with_credentials(
            UserCredential(
                self.username,
                self.password
            )
        )
        return conn

    def _get_files_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f'{self.sharepoint_doc}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Files", "Folders"]).get().execute_query()
        return root_folder.files
    
    def get_folder_list(self, folder_name:str)->list:
        conn = self._auth()
        target_folder_url = f'{self.sharepoint_doc}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Folders"]).get().execute_query()
        folder_names = list()
        for folder in root_folder.folders:
            folder_path = folder.properties['Name']
            folder_names.append(folder_path)
        return folder_names
        
    def get_file_list(self, folder_name:str)->list:
        conn = self._auth()
        root_folder = conn.web.get_folder_by_server_relative_url(f"{self.sharepoint_doc}/{folder_name}")
        files = root_folder.expand(["Files"]).get().execute_query()
        file_names = list()
        for file in files.files:
            file_path = file.properties['ServerRelativeUrl']
            file_name = file_path.split('/')[-1]
            file_names.append(file_name)
        return file_names
    
    def show_items_folder(self,  folder_name:str)->list:
        folder_names = self.get_folder_list(folder_name)
        file_names = self.get_file_list(folder_name)
        print(folder_name+"/"+"\n")
        for folder in folder_names:
            print("\t"+folder+"/")
        for file in file_names:
            print("\t"+file)


    def read_parquet(self,file_path:str, **kwargs):
        buffer = BytesIO()
        conn = self._auth()
        file_url = f'/sites/{self.sharepoint_site_name}/{self.sharepoint_doc}/{file_path}'
        file = File.open_binary(conn, file_url)
        buffer.write(file.content)
        df = pd.read_parquet(buffer, **kwargs)
        return df


    def read_csv(self, file_path:str, **kwargs):
        conn = self._auth()
        file_url = f'/sites/{self.sharepoint_site_name}/{self.sharepoint_doc}/{file_path}'
        
        file = File.open_binary(conn, file_url)
        df = pd.read_csv(BytesIO(file.content), **kwargs)
        return df


    def write_parquet(self, df:pd.DataFrame, file_path:str,  **kwargs):
        conn = self._auth()

        buffer = BytesIO()
        df.to_parquet(buffer, **kwargs)

        folder_name = "/".join(file_path.split("/")[:-1])
        file_name = file_path.split("/")[-1]
        target_folder_url = f'{self.sharepoint_doc}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)

        root_folder.upload_file(file_name, buffer.getvalue()).execute_query()
        
        print_success("File uploaded successfully")


    def write_csv(self, df:pd.DataFrame, file_path:str,  **kwargs):
        conn = self._auth()

        buffer = BytesIO()
        df.to_csv(buffer, index=False, mode="wb",encoding="utf-8", **kwargs)

        folder_name = "/".join(file_path.split("/")[:-1])
        file_name = file_path.split("/")[-1]
        target_folder_url = f'{self.sharepoint_doc}/{folder_name}'

        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        try:
            root_folder.upload_file(file_name, buffer.getvalue()).execute_query()
        except:
            raise ValueError("O caminho do arquivo não existe")
        
        print_success("File uploaded successfully")



if __name__ == "__main__":
    
    sp = SharePointConnection("email@example.com", "******")

    folder = "dados/1_entrada"

    sp.show_items_folder(folder)
    data = sp.read_csv("dados/1_entrada/awards.csv")
    display(data.head())
    sp.write_csv(data, "dados/lot/1_entrada/dados_teste1.csv")

    data2 = sp.read_csv("dados/1_entrada/dados_teste1.csv")
    display(data2.head())

    sp.write_parquet(data, "dados/1_entrada/dados_teste1.parquet")
    data3 = sp.read_parquet("dados/1_entrada/dados_teste1.parquet")
    display(data3.head())



