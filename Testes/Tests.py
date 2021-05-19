import pdSharepoint as pds
from configparser import RawConfigParser
from os import remove
from os.path import dirname, abspath, isfile
from sys import platform

#Windows ou Linux
split_char = '/'
if 'linux' in platform:
    path = dirname(abspath(__file__))
elif 'win' in platform:
    split_char = '\\'
    path = abspath('')

#Ler arquivo de configuração
config = RawConfigParser(allow_no_value=True)
config.read(path + split_char + 'config.ini', encoding='utf8')
config_file = dict(config._sections)

for k in config_file:
    config_file[k] = dict(config._defaults, **config_file[k])
    config_file[k].pop('__name__', None)

#Obtendo arquivo excel do Sharepoint
## Variáveis
file_path = '/General/Testes/Pasta.xlsx'
host = 'anatel365.sharepoint.com'
site = 'Pessoal661' #Nome da Equipe no Teams
library = 'Documentos'
client_id = config_file['oauth']['client_id']
client_secret = config_file['oauth']['client_secret']
credentials = (client_id, client_secret)
token_path = config_file['oauth']['token_path']
token_filename = config_file['oauth']['token_filename']
token_filepath = token_path + token_filename
## Consulta ao Sharepoint
df_excel = pds.read_sharepoint_excel(file_path=file_path, host=host, site=site, library=library,
                    credentials=credentials, token_filepath=token_filepath)
## Resultado
print('Head do arquivo Excel do Sharepoint')
print(df_excel.head())
print('-----------------------------------')

#Obtendo lista do Sharepoint
## Variáveis
list_name = 'Pasta'
host = 'anatel365.sharepoint.com'
site = 'Pessoal661' #Nome da Equipe no Teams
cols = ['Coluna A', 'Coluna B', 'Coluna C', 'Coluna D', 'Coluna E']
client_id = config_file['oauth']['client_id']
client_secret = config_file['oauth']['client_secret']
credentials = (client_id, client_secret)
token_path = config_file['oauth']['token_path']
token_filename = config_file['oauth']['token_filename']
token_filepath = token_path + token_filename
## Consulta ao Sharepoint
df_list = pds.read_sharepoint_list(list_name=list_name, host=host, site=site,
                    credentials=credentials, token_filepath=token_filepath)
## Resultado
print('Info da lista do Sharepoint')
print(df_list.info())
print('---------------------------')

#Fazendo download de arquivo excel do Sharepoint
## Variáveis
file_path = '/General/Testes/Pasta.xlsx'
host = 'anatel365.sharepoint.com'
site = 'Pessoal661' #Nome da Equipe no Teams
library = 'Documentos'
client_id = config_file['oauth']['client_id']
client_secret = config_file['oauth']['client_secret']
credentials = (client_id, client_secret)
token_path = config_file['oauth']['token_path']
token_filename = config_file['oauth']['token_filename']
token_filepath = token_path + token_filename
io = path + split_char + 'Pasta.xlsx'
## Consulta ao Sharepoint
pds.download_sharepoint_excel(file_path=file_path, host=host, site=site, library=library,
                   credentials=credentials, token_filepath=token_filepath, io=io)
## Resultados
if isfile(io):
    print(f'Arquivo armazenado em {io}')
else:
    print('Arquivo não armazenado')
print('---------------------------')
remove(io)
