# pdSharepoint
Pacote para obter arquivos excel ou listas do Sharepoint em formato DataFrame Pandas. Esse pacote usa o <a href='https://pypi.org/project/O365/'>projeto O365</a>.

## Autenticação Sharepoint

É necessário ter uma conta da Microsoft com acesso a um Sharepoint.

Passo-a-passo para autenticação do script:

Para permitir autenticação, primeiro é necessário registrar uma aplicação no <a href='https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade'>Registro de Aplicativos do Azure</a>:
1.	Logar no Portal do Azure (Registro de Aplicações) com a sua conta do Sharepoint.
2.	Clique em “+Novo Registro”
3.	Dê um nome para a sua aplicação.
4.	Em “Tipos de conta com suporte” e escolha a opção “Contas em qualquer diretório organizacional (Qualquer diretório do Azure AD – Multilocatário)”
5.	Coloque a URI de redirecionamento do tipo Web para https://login.microsoftonline.com/common/oauth2/nativeclient e clique em Registrar. O link precisa ser inserido no campo de texto, pois simplesmente aceitar o check box do lado do campo de texto não é suficiente.
6.	Anote o id da Aplicação (id do cliente). Ele será usado pelo pdSharepoint. Não compartilhe esse id com ninguém.
7.	Na lista à esquerda, clique em “Certificados e Segredos” e gere um novo segredo de cliente. Coloque a data de expiração para “Nunca”. Anote o valor do segredo criado agora, pois ele ficará oculto daqui para frente. O pdSharepoint irá utilizar esse segredo. Não compartilhe esse segredo com ninguém.
8.	Nas lista à esquerda, clique em “Permissões de APIs”, clique em adicionar uma permissão.
9.	Clique em Microsoft Graph
10. Procure na lista e expanda Sites.
11. Selecione Sites.Read.All e Sites.ReadWrite.All
12. Procure na lista e expanda User.
13. Selecione User.Read
14. Clique em Todas as APIs no canto superior esquerda para voltar a tela anterior.
15.	Procure na lista e clique em Sharepoint
16.	Clique em Permissões delegadas
17.	Expanda AllSites e selecione: AllSites.Read, AllSites.Write
18.	Expanda MyFiles e selecione: MyFiles.Read, MyFiles.Write
19.	Expanda Sites e selecione: Sites.Search.All
20.	Clique em “Adicionar permissões”

Sugestão: criar um arquivo config.ini com os seguintes campos. Atenção: esse arquivo deve ser armazenado em local seguro!
```
[oauth]
token_path:<pasta local onde o token será armazenado>
token_filename:o365_token.txt
client_id:<id do cliente gerado no passo 6>
client_secret:<segredo do cliente gerado no passo 7>
```

A primeira vez que rodar seu script usando o pdSharepoint, o Azure irá solicitar a geração de um token. Um link será fornecido ao usuário.

![img.png](img.png)

O usuário deverá entrar no link e fornecer consentimento para o script. Uma vez fornecido consentimento, deve-se copiar o endereço que consta na barra de endereço do navegador e colar no campo solicitado no prompt.
Se tudo der certo, a mensagem abaixo aparecerá no prompt.

![img_1.png](img_1.png)

O token será armazenado no arquivo identificado pelos campos token_path e token_filename que constam no config.ini. O arquivo com o token deve ser guardado com segurança, pois ele dará acesso ao Sharepoint do usuário. O seu script precisará ter acesso a esse arquivo de agora em diante.
O token tem validade de 90 dias, porém a cada nova execução do script, o token será renovado com a mesma validade. Assim, contando que o script rode novamente em menos de 90 dias, não será necessária nova interação com o usuário.
## Instalar

```
pip install pdSharepoint
```
## Importar pacote
```
import pdSharepoint as pds
```
## Fazer o download de um arquivo excel do Sharepoint
Crie uma variável com o endereço do host Sharepoint.
```
host = 'meu_servidor.sharepoint.com'
```
Crie uma variável com o nome da Equipe/Site do Sharepoint.
```
site = 'MinhaEquipe'
```
Crie uma variável com o nome da biblioteca no Site que contém o arquivo desejado.
```
library = 'Documentos'
```
Crie uma variável com o caminho para o arquivo excel do Sharepoint a partir de dentro da biblioteca do site.
```
file_path = '/General/MeuArquivo.xlsx'
```
Crie um variável com o id do cliente criado no MS Azure.
```
#pode ser obtido do arquivo config.ini criado acima
client_id = '{seu id de cliente}'
```
Crie um variável com o segredo do cliente criado no MS Azure.
```
#pode ser obtido do arquivo config.ini criado acima
client_secret = '{seu segredo de cliente}'
```
Crie uma tupla com as duas credenciais acima.
```
credentials = (client_id, client_secret)
```
Crie uma variável com o nome e caminho do arquivo que armazenará o token. Lembre de manter esse arquivo em local seguro.
```
#pode ser obtido do arquivo config.ini criado acima
token_filepath = '{nome e caminho para seu arquivo de token}'
```
Crie uma variável com o pasta local onde deseja armazenar o arquivo excel do Sharepoint. Se io não for informado, o arquivo será armazenado na pasta de execução do script.
```
io = 'C:\\MeuArquivo.xlsx'
```
Faça o download do arquivo.
```
pds.download_sharepoint_excel(file_path=file_path, host=host, site=site, library=library,
                   credentials=credentials, token_filepath=token_filepath, io=io)
```
Código completo para download.
```
import pdSharepoint as pds

host = 'meu_servidor.sharepoint.com'
site = 'MinhaEquipe'
library = 'Documentos'
file_path = '/General/MeuArquivo.xlsx'
client_id = '{seu id de cliente}' #pode ser obtido do arquivo config.ini criado acima
client_secret = '{seu segredo de cliente}' #pode ser obtido do arquivo config.ini criado acima
credentials = (client_id, client_secret)
token_filepath = '{nome e caminho para seu arquivo de token}' #pode ser obtido do arquivo config.ini criado acima
io = 'C:\\MeuArquivo.xlsx'

pds.download_sharepoint_excel(file_path=file_path, host=host, site=site, library=library,
                   credentials=credentials, token_filepath=token_filepath, io=io)
```
## Obter um arquivo excel do Sharepoint em um DataFrame Pandas.
Use as mesmas variáveis criadas acima. O pdSharepoint irá fazer o download do arquivo excel na pasta de execução do script, criar um DataFrame pandas com esse arquivo e depois deletar o arquivo, mantendo o objeto DataFrame para ser usado pelo seu script.

Todos os argumentos de pandas.read_excel podem ser utilizados (exceto io).
```
import pdSharepoint as pds
import pandas as pd

host = 'meu_servidor.sharepoint.com'
site = 'MinhaEquipe'
library = 'Documentos'
file_path = '/General/MeuArquivo.xlsx'
client_id = '{seu id de cliente}' #pode ser obtido do arquivo config.ini criado acima
client_secret = '{seu segredo de cliente}' #pode ser obtido do arquivo config.ini criado acima
credentials = (client_id, client_secret)
token_filepath = '{nome e caminho para seu arquivo de token}' #pode ser obtido do arquivo config.ini criado acima

df = pds.read_sharepoint_excel(file_path=file_path, host=host, site=site, library=library,
                    credentials=credentials, token_filepath=token_filepath)
```
## Obter uma lista do Sharepoint em um DataFrame Pandas
Além das variáveis acima, crie também uma variável para conter o nome da lista.
```
list_name = 'MinhaLista'
```
Faça a consulta ao Sharepoint. O pdSharepoint irá criar um pandas DataFrame com os dados da lista.
```
df = pds.read_sharepoint_list(list_name=list_name, host=host, site=site,
                    credentials=credentials, token_filepath=token_filepath)
```
Código completo para obter listas do Sharepoint em DataFrames pandas.
```
import pdSharepoint as pds
import pandas as pd

host = 'meu_servidor.sharepoint.com'
site = 'MinhaEquipe'
list_name = 'MinhaLista'
client_id = '{seu id de cliente}' #pode ser obtido do arquivo config.ini criado acima
client_secret = '{seu segredo de cliente}' #pode ser obtido do arquivo config.ini criado acima
credentials = (client_id, client_secret)
token_filepath = '{nome e caminho para seu arquivo de token}' #pode ser obtido do arquivo config.ini criado acima

df = pds.read_sharepoint_list(list_name=list_name, host=host, site=site,
                    credentials=credentials, token_filepath=token_filepath)
```
Na esperança do pandas um dia incorporar isso tudo!
