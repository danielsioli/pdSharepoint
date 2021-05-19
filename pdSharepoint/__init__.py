from os import remove
from os.path import dirname, abspath, isfile
from O365 import Account, MSGraphProtocol, FileSystemTokenBackend
from sys import platform
import pandas as pd

def get_path_name(obj_name, io=None):
    """ Returns a path and a name for where the downloaded objects (list or excel file) will be stored.
    :param str obj_name: a list name or a path to an excel file in sharepoint
    :param str io: path to where downloaded objects (list or excel file) will be stored.
    :return: a tuple with a path and a name
    :rtype: tuple
    """
    # obter o caracter de split do filepath
    split_char = '/'
    if 'linux' in platform:
        path = dirname(abspath(__file__))
    elif 'win' in platform:
        split_char = '\\'
        path = abspath('')

    # obter to_path e name
    if not io:
        to_path = path + split_char
        name = obj_name.split('/')[-1]
    else:
        to_path = split_char.join(io.split(split_char)[:-1]) + split_char
        name = io.split(split_char)[-1]
    return split_char, to_path, name

def get_team(obj_name, credentials, host, site,
              token_filepath=None, io=None):
    """ Returns a Sharepoint site object
    :param str obj_name: a list name or a path to an excel file in sharepoint
    :param tuple credentials: a tuple with Office 365 client_id and client_secret
    :param str host: A sharepoint host like anatel365.sharepoint.com
    :param str site: a site's name or team's name
    :param str token_filepath: path where the Office 365 token is or will be stored.
    :param str io: path to where downloaded objects (list or excel file) will be stored.
    :return: a Sharepoint Site
    :rtype: Site
    """

    # obter o caracter de split do filepath
    split_char, to_path, name = get_path_name(obj_name=obj_name,io=io)

    # separar o token_filename do token_path
    if not token_filepath:
        token_filename = 'o365token.txt'
        token_path = to_path + split_char
    else:
        token_filename = token_filepath.split(split_char)[-1]
        token_path = split_char.join(token_filepath.split(split_char)[:-1]) + split_char

    # Obter o token
    token_backend = FileSystemTokenBackend(token_path=token_path, token_filename=token_filename)
    # criar o objeto account
    account = Account(credentials, token_backend=token_backend)
    # scope para sharepoint
    scopes = ['basic'] + MSGraphProtocol().get_scopes_for(user_provided_scopes='sharepoint')

    # Autenticação
    if not isfile(token_path + token_filename):
        if not account.authenticate(scopes=scopes):
            raise Exception('Não foi possível autenticar com as credenciais e token informados')

    sharepoint = account.sharepoint()
    team = sharepoint.get_site(host, 'sites/' + site)
    return team

def download_sharepoint_excel(file_path, host, site, library, credentials,
                   token_filepath=None, io=None):
    """ Downloads a Sharepoint excel file to local storage
    :param str file_path: a file path to an excel file in sharepoint
     within a site's library (i.e. /General/file.xlsx)
    :param str host: A sharepoint host like anatel365.sharepoint.com
    :param str site: a site's name or team's name
    :param str library: a library in the Sharepoint site (i.e. Documents)
    :param tuple credentials: a tuple with Office 365 client_id and client_secret
    :param str token_filepath: path where the Office 365 token is or will be stored.
    :param str io: path to where downloaded objects (list or excel file) will be stored.
    """
    #Obtendo a Equipe
    team = get_team(obj_name=file_path, credentials=credentials,
                    host=host, site=site, token_filepath=token_filepath,
                    io=io)
    split_char, to_path, name = get_path_name(obj_name=file_path, io=io)

    drives = team.list_document_libraries()
    team_drive = None
    for drive in drives:
        if library == drive.name:
            team_drive = drive
            break
    if team_drive:
        arquivo = team_drive.get_item_by_path(file_path)
        arquivo.download(to_path=to_path, name=name)
        return to_path + name
    else:
        raise ValueError(f'Library {library} not found!')

def read_sharepoint_list(list_name, host, site, credentials,
              token_filepath=None):
    """ Returns a panda DataFrame with the contents of a Sharepoint list
    :param str list_name: a list name in sharepoint
    :param str host: A sharepoint host like anatel365.sharepoint.com
    :param str site: a site's name or team's name
    :param tuple credentials: a tuple with Office 365 client_id and client_secret
    :param str token_filepath: path where the Office 365 token is or will be stored.
    :return: a pandas DataFrame with the Sharepoint list's contents.
    :rtype: pandas.DataFrame
    """
    # Obtendo a Equipe
    team = get_team(obj_name=list_name, credentials=credentials,
                                             host=host, site=site, token_filepath=token_filepath,
                                             io=None)

    lt = team.get_list_by_name(list_name)
    list_columns = lt.get_list_columns()
    no_columns = ['ComplianceAssetId', 'Author', 'Editor', 'DocIcon',
                  '_IsRecord', 'AppAuthor', 'AppEditor']

    # Cria lista de dicionários para criar um pandas DataFrame
    items = lt.get_items()
    id_list = []
    rows = []
    for item in items:
        id_list.append(item.object_id)

    for item_id in id_list:
        fields = lt.get_item_by_id(item_id).fields
        fields.pop('@odata.etag', None)
        for list_column in list_columns:
            if list_column.display_name != list_column.internal_name\
                    and list_column.internal_name not in no_columns:
                if 'Title' in list_column.internal_name:
                    if list_column.internal_name == 'Title':
                        fields[list_column.display_name] = fields[list_column.internal_name]
                        fields.pop(list_column.internal_name, None)
                else:
                    fields[list_column.display_name] = fields[list_column.internal_name]
                    fields.pop(list_column.internal_name, None)

        rows.append(fields)
    df = pd.DataFrame(rows)
    #Tenta transformar os campos de data em datetime
    df = df.select_dtypes(include=object).apply(pd.to_datetime, errors='ignore').combine_first(df)
    return df

def read_sharepoint_excel(file_path, host, site, library, credentials,
               token_filepath=None, **kwargs):
    """ Returns a panda DataFrame with the contents of a Sharepoint list
    :param str file_path: a file path to an excel file in sharepoint
     within a site's library (i.e. /General/file.xlsx)
    :param str host: A sharepoint host like anatel365.sharepoint.com
    :param str site: a site's name or team's name
    :param str library: a library in the Sharepoint site (i.e. Documents)
    :param tuple credentials: a tuple with Office 365 client_id and client_secret
    :param str token_filepath: path where the Office 365 token is or will be stored.
    :param kwargs: every param from pandas.read_excel, except io.
    (kwargs)
    :return: a pandas DataFrame with the Sharepoint excel file's contents.
    :param
    :rtype: pandas.DataFrame
    """
    #Argumentos do pd.read_excel
    sheet_name = kwargs.pop('sheet_name', 0)
    header = kwargs.pop('header', 0)
    names = kwargs.pop('names', None)
    index_col = kwargs.pop('index_col', None)
    usecols = kwargs.pop('usecols', None)
    squeeze = kwargs.pop('squeeze', False)
    dtype = kwargs.pop('dtype', None)
    engine = kwargs.pop('engine', None)
    converters = kwargs.pop('converters', None)
    true_values = kwargs.pop('true_values', None)
    false_values = kwargs.pop('false_values', None)
    skiprows = kwargs.pop('skiprows', None)
    nrows = kwargs.pop('nrows', None)
    na_values = kwargs.pop('na_values', None)
    keep_default_na = kwargs.pop('keep_default_na', True)
    na_filter = kwargs.pop('na_filter', True)
    verbose = kwargs.pop('verbose', False)
    parse_dates = kwargs.pop('parse_dates', False)
    date_parser = kwargs.pop('date_parser', None)
    thousands = kwargs.pop('thousands', None)
    comment = kwargs.pop('comment', None)
    skipfooter = kwargs.pop('skipfooter', 0)
    convert_float = kwargs.pop('convert_float', True)
    mangle_dupe_cols = kwargs.pop('mangle_dupe_cols', True)
    storage_options = kwargs.pop('storage_options', None)

    #Download do arquivo do Sharepoint
    io = download_sharepoint_excel(file_path=file_path, host=host, site=site,library=library,
                        credentials=credentials, token_filepath=token_filepath)

    #Ler o arquivo em um pandas
    df = pd.read_excel(io=io, sheet_name=sheet_name, header=header, names=names,
                       index_col=index_col, usecols=usecols, squeeze=squeeze,
                       dtype=dtype, engine=engine, converters=converters,
                       true_values=true_values, false_values=false_values,
                       skiprows=skiprows, nrows=nrows, na_values=na_values,
                       keep_default_na=keep_default_na, na_filter=na_filter,
                       verbose=verbose, parse_dates=parse_dates, date_parser=date_parser,
                       thousands=thousands, comment=comment, skipfooter=skipfooter,
                       convert_float=convert_float, mangle_dupe_cols=mangle_dupe_cols,
                       storage_options=storage_options)

    #Apagar arquivo temporário
    remove(io)

    #Retornar um Pandas DataFrame
    return df