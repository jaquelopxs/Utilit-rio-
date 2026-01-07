import gspread
import pandas as pd
import os
from dotenv import load_dotenv
import sys


ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

def obter_path_env():
    if getattr(sys, 'frozen', False):
        return os.path.join(os.path.dirname(sys.executable), ".env")
    else:
        return os.path.join(os.path.dirname(__file__), "..", ".env")
    
ENV_PATH = obter_path_env()

load_dotenv(dotenv_path=ENV_PATH, override=True)

credencials_file = os.getenv("GS_CREDENTIALS_FILE")
spreadsheet_id = os.getenv("GS_SPREADSHEET_ID")



def conectar():
    """
    Conecta-se ao Google Sheets usando a Conta de Serviço e retorna o DataFrame.
    Retorna um DataFrame vazio em caso de erro.
    """
    try:
        print("Tentando autenticar e conectar ao Google Sheets...")
        gc = gspread.service_account(filename=credencials_file)
        spreadsheet = gc.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.sheet1
        data = worksheet.get_all_values()
        
        # Cria o DataFrame
        df = pd.DataFrame(data[1:], columns=data[0], dtype=str)
        print(f"Conexão SUCESSO! DataFrame carregado com {len(df)} linhas.")
        
        # IMPORTANTE: Retorna o DF
        return df

    except gspread.exceptions.APIError as e:
        if '403' in str(e):
            print(f"\nERRO DE PERMISSÃO (403): O acesso foi negado.")
            print(">>> Ação: Verifique se o 'client_email' do JSON foi adicionado como 'Leitor' na planilha.")
        else:
            print(f"\nERRO DE API: {e}")
        return pd.DataFrame() # Retorna um DF vazio em caso de erro

    except gspread.exceptions.SpreadsheetNotFound:
        print(f"\nERRO: Planilha não encontrada. Verifique se o ID está correto.")
        return pd.DataFrame()
        
    except Exception as e:
        print(f"\nOcorreu um erro inesperado: {e}")
        return pd.DataFrame()
    
def salvar_status(df):
    try:
        gc = gspread.service_account(filename=credencials_file)
        sh = gc.open_by_key(spreadsheet_id)
        ws = sh.sheet1

        valores = [df.columns.tolist()] + df.fillna("").values.tolist()
        ws.update(valores)

        print("✓ Planilha do Google Sheets atualizada com sucesso.")

    except Exception as e:
        print(f"Erro ao atualizar Google Sheets: {e}")