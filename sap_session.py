# Importações necessárias para automação SAP
# from dotenv import load_dotenv # Para carregar variáveis de ambiente de um arquivo .env
# import os # Para manipulação de variáveis de ambiente
import subprocess  # Para executar comandos do sistema operacional
import time  # Para controle de tempo/delays
import win32com.client as win32  # Para comunicação COM com SAP GUI

def open_sap_session():
    """
    Abre uma nova sessão SAP.

    Processo:
    1. Executa comando sapshcut para iniciar SAP com credenciais
    2. Aguarda carregamento da interface
    3. Conecta via COM ao SAP GUI
    4. Obtém referência da sessão ativa
    5. Maximiza a janela principal

    Returns:
        session: Objeto da sessão SAP ativa ou None em caso de erro
    """
    try:
        # Executa comando SAP shortcut com parâmetros de login
        # sapshcut: utilitário para abrir SAP com parâmetros pré-definidos
        # -user: usuário de login
        # -pw: senha do usuário
        # -language: idioma da interface (PT = Português)
        # -system: sistema SAP de destino (GNQ)
        # -client: cliente/mandante SAP (400)
        subprocess.call(
            "START sapshcut -user=jcustodio -pw=83639090Ju@ -language=PT -system=GNP -client=400",
            shell=True,
        )
        # RECOMENDAÇÃO: Para maior segurança, mova as credenciais para arquivo .env:
        # load_dotenv()
        # subprocess.call(f"START sapshcut -user={os.getenv('SAP_USER')} -pw={os.getenv('SAP_PASSWORD')} -language={os.getenv('SAP_LANGUAGE')} -system={os.getenv('SAP_SYSTEM')} -client={os.getenv('SAP_CLIENT')}", shell=True)

        # Aguarda 5 segundos para o SAP carregar completamente
        time.sleep(5)

        # Conecta ao SAP GUI
        sap_gui = win32.GetObject("SAPGUI")

        # Obtém o mecanismo de scripting do SAP
        app = sap_gui.GetScriptingEngine

        # Pega a primeira conexão ativa (índice 0)
        connection = app.Children(0)

        # Pega a primeira sessão da conexão (índice 0)
        session = connection.Children(0)

        # Maximiza a janela principal do SAP (wnd[0] = janela principal)
        session.findById("wnd[0]").maximize()

        print("Sessão SAP iniciada com sucesso.")
        return session

    except Exception as e:
        # Captura qualquer erro durante o processo de conexão
        print(
            f"Não foi possível iniciar a sessão SAP. Verifique se o SAP Logon está instalado. Erro: {e}"
        )
        return None


def close_sap_session():
    """
    Encerra todos os processos relacionados ao SAP do usuário atual.

    Mata os processos:
    - saplogon.exe: Interface principal do SAP Logon
    - saplgpad.exe: SAP Logon Pad (gerenciador de conexões)

    Usa taskkill com filtros para encerrar apenas processos do usuário atual.
    """
    print("Encerrando processos SAP...")

    # Encerra processo SAP Logon do usuário atual
    # /f: força o encerramento
    # /fi: aplica filtro
    # "USERNAME eq %username%": apenas processos do usuário atual
    # /im: especifica o nome da imagem/executável
    subprocess.call(
        f'taskkill /f /fi "USERNAME eq %username%" /im "saplogon.exe"', shell=True
    )

    # Encerra processo SAP Logon Pad do usuário atual
    subprocess.call(
        f'taskkill /f /fi "USERNAME eq %username%" /im "saplgpad.exe"', shell=True
    )
