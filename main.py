# importações de módulos
from sap_session import open_sap_session, close_sap_session
import os
import time
import sys


# função principal
def main():
    # Verifica ambiente antes de executar
    # if not verificar_ambiente():
    #     return
    # Inicia a sessão SAP
    session = open_sap_session()
    # Se a sessão não for iniciada, encerra o programa
    if not session:
        return

    # Função para garantir que estamos no menu principal do SAP
    # def para criar funções - nome_função(parâmetros):
    def garantir_menu_principal(session):
        # procura pelo id wnd[0]/tbar[0]/okcd e insere o comando /n para voltar ao menu principal
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        # procura pelo id wnd[0] e envia a tecla Enter (VKey 0)
        session.findById("wnd[0]").sendVKey(0)

    # Exemplo de função para realizar uma ação no SAP
    def exemplo_acao(session):
        # procura pelo id wnd[0]/tbar[0]/okcd e insere o comando se16n
        session.findById("wnd[0]/tbar[0]/okcd").text = "se16n"
        # procura pelo id wnd[0] e envia a tecla Enter (VKey 0)
        session.findById("wnd[0]").sendVKey(0)

    # Exemplo de chamadadas das funções
    garantir_menu_principal(session)
    time.sleep(5)  # espera 5 segundo para garantir que a tela foi carregada
    exemplo_acao(session)
    time.sleep(5)
    garantir_menu_principal(session)

    # encerrar a sessão SAP
    close_sap_session()
    print("\nProcessos SAP encerrados.")


if __name__ == "__main__":
    main()
