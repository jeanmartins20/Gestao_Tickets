import os
import sys
import winshell
from win32com.client import Dispatch

def criar_atalho():
    desktop = winshell.desktop()
    # Caminho do aplicativo principal, ajustado para o novo nome do arquivo
    caminho_app = os.path.abspath("App_Gestão_V3_Grafico_Modernizado.py")
    caminho_python = sys.executable
    atalho_path = os.path.join(desktop, "Gestão de Tickets V3.lnk")
    # Certifique-se de que o ícone 'ticket_icon (2).ico' está no mesmo diretório
    # ou forneça o caminho completo para ele.
    icon_path = os.path.abspath("ticket_icon (2).ico") 

    shell = Dispatch('WScript.Shell')
    atalho = shell.CreateShortCut(atalho_path)
    atalho.TargetPath = caminho_python
    atalho.Arguments = f'"{caminho_app}"'
    atalho.WorkingDirectory = os.path.dirname(caminho_app)
    atalho.IconLocation = icon_path
    atalho.save()

if __name__ == "__main__":
    criar_atalho()
    print("Atalho criado com sucesso na Área de Trabalho.")