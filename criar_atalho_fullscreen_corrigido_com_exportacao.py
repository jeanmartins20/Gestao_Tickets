import os
import sys
import winshell
from win32com.client import Dispatch

# Garante que o terminal aceite UTF-8 se suportado
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass  # Ignora se a função não estiver disponível (Python < 3.7)

def criar_atalho():
    desktop = winshell.desktop()
    # O caminho do aplicativo principal já está correto, apontando para o arquivo mais recente.
    caminho_app = os.path.abspath("app_corrigido_fullscreen_corrigido.py")
    caminho_python = sys.executable
    atalho_path = os.path.join(desktop, "Gestão de Tickets.lnk")
    # Certifique-se de que o arquivo 'ticket_icon (2).ico' está no mesmo diretório do script.
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
