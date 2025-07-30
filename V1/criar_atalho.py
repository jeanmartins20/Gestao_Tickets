# Criar Atalho para

import os
import sys
import winshell
from win32com.client import Dispatch

def criar_atalho():
    desktop = winshell.desktop()
    caminho_app = os.path.abspath("app.py")
    caminho_python = sys.executable
    atalho_path = os.path.join(desktop, "Gestão de Tickets.lnk")
    icon_path = os.path.abspath("ticket_icon.ico")

    shell = Dispatch('WScript.Shell')
    atalho = shell.CreateShortCut(atalho_path)
    atalho.TargetPath = caminho_python
    atalho.Arguments = f'"{caminho_app}"'
    atalho.WorkingDirectory = os.path.dirname(caminho_app)
    atalho.IconLocation = icon_path
    atalho.save()

if __name__ == "__main__":
    criar_atalho()
    print("✅ Atalho criado na Área de Trabalho.")
