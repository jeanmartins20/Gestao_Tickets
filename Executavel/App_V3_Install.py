import os
import sys
import winshell
from win32com.client import Dispatch
from tkinter import messagebox

def criar_atalho():
    desktop = winshell.desktop()
    
    # Obter o diretório atual do script de atalho
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Caminho para o executável gerado pelo PyInstaller.
    # Ele estará dentro da pasta 'dist' que o PyInstaller cria.
    caminho_executavel = os.path.join(script_dir, "dist", "App_V3_Modernizado_corrigido.exe")
    
    # Caminho para o arquivo de ícone.
    icon_path = os.path.join(script_dir, "ticket_icon (2).ico") 

    atalho_path = os.path.join(desktop, "Gestão de Tickets V3.lnk")

    # Verificar se o executável existe antes de tentar criar o atalho
    if not os.path.exists(caminho_executavel):
        messagebox.showerror(
            "Erro ao Criar Atalho", 
            f"O arquivo executável não foi encontrado em '{caminho_executavel}'.\n"
            "Por favor, certifique-se de ter gerado o .exe com PyInstaller antes de rodar este script."
        )
        return

    try:
        shell = Dispatch('WScript.Shell')
        atalho = shell.CreateShortCut(atalho_path)
        atalho.TargetPath = caminho_executavel  # Aponta diretamente para o executável
        atalho.Arguments = ''
        atalho.WorkingDirectory = os.path.dirname(caminho_executavel)
        atalho.IconLocation = icon_path
        atalho.save()
        messagebox.showinfo("Atalho Criado", "Atalho 'Gestão de Tickets V3.lnk' criado com sucesso na Área de Trabalho.")
    except Exception as e:
        messagebox.showerror("Erro ao Criar Atalho", f"Ocorreu um erro ao criar o atalho: {e}")


if __name__ == "__main__":
    criar_atalho()

