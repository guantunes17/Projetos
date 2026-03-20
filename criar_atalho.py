"""
Execute este script UMA VEZ para criar o atalho da área de trabalho.
Depois pode deletar este arquivo e o .bat.
"""
import os
import sys

def criar_atalho():
    try:
        import winshell
        from win32com.client import Dispatch
    except ImportError:
        # Instala dependências se necessário
        os.system('pip install winshell pywin32 -q')
        import winshell
        from win32com.client import Dispatch

    # Caminhos
    pasta_projeto = r'C:\Users\reser\Documents\Codigos Python\Projetos'
    script        = os.path.join(pasta_projeto, 'central_relatorios.py')
    icone         = os.path.join(pasta_projeto, 'central_relatorios_icone.ico')
    area_trabalho = winshell.desktop()
    atalho_path   = os.path.join(area_trabalho, 'Central de Relatórios.lnk')

    # Encontra pythonw.exe (roda sem janela de terminal)
    pythonw = os.path.join(os.path.dirname(sys.executable), 'pythonw.exe')
    if not os.path.isfile(pythonw):
        pythonw = sys.executable  # fallback para python.exe

    # Cria atalho
    shell   = Dispatch('WScript.Shell')
    atalho  = shell.CreateShortCut(atalho_path)
    atalho.Targetpath      = pythonw
    atalho.Arguments       = f'"{script}"'
    atalho.WorkingDirectory = pasta_projeto
    atalho.IconLocation    = icone
    atalho.Description     = 'Central de Relatórios — Baia 4 Logística'
    atalho.save()

    print(f'✅ Atalho criado em: {atalho_path}')
    print('   Pode deletar o .bat da área de trabalho agora!')
    input('\nPressione Enter para fechar...')

criar_atalho()