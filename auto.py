import os
import time
import threading
import pyautogui
import pyperclip
import win32com.client as win32
from PyPDF2 import PdfReader
import tkinter as tk
import pythoncom  # IMPORTANTE para rodar COM dentro da thread
import re

# Caminho da pasta dos boletos

PASTA_BOLETOS = r'C:\ATX'

def pegar_ultimo_boleto():
    arquivos = [os.path.join(PASTA_BOLETOS, f) for f in os.listdir(PASTA_BOLETOS) if f.endswith('.pdf')]
    if not arquivos:
        print("❌ Nenhum boleto PDF encontrado em C:\\ATX")
        return None
    return max(arquivos, key=os.path.getctime)

def extrair_nome_cliente(caminho_pdf):
    try:
        leitor = PdfReader(caminho_pdf)
        texto = leitor.pages[0].extract_text()

        for linha in texto.split('\n'):
            if "Pagador:" in linha:
                nome = linha.split("Pagador:")[1].strip()
                nome = nome.split('-')[0].strip()
                return nome
        return "Cliente"
    except Exception as e:
        print("❌ Erro ao ler o PDF:", e)
        return "Cliente"

def copiar_email_com_pyautogui(x, y):
    pyautogui.click(x, y, clicks=2, interval=0.1)  # clique triplo
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)
    email = pyperclip.paste().strip().replace('\n', ' ')
    return email


def enviar_email(destinatario, nome_cliente, caminho_pdf, numero_boleto):
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    email.Display()
    email.Subject = f"BOLETO - {numero_boleto} - {nome_cliente}"
    email.To = destinatario
    email.Attachments.Add(caminho_pdf)

    corpo = f"""Olá {nome_cliente},

Segue em anexo o boleto referente ao seu faturamento.

Qualquer dúvida, estamos à disposição.
"""

    email.HTMLBody = corpo.replace('\n', '<br>') + email.HTMLBody
    email.Send()
    print(f"E-mail enviado para {destinatario}")


def executar_automacao():
    pythoncom.CoInitialize()

    status_label.config(text="Buscando boleto...")
    boleto = pegar_ultimo_boleto()
    numero_boleto = os.path.basename(boleto).split('_')[0]
    if not boleto:
        status_label.config(text="❌ Nenhum boleto encontrado.")
        return

    nome_cliente = extrair_nome_cliente(boleto)
    print(f"Boleto: {os.path.basename(boleto)}")
    print(f"Cliente: {nome_cliente}")

    status_label.config(text="Posicione o mouse sobre o campo de e-mail no XPERT (5s)...")
    #cotagem regressiva para o usuário posicionar o mouse
    for i in range(5, 0, -1):
        status_label.config(text=f"Posicione o mouse sobre o campo de e-mail no XPERT ({i}s)...")
        print(f"Mova o mouse para o campo de e-mail no XPERT... ({i}s)")
        time.sleep(1)

    time.sleep(5)

    email_cliente = copiar_email_com_pyautogui(x=947, y=424)
    print(f"Conteúdo copiado: {email_cliente}")

    if not email_cliente:
        status_label.config(text="❌ Nenhum e-mail copiado!")
        print("❌ Nenhum e-mail copiado do XPERT!")
        return

    # Verificação com expressão regular para pegar todos os e-mails
    regex_email = r'[\w\.-]+@[\w\.-]+\.\w+'
    emails_encontrados = re.findall(regex_email, email_cliente)

    if not emails_encontrados:
        status_label.config(text="❌ Nenhum e-mail válido copiado!")
        print("❌ Nenhum e-mail válido copiado do XPERT!")
        return

    emails_validos = ';'.join(emails_encontrados)
    print(f"E-mails validados: {emails_validos}")
    status_label.config(text="Enviando e-mail...")
    
    try:
        enviar_email(emails_validos, nome_cliente, boleto, numero_boleto)
        status_label.config(text=f"E-mail enviado para {emails_validos}")
    except Exception as e:
        status_label.config(text="❌ Erro ao enviar e-mail.")
        print("❌ Erro ao enviar:", e)

def iniciar_thread():
    threading.Thread(target=executar_automacao).start()

# Interface Gráfica
janela = tk.Tk()
janela.title("Robô de Envio Automático")
janela.geometry("450x300")

titulo = tk.Label(janela, text="🤖 Robô de Envio Automático", font=("Arial", 20, "bold"))
titulo.pack(pady=10)

btn = tk.Button(janela, text="Enviar", font=("Arial", 12), bg="#09FF00", fg="white", width=20, command=iniciar_thread)
btn.pack(pady=10)

barra_lateral = tk.Frame(janela, width=250, bg="#f0f0f0")
barra_lateral.pack(side=tk.RIGHT, fill=tk.Y)

data = time.strftime("%d/%m/%Y")
rodape = tk.Label(janela, text=f"Desenvolvido por Cainã - {data}", font=("Arial", 8))
rodape.pack(side=tk.BOTTOM, pady=5)



status_label = tk.Label(janela, text="", font=("Arial", 10), wraplength=280)
status_label.pack(pady=20)

janela.mainloop()
