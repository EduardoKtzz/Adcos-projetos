import customtkinter as ctk
from tkinter import messagebox
import webbrowser
import pythoncom
from automocao import rodar_automacao
from threading import Thread
import traceback

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Função para abrir o link do procedimento
def abrir_procedimento():
    url = ("https://adcos.sharepoint.com/:w:/r/sites/ccad/Documentos%20Compartilhados/Cadastro%20de%20Clientes%20Profisional/VTEX/Atualizar%20planilha/Procedimento.docx?d=wc097263fee0b46f0a565da13ad600372&csf=1&web=1&e=J0SBU2")
    webbrowser.open(url)

# Criar a janela principal
app = ctk.CTk()
app.geometry("300x250")
app.title("Atualização de clientes")

# Função chamada ao clicar em 'Entrar'
def login():
    email = entry_email.get()
    senha = entry_senha.get()

    if not email or not senha:
        messagebox.showerror("Erro", "Por favor, preencha todos os campos.", parent=app)
        if not email:
            entry_email.focus()
        elif not senha:
            entry_senha.focus()
        return

    def executar():
        pythoncom.CoInitialize()
        try:
            botao_login.configure(state="disabled", text="Executando...")
            rodar_automacao(email, senha)
            app.after(0, lambda: messagebox.showinfo("Sucesso", "Automação executada com sucesso!", parent=app))

        except Exception as e:
            import traceback
            erro_completo = traceback.format_exc()
            with open("erro.txt", "w", encoding="utf-8") as f:
                f.write(erro_completo)

            erro_str = str(e)
            app.after(0, lambda: messagebox.showerror("Erro na automação", erro_str, parent=app))

        finally:
            app.after(0, lambda: botao_login.configure(state="normal", text="Entrar"))
            pythoncom.CoUninitialize()

    Thread(target=executar, daemon=True).start()

# Widgets
label_titulo = ctk.CTkLabel(app, text="Login com a conta VTEX", font=("Arial", 20))
label_titulo.pack(pady=20)

entry_email = ctk.CTkEntry(app, placeholder_text="Email", width=230)
entry_email.pack(pady=5)

entry_senha = ctk.CTkEntry(app, placeholder_text="Senha", show="*", width=230)
entry_senha.pack(pady=5)

botao_login = ctk.CTkButton(app, text="Entrar", command=login)
botao_login.pack(pady=5)

botao_instrucao = ctk.CTkButton(app, text="Como usar?", command=abrir_procedimento)
botao_instrucao.pack(pady=5)

# Captura qualquer erro geral (fora das threads)
try:
    app.mainloop()
except Exception as e:
    erro = traceback.format_exc()
    with open("erro_fatal.txt", "w", encoding="utf-8") as f:
        f.write(erro)
    messagebox.showerror("Erro crítico", f"Ocorreu um erro inesperado:\n\n{str(e)}")
