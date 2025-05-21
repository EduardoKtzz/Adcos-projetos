import customtkinter as ctk
from PIL import Image
import webbrowser
from concurrent.futures import ThreadPoolExecutor
import threading
from kitsOficial import preencherKits
from KitsData import alterar_data

executor = ThreadPoolExecutor(max_workers=2)  # Define um limite de threads ativas

def iniciar(log_area, entrada_kit):
    def executar():
        try:
            kit_codigo = entrada_kit.get()
            if not kit_codigo.isdigit():
                kit_codigo = 0
            else:
                kit_codigo = int(kit_codigo)

            log(f"🚀 Iniciando a automação...", log_area)
            preencherKits(log, log_area, kit_codigo)
        except Exception as e:
            log(f"❌ Erro: {e}", log_area)

        log(f"🚀 Iniciando a automação...", log_area)
        preencherKits(log, log_area, kit_codigo)  # Passa o kit como argumento

    executor.submit(executar)  # Submete a tarefa para execução em uma thread


def iniciar_data(log_area):
    def executar_data():
        log(f"🚀 Iniciando a automação...", log_area)
        alterar_data(log, log_area)

    thread = threading.Thread(target=executar_data)
    thread.start()

def log(mensagem, log_area):

    log_area.after(0, lambda: log_area.insert("end", mensagem + "\n"))
    log_area.after(0, lambda: log_area.see("end"))  # Rola automaticamente até o final

# Função para exibir a tela de "Alterar Data"
def mostrar_alterar_data(frame_principal):
    for widget in frame_principal.winfo_children():
        widget.destroy()

    ctk.CTkLabel(frame_principal, text="Área de alteração de datas", font=("Arial", 24, "bold"), text_color="white").pack(pady=10)

    ctk.CTkLabel(frame_principal, text="Preencha as informações na planilha para realizar a alteração de data.", font=("Arial", 16), text_color="white").pack(pady=10)

 #  Área de logs 
    ctk.CTkLabel(frame_principal, text="Logs do Sistema:", font=("Arial", 16), text_color="white").pack(pady=(20, 5))
    log_area = ctk.CTkTextbox(frame_principal, width=600, height=200, wrap="word")
    log_area.pack(padx=20, pady=10, fill="both", expand=True)

    # Função para baixar a planilha
    def planilha_modelo2():
        url = r"https://adcos.sharepoint.com/:x:/r/sites/ccad/Documentos%20Compartilhados/Eduardo%20Klitzke/Kits/Kits.xlsx?d=w94136bba54da4250ba2a638ae01d3b42&csf=1&web=1&e=HGnnat&nav=MTVfe0ZFQTYxMUFELTY2QjItNDREQi05NDkzLTE4NDU5QzczMjRBRn0"  # Altere para o link correto
        webbrowser.open(url)
        log("Planilha de modelo aberta com sucesso!\n", log_area)

    # Botão de baixar planilha modelo
    ctk.CTkButton(frame_principal, text="Entrar na planilha modelo", command=planilha_modelo2, fg_color="#3498db", hover_color="#2980b9", corner_radius=10, width=250).pack(pady=(5, 15))

    #Botão de exec
    ctk.CTkButton(frame_principal, text="Começar a alteração", command=lambda: iniciar_data(log_area), fg_color="#3498db", hover_color="#2980b9", corner_radius=10, width=250).pack(pady=(0, 10))

    # Botão para voltar para o menu de edição
    ctk.CTkButton(frame_principal, text="Voltar", command=lambda: mostrar_editar(frame_principal), fg_color="#e74c3c", hover_color="#c0392b", corner_radius=10).pack(pady=20)

# Função para exibir a tela de "Adicionar Tabela"
def mostrar_adicionar_tabela(frame_principal):
    for widget in frame_principal.winfo_children():
        widget.destroy()

    ctk.CTkLabel(frame_principal, text="Área para adicionar tabelas", font=("Arial", 24, "bold"), text_color="white").pack(pady=10)
    ctk.CTkLabel(frame_principal, text="Preencha as informações na planilha para realizar adição de tabela. Lembrando que aqui só conseguimos adicionar duas tabelas(A72 e A660)", font=("Arial", 16), wraplength=400,  text_color="white").pack(pady=10)


 #  Área de logs 
    ctk.CTkLabel(frame_principal, text="Logs do Sistema:", font=("Arial", 16), text_color="white").pack(pady=(20, 5))
    log_area = ctk.CTkTextbox(frame_principal, width=600, height=200, wrap="word")
    log_area.pack(padx=20, pady=10, fill="both", expand=True)

    # Função para baixar a planilha
    def planilha_modelo3():
        url = r"https://adcos.sharepoint.com/:x:/r/sites/ccad/Documentos%20Compartilhados/Eduardo%20Klitzke/Kits/Kits.xlsx?d=w94136bba54da4250ba2a638ae01d3b42&csf=1&web=1&e=iyWPjW&nav=MTVfezVGQUFCQUMwLTkxNzItNEUwMi1CNTBFLTg2Q0Q0OUUxQkVBOX0"  # Altere para o link correto
        webbrowser.open(url)
        log("Planilha de modelo aberta com sucesso!\n", log_area)

    # Botão de baixar planilha modelo
    ctk.CTkButton(frame_principal, text="Entrar na planilha modelo", command=planilha_modelo3, fg_color="#3498db", hover_color="#2980b9", corner_radius=10, width=250).pack(pady=(5, 15))

    #Botão de exec
    ctk.CTkButton(frame_principal, text="Começar a criação", command=lambda: iniciar_data(log_area), fg_color="#3498db", hover_color="#2980b9", corner_radius=10, width=250).pack(pady=(0, 10))

    # Botão para voltar para o menu de edição
    ctk.CTkButton(frame_principal, text="Voltar", command=lambda: mostrar_editar(frame_principal), fg_color="#e74c3c", hover_color="#c0392b", corner_radius=10).pack(pady=20)

# Função para exibir o menu de "EDITAR"
def mostrar_editar(frame_principal):
    for widget in frame_principal.winfo_children():
        widget.destroy()

    ctk.CTkLabel(frame_principal, text="Selecione a opção desejada:", font=("Arial", 24, "bold"), text_color="white").pack(pady=20)

    # Botões para acessar as opções
    ctk.CTkButton(frame_principal, text="Alterar Data", command=lambda: mostrar_alterar_data(frame_principal), fg_color="#3498db", hover_color="#2980b9", corner_radius=10, width=150, height=40).pack(pady=5)
    ctk.CTkButton(frame_principal, text="Adicionar Tabela", command=lambda: mostrar_adicionar_tabela(frame_principal), fg_color="#3498db", hover_color="#2980b9", corner_radius=10, width=150, height=40).pack(pady=5)

# Função para alternar o conteúdo no painel principal
def mostrar_conteudo(opcao, frame_principal):
    for widget in frame_principal.winfo_children():
        widget.destroy()

    if opcao == "CRIAR":

        ctk.CTkLabel(frame_principal, text="Criação de Kit", font=("Arial", 24, "bold"), text_color="white").pack(pady=20)

        # Campo para inserir o número do kit
        ctk.CTkLabel(frame_principal, text="Número do Kit:", font=("Arial", 16), text_color="white").pack(pady=(0, 5))
        numero_kit = ctk.CTkEntry(frame_principal, width=100, height=40)
        numero_kit.pack(pady=(0, 10))

        ctk.CTkLabel(frame_principal, text="Preencha a planilha modelo antes de iniciar o programa", font=("Arial", 14), text_color="white").pack(pady=(0, 5))

        # Área de logs 
        ctk.CTkLabel(frame_principal, text="Logs do Sistema:", font=("Arial", 16), text_color="white").pack(pady=(20, 5))
        log_area = ctk.CTkTextbox(frame_principal, width=600, height=200, wrap="word")
        log_area.pack(padx=20, pady=10, fill="both", expand=True)

        # Função para baixar a planilha
        def planilha_modelo():
            url = "https://adcos.sharepoint.com/:x:/r/sites/ccad/Documentos%20Compartilhados/Eduardo%20Klitzke/Kits/Kits.xlsx?d=w94136bba54da4250ba2a638ae01d3b42&csf=1&web=1&e=G0Gx6g"  # Altere para o link correto
            webbrowser.open(url)
            log("Planilha de modelo aberta com sucesso!\n", log_area)

        # Botão de baixar planilha modelo
        ctk.CTkButton(frame_principal, text="Entrar na planilha modelo", command=planilha_modelo, fg_color="#3498db", hover_color="#2980b9", corner_radius=10, width=250).pack(pady=(5, 15))

        #Botão de exec
        ctk.CTkButton(frame_principal, text="Começar a criação", command=lambda: iniciar(log_area, numero_kit), fg_color="#3498db", hover_color="#2980b9", corner_radius=10, width=250).pack(pady=(0, 10))

    elif opcao == "EDITAR":
        ctk.CTkLabel(frame_principal, text="Selecione a opção desejada:", font=("Arial", 24, "bold"), text_color="white").pack(pady=20)

    else:
        ctk.CTkLabel(frame_principal, text="Bem-vindo ao Sistema!", font=("Arial", 24, "bold"), text_color="white").pack(pady=20)

# Função para abrir o link do Word
def abrir_procedimento():
    url = r"https://adcos.sharepoint.com/:w:/r/sites/ccad/Documentos%20Compartilhados/Eduardo%20Klitzke/Kits/PROCEDIMENTO.docx?d=w171190a255434c16913f9c6aed266a1c&csf=1&web=1&e=Tt2duO"  # Insira o link do seu documento
    webbrowser.open(url)

# Função principal
def criar_interface():
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")

    app = ctk.CTk()
    app.title("Cadastro de Kit")
    app.geometry("900x600")
    app.resizable(False, False)  # Impede o redimensionamento

    # Barra lateral (Esquerda) com tamanho fixo
    sidebar = ctk.CTkFrame(app, width=220, height=600, fg_color="#2e2e2e", corner_radius=10)
    sidebar.pack_propagate(False)  # Impede ajuste automático
    sidebar.pack(side="left", padx=(5, 0), pady=5)

    # Área principal (Direita) com tamanho fixo
    frame_principal = ctk.CTkFrame(app, width=680, height=600, fg_color="#1E1E1E", corner_radius=10)
    frame_principal.pack_propagate(False)  # Impede ajuste automático
    frame_principal.pack(side="right", padx=(0, 5), pady=5)

    # Logo com tamanho fixo
    logo = ctk.CTkImage(light_image=Image.open(r"C:\Users\eduardo.klitzke\OneDrive - ADCOS PARTICIPAÇÕES LTDA\Documentos\Programação\Programação\KitPW\logo2.png"), size=(130, 130))
    logo_label = ctk.CTkLabel(sidebar, image=logo, text="")
    logo_label.pack(pady=(20, 10))

    # Nome do Programa
    ctk.CTkLabel(sidebar, text="Cadastro de Kits", font=("Arial", 18, "bold"), text_color="white").pack(pady=(5, 20))

    # Botões de navegação com tamanho fixo
    botoes = [
    ("CRIAR", lambda: mostrar_conteudo("CRIAR", frame_principal)),
    ("EDITAR", lambda: mostrar_editar(frame_principal)),  # Atualizado para chamar o menu de edição
    ("PROCEDIMENTO", lambda: abrir_procedimento())

]

    for texto, comando in botoes:
        ctk.CTkButton(
            sidebar,
            text=texto,
            command=comando,
            width=170,
            height=40,
            fg_color="#1ABC9C",
            hover_color="#16A085",
            corner_radius=5,
            border_width=0.5,
            border_color="#000"
        ).pack(pady=10)

    # Mostrar conteúdo inicial
    mostrar_conteudo("INICIO", frame_principal)

    app.mainloop()

criar_interface()