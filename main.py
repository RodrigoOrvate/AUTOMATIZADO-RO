import tkinter as tk
from tkinter import filedialog
import pandas as pd
import openpyxl
import subprocess
import os
import sys
from procurar_objeto import procurar
from procurar_distvel import organizar

caminho_arquivo1 = ""
caminho_arquivo2 = ""
conjuntos_objetos = []  # Declaração global removida aqui
global_workbook = openpyxl.Workbook()  # Declaração global do workbook
global_excel_filename_obj = "dados_filtrados_obj.xlsx"  # Variável global para o nome do arquivo Excel obj
global_excel_filename_distvel = "dados_filtrados_distvel.xlsx" # Variável global para o nome do arquivo Excel distvel
erro_exibido = False  # Atualiza o status para indicar que a mensagem de erro não foi exibida
colunas_desejadas = ['DAY', 'ANIMAL', 'OBJECTS', 'Total Bouts', 'Total Duration(Second)', 'Latency(Second)', 'Ending time(Second) of First Bout']


def pesquisar_arquivo1():
    global caminho_arquivo1
    filename1 = filedialog.askopenfilename()
    caminho_arquivo1 = filename1
    caminho_entry1.config(state='normal')  # Definir o estado como normal
    caminho_entry1.insert(0, filename1)
    caminho_entry1.config(state='readonly')  # Definir o estado como somente leitura
    liberar_criar_conjuntos()

def pesquisar_arquivo2():
    global caminho_arquivo2
    filename2 = filedialog.askopenfilename()
    caminho_arquivo2 = filename2
    caminho_entry2.config(state='normal')  # Definir o estado como normal
    caminho_entry2.insert(0, filename2)  # Insere o caminho do arquivo no Entry
    caminho_entry2.config(state='readonly')  # Definir o estado como somente leitura
    liberar_botaodois()

def limpar_entry1():
    global caminho_entry1
    
    caminho_entry1.config(state='normal')  
    caminho_entry1.delete(0, tk.END)  
    caminho_entry1.config(state='readonly')
    
    # Limpar quantidade_entry
    quantidade_entry.delete(0, tk.END)

def limpar_entry2():
    global caminho_entry2
    
    caminho_entry2.config(state='normal')  
    caminho_entry2.delete(0, tk.END)  
    caminho_entry2.config(state='readonly')
    
    # Limpar quantidade_entry
    quantidade_entry.delete(0, tk.END)

def criar_conjunto_labels(conjuntos_frame, numero_conjunto):
    conjunto_frame = tk.Frame(conjuntos_frame, relief="solid", bd=1)
    conjunto_frame.grid(row=numero_conjunto-1, column=0, padx=5, pady=5, sticky="ew")
    
    titulo_label = tk.Label(conjunto_frame, text=f"Planilha {numero_conjunto}", font=("Arial", 12, "bold"))
    titulo_label.grid(row=0, column=0, columnspan=4, pady=(10, 5))

    objeto1_entry = tk.Entry(conjunto_frame, font=("Helvetica", 10, "bold"))
    objeto1_entry.insert(0, "Primeiro par_objeto")
    objeto1_entry.config(fg="grey")  # Define a cor do texto como cinza
    objeto1_entry.bind("<FocusIn>", lambda event: on_entry_click(objeto1_entry, "Primeiro par_objeto"))  # Remove o texto ao clicar
    objeto1_entry.bind("<FocusOut>", lambda event: on_focusout(objeto1_entry, "Primeiro par_objeto"))  # Restaura o texto se estiver vazio ao sair do foco
    objeto1_entry.grid(row=1, column=1, padx=(0, 5), pady=5, sticky="ew")

    objeto2_entry = tk.Entry(conjunto_frame, font=("Helvetica", 10, "bold"))
    objeto2_entry.insert(0, "Segundo par_objeto")
    objeto2_entry.config(fg="grey")
    objeto2_entry.bind("<FocusIn>", lambda event: on_entry_click(objeto2_entry, "Segundo par_objeto"))
    objeto2_entry.bind("<FocusOut>", lambda event: on_focusout(objeto2_entry, "Segundo par_objeto"))
    objeto2_entry.grid(row=1, column=3, padx=(0, 5), pady=5, sticky="ew")

    obj1_entry = tk.Entry(conjunto_frame, font=("Helvetica", 10, "bold"))
    obj1_entry.insert(0, "Primeiro OBJ")
    obj1_entry.config(fg="grey")
    obj1_entry.bind("<FocusIn>", lambda event: on_entry_click(obj1_entry, "Primeiro OBJ"))
    obj1_entry.bind("<FocusOut>", lambda event: on_focusout(obj1_entry, "Primeiro OBJ"))
    obj1_entry.grid(row=2, column=1, padx=(0, 5), pady=5, sticky="ew")

    obj2_entry = tk.Entry(conjunto_frame, font=("Helvetica", 10, "bold"))
    obj2_entry.insert(0, "Segundo OBJ")
    obj2_entry.config(fg="grey")
    obj2_entry.bind("<FocusIn>", lambda event: on_entry_click(obj2_entry, "Segundo OBJ"))
    obj2_entry.bind("<FocusOut>", lambda event: on_focusout(obj2_entry, "Segundo OBJ"))
    obj2_entry.grid(row=2, column=3, padx=(0, 5), pady=5, sticky="ew")

    return objeto1_entry, objeto2_entry, obj1_entry, obj2_entry

# Funções para manipular o texto do placeholder nos Entry
def on_entry_click(entry, placeholder):
    if entry.cget("fg") == "grey" and entry.get() == placeholder:
        entry.delete(0, "end")
        entry.config(fg="black")

def on_focusout(entry, placeholder):
    if entry.get() == "":
        entry.insert(0, placeholder)
        entry.config(fg="grey")

def reiniciar_programa():
    """Reinicia o programa forçadamente."""
    try:
        python = sys.executable
        os.chdir(os.path.dirname(sys.argv[0]))
        subprocess.Popen([python] + sys.argv)
        root.destroy()  # Fecha a janela atual
    except Exception as e:
        print(f"Erro ao reiniciar o programa: {e}")

def mostrar_erro(mensagem):
    global erro_exibido
    if not erro_exibido:  # Verifica se a mensagem de erro já foi exibida
        erro_window = tk.Toplevel()
        erro_window.title("Erro")
        
        label_erro = tk.Label(erro_window, text=mensagem, font=("Helvetica", 10))
        label_erro.pack(padx=20, pady=10)
        
        ok_button = tk.Button(erro_window, text="OK", font=("Helvetica", 10, "bold"), command=erro_window.destroy)
        ok_button.pack(pady=5)
        
        erro_window.geometry("470x100")
        erro_window.resizable(False, False)
        
        x_erro = (root.winfo_screenwidth() - erro_window.winfo_reqwidth()) / 2
        y_erro = (root.winfo_screenheight() - erro_window.winfo_reqheight()) / 2
        erro_window.geometry("+%d+%d" % (x_erro, y_erro))

        erro_exibido = True #Mostrar que o erro foi exibido

def mostrar_erro2(mensagem):
    global erro_exibido
    if not erro_exibido:  # Verifica se a mensagem de erro já foi exibida
        erro_window = tk.Toplevel()
        erro_window.title("Erro")
        
        label_erro = tk.Label(erro_window, text=mensagem, font=("Helvetica", 10))
        label_erro.pack(padx=20, pady=10)
        
        ok_button = tk.Button(erro_window, text="OK", font=("Helvetica", 10, "bold"), command=erro_window.destroy)
        ok_button.pack(pady=5)
        
        erro_window.geometry("470x100")
        erro_window.resizable(False, False)
        
        x_erro = (root.winfo_screenwidth() - erro_window.winfo_reqwidth()) / 2
        y_erro = (root.winfo_screenheight() - erro_window.winfo_reqheight()) / 2
        erro_window.geometry("+%d+%d" % (x_erro, y_erro))

        erro_exibido = True #Mostrar que o erro foi exibido

# Função para validar a entrada de quantidade de conjuntos
def validar_quantidade_entry(*args):
    quantidade_texto = quantidade_entry.get()
    if quantidade_texto == "":
        # Se o campo estiver vazio, não há necessidade de validação
        return
    if not quantidade_texto.isdigit():
        mostrar_erro("Por favor, digite apenas números inteiros para a quantidade de conjuntos.")
        quantidade_entry.delete(0, tk.END)  # Limpa o Entry

def criar_conjuntos():
    global erro_exibido
    erro_exibido = False
    global global_workbook
    try:
        num_conjuntos = int(quantidade_entry.get())
    except ValueError:
        mostrar_erro2("Por favor, digite um número válido para a quantidade de conjuntos.")
        return

    if num_conjuntos == 0:
        limite_label.config(text="Mínimo de planilhas é 1.", fg="red")
        return
    if num_conjuntos > 100:
        limite_label.config(text="Limite máximo de 100 planilhas alcançado.", fg="red")
        return
    else:
        limite_label.config(text="")

    global conjuntos_objetos

    conjuntos_objetos = []  # Limpa os objetos existentes
    # Limpa os conjuntos existentes
    for child in conjuntos_frame.winfo_children():
        child.destroy()

    global_workbook = openpyxl.Workbook()

    global caminho_arquivo1

    # Função para criar os novos conjuntos
    for i in range(1, num_conjuntos + 1):
        primeiro_objeto_entry, segundo_objeto_entry, primeiro_obj_entry, segundo_obj_entry = criar_conjunto_labels(conjuntos_frame, i)
        conjuntos_objetos.append((primeiro_objeto_entry, segundo_objeto_entry, primeiro_obj_entry, segundo_obj_entry))

    global_workbook.save(global_excel_filename_obj)  # Salva o workbook

    default_sheet = global_workbook['Sheet']
    global_workbook.remove(default_sheet)

    # Atualiza a altura do canvas conforme a quantidade de conjuntos
    conjuntos_canvas.update_idletasks()  # Atualiza o tamanho do canvas
    altura_conjuntos = conjuntos_frame.winfo_reqheight()  # Altura do frame dos conjuntos
    conjuntos_canvas.config(scrollregion=(0, 0, conjuntos_canvas.winfo_width(), altura_conjuntos))

    # Atualiza a altura da barra de rolagem de acordo com a altura dos conjuntos e do canvas
    altura_canvas = conjuntos_canvas.winfo_height()
    if num_conjuntos > 0:
        scrollbar.pack(side="right", fill="y")  # Exibe a barra de rolagem
        altura_scrollbar = min(1.0, altura_canvas / altura_conjuntos)
        scrollbar.config(command=conjuntos_canvas.yview)
        scrollbar.set(0, altura_scrollbar)
    else:
        scrollbar.pack_forget()  # Oculta a barra de rolagem quando não há conjuntos
    liberar_botaoum()

def liberar_criar_conjuntos():
    global caminho_arquivo1, conjuntos_objetos
    if caminho_arquivo1 and caminho_entry1.get():  # Verifica se caminho_arquivo1 está definido e se caminho_entry1 não está vazio
        criar_button.config(state="normal")  # Habilita o botão "Criar Conjuntos"
    else:
        criar_button.config(state="disabled")  # Desabilita o botão "Criar Conjuntos" se o caminho do arquivo 1 ou os conjuntos de objetos não estiverem preenchidos
        # Limpa os conjuntos
        for child in conjuntos_frame.winfo_children():
            child.destroy()

    # Chama novamente após 100ms para continuar verificando
    root.after(100, liberar_criar_conjuntos)

def liberar_botaoum():
    global conjuntos_objetos, caminho_arquivo1
    
    # Inicialmente assume que todos os conjuntos estão preenchidos corretamente
    all_entries_filled = True
    
    for objeto1_entry, objeto2_entry, obj1_entry, obj2_entry in conjuntos_objetos:
        # Verifica se os Entry ainda existem antes de tentar acessá-los
        if objeto1_entry.winfo_exists() and objeto2_entry.winfo_exists() and obj1_entry.winfo_exists() and obj2_entry.winfo_exists():
            # Verifica se algum dos campos está vazio ou contém o placeholder
            if not objeto1_entry.get() or objeto1_entry.get() == "Primeiro par_objeto" \
                or not objeto2_entry.get() or objeto2_entry.get() == "Segundo par_objeto" \
                or not obj1_entry.get() or obj1_entry.get() == "Primeiro OBJ" \
                or not obj2_entry.get() or obj2_entry.get() == "Segundo OBJ":
                # Se algum dos campos estiver vazio ou contiver o placeholder, considera que nem todos os campos estão preenchidos
                all_entries_filled = False
                break

    # Habilita ou desabilita o botão dependendo do estado de preenchimento dos campos
    if all_entries_filled and caminho_arquivo1:
        procurar_button1.config(state="normal")
    else:
        procurar_button1.config(state="disabled")
    
    # Chama novamente após 100ms para continuar verificando
    root.after(100, liberar_botaoum)

def liberar_botaodois():
    global caminho_arquivo2
    if caminho_arquivo2 and caminho_entry2.get():
        procurar_button2.config(state="normal")  # Habilita o botão se o caminho do arquivo 1 estiver preenchido
    else:
        procurar_button2.config(state="disabled")  # Desabilita o botão se o caminho do arquivo 1 não estiver preenchido
    root.after(100, liberar_botaodois)

def procurar_objetos():
    global caminho_arquivo1, conjuntos_objetos, global_workbook
    for primeiro_objeto, segundo_objeto, primeiro_obj, segundo_obj in conjuntos_objetos:
        procurar(primeiro_objeto.get().upper(), segundo_objeto.get().upper(), primeiro_obj.get().upper(),
                 segundo_obj.get().upper(), caminho_arquivo1, global_workbook, colunas_desejadas)
    global_workbook.save(global_excel_filename_obj)  # Salva o workbook
    # Ocultar a janela principal
    root.withdraw()
    # Exibir a mensagem de alerta
    mostrar_alerta()

def organizar_distvel():
    global caminho_arquivo2, global_workbook
    # Realiza a organização dos dados
    organizar(caminho_arquivo2, global_workbook)
    # Remove todas as folhas do arquivo, exceto "TR" e "TT"
    for sheet_name in global_workbook.sheetnames:
        if sheet_name not in ["TR", "TT"]:
            del global_workbook[sheet_name]
    # Salva o arquivo Excel
    global_workbook.save(global_excel_filename_distvel)
    # Mostra a mensagem de alerta
    mostrar_alerta()

entry_label_style = {"bg": "#F0F0F0",  # Cor de fundo cinza claro
               "fg": "#000000",  # Cor do texto preto
               "font": ("Helvetica", 10),  # Fonte Arial tamanho 10
               "bd": 0,  # Desativando a borda padrão
               "highlightthickness": 2,  # Espessura do destaque ao focar
               "highlightbackground": "#4CAF50",  # Cor do destaque ao focar
               "highlightcolor": "#4CAF50",  # Cor do destaque ao focar
               "relief": "flat",  # Estilo de relevo plano
}

from functools import partial
def mostrar_alerta():
    alerta = tk.Toplevel()
    alerta.title("Alerta")
    alerta.configure(bg="#f0f0f0")  # Define a cor de fundo do alerta
    alerta.geometry("400x100")  # Define o tamanho da janela de alerta
    
    mensagem_label = tk.Label(alerta, text="Dados filtrados com sucesso!\nVocê pode alterar os componentes se possível.", font=("Helvetica", 12), bg="#f0f0f0")  # Define a fonte e cor do texto
    mensagem_label.pack(padx=20, pady=10)
    
    ok_button = tk.Button(alerta, text="OK", font=("Helvetica", 12, "bold"), bg="#4CAF50", fg="white", command=partial(fechar_alerta, alerta))  # Define a fonte, cor do texto e cor do botão
    ok_button.pack(pady=5)
    
    # Desativa a maximização da janela de alerta
    alerta.resizable(False, False)
    
    # Define a posição da janela de alerta no centro da tela
    largura_tela = alerta.winfo_screenwidth()
    altura_tela = alerta.winfo_screenheight()
    x_alerta = (largura_tela - 400) / 2
    y_alerta = (altura_tela - 100) / 2
    alerta.geometry("+%d+%d" % (x_alerta, y_alerta))
    
    # Desativa o botão de fechar da janela de alerta
    alerta.protocol("WM_DELETE_WINDOW", lambda: None)

def fechar_alerta(alerta):
    alerta.destroy()
    abrir_arquivo_excel()

def abrir_arquivo_excel():
    # Verifica se ambos os arquivos existem
    if os.path.exists(global_excel_filename_obj) and os.path.exists(global_excel_filename_distvel):
        # Obtém o tempo de modificação de cada arquivo
        tempo_mod_obj = os.path.getmtime(global_excel_filename_obj)
        tempo_mod_distvel = os.path.getmtime(global_excel_filename_distvel)
        
        # Verifica qual arquivo foi modificado mais recentemente
        if tempo_mod_obj > tempo_mod_distvel:
            arquivo_recente = global_excel_filename_obj
        else:
            arquivo_recente = global_excel_filename_distvel
        
        # Abre o arquivo mais recentemente modificado
        os.system(f'start excel "{arquivo_recente}"')
    elif os.path.exists(global_excel_filename_obj):
        os.system(f'start excel "{global_excel_filename_obj}"')
    elif os.path.exists(global_excel_filename_distvel):
        os.system(f'start excel "{global_excel_filename_distvel}"')
    else:
        # Se nenhum dos arquivos existir, imprime uma mensagem indicando isso
        print("Nenhum dos arquivos globais foi encontrado.")
    
    # Fecha a aplicação Tkinter
    root.quit()

root = tk.Tk()
root.title("AUTOMATIZADO")

# Função para obter a largura e a altura da tela
largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()

# Quadro para pesquisa de arquivo
pesquisa_frame = tk.Frame(root)
pesquisa_frame.grid(row=0, column=0, columnspan=2, pady=10)

# Botão para pesquisar arquivo 1
pesquisar1_button = tk.Button(pesquisa_frame, text="Pesquisar Arquivo (OBJ)", font=("Helvetica", 10, "bold"), bg="#4CAF50", fg="white", command=pesquisar_arquivo1)
pesquisar1_button.grid(row=0, column=0, padx=10, pady=5)

# Botão para pesquisar arquivo 2
pesquisar2_button = tk.Button(pesquisa_frame, text="Pesquisar Arquivo (DIST/VEL)", font=("Helvetica", 10, "bold"), bg="#4CAF50", fg="white", command=pesquisar_arquivo2)
pesquisar2_button.grid(row=0, column=1, padx=10, pady=5)

# Espaço em branco para mostrar o caminho do arquivo 1
caminho_entry1 = tk.Entry(root, width=50, **entry_label_style)
caminho_entry1.grid(row=1, column=1, padx=(0, 5), pady=5)

# Botão para limpar o Entry 1
limpar1_button = tk.Button(root, text="x", font=("Arial", 10, "bold"), fg="red", bd=0, highlightthickness=0, command=limpar_entry1)
limpar1_button.grid(row=1, column=0, sticky="e", padx=(5, 2), pady=5)

# Espaço em branco para mostrar o caminho do arquivo 2
caminho_entry2 = tk.Entry(root, width=50, **entry_label_style)
caminho_entry2.grid(row=2, column=1, padx=(0, 5), pady=5)

# Botão para limpar o Entry 2
limpar2_button = tk.Button(root, text="x", font=("Arial", 10, "bold"), fg="red", bd=0, highlightthickness=0, command=limpar_entry2)
limpar2_button.grid(row=2, column=0, sticky="e", padx=(5, 2), pady=5)

# Quadro para criar conjuntos
criar_conjuntos_frame = tk.Frame(root)
criar_conjuntos_frame.grid(row=3, column=0, columnspan=2, pady=10)

quantidade_label = tk.Label(criar_conjuntos_frame, text="Quantidade de conjuntos:", font=("Helvetica", 10, "bold"))
quantidade_label.grid(row=0, column=0, padx=5)

quantidade_entry = tk.Entry(criar_conjuntos_frame, width=15, **entry_label_style)
quantidade_entry.grid(row=0, column=1, padx=5)
quantidade_entry.bind("<KeyRelease>", validar_quantidade_entry)

criar_button = tk.Button(criar_conjuntos_frame, text="Criar Conjuntos", font=("Helvetica", 10, "bold"), bg="#4CAF50", fg="white", command=criar_conjuntos, state="disabled")
criar_button.grid(row=0, column=2, padx=5)

# Definir pares_objetos e objs como variáveis globais
pares_objetos = set()
objs = set()

# Definir pares_objetos e objs como variáveis globais
pares_objetos = set()
objs = set()

def procurar_colunas(caminho_arquivo):
    global pares_objetos, objs
    
    try:
        # Carregar dados do Excel
        df = pd.read_excel(caminho_arquivo, header=6)
        
        # Extrair valores únicos das colunas 'OBJECTS' e 'Events'
        objetos = set(df['OBJECTS'].astype(str))  # Convertendo para strings
        events = set(df['Events'].astype(str))  # Convertendo para strings

        pares_objetos.clear()
        for objeto in objetos:
            if objeto.strip():
                pares = objeto.split(' & ')
                pares_objetos.update(pares)
        
        objs.clear()
        for event in events:
            # Extrair apenas os valores que começam com "OBJ"
            obj = event.split("OBJ")[1].split()[0] if "OBJ" in event else None
            if obj:
                objs.add(obj)
        
    except Exception as e:
        print(f"Erro ao abrir o arquivo Excel: {e}")

# Função para atualizar os rótulos com os pares de objetos e OBJs encontrados
def atualizar_rotulos():
    global pares_objetos, objs
    
    if caminho_arquivo1:
        procurar_colunas(caminho_arquivo1)
        texto = f"Pares de Objetos: {', '.join(pares_objetos)}\nOBJs: {', '.join(objs)}"
        pares_objetos_var.set(texto)  # Atualiza o valor do texto com os pares de objetos e OBJs

# Botão para atualizar os rótulos com os pares de objetos e OBJs
atualizar_rotulos_button = tk.Button(criar_conjuntos_frame, text="Ver objetos", font=("Helvetica", 10, "bold"), bg="#4CAF50", fg="white", command=atualizar_rotulos)
atualizar_rotulos_button.grid(row=1, column=2, padx=5, pady=5)

# Rótulo para exibir os pares de objetos e OBJs
pares_objetos_var = tk.StringVar()
pares_objetos_value_label = tk.Label(criar_conjuntos_frame, textvariable=pares_objetos_var, font=("Helvetica", 10), wraplength=500, justify="left")
pares_objetos_value_label.grid(row=3, column=0, columnspan=3, padx=10, pady=(0, 5), sticky="w")

# Rótulo para exibir os OBJs encontrados
objs_var = tk.StringVar()
objs_value_label = tk.Label(criar_conjuntos_frame, textvariable=objs_var, font=("Helvetica", 10), wraplength=500, justify="left")
objs_value_label.grid(row=4, column=1, columnspan=2, pady=(0, 5), sticky="w")

# Chamar a função para inicializar os valores dos rótulos
atualizar_rotulos()

# Botão para atualizar os rótulos com os pares de objetos e OBJs
atualizar_rotulos_button = tk.Button(criar_conjuntos_frame, text="Ver objetos", font=("Helvetica", 10, "bold"), bg="#4CAF50", fg="white", command=atualizar_rotulos)
atualizar_rotulos_button.grid(row=1, column=2, padx=5, pady=5)

limite_label = tk.Label(criar_conjuntos_frame, text="", fg="red")
limite_label.grid(row=1, columnspan=3, pady=(5, 0))

# Quadro para conjuntos de labels e barra de rolagem
conjuntos_scroll_frame = tk.Frame(root)
conjuntos_scroll_frame.grid(row=4, column=0, columnspan=2, pady=10)

conjuntos_canvas = tk.Canvas(conjuntos_scroll_frame, bg="#E0E0E0")
conjuntos_canvas.pack(side="left", fill="both", expand=True)

scrollbar = tk.Scrollbar(conjuntos_scroll_frame, orient="vertical", command=conjuntos_canvas.yview)
scrollbar.pack(side="right", fill="y")

conjuntos_frame = tk.Frame(conjuntos_canvas, bg="#E0E0E0")
conjuntos_frame.grid(row=0, column=0, sticky="nsew")

conjuntos_canvas.create_window((0, 0), window=conjuntos_frame, anchor="nw")

# Botões de procurar objeto
procurar_button_frame = tk.Frame(root)
procurar_button_frame.grid(row=5, column=0, columnspan=2, pady=10)

# Botão para procurar objeto 1
procurar_button1 = tk.Button(procurar_button_frame, text="Procurar Objetos", font=("Helvetica", 12, "bold"), bg="#4CAF50", fg="white", command=procurar_objetos, state="disabled")
procurar_button1.pack(side="left", padx=10)

# Botão para procurar objeto 2
procurar_button2 = tk.Button(procurar_button_frame, text="Organizar Dist/Vel", font=("Helvetica", 12, "bold"), bg="#4CAF50", fg="white", command=organizar_distvel, state="disabled")
procurar_button2.pack(side="left", padx=10)

reiniciar_button = tk.Button(procurar_button_frame, text="Reiniciar", font=("Helvetica", 10, "bold"), bg="#4CAF50", fg="white", command=reiniciar_programa)
reiniciar_button.pack(side="left", padx=10)

# Define o tamanho da janela principal como fixo
root.resizable(False, False)
# Define a posição da janela principal na tela após a execução do mainloop()
root.update_idletasks()
x_root = (largura_tela - root.winfo_width()) / 2
y_root = (altura_tela - root.winfo_height()) / 2
root.geometry("+%d+%d" % (x_root, y_root))

# Importe a biblioteca PIL para manipulação de imagens
from PIL import Image, ImageTk
# Defina o caminho para o arquivo de ícone (.ico)
caminho_icone = "memorylab.ico"
# Carregue o ícone como uma imagem usando PIL
icone = Image.open(caminho_icone)
# Converta a imagem para o formato TKinter PhotoImage
icone_tk = ImageTk.PhotoImage(icone)
# Defina o ícone da janela principal
root.iconphoto(True, icone_tk)

root.mainloop()