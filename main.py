import pandas as pd
import tkinterweb as tkweb
from tkinterweb import ttkweb, messageboxweb
from PIL import Image, ImageTk

# Carregar o arquivo Excel para um DataFrame
df = pd.read_excel('./Dados/CMED-2024-Teste-.xlsx')

# Função para realizar a pesquisa
def pesquisar(dados, tipo, termo):
    if tipo == "SUBSTÂNCIA":
        colunas = ['SUBSTÂNCIA','LABORATÓRIO', 'REGISTRO', 'PRODUTO', 'APRESENTAÇÃO', 'CLASSE TERAPÊUTICA',
                   'TIPO DE PRODUTO (STATUS DO PRODUTO)', 'RESTRIÇÃO HOSPITALAR', 'COMERCIALIZAÇÃO 2022',
                   'TARJA', 'SUBSTÂNCIA']
        resultado = dados[dados['SUBSTÂNCIA'].str.contains(termo, case=False)][colunas]
    elif tipo == "PRODUTO":
        colunas = ['SUBSTÂNCIA', 'LABORATÓRIO', 'REGISTRO', 'PRODUTO', 'APRESENTAÇÃO', 'CLASSE TERAPÊUTICA',
                   'TIPO DE PRODUTO (STATUS DO PRODUTO)', 'RESTRIÇÃO HOSPITALAR', 'COMERCIALIZAÇÃO 2022',
                   'TARJA']
        resultado = dados[dados['PRODUTO'].str.contains(termo, case=False)][colunas]
    else:
        resultado = pd.DataFrame(columns=dados.columns)  # Retorna um DataFrame vazio se o tipo não for válido
    return resultado

# Função para lidar com a ação do botão "Pesquisar"
def realizar_pesquisa():
    termo_de_pesquisa = termo_pesquisa.get()
    tipo_de_pesquisa = tipo_combobox.get()
    resultado_da_pesquisa = pesquisar(df, tipo_de_pesquisa, termo_de_pesquisa)

    # Limpar a tabela de resultados
    for item in resultados_treeview.get_children():
        resultados_treeview.delete(item)
    
    # Exibir a mensagem de busca inválida se não houver resultados
    if resultado_da_pesquisa.empty:
        messageboxweb.showinfo("Busca Inválida", "Nenhum resultado encontrado.")
    
    # Inserir os novos resultados na tabela
    for index, row in resultado_da_pesquisa.iterrows():
        resultados_treeview.insert("", "end", values=list(row))
    
    # Encontrar concorrentes e filtrar EUROFARMA LABORATÓRIOS S.A.
    concorrentes = resultado_da_pesquisa[resultado_da_pesquisa['LABORATÓRIO'] != "EUROFARMA LABORATÓRIOS S.A."]['LABORATÓRIO'].unique()

    # Criar um bloco de texto para exibir os concorrentes
    texto_concorrentes = "Concorrentes:\n"
    for concorrente in concorrentes:
        texto_concorrentes += f"- {concorrente}\n"

    # Exibir concorrentes abaixo da tabela
    label_concorrentes.config(text=texto_concorrentes)

# Função para limpar a pesquisa
def limpar_pesquisa():
    termo_pesquisa.delete(0, tkweb.END)
    tipo_combobox.set("SUBSTÂNCIA")
    for item in resultados_treeview.get_children():
        resultados_treeview.delete(item)
    label_concorrentes.config(text="")  # Limpar o bloco de texto dos concorrentes

# Criar a interface gráfica
root = tkweb.tkinter.Tk()
root.title("Consulta CMED")

# Frame para exibir a logo da empresa
logo_frame = ttkweb.Frame(root)
logo_frame.pack(side=tkweb.TOP, padx=10, pady=10, anchor='nw')

# Carregar a logo da empresa
logo_image = Image.open("images/logoeufa.png")
logo_image = logo_image.resize((100, 100))
logo_photo = ImageTk.PhotoImage(logo_image)

# Exibir a logo da empresa
logo_label = ttkweb.Label(logo_frame, image=logo_photo)
logo_label.image = logo_photo
logo_label.pack(side=tkweb.LEFT)

# Frame para os inputs
input_frame = ttkweb.Frame(root)
input_frame.pack(pady=10)

# Label e Entry para o termo de pesquisa
ttkweb.Label(input_frame, text="Termo de Pesquisa:").grid(row=0, column=0, padx=5, pady=5)
termo_pesquisa = ttkweb.Entry(input_frame)
termo_pesquisa.grid(row=0, column=1, padx=5, pady=5)

# Combobox para selecionar o tipo de pesquisa
ttkweb.Label(input_frame, text="Tipo de Pesquisa:").grid(row=1, column=0, padx=5, pady=5)
tipo_combobox = ttkweb.Combobox(input_frame, values=["SUBSTÂNCIA", "PRODUTO"])
tipo_combobox.grid(row=1, column=1, padx=5, pady=5)
tipo_combobox.set("SUBSTÂNCIA")

# Botão para realizar a pesquisa
pesquisar_button = ttkweb.Button(input_frame, text="Pesquisar", command=realizar_pesquisa)
pesquisar_button.grid(row=2, columnspan=2, padx=5, pady=5)

# Botão para limpar a pesquisa
limpar_button = ttkweb.Button(input_frame, text="Limpar Pesquisa", command=limpar_pesquisa)
limpar_button.grid(row=3, columnspan=2, padx=5, pady=5)

# Frame para exibir os resultados
resultados_frame = ttkweb.Frame(root)
resultados_frame.pack(pady=10)

# Tabela para exibir os resultados
colunas = ['SUBSTÂNCIA','LABORATÓRIO', 'REGISTRO', 'PRODUTO', 'APRESENTAÇÃO', 'CLASSE TERAPÊUTICA',
           'TIPO DE PRODUTO (STATUS DO PRODUTO)', 'RESTRIÇÃO HOSPITALAR', 'COMERCIALIZAÇÃO 2022',
           'TARJA',]
resultados_treeview = ttkweb.Treeview(resultados_frame, columns=colunas, show="headings")
for coluna in colunas:
    resultados_treeview.heading(coluna, text=coluna)
resultados_treeview.pack()

# Bloco de texto para exibir os concorrentes
label_concorrentes = ttkweb.Label(root)
label_concorrentes.pack(pady=10)

# Rodapé
rodape_label = ttkweb.Label(root, text="Criado por Lucas Lopes da Silva - Departamento de Terceiros Globais", anchor=tkweb.E)
rodape_label.pack(side=tkweb.RIGHT, padx=8, pady=8, anchor='se')

# Iniciar a interface gráfica
root.mainloop()
