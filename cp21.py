import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import scrolledtext, messagebox
from tkinter import ttk
from datetime import datetime
#usei pandas e numpy para os calculos e thinker para fazer a interface bonita caso não rode de primeira no seu pc tente baixar essas bibliotecas que da tudo certo
#nessa parte aqui é onde os dados são acrescentados em xlsx, como não sei acrescentar o csv
try:
    df = pd.read_excel('Dados1.xlsx')
    # Essa parte daqui é trivial só para fazer a conversão de date para data
    df['Date'] = pd.to_datetime(df['Date']).dt.date
    receita = df['Total Revenue']
    total_vendas = receita.sum()
except Exception as e:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Erro", f"Não foi possível carregar o arquivo Dados1.xlsx.\n{e}")
    exit()
#a ideia de ter essa mensagem aqui é caso quiser trocar esses dados, ai é só mudar a tabela e os titulos de tabela
#nessa parte foi feito o calculo das estatisticas criando essa definição
def calcular_estatisticas():
    media = receita.mean()
    desvio_absoluto_medio = (receita - media).abs().mean()
    variancia = receita.var(ddof=0)  
    desvio_padrao = receita.std(ddof=0)
    return (f"Média: {media:,.2f}\n"
            f"Desvio Absoluto Médio: {desvio_absoluto_medio:,.2f}\n"
            f"Variância: {variancia:,.2f}\n"
            f"Desvio Padrão: {desvio_padrao:,.2f}")

def cinco_maiores():
    maiores = df.nlargest(5, 'Total Revenue')[['Transaction ID', 'Date', 'Product Category', 'Product Name', 'Total Revenue']]
    #Nessa definição (me desculpe professor mas não lembro o nome disso) foram organizadas para começar pelas maiores usando a função df.nlargest
    linhas = []
    for _, row in maiores.iterrows():
        linhas.append(f"{row['Transaction ID']}  {row['Date']}  {row['Product Category']:15}  {row['Product Name']:30}  {row['Total Revenue']:>10,.2f}")
    return "\n".join(linhas) if linhas else "Nenhum dado."

def cinco_menores():
    menores = df.nsmallest(5, 'Total Revenue')[['Transaction ID', 'Date', 'Product Category', 'Product Name', 'Total Revenue']]
#aqui so usei a mesma logica mas com a função inversa que 
    linhas = []
    for _, row in menores.iterrows():
        linhas.append(f"{row['Transaction ID']}  {row['Date']}  {row['Product Category']:15}  {row['Product Name']:30}  {row['Total Revenue']:>10,.2f}")
    return "\n".join(linhas) if linhas else "Nenhum dado."

def desconto_5():
    desconto = total_vendas * 0.05
    return f"Valor total das vendas: {total_vendas:,.2f}\nDesconto de 5%: {desconto:,.2f}"

def taxa_desconto(valor):
    if 0.99 <= valor <= 99.99:
        return 0.02
    elif 100.00 <= valor <= 199.99:
        return 0.03
    elif 200.00 <= valor <= 499.99:
        return 0.06
    elif 500.00 <= valor <= 999.99:
        return 0.10
    elif 3000.00 <= valor <= 3999.99:
        return 0.12
    else:
        return 0.0

def desconto_escalonado():
    faixas = [
        ('0,99 a 99,99', 0.99, 99.99, 0.02),
        ('100,00 a 199,99', 100.00, 199.99, 0.03),
        ('200,00 a 499,99', 200.00, 499.99, 0.06),
        ('500,00 a 999,99', 500.00, 999.99, 0.10),
        ('3000,00 a 3999,99', 3000.00, 3999.99, 0.12),
        ('Outros / Sem desconto', None, None, 0.0)
    ]
    relatorio = []
    desconto_total = 0
    for descricao, lim_inf, lim_sup, taxa in faixas:
        if lim_inf is None:
            condicoes = []
            for d, li, ls, tx in faixas[:-1]:
                condicoes.append((df['Total Revenue'] >= li) & (df['Total Revenue'] <= ls))
            mascara = ~np.any(condicoes, axis=0)
        else:
            mascara = (df['Total Revenue'] >= lim_inf) & (df['Total Revenue'] <= lim_sup)

        itens = df[mascara]
        qtd = len(itens)
        valor_total = itens['Total Revenue'].sum()
        desconto = valor_total * taxa
        desconto_total += desconto
        relatorio.append(f"{descricao:<20} | {qtd:>5} | {valor_total:>15,.2f} | {taxa*100:>5.1f}% | {desconto:>15,.2f}")
    # Linha total
    relatorio.append("-" * 70)
    relatorio.append(f"{'TOTAL GERAL':<20} | {len(df):>5} | {total_vendas:>15,.2f} | {'':>5} | {desconto_total:>15,.2f}")
    return "\n".join(relatorio)

# Funções de callback para os botões
def mostrar_estatisticas():
    texto = calcular_estatisticas()
    resultado_text.delete(1.0, tk.END)
    resultado_text.insert(tk.END, texto)

def mostrar_maiores():
    texto = cinco_maiores()
    resultado_text.delete(1.0, tk.END)
    resultado_text.insert(tk.END, texto)

def mostrar_menores():
    texto = cinco_menores()
    resultado_text.delete(1.0, tk.END)
    resultado_text.insert(tk.END, texto)

def mostrar_desconto5():
    texto = desconto_5()
    resultado_text.delete(1.0, tk.END)
    resultado_text.insert(tk.END, texto)

def mostrar_escalonado():
    texto = desconto_escalonado()
    resultado_text.delete(1.0, tk.END)
    resultado_text.insert(tk.END, texto)

# Configuração da janela principal
root = tk.Tk()
root.title("Análise de Vendas - Dados1.xlsx")
root.geometry("900x600")
root.configure(bg="#f0f0f0")

# Frame superior para os botões
frame_botoes = tk.Frame(root, bg="#f0f0f0")
frame_botoes.pack(pady=10)

# Estilo dos botões
btn_style = {"font": ("Arial", 10, "bold"), "bg": "#4CAF50", "fg": "white", "width": 20, "height": 2}

btn_a = tk.Button(frame_botoes, text="a) Estatísticas", command=mostrar_estatisticas, **btn_style)
btn_a.grid(row=0, column=0, padx=5, pady=5)

btn_b = tk.Button(frame_botoes, text="b) 5 Maiores Vendas", command=mostrar_maiores, **btn_style)
btn_b.grid(row=0, column=1, padx=5, pady=5)

btn_c = tk.Button(frame_botoes, text="c) 5 Menores Vendas", command=mostrar_menores, **btn_style)
btn_c.grid(row=0, column=2, padx=5, pady=5)

btn_d = tk.Button(frame_botoes, text="d) Desconto 5%", command=mostrar_desconto5, **btn_style)
btn_d.grid(row=1, column=0, padx=5, pady=5)

btn_e = tk.Button(frame_botoes, text="e) Desconto Escalonado", command=mostrar_escalonado, **btn_style)
btn_e.grid(row=1, column=1, padx=5, pady=5)

btn_sair = tk.Button(frame_botoes, text="Sair", command=root.quit, bg="#f44336", fg="white", font=("Arial", 10, "bold"), width=20, height=2)
btn_sair.grid(row=1, column=2, padx=5, pady=5)

# Frame para a área de texto com rolagem
frame_texto = tk.Frame(root, bg="#f0f0f0")
frame_texto.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

resultado_text = scrolledtext.ScrolledText(frame_texto, wrap=tk.NONE, font=("Courier New", 10))
resultado_text.pack(fill=tk.BOTH, expand=True)

# Barra de rolagem horizontal (opcional, mas ScrolledText já tem vertical, a horizontal pode ser adicionada)
# Vamos configurar wrap para NONE e adicionar barra horizontal
resultado_text.config(wrap=tk.NONE)
hbar = tk.Scrollbar(frame_texto, orient=tk.HORIZONTAL, command=resultado_text.xview)
resultado_text.configure(xscrollcommand=hbar.set)
hbar.pack(side=tk.BOTTOM, fill=tk.X)

# Instrução inicial
resultado_text.insert(tk.END, "Clique em um botão para ver os resultados...")

root.mainloop()