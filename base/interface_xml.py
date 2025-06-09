import xml.etree.ElementTree as ET
import pandas as pd
import os
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import sys

# Setup base path
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

# Caminho dos arquivos
tabela_produtos = os.path.join(base_path, 'tabelas', 'codInterno.csv')
tabela_cfop = os.path.join(base_path, 'tabelas', 'TESXCFOP.csv')

# Caminho do ícone
icone_caminho = os.path.join(base_path, 'lc1.ico')

# Lê os CSVs
df_codigos = pd.read_csv(tabela_produtos, delimiter=';', encoding='utf-8')
df_tes_cfop = pd.read_csv(tabela_cfop, delimiter=';', encoding='latin1')

resultados = []

def processar_xmls(pasta_xml):
    resultados.clear()
    arquivos_xml = [f for f in os.listdir(pasta_xml) if f.lower().endswith('.xml')]

    for arquivo in arquivos_xml:
        caminho_xml = os.path.join(pasta_xml, arquivo)
        try:
            tree = ET.parse(caminho_xml)
            root = tree.getroot()
            ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

            num_nf = root.find('.//nfe:ide/nfe:nNF', ns)
            numero_nota = num_nf.text.zfill(9) if num_nf is not None else 'N/A'

            dest_cnpj = root.find('.//nfe:dest/nfe:CNPJ', ns)
            cnpj_cliente = dest_cnpj.text.strip() if dest_cnpj is not None else 'N/A'

            dest_nome = root.find('.//nfe:dest/nfe:xNome', ns)
            nome_cliente = dest_nome.text.strip() if dest_nome is not None else 'N/A'

            icms_valor = root.find('.//nfe:ICMSTot/nfe:vICMS', ns)
            ipi_valor = root.find('.//nfe:ICMSTot/nfe:vIPI', ns)

            icms_valor = float(icms_valor.text) if icms_valor is not None else 0.0
            ipi_valor = float(ipi_valor.text) if ipi_valor is not None else 0.0

            possui_icms = 'S' if icms_valor > 0 else 'N'
            possui_ipi = 'S' if ipi_valor > 0 else 'N'

            itens = root.findall('.//nfe:det', ns)
            for item in itens:
                prod = item.find('nfe:prod', ns)
                if prod is not None:
                    cod_prod = prod.find('nfe:cProd', ns).text.strip()
                    nome_prod = prod.find('nfe:xProd', ns).text.strip()
                    cfop = prod.find('nfe:CFOP', ns).text.strip() if prod.find('nfe:CFOP', ns) is not None else ''
                    nat_op = root.find('.//nfe:ide/nfe:natOp', ns)
                    nat_op = nat_op.text.strip() if nat_op is not None else ''
                    qcom = prod.find('nfe:qCom', ns).text.strip() if prod.find('nfe:qCom', ns) is not None else ''
                    vuncom = prod.find('nfe:vUnCom', ns).text.strip() if prod.find('nfe:vUnCom', ns) is not None else ''
                    vprod = prod.find('nfe:vProd', ns).text.strip() if prod.find('nfe:vProd', ns) is not None else ''

                    filtro = df_codigos[df_codigos['codigo_interno'].astype(str).str[:10] == cod_prod[:10]]
                    cod_protheus = filtro['codigo'].astype(str).values[0] if not filtro.empty else 'NÃO ENCONTRADO'

                    linha_tes = df_tes_cfop[
                        (df_tes_cfop['F4_CF'].astype(str).str.strip() == cfop.strip()) &
                        (df_tes_cfop['F4_ICM'].astype(str).str.strip() == possui_icms) &
                        (df_tes_cfop['F4_IPI'].astype(str).str.strip() == possui_ipi)
                    ]
                    tes = linha_tes['F4_CODIGO'].astype(str).values[0] if not linha_tes.empty else 'NÃO ENCONTRADO'

                    # Formata quantidade
                    try:
                        qcom_float = float(qcom)
                        if qcom_float.is_integer():
                            quantidade_formatada = f"{int(qcom_float)}"
                        else:
                            quantidade_formatada = f"{qcom_float:.4f}".rstrip('0').rstrip('.').replace('.', ',')
                    except:
                        quantidade_formatada = qcom

                    # Formata valor unitário
                    try:
                        vuncom_float = float(vuncom)
                        valor_unitario_formatado = f"{vuncom_float:.4f}".replace('.', ',')
                    except:
                        valor_unitario_formatado = vuncom

                    # Formata valor total
                    try:
                        vprod_float = float(vprod)
                        valor_total_formatado = f"{vprod_float:.2f}".replace('.', ',')
                    except:
                        valor_total_formatado = vprod

                    # Formata ICMS e IPI
                    icms_formatado = f"{icms_valor:.2f}".replace('.', ',')
                    ipi_formatado = f"{ipi_valor:.2f}".replace('.', ',')

                    resultado = {
                        'Nota': numero_nota,
                        'CNPJ Cliente': cnpj_cliente,
                        'Nome Cliente': nome_cliente,
                        'Cod. Protheus': cod_protheus,
                        'Quant': quantidade_formatada,
                        'Valor Uni': valor_unitario_formatado,
                        'Valor Total': valor_total_formatado,
                        'TES': tes,
                        'CFOP': cfop,
                        'IPI': ipi_formatado,
                        'ICMS': icms_formatado,
                        'NatOp': nat_op,
                        'Cod. Interno': cod_prod,
                        'Descrição': nome_prod
                    }
                    resultados.append(resultado)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar '{arquivo}': {e}")

    atualizar_tabela()

def atualizar_tabela():
    for row in tree.get_children():
        tree.delete(row)
    for res in resultados:
        tree.insert("", tk.END, values=list(res.values()))

def escolher_pasta():
    pasta = filedialog.askdirectory()
    if pasta:
        processar_xmls(pasta)

def exportar_csv():
    if resultados:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"resultado_final_{timestamp}.csv"
        df_result = pd.DataFrame(resultados)
        df_result.to_csv(nome_arquivo, index=False)
        messagebox.showinfo("Exportado", f"Arquivo '{nome_arquivo}' salvo com sucesso!")
    else:
        messagebox.showwarning("Aviso", "Nenhum dado para exportar.")

# --------- INTERFACE -----------
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("green")

janela = ctk.CTk()
janela.title("Leitor de XML - NFe - LC1 Contadores Associados")
janela.geometry("1200x600")

if os.path.exists(icone_caminho):
    janela.iconbitmap(icone_caminho)

btn_pasta = ctk.CTkButton(janela, text="Selecionar Pasta de XMLs", command=escolher_pasta, corner_radius=15)
btn_pasta.pack(pady=10)

tree_frame = ctk.CTkFrame(janela, corner_radius=15)
tree_frame.pack(padx=20, pady=10, fill="both", expand=True)

cols = ['Nota', 'CNPJ Cliente', 'Nome Cliente', 'Cod. Protheus', 'Quant', 'Valor Uni', 'Valor Total',
        'TES', 'CFOP', 'IPI', 'ICMS', 'NatOp', 'Cod. Interno', 'Descrição']

style = ttk.Style()
style.theme_use("default")
style.configure("Treeview",
                background="white",
                foreground="black",
                rowheight=30,
                fieldbackground="white",
                bordercolor="#D3D3D3",
                borderwidth=0,
                font=('Segoe UI', 10))
style.map('Treeview', background=[('selected', '#4CAF50')])
style.configure("Treeview.Heading", font=('Segoe UI', 11, 'bold'), background="#4CAF50", foreground="white")

# Cria Treeview + Scrollbars
tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical")
tree_scroll_y.pack(side="right", fill="y")

tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
tree_scroll_x.pack(side="bottom", fill="x")

tree = ttk.Treeview(tree_frame, columns=cols, show='headings',
                    yscrollcommand=tree_scroll_y.set,
                    xscrollcommand=tree_scroll_x.set)

for col in cols:
    tree.heading(col, text=col)
    tree.column(col, width=120, anchor='w') 

tree.pack(fill="both", expand=True)

tree_scroll_y.config(command=tree.yview)
tree_scroll_x.config(command=tree.xview)

def copiar_celula(event):
    item = tree.identify_row(event.y)
    coluna = tree.identify_column(event.x)
    if item and coluna:
        coluna_index = int(coluna.replace('#', '')) - 1
        valores = tree.item(item)['values']
        valor = valores[coluna_index]

        if cols[coluna_index] == 'Nota':
            valor = str(valor).zfill(9)
        elif cols[coluna_index] == 'CNPJ Cliente' and valor != 'N/A':
            valor = str(valor).zfill(14)

        janela.clipboard_clear()
        janela.clipboard_append(str(valor))
        janela.update()

tree.bind("<Double-1>", copiar_celula)

btn_exportar = ctk.CTkButton(janela, text="Exportar CSV", command=exportar_csv, corner_radius=15)
btn_exportar.pack(pady=10)

janela.mainloop()
