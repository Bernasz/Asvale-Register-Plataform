import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from tkcalendar import DateEntry
import os
from datetime import datetime
from PIL import Image, ImageTk

class SistemaEntregas:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Cadastro de Entregas")
        
        # Maximizar a janela
        self.root.state('zoomed')
        
        # Frame principal
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill='both', expand=True)
        
        # Frame para o logo
        self.logo_frame = ttk.Frame(self.main_frame)
        self.logo_frame.pack(fill='x', padx=10, pady=5)
        
        # Carregar e exibir o logo
        try:
            # Carregar a imagem JPG usando PIL
            imagem_pil = Image.open('logo.png')
            # Redimensionar a imagem se necessário (ajuste os valores conforme necessário)
            imagem_pil = imagem_pil.resize((200, 100), Image.Resampling.LANCZOS)
            # Converter para PhotoImage
            self.logo_img = ImageTk.PhotoImage(imagem_pil)
            self.logo_label = ttk.Label(self.logo_frame, image=self.logo_img)
            self.logo_label.pack(pady=5)
        except Exception as e:
            print(f"Erro ao carregar o logo: {e}")
        
        # Configuração do arquivo Excel
        self.file_name = 'cadastro_entregas.xlsx'
        
        # Criar notebook (abas)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Criar abas
        self.aba_cadastro = ttk.Frame(self.notebook)
        self.aba_pesquisa = ttk.Frame(self.notebook)
        
        self.notebook.add(self.aba_cadastro, text='Cadastro')
        self.notebook.add(self.aba_pesquisa, text='Pesquisa')
        
        self.criar_aba_cadastro()
        self.criar_aba_pesquisa()
        
        # Carregar dados existentes com tipos corretos
        if os.path.exists(self.file_name):
            self.df = pd.read_excel(self.file_name, dtype={
                'Produtor/Empresa': str,
                'CPF/CNPJ': str,
                'I.E/Produtor': str,
                'Fone': str,
                'Data': str
            })
            
            # Configurações para evitar notação científica
            pd.set_option('display.float_format', '{:,.2f}'.format)
            pd.options.display.float_format = '{:,.0f}'.format
            pd.set_option('display.max_colwidth', None)
        else:
            self.criar_dataframe_inicial()

    def criar_aba_cadastro(self):
        # Frame principal da aba com configuração de peso
        main_frame = ttk.Frame(self.aba_cadastro)
        main_frame.pack(fill='both', expand=True)
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Frame para dados básicos
        frame_dados = ttk.LabelFrame(main_frame, text="Dados do Produtor/Empresa")
        frame_dados.pack(fill='x', padx=10, pady=5)
        frame_dados.grid_columnconfigure(1, weight=1)

        # Criar campos de entrada básicos com grid weights
        campos = [
            ('Produtor/Empresa:', 'produtor'),
            ('CPF/CNPJ:', 'cpf_cnpj'),
            ('I.E/Produtor:', 'ie_produtor'),
            ('Endereço:', 'endereco'),
            ('Município:', 'municipio'),
            ('Revenda(s):', 'revendas'),
            ('UF:', 'uf'),
            ('Fone:', 'fone')
        ]

        self.entradas = {}
        for i, (label, campo) in enumerate(campos):
            ttk.Label(frame_dados, text=label).grid(row=i, column=0, padx=5, pady=2, sticky='e')
            entrada = ttk.Entry(frame_dados)
            entrada.grid(row=i, column=1, padx=5, pady=2, sticky='ew')
            self.entradas[campo] = entrada

        # Data
        ttk.Label(frame_dados, text="Data:").grid(row=len(campos), column=0, padx=5, pady=2, sticky='e')
        self.entradas['data'] = DateEntry(frame_dados, width=12, background='darkblue',
                                        foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.entradas['data'].grid(row=len(campos), column=1, padx=5, pady=2, sticky='w')

        # Frame para embalagens
        frame_embalagens = ttk.LabelFrame(self.aba_cadastro, text="Embalagens")
        frame_embalagens.pack(fill='x', padx=10, pady=5)

        # Criar campos para cada tipo de embalagem
        self.criar_campos_embalagens(frame_embalagens)

        # Botão de salvar
        ttk.Button(self.aba_cadastro, text="Salvar Cadastro", 
                  command=self.salvar_cadastro).pack(pady=10)

    def criar_campos_embalagens(self, frame):
        frame.grid_columnconfigure((0,1,2), weight=1)  # Distribuir peso igualmente
        
        tipos_embalagens = ['Lavadas', 'Não Lavadas', 'Impróprias']
        self.entradas_embalagens = {}

        for i, tipo in enumerate(tipos_embalagens):
            frame_tipo = ttk.LabelFrame(frame, text=f"Embalagens {tipo}")
            frame_tipo.grid(row=0, column=i, padx=5, pady=5, sticky='nsew')
            frame_tipo.grid_columnconfigure(0, weight=1)

            # Checkbox com frame próprio
            checkbox_frame = ttk.Frame(frame_tipo)
            checkbox_frame.pack(fill='x', padx=2, pady=2)
            
            var = tk.BooleanVar()
            self.entradas_embalagens[f'Embalagens {tipo}'] = var
            ttk.Checkbutton(checkbox_frame, text=f"Tem embalagens {tipo}?", 
                          variable=var).pack(pady=2)

            # Frame para os campos de entrada
            campos_frame = ttk.Frame(frame_tipo)
            campos_frame.pack(fill='both', expand=True)
            campos_frame.grid_columnconfigure(0, weight=1)

            for j, tamanho in enumerate(['1L', '5L', '10L', '20L', 'ML', 'GR', 'KG']):
                frame_qtd = ttk.Frame(campos_frame)
                frame_qtd.pack(fill='x', padx=2, pady=2)
                frame_qtd.grid_columnconfigure(1, weight=1)
                
                ttk.Label(frame_qtd, text=f"{tamanho}:", width=8).pack(side='left')
                entrada = ttk.Entry(frame_qtd, width=8)
                entrada.pack(side='left', padx=2)
                entrada.insert(0, "0")
                
                chave = f'Embalagens {tipo} {tamanho}'
                self.entradas_embalagens[chave] = entrada

                if tamanho in ['ML', 'GR', 'KG']:
                    ttk.Label(frame_qtd, text="Peso/Vol:", width=8).pack(side='left', padx=2)
                    entrada_peso = ttk.Entry(frame_qtd, width=8)
                    entrada_peso.pack(side='left')
                    entrada_peso.insert(0, "0")
                    self.entradas_embalagens[f'Peso {tipo} {tamanho}'] = entrada_peso

    def criar_aba_pesquisa(self):
        # Frame para critérios de pesquisa
        frame_pesquisa = ttk.LabelFrame(self.aba_pesquisa, text="Pesquisar")
        frame_pesquisa.pack(fill='x', padx=10, pady=5)

        # Opções de pesquisa
        self.opcao_pesquisa = tk.StringVar(value="1")
        ttk.Radiobutton(frame_pesquisa, text="Produtor/Empresa", 
                       variable=self.opcao_pesquisa, value="1").pack(anchor='w')
        ttk.Radiobutton(frame_pesquisa, text="CPF/CNPJ", 
                       variable=self.opcao_pesquisa, value="2").pack(anchor='w')
        ttk.Radiobutton(frame_pesquisa, text="Data", 
                       variable=self.opcao_pesquisa, value="3").pack(anchor='w')

        # Campo de pesquisa
        self.entrada_pesquisa = ttk.Entry(frame_pesquisa)
        self.entrada_pesquisa.pack(fill='x', padx=5, pady=5)

        # Botão de pesquisa
        ttk.Button(frame_pesquisa, text="Pesquisar", 
                  command=self.realizar_pesquisa).pack(pady=5)

        # Treeview para resultados
        self.criar_treeview()

    def criar_treeview(self):
        # Frame para a tabela de resultados
        frame_resultados = ttk.LabelFrame(self.aba_pesquisa, text="Resultados")
        frame_resultados.pack(fill='both', expand=True, padx=10, pady=5)

        # Criar Treeview
        colunas = ['Produtor/Empresa', 'CPF/CNPJ', 'Data', 'Total Embalagens']
        self.tree = ttk.Treeview(frame_resultados, columns=colunas, show='headings')

        # Configurar cabeçalhos
        for col in colunas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)

        # Adicionar scrollbar
        scrollbar = ttk.Scrollbar(frame_resultados, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Posicionar elementos
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

    def criar_dataframe_inicial(self):
        # Criar DataFrame vazio com as colunas necessárias
        colunas = [
            'Produtor/Empresa', 'CPF/CNPJ', 'I.E/Produtor', 'Endereço', 'Município',
            'Revenda(s)', 'UF', 'Fone', 'Data'
            # Adicionar todas as outras colunas necessárias
        ]
        self.df = pd.DataFrame(columns=colunas)
        self.df.to_excel(self.file_name, index=False)

    def salvar_cadastro(self):
        # Implementar a lógica de salvamento
        try:
            dados = self.coletar_dados_formulario()
            self.df = pd.concat([self.df, pd.DataFrame([dados])], ignore_index=True)
            self.df.to_excel(self.file_name, index=False)
            messagebox.showinfo("Sucesso", "Cadastro realizado com sucesso!")
            self.limpar_formulario()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar: {str(e)}")

    def realizar_pesquisa(self):
        criterio = self.opcao_pesquisa.get()
        valor = self.entrada_pesquisa.get()

        # Limpar resultados anteriores
        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            if criterio == "1":  # Produtor/Empresa
                resultados = self.df[self.df['Produtor/Empresa'].str.contains(valor, case=False, na='')]
                titulo = f"Produtor/Empresa: {valor}"
            elif criterio == "2":  # CPF/CNPJ
                resultados = self.df[self.df['CPF/CNPJ'].str.contains(valor, na='')]
                titulo = f"CPF/CNPJ: {valor}"
            elif criterio == "3":  # Data
                if '-' in valor:  # Pesquisa por intervalo de datas
                    try:
                        data_inicio, data_fim = valor.split('-')
                        data_inicio = pd.to_datetime(data_inicio.strip(), format='%d/%m/%Y')
                        data_fim = pd.to_datetime(data_fim.strip(), format='%d/%m/%Y')

                        resultados = self.df[
                            (pd.to_datetime(self.df['Data'], format='%d/%m/%Y') >= data_inicio) & 
                            (pd.to_datetime(self.df['Data'], format='%d/%m/%Y') <= data_fim)
                        ]
                        titulo = f"Período: {data_inicio.strftime('%d/%m/%Y')} até {data_fim.strftime('%d/%m/%Y')}"
                    except:
                        messagebox.showerror("Erro", "Formato de data inválido. Use dd/mm/yyyy-dd/mm/yyyy")
                        return
                else:  # Pesquisa por data específica
                    try:
                        data_pesquisa = pd.to_datetime(valor.strip(), format='%d/%m/%Y')
                        resultados = self.df[pd.to_datetime(self.df['Data'], format='%d/%m/%Y') == data_pesquisa]
                        titulo = f"Data: {valor}"
                    except:
                        messagebox.showerror("Erro", "Formato de data inválido. Use dd/mm/yyyy")
                        return

            if not resultados.empty:
                # Variáveis para totais gerais
                total_geral_entregas = len(resultados)
                total_geral_lavadas = 0
                total_geral_nao_lavadas = 0
                total_geral_improprias = 0
                
                # Dicionário para armazenar totais por tipo e tamanho
                totais_detalhados = {
                    'Lavadas': {'1L': 0, '5L': 0, '10L': 0, '20L': 0, 'ML': 0, 'GR': 0, 'KG': 0},
                    'Não Lavadas': {'1L': 0, '5L': 0, '10L': 0, '20L': 0, 'ML': 0, 'GR': 0, 'KG': 0},
                    'Impróprias': {'1L': 0, '5L': 0, '10L': 0, '20L': 0, 'ML': 0, 'GR': 0, 'KG': 0}
                }

                for _, row in resultados.iterrows():
                    # Calcular totais por categoria para esta entrega
                    total_lavadas = 0
                    total_nao_lavadas = 0
                    total_improprias = 0

                    # Calcular totais detalhados
                    for tipo in ['Lavadas', 'Não Lavadas', 'Impróprias']:
                        for tamanho in ['1L', '5L', '10L', '20L', 'ML', 'GR', 'KG']:
                            coluna = f'Embalagens {tipo} {tamanho}'
                            valor = int(row.get(coluna, 0))
                            totais_detalhados[tipo][tamanho] += valor
                            
                            if tipo == 'Lavadas':
                                total_lavadas += valor
                            elif tipo == 'Não Lavadas':
                                total_nao_lavadas += valor
                            else:
                                total_improprias += valor

                    total_geral_lavadas += total_lavadas
                    total_geral_nao_lavadas += total_nao_lavadas
                    total_geral_improprias += total_improprias

                    total_esta_entrega = total_lavadas + total_nao_lavadas + total_improprias

                    # Inserir na TreeView
                    self.tree.insert('', 'end', values=(
                        row['Produtor/Empresa'],
                        row['CPF/CNPJ'],
                        row['Data'],
                        f"{total_esta_entrega} ({total_lavadas}L/{total_nao_lavadas}NL/{total_improprias}I)"
                    ))

                total_geral_embalagens = total_geral_lavadas + total_geral_nao_lavadas + total_geral_improprias

                # Criar mensagem detalhada
                mensagem = f"Resumo para {titulo}\n\n"
                mensagem += f"Total de entregas realizadas: {total_geral_entregas}\n"
                mensagem += f"Média de embalagens por entrega: {total_geral_embalagens/total_geral_entregas:.1f}\n\n"
                mensagem += f"Total geral de embalagens: {total_geral_embalagens}\n"
                mensagem += f"- Lavadas: {total_geral_lavadas}\n"
                mensagem += f"- Não Lavadas: {total_geral_nao_lavadas}\n"
                mensagem += f"- Impróprias: {total_geral_improprias}\n\n"
                
                mensagem += "Detalhamento por tipo e tamanho:\n"
                for tipo in ['Lavadas', 'Não Lavadas', 'Impróprias']:
                    mensagem += f"\n{tipo}:\n"
                    for tamanho in ['1L', '5L', '10L', '20L', 'ML', 'GR', 'KG']:
                        qtd = totais_detalhados[tipo][tamanho]
                        if qtd > 0:  # Só mostra se tiver quantidade
                            mensagem += f"  {tamanho}: {qtd}\n"
                            if tamanho in ['ML', 'GR', 'KG']:
                                coluna_peso = f'Peso {tipo} {tamanho}'
                                peso_total = sum(float(row.get(coluna_peso, 0)) for _, row in resultados.iterrows())
                                if peso_total > 0:
                                    mensagem += f"    Peso/Volume Total: {peso_total:.2f} {tamanho}\n"

                messagebox.showinfo("Resumo Detalhado", mensagem)
            else:
                messagebox.showinfo("Resultado", "Nenhum registro encontrado")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro na pesquisa: {str(e)}")

    def coletar_dados_formulario(self):
        dados = {}
        
        # Coletar dados básicos
        for campo, entrada in self.entradas.items():
            if isinstance(entrada, DateEntry):
                dados['Data'] = entrada.get_date().strftime('%d/%m/%Y')
            else:
                valor = entrada.get().strip()
                # Tratamento especial para campos numéricos longos
                if campo in ['cpf_cnpj', 'ie_produtor', 'fone']:
                    # Remove caracteres não numéricos
                    valor = ''.join(filter(str.isdigit, valor))
                    # Preenche com zeros à esquerda se necessário
                    if campo == 'cpf_cnpj':
                        valor = valor.zfill(14 if len(valor) > 11 else 11)
                    elif campo == 'fone':
                        valor = valor.zfill(11)
                
                mapeamento = {
                    'produtor': 'Produtor/Empresa',
                    'cpf_cnpj': 'CPF/CNPJ',
                    'ie_produtor': 'I.E/Produtor',
                    'endereco': 'Endereço',
                    'municipio': 'Município',
                    'revendas': 'Revenda(s)',
                    'uf': 'UF',
                    'fone': 'Fone'
                }
                nome_campo = mapeamento.get(campo, campo)
                dados[nome_campo] = valor

        # Coletar dados das embalagens
        for chave, entrada in self.entradas_embalagens.items():
            if isinstance(entrada, tk.BooleanVar):
                dados[chave] = "Sim" if entrada.get() else "Não"
            else:
                valor = entrada.get()
                dados[chave] = valor if valor else "0"

        return dados

    def limpar_formulario(self):
        # Limpar campos de entrada
        for entrada in self.entradas.values():
            if isinstance(entrada, DateEntry):
                entrada.set_date(datetime.now())
            else:
                entrada.delete(0, tk.END)

        # Limpar campos de embalagens
        for entrada in self.entradas_embalagens.values():
            if isinstance(entrada, tk.BooleanVar):
                entrada.set(False)
            else:
                entrada.delete(0, tk.END)
                entrada.insert(0, "0")

if __name__ == "__main__":
    root = tk.Tk()
    app = SistemaEntregas(root)
    root.mainloop()
