# -*- coding: utf-8 -*-
r"""
triagem_importacao_rh.py
==========================
Diálogo modal exibido durante a importação de planilhas de RH (Transporte,
Folha RH, Diárias) quando `conciliar_importacao.classificar_colaboradores`
encontra colaboradores ausentes do cadastro de fornecedores ou com nome
divergente do que já está cadastrado.

Mesmo padrão de UX já usado em `solicitar_etapa_obra()` (ImportadorRH): um
único modal bloqueante para o lote inteiro, não uma caixa de diálogo por
pessoa — planilhas de diária costumam ter vários colaboradores pendentes de
uma vez, e confirmar um por um seria fricção sem necessidade.
"""

import tkinter as tk
from tkinter import ttk

from src.config.config import ARQUIVO_FORNECEDORES
from src.config.utils import custom_messagebox, validar_documento
from src.fornecedores.conciliar_importacao import upsert_fornecedores_lote


class TriagemImportacaoRH:
    """
    Uso:
        triagem = TriagemImportacaoRH(sistema.root, sistema, pendencias)
        resultado = triagem.executar()
        # resultado['cancelado']       -> True se o usuário cancelou a
        #                                  importação inteira; nada foi
        #                                  gravado no cadastro nesse caso.
        # resultado['cpfs_excluidos']  -> lista de CPFs (só dígitos) que
        #                                  NÃO devem ser incluídos nos
        #                                  lançamentos desta importação
        #                                  (o chamador deve filtrar
        #                                  `registros_lote` por este
        #                                  critério antes de estender
        #                                  `sistema.dados_para_incluir`).
    """

    def __init__(self, root, sistema, pendencias):
        self.root = root
        self.sistema = sistema
        self.pendencias = pendencias
        self.resultado = {'cpfs_excluidos': [], 'cancelado': False}
        self.linhas = {}  # iid da Treeview -> dict com os campos editáveis

    def executar(self):
        self._montar_janela()
        self.janela.wait_window()
        return self.resultado

    # ------------------------------------------------------------------
    def _montar_janela(self):
        self.janela = tk.Toplevel(self.root)
        self.janela.title("⚠️ Conciliação de Fornecedores — Importação de RH")
        self.janela.geometry("1000x540")
        self.janela.transient(self.root)
        self.janela.grab_set()
        self.janela.protocol("WM_DELETE_WINDOW", self._cancelar_tudo)

        self.janela.update_idletasks()
        largura, altura = 1000, 540
        x = (self.janela.winfo_screenwidth() // 2) - (largura // 2)
        y = (self.janela.winfo_screenheight() // 2) - (altura // 2)
        self.janela.geometry(f"{largura}x{altura}+{x}+{y}")

        main = ttk.Frame(self.janela, padding=15)
        main.pack(fill='both', expand=True)

        ttk.Label(main, text="⚠️ Colaboradores não conferem com o cadastro de fornecedores",
                  font=('Arial', 13, 'bold')).pack(anchor='w')
        ttk.Label(main,
                  text="Dê duplo clique numa linha para editar Nome / Dados Bancários / Categoria "
                       "antes de confirmar. Clique na coluna ✓ para excluir uma linha desta importação "
                       "(o lançamento dela não será importado agora).",
                  font=('Arial', 9), foreground='gray', justify='left', wraplength=950).pack(anchor='w', pady=(0, 10))

        tree_frame = ttk.Frame(main)
        tree_frame.pack(fill='both', expand=True)

        colunas = ('incluir', 'status', 'cpf', 'nome_planilha', 'nome_cadastro',
                   'dados_bancarios', 'categoria')
        self.tree = ttk.Treeview(tree_frame, columns=colunas, show='headings', height=14)
        titulos = {
            'incluir': '✓', 'status': 'Situação', 'cpf': 'CPF',
            'nome_planilha': 'Nome (planilha)', 'nome_cadastro': 'Nome (cadastro atual)',
            'dados_bancarios': 'Dados Bancários (a gravar)', 'categoria': 'Categoria',
        }
        larguras = {'incluir': 30, 'status': 110, 'cpf': 110, 'nome_planilha': 220,
                    'nome_cadastro': 220, 'dados_bancarios': 190, 'categoria': 70}
        for c in colunas:
            self.tree.heading(c, text=titulos[c])
            self.tree.column(c, width=larguras[c], anchor='w')

        scroll_y = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side='left', fill='both', expand=True)
        scroll_y.pack(side='right', fill='y')

        for p in self.pendencias:
            cpf_fmt = f"{p.cpf[:3]}.{p.cpf[3:6]}.{p.cpf[6:9]}-{p.cpf[9:]}"
            situacao = "🆕 Ausente" if p.status == 'ausente' else f"⚠️ Nome diverge ({p.similaridade:.0%})"
            iid = self.tree.insert('', 'end', values=(
                '☑', situacao, cpf_fmt, p.nome_planilha,
                p.nome_atual or '—', p.dados_bancarios_planilha, p.categoria_sugerida
            ))
            self.linhas[iid] = {
                'pendencia': p,
                'incluir': True,
                'nome_final': p.nome_planilha,
                'dados_bancarios_final': p.dados_bancarios_planilha,
                'categoria_final': p.categoria_sugerida,
            }

        self.tree.bind('<Button-1>', self._toggle_incluir)
        self.tree.bind('<Double-1>', self._editar_linha)

        botoes = ttk.Frame(main)
        botoes.pack(fill='x', pady=(12, 0))
        ttk.Button(botoes, text="✅ Cadastrar/Corrigir e Continuar Importação",
                   command=self._confirmar).pack(side='left', padx=5)
        ttk.Button(botoes, text="⏭️ Continuar Sem Gravar (pula todas as linhas acima)",
                   command=self._continuar_sem_gravar).pack(side='left', padx=5)
        ttk.Button(botoes, text="❌ Cancelar Importação Inteira",
                   command=self._cancelar_tudo).pack(side='right', padx=5)

    # ------------------------------------------------------------------
    def _toggle_incluir(self, event):
        regiao = self.tree.identify_region(event.x, event.y)
        coluna = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if regiao != 'cell' or coluna != '#1' or not iid:
            return
        linha = self.linhas[iid]
        linha['incluir'] = not linha['incluir']
        valores = list(self.tree.item(iid, 'values'))
        valores[0] = '☑' if linha['incluir'] else '☐'
        self.tree.item(iid, values=valores)

    def _editar_linha(self, event):
        iid = self.tree.identify_row(event.y)
        coluna = self.tree.identify_column(event.x)
        if not iid or coluna == '#1':
            return
        linha = self.linhas[iid]

        edit = tk.Toplevel(self.janela)
        edit.title("Editar antes de gravar")
        edit.transient(self.janela)
        edit.grab_set()
        frame = ttk.Frame(edit, padding=15)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text=f"CPF: {linha['pendencia'].cpf}",
                  font=('Arial', 9, 'bold')).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))

        ttk.Label(frame, text="Nome:").grid(row=1, column=0, sticky='w', pady=4)
        e_nome = ttk.Entry(frame, width=45)
        e_nome.insert(0, linha['nome_final'])
        e_nome.grid(row=1, column=1, pady=4)

        ttk.Label(frame, text="Dados Bancários:").grid(row=2, column=0, sticky='w', pady=4)
        e_dados = ttk.Entry(frame, width=45)
        e_dados.insert(0, linha['dados_bancarios_final'])
        e_dados.grid(row=2, column=1, pady=4)

        ttk.Label(frame, text="Categoria:").grid(row=3, column=0, sticky='w', pady=4)
        e_cat = ttk.Entry(frame, width=15)
        e_cat.insert(0, linha['categoria_final'])
        e_cat.grid(row=3, column=1, sticky='w', pady=4)

        def salvar():
            linha['nome_final'] = e_nome.get().strip().upper()
            linha['dados_bancarios_final'] = e_dados.get().strip()
            linha['categoria_final'] = e_cat.get().strip().upper()
            valores = list(self.tree.item(iid, 'values'))
            valores[3] = linha['nome_final']
            valores[5] = linha['dados_bancarios_final']
            valores[6] = linha['categoria_final']
            self.tree.item(iid, values=valores)
            edit.destroy()

        botoes = ttk.Frame(frame)
        botoes.grid(row=4, column=0, columnspan=2, pady=(15, 0))
        ttk.Button(botoes, text="Salvar", command=salvar).pack(side='left', padx=5)
        ttk.Button(botoes, text="Cancelar", command=edit.destroy).pack(side='left', padx=5)

    # ------------------------------------------------------------------
    def _montar_dados_para_gravar(self):
        a_gravar, excluidos = [], []
        for linha in self.linhas.values():
            p = linha['pendencia']
            if not linha['incluir']:
                excluidos.append(p.cpf)
                continue

            cpf_fmt = f"{p.cpf[:3]}.{p.cpf[3:6]}.{p.cpf[6:9]}-{p.cpf[9:]}"
            if not validar_documento(cpf_fmt, 'PF'):
                custom_messagebox("warning", "CPF inválido",
                    f"O CPF {cpf_fmt} não passa na validação de dígito verificador.\n\n"
                    f"Esta linha foi excluída automaticamente desta importação — "
                    f"corrija o CPF na planilha de origem e importe este colaborador separadamente.")
                excluidos.append(p.cpf)
                continue

            chave_pix = linha['dados_bancarios_final']
            if chave_pix.upper().startswith('PIX') and ':' in chave_pix:
                chave_pix = chave_pix.split(':', 1)[1].strip()

            a_gravar.append({
                'cnpj_cpf': p.cpf,
                'tipo_pessoa': 'PF',
                'razao_social': linha['nome_final'],
                'nome': linha['nome_final'],
                'categoria': linha['categoria_final'] or 'MO',
                'chave_pix': chave_pix,
                'dados_bancarios': linha['dados_bancarios_final'],
            })
        return a_gravar, excluidos

    def _confirmar(self):
        a_gravar, excluidos = self._montar_dados_para_gravar()
        if a_gravar:
            try:
                upsert_fornecedores_lote(a_gravar, ARQUIVO_FORNECEDORES)
            except PermissionError:
                custom_messagebox("error", "Erro de Permissão",
                    "Base_Fornecedores.xlsx está aberta no Excel. Feche o arquivo e tente novamente.")
                return
            except Exception as e:
                custom_messagebox("error", "Erro", f"Erro ao gravar fornecedores:\n{e}")
                return

            # A cache de fornecedores do sistema é indexada por mtime do
            # arquivo (carregar_cache_se_necessario) — como acabamos de
            # salvar, a próxima leitura já deve vir atualizada. Se a
            # implementação da cache não checar mtime automaticamente,
            # force a invalidação aqui, ex.:
            # self.sistema.cache_fornecedores.invalidar()

        self.resultado = {'cpfs_excluidos': excluidos, 'cancelado': False}
        msg = f"✅ {len(a_gravar)} fornecedor(es) cadastrado(s)/atualizado(s)."
        if excluidos:
            msg += f"\n⏭️ {len(excluidos)} linha(s) não serão importadas agora."
        custom_messagebox("info", "Conciliação Concluída", msg)
        self.janela.destroy()

    def _continuar_sem_gravar(self):
        # Ninguém é cadastrado/corrigido agora. Por segurança, TODAS as
        # pendências (não só as desmarcadas) ficam de fora da importação —
        # do contrário voltaríamos ao problema original: lançamento sem
        # fornecedor correspondente, pesquisável só por acidente.
        todos_cpfs = [linha['pendencia'].cpf for linha in self.linhas.values()]
        self.resultado = {'cpfs_excluidos': todos_cpfs, 'cancelado': False}
        self.janela.destroy()

    def _cancelar_tudo(self):
        self.resultado = {'cpfs_excluidos': [], 'cancelado': True}
        self.janela.destroy()
