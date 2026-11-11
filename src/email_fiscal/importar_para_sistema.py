# -*- coding: utf-8 -*-
"""
importar_para_sistema.py - Tela "📧 E-mail" dentro da Entrada de Dados.

Fluxo:
  1. Sincroniza a triagem: o que já foi SALVO na planilha do cliente vira
     IMPORTADO (pelo marcador "IMPORTADO EMAIL #<id>" em OBSERVAÇÃO).
  2. Mostra as sugestões PENDENTES do cliente selecionado.
  3. Os itens marcados entram em sistema.dados_para_incluir - exatamente como
     a importação da Folha RH - e você confere no visualizador e clica Enviar.
     Todas as validações que o sistema já faz continuam valendo.
"""

VERSAO = "2026-11-12.8"   # igual em todos os arquivos do módulo
import os
from pathlib import Path
import threading
import webbrowser
import tkinter as tk
from tkinter import simpledialog
from datetime import datetime, date
from tkinter import ttk

from dateutil.relativedelta import relativedelta

from . import config_email as cfg
from .caixa_entrada import CaixaEntrada, ArquivoEmUso
from .roteador import BaseDados, so_digitos
from .leitor_boleto import cnpj_valido, cpf_valido
from .leitor_nfe import ler_emitente_arquivo
from .visualizar_nota import abrir_documento

try:
    from src.config.dialogs import custom_messagebox
except ImportError:
    custom_messagebox = None


def _msg(tipo, titulo, texto):
    if custom_messagebox:
        return custom_messagebox(tipo, titulo, texto)
    from tkinter import messagebox
    return {"yesno": messagebox.askyesno, "error": messagebox.showerror,
            "warning": messagebox.showwarning}.get(tipo, messagebox.showinfo)(titulo, texto)


def _preencher(widget, valor, forcar=False):
    """Escreve num Entry ou Combobox (inclusive readonly) sem apagar o que já estiver preenchido."""
    if widget is None or valor in (None, ""):
        return
    try:
        if widget.get().strip() and not forcar:
            return
    except Exception:
        pass
    try:
        widget.delete(0, "end")
        widget.insert(0, valor)
        if widget.get() != str(valor):
            raise ValueError
    except Exception:
        try:
            widget.set(valor)
        except Exception:
            pass


def _valor(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return round(float(v), 2)
    s = str(v).strip().replace("R$", "").strip()
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def _data(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return datetime.strptime(str(v).strip()[:10], "%d/%m/%Y").date()
    except (ValueError, TypeError):
        try:
            return datetime.strptime(str(v).strip()[:10], "%Y-%m-%d").date()
        except (ValueError, TypeError):
            return None


def _data_relatorio(sistema):
    """Mesma regra usada na importação da Folha RH."""
    try:
        valor = sistema.data_rel_entry.get()
        if valor:
            return valor
    except Exception:
        pass
    hoje = datetime.now()
    if 6 <= hoje.day <= 20:
        return hoje.replace(day=20).strftime("%d/%m/%Y")
    if hoje.day > 20:
        return (hoje.replace(day=1) + relativedelta(months=1)).replace(day=5).strftime("%d/%m/%Y")
    return hoje.replace(day=5).strftime("%d/%m/%Y")


def marcar_repetidos(caixa):
    """Mesmo boleto (linha digitável), mesma parcela de nota (chave+valor+vencimento)
    ou mesmo arquivo manual aparecendo mais de uma vez: a 1ª ocorrência fica,
    as PENDENTES repetidas viram IGNORADO com a indicação do ID original."""
    vistos, n = {}, 0
    ordem = sorted(caixa.lancamentos, key=lambda x: int(float(x.get("ID_SUGESTAO") or 0)))
    for l in ordem:
        st = str(l.get("STATUS") or "").upper()
        if st in ("IGNORADO", "CANCELADA"):
            continue
        chaves = []
        if l.get("LINHA_DIGITAVEL"):
            chaves.append("L" + str(l["LINHA_DIGITAVEL"]).strip())
        if l.get("CHAVE_DOC"):
            chaves.append(f"C{l['CHAVE_DOC']}|{_valor(l.get('VALOR'))}|{_data(l.get('DT_VENCTO'))}")
        if str(l.get("TIPO_DOC")) == "MANUAL" and l.get("ARQUIVOS"):
            chaves.append("A" + str(l["ARQUIVOS"]))
        original = next((vistos[c] for c in chaves if c in vistos), None)
        if original and st == "PENDENTE":
            l["STATUS"] = "IGNORADO"
            l["ALERTAS"] = f"REPETIDO do ID {original} | {l.get('ALERTAS') or ''}".strip(" |")
            n += 1
            continue
        for c in chaves:
            vistos.setdefault(c, l["ID_SUGESTAO"])
    return n


def dados_para_cadastro(l, base):
    """Reúne o que se sabe do fornecedor de uma sugestão: XML salvo (se houver),
    outra filial já cadastrada (mesma raiz de CNPJ) e a própria linha."""
    doc = so_digitos(l.get("CNPJ_CPF"))
    emit = {}
    for caminho in str(l.get("ARQUIVOS") or "").split(" ; "):
        caminho = caminho.strip()
        caminho = cfg.resolver_arquivo(caminho)
        if caminho.lower().endswith(".xml") and os.path.exists(caminho):
            emit = ler_emitente_arquivo(caminho)
            if emit.get("doc") == doc:
                break
            emit = {}
    parente = base.buscar_mesma_empresa(doc)
    razao = emit.get("razao_social") or str(l.get("NOME") or "").upper()
    nome = (parente or {}).get("nome") or emit.get("fantasia") or razao
    return {
        "cnpj_cpf": doc,
        "tipo_pessoa": "PJ" if len(doc) == 14 else "PF",
        "razao_social": razao,
        "nome": nome,
        "telefone": emit.get("telefone", ""),
        "email": emit.get("email", ""),
        "endereco": emit.get("endereco", ""),
        "categoria": (parente or {}).get("categoria") or str(l.get("CATEGORIA") or ""),
        "vinculo": "FORNECEDOR",
        "_parente": parente,
    }


def atualizar_fornecedores(caixa, base):
    """Fornecedor cadastrado depois da coleta: some o alerta de "não cadastrado"
    / "FILIAL NOVA" e NOME/CATEGORIA passam a ser os do cadastro."""
    n = 0
    for l in caixa.lancamentos:
        if str(l.get("STATUS") or "").upper() != "PENDENTE":
            continue
        alertas = str(l.get("ALERTAS") or "")
        if "NÃO cadastrado" not in alertas and "FILIAL NOVA" not in alertas:
            continue
        forn = base.buscar_fornecedor(l.get("CNPJ_CPF"))
        if not forn:
            continue
        l["NOME"] = forn["nome"]
        l["CATEGORIA"] = forn["categoria"] or l.get("CATEGORIA")
        l["ALERTAS"] = " | ".join(a.strip() for a in alertas.split("|")
                                  if a.strip() and "NÃO cadastrado" not in a and not a.strip().startswith("FILIAL NOVA"))
        n += 1
    return n


def sincronizar_status(caixa, base, so_cliente=None):
    """Marca IMPORTADO (marcador encontrado) ou JA_LANCADO (lançamento igual feito à mão).
    so_cliente: confere só as sugestões desse cliente (abre só a planilha dele)."""
    alterou = 0
    for l in caixa.lancamentos:
        if str(l.get("STATUS") or "").upper() != "PENDENTE" or not l.get("CLIENTE"):
            continue
        cliente = str(l["CLIENTE"]).strip()
        if so_cliente and cliente != so_cliente:
            continue
        achou, _ = base.ja_lancado(cliente, id_sugestao=l.get("ID_SUGESTAO"), doc=l.get("CNPJ_CPF"),
                                   valor=_valor(l.get("VALOR")))
        if achou:
            l["STATUS"], l["DATA_IMPORTACAO"] = "IMPORTADO", datetime.now()
            alterou += 1
            continue
        achou, motivo = base.ja_lancado(cliente, doc=l.get("CNPJ_CPF"), nf=l.get("NF"),
                                        valor=_valor(l.get("VALOR")), vencimento=_data(l.get("DT_VENCTO")),
                                        nfs_provaveis=str(l.get("NF_PROVAVEL") or "").split(";"),
                                        parcelado=" - PARC " in str(l.get("REFERENCIA") or ""))
        if achou:
            l["STATUS"] = "JA_LANCADO"
            l["ALERTAS"] = f"Lançado manualmente: {motivo} | {l.get('ALERTAS') or ''}".strip(" |")
            alterou += 1
            continue
        dup = base.possivel_duplicidade(cliente, doc=l.get("CNPJ_CPF"), nf=l.get("NF"),
                                        valor=_valor(l.get("VALOR")), vencimento=_data(l.get("DT_VENCTO")))
        alertas = str(l.get("ALERTAS") or "")
        if dup and "POSSÍVEL DUPLICIDADE" not in alertas and "Duplicidade descartada" not in alertas:
            l["ALERTAS"] = f"{dup} | {alertas}".strip(" |")
            alterou += 1
    return alterou


def data_rel_do_item(venc, data_rel_atual_str):
    """Relatório (dias de cfg.DIAS_RELATORIO) mais recente até o vencimento,
    nunca antes do relatório em uso. Devolve dd/mm/aaaa."""
    atual = _data(data_rel_atual_str)
    venc = _data(venc)
    if not cfg.DATA_REL_PELO_VENCIMENTO or not atual or not venc or venc <= atual:
        return data_rel_atual_str
    dias = sorted(cfg.DIAS_RELATORIO)
    candidatos = [d for d in dias if d <= venc.day]
    if candidatos:
        rel = venc.replace(day=candidatos[-1])
    else:  # antes do 1º dia de relatório do mês -> último relatório do mês anterior
        rel = (venc.replace(day=1) - relativedelta(days=1)).replace(day=dias[-1])
    return max(rel, atual).strftime("%d/%m/%Y")


def tp_desp_por_data(l, data_rel_str):
    """Regra do usuário (ver config_email.REGRA_TP_DESP_POR_DATA). None = regra não se aplica."""
    if not cfg.REGRA_TP_DESP_POR_DATA or str(l.get("TIPO_DOC")) not in cfg.TIPOS_COM_REGRA_POR_DATA:
        return None
    data_rel, email, venc = _data(data_rel_str), _data(l.get("EMAIL_DATA")), _data(l.get("DT_VENCTO"))
    if not (data_rel and email and venc) or email >= data_rel:
        return None
    return "3" if venc > data_rel else "5"


def _doc_formatado(doc):
    """CNPJ/CPF com pontuação, como o sistema grava (00.000.000/0000-00 / 000.000.000-00).
    Também evita que a lista do visualizador (Treeview) converta o documento em número
    e perca os zeros à esquerda."""
    d = so_digitos(doc)
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    if len(d) == 11:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    return d


def _referencia_importacao(l):
    ref = str(l.get("REFERENCIA") or "").strip()
    prov = str(l.get("NF_PROVAVEL") or "").split(";")[0].strip()
    if (cfg.NF_PROVAVEL_NA_REFERENCIA and not str(l.get("NF") or "").strip() and prov and prov != "-"
            and ref.upper().startswith("BOLETO - AGUARDANDO NF")):
        return f"BOLETO - NF {prov} (A CONFIRMAR)"
    return ref


def _para_registro(l, data_rel):
    """Converte a linha da triagem no dicionário que o SistemaEntradaDados usa."""
    valor = _valor(l.get("VALOR"))
    venc = _data(l.get("DT_VENCTO"))
    return {
        "data": data_rel,
        "cnpj_cpf": _doc_formatado(l.get("CNPJ_CPF")),
        "nome": str(l.get("NOME") or "").strip(),
        "categoria": str(l.get("CATEGORIA") or "").strip(),
        "tp_desp": l.get("_TP") or str(l.get("TP_DESP") or "").strip(),
        "referencia": _referencia_importacao(l),
        "etapa_obra": "",
        "insumo": "",
        "nf": str(l.get("NF") or "").strip(),
        "vr_unit": f"{valor:.2f}",
        "dias": 1,
        "valor": f"{valor:.2f}",
        "dt_vencto": venc.strftime("%d/%m/%Y"),
        "dados_bancarios": str(l.get("DADOS_BANCARIOS") or "").strip(),
        "observacao": f"{cfg.PREFIXO_OBSERVACAO} #{l['ID_SUGESTAO']} - {datetime.now():%d/%m/%Y %H:%M:%S}",
        "forma_pagamento": "",
    }


class JanelaImportacaoEmail:
    COLS = ("ID", "IDENT.", "FORNECEDOR", "REFERÊNCIA", "NF", "VALOR", "VENCTO", "ALERTAS")
    FILTROS = ("Todos os pendentes", "Só possíveis duplicidades", "Só com alertas",
               "Só sem alertas (prontos)", "Ignorados (para restaurar)")

    def __init__(self, sistema):
        self.sistema = sistema
        self.cliente = getattr(sistema, "cliente_atual", None)
        self.jan = tk.Toplevel(sistema.root)
        self.jan.title(f"Documentos recebidos por e-mail - {self.cliente}")
        self.jan.geometry("1250x600")
        self.jan.minsize(980, 420)
        self.jan.transient(sistema.root)

        topo = ttk.Frame(self.jan)
        topo.pack(fill="x", padx=10, pady=6)
        ttk.Label(topo, text=f"Cliente: {self.cliente}", font=("Helvetica", 11, "bold")).pack(side="left")
        self.var_sem_cliente = tk.BooleanVar(value=False)
        ttk.Checkbutton(topo, text="Mostrar também itens SEM cliente definido",
                        variable=self.var_sem_cliente, command=self.recarregar).pack(side="left", padx=20)
        ttk.Label(topo, text="Mostrar:").pack(side="left", padx=(10, 4))
        self.var_filtro = tk.StringVar(value=self.FILTROS[0])
        cb = ttk.Combobox(topo, textvariable=self.var_filtro, values=self.FILTROS, state="readonly", width=30)
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda e: self._preencher_lista())

        meio = ttk.Frame(self.jan)
        meio.pack(fill="both", expand=True, padx=10)
        self.tree = ttk.Treeview(meio, columns=self.COLS, show="headings", selectmode="extended")
        larg = (50, 110, 230, 250, 70, 90, 85, 340)
        for c, w in zip(self.COLS, larg):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="e" if c == "VALOR" else "w")
        self.tree.tag_configure("alerta", background="#FFF3CD")
        self.tree.tag_configure("duplicidade", background="#F8D7DA")
        self.tree.tag_configure("ignorado", foreground="#888888")
        sb = ttk.Scrollbar(meio, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", lambda e: self.abrir_documento())
        sb.pack(side="right", fill="y")

        self.log = tk.Text(self.jan, height=6, font=("Consolas", 9))
        self.log.pack(fill="x", padx=10, pady=4)

        # Botões em três fileiras, empacotadas ANTES da lista: se a janela ficar
        # pequena, quem encolhe é a lista - os botões (e o Fechar) nunca somem.
        coleta = ttk.Frame(self.jan)
        coleta.pack(side="bottom", fill="x", padx=10, pady=(2, 8), before=meio)
        acoes = ttk.Frame(self.jan)
        acoes.pack(side="bottom", fill="x", padx=10, pady=2, before=meio)
        triagem = ttk.Frame(self.jan)
        triagem.pack(side="bottom", fill="x", padx=10, pady=(4, 2), before=meio)
        self.log.pack_forget()
        self.log.pack(side="bottom", fill="x", padx=10, pady=4, before=meio)

        ttk.Label(triagem, text="Triagem:", width=11).pack(side="left")
        for texto, cmd in (("⚖️ Comparar com o lançamento existente", self.comparar),
                           ("📝 Editar sugestão", self.editar),
                           ("✏️ Informar CNPJ", self.informar_cnpj),
                           ("🏢 Cadastrar fornecedor", self.cadastrar_fornecedor),
                           ("🚫 Ignorar selecionados", self.ignorar),
                           ("↩️ Voltar para pendente", self.restaurar)):
            ttk.Button(triagem, text=texto, command=cmd).pack(side="left", padx=3)

        ttk.Label(acoes, text="Documentos:", width=11).pack(side="left")
        for texto, cmd in (("✅ Importar selecionados", self.importar),
                           ("📄 Abrir documento", self.abrir_documento),
                           ("✉️ Abrir e-mail", self.abrir_email)):
            ttk.Button(acoes, text=texto, command=cmd).pack(side="left", padx=3)

        # Fechar é empacotado primeiro, à direita: tem espaço garantido
        ttk.Button(coleta, text="Fechar", command=self.jan.destroy).pack(side="right", padx=3)
        ttk.Label(coleta, text="Coleta:", width=11).pack(side="left")
        for texto, cmd in (("📥 Buscar e-mails agora", self.buscar_emails),
                           ("📂 Ler pasta(s)", self.ler_pasta),
                           ("🔄 Reprocessar triagem", self.reprocessar),
                           ("📊 Abrir planilha de triagem", self.abrir_planilha),
                           ("🔃 Recarregar", self.recarregar)):
            ttk.Button(coleta, text=texto, command=cmd).pack(side="left", padx=3)

        self._conferir_versoes()
        self.recarregar()

    def _conferir_versoes(self):
        """Avisa se algum arquivo do módulo ficou de uma versão anterior."""
        import importlib
        desatualizados = []
        for nome in ("config_email", "gmail_cliente", "leitor_nfe", "leitor_boleto",
                     "roteador", "caixa_entrada", "coletor", "localizar_documentos", "visualizar_nota"):
            try:
                mod = importlib.import_module(f"{__package__}.{nome}")
                if getattr(mod, "VERSAO", None) != VERSAO:
                    desatualizados.append(f"{nome}.py ({getattr(mod, 'VERSAO', 'sem versão')})")
            except Exception as e:
                desatualizados.append(f"{nome}.py (erro ao carregar: {e})")
        if desatualizados:
            self.escrever(f"ATENÇÃO: arquivos de versão diferente de {VERSAO}: " + ", ".join(desatualizados)
                          + " - substitua-os pelos da última entrega.")
        else:
            self.escrever(f"Módulo email_fiscal versão {VERSAO} (todos os arquivos conferidos).")

    # ---------------------------------------------------------------------
    def escrever(self, texto):
        self.log.insert("end", texto + "\n")
        self.log.see("end")

    def recarregar(self):
        try:
            self.base, self.caixa = BaseDados(), CaixaEntrada()
        except Exception as e:
            _msg("error", "Erro", f"Não foi possível ler as planilhas:\n{e}")
            return
        # Abertura LEVE: nada de abrir PDFs aqui (isso travaria a tela).
        # Releitura de PDFs: botão "🔄 Reprocessar triagem" e fim de cada coleta.
        from .coletor import reidentificar_clientes
        rid, tocados = reidentificar_clientes(self.caixa, self.base, ler_pdf=False)
        rel = 0
        if rid:
            self.escrever(f"{rid} sugestão(ões) com cliente confirmado pelo CPF/CNPJ do cadastro atual.")
        forn = atualizar_fornecedores(self.caixa, self.base)
        if forn:
            self.escrever(f"{forn} sugestão(ões) atualizada(s) com fornecedor cadastrado.")
        from .coletor import ignorar_sem_cobranca
        semcob = ignorar_sem_cobranca(self.caixa, so_sem_vencimento=True)
        if semcob:
            self.escrever(f"{semcob} NF-e sem cobrança (remessa/retorno/sem pagamento) marcada(s) como IGNORADO.")
        tocados += semcob
        comp = self.caixa.anexar_complementares()
        if comp:
            self.escrever(f"{comp} documento(s) complementar(es) anexado(s) à nota correspondente.")
        refs = self.caixa.ignorar_chaves_referenciadas() + comp
        if refs:
            self.escrever(f"{refs} NFC-e apenas citada(s) em outra nota marcada(s) como IGNORADO.")
        tocados += forn + refs
        rep = marcar_repetidos(self.caixa)
        if rep:
            self.escrever(f"{rep} sugestão(ões) repetida(s) marcada(s) como IGNORADO.")
        uni = self.caixa.unir_boletos_e_notas()
        if uni:
            self.escrever(f"{uni} boleto(s) unido(s) à nota fiscal correspondente (e-mails diferentes).")
        n = sincronizar_status(self.caixa, self.base, so_cliente=self.cliente) + rep + uni + tocados
        if n:
            try:
                self.caixa.salvar(self.base.nomes_clientes())
                self.escrever(f"Triagem atualizada ({n} alteração(ões)).")
            except ArquivoEmUso as e:
                self.escrever(str(e))

        self._preencher_lista()

    def _preencher_lista(self):
        filtro = self.var_filtro.get()
        status = ("IGNORADO",) if filtro.startswith("Ignorados") else ("PENDENTE",)
        em_andamento = {str(r.get("observacao", "")).split(" - ")[0]
                        for r in (getattr(self.sistema, "dados_para_incluir", None) or [])}
        # Itens que estavam na lista de envio e saíram dela: foram GRAVADOS (Enviar) ou
        # descartados. Confere na planilha do cliente antes de mostrá-los de novo.
        saidos = getattr(self, "_em_envio_anterior", set()) - em_andamento
        self._em_envio_anterior = em_andamento
        if saidos:
            self._conferir_gravados()
        self.itens = {}
        self.tree.delete(*self.tree.get_children())
        pre_selecionar = []
        for l in self.caixa.pendentes(self.cliente, self.var_sem_cliente.get(), status=status):
            if f"{cfg.PREFIXO_OBSERVACAO} #{l['ID_SUGESTAO']}" in em_andamento:
                continue  # já está na lista de envio desta sessão
            alertas = str(l.get("ALERTAS") or "")
            duplicidade = "POSSÍVEL DUPLICIDADE" in alertas
            if filtro.startswith("Só possíveis") and not duplicidade:
                continue
            if filtro.startswith("Só com alertas") and not alertas:
                continue
            if filtro.startswith("Só sem alertas") and alertas:
                continue
            v, d = _valor(l.get("VALOR")), _data(l.get("DT_VENCTO"))
            iid = self.tree.insert("", "end", values=(
                l["ID_SUGESTAO"], l.get("IDENTIFICACAO") or "", l.get("NOME") or "",
                _referencia_importacao(l),   # exatamente o texto que irá para a planilha
                l.get("NF") or (str(l.get("NF_PROVAVEL")).split(";")[0] + "?"
                                if str(l.get("NF_PROVAVEL") or "-") != "-" else ""),
                f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if v is not None else "",
                d.strftime("%d/%m/%Y") if d else "", alertas),
                tags=(("ignorado",) if status == ("IGNORADO",) else
                      ("duplicidade",) if duplicidade else ("alerta",) if alertas else ()))
            self.itens[iid] = l
            if status == ("PENDENTE",) and not duplicidade:
                pre_selecionar.append(iid)
        # possíveis duplicidades NÃO vêm marcadas: só entram se você selecionar
        self.tree.selection_set(pre_selecionar)
        dups = 0 if status != ("PENDENTE",) else \
            sum(1 for l in self.itens.values() if "POSSÍVEL DUPLICIDADE" in str(l.get("ALERTAS") or ""))
        self.escrever(f"{len(self.itens)} sugestão(ões) listada(s) para {self.cliente} ({filtro.lower()})"
                      + (f"; {dups} possível(is) duplicidade(s) em vermelho, fora da seleção." if dups else "."))

    # ---------------------------------------------------------------------
    def _conferir_gravados(self):
        """Relê a planilha do cliente e marca IMPORTADO o que já foi gravado com o
        marcador "IMPORTADO EMAIL #n" (evita importar de novo a mesma sugestão)."""
        try:
            self.base.limpar_cache_cliente(self.cliente)
            n = 0
            for l in self.caixa.pendentes(self.cliente, False):
                achou, _ = self.base.ja_lancado(self.cliente, id_sugestao=l.get("ID_SUGESTAO"),
                                                doc=l.get("CNPJ_CPF"), valor=_valor(l.get("VALOR")))
                if achou:
                    l["STATUS"], l["DATA_IMPORTACAO"] = "IMPORTADO", datetime.now()
                    n += 1
            if n:
                self._gravar_triagem()
                self.escrever(f"{n} sugestão(ões) já gravada(s) na planilha do cliente: marcadas como IMPORTADO.")
            return n
        except Exception as e:
            self.escrever(f"Não foi possível conferir a planilha do cliente: {e}")
            return 0

    def _fechar_se_vazia(self):
        """Depois de uma ação: se não sobrou nada pendente deste cliente, fecha a janela."""
        if not cfg.FECHAR_JANELA_QUANDO_VAZIA:
            return False
        # sugestões já levadas para a lista de envio não contam como pendentes
        em_envio = {str(r.get("observacao", "")).split(" - ")[0]
                    for r in (getattr(self.sistema, "dados_para_incluir", None) or [])}
        if any(f"{cfg.PREFIXO_OBSERVACAO} #{l['ID_SUGESTAO']}" not in em_envio
               for l in self.caixa.pendentes(self.cliente, False)):
            return False
        try:
            self.jan.destroy()
        except Exception:
            pass
        return True

    def _gravar_triagem(self):
        try:
            self.caixa.salvar(self.base.nomes_clientes())
            return True
        except ArquivoEmUso as e:
            _msg("error", "Planilha de triagem aberta", str(e))
            return False

    def ignorar(self, itens=None):
        itens = itens or [self.itens[i] for i in self.tree.selection()]
        itens = [l for l in itens if str(l.get("STATUS")).upper() == "PENDENTE"]
        if not itens:
            _msg("warning", "Aviso", "Selecione sugestões pendentes para ignorar.")
            return
        if not _msg("yesno", "Ignorar sugestões",
                    f"Marcar {len(itens)} sugestão(ões) como IGNORADO?\n\n"
                    f"Elas saem da lista e não serão importadas. Podem voltar pelo filtro "
                    f"'Ignorados (para restaurar)'."):
            return
        hoje = datetime.now().strftime("%d/%m/%Y")
        for l in itens:
            l["STATUS"] = "IGNORADO"
            l["ALERTAS"] = f"IGNORADO pelo usuário em {hoje} | {l.get('ALERTAS') or ''}".strip(" |")
        if self._gravar_triagem():
            self.escrever(f"{len(itens)} sugestão(ões) marcada(s) como IGNORADO.")
        self._preencher_lista()
        self._fechar_se_vazia()

    def editar(self):
        """Edita uma sugestão numa janela (sem abrir o Excel), com o documento ao lado."""
        l = self._um_selecionado()
        if not l:
            return
        jan = tk.Toplevel(self.jan)
        jan.title(f"Editar sugestão #{l['ID_SUGESTAO']}")
        jan.geometry("640x560")
        jan.transient(self.jan)
        frm = ttk.Frame(jan, padding=12)
        frm.pack(fill="both", expand=True)

        venc = _data(l.get("DT_VENCTO"))
        v = _valor(l.get("VALOR"))
        campos = [
            ("CNPJ_CPF", "CNPJ/CPF do fornecedor", l.get("CNPJ_CPF") or ""),
            ("NOME", "Nome do fornecedor", l.get("NOME") or ""),
            ("REFERENCIA", "Referência", l.get("REFERENCIA") or ""),
            ("NF", "NF", l.get("NF") or ""),
            ("VALOR", "Valor (ex.: 1.234,56)", f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if v is not None else ""),
            ("DT_VENCTO", "Vencimento (dd/mm/aaaa)", venc.strftime("%d/%m/%Y") if venc else ""),
            ("CATEGORIA", "Categoria", l.get("CATEGORIA") or ""),
            ("TP_DESP", "TP_DESP (vazio = regra de datas)", l.get("TP_DESP") or ""),
            ("DADOS_BANCARIOS", "Dados bancários", l.get("DADOS_BANCARIOS") or ""),
        ]
        entradas = {}
        for i, (chave, rotulo, valor) in enumerate(campos):
            ttk.Label(frm, text=rotulo).grid(row=i, column=0, sticky="w", pady=3)
            e = ttk.Entry(frm, width=55)
            e.insert(0, str(valor))
            e.grid(row=i, column=1, sticky="we", pady=3)
            entradas[chave] = e
        i = len(campos)
        ttk.Label(frm, text="Cliente").grid(row=i, column=0, sticky="w", pady=3)
        cb = ttk.Combobox(frm, values=self.base.nomes_clientes(), width=52, state="readonly")
        cb.set(str(l.get("CLIENTE") or ""))
        cb.grid(row=i, column=1, sticky="we", pady=3)
        aviso = ttk.Label(frm, text="", foreground="#b00020", wraplength=600)
        aviso.grid(row=i + 1, column=0, columnspan=2, sticky="w", pady=6)
        ttk.Label(frm, text="Alertas atuais: " + (str(l.get("ALERTAS") or "") or "nenhum"),
                  wraplength=600, foreground="#555").grid(row=i + 2, column=0, columnspan=2, sticky="w")
        frm.columnconfigure(1, weight=1)

        def buscar_fornecedor(_e=None):
            doc = so_digitos(entradas["CNPJ_CPF"].get())
            forn = self.base.buscar_fornecedor(doc) if doc else None
            if forn:
                for k, val in (("NOME", forn["nome"]), ("CATEGORIA", forn["categoria"])):
                    entradas[k].delete(0, "end")
                    entradas[k].insert(0, val or "")
                aviso.config(text=f"Fornecedor encontrado na base: {forn['nome']}", foreground="#0a6b1f")
            elif doc:
                aviso.config(text="CNPJ/CPF não está na Base_Fornecedores (use 'Cadastrar fornecedor' depois).",
                             foreground="#b00020")
        entradas["CNPJ_CPF"].bind("<FocusOut>", buscar_fornecedor)

        def buscar_pelo_nome(_e=None):
            if so_digitos(entradas["CNPJ_CPF"].get()):
                return
            from .roteador import normalizar_texto
            alvo = normalizar_texto(entradas["NOME"].get())
            if len(alvo) < 4:
                return
            achados = [f for f in self.base.fornecedores.values()
                       if alvo in normalizar_texto(f["nome"]) or alvo in normalizar_texto(f["razao"])]
            if len(achados) == 1:
                f = achados[0]
                entradas["CNPJ_CPF"].insert(0, f["doc"])
                for k, val in (("NOME", f["nome"]), ("CATEGORIA", f["categoria"])):
                    entradas[k].delete(0, "end")
                    entradas[k].insert(0, val or "")
                aviso.config(text=f"Fornecedor encontrado pelo nome: {f['nome']} ({f['doc']}) - confira.",
                             foreground="#0a6b1f")
            elif achados:
                lista = "; ".join(f"{f['nome']} ({f['doc']})" for f in achados[:6])
                aviso.config(text=f"{len(achados)} fornecedores com esse nome: {lista}"
                                  + (" ..." if len(achados) > 6 else "") + " - digite o CNPJ.",
                             foreground="#b00020")
        entradas["NOME"].bind("<FocusOut>", buscar_pelo_nome)

        def salvar():
            doc = so_digitos(entradas["CNPJ_CPF"].get())
            if doc and not (cnpj_valido(doc) or cpf_valido(doc)):
                aviso.config(text="CNPJ/CPF inválido.", foreground="#b00020")
                return
            valor_txt = entradas["VALOR"].get().strip()
            novo_valor = _valor(valor_txt) if valor_txt else None
            if valor_txt and novo_valor is None:
                aviso.config(text="Valor inválido. Use, por exemplo, 1.234,56", foreground="#b00020")
                return
            venc_txt = entradas["DT_VENCTO"].get().strip()
            novo_venc = _data(venc_txt) if venc_txt else None
            if venc_txt and not novo_venc:
                aviso.config(text="Vencimento inválido. Use dd/mm/aaaa.", foreground="#b00020")
                return
            for k in ("NOME", "REFERENCIA", "NF", "CATEGORIA", "TP_DESP", "DADOS_BANCARIOS"):
                l[k] = entradas[k].get().strip().upper() if k in ("NOME", "REFERENCIA") else entradas[k].get().strip()
            l["CNPJ_CPF"], l["VALOR"], l["DT_VENCTO"] = doc, novo_valor, novo_venc
            if cb.get() and cb.get() != str(l.get("CLIENTE") or ""):
                l["CLIENTE"], l["IDENTIFICACAO"] = cb.get(), "MANUAL"
            # alertas que deixam de valer
            alertas = [a.strip() for a in str(l.get("ALERTAS") or "").split("|") if a.strip()]
            remover = []
            if novo_valor is not None:
                remover.append("Sem VALOR")
            if novo_venc:
                remover.append("Sem vencimento")
            if doc:
                remover.append("CNPJ do fornecedor não identificado")
                if self.base.buscar_fornecedor(doc):
                    remover += ["Fornecedor ", "FILIAL NOVA"]
            if cb.get():
                remover.append("Confira/escolha o CLIENTE")
            if doc and novo_valor is not None and novo_venc:
                remover += ["NFS-e só em PDF", "NF só em PDF", "PDF é imagem", "PDF protegido",
                            "PDF não lido", "PDF muito grande", "Relido:"]
            alertas = [a for a in alertas if not any(a.startswith(r) for r in remover)]
            alertas.append(f"Editado pelo usuário em {datetime.now():%d/%m/%Y}")
            l["ALERTAS"] = " | ".join(dict.fromkeys(alertas))
            if self._gravar_triagem():
                self.escrever(f"Sugestão #{l['ID_SUGESTAO']} atualizada.")
                jan.destroy()
                self._preencher_lista()

        bt = ttk.Frame(jan, padding=(12, 0, 12, 12))
        bt.pack(fill="x")
        ttk.Button(bt, text="💾 Salvar", command=salvar).pack(side="left", padx=4)
        ttk.Button(bt, text="📄 Abrir documento", command=self.abrir_documento).pack(side="left", padx=4)
        ttk.Button(bt, text="Cancelar", command=jan.destroy).pack(side="right", padx=4)

    def restaurar(self):
        itens = [self.itens[i] for i in self.tree.selection()]
        itens = [l for l in itens if str(l.get("STATUS")).upper() == "IGNORADO"]
        if not itens:
            _msg("info", "Aviso", "Escolha o filtro 'Ignorados (para restaurar)' e selecione as sugestões.")
            return
        for l in itens:
            l["STATUS"] = "PENDENTE"
            partes = [a.strip() for a in str(l.get("ALERTAS") or "").split("|")
                      if a.strip() and not a.strip().startswith("IGNORADO pelo usuário")]
            l["ALERTAS"] = " | ".join(partes)
        if self._gravar_triagem():
            self.escrever(f"{len(itens)} sugestão(ões) voltaram para PENDENTE.")
        self._preencher_lista()

    def comparar(self):
        import re as _re
        l = self._um_selecionado()
        if not l:
            return
        m = _re.search(r"ID_LANCAMENTO (\d+)", str(l.get("ALERTAS") or ""))
        if not m:
            _msg("info", "Comparar", "Esta sugestão não aponta para um lançamento existente "
                                     "(não há 'ID_LANCAMENTO' nos alertas).")
            return
        id_lanc = m.group(1)
        existente = next((x for x in self.base._lancamentos_cliente(str(l.get("CLIENTE") or self.cliente))
                          if str(x.get("ID_LANCAMENTO") or "").split(".")[0] == id_lanc), None)
        if not existente:
            _msg("warning", "Comparar", f"Lançamento {id_lanc} não encontrado na planilha do cliente.")
            return

        def fmt_v(v):
            v = _valor(v)
            return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if v is not None else ""

        def fmt_d(v):
            d = _data(v) or (_data(str(v)) if v else None)
            try:
                from .roteador import _para_data
                d = d or _para_data(v)
            except Exception:
                pass
            return d.strftime("%d/%m/%Y") if d else str(v or "")

        linhas = [
            ("Fornecedor", l.get("NOME"), existente.get("NOME")),
            ("CNPJ/CPF", l.get("CNPJ_CPF"), existente.get("CNPJ_CPF")),
            ("NF", l.get("NF"), existente.get("NF")),
            ("Valor", fmt_v(l.get("VALOR")), fmt_v(existente.get("VALOR"))),
            ("Vencimento", fmt_d(l.get("DT_VENCTO")), fmt_d(existente.get("DT_VENCTO"))),
            ("Referência", l.get("REFERENCIA"), existente.get("REFERENCIA")),
            ("Data do relatório", "", fmt_d(existente.get("DATA_REL"))),
            ("Observação", "", existente.get("OBSERVACAO")),
        ]
        jan = tk.Toplevel(self.jan)
        jan.title(f"Comparar sugestão #{l['ID_SUGESTAO']} com lançamento {id_lanc}")
        jan.geometry("900x330")
        jan.transient(self.jan)
        tv = ttk.Treeview(jan, columns=("c", "s", "e"), show="headings", height=9)
        for c, t, w in (("c", "", 140), ("s", f"Sugestão #{l['ID_SUGESTAO']} (e-mail)", 360),
                        ("e", f"Lançamento {id_lanc} (planilha)", 360)):
            tv.heading(c, text=t)
            tv.column(c, width=w, anchor="w")
        tv.tag_configure("dif", background="#FFF3CD")
        for campo, a, b in linhas:
            diferente = str(a or "").strip() != str(b or "").strip() and a not in ("", None)
            tv.insert("", "end", values=(campo, a or "", b or ""), tags=("dif",) if diferente else ())
        tv.pack(fill="both", expand=True, padx=10, pady=8)
        ttk.Label(jan, text="Em amarelo: campos diferentes.").pack(anchor="w", padx=10)
        bt = ttk.Frame(jan)
        bt.pack(fill="x", padx=10, pady=8)

        def eh_duplicado():
            jan.destroy()
            self.ignorar([l])

        def nao_eh():
            partes = [a.strip() for a in str(l.get("ALERTAS") or "").split("|")
                      if a.strip() and not a.strip().startswith("POSSÍVEL DUPLICIDADE")]
            partes.append(f"Duplicidade descartada pelo usuário (lançamento {id_lanc} é outra conta)")
            l["ALERTAS"] = " | ".join(partes)
            if self._gravar_triagem():
                self.escrever(f"Sugestão #{l['ID_SUGESTAO']}: duplicidade descartada; continua pendente.")
            jan.destroy()
            self._preencher_lista()

        ttk.Button(bt, text="🚫 É duplicado: ignorar a sugestão", command=eh_duplicado).pack(side="left", padx=4)
        ttk.Button(bt, text="✅ Não é duplicado: manter pendente", command=nao_eh).pack(side="left", padx=4)
        ttk.Button(bt, text="📄 Abrir documento", command=self.abrir_documento).pack(side="left", padx=4)
        ttk.Button(bt, text="Fechar", command=jan.destroy).pack(side="right", padx=4)

    # ---------------------------------------------------------------------
    def importar(self):
        selecionados = [self.itens[i] for i in self.tree.selection()]
        if not selecionados:
            _msg("warning", "Aviso", "Selecione ao menos um item.")
            return

        # segurança: sugestão já gravada (por exemplo, importada antes nesta mesma sessão)
        if self._conferir_gravados():
            selecionados = [l for l in selecionados if str(l.get("STATUS")).upper() == "PENDENTE"]
            if not selecionados:
                self._preencher_lista()
                _msg("info", "Já importado", "As sugestões selecionadas já estão gravadas na planilha do cliente.")
                return

        # segurança extra: nunca envia o mesmo boleto duas vezes no mesmo lote
        unicos, linhas = [], set()
        for l in selecionados:
            ld = str(l.get("LINHA_DIGITAVEL") or "").strip()
            if ld and ld in linhas:
                continue
            linhas.add(ld)
            unicos.append(l)
        if len(unicos) < len(selecionados):
            self.escrever(f"{len(selecionados) - len(unicos)} item(ns) com boleto repetido foram desconsiderados.")
        selecionados = unicos

        data_rel = _data_relatorio(self.sistema)
        for l in selecionados:
            l["_DATA_REL"] = data_rel_do_item(l.get("DT_VENCTO"), data_rel)
            manual = str(l.get("TP_DESP") or "").strip() if "Editado pelo usuário" in str(l.get("ALERTAS") or "") else ""
            l["_TP"] = manual or tp_desp_por_data(l, l["_DATA_REL"]) or str(l.get("TP_DESP") or "").strip()
        incompletos = [l for l in selecionados
                       if _valor(l.get("VALOR")) is None or not _data(l.get("DT_VENCTO"))
                       or not so_digitos(l.get("CNPJ_CPF")) or not l.get("NOME") or not l.get("_TP")]
        if incompletos:
            def faltas(l):
                f = []
                if _valor(l.get("VALOR")) is None: f.append("valor")
                if not _data(l.get("DT_VENCTO")): f.append("vencimento")
                if not so_digitos(l.get("CNPJ_CPF")): f.append("CNPJ/CPF")
                if not l.get("NOME"): f.append("nome")
                if not l.get("_TP"):
                    f.append("TP_DESP (a regra de datas não se aplica: e-mail na DATA_REL ou depois, ou sem data)")
                return ", ".join(f)
            linhas = "\n".join(f"• Item {l['ID_SUGESTAO']}: falta {faltas(l)}" for l in incompletos[:10])
            _msg("warning", "Itens incompletos",
                 f"{linhas}\n\nComplete com o botão '📝 Editar sugestão'. "
                 f"Os demais itens selecionados serão importados agora.")
            selecionados = [l for l in selecionados if l not in incompletos]
            if not selecionados:
                return

        suspeitos = [l for l in selecionados if "POSSÍVEL DUPLICIDADE" in str(l.get("ALERTAS") or "")]
        if suspeitos:
            lista = "\n".join(f"• Item {l['ID_SUGESTAO']}: {str(l.get('ALERTAS')).split(' | ')[0]}"
                              for l in suspeitos[:8])
            if not _msg("yesno", "Possível duplicidade",
                        f"Estes itens têm a mesma NF, valor e vencimento de um lançamento já "
                        f"existente, com CNPJ de outra empresa:\n\n{lista}\n\n"
                        f"Importar mesmo assim?\n(Não = desmarque-os e marque IGNORADO na triagem)"):
                return

        sem_cadastro = [l for l in selecionados if not self.base.buscar_fornecedor(l.get("CNPJ_CPF"))]
        if sem_cadastro:
            nomes = "\n".join(f"• {l.get('NOME')} ({l.get('CNPJ_CPF')})" for l in sem_cadastro[:10])
            if not _msg("yesno", "Fornecedor não cadastrado",
                        f"Estes fornecedores não estão na Base_Fornecedores:\n\n{nomes}\n\n"
                        f"Recomendado: cadastrá-los antes. Importar mesmo assim?"):
                return

        # itens que estavam sem cliente passam a pertencer ao cliente atual
        sem_cliente = [l for l in selecionados if not str(l.get("CLIENTE") or "").strip()]
        if sem_cliente:
            for l in sem_cliente:
                l["CLIENTE"], l["IDENTIFICACAO"] = self.cliente, "MANUAL"
            try:
                self.caixa.salvar(self.base.nomes_clientes())
            except ArquivoEmUso as e:
                for l in sem_cliente:
                    l["CLIENTE"] = ""
                _msg("error", "Planilha de triagem aberta", str(e))
                return

        registros = [_para_registro(l, l.get("_DATA_REL") or data_rel) for l in selecionados]
        from collections import Counter
        por_rel = Counter(r["data"] for r in registros)
        self.escrever("DATA_REL pelo vencimento: " + ", ".join(
            f"{qtd} item(ns) em {d}" for d, qtd in sorted(por_rel.items(), key=lambda x: _data(x[0]))))
        tps = [r["tp_desp"] for r in registros]
        self.escrever(f"TP_DESP pela regra de datas: "
                      f"{tps.count('3')} com tipo 3, {tps.count('5')} com tipo 5"
                      + (f", {len(tps) - tps.count('3') - tps.count('5')} com outro tipo informado." if
                         len(tps) - tps.count('3') - tps.count('5') else "."))

        if hasattr(self.sistema, "solicitar_etapa_obra"):
            etapa = self.sistema.solicitar_etapa_obra()
            if etapa is None:
                _msg("info", "Importação cancelada", "Nenhum item foi importado.")
                return
            for r in registros:
                r["etapa_obra"] = etapa

        self.sistema.dados_para_incluir.extend(registros)
        self.escrever(f"{len(registros)} item(ns) adicionados à lista de envio.")
        self.recarregar()
        self._fechar_se_vazia_apos_importar = True

        if _msg("yesno", "Importação por e-mail",
                f"{len(registros)} lançamento(s) prontos.\n\n"
                f"Eles só serão gravados na planilha do cliente quando você clicar em ENVIAR.\n\n"
                f"Deseja visualizar agora?"):
            self.sistema.visualizar_lancamentos()
        if getattr(self, "_fechar_se_vazia_apos_importar", False):
            self._fechar_se_vazia_apos_importar = False
            self._fechar_se_vazia()

    # ---------------------------------------------------------------------
    def _um_selecionado(self):
        sel = self.tree.selection()
        if len(sel) != 1:
            _msg("warning", "Aviso", "Selecione UMA linha.")
            return None
        return self.itens[sel[0]]

    def abrir_documento(self):
        l = self._um_selecionado()
        if not l:
            return
        caminhos = [c.strip() for c in str(l.get("ARQUIVOS") or "").split(" ; ") if c.strip()]
        existentes = [cfg.resolver_arquivo(c) for c in caminhos if os.path.exists(cfg.resolver_arquivo(c))]
        if not existentes:
            _msg("warning", "Documento", "Arquivo não encontrado:\n" + "\n".join(caminhos or ["(sem arquivo)"]))
            return
        # PDF primeiro (é o que se lê); XML depois
        for c in sorted(existentes, key=lambda x: not x.lower().endswith(".pdf"))[:2]:
            try:
                abrir_documento(c)
            except Exception as e:
                _msg("error", "Erro", f"Não foi possível abrir {c}:\n{e}")

    def abrir_email(self):
        l = self._um_selecionado()
        if not l:
            return
        msg_id = str(l.get("MSG_ID") or "")
        if not msg_id or msg_id.startswith("ARQ-"):
            _msg("info", "E-mail", "Este item veio de um arquivo local, não de e-mail. Use 'Abrir documento'.")
            return
        self.escrever(f"E-mail original: {l.get('EMAIL_REMETENTE')} | {l.get('EMAIL_ASSUNTO')}")
        webbrowser.open(f"https://mail.google.com/mail/u/0/#all/{msg_id}")

    def informar_cnpj(self):
        l = self._um_selecionado()
        if not l:
            return
        valor = simpledialog.askstring("Informar CNPJ/CPF",
                                       f"CNPJ ou CPF do fornecedor (item {l['ID_SUGESTAO']}):",
                                       parent=self.jan, initialvalue=l.get("CNPJ_CPF") or "")
        if valor is None:
            return
        doc = so_digitos(valor)
        if not (cnpj_valido(doc) or cpf_valido(doc)):
            _msg("error", "Inválido", f"{valor} não é um CNPJ/CPF válido.")
            return
        l["CNPJ_CPF"] = doc
        forn = self.base.buscar_fornecedor(doc)
        alertas = [a.strip() for a in str(l.get("ALERTAS") or "").split("|")
                   if a.strip() and not a.strip().startswith(("CNPJ do fornecedor", "Fornecedor ", "FILIAL NOVA"))]
        if forn:
            l["NOME"], l["CATEGORIA"] = forn["nome"], forn["categoria"] or l.get("CATEGORIA")
        else:
            alertas.append(f"Fornecedor {doc} NÃO cadastrado na Base_Fornecedores - cadastre antes de importar")
        l["ALERTAS"] = " | ".join(alertas)
        try:
            self.caixa.salvar(self.base.nomes_clientes())
        except ArquivoEmUso as e:
            _msg("error", "Planilha de triagem aberta", str(e))
            return
        self.recarregar()
        if not forn and _msg("yesno", "Fornecedor", "Fornecedor não cadastrado. Cadastrar agora?"):
            for iid, item in self.itens.items():
                if item is l or item.get("ID_SUGESTAO") == l.get("ID_SUGESTAO"):
                    self.tree.selection_set(iid)
            self.cadastrar_fornecedor()

    def cadastrar_fornecedor(self):
        l = self._um_selecionado()
        if not l:
            return
        doc = so_digitos(l.get("CNPJ_CPF"))
        if not doc:
            _msg("warning", "Sem CNPJ", "Este item não tem CNPJ. Use 'Abrir documento' para lê-lo "
                                        "e depois 'Informar CNPJ'.")
            return
        if self.base.buscar_fornecedor(doc):
            _msg("info", "Fornecedor", "Este fornecedor já está cadastrado. Clique em Recarregar.")
            return
        if not hasattr(self.sistema, "novo_fornecedor"):
            _msg("error", "Erro", "Tela de cadastro de fornecedor não encontrada no sistema.")
            return
        dados = dados_para_cadastro(l, self.base)
        self.sistema.novo_fornecedor()
        janela = getattr(self.sistema, "janela_fornecedor", None)
        campos = getattr(self.sistema, "campos_form", {})
        for chave in ("tipo_pessoa", "cnpj_cpf", "razao_social", "nome", "telefone",
                      "email", "endereco", "categoria", "vinculo"):
            _preencher(campos.get(chave), dados.get(chave), forcar=(chave == "tipo_pessoa"))
        p = dados["_parente"]
        self.escrever(f"Cadastro aberto com dados {'do XML' if dados['endereco'] else 'da sugestão'}"
                      + (f" e da filial já cadastrada {p['nome']} ({p['doc']})" if p else "")
                      + ". Complete o que faltar e salve.")
        if janela is not None:
            janela.wait_window()
            self.recarregar()

    # ---------------------------------------------------------------------
    def abrir_planilha(self):
        try:
            os.startfile(str(cfg.ARQUIVO_CAIXA))  # Windows
        except Exception as e:
            _msg("error", "Erro", f"Não foi possível abrir {cfg.ARQUIVO_CAIXA}:\n{e}")

    def ler_pasta(self):
        """Lê XML/PDF/ZIP/EML de uma pasta (e subpastas), em segundo plano: documentos
        recebidos por WhatsApp, baixados à mão, ou a pasta Documentos_Email copiada
        de outro ambiente."""
        from tkinter import filedialog
        from .coletor import executar_pasta_protegido
        pastas, inicial = [], str(cfg.PASTA_DOCUMENTOS)
        while True:
            pasta = filedialog.askdirectory(
                parent=self.jan, initialdir=inicial,
                title=f"Pasta {len(pastas) + 1} com documentos (subpastas incluídas)")
            if not pasta:
                break
            if pasta not in pastas:
                pastas.append(pasta)
            inicial = str(Path(pasta).parent)
            if not _msg("yesno", "Ler pastas",
                        "Pastas escolhidas:\n\n" + "\n".join(f"• {p}" for p in pastas)
                        + "\n\nAdicionar outra pasta?"):
                break
        if not pastas:
            return

        def progresso(t):
            try:
                self.jan.after(0, self.escrever, t)
            except Exception:
                pass

        def tarefa():
            try:
                for pasta in pastas:
                    progresso(f"--- Pasta: {pasta}")
                    executar_pasta_protegido(pasta, progresso=progresso)
            except ArquivoEmUso as e:
                progresso(str(e))
            except Exception as e:
                progresso(f"ERRO: {e}")
            finally:
                try:
                    self.jan.after(0, self.recarregar)
                except Exception:
                    pass

        self.escrever(f"Lendo {len(pastas)} pasta(s) em segundo plano...")
        threading.Thread(target=tarefa, daemon=True).start()

    def reprocessar(self):
        """Relê PDFs da triagem (boletos sem CNPJ, PDFs protegidos, clientes a confirmar)
        em segundo plano, sem travar a tela."""
        from .coletor import reprocessar_triagem, _COLETA

        def progresso(t):
            try:
                self.jan.after(0, self.escrever, t)
            except Exception:
                pass

        def tarefa():
            try:
                reprocessar_triagem(progresso=progresso)
            except ArquivoEmUso as e:
                progresso(str(e))
            except Exception as e:
                progresso(f"ERRO: {e}")
            finally:
                try:
                    self.jan.after(0, self.recarregar)
                except Exception:
                    pass

        self.escrever("Reprocessando a triagem em segundo plano (a tela continua livre)...")
        threading.Thread(target=tarefa, daemon=True).start()

    def buscar_emails(self):
        from .coletor import executar_gmail

        def progresso(t):
            try:  # se a janela foi fechada, a coleta continua e grava normalmente
                self.jan.after(0, self.escrever, t)
            except Exception:
                pass

        def tarefa():
            try:
                executar_gmail(progresso=progresso)
            except ArquivoEmUso as e:
                progresso(str(e))
            except Exception as e:
                progresso(f"ERRO: {e}")
            finally:
                try:
                    self.jan.after(0, self.recarregar)
                except Exception:
                    pass  # janela fechada durante a coleta

        self.escrever("Buscando no Gmail... (na 1ª vez o navegador abrirá para autorização)")
        threading.Thread(target=tarefa, daemon=True).start()


def abrir_importacao_email(sistema):
    if not getattr(sistema, "cliente_atual", None):
        _msg("warning", "Aviso", "Selecione um cliente antes de importar documentos do e-mail.")
        return
    JanelaImportacaoEmail(sistema)
