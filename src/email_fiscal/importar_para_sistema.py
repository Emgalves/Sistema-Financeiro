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

VERSAO = "2026-09-26.4"   # igual em todos os arquivos do módulo
import os
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


def sincronizar_status(caixa, base):
    """Marca IMPORTADO (marcador encontrado) ou JA_LANCADO (lançamento igual feito à mão)."""
    alterou = 0
    for l in caixa.lancamentos:
        if str(l.get("STATUS") or "").upper() != "PENDENTE" or not l.get("CLIENTE"):
            continue
        cliente = str(l["CLIENTE"]).strip()
        achou, _ = base.ja_lancado(cliente, id_sugestao=l.get("ID_SUGESTAO"))
        if achou:
            l["STATUS"], l["DATA_IMPORTACAO"] = "IMPORTADO", datetime.now()
            alterou += 1
            continue
        achou, motivo = base.ja_lancado(cliente, doc=l.get("CNPJ_CPF"), nf=l.get("NF"),
                                        valor=_valor(l.get("VALOR")), vencimento=_data(l.get("DT_VENCTO")))
        if achou:
            l["STATUS"] = "JA_LANCADO"
            l["ALERTAS"] = f"Lançado manualmente: {motivo} | {l.get('ALERTAS') or ''}".strip(" |")
            alterou += 1
            continue
        dup = base.possivel_duplicidade(cliente, doc=l.get("CNPJ_CPF"), nf=l.get("NF"),
                                        valor=_valor(l.get("VALOR")), vencimento=_data(l.get("DT_VENCTO")))
        alertas = str(l.get("ALERTAS") or "")
        if dup and "POSSÍVEL DUPLICIDADE" not in alertas:
            l["ALERTAS"] = f"{dup} | {alertas}".strip(" |")
            alterou += 1
    return alterou


def tp_desp_por_data(l, data_rel_str):
    """Regra do usuário (ver config_email.REGRA_TP_DESP_POR_DATA). None = regra não se aplica."""
    if not cfg.REGRA_TP_DESP_POR_DATA or str(l.get("TIPO_DOC")) not in cfg.TIPOS_COM_REGRA_POR_DATA:
        return None
    data_rel, email, venc = _data(data_rel_str), _data(l.get("EMAIL_DATA")), _data(l.get("DT_VENCTO"))
    if not (data_rel and email and venc) or email >= data_rel:
        return None
    return "3" if venc > data_rel else "5"


def _para_registro(l, data_rel):
    """Converte a linha da triagem no dicionário que o SistemaEntradaDados usa."""
    valor = _valor(l.get("VALOR"))
    venc = _data(l.get("DT_VENCTO"))
    return {
        "data": data_rel,
        "cnpj_cpf": so_digitos(l.get("CNPJ_CPF")),
        "nome": str(l.get("NOME") or "").strip(),
        "categoria": str(l.get("CATEGORIA") or "").strip(),
        "tp_desp": l.get("_TP") or str(l.get("TP_DESP") or "").strip(),
        "referencia": str(l.get("REFERENCIA") or "").strip(),
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

    def __init__(self, sistema):
        self.sistema = sistema
        self.cliente = getattr(sistema, "cliente_atual", None)
        self.jan = tk.Toplevel(sistema.root)
        self.jan.title(f"Documentos recebidos por e-mail - {self.cliente}")
        self.jan.geometry("1250x560")
        self.jan.transient(sistema.root)

        topo = ttk.Frame(self.jan)
        topo.pack(fill="x", padx=10, pady=6)
        ttk.Label(topo, text=f"Cliente: {self.cliente}", font=("Helvetica", 11, "bold")).pack(side="left")
        self.var_sem_cliente = tk.BooleanVar(value=False)
        ttk.Checkbutton(topo, text="Mostrar também itens SEM cliente definido",
                        variable=self.var_sem_cliente, command=self.recarregar).pack(side="left", padx=20)

        meio = ttk.Frame(self.jan)
        meio.pack(fill="both", expand=True, padx=10)
        self.tree = ttk.Treeview(meio, columns=self.COLS, show="headings", selectmode="extended")
        larg = (50, 110, 230, 250, 70, 90, 85, 340)
        for c, w in zip(self.COLS, larg):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="e" if c == "VALOR" else "w")
        self.tree.tag_configure("alerta", background="#FFF3CD")
        sb = ttk.Scrollbar(meio, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", lambda e: self.abrir_documento())
        sb.pack(side="right", fill="y")

        self.log = tk.Text(self.jan, height=6, font=("Consolas", 9))
        self.log.pack(fill="x", padx=10, pady=4)

        rodape = ttk.Frame(self.jan)
        rodape.pack(fill="x", padx=10, pady=6)
        ttk.Button(rodape, text="✅ Importar selecionados", command=self.importar).pack(side="left", padx=4)
        ttk.Button(rodape, text="📄 Abrir documento", command=self.abrir_documento).pack(side="left", padx=4)
        ttk.Button(rodape, text="✉️ Abrir e-mail", command=self.abrir_email).pack(side="left", padx=4)
        ttk.Button(rodape, text="✏️ Informar CNPJ", command=self.informar_cnpj).pack(side="left", padx=4)
        ttk.Button(rodape, text="🏢 Cadastrar fornecedor", command=self.cadastrar_fornecedor).pack(side="left", padx=4)
        ttk.Button(rodape, text="📥 Buscar e-mails agora", command=self.buscar_emails).pack(side="left", padx=4)
        ttk.Button(rodape, text="📊 Abrir planilha de triagem", command=self.abrir_planilha).pack(side="left", padx=4)
        ttk.Button(rodape, text="🔄 Recarregar", command=self.recarregar).pack(side="left", padx=4)
        ttk.Button(rodape, text="Fechar", command=self.jan.destroy).pack(side="right", padx=4)

        self._conferir_versoes()
        self.recarregar()

    def _conferir_versoes(self):
        """Avisa se algum arquivo do módulo ficou de uma versão anterior."""
        import importlib
        desatualizados = []
        for nome in ("config_email", "gmail_cliente", "leitor_nfe", "leitor_boleto",
                     "roteador", "caixa_entrada", "coletor"):
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
        from .coletor import reler_pendentes
        rel, tocados = reler_pendentes(self.caixa, self.base)
        if rel:
            self.escrever(f"{rel} sugestão(ões) antiga(s) relida(s) do PDF e corrigida(s).")
        rep = marcar_repetidos(self.caixa)
        if rep:
            self.escrever(f"{rep} sugestão(ões) repetida(s) marcada(s) como IGNORADO.")
        uni = self.caixa.unir_boletos_e_notas()
        if uni:
            self.escrever(f"{uni} boleto(s) unido(s) à nota fiscal correspondente (e-mails diferentes).")
        n = sincronizar_status(self.caixa, self.base) + rep + uni + tocados
        if n:
            try:
                self.caixa.salvar(self.base.nomes_clientes())
                self.escrever(f"Triagem atualizada ({n} alteração(ões)).")
            except ArquivoEmUso as e:
                self.escrever(str(e))

        em_andamento = {str(r.get("observacao", "")).split(" - ")[0]
                        for r in (getattr(self.sistema, "dados_para_incluir", None) or [])}
        self.itens = {}
        self.tree.delete(*self.tree.get_children())
        for l in self.caixa.pendentes(self.cliente, self.var_sem_cliente.get()):
            if f"{cfg.PREFIXO_OBSERVACAO} #{l['ID_SUGESTAO']}" in em_andamento:
                continue  # já está na lista de envio desta sessão
            v, d = _valor(l.get("VALOR")), _data(l.get("DT_VENCTO"))
            iid = self.tree.insert("", "end", values=(
                l["ID_SUGESTAO"], l.get("IDENTIFICACAO") or "", l.get("NOME") or "",
                l.get("REFERENCIA") or "", l.get("NF") or "",
                f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if v is not None else "",
                d.strftime("%d/%m/%Y") if d else "", l.get("ALERTAS") or ""),
                tags=("alerta",) if l.get("ALERTAS") else ())
            self.itens[iid] = l
        self.tree.selection_set(self.tree.get_children())
        self.escrever(f"{len(self.itens)} sugestão(ões) pendente(s) para {self.cliente}.")

    # ---------------------------------------------------------------------
    def importar(self):
        selecionados = [self.itens[i] for i in self.tree.selection()]
        if not selecionados:
            _msg("warning", "Aviso", "Selecione ao menos um item.")
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
            l["_TP"] = tp_desp_por_data(l, data_rel) or str(l.get("TP_DESP") or "").strip()
        incompletos = [l for l in selecionados
                       if _valor(l.get("VALOR")) is None or not _data(l.get("DT_VENCTO"))
                       or not so_digitos(l.get("CNPJ_CPF")) or not l.get("NOME") or not l.get("_TP")]
        if incompletos:
            ids = ", ".join(str(l["ID_SUGESTAO"]) for l in incompletos)
            _msg("warning", "Itens incompletos",
                 f"Os itens {ids} estão sem VALOR, DT_VENCTO, CNPJ_CPF, NOME ou TP_DESP.\n"
                 f"(TP_DESP só é calculado quando a data do e-mail é anterior à DATA_REL {data_rel};\n"
                 f"nos outros casos, preencha TP_DESP na planilha.)\n\n"
                 f"Complete-os na planilha de triagem (botão 'Abrir planilha'), salve, feche o Excel "
                 f"e clique em Recarregar. Os demais itens serão importados agora.")
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

        registros = [_para_registro(l, data_rel) for l in selecionados]
        tps = [r["tp_desp"] for r in registros]
        self.escrever(f"TP_DESP pela regra de datas (DATA_REL {data_rel}): "
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
        self.escrever(f"{len(registros)} item(ns) adicionados à lista de envio (DATA_REL {data_rel}).")
        self.recarregar()

        if _msg("yesno", "Importação por e-mail",
                f"{len(registros)} lançamento(s) prontos.\n\n"
                f"Eles só serão gravados na planilha do cliente quando você clicar em ENVIAR.\n\n"
                f"Deseja visualizar agora?"):
            self.sistema.visualizar_lancamentos()

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
        existentes = [c for c in caminhos if os.path.exists(c)]
        if not existentes:
            _msg("warning", "Documento", "Arquivo não encontrado:\n" + "\n".join(caminhos or ["(sem arquivo)"]))
            return
        # PDF primeiro (é o que se lê); XML depois
        for c in sorted(existentes, key=lambda x: not x.lower().endswith(".pdf"))[:2]:
            try:
                os.startfile(c)
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

    def buscar_emails(self):
        from .coletor import executar_gmail

        def progresso(t):
            self.jan.after(0, self.escrever, t)

        def tarefa():
            try:
                executar_gmail(progresso=progresso)
            except ArquivoEmUso as e:
                progresso(str(e))
            except Exception as e:
                progresso(f"ERRO: {e}")
            finally:
                self.jan.after(0, self.recarregar)

        self.escrever("Buscando no Gmail... (na 1ª vez o navegador abrirá para autorização)")
        threading.Thread(target=tarefa, daemon=True).start()


def abrir_importacao_email(sistema):
    if not getattr(sistema, "cliente_atual", None):
        _msg("warning", "Aviso", "Selecione um cliente antes de importar documentos do e-mail.")
        return
    JanelaImportacaoEmail(sistema)
