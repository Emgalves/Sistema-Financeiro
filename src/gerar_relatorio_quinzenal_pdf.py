"""
Relatório Quinzenal de Medições em PDF
Layout redesenhado — extrato completo por fornecedor
"""

import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, PageBreak
)
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT

# ---------------------------------------------------------------------------
# Paleta
# ---------------------------------------------------------------------------
AZUL        = colors.HexColor('#1F4788')
AZUL_CLARO  = colors.HexColor('#E8EAF6')
CINZA_CLARO = colors.HexColor('#F5F5F5')
CINZA_BORDA = colors.HexColor('#CCCCCC')
AMARELO     = colors.HexColor('#FFF3CD')
BRANCO      = colors.white
PRETO       = colors.black
VERMELHO    = colors.HexColor('#CC0000')
VERDE       = colors.HexColor('#006600')

PAGE_W, PAGE_H = A4
MARGIN_L = MARGIN_R = 15 * mm
MARGIN_T = 48 * mm
MARGIN_B = 22 * mm
BODY_W   = PAGE_W - MARGIN_L - MARGIN_R


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _limpar(texto):
    if not texto or (isinstance(texto, float)):
        return ''
    return ' '.join(str(texto).split())


def _normalizar_doc(doc):
    import re
    return re.sub(r'\D', '', str(doc or ''))


def _fmt_moeda(valor):
    try:
        return "R$ {:,.2f}".format(float(valor)).replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return 'R$ 0,00'


def _fmt_num(valor):
    """Valor monetário sem símbolo R$ — para colunas estreitas."""
    try:
        return "{:,.2f}".format(float(valor)).replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return '0,00'


def _fmt_data(data):
    if data is None:
        return ''
    if isinstance(data, str):
        try:
            data = pd.to_datetime(data)
        except Exception:
            return data
    try:
        if pd.isna(data):
            return ''
    except Exception:
        pass
    try:
        return data.strftime('%d/%m/%Y')
    except Exception:
        return str(data)

def _fmt_doc(doc, tipo_pessoa=None):
    """
    Formata CPF ou CNPJ com mascara.

    tipo_pessoa: valor da coluna tipo_pessoa da base_fornecedores.xlsx.
      - 'F', 'PF' ou 'FISICA'   -> formata como CPF  (000.000.000-00)
      - 'J', 'PJ' ou 'JURIDICA' -> formata como CNPJ (00.000.000/0000-00)
      - None -> fallback por contagem de digitos (11=CPF, 14=CNPJ)
    Garante zeros a esquerda antes de aplicar a mascara.
    """
    if not doc:
        return ''
    d = _normalizar_doc(doc)
    if not d:
        return str(doc)

    # Determinar tipo pelo campo tipo_pessoa quando disponivel
    tp = str(tipo_pessoa or '').strip().upper()
    is_cpf  = tp in ('F', 'PF', 'FISICA', 'PESSOA FISICA')
    is_cnpj = tp in ('J', 'PJ', 'JURIDICA', 'PESSOA JURIDICA')

    # Fallback por contagem de digitos se tipo_pessoa nao for reconhecido
    if not is_cpf and not is_cnpj:
        is_cpf  = len(d) <= 11
        is_cnpj = not is_cpf

    if is_cpf:
        d = d.zfill(11)  # garante 11 digitos com zeros a esquerda
        return f'{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}'
    else:
        d = d.zfill(14)  # garante 14 digitos com zeros a esquerda
        return f'{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}'



def _fmt_fone(fone):
    """Formata telefone: (XX) XXXXX-XXXX (celular) ou (XX) XXXX-XXXX (fixo)."""
    if not fone:
        return ''
    import re
    d = re.sub(r'\D', '', str(fone))
    if not d:
        return str(fone)
    # Remover prefixo +55 se presente
    if d.startswith('55') and len(d) in (12, 13):
        d = d[2:]
    if len(d) == 11:
        return f'({d[:2]}) {d[2:7]}-{d[7:]}'
    if len(d) == 10:
        return f'({d[:2]}) {d[2:6]}-{d[6:]}'
    return str(fone)  # retorna como esta se nao reconhecer

def _p(texto, fs=8, bold=False, align=TA_LEFT, color=PRETO, leading=None):
    """Atalho para criar Paragraph com estilo inline."""
    fname = 'Helvetica-Bold' if bold else 'Helvetica'
    kwargs = dict(fontSize=fs, fontName=fname, textColor=color, alignment=align)
    if leading:
        kwargs['leading'] = leading
    st = ParagraphStyle('_', **kwargs)
    return Paragraph(str(texto), st)


def _titulo_secao(texto):
    """Barra azul com texto branco — título de seção."""
    st = ParagraphStyle('ts', fontSize=9, fontName='Helvetica-Bold',
                        textColor=BRANCO, alignment=TA_CENTER)
    tbl = Table([[Paragraph(texto, st)]], colWidths=[BODY_W])
    tbl.setStyle(TableStyle([
        ('BACKGROUND',    (0, 0), (-1, -1), AZUL),
        ('TOPPADDING',    (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING',   (0, 0), (-1, -1), 6),
    ]))
    return tbl


# ---------------------------------------------------------------------------
# Classe principal
# ---------------------------------------------------------------------------

class RelatorioQuinzenalPDF:

    EMPRESA = {
        'nome':     'VASCONCELOS&RINALDI',
        'linha2':   'ENGENHARIA',
        'endereco': 'Rua Zodiaco, 87 Sala 07 \u2013 Santa L\u00facia - Belo Horizonte - MG',
        'fones':    '(31) 3654-6616 / (31) 99974-1241 / (31) 98711-1139',
        'email':    'rvr.engenharia@gmail.com',
    }

    def __init__(self, arquivo_cliente, arquivo_clientes, arquivo_fornecedores=None):
        self.arquivo_cliente      = Path(arquivo_cliente)
        self.arquivo_clientes     = Path(arquivo_clientes)
        self.arquivo_fornecedores = (
            Path(arquivo_fornecedores) if arquivo_fornecedores
            else self._localizar_fornecedores()
        )
        self.cliente_info = {}
        self.logo_path    = Path(__file__).parent / 'logo.png'

    # ------------------------------------------------------------------
    def _localizar_fornecedores(self):
        """
        Resolve o caminho para base_fornecedores.xlsx em ordem de prioridade:
        1. ARQUIVO_FORNECEDORES definido em src.config.config (fonte canonica do sistema)
        2. Caminhos relativos ao arquivo do cliente (fallback legado)
        """
        # Prioridade 1: constante central do sistema
        try:
            from src.config.config import ARQUIVO_FORNECEDORES
            p = Path(ARQUIVO_FORNECEDORES)
            if p.exists():
                return p
        except Exception:
            pass
        # Prioridade 2: caminhos relativos ao arquivo do cliente
        base = self.arquivo_cliente.parent.parent
        for c in [
            base / 'base_fornecedores.xlsx',
            base / 'Fornecedores' / 'base_fornecedores.xlsx',
            Path(__file__).parent.parent / 'base_fornecedores.xlsx',
        ]:
            if c.exists():
                return c
        return None

    # ------------------------------------------------------------------
    # Carga de dados
    # ------------------------------------------------------------------

    def _carregar_cliente(self):
        """
        Tenta carregar nome e endereço do cliente em três etapas:
        1. Aba 'Clientes' do arquivo_clientes (sistema padrão)
        2. Aba 'RESUMO' do próprio arquivo_cliente (fallback — linhas 3 e 4)
        3. Nome derivado do nome do arquivo
        """
        nome_arq = self.arquivo_cliente.stem.replace('_', ' ')

        # Tentativa 1: arquivo_clientes / aba Clientes
        try:
            df = pd.read_excel(self.arquivo_clientes, sheet_name='Clientes')
            for _, row in df.iterrows():
                if nome_arq.upper() in str(row.get('Nome', '')).upper():
                    self.cliente_info = {
                        'nome':     str(row['Nome']).strip(),
                        'endereco': _limpar(row.get('Endere\u00e7o', '')),
                    }
                    return
        except Exception:
            pass

        # Tentativa 2: aba RESUMO do próprio arquivo_cliente (row 3 = nome, row 4 = endereço)
        try:
            wb = __import__('openpyxl').load_workbook(self.arquivo_cliente, data_only=True)
            if 'RESUMO' in wb.sheetnames:
                ws = wb['RESUMO']
                # Nome está na célula A3 (row 3, col 1)
                nome_resumo = ws.cell(row=3, column=1).value
                end_resumo  = ws.cell(row=4, column=1).value
                if nome_resumo:
                    self.cliente_info = {
                        'nome':     str(nome_resumo).strip(),
                        'endereco': _limpar(end_resumo or ''),
                    }
                    return
        except Exception as e:
            print(f'Aviso RESUMO: {e}')

        # Tentativa 3: derivar do nome do arquivo
        self.cliente_info = {'nome': nome_arq, 'endereco': ''}

    def _carregar_fornecedor_base(self, cnpj):
        if not self.arquivo_fornecedores or not self.arquivo_fornecedores.exists():
            return {}
        try:
            cnpj_num = _normalizar_doc(cnpj)
            df = pd.read_excel(self.arquivo_fornecedores, sheet_name='Fornecedores')
            for _, row in df.iterrows():
                if _normalizar_doc(row.get('CNPJ/CPF', '')).lstrip('0') == cnpj_num.lstrip('0') and _normalizar_doc(row.get('CNPJ/CPF', '')):
                    # NOME tem prioridade; RAZAO SOCIAL como fallback
                    nome = (_limpar(row.get('NOME', ''))
                            or _limpar(row.get('RAZÃO SOCIAL', '')))
                    # Telefone: tentar com e sem dois-pontos no cabecalho
                    telefone = _fmt_fone(
                        _limpar(row.get('TELEFONE:', ''))
                        or _limpar(row.get('TELEFONE', ''))
                    )
                    # tipo_pessoa define se e CPF ou CNPJ para formatacao correta
                    tipo_pessoa = row.get('tipo_pessoa', None)
                    return {
                        'nome':            nome,
                        'cnpj':            _fmt_doc(row.get('CNPJ/CPF', ''), tipo_pessoa),
                        'endereco':        _limpar(row.get('ENDEREÇO', '')),
                        'telefone':        telefone,
                        'dados_bancarios': _limpar(row.get('DADOS BANCÁRIOS', '')),
                    }
        except Exception as e:
            print(f'Erro ao buscar fornecedor: {e}')
        return {}

    def _identificar_quinzena(self, data_ref):
        """
        Define a faixa de Prev. Pagto que pertence ao relatorio da data_ref.

        Logica de negocio:
        - Relatorio do dia 5  (data_ref com dia <= 9):
            cobre Prev. Pagto de 05 ate 19 do mesmo mes.
        - Relatorio do dia 20 (data_ref com dia >= 10):
            cobre Prev. Pagto de 20 do mes atual ate 04 do mes seguinte.

        O usuario informa qualquer data do mes; o codigo snap para o
        relatorio mais proximo (dia 5 ou dia 20).
        """
        dia, mes, ano = data_ref.day, data_ref.month, data_ref.year

        if dia <= 9:
            # Quinzena do dia 5: Prev. Pagto entre 05 e 19 do mesmo mes
            data_ini = datetime(ano, mes, 5)
            data_fim = datetime(ano, mes, 19)
        else:
            # Quinzena do dia 20: Prev. Pagto entre 20 do mes atual e 04 do mes seguinte
            data_ini = datetime(ano, mes, 20)
            if mes == 12:
                data_fim = datetime(ano + 1, 1, 4)
            else:
                data_fim = datetime(ano, mes + 1, 4)

        return data_ini, data_fim

    def _carregar_dados_planilha(self, data_ref, cnpj_filtro=None):
        data_ini, data_fim = self._identificar_quinzena(data_ref)
        cnpj_filtro_num = _normalizar_doc(cnpj_filtro) if cnpj_filtro else None

        df_c = pd.read_excel(self.arquivo_cliente, sheet_name='Contratos_Medicao')
        df_m = pd.read_excel(self.arquivo_cliente, sheet_name='Medicoes')

        # Excluir medições EXCLUÍDO
        df_m = df_m[
            df_m['Status'].isna() | (df_m['Status'] != 'EXCLU\u00cdDO')
        ].copy()

        df_m['Data_Medicao']   = pd.to_datetime(df_m['Data_Medicao'],   errors='coerce')
        df_m['Data_Pagamento'] = pd.to_datetime(df_m['Data_Pagamento'], errors='coerce')

        # IDs da quinzena: filtrar por Data_Pagamento (Prev. Pagto),
        # que e o criterio de negocio para definir a qual relatorio a medicao pertence.
        mask_q = (df_m['Data_Pagamento'] >= data_ini) & (df_m['Data_Pagamento'] <= data_fim)
        ids_q  = set(zip(df_m.loc[mask_q, 'ID_Contrato'], df_m.loc[mask_q, 'ID_Medicao']))
        # Conjunto dos ID_Contrato que tem medicao na quinzena (para filtrar fornecedores)
        contratos_com_medicao_q = set(df_m.loc[mask_q, 'ID_Contrato'].astype(int))

        # Construir sequência global por fornecedor (Nº/Emp)
        seq_global = {}
        contador   = {}
        for _, row in df_c.iterrows():
            n = _normalizar_doc(row['CNPJ_Fornecedor'])
            contador[n] = contador.get(n, 0) + 1
            seq_global[int(row['ID_Contrato'])] = contador[n]

        # Fornecedores distintos na planilha
        vistos = []
        seen   = set()
        for _, row in df_c.iterrows():
            n = _normalizar_doc(row['CNPJ_Fornecedor'])
            if n not in seen:
                seen.add(n)
                vistos.append((str(row['CNPJ_Fornecedor']), str(row['Nome_Fornecedor'])))

        resultado = []
        for cnpj_raw, nome_forn in vistos:
            cnpj_num = _normalizar_doc(cnpj_raw)

            # Aplicar filtro externo
            if cnpj_filtro_num and cnpj_num != cnpj_filtro_num:
                continue

            contratos_forn = df_c[
                df_c['CNPJ_Fornecedor'].apply(_normalizar_doc) == cnpj_num
            ]
            ids_forn = set(int(x) for x in contratos_forn['ID_Contrato'])

            # Só incluir se houver medição na quinzena para algum contrato deste fornecedor
            if not ids_forn.intersection(contratos_com_medicao_q):
                continue

            dados_base = self._carregar_fornecedor_base(cnpj_raw)

            contratos_lista = []
            for _, row_c in contratos_forn.sort_values('Data_Inicio').iterrows():
                id_c    = int(row_c['ID_Contrato'])
                num_emp = seq_global.get(id_c, '?')

                meds = df_m[df_m['ID_Contrato'] == id_c].sort_values('Data_Medicao').to_dict('records')
                for m in meds:
                    m['_quinzena'] = (id_c, int(m['ID_Medicao'])) in ids_q

                vg = float(row_c.get('Valor_Global', 0) or 0)
                vp = float(row_c.get('Valor_Pago',   0) or 0)

                contratos_lista.append({
                    'id':           id_c,
                    'num_emp':      num_emp,
                    'descricao':    _limpar(row_c.get('Descricao', '')),
                    'data_inicio':  row_c.get('Data_Inicio'),
                    'data_final':   row_c.get('Data_Final'),
                    'valor_global': vg,
                    'valor_pago':   vp,
                    'saldo':        vg - vp,
                    'status':       _limpar(row_c.get('Status', 'ATIVO')),
                    'medicoes':     meds,
                })

            total_global = sum(c['valor_global'] for c in contratos_lista)
            total_pago   = sum(c['valor_pago']   for c in contratos_lista)
            total_saldo  = total_global - total_pago
            perc_exec    = (total_pago / total_global * 100) if total_global > 0 else 0
            med_atual    = sum(
                float(m['Valor'] or 0)
                for c in contratos_lista
                for m in c['medicoes']
                if m['_quinzena']
            )

            # Extrato flat ordenado cronologicamente por data de pagamento
            todas_meds_flat = []
            for c in contratos_lista:
                for m in c['medicoes']:
                    todas_meds_flat.append((m['Data_Medicao'], c, m))
            todas_meds_flat.sort(key=lambda x: x[0] if pd.notna(x[0]) else datetime.min)

            extrato = []
            seq = 0
            for _dt, c, m in todas_meds_flat:
                seq += 1
                extrato.append({
                    'seq':        seq,
                    'num_emp':    c['num_emp'],
                    'id_medicao': int(m['ID_Medicao']),
                    'data_pagto': m['Data_Pagamento'],
                    'valor':      float(m['Valor'] or 0),
                    'referencia': _limpar(m['Referencia']),
                    '_quinzena':  m['_quinzena'],
                })

            resultado.append({
                'cnpj':           cnpj_raw,
                'nome':           _limpar(nome_forn),
                'dados_base':     dados_base,
                'contratos':      contratos_lista,
                'extrato':        extrato,
                'total_global':   total_global,
                'total_pago':     total_pago,
                'total_saldo':    total_saldo,
                'perc_exec':      perc_exec,
                'med_atual':      med_atual,
                'data_ini_quinz': data_ini,
                'data_fim_quinz': data_fim,
            })

        return resultado

    # ------------------------------------------------------------------
    # Elementos de conteúdo
    # ------------------------------------------------------------------

    def _bloco_contratada(self, forn):
        db    = forn['dados_base']
        nome  = db.get('nome')  or forn['nome']
        cnpj  = _fmt_doc(db.get('cnpj') or forn['cnpj'])
        end   = db.get('endereco', '')        or ''
        fone  = db.get('telefone', '')         or ''
        banco = db.get('dados_bancarios', '') or ''

        st_l = ParagraphStyle('l', fontSize=8, fontName='Helvetica-Bold', textColor=AZUL)
        st_v = ParagraphStyle('v', fontSize=8, fontName='Helvetica')

        cw   = [28*mm, 72*mm, 32*mm, 48*mm]
        rows = [
            [_p('NOME / RAZÃO SOCIAL:', 8, True, color=AZUL), _p(nome, 8),
             _p('CPF / CNPJ:', 8, True, color=AZUL),          _p(cnpj, 8)],
            [_p('ENDEREÇO:', 8, True, color=AZUL),             _p(end or '\u2014', 8),
             _p('RESPONSÁVEL:', 8, True, color=AZUL),           _p('', 8)],
            [_p('CONTATOS:', 8, True, color=AZUL),             _p(fone or '\u2014', 8),
             _p('DADOS PAGAMENTO:', 8, True, color=AZUL),      _p(banco or '\u2014', 8)],
        ]
        tbl = Table(rows, colWidths=cw)
        tbl.setStyle(TableStyle([
            ('TOPPADDING',    (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('LEFTPADDING',   (0, 0), (-1, -1), 4),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
            ('GRID',          (0, 0), (-1, -1), 0.4, CINZA_BORDA),
            ('BACKGROUND',    (0, 0), (0, -1), CINZA_CLARO),
            ('BACKGROUND',    (2, 0), (2, -1), CINZA_CLARO),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        return [_titulo_secao('DADOS DA CONTRATADA'), tbl]

    def _bloco_contratos(self, forn):
        st_cab = ParagraphStyle('cb', fontSize=7.5, fontName='Helvetica-Bold',
                                textColor=BRANCO, alignment=TA_CENTER)
        st_d   = ParagraphStyle('dc', fontSize=7.5, fontName='Helvetica', leading=9)

        cw  = [10*mm, 19*mm, 19*mm, 69*mm, 23*mm, 23*mm, 17*mm]
        cab = [
            _p('Nº', 7.5, True, TA_CENTER, BRANCO),
            _p('Início', 7.5, True, TA_CENTER, BRANCO),
            _p('Fim', 7.5, True, TA_CENTER, BRANCO),
            _p('Descrição do Serviço', 7.5, True, TA_CENTER, BRANCO),
            _p('Valor Global', 7.5, True, TA_CENTER, BRANCO),
            _p('Valor Pago', 7.5, True, TA_CENTER, BRANCO),
            _p('Saldo', 7.5, True, TA_CENTER, BRANCO),
        ]
        data = [cab]
        for c in forn['contratos']:
            tem_saldo = c['saldo'] > 0.01
            saldo_str = _fmt_num(c['saldo'])  # sem R$
            data.append([
                _p(str(c['num_emp']),           7.5, False, TA_CENTER),
                _p(_fmt_data(c['data_inicio']), 7.5, False, TA_CENTER),
                _p(_fmt_data(c['data_final']),  7.5, False, TA_CENTER),
                Paragraph(c['descricao'], st_d),
                _p(_fmt_num(c['valor_global']), 7.5, False, TA_RIGHT),
                _p(_fmt_num(c['valor_pago']),   7.5, False, TA_RIGHT),
                _p(saldo_str, 7.5, tem_saldo, TA_RIGHT),  # negrito se saldo aberto
            ])

        tbl = Table(data, colWidths=cw)
        style = [
            ('BACKGROUND',    (0, 0), (-1, 0), AZUL),
            ('TOPPADDING',    (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('LEFTPADDING',   (0, 0), (-1, -1), 3),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 3),
            ('GRID',          (0, 0), (-1, -1), 0.4, CINZA_BORDA),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ]
        for i in range(1, len(data)):
            style.append(('BACKGROUND', (0, i), (-1, i),
                           CINZA_CLARO if i % 2 == 0 else BRANCO))
        tbl.setStyle(TableStyle(style))
        return [_titulo_secao('DADOS DE CONTRATOS'), tbl]

    def _bloco_totais(self, forn):
        cw = [BODY_W / 4] * 4
        data = [
            [_p('VALOR GLOBAL',   8, True, TA_CENTER),
             _p('VALOR PAGO',     8, True, TA_CENTER),
             _p('SALDO',          8, True, TA_CENTER),
             _p('MEDI\u00c7\u00c3O ATUAL', 8, True, TA_CENTER)],
            [_p(_fmt_moeda(forn['total_global']), 9, True, TA_CENTER, AZUL),
             _p(_fmt_moeda(forn['total_pago']),   9, True, TA_CENTER, AZUL),
             _p(_fmt_moeda(forn['total_saldo']),  9, True, TA_CENTER, AZUL),
             _p(_fmt_moeda(forn['med_atual']),    9, True, TA_CENTER, VERMELHO)],
        ]
        tbl = Table(data, colWidths=cw)
        tbl.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (-1, 0), AZUL_CLARO),
            ('BACKGROUND',    (0, 1), (-1, 1), BRANCO),
            ('BOX',           (0, 0), (-1, -1), 1, AZUL),
            ('INNERGRID',     (0, 0), (-1, -1), 0.4, CINZA_BORDA),
            ('TOPPADDING',    (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        return tbl

    def _bloco_extrato(self, forn):
        titulo = (
            f"EXTRATO DE MEDI\u00c7\u00d5ES  \u2014  "
            # f"Prev. Pagto: {_fmt_data(forn['data_ini_quinz'])} a "
            # f"{_fmt_data(forn['data_fim_quinz'])}  \u2014  "
            f"% Executado: {forn['perc_exec']:.1f}%"
        )
        st_d = ParagraphStyle('at', fontSize=7.5, fontName='Helvetica', leading=9)

        cw  = [8*mm, 22*mm, 26*mm, 14*mm, 110*mm]
        cab = [
            _p('N\u00ba', 7.5, True, TA_CENTER, BRANCO),
            _p('Prev. Pagto', 7.5, True, TA_CENTER, BRANCO),
            _p('Valor Medido', 7.5, True, TA_CENTER, BRANCO),
            _p('Ref.', 7.5, True, TA_CENTER, BRANCO),
            _p('Atividade', 7.5, True, TA_CENTER, BRANCO),
        ]
        data = [cab]
        for m in forn['extrato']:
            data.append([
                _p(str(m['seq']),              7.5, False, TA_CENTER),
                _p(_fmt_data(m['data_pagto']), 7.5, False, TA_CENTER),
                _p(_fmt_num(m['valor']),       7.5, False, TA_RIGHT),
                _p(str(m['num_emp']),          7.5, False, TA_CENTER),
                Paragraph(m['referencia'], st_d),
            ])

        tbl = Table(data, colWidths=cw, repeatRows=1)
        style = [
            ('BACKGROUND',    (0, 0), (-1, 0), AZUL),
            ('TOPPADDING',    (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('LEFTPADDING',   (0, 0), (-1, -1), 3),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 3),
            ('GRID',          (0, 0), (-1, -1), 0.4, CINZA_BORDA),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ]
        for i, m in enumerate(forn['extrato'], start=1):
            if m['_quinzena']:
                style += [
                    ('BACKGROUND', (0, i), (-1, i), AMARELO),
                    ('FONTNAME',   (0, i), (-1, i), 'Helvetica-Bold'),
                ]
            else:
                style.append(('BACKGROUND', (0, i), (-1, i),
                               CINZA_CLARO if i % 2 == 0 else BRANCO))
        tbl.setStyle(TableStyle(style))

        st_leg = ParagraphStyle('lg', fontSize=7, fontName='Helvetica',
                                textColor=colors.HexColor('#666666'))
        legenda = Paragraph(
            '<b>Legenda:</b> Linhas em amarelo = medi\u00e7\u00f5es com Prev. Pagto no per\u00edodo. '
            'Coluna Ref. = N\u00ba sequencial do contrato do empreiteiro (Nº/Emp).',
            st_leg
        )
        return [_titulo_secao(titulo), tbl, Spacer(1, 2*mm), legenda]

    # ------------------------------------------------------------------
    # Cabeçalhos e rodapé de página
    # ------------------------------------------------------------------

    def _draw_logo(self, cnv, x, y, w, h):
        if self.logo_path.exists():
            try:
                cnv.drawImage(str(self.logo_path), x, y, width=w, height=h,
                              preserveAspectRatio=True, mask='auto')
                return
            except Exception:
                pass
        cnv.setFont('Helvetica-Bold', 9)
        cnv.setFillColor(PRETO)
        cnv.drawString(x, y + h - 4*mm, self.EMPRESA['nome'])
        cnv.setFont('Helvetica', 7)
        cnv.drawString(x, y + h - 8*mm, self.EMPRESA['linha2'])

    def _cab_completo(self, cnv, doc, cliente_nome, cliente_end,
                       forn_nome, forn_cnpj, data_emissao):
        cnv.saveState()
        ml = doc.leftMargin
        mr = doc.leftMargin + doc.width

        # Logo — proporção 404x124, altura=14mm → largura≈46mm
        self._draw_logo(cnv, ml, 274*mm, 46*mm, 14*mm)

        # Dados empresa (direita, alinhados ao topo)
        cnv.setFont('Helvetica', 6.5)
        cnv.setFillColor(PRETO)
        cnv.drawRightString(mr, 286*mm, self.EMPRESA['endereco'])
        cnv.drawRightString(mr, 282*mm, self.EMPRESA['fones'])
        cnv.drawRightString(mr, 278*mm, self.EMPRESA['email'])

        # Data emissão — linha separada, negrito, alinhada à direita
        cnv.setFillColor(colors.HexColor('#444444'))
        cnv.setFont('Helvetica', 6.5)
        cnv.drawRightString(mr, 274*mm, 'Data:')
        cnv.setFont('Helvetica-Bold', 8)
        cnv.setFillColor(AZUL)
        cnv.drawRightString(mr, 270*mm, data_emissao)

        cx = (ml + mr) / 2

        # Título
        cnv.setFont('Helvetica-Bold', 12)
        cnv.setFillColor(AZUL)
        cnv.drawCentredString(cx, 262*mm, 'MEDI\u00c7\u00c3O DE SUB-EMPREITEIRO')

        # Nome do cliente — destaque
        cnv.setFont('Helvetica-Bold', 13)
        cnv.setFillColor(PRETO)
        cnv.drawCentredString(cx, 256*mm, cliente_nome.upper())

        # Endereço cliente
        cnv.setFont('Helvetica', 8)
        cnv.setFillColor(colors.HexColor('#444444'))
        end = cliente_end[:130] + ('...' if len(cliente_end) > 130 else '')
        cnv.drawCentredString(cx, 251*mm, end)

        # Linha separadora
        cnv.setStrokeColor(AZUL)
        cnv.setLineWidth(1.0)
        cnv.line(ml, 247*mm, mr, 247*mm)

        cnv.restoreState()

    def _cab_reduzido(self, cnv, doc, cliente_nome,
                       forn_nome, forn_cnpj, data_emissao):
        cnv.saveState()
        ml = doc.leftMargin
        mr = doc.leftMargin + doc.width

        self._draw_logo(cnv, ml, 278*mm, 25*mm, 8*mm)

        cnv.setFont('Helvetica', 6.5)
        cnv.setFillColor(PRETO)
        cnv.drawRightString(mr, 285*mm, self.EMPRESA['endereco'])
        cnv.drawRightString(mr, 281*mm, self.EMPRESA['fones'])
        cnv.drawRightString(mr, 277*mm, self.EMPRESA['email'])

        cnv.setFont('Helvetica-Bold', 9)
        cnv.setFillColor(AZUL)
        cnv.drawString(ml, 272*mm, cliente_nome.upper())

        cnv.setFont('Helvetica', 8)
        cnv.setFillColor(PRETO)
        cnv.drawString(ml, 268*mm, f'Contratada: {forn_nome}  |  {forn_cnpj}')

        cnv.setFont('Helvetica-Bold', 7.5)
        cnv.drawRightString(mr, 268*mm, data_emissao)

        cnv.setStrokeColor(AZUL)
        cnv.setLineWidth(0.6)
        cnv.line(ml, 265*mm, mr, 265*mm)

        cnv.restoreState()

    def _rodape(self, cnv, doc):
        cnv.saveState()
        ml = doc.leftMargin
        mr = doc.leftMargin + doc.width
        cnv.setFont('Helvetica', 7.5)
        cnv.setFillColor(colors.HexColor('#666666'))
        cnv.drawRightString(mr, 13*mm, f'P\u00e1gina {cnv.getPageNumber()}')
        cnv.setStrokeColor(CINZA_BORDA)
        cnv.setLineWidth(0.4)
        cnv.line(ml, 16*mm, mr, 16*mm)
        cnv.restoreState()

    # ------------------------------------------------------------------
    # Geração
    # ------------------------------------------------------------------

    def gerar_pdf(self, data_referencia, arquivo_saida=None, cnpj_fornecedor=None):
        """
        Gera o relatório quinzenal em um único PDF sem dependências externas de merge.

        Estratégia: todos os fornecedores são montados em uma única lista de elementos
        e um único SimpleDocTemplate. O callback on_page recebe o número da página e
        consulta um mapa página→fornecedor construído numa primeira passagem (two-pass):
        primeiro geramos para /dev/null contando páginas por fornecedor, depois
        geramos o arquivo final com cabeçalho correto em cada página.

        Args:
            data_referencia:  datetime
            arquivo_saida:    Path/str opcional
            cnpj_fornecedor:  CNPJ para filtrar (None = todos com medições na quinzena)
        """
        self._carregar_cliente()
        fornecedores = self._carregar_dados_planilha(data_referencia, cnpj_fornecedor)

        if not fornecedores:
            print('Nenhuma medi\u00e7\u00e3o encontrada na quinzena especificada.')
            return None

        if not arquivo_saida:
            data_str      = data_referencia.strftime('%d-%m-%Y')
            nome_cli      = self.cliente_info['nome'].replace(' ', '_').upper()
            arquivo_saida = self.arquivo_cliente.parent / f'REL_MED_{nome_cli}_{data_str}.pdf'
        arquivo_saida = Path(arquivo_saida)

        cliente_nome = self.cliente_info.get('nome', 'CLIENTE')
        cliente_end  = self.cliente_info.get('endereco', '')
        data_emissao = data_referencia.strftime('%d/%m/%Y')

        def _montar_elements(forn_list):
            """Monta a lista completa de elementos para todos os fornecedores."""
            els = []
            for idx, forn in enumerate(forn_list):
                els += self._bloco_contratada(forn)
                els.append(Spacer(1, 4*mm))
                els += self._bloco_contratos(forn)
                els.append(Spacer(1, 4*mm))
                els.append(self._bloco_totais(forn))
                els.append(PageBreak())
                els += self._bloco_extrato(forn)
                if idx < len(forn_list) - 1:
                    els.append(PageBreak())
            return els

        # --- PASS 1: contar páginas por fornecedor via dry-run em /dev/null ---
        # Usamos um canvas counter — cada fornecedor começa sempre na página 1
        # do seu bloco e ocupa N páginas. Determinamos N simulando o build.
        paginas_por_forn = self._contar_paginas_por_fornecedor(
            fornecedores, cliente_nome, cliente_end, data_emissao
        )

        # Montar mapa página_global → (fornecedor, é_primeira_pagina_do_fornecedor)
        mapa_pagina = {}  # página (1-based) → (forn_dict, bool primeira_pag)
        pagina_atual = 1
        for idx, forn in enumerate(fornecedores):
            n_pags = paginas_por_forn[idx]
            for p in range(n_pags):
                mapa_pagina[pagina_atual + p] = (forn, p == 0)
            pagina_atual += n_pags

        # --- PASS 2: gerar PDF final ---
        elements = _montar_elements(fornecedores)

        doc = SimpleDocTemplate(
            str(arquivo_saida),
            pagesize=A4,
            rightMargin=MARGIN_R, leftMargin=MARGIN_L,
            topMargin=MARGIN_T,   bottomMargin=MARGIN_B,
        )

        def on_page(cnv, doc_inner):
            pag = cnv.getPageNumber()
            forn_cur, primeira = mapa_pagina.get(pag, (fornecedores[-1], False))
            forn_nome = forn_cur['dados_base'].get('nome') or forn_cur['nome']
            forn_cnpj = forn_cur['dados_base'].get('cnpj') or forn_cur['cnpj']
            if primeira:
                self._cab_completo(cnv, doc_inner, cliente_nome, cliente_end,
                                   forn_nome, forn_cnpj, data_emissao)
            else:
                self._cab_reduzido(cnv, doc_inner, cliente_nome,
                                   forn_nome, forn_cnpj, data_emissao)
            self._rodape(cnv, doc_inner)

        doc.build(elements, onFirstPage=on_page, onLaterPages=on_page)

        print(f'Relat\u00f3rio gerado: {arquivo_saida}')
        return str(arquivo_saida)

    def _contar_paginas_por_fornecedor(self, fornecedores, cliente_nome,
                                        cliente_end, data_emissao):
        """
        Faz um dry-run do build para cada fornecedor individualmente e
        retorna uma lista com o número de páginas de cada um.
        Usa io.BytesIO como destino para não gravar em disco.
        """
        import io
        contagens = []
        for forn in fornecedores:
            forn_nome = forn['dados_base'].get('nome') or forn['nome']
            forn_cnpj = forn['dados_base'].get('cnpj') or forn['cnpj']

            els = []
            els += self._bloco_contratada(forn)
            els.append(Spacer(1, 4*mm))
            els += self._bloco_contratos(forn)
            els.append(Spacer(1, 4*mm))
            els.append(self._bloco_totais(forn))
            els.append(PageBreak())
            els += self._bloco_extrato(forn)

            buf = io.BytesIO()
            doc_tmp = SimpleDocTemplate(
                buf, pagesize=A4,
                rightMargin=MARGIN_R, leftMargin=MARGIN_L,
                topMargin=MARGIN_T,   bottomMargin=MARGIN_B,
            )

            contador = [0]

            def on_page_count(cnv, _doc):
                contador[0] = cnv.getPageNumber()

            doc_tmp.build(els, onFirstPage=on_page_count, onLaterPages=on_page_count)
            contagens.append(contador[0])

        return contagens


# ---------------------------------------------------------------------------
def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('arquivo_cliente')
    parser.add_argument('arquivo_clientes')
    parser.add_argument('--fornecedores', default=None)
    parser.add_argument('--data',         default=None)
    parser.add_argument('--cnpj',         default=None)
    parser.add_argument('--output',       default=None)
    args = parser.parse_args()

    data_ref = datetime.strptime(args.data, '%d/%m/%Y') if args.data else datetime.now()
    gerador  = RelatorioQuinzenalPDF(args.arquivo_cliente, args.arquivo_clientes,
                                     args.fornecedores)
    gerador.gerar_pdf(data_ref, args.output, args.cnpj)


if __name__ == '__main__':
    main()
