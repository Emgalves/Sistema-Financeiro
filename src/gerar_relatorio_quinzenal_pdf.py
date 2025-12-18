"""
Script para geração de Relatório Quinzenal de Medições em PDF
Autor: Sistema de Gestão de Contratos
Data: 2025
"""

import os
import sys
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, 
    Spacer, PageBreak, Image
)
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.pdfgen import canvas


class RelatorioQuinzenalPDF:
    """Classe para geração de relatório quinzenal de medições em PDF"""
    
    def __init__(self, arquivo_cliente, arquivo_clientes):
        """
        Inicializa o gerador de relatório
        
        Args:
            arquivo_cliente: Caminho para o arquivo Excel do cliente
            arquivo_clientes: Caminho para o arquivo Excel com dados dos clientes
        """
        self.arquivo_cliente = Path(arquivo_cliente)
        self.arquivo_clientes = Path(arquivo_clientes)
        self.cliente_info = None
        self.contratos_quinzena = []
        
    def carregar_dados_cliente(self):
        """Carrega informações do cliente a partir do arquivo clientes.xlsx"""
        try:
            # Extrair CPF do nome do arquivo
            nome_arquivo = self.arquivo_cliente.stem
            
            # Carregar planilha de clientes
            df_clientes = pd.read_excel(self.arquivo_clientes, sheet_name='Clientes')
            
            # Procurar cliente pelo nome (simplificado - pode precisar ajuste)
            # O nome do arquivo é o nome do cliente
            cliente_nome = nome_arquivo.replace('_', ' ')
            
            # Tentar encontrar cliente
            for _, row in df_clientes.iterrows():
                nome_cliente_base = str(row['Nome']).strip().upper()
                if cliente_nome.upper() in nome_cliente_base:
                    self.cliente_info = {
                        'nome': row['Nome'],
                        'endereco': str(row.get('Endereço', '')) if pd.notna(row.get('Endereço')) else ''
                    }
                    break
            
            if not self.cliente_info:
                # Se não encontrou, usar dados do arquivo
                self.cliente_info = {
                    'nome': cliente_nome,
                    'endereco': 'NÃO INFORMADO'
                }
                
        except Exception as e:
            print(f"Erro ao carregar dados do cliente: {str(e)}")
            import traceback
            traceback.print_exc()
            self.cliente_info = {
                'nome': 'CLIENTE',
                'endereco': 'NÃO INFORMADO'
            }
    
    def identificar_quinzena(self, data_referencia):
        """
        Identifica a quinzena baseada na data de referência
        
        Args:
            data_referencia: datetime object
            
        Returns:
            tuple: (data_inicio, data_fim) da quinzena
        """
        dia = data_referencia.day
        mes = data_referencia.month
        ano = data_referencia.year
        
        if dia <= 5:
            # Primeira quinzena - do dia 21 do mês anterior até dia 5
            data_fim = datetime(ano, mes, 5)
            if mes == 1:
                data_inicio = datetime(ano - 1, 12, 21)
            else:
                data_inicio = datetime(ano, mes - 1, 21)
        else:
            # Segunda quinzena - do dia 6 até dia 20
            data_inicio = datetime(ano, mes, 6)
            data_fim = datetime(ano, mes, 20)
        
        return data_inicio, data_fim
    
    def filtrar_medicoes_quinzena(self, data_referencia):
        """
        Filtra medições que foram realizadas na quinzena
        
        Args:
            data_referencia: datetime object da data de referência
            
        Returns:
            list: Lista de contratos com medições na quinzena
        """
        data_inicio, data_fim = self.identificar_quinzena(data_referencia)
        
        try:
            # Carregar aba de contratos
            df_contratos = pd.read_excel(self.arquivo_cliente, sheet_name='Contratos_Medicao')
            
            # Carregar aba de medições
            df_medicoes = pd.read_excel(self.arquivo_cliente, sheet_name='Medicoes')
            
            # === FILTRO CRÍTICO: Excluir medições com status EXCLUÍDO ===
            # Medições excluídas não devem aparecer em nenhum relatório
            total_medicoes_antes = len(df_medicoes)
            df_medicoes = df_medicoes[
                (df_medicoes['Status'].isna()) | 
                (df_medicoes['Status'] != 'EXCLUÍDO')
            ].copy()
            total_medicoes_depois = len(df_medicoes)
            medicoes_excluidas = total_medicoes_antes - total_medicoes_depois
            
            if medicoes_excluidas > 0:
                print(f"✓ {medicoes_excluidas} medição(ões) com status EXCLUÍDO foram filtradas do relatório")
            
            # Converter datas
            df_medicoes['Data_Medicao'] = pd.to_datetime(df_medicoes['Data_Medicao'], errors='coerce')
            df_medicoes['Data_Pagamento'] = pd.to_datetime(df_medicoes['Data_Pagamento'], errors='coerce')
            
            # Filtrar medições na quinzena
            medicoes_quinzena = df_medicoes[
                (df_medicoes['Data_Medicao'] >= data_inicio) & 
                (df_medicoes['Data_Medicao'] <= data_fim)
            ].copy()
            
            if medicoes_quinzena.empty:
                print(f"Nenhuma medição encontrada entre {data_inicio.strftime('%d/%m/%Y')} e {data_fim.strftime('%d/%m/%Y')}")
                return []
            
            # Agrupar por contrato
            contratos_com_medicoes = []
            
            for id_contrato in medicoes_quinzena['ID_Contrato'].unique():
                # Buscar informações do contrato
                contrato_rows = df_contratos[df_contratos['ID_Contrato'] == id_contrato]
                if contrato_rows.empty:
                    print(f"Contrato ID {id_contrato} não encontrado na aba Contratos_Medicao")
                    continue
                    
                contrato = contrato_rows.iloc[0]
                
                # Buscar TODAS as medições do contrato (histórico completo)
                todas_medicoes = df_medicoes[
                    df_medicoes['ID_Contrato'] == id_contrato
                ].sort_values('Data_Medicao')
                
                # Identificar quais medições são da quinzena (para destaque visual)
                medicoes_quinzena_ids = medicoes_quinzena[
                    medicoes_quinzena['ID_Contrato'] == id_contrato
                ]['ID_Medicao'].tolist()
                
                contratos_com_medicoes.append({
                    'id': int(id_contrato),
                    'fornecedor': contrato['Nome_Fornecedor'],
                    'descricao': contrato['Descricao'],
                    'valor_global': float(contrato.get('Valor_Global', 0)),
                    'data_inicio': contrato.get('Data_Inicio'),
                    'data_final': contrato.get('Data_Final'),
                    'status': contrato.get('Status', 'ATIVO'),
                    'todas_medicoes': todas_medicoes.to_dict('records'),
                    'medicoes_quinzena_ids': medicoes_quinzena_ids,  # IDs para destacar
                    'valor_executado_total': float(todas_medicoes['Valor'].sum()),
                    'qtd_medicoes_quinzena': len(medicoes_quinzena_ids)
                })
            
            # === ORDENAR CONTRATOS POR ID ===
            contratos_com_medicoes.sort(key=lambda x: x['id'])
            
            self.contratos_quinzena = contratos_com_medicoes
            return contratos_com_medicoes
            
        except Exception as e:
            print(f"Erro ao filtrar medições: {str(e)}")
            import traceback
            traceback.print_exc()
            return []
    
    def formatar_moeda(self, valor):
        """Formata valor como moeda brasileira"""
        try:
            return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        except:
            return "R$ 0,00"
    
    def formatar_data(self, data):
        """Formata data para padrão brasileiro"""
        if pd.isna(data):
            return ""
        if isinstance(data, str):
            try:
                data = pd.to_datetime(data)
            except:
                return data
        try:
            return data.strftime('%d/%m/%Y')
        except:
            return str(data)
    
    def criar_cabecalho(self, canvas, doc):
        """
        Cria cabeçalho do documento com dados da empresa e do cliente
        Este cabeçalho aparece em TODAS as páginas do relatório
        IMPORTANTE: CPF não é exibido por ser dado sensível (LGPD)
        """
        canvas.saveState()
        
        # === MARGENS ALINHADAS COM O CORPO DO DOCUMENTO ===
        margin_left = doc.leftMargin
        margin_right = doc.width + doc.leftMargin
        
        # === SEÇÃO EMPRESA (Vasconcelos&Rinaldi) ===
        # Logo - verificar se existe
        logo_path = Path(__file__).parent / 'logo.png'
        if logo_path.exists():
            try:
                canvas.drawImage(
                    str(logo_path), 
                    margin_left, 
                    273*mm, 
                    width=40*mm, 
                    height=10*mm,
                    preserveAspectRatio=True,
                    mask='auto'
                )
            except:
                # Se falhar ao carregar logo, mostrar texto
                canvas.setFont('Helvetica-Bold', 9)
                canvas.drawString(margin_left, 280*mm, "VASCONCELOS&RINALDI")
                canvas.setFont('Helvetica', 7)
                canvas.drawString(margin_left, 277*mm, "ENGENHARIA")
        else:
            # Sem logo, mostrar texto
            canvas.setFont('Helvetica-Bold', 9)
            canvas.drawString(margin_left, 280*mm, "VASCONCELOS&RINALDI")
            canvas.setFont('Helvetica', 7)
            canvas.drawString(margin_left, 277*mm, "ENGENHARIA")
        
        # Informações de contato da empresa (lado direito superior)
        canvas.setFont('Helvetica', 6)
        y_pos = 280*mm
        canvas.drawRightString(margin_right, y_pos, "Rua Zodiaco, 87 Sala 07 – Santa Lúcia - Belo Horizonte - MG")
        y_pos -= 3*mm
        canvas.drawRightString(margin_right, y_pos, "(31) 3654-6616 / (31) 99974-1241 / (31) 98711-1139")
        y_pos -= 3*mm
        canvas.drawRightString(margin_right, y_pos, "rvr.engenharia@gmail.com")
        
        # === ESPAÇAMENTO MAIOR ANTES DA LINHA SEPARADORA ===
        # Linha separadora mais abaixo para criar espaçamento visual
        # canvas.setStrokeColor(colors.HexColor('#1F4788'))
        # canvas.setLineWidth(0.5)
        # canvas.line(margin_left, 265*mm, margin_right, 265*mm)  # Desceu de 271mm para 269mm
        
        # === SEÇÃO CLIENTE (Dados do cliente em destaque - SEM CPF) ===
        # Área sombreada para dados do cliente - com margens alinhadas
        # canvas.setFillColor(colors.HexColor('#F5F5F5'))
        # canvas.rect(margin_left, 264*mm, doc.width, 8*mm, fill=1, stroke=0)
        
        # Nome do cliente (DESTAQUE) - com margem interna
        canvas.setFillColor(colors.HexColor('#1F4788'))
        canvas.setFont('Helvetica-Bold', 11)
        canvas.drawString(margin_left + 2*mm, 264*mm, self.cliente_info['nome'].upper())
        
        # Endereço do cliente (abaixo do nome) - com margem interna
        canvas.setFillColor(colors.black)
        canvas.setFont('Helvetica', 8)
        endereco_completo = self.cliente_info['endereco']
        # if self.cliente_info['cidade']:
        #     endereco_completo += f", {self.cliente_info['cidade']}"
        # if self.cliente_info['estado']:
        #     endereco_completo += f" / {self.cliente_info['estado']}"
        
        # Calcular largura disponível para o endereço
        largura_disponivel = doc.width - 4*mm  # Desconta margens internas
        
        # Truncar endereço se muito longo
        max_chars = 110
        if len(endereco_completo) > max_chars:
            endereco_completo = endereco_completo[:max_chars-3] + "..."
        
        canvas.drawString(margin_left + 2*mm, 260*mm, endereco_completo)
        
        canvas.restoreState()
    
    def criar_rodape(self, canvas, doc):
        """Cria rodapé do documento com numeração de página"""
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        page_num = canvas.getPageNumber()
        text = f"Página {page_num}"
        canvas.drawRightString(180*mm, 15*mm, text)
        canvas.restoreState()
    
    def gerar_pdf(self, data_referencia, arquivo_saida=None):
        """
        Gera o PDF do relatório quinzenal
        
        Args:
            data_referencia: datetime object da data de referência
            arquivo_saida: Caminho do arquivo de saída (opcional)
        """
        # Carregar dados do cliente
        self.carregar_dados_cliente()
        
        # Filtrar medições da quinzena
        contratos = self.filtrar_medicoes_quinzena(data_referencia)
        
        if not contratos:
            print("Nenhum contrato com medições na quinzena especificada.")
            return None
        
        # Definir nome do arquivo de saída
        if not arquivo_saida:
            data_str = data_referencia.strftime('%d-%m-%Y')
            nome_cliente_limpo = self.cliente_info['nome'].replace(' ', '_').upper()
            arquivo_saida = f"REL_-_{nome_cliente_limpo}_-_{data_str}.pdf"
        
        arquivo_saida = Path(arquivo_saida)
        
        # Criar documento PDF
        doc = SimpleDocTemplate(
            str(arquivo_saida),
            pagesize=A4,
            rightMargin=20*mm,
            leftMargin=20*mm,
            topMargin=42*mm,  # Ajustado para 42mm (cabeçalho menor sem CPF)
            bottomMargin=25*mm
        )
        
        # Configurar estilos
        styles = getSampleStyleSheet()
        
        # Estilo para título
        style_titulo = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            textColor=colors.HexColor('#1F4788'),
            spaceAfter=6,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        )
        
        # Estilo para subtítulo
        style_subtitulo = ParagraphStyle(
            'CustomSubtitle',
            parent=styles['Normal'],
            fontSize=10,
            spaceAfter=12,
            alignment=TA_CENTER,
            fontName='Helvetica'
        )
        
        # Estilo para seção
        style_secao = ParagraphStyle(
            'SectionTitle',
            parent=styles['Heading2'],
            fontSize=11,
            textColor=colors.HexColor('#1F4788'),
            spaceAfter=8,
            spaceBefore=12,
            fontName='Helvetica-Bold',
            backColor=colors.HexColor('#E8EAF6'),
            borderPadding=5
        )
        
        # Container para elementos do PDF
        elements = []
        
        # === NOTA: Título e endereço do cliente agora estão no cabeçalho ===
        # Isso garante que apareçam em TODAS as páginas do relatório
        
        # Seção de contratos e medições - com data no título
        data_inicio, data_fim = self.identificar_quinzena(data_referencia)
        
        secao_titulo = Paragraph(
            f"MEDIÇÕES DA QUINZENA - Data: {data_referencia.strftime('%d/%m/%Y')}", 
            style_secao
        )
        elements.append(secao_titulo)
        elements.append(Spacer(1, 5*mm))
        
        # Processar cada contrato
        for idx, contrato in enumerate(contratos):
            # Título do contrato
            contrato_titulo = Paragraph(
                f"<b>Contrato #{contrato['id']} - {contrato['fornecedor']}</b>",
                styles['Heading3']
            )
            elements.append(contrato_titulo)
            elements.append(Spacer(1, 3*mm))
            
            # Informações do contrato em formato de tabela otimizado
            # Calcular valores do resumo
            percentual_exec = (contrato['valor_executado_total'] / contrato['valor_global'] * 100) if contrato['valor_global'] > 0 else 0
            saldo = contrato['valor_global'] - contrato['valor_executado_total']
            
            # Criar estilo para descrição com quebra de linha
            style_descricao = ParagraphStyle(
                'Descricao',
                parent=styles['Normal'],
                fontSize=9,
                leading=11,
                fontName='Helvetica'
            )
            
            # Usar Paragraph para descrição permitir quebra automática
            descricao_paragraph = Paragraph(contrato['descricao'], style_descricao)
            
            # Criar tabela com descrição em linha separada
            # Linha 1: Descrição (ocupa toda a largura)
            # Linhas 2-4: Informações lado a lado (Label|Valor | Label|Valor)
            info_contrato = [
                ['Descrição:', descricao_paragraph, '', ''],
                ['Valor Global:', self.formatar_moeda(contrato['valor_global']), 'Executado Total:', self.formatar_moeda(contrato['valor_executado_total'])],
                ['Status:', contrato['status'], 'Saldo:', self.formatar_moeda(saldo)],
                ['Período:', f"{self.formatar_data(contrato['data_inicio'])} a {self.formatar_data(contrato['data_final'])}", '% Executado:', f"{percentual_exec:.1f}%"],
            ]
            
            table_contrato = Table(info_contrato, colWidths=[30*mm, 65*mm, 35*mm, 50*mm])
            table_contrato.setStyle(TableStyle([
                # Primeira linha - Descrição (merge das colunas 1-3 para dar mais espaço)
                ('SPAN', (1, 0), (3, 0)),  # Mesclar colunas 1-3 na primeira linha
                ('BACKGROUND', (1, 0), (3, 0), colors.HexColor('#FFFACD')),  # Destaque amarelo claro
                
                # Labels (colunas 0 e 2)
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('FONTNAME', (2, 0), (2, -1), 'Helvetica-Bold'),
                ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#555555')),
                ('TEXTCOLOR', (2, 0), (2, -1), colors.HexColor('#1F4788')),
                
                # Valores (colunas 1 e 3)
                ('FONTNAME', (1, 1), (1, -1), 'Helvetica'),  # Linha 1 usa Paragraph
                ('FONTNAME', (3, 0), (3, -1), 'Helvetica-Bold'),
                ('TEXTCOLOR', (3, 0), (3, -1), colors.HexColor('#1F4788')),
                
                # Alinhamentos
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (1, -1), 'LEFT'),
                ('ALIGN', (2, 0), (2, -1), 'LEFT'),
                ('ALIGN', (3, 0), (3, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                
                # Tamanho da fonte
                ('FONTSIZE', (0, 1), (-1, -1), 9),  # Linhas 2-4
                
                # Bordas leves
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#DDDDDD')),
            ]))
            
            elements.append(table_contrato)
            elements.append(Spacer(1, 5*mm))
            
            # Tabela de medições - HISTÓRICO COMPLETO DO CONTRATO
            medicoes_data = [['ID', 'Data Medição', 'Dt. Pagto', 'Referência', 'Valor']]
            
            for medicao in contrato['todas_medicoes']:
                # Adicionar quebra de linha na referência se for muito longa
                referencia = str(medicao['Referencia'])
                if len(referencia) > 45:
                    # Quebrar em aproximadamente 45 caracteres no último espaço
                    pos_quebra = referencia[:45].rfind(' ')
                    if pos_quebra > 0:
                        referencia = referencia[:pos_quebra] + '\n' + referencia[pos_quebra+1:]
                
                medicoes_data.append([
                    str(medicao['ID_Medicao']),
                    self.formatar_data(medicao['Data_Medicao']),
                    self.formatar_data(medicao['Data_Pagamento']),
                    referencia,
                    self.formatar_moeda(medicao['Valor'])
                ])
            
            col_widths = [10*mm, 25*mm, 25*mm, 90*mm, 30*mm]
            
            table_medicoes = Table(medicoes_data, colWidths=col_widths)
            
            # Estilo base da tabela
            table_style = [
                # Cabeçalho
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4788')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 8),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                
                # Corpo da tabela
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # ID
                ('ALIGN', (1, 1), (2, -1), 'CENTER'),  # Datas
                ('ALIGN', (3, 1), (3, -1), 'LEFT'),    # Referência
                ('ALIGN', (4, 1), (4, -1), 'RIGHT'),   # Valor
                
                # Bordas
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]
            
            # Destacar medições da quinzena com fundo amarelo claro
            for med_idx, medicao in enumerate(contrato['todas_medicoes'], start=1):
                if medicao['ID_Medicao'] in contrato['medicoes_quinzena_ids']:
                    # Linha da medição que está na quinzena - destaque amarelo
                    table_style.append(('BACKGROUND', (0, med_idx), (-1, med_idx), colors.HexColor('#FFFACD')))
                    table_style.append(('FONTNAME', (0, med_idx), (-1, med_idx), 'Helvetica-Bold'))
                else:
                    # Linhas alternadas para histórico
                    if med_idx % 2 == 0:
                        table_style.append(('BACKGROUND', (0, med_idx), (-1, med_idx), colors.HexColor('#F5F5F5')))
                    else:
                        table_style.append(('BACKGROUND', (0, med_idx), (-1, med_idx), colors.white))
            
            table_medicoes.setStyle(TableStyle(table_style))
            
            elements.append(table_medicoes)
            
            # Adicionar legenda se houver medições destacadas
            # if contrato['medicoes_quinzena_ids']:
            #     elements.append(Spacer(1, 2*mm))
            #     legenda_style = ParagraphStyle(
            #         'Legenda',
            #         parent=styles['Normal'],
            #         fontSize=7,
            #         textColor=colors.HexColor('#666666'),
            #         italic=True
            #     )
            #     legenda = Paragraph(
            #         f"<b>Legenda:</b> As linhas em <font backColor='#FFFACD'>amarelo</font> "
            #         f"indicam medições realizadas na quinzena atual ({contrato['qtd_medicoes_quinzena']} medição{'ões' if contrato['qtd_medicoes_quinzena'] > 1 else ''}). "
            #         f"As demais linhas mostram o histórico completo do contrato.",
            #         legenda_style
            #     )
            #     elements.append(legenda)
            
            # Espaço entre contratos e quebra de página
            if idx < len(contratos) - 1:
                # Sempre adicionar quebra de página entre contratos
                elements.append(PageBreak())
        
        # Construir PDF com cabeçalho e rodapé em todas as páginas
        doc.build(
            elements, 
            onFirstPage=lambda c, d: (self.criar_cabecalho(c, d), self.criar_rodape(c, d)),
            onLaterPages=lambda c, d: (self.criar_cabecalho(c, d), self.criar_rodape(c, d))
        )
        
        print(f"\nRelatório gerado com sucesso: {arquivo_saida}")
        print(f"Total de contratos no relatório: {len(contratos)}")
        
        return str(arquivo_saida)


def main():
    """Função principal para testes"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Gerador de Relatório Quinzenal de Medições')
    parser.add_argument('arquivo_cliente', help='Caminho do arquivo Excel do cliente')
    parser.add_argument('arquivo_clientes', help='Caminho do arquivo clientes.xlsx')
    parser.add_argument('--data', help='Data de referência (DD/MM/YYYY)', default=None)
    parser.add_argument('--output', help='Arquivo de saída', default=None)
    
    args = parser.parse_args()
    
    # Processar data
    if args.data:
        data_ref = datetime.strptime(args.data, '%d/%m/%Y')
    else:
        data_ref = datetime.now()
    
    # Gerar relatório
    gerador = RelatorioQuinzenalPDF(args.arquivo_cliente, args.arquivo_clientes)
    gerador.gerar_pdf(data_ref, args.output)


if __name__ == '__main__':
    main()
