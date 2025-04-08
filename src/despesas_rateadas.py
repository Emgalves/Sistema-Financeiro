# Imports da biblioteca padrão Python
import os
import sys
from pathlib import Path
import re
from datetime import datetime
from decimal import Decimal

# Imports relacionados ao Tkinter
import tkinter as tk
from tkinter import ttk, messagebox, StringVar
from tkinter import *
from tkcalendar import DateEntry, Calendar

# Imports para manipulação de dados e Excel
import pandas as pd
import xlwings as xw
from openpyxl import load_workbook
import openpyxl
import babel
from dateutil.relativedelta import relativedelta

# Imports para validação
from validate_docbr import CPF, CNPJ

def add_project_root():
    import sys
    from pathlib import Path
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent
    if str(project_root) not in sys.path:
        sys.path.append(str(project_root))

add_project_root()

# Importar logger
try:
    from config.logger_config import system_logger, log_action
    logger = system_logger.get_logger()
    logger.info("Logger importado com sucesso")
except Exception as e:
    print(f"Erro ao importar logger: {str(e)}")

class GerenciadorDespesasRateadas:

    def carregar_clientes_ativos(self):
        """Carrega todos os clientes ativos do sistema"""
        clientes = []
        try:
            wb = load_workbook(ARQUIVO_CLIENTES)
            ws = wb['Clientes']
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:  # Nome não vazio
                    clientes.append({
                        'nome': row[0],
                        'percentual': 0,
                        'valor': 0,
                        'arquivo': PASTA_CLIENTES / f"{row[0]}.xlsx"
                    })
            
            return clientes
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar clientes: {str(e)}")
            return []
        
    def calcular_rateio(self):
        """Calcula o rateio baseado nos percentuais ou valores definidos"""
        if self.modo_rateio.get() == "percentual":
            # Verificar se o total é 100%
            total_percentual = sum(cliente['percentual'] for cliente in self.clientes)
            if not (99.9 <= total_percentual <= 100.1):  # Tolerância para arredondamentos
                messagebox.showerror("Erro", f"O total de percentuais deve ser 100%. Atual: {total_percentual}%")
                return False
                
            # Calcular valores baseados nos percentuais
            valor_total = float(self.valor_total.get().replace(',', '.'))
            for cliente in self.clientes:
                cliente['valor'] = (cliente['percentual'] / 100) * valor_total
        else:  # modo = valor
            # Verificar se o total corresponde ao valor da despesa
            total_valores = sum(cliente['valor'] for cliente in self.clientes)
            valor_total = float(self.valor_total.get().replace(',', '.'))
            
            if abs(total_valores - valor_total) > 0.01:  # Tolerância de 1 centavo
                messagebox.showerror("Erro", 
                                    f"O total dos valores ({total_valores:.2f}) não corresponde ao valor da despesa ({valor_total:.2f})")
                return False
                
        return True

    def aplicar_rateio(self):
        """Aplica o rateio nos arquivos de cada cliente"""
        data_ref = self.data_ref.get_date()
        descricao = self.descricao.get()
        tipo_despesa = self.tipo_despesa.get()
        observacao = self.observacao.get()
        
        # Lista para registrar resultados
        registros = []
        
        for cliente in self.clientes:
            if cliente['valor'] <= 0:
                continue  # Pular clientes sem valor
                
            try:
                wb = load_workbook(cliente['arquivo'])
                ws = wb["Dados"]
                
                # Preparar dados do lançamento
                proxima_linha = ws.max_row + 1
                
                # Data do Relatório (formatada)
                ws.cell(row=proxima_linha, column=1, value=data_ref)
                ws.cell(row=proxima_linha, column=1).number_format = 'DD/MM/YYYY'
                
                # Tipo de Despesa
                ws.cell(row=proxima_linha, column=2, value=int(tipo_despesa))
                
                # CNPJ/CPF do sistema (se aplicável)
                ws.cell(row=proxima_linha, column=3, value="")
                
                # Nome do sistema
                ws.cell(row=proxima_linha, column=4, value="SISTEMA")
                
                # Referência
                ws.cell(row=proxima_linha, column=5, value=f"RATEIO: {descricao}")
                
                # NF (vazio para rateios)
                ws.cell(row=proxima_linha, column=6, value="")
                
                # Valor Unitário
                ws.cell(row=proxima_linha, column=7, value=cliente['valor'])
                ws.cell(row=proxima_linha, column=7).number_format = '#,##0.00'
                
                # Dias (1 para despesas rateadas)
                ws.cell(row=proxima_linha, column=8, value=1)
                
                # Valor Total
                ws.cell(row=proxima_linha, column=9, value=cliente['valor'])
                ws.cell(row=proxima_linha, column=9).number_format = '#,##0.00'
                
                # Data de Vencimento (mesma data do relatório por padrão)
                ws.cell(row=proxima_linha, column=10, value=data_ref)
                ws.cell(row=proxima_linha, column=10).number_format = 'DD/MM/YYYY'
                
                # Categoria
                ws.cell(row=proxima_linha, column=11, value="RATEIO")
                
                # Dados Bancários (vazio para rateios)
                ws.cell(row=proxima_linha, column=12, value="")
                
                # Observação
                ws.cell(row=proxima_linha, column=13, value=observacao)
                
                # Salvar planilha
                wb.save(cliente['arquivo'])
                
                # Registrar sucesso
                registros.append({
                    'cliente': cliente['nome'],
                    'valor': cliente['valor'],
                    'status': 'SUCESSO'
                })
                
            except Exception as e:
                # Registrar falha
                registros.append({
                    'cliente': cliente['nome'],
                    'valor': cliente['valor'],
                    'status': f'FALHA: {str(e)}'
                })
        
        # Registrar o rateio no histórico
        self.registrar_historico(registros)
        
        # Exibir resultados
        self.mostrar_resultado_rateio(registros)

    def registrar_historico(self, registros):
        """Registra o rateio no histórico"""
        try:
            data_atual = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
            data_ref = self.data_ref.get()
            descricao = self.descricao.get()
            valor_total = float(self.valor_total.get().replace(',', '.'))
            tipo_despesa = self.tipo_despesa.get()
            
            # Criar arquivo de histórico se não existir
            historico_path = Path('historico_rateios.xlsx')
            if not historico_path.exists():
                wb = Workbook()
                ws = wb.active
                ws.title = "Histórico"
                
                # Cabeçalhos
                headers = ['Data Registro', 'Data Referência', 'Descrição', 'Valor Total', 
                        'Tipo Despesa', 'Qtd Clientes', 'Status']
                for col, header in enumerate(headers, 1):
                    ws.cell(row=1, column=col, value=header)
                    
                wb.save(historico_path)
            
            # Abrir arquivo de histórico
            wb = load_workbook(historico_path)
            ws = wb["Histórico"]
            
            # Adicionar registro principal
            proxima_linha = ws.max_row + 1
            ws.cell(row=proxima_linha, column=1, value=data_atual)
            ws.cell(row=proxima_linha, column=2, value=data_ref)
            ws.cell(row=proxima_linha, column=3, value=descricao)
            ws.cell(row=proxima_linha, column=4, value=valor_total)
            ws.cell(row=proxima_linha, column=5, value=tipo_despesa)
            ws.cell(row=proxima_linha, column=6, value=len(registros))
            
            # Verificar status geral
            falhas = [r for r in registros if r['status'].startswith('FALHA')]
            status = "SUCESSO" if not falhas else f"PARCIAL ({len(falhas)} falhas)"
            ws.cell(row=proxima_linha, column=7, value=status)
            
            # Adicionar detalhes em outra aba se não existir
            if "Detalhes" not in wb.sheetnames:
                ws_details = wb.create_sheet("Detalhes")
                # Cabeçalhos
                headers = ['ID Rateio', 'Cliente', 'Valor', 'Status']
                for col, header in enumerate(headers, 1):
                    ws_details.cell(row=1, column=col, value=header)
            else:
                ws_details = wb["Detalhes"]
            
            # Adicionar detalhes de cada cliente
            id_rateio = proxima_linha - 1  # Usar a linha como ID do rateio
            for registro in registros:
                proxima_linha_det = ws_details.max_row + 1
                ws_details.cell(row=proxima_linha_det, column=1, value=id_rateio)
                ws_details.cell(row=proxima_linha_det, column=2, value=registro['cliente'])
                ws_details.cell(row=proxima_linha_det, column=3, value=registro['valor'])
                ws_details.cell(row=proxima_linha_det, column=4, value=registro['status'])
            
            # Salvar histórico
            wb.save(historico_path)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao registrar histórico: {str(e)}")