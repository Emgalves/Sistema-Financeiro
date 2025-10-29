#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de Configuração de Ambiente - Sistema de Gestão Financeira
VERSÃO DEBUG - Com logs detalhados para diagnóstico
"""

import os
import sys
from pathlib import Path
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv()

class ConfiguracaoAmbiente:
    """Gerencia a configuração do ambiente (Teste/Produção)"""
    
    # Definição dos ambientes
    PRODUCAO = "PRODUCAO"
    TESTE = "TESTE"
    
    def __init__(self):
        print("\n" + "="*70)
        print("DEBUG - DETECÇÃO DE AMBIENTE")
        print("="*70)
        
        # MÉTODO 1: Detectar pelo NOME DO EXECUTÁVEL
        executavel = self._get_nome_executavel()
        print(f"1. Nome do executável: '{executavel}'")
        
        # Verificar como PyInstaller está executando
        if getattr(sys, 'frozen', False):
            print(f"2. Executando como: EXECUTÁVEL COMPILADO")
            print(f"   sys.executable: {sys.executable}")
        else:
            print(f"2. Executando como: SCRIPT PYTHON")
            print(f"   sys.argv[0]: {sys.argv[0]}")
        
        ambiente_detectado = False
        nome_upper = executavel.upper()
        print(f"3. Nome em maiúsculas: '{nome_upper}'")
        
        # DEBUG: Testar cada condição
        print(f"\n4. Testando condições:")
        print(f"   - Termina com '_PRODUCAO'? {nome_upper.endswith('_PRODUCAO')}")
        print(f"   - Termina com '_PROD'? {nome_upper.endswith('_PROD')}")
        print(f"   - Termina com '_TESTE'? {nome_upper.endswith('_TESTE')}")
        print(f"   - Termina com '_TEST'? {nome_upper.endswith('_TEST')}")
        print(f"   - Contém 'PRODUCAO'? {'PRODUCAO' in nome_upper}")
        print(f"   - Contém 'TESTE'? {'TESTE' in nome_upper}")
        
        # REGRA 1: Termina com _PRODUCAO ou _PROD
        if nome_upper.endswith("_PRODUCAO") or nome_upper.endswith("_PROD"):
            self.ambiente = self.PRODUCAO
            ambiente_detectado = True
            print(f"\n✅ RESULTADO: Detectado como PRODUCAO (termina com _PRODUCAO/_PROD)")
        
        # REGRA 2: Termina com _TESTE ou _TEST
        elif nome_upper.endswith("_TESTE") or nome_upper.endswith("_TEST"):
            self.ambiente = self.TESTE
            ambiente_detectado = True
            print(f"\n✅ RESULTADO: Detectado como TESTE (termina com _TESTE/_TEST)")
        
        # REGRA 3: Contém PRODUCAO mas NÃO contém TESTE
        elif "PRODUCAO" in nome_upper and "TESTE" not in nome_upper:
            self.ambiente = self.PRODUCAO
            ambiente_detectado = True
            print(f"\n✅ RESULTADO: Detectado como PRODUCAO (contém PRODUCAO, não contém TESTE)")
        
        # REGRA 4: Contém TESTE
        elif "TESTE" in nome_upper:
            self.ambiente = self.TESTE
            ambiente_detectado = True
            print(f"\n✅ RESULTADO: Detectado como TESTE (contém TESTE)")
        
        # MÉTODO 2: Tentar .env
        if not ambiente_detectado:
            print(f"\n⚠️ Não detectado pelo nome, tentando .env...")
            
            env_value = os.getenv("AMBIENTE_SISTEMA", "").upper()
            print(f"   AMBIENTE_SISTEMA = '{env_value}'")
            
            if not env_value:
                env_value = os.getenv("SISTEMA_AMBIENTE", "").upper()
                print(f"   SISTEMA_AMBIENTE = '{env_value}'")
            
            if env_value in [self.PRODUCAO, self.TESTE]:
                self.ambiente = env_value
                ambiente_detectado = True
                print(f"\n✅ RESULTADO: Detectado pelo .env como {self.ambiente}")
        
        # MÉTODO 3: Padrão TESTE
        if not ambiente_detectado:
            self.ambiente = self.TESTE
            print(f"\n⚠️ RESULTADO: Usando padrão TESTE (nenhuma regra aplicou)")
        
        print("="*70)
        print(f"🎯 AMBIENTE FINAL: {self.ambiente}")
        print("="*70 + "\n")
    
    def _get_nome_executavel(self):
        """Obtém o nome do executável/script em execução"""
        try:
            if getattr(sys, 'frozen', False):
                # Executável compilado com PyInstaller
                exe_path = sys.executable
            else:
                # Script Python normal
                exe_path = sys.argv[0]
            
            # Retornar apenas o nome do arquivo sem extensão
            nome = Path(exe_path).stem
            return nome
        except Exception as e:
            print(f"⚠️ Erro ao detectar nome: {e}")
            return ""
    
    def eh_producao(self):
        """Verifica se está em ambiente de produção"""
        return self.ambiente == self.PRODUCAO
    
    def eh_teste(self):
        """Verifica se está em ambiente de teste"""
        return self.ambiente == self.TESTE
    
    def get_nome_ambiente(self):
        """Retorna o nome do ambiente atual"""
        return self.ambiente
    
    def get_config_visual(self):
        """Retorna configurações visuais baseadas no ambiente"""
        if self.eh_producao():
            return {
                'cor_fundo': '#f0f0f0',
                'cor_banner': '#2e7d32',
                'cor_texto_banner': 'white',
                'cor_titulo': '#000000',
                'prefixo_titulo': '🟢 PRODUÇÃO',
                'mostrar_banner': False,
                'cor_card': 'white',
                'borda_janela': 'normal'
            }
        else:
            return {
                'cor_fundo': '#fff9e6',
                'cor_banner': '#ff6b00',
                'cor_texto_banner': 'white',
                'cor_titulo': '#d84315',
                'prefixo_titulo': '⚠️ AMBIENTE DE TESTE',
                'mostrar_banner': True,
                'cor_card': '#fffbf0',
                'borda_janela': '#ff6b00'
            }
    
    def get_titulo_janela(self, titulo_base):
        """Adiciona prefixo ao título baseado no ambiente"""
        config = self.get_config_visual()
        return f"{config['prefixo_titulo']} - {titulo_base}"
    
    def exibir_info_ambiente(self):
        """Exibe informações sobre o ambiente atual"""
        separador = "=" * 60
        print(separador)
        if self.eh_producao():
            print("🟢 MODO: PRODUÇÃO")
            print("   Todas as operações afetarão dados REAIS!")
        else:
            print("⚠️  MODO: TESTE")
            print("   Ambiente seguro para testes e experimentação")
        print(separador)


# Instância global
config_ambiente = ConfiguracaoAmbiente()


def aplicar_estilo_ambiente(widget, tipo='janela'):
    """Aplica estilos visuais baseados no ambiente"""
    config = config_ambiente.get_config_visual()
    
    try:
        if tipo == 'janela':
            widget.configure(bg=config['cor_fundo'])
        elif tipo == 'frame':
            widget.configure(bg=config['cor_fundo'])
        elif tipo == 'card':
            widget.configure(bg=config['cor_card'])
        elif tipo == 'label':
            widget.configure(bg=config['cor_fundo'])
    except Exception as e:
        print(f"Erro ao aplicar estilo: {e}")


def criar_banner_ambiente(parent):
    """Cria um banner visual no topo da janela"""
    import tkinter as tk
    
    config = config_ambiente.get_config_visual()
    
    if not config['mostrar_banner']:
        return None
    
    banner = tk.Frame(parent, bg=config['cor_banner'], height=50)
    banner.pack(side='top', fill='x')
    banner.pack_propagate(False)
    
    texto = "⚠️  AMBIENTE DE TESTE - Dados não são reais  ⚠️"
    label = tk.Label(
        banner,
        text=texto,
        bg=config['cor_banner'],
        fg=config['cor_texto_banner'],
        font=('Helvetica', 12, 'bold')
    )
    label.pack(expand=True)
    
    return banner


def configurar_ttk_style_ambiente():
    """Configura os estilos TTK baseado no ambiente"""
    from tkinter import ttk
    
    config = config_ambiente.get_config_visual()
    style = ttk.Style()
    
    style.configure('Menu.TFrame', background=config['cor_fundo'])
    style.configure('Card.TFrame', background=config['cor_card'])
    style.configure('CardTitle.TLabel', 
                   font=('Helvetica', 14, 'bold'),
                   background=config['cor_card'])
    style.configure('CardDesc.TLabel',
                   font=('Helvetica', 10),
                   background=config['cor_card'],
                   wraplength=300)
    style.configure('Action.TButton',
                   font=('Helvetica', 12),
                   padding=10)
    
    if config_ambiente.eh_teste():
        style.configure('Action.TButton',
                       font=('Helvetica', 12, 'bold'),
                       padding=10)


def get_cor_status():
    """Retorna a cor de status baseada no ambiente"""
    return '#2e7d32' if config_ambiente.eh_producao() else '#ff6b00'


# Exibir informações ao carregar o módulo
config_ambiente.exibir_info_ambiente()