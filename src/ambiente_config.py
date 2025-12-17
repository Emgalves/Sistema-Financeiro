#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de Configuração de Ambiente - Sistema de Gestão Financeira
VERSÃO MELHORADA - Prioriza execução como script Python

REGRA PRINCIPAL:
    - Script Python (.py)     → SEMPRE usa ambiente TESTE
    - Executável compilado    → Detecta pelo nome do arquivo
        - Nome termina com _PRODUCAO ou _PROD → PRODUÇÃO
        - Nome termina com _TESTE ou _TEST    → TESTE
        - Fallback: busca Google Drive        → PRODUÇÃO se encontrar
        - Padrão seguro                       → TESTE

SEMPRE VERIFICAR ESTES ARQUIVOS PARA MANTER CONSISTÊNCIA:
 - src/ambiente_config.py
    - src/config/paths.py
    - src/config/config.py
    - src/config/__init__.py
"""

import os
import sys
from pathlib import Path
from dotenv import load_dotenv

# Carregar variáveis de ambiente (mas não usar como prioridade!)
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
        
        # ===================================================================
        # VERIFICAÇÃO PRIORITÁRIA: Script Python = SEMPRE TESTE
        # ===================================================================
        eh_script_python = not getattr(sys, 'frozen', False)
        
        if eh_script_python:
            print(f"🐍 EXECUTANDO COMO: SCRIPT PYTHON (.py)")
            print(f"   sys.argv[0]: {sys.argv[0]}")
            print(f"\n⚠️  REGRA ABSOLUTA: Scripts Python SEMPRE usam ambiente TESTE")
            print(f"   (Para usar PRODUÇÃO, compile como executável)")
            
            self.ambiente = self.TESTE
            print("="*70)
            print(f"🎯 AMBIENTE FINAL: {self.ambiente}")
            print("="*70 + "\n")
            return  # ✅ Sai aqui, não continua a detecção
        
        # ===================================================================
        # Se chegou aqui, é executável compilado (.exe)
        # ===================================================================
        print(f"📦 EXECUTANDO COMO: EXECUTÁVEL COMPILADO (.exe)")
        print(f"   sys.executable: {sys.executable}")
        
        # MÉTODO 1: Detectar pelo NOME DO EXECUTÁVEL (só para executáveis!)
        executavel = self._get_nome_executavel()
        print(f"1. Nome do executável: '{executavel}'")
        
        ambiente_detectado = False
        nome_upper = executavel.upper()
        print(f"2. Nome em maiúsculas: '{nome_upper}'")
        
        # DEBUG: Testar cada condição
        print(f"\n3. Testando condições:")
        print(f"   - Termina com '_PRODUCAO'? {nome_upper.endswith('_PRODUCAO')}")
        print(f"   - Termina com '_PROD'? {nome_upper.endswith('_PROD')}")
        print(f"   - Termina com '_TESTE'? {nome_upper.endswith('_TESTE')}")
        print(f"   - Termina com '_TEST'? {nome_upper.endswith('_TEST')}")
        print(f"   - Contém 'PRODUCAO'? {'PRODUCAO' in nome_upper}")
        print(f"   - Contém 'TESTE'? {'TESTE' in nome_upper}")
        
        # ===================================================================
        # PRIORIDADE MÁXIMA: NOME DO EXECUTÁVEL
        # ===================================================================
        
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
        
        # ===================================================================
        # MÉTODO 2: FALLBACK - BUSCAR GOOGLE DRIVE (NOVO!)
        # Só usar se nome não tiver sufixo identificável
        # ===================================================================
        if not ambiente_detectado:
            print(f"\n⚠️ Nome sem sufixo identificável")
            print(f"🔍 Buscando Google Drive como fallback...")
            
            google_drive_encontrado = self._buscar_google_drive()
            
            if google_drive_encontrado:
                self.ambiente = self.PRODUCAO
                ambiente_detectado = True
                print(f"\n✅ RESULTADO: Detectado como PRODUCAO (Google Drive encontrado)")
            else:
                print(f"   ❌ Google Drive não encontrado")
        
        # ===================================================================
        # MÉTODO 2: Padrão TESTE para executáveis sem identificação clara
        # ===================================================================
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
    
    def _buscar_google_drive(self):
        """
        Busca Google Drive em caminhos conhecidos
        Retorna True se encontrar, False caso contrário
        """
        import platform
        
        caminhos_windows = [
            Path("H:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios"),
            Path("G:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios"),
            Path("H:/Drives compartilhados/Relatórios"),
            Path("G:/Drives compartilhados/Relatórios"),
            Path("H:/Relatórios"),
            Path("G:/Relatórios"),
            Path("F:/Relatórios"),
            Path("E:/Relatórios"),
        ]
        
        caminhos_mac = [
            Path(os.path.expanduser("~")) / "Library/CloudStorage/GoogleDrive-emilia.mga@gmail.com/Meu Drive",
            Path(os.path.expanduser("~")) / "Google Drive",
        ]
        
        caminhos = caminhos_windows if platform.system() == 'Windows' else caminhos_mac
        
        for idx, caminho in enumerate(caminhos, 1):
            print(f"   [{idx}/{len(caminhos)}] Testando: {caminho}")
            if caminho.exists():
                print(f"   ✅ ENCONTRADO: {caminho}")
                return True
            else:
                print(f"   ❌ Não existe")
        
        return False
    
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
    
    banner = tk.Frame(parent, bg=config['cor_banner'], height=30)
    banner.pack(side='top', fill='x')
    banner.pack_propagate(False)
    
    texto = "⚠️  AMBIENTE DE TESTE - Dados não são reais  ⚠️"
    label = tk.Label(
        banner,
        text=texto,
        bg=config['cor_banner'],
        fg=config['cor_texto_banner'],
        font=('Helvetica', 10, 'bold')
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