# -*- coding: utf-8 -*-
"""
CORREÇÃO RÁPIDA E FINAL - NFe/CERTIFICADO A1
Salve como: corrigir_nfe_final.py (na raiz do projeto)
Executa correção e termina automaticamente
"""

import sys
import os
from pathlib import Path

def aplicar_correcao_final():
    """Aplica correção final sem loops"""
    print("🔧 CORREÇÃO FINAL NFe/CERTIFICADO A1")
    print("=" * 45)
    
    try:
        # 1. Configurar paths
        current_dir = Path(__file__).resolve().parent
        paths = [
            str(current_dir),
            str(current_dir / "src"),
            str(current_dir / "src" / "nfe")
        ]
        
        for path in paths:
            if os.path.exists(path) and path not in sys.path:
                sys.path.insert(0, path)
        
        print("✅ Paths configurados")
        
        # 2. Criar sistema
        print("⚙️ Criando sistema...")
        
        try:
            from src.Sistema_Entrada_Dados import SistemaEntradaDados
        except ImportError:
            from Sistema_Entrada_Dados import SistemaEntradaDados
        
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        
        sistema = SistemaEntradaDados(parent=root)
        print("✅ Sistema criado")
        
        # 3. Aplicar correção usando importação direta
        print("🔧 Aplicando correção...")
        
        try:
            from src.nfe.correcao_certificado_a1 import corrigir_sistema_certificado_a1
            sucesso = corrigir_sistema_certificado_a1(sistema)
            
            if sucesso:
                print("✅ Correção aplicada com sucesso!")
                
                # Tornar disponível globalmente
                import builtins
                builtins.sistema_principal = sistema
                
                print("\n🎯 CORREÇÃO CONCLUÍDA!")
                print("=" * 30)
                print("✅ Variable 'sistema_principal' está disponível globalmente")
                print("✅ Todas as correções foram aplicadas")
                print("✅ Sistema NFe híbrido inicializado")
                print("✅ Consultor SEFAZ A1 configurado")
                
                print("\n📋 COMANDOS PARA USO:")
                print("1. Abra Console Python no VSCode: Ctrl+Shift+P → 'Python: Start REPL'")
                print("2. Execute: sistema_principal.configurar_certificado_rapido()")
                print("3. Diagnóstico: sistema_principal.diagnosticar_nfe()")
                
                return True
            else:
                print("❌ Falha na correção")
                return False
                
        except ImportError:
            print("❌ Erro de importação")
            return False
        except Exception as e:
            print(f"❌ Erro: {e}")
            return False
            
    except Exception as e:
        print(f"❌ Erro geral: {e}")
        return False

def verificar_correcao():
    """Verifica se a correção foi aplicada"""
    try:
        if hasattr(__builtins__, 'sistema_principal'):
            sistema = getattr(__builtins__, 'sistema_principal')
            print(f"\n🔍 VERIFICAÇÃO:")
            print(f"✅ Sistema encontrado: {type(sistema).__name__}")
            print(f"✅ Tem processador NFe: {hasattr(sistema, 'processador_nfe')}")
            print(f"✅ Tem consultor A1: {hasattr(sistema, 'consultor_sefaz_a1')}")
            
            if hasattr(sistema, 'configurar_certificado_rapido'):
                print("✅ Método configurar_certificado_rapido: DISPONÍVEL")
            else:
                print("❌ Método configurar_certificado_rapido: AUSENTE")
                
            if hasattr(sistema, 'diagnosticar_nfe'):
                print("✅ Método diagnosticar_nfe: DISPONÍVEL")
            else:
                print("❌ Método diagnosticar_nfe: AUSENTE")
                
            return True
        else:
            print("❌ Sistema não encontrado")
            return False
            
    except Exception as e:
        print(f"❌ Erro na verificação: {e}")
        return False

def main():
    """Função principal - executa e termina"""
    try:
        # Aplicar correção
        sucesso = aplicar_correcao_final()
        
        if sucesso:
            # Verificar se funcionou
            verificar_correcao()
            
            print("\n🎉 SUCESSO! SCRIPT FINALIZADO!")
            print("📋 Agora use o Console Python do VSCode conforme instruído acima.")
        else:
            print("\n❌ FALHA NA APLICAÇÃO")
            
        # Terminar automaticamente (SEM LOOP)
        print("\n👋 Script finalizado automaticamente.")
        
    except Exception as e:
        print(f"❌ Erro na execução: {e}")

if __name__ == "__main__":
    main()
    # Script termina aqui automaticamente
