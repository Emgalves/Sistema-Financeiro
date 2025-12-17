# src/config/config.py
"""
Configuração de caminhos do sistema
Usa ambiente_config.py como fonte única de verdade para detecção de ambiente

SEMPRE VERIFICAR ESTES ARQUIVOS PARA MANTER CONSISTÊNCIA:
 - src/ambiente_config.py
    - src/config/paths.py
    - src/config/config.py
    - src/config/__init__.py
"""

from pathlib import Path
import platform
import os
import logging

# Configurar logging básico
logger = logging.getLogger(__name__)

# ============================================================================
# IMPORTAR DETECÇÃO DE AMBIENTE (FONTE ÚNICA DE VERDADE)
# ============================================================================

try:
    from src.ambiente_config import config_ambiente
    ENV = 'producao' if config_ambiente.eh_producao() else 'teste'
    print(f"✅ config.py: Usando ambiente detectado por ambiente_config: {ENV.upper()}")
except ImportError as e:
    logger.warning(f"⚠️ Não foi possível importar ambiente_config: {e}")
    logger.warning("⚠️ Usando fallback (não recomendado)")
    ENV = 'teste'

print(f"\n{'='*70}")
print(f"{'🟢' if ENV == 'producao' else '🟨'} config.py - AMBIENTE: {ENV.upper()}")
print(f"{'='*70}")

# Detecta o sistema operacional
IS_WINDOWS = platform.system() == 'Windows'
IS_MAC = platform.system() == 'Darwin'

print(f"💻 Sistema operacional: {platform.system()}")

# Inicializa a variável GOOGLE_DRIVE_PATH como None
GOOGLE_DRIVE_PATH = None

# ============================================================================
# CONFIGURAÇÃO DE CAMINHOS BASEADA NO AMBIENTE
# ============================================================================

if ENV == 'teste':
    print("\n🧪 MODO TESTE ATIVADO")
    print("-" * 70)
    
    # CAMINHOS DE TESTE (pasta separada com dados de teste)
    BASE_PATH = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Planilhas_Base')
    PASTA_CLIENTES = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Clientes')
    
    print(f"📁 BASE_PATH (TESTE): {BASE_PATH}")
    print(f"📁 PASTA_CLIENTES (TESTE): {PASTA_CLIENTES}")
    print(f"✅ BASE_PATH existe? {BASE_PATH.exists()}")
    print(f"✅ PASTA_CLIENTES existe? {PASTA_CLIENTES.exists()}")
    
else:  # ENV == 'producao'
    print("\n🏭 MODO PRODUÇÃO ATIVADO")
    print("-" * 70)
    
    # Buscar Google Drive
    if IS_WINDOWS:
        possiveis_caminhos = [
            Path("H:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios"),
            Path("G:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios"),
            Path("H:/Drives compartilhados/Relatórios"),
            Path("G:/Drives compartilhados/Relatórios"),
            Path("H:/Relatórios"),
            Path("G:/Relatórios"),
            Path("F:/Relatórios"),
            Path("E:/Relatórios"),
        ]
        
        print(f"\n🔍 BUSCANDO GOOGLE DRIVE:")
        for idx, caminho in enumerate(possiveis_caminhos, 1):
            print(f"   [{idx}/{len(possiveis_caminhos)}] {caminho}")
            
            if caminho.exists():
                GOOGLE_DRIVE_PATH = caminho
                print(f"   ✅ ENCONTRADO!")
                break
            else:
                print(f"   ❌ Não existe")
        
        print()
                
    elif IS_MAC:
        possiveis_caminhos = [
            Path(os.path.expanduser("~")) / "Library/CloudStorage/GoogleDrive-emilia.mga@gmail.com/Meu Drive",
            Path(os.path.expanduser("~")) / "Google Drive",
        ]
        
        print(f"\n🔍 BUSCANDO GOOGLE DRIVE (Mac):")
        for idx, caminho in enumerate(possiveis_caminhos, 1):
            print(f"   [{idx}/{len(possiveis_caminhos)}] {caminho}")
            
            if caminho.exists():
                GOOGLE_DRIVE_PATH = caminho
                print(f"   ✅ ENCONTRADO!")
                break
            else:
                print(f"   ❌ Não existe")
        
        print()
    
    # ====== VALIDAÇÃO CRÍTICA ======
    if GOOGLE_DRIVE_PATH is None:
        print("❌" * 20)
        print("❌  ERRO CRÍTICO: MODO PRODUÇÃO MAS GOOGLE DRIVE NÃO ENCONTRADO!")
        print("❌" * 20)
        print("")
        print("O ambiente foi detectado como PRODUÇÃO, mas não")
        print("foi possível encontrar o Google Drive.")
        print("")
        print("Isso pode acontecer se:")
        print("  • O executável tem sufixo _PRODUCAO no nome")
        print("  • MAS o Google Drive não está sincronizado")
        print("")
        print("AÇÃO: Forçando modo TESTE para evitar erros")
        print("")
        
        ENV = 'teste'
        BASE_PATH = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Planilhas_Base')
        PASTA_CLIENTES = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Clientes')
        
        print(f"📁 BASE_PATH (FALLBACK TESTE): {BASE_PATH}")
        print(f"📁 PASTA_CLIENTES (FALLBACK TESTE): {PASTA_CLIENTES}")
        print("")
        print("❌" * 20)
    else:
        # Define os caminhos base
        BASE_PATH = GOOGLE_DRIVE_PATH / "Financeiro/Planilhas_Base"
        PASTA_CLIENTES = GOOGLE_DRIVE_PATH / "Financeiro/Clientes"
        
        print(f"☁️  GOOGLE_DRIVE_PATH: {GOOGLE_DRIVE_PATH}")
        print(f"📁 BASE_PATH: {BASE_PATH}")
        print(f"📁 PASTA_CLIENTES: {PASTA_CLIENTES}")
        print()
        
        # VALIDAÇÃO: Verificar se os caminhos realmente existem
        if not BASE_PATH.exists():
            print(f"❌ ERRO: BASE_PATH não existe!")
            print(f"   Caminho: {BASE_PATH}")
            raise FileNotFoundError(f"BASE_PATH não encontrado: {BASE_PATH}")
        
        if not PASTA_CLIENTES.exists():
            print(f"❌ ERRO: PASTA_CLIENTES não existe!")
            print(f"   Caminho: {PASTA_CLIENTES}")
            raise FileNotFoundError(f"PASTA_CLIENTES não encontrado: {PASTA_CLIENTES}")
        
        print(f"✅ Todos os caminhos base existem e estão acessíveis!")
        print()

print(f"{'='*70}\n")

# Define caminhos específicos
ARQUIVO_CLIENTES = BASE_PATH / "Clientes.xlsx"
ARQUIVO_FORNECEDORES = BASE_PATH / "base_fornecedores.xlsx"
ARQUIVO_MODELO = BASE_PATH / "MODELO.xlsx"
ARQUIVO_CONTROLE = BASE_PATH / "controle_taxa_adm.xlsx"
PASTA_RH = BASE_PATH / "Planilhas_RH"
ARQUIVO_PARAMETROS_MATERIAIS = BASE_PATH / "parametros_materiais.json"

# Verificação final dos diretórios
print(f"📋 VERIFICAÇÃO FINAL DE CAMINHOS:")
print(f"=" * 70)

if GOOGLE_DRIVE_PATH is not None:
    print(f"☁️  GOOGLE_DRIVE_PATH: {GOOGLE_DRIVE_PATH}")
    print(f"    Existe? {'✅' if GOOGLE_DRIVE_PATH.exists() else '❌'}")

print(f"\n📁 BASE_PATH: {BASE_PATH}")
print(f"    Existe? {'✅' if BASE_PATH.exists() else '❌'}")

print(f"\n📁 PASTA_CLIENTES: {PASTA_CLIENTES}")
print(f"    Existe? {'✅' if PASTA_CLIENTES.exists() else '❌'}")

print(f"\n📄 ARQUIVOS CRÍTICOS:")
arquivos_verificar = [
    ("Clientes", ARQUIVO_CLIENTES),
    ("Fornecedores", ARQUIVO_FORNECEDORES),
    ("Modelo", ARQUIVO_MODELO),
]

for nome, arquivo in arquivos_verificar:
    existe = arquivo.exists()
    print(f"    {nome:15} {'✅' if existe else '❌'} {arquivo.name}")

# Criar as pastas se não existirem (só em modo teste)
if ENV == 'teste':
    try:
        print(f"\n🔧 Criando pastas de teste se necessário...")
        BASE_PATH.mkdir(parents=True, exist_ok=True)
        PASTA_CLIENTES.mkdir(parents=True, exist_ok=True)
        print(f"✅ Pastas criadas/verificadas")
    except Exception as e:
        print(f"❌ Erro ao criar pastas: {e}")

def verificar_arquivos():
    """Verifica se todos os arquivos necessários estão acessíveis"""
    print(f"\n🔍 VERIFICAÇÃO DETALHADA DE ARQUIVOS:")
    print(f"=" * 70)
    
    arquivos = [
        ('CLIENTES', ARQUIVO_CLIENTES),
        ('FORNECEDORES', ARQUIVO_FORNECEDORES), 
        ('MODELO', ARQUIVO_MODELO),
        ('CONTROLE', ARQUIVO_CONTROLE)
    ]
    
    erros = []
    
    for nome, arquivo in arquivos:
        existe = arquivo.exists()
        status = '✅' if existe else '❌'
        print(f"{status} {nome}: {arquivo}")
        
        if existe:
            # Testar se consegue abrir para leitura
            try:
                with open(arquivo, 'rb') as f:
                    f.read(1)
                print(f"    └─ Leitura: ✅")
            except Exception as e:
                print(f"    └─ Leitura: ❌ {e}")
                erros.append(f"Sem permissão de leitura: {arquivo}")
        else:
            erros.append(f"Arquivo não encontrado: {arquivo}")
    
    if erros:
        print(f"\n❌ ERROS ENCONTRADOS:")
        for erro in erros:
            print(f"   • {erro}")
        raise FileNotFoundError(f"{len(erros)} arquivo(s) com problema(s)")
    else:
        print(f"\n✅ Todos os arquivos estão acessíveis!")

print(f"\n🎯 Configuração concluída - Ambiente: {ENV.upper()}")
print(f"=" * 70)
print()