# config.py
from pathlib import Path
import platform
import os

# ============================================================================
# CORREÇÃO PRINCIPAL: USAR O AMBIENTE_CONFIG.PY
# ============================================================================
# Em vez de detectar ambiente de forma independente,
# vamos IMPORTAR do ambiente_config.py que já funciona!
# ============================================================================

try:
    # Importar o ambiente já detectado corretamente
    from src.ambiente_config import config_ambiente
    
    # Usar a detecção que JÁ FUNCIONA
    if config_ambiente.eh_producao():
        ENV = 'producao'
        print("🟢 AMBIENTE DETECTADO: PRODUÇÃO (via ambiente_config.py)")
    else:
        ENV = 'teste'
        print("🟨 AMBIENTE DETECTADO: TESTE (via ambiente_config.py)")
        
except ImportError:
    # Fallback apenas para desenvolvimento (quando executado diretamente)
    print("⚠️ ambiente_config não disponível - usando fallback")
    ENV = os.getenv('SISTEMA_AMBIENTE', 'teste').lower()
    print(f"🎯 Ambiente (fallback): {ENV}")

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
    print("🧪 MODO TESTE ATIVADO")
    
    # Caminhos de teste - SEMPRE usar local
    BASE_PATH = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Planilhas_Base')
    PASTA_CLIENTES = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Clientes')
    
    print(f"📁 BASE_PATH (TESTE): {BASE_PATH}")
    print(f"📁 PASTA_CLIENTES (TESTE): {PASTA_CLIENTES}")
    
else:  # ENV == 'producao'
    print("🏭 MODO PRODUÇÃO ATIVADO")
    
    if IS_WINDOWS:
        # Lista de possíveis caminhos do Google Drive em Windows
        possiveis_caminhos = [
            Path("H:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios"),
            # Caminho alternativo comum para Google Drive no Windows
            Path(os.path.expanduser("~")) / "Google Drive",
            # Outro formato possível 
            Path(os.path.expanduser("~")) / "AppData/Local/Google/Drive/shared_drives"
        ]
        
        # Tenta encontrar um caminho válido
        for caminho in possiveis_caminhos:
            print(f"🔍 Testando caminho: {caminho}")
            if caminho.exists():
                GOOGLE_DRIVE_PATH = caminho
                print(f"✅ Google Drive encontrado em: {GOOGLE_DRIVE_PATH}")
                break
            else:
                print(f"❌ Caminho não existe: {caminho}")
                
    elif IS_MAC:
        # Lista de possíveis caminhos do Google Drive em Mac
        possiveis_caminhos = [
            Path("/Users/emiliamargareth/Library/CloudStorage/GoogleDrive-emilia.mga@gmail.com/Meu Drive"),
            Path(os.path.expanduser("~")) / "Library/CloudStorage/GoogleDrive-emilia.mga@gmail.com/Meu Drive",
            Path(os.path.expanduser("~")) / "Google Drive"
        ]
        
        # Tenta encontrar um caminho válido
        for caminho in possiveis_caminhos:
            print(f"🔍 Testando caminho: {caminho}")
            if caminho.exists():
                GOOGLE_DRIVE_PATH = caminho
                print(f"✅ Google Drive encontrado em: {GOOGLE_DRIVE_PATH}")
                break
            else:
                print(f"❌ Caminho não existe: {caminho}")
    
    # Se não encontrou o Google Drive, usar fallback
    if GOOGLE_DRIVE_PATH is None:
        print("⚠️ AVISO: Google Drive não encontrado. Usando caminho local como fallback.")
        # Usar caminho local como fallback
        BASE_PATH = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Planilhas_Base')
        PASTA_CLIENTES = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Clientes')
    else:
        # Define os caminhos base para diferentes pastas
        BASE_PATH = GOOGLE_DRIVE_PATH / "Financeiro/Planilhas_Base"
        PASTA_CLIENTES = GOOGLE_DRIVE_PATH / "Financeiro/Clientes"
        
    print(f"📁 BASE_PATH (PRODUÇÃO): {BASE_PATH}")
    print(f"📁 PASTA_CLIENTES (PRODUÇÃO): {PASTA_CLIENTES}")

# Define caminhos específicos
ARQUIVO_CLIENTES = BASE_PATH / "clientes.xlsx"
ARQUIVO_FORNECEDORES = BASE_PATH / "base_fornecedores.xlsx"
ARQUIVO_MODELO = BASE_PATH / "MODELO.xlsx"
ARQUIVO_CONTROLE = BASE_PATH / "controle_taxa_adm.xlsx"
PASTA_RH = BASE_PATH / "Planilhas_RH"
ARQUIVO_PARAMETROS_MATERIAIS = BASE_PATH / "parametros_materiais.json"

# Verificação dos diretórios
print(f"\n📋 VERIFICAÇÃO DE CAMINHOS:")
print(f"=" * 40)

if GOOGLE_DRIVE_PATH is not None:
    print(f"📁 GOOGLE_DRIVE_PATH: {GOOGLE_DRIVE_PATH}")
    print(f"   Existe? {'✅' if GOOGLE_DRIVE_PATH.exists() else '❌'}")

print(f"📁 BASE_PATH: {BASE_PATH}")
print(f"   Existe? {'✅' if BASE_PATH.exists() else '❌'}")

print(f"📁 PASTA_CLIENTES: {PASTA_CLIENTES}")
print(f"   Existe? {'✅' if PASTA_CLIENTES.exists() else '❌'}")

print(f"\n📄 ARQUIVOS IMPORTANTES:")
print(f"📄 ARQUIVO_CLIENTES: {ARQUIVO_CLIENTES}")
print(f"   Existe? {'✅' if ARQUIVO_CLIENTES.exists() else '❌'}")

print(f"📄 ARQUIVO_MODELO: {ARQUIVO_MODELO}")
print(f"   Existe? {'✅' if ARQUIVO_MODELO.exists() else '❌'}")

# Criar as pastas se não existirem
try:
    print(f"\n🔧 CRIANDO PASTAS NECESSÁRIAS...")
    BASE_PATH.mkdir(parents=True, exist_ok=True)
    PASTA_CLIENTES.mkdir(parents=True, exist_ok=True)
    print(f"✅ Pastas criadas/verificadas com sucesso")
except Exception as e:
    print(f"❌ Erro ao criar pastas: {e}")

def verificar_arquivos():
    """Verifica se todos os arquivos necessários estão acessíveis"""
    print(f"\n🔍 VERIFICAÇÃO COMPLETA DE ARQUIVOS:")
    print(f"=" * 45)
    
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
        
        if not existe:
            erros.append(f"Arquivo não encontrado: {arquivo}")
    
    if erros:
        print(f"\n❌ ERROS ENCONTRADOS:")
        for erro in erros:
            print(f"   • {erro}")
        raise FileNotFoundError(f"{len(erros)} arquivo(s) não encontrado(s)")
    else:
        print(f"\n✅ Todos os arquivos estão acessíveis!")

print(f"\n🏁 Configuração concluída - Ambiente: {ENV}")
print(f"="*50)