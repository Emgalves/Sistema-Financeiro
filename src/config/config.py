# config.py
from pathlib import Path
import platform
import os

# CORREÇÃO 1: Verificação mais robusta da variável de ambiente
def obter_ambiente():
    """Obtém o ambiente atual com múltiplas verificações"""
    
    # Método 1: Variável de ambiente do sistema
    env_sistema = os.getenv('SISTEMA_AMBIENTE', '').lower().strip()
    print(f"🔍 ENV do sistema (SISTEMA_AMBIENTE): '{env_sistema}'")
    
    # Método 2: Variável alternativa
    env_alt = os.getenv('AMBIENTE', '').lower().strip()
    print(f"🔍 ENV alternativo (AMBIENTE): '{env_alt}'")
    
    # Método 3: Verificar arquivo de configuração local
    try:
        config_file = Path(__file__).parent / '.env_config'
        if config_file.exists():
            with open(config_file, 'r', encoding='utf-8') as f:
                env_arquivo = f.read().strip().lower()
                print(f"🔍 ENV do arquivo (.env_config): '{env_arquivo}'")
        else:
            env_arquivo = ''
    except Exception as e:
        print(f"⚠️ Erro ao ler arquivo de config: {e}")
        env_arquivo = ''
    
    # Priorizar: arquivo > SISTEMA_AMBIENTE > AMBIENTE > padrão
    if env_arquivo and env_arquivo in ['teste', 'test', 'desenvolvimento', 'dev']:
        return 'teste'
    elif env_sistema and env_sistema in ['teste', 'test', 'desenvolvimento', 'dev']:
        return 'teste'
    elif env_alt and env_alt in ['teste', 'test', 'desenvolvimento', 'dev']:
        return 'teste'
    else:
        return 'producao'

# USAR A FUNÇÃO CORRIGIDA
ENV = obter_ambiente()
print(f"🎯 Ambiente determinado: {ENV}")

# CORREÇÃO 2: Adicionar modo de debug para variáveis
def debug_variaveis_ambiente():
    """Mostra todas as variáveis de ambiente relacionadas"""
    print("\n" + "="*50)
    print("🔧 DEBUG - VARIÁVEIS DE AMBIENTE")
    print("="*50)
    
    variaveis_interesse = [
        'SISTEMA_AMBIENTE', 'AMBIENTE', 'ENV', 'ENVIRONMENT',
        'COMPUTERNAME', 'USERNAME', 'USERDOMAIN'
    ]
    
    for var in variaveis_interesse:
        valor = os.getenv(var, 'NÃO DEFINIDA')
        print(f"  {var}: {valor}")
    
    print("="*50 + "\n")

# Executar debug se necessário
debug_variaveis_ambiente()

# Detecta o sistema operacional
IS_WINDOWS = platform.system() == 'Windows'
IS_MAC = platform.system() == 'Darwin'

print(f"💻 Sistema operacional: {platform.system()}")

# Inicializa a variável GOOGLE_DRIVE_PATH como None
GOOGLE_DRIVE_PATH = None

# CORREÇÃO 3: Configuração mais clara para cada ambiente
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

# CORREÇÃO 4: Verificação melhorada dos diretórios
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

# CORREÇÃO 5: Funções auxiliares para mudança de ambiente
def criar_arquivo_ambiente_teste():
    """Cria arquivo local para forçar ambiente de teste"""
    try:
        config_file = Path(__file__).parent / '.env_config'
        with open(config_file, 'w', encoding='utf-8') as f:
            f.write('teste')
        print(f"✅ Arquivo de configuração criado: {config_file}")
        print(f"💡 Reinicie o sistema para aplicar o ambiente de TESTE")
        return True
    except Exception as e:
        print(f"❌ Erro ao criar arquivo de config: {e}")
        return False

def criar_arquivo_ambiente_producao():
    """Cria arquivo local para forçar ambiente de produção"""
    try:
        config_file = Path(__file__).parent / '.env_config'
        with open(config_file, 'w', encoding='utf-8') as f:
            f.write('producao')
        print(f"✅ Arquivo de configuração criado: {config_file}")
        print(f"💡 Reinicie o sistema para aplicar o ambiente de PRODUÇÃO")
        return True
    except Exception as e:
        print(f"❌ Erro ao criar arquivo de config: {e}")
        return False

def remover_arquivo_ambiente():
    """Remove arquivo de configuração local"""
    try:
        config_file = Path(__file__).parent / '.env_config'
        if config_file.exists():
            config_file.unlink()
            print(f"✅ Arquivo de configuração removido: {config_file}")
            print(f"💡 Reinicie o sistema - usará variável de ambiente do sistema")
        else:
            print(f"ℹ️ Arquivo de configuração não existe")
        return True
    except Exception as e:
        print(f"❌ Erro ao remover arquivo de config: {e}")
        return False

# CORREÇÃO 6: Mostrar instruções para mudança de ambiente
def mostrar_instrucoes_ambiente():
    """Mostra instruções para mudança de ambiente"""
    print(f"\n" + "="*60)
    print(f"🎯 COMO MUDAR DE AMBIENTE:")
    print(f"="*60)
    print(f"")
    print(f"OPÇÃO 1 - Via arquivo local (RECOMENDADO):")
    print(f"   Para TESTE:")
    print(f"   >>> from src.config.config import criar_arquivo_ambiente_teste")
    print(f"   >>> criar_arquivo_ambiente_teste()")
    print(f"")
    print(f"   Para PRODUÇÃO:")
    print(f"   >>> from src.config.config import criar_arquivo_ambiente_producao")
    print(f"   >>> criar_arquivo_ambiente_producao()")
    print(f"")
    print(f"OPÇÃO 2 - Via variável do sistema Windows:")
    print(f"   1. Abra o Prompt como Administrador")
    print(f"   2. Para TESTE: setx SISTEMA_AMBIENTE \"teste\" /M")
    print(f"   3. Para PRODUÇÃO: setx SISTEMA_AMBIENTE \"producao\" /M")
    print(f"   4. Reinicie o computador")
    print(f"")
    print(f"OPÇÃO 3 - Via variável da sessão (temporário):")
    print(f"   No prompt antes de executar:")
    print(f"   set SISTEMA_AMBIENTE=teste")
    print(f"   python seu_script.py")
    print(f"")
    print(f"="*60)

# Mostrar instruções
mostrar_instrucoes_ambiente()

print(f"\n🏁 Configuração concluída - Ambiente: {ENV}")
print(f"="*50)