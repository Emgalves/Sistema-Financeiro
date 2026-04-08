# src/config/config.py
"""
Configuração de caminhos do sistema
Usa ambiente_config.py como fonte única de verdade para detecção de ambiente

SEMPRE VERIFICAR ESTES ARQUIVOS PARA MANTER CONSISTÊNCIA:
 - src/ambiente_config.py
 - src/config/paths.py
 - src/config/config.py
 - src/config/__init__.py

CAMINHOS DE PRODUÇÃO:
  Os caminhos de produção são lidos do arquivo config_caminhos.json,
  localizado na mesma pasta do executável (ex: S:\Gestão\config_caminhos.json).
  Se o servidor ou letra de drive mudar, basta editar esse arquivo —
  sem necessidade de rebuild do executável.

  O JSON deve apontar para a pasta que contém diretamente
  'Planilhas_Base' e 'Clientes', ou seja:
    - Sua máquina : H:\...\.shortcut-targets-by-id\...\Relatórios\Financeiro
    - Cliente     : Z:\Servidor\Relatórios\Financeiro
"""

from pathlib import Path
import platform
import os
import sys
import json
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

# BASE_DADOS = pasta que contém diretamente 'Planilhas_Base' e 'Clientes'
# Exemplos:
#   Sua máquina : H:\...\.shortcut-targets-by-id\...\Relatórios\Financeiro
#   Cliente     : Z:\Servidor\Relatórios\Financeiro
BASE_DADOS = None
GOOGLE_DRIVE_PATH = None  # mantido por compatibilidade com outros módulos

# ============================================================================
# CONFIGURAÇÃO DE CAMINHOS BASEADA NO AMBIENTE
# ============================================================================

if ENV == 'teste':
    print("\n🧪 MODO TESTE ATIVADO")
    print("-" * 70)

    BASE_PATH = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Planilhas_Base')
    PASTA_CLIENTES = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Clientes')

    print(f"📁 BASE_PATH (TESTE): {BASE_PATH}")
    print(f"📁 PASTA_CLIENTES (TESTE): {PASTA_CLIENTES}")
    print(f"✅ BASE_PATH existe? {BASE_PATH.exists()}")
    print(f"✅ PASTA_CLIENTES existe? {PASTA_CLIENTES.exists()}")

else:  # ENV == 'producao'
    print("\n🏭 MODO PRODUÇÃO ATIVADO")
    print("-" * 70)

    if IS_WINDOWS:

        # ====================================================================
        # PASSO 1: Tentar ler caminho do arquivo JSON externo
        # ====================================================================

        def encontrar_config_json():
            """
            Busca config_caminhos.json em locais possíveis:
            1. Mesma pasta do executável (produção)
            2. Diretório de trabalho atual
            3. Pasta raiz do projeto (desenvolvimento)
            """
            candidatos = []

            if getattr(sys, 'frozen', False):
                candidatos.append(Path(sys.executable).parent / "config_caminhos.json")

            candidatos.append(Path(os.getcwd()) / "config_caminhos.json")
            candidatos.append(Path(__file__).resolve().parent.parent.parent / "config_caminhos.json")

            for candidato in candidatos:
                print(f"   🔍 Buscando JSON em: {candidato}")
                if candidato.exists():
                    print(f"   ✅ Encontrado: {candidato}")
                    return candidato

            return None

        caminho_json = encontrar_config_json()

        if caminho_json:
            try:
                with open(caminho_json, 'r', encoding='utf-8') as f:
                    dados_json = json.load(f)

                caminho_principal   = dados_json.get("caminho_dados", "").strip()
                caminho_alternativo = dados_json.get("caminho_dados_alternativo", "").strip()

                print(f"\n📄 config_caminhos.json lido com sucesso")
                print(f"   Caminho principal  : {caminho_principal}")
                print(f"   Caminho alternativo: {caminho_alternativo}")

                for label, caminho_str in [("principal", caminho_principal), ("alternativo", caminho_alternativo)]:
                    if caminho_str:
                        p = Path(caminho_str)
                        if p.exists():
                            BASE_DADOS = p
                            print(f"   ✅ Caminho {label} acessível: {p}")
                            break
                        else:
                            print(f"   ❌ Caminho {label} não acessível: {p}")

            except Exception as e:
                print(f"   ⚠️ Erro ao ler config_caminhos.json: {e}")
        else:
            print(f"\n⚠️ config_caminhos.json não encontrado — usando busca automática")

        # ====================================================================
        # PASSO 2: Busca automática se JSON não resolveu
        # Todos os caminhos apontam para a pasta que contém
        # diretamente 'Planilhas_Base' e 'Clientes'
        # ====================================================================

        if BASE_DADOS is None:
            print(f"\n🔍 BUSCANDO CAMINHO DE DADOS AUTOMATICAMENTE:")

            possiveis_caminhos = [
                # Servidor por nome UNC (independe de letra de drive)
                Path("//servidor/Servidor/Relatórios/Financeiro"),
                Path("//servidor/Servidor/Relatorios/Financeiro"),
                Path("//servidor/Relatórios/Financeiro"),
                Path("//servidor/Relatorios/Financeiro"),
                # Letra Z (servidor mapeado — cliente)
                Path("Z:/Servidor/Relatórios/Financeiro"),
                Path("Z:/Servidor/Relatorios/Financeiro"),
                Path("Z:/Relatórios/Financeiro"),
                Path("Z:/Relatorios/Financeiro"),
                # Letra Y
                Path("Y:/Servidor/Relatórios/Financeiro"),
                Path("Y:/Servidor/Relatorios/Financeiro"),
                # Google Drive — sua máquina
                Path("H:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios/Financeiro"),
                Path("G:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios/Financeiro"),
                Path("H:/Drives compartilhados/Relatórios/Financeiro"),
                Path("G:/Drives compartilhados/Relatórios/Financeiro"),
                Path("H:/Relatórios/Financeiro"),
                Path("G:/Relatórios/Financeiro"),
                Path("F:/Relatórios/Financeiro"),
                Path("E:/Relatórios/Financeiro"),
            ]

            for idx, caminho in enumerate(possiveis_caminhos, 1):
                print(f"   [{idx}/{len(possiveis_caminhos)}] {caminho}")
                if caminho.exists():
                    BASE_DADOS = caminho
                    print(f"   ✅ ENCONTRADO!")
                    break
                else:
                    print(f"   ❌ Não existe")

        # ====================================================================
        # PASSO 3: Validação — fallback com mensagem clara
        # ====================================================================

        if BASE_DADOS is None:
            print("\n" + "❌" * 20)
            print("❌  ERRO: CAMINHO DE DADOS NÃO ENCONTRADO!")
            print("❌" * 20)
            print("""
Para corrigir SEM rebuild:
  1. Abra o arquivo config_caminhos.json na pasta do executável
     (ex: S:\\Gestão\\config_caminhos.json)
  2. Edite "caminho_dados" com o caminho correto até a pasta Financeiro
  3. Salve e reabra o sistema

O caminho deve apontar para a pasta que contém
diretamente as pastas 'Planilhas_Base' e 'Clientes'.

Exemplos:
  Z:/Servidor/Relatórios/Financeiro
  //servidor/Servidor/Relatórios/Financeiro
  H:/.shortcut-targets-by-id/.../Relatórios/Financeiro
""")
            print("⚠️ Forçando modo TESTE como fallback de emergência")
            ENV = 'teste'
            BASE_PATH = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Planilhas_Base')
            PASTA_CLIENTES = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Clientes')
            print(f"📁 BASE_PATH (FALLBACK TESTE): {BASE_PATH}")
            print(f"📁 PASTA_CLIENTES (FALLBACK TESTE): {PASTA_CLIENTES}")
            print("❌" * 20)

    elif IS_MAC:
        possiveis_caminhos_mac = [
            Path(os.path.expanduser("~")) / "Library/CloudStorage/GoogleDrive-emilia.mga@gmail.com/Meu Drive/Relatórios/Financeiro",
            Path(os.path.expanduser("~")) / "Google Drive/Relatórios/Financeiro",
        ]

        print(f"\n🔍 BUSCANDO CAMINHO DE DADOS (Mac):")
        for idx, caminho in enumerate(possiveis_caminhos_mac, 1):
            print(f"   [{idx}/{len(possiveis_caminhos_mac)}] {caminho}")
            if caminho.exists():
                BASE_DADOS = caminho
                print(f"   ✅ ENCONTRADO!")
                break
            else:
                print(f"   ❌ Não existe")

    # ====================================================================
    # Definir BASE_PATH e PASTA_CLIENTES a partir de BASE_DADOS
    # ====================================================================

    if BASE_DADOS is not None and ENV == 'producao':
        GOOGLE_DRIVE_PATH = BASE_DADOS  # compatibilidade com outros módulos
        BASE_PATH      = BASE_DADOS / "Planilhas_Base"
        PASTA_CLIENTES = BASE_DADOS / "Clientes"

        print(f"\n📂 BASE_DADOS    : {BASE_DADOS}")
        print(f"📁 BASE_PATH     : {BASE_PATH}")
        print(f"📁 PASTA_CLIENTES: {PASTA_CLIENTES}")

        if not BASE_PATH.exists():
            print(f"❌ ERRO: BASE_PATH não existe: {BASE_PATH}")
            raise FileNotFoundError(f"BASE_PATH não encontrado: {BASE_PATH}")

        if not PASTA_CLIENTES.exists():
            print(f"❌ ERRO: PASTA_CLIENTES não existe: {PASTA_CLIENTES}")
            raise FileNotFoundError(f"PASTA_CLIENTES não encontrado: {PASTA_CLIENTES}")

        print(f"✅ Todos os caminhos base existem e estão acessíveis!")

print(f"\n{'='*70}\n")

# ============================================================================
# ARQUIVOS ESPECÍFICOS
# ============================================================================

ARQUIVO_CLIENTES             = BASE_PATH / "Clientes.xlsx"
ARQUIVO_FORNECEDORES         = BASE_PATH / "base_fornecedores.xlsx"
ARQUIVO_MODELO               = BASE_PATH / "MODELO.xlsx"
ARQUIVO_CONTROLE             = BASE_PATH / "controle_taxa_adm.xlsx"
PASTA_RH                     = BASE_PATH / "Planilhas_RH"
ARQUIVO_PARAMETROS_MATERIAIS = BASE_PATH / "parametros_materiais.json"

# ============================================================================
# VERIFICAÇÃO FINAL
# ============================================================================

print(f"📋 VERIFICAÇÃO FINAL DE CAMINHOS:")
print(f"{'='*70}")
print(f"\n📁 BASE_PATH: {BASE_PATH}")
print(f"    Existe? {'✅' if BASE_PATH.exists() else '❌'}")
print(f"\n📁 PASTA_CLIENTES: {PASTA_CLIENTES}")
print(f"    Existe? {'✅' if PASTA_CLIENTES.exists() else '❌'}")
print(f"\n📄 ARQUIVOS CRÍTICOS:")

for nome, arquivo in [
    ("Clientes",     ARQUIVO_CLIENTES),
    ("Fornecedores", ARQUIVO_FORNECEDORES),
    ("Modelo",       ARQUIVO_MODELO),
]:
    print(f"    {nome:15} {'✅' if arquivo.exists() else '❌'} {arquivo.name}")

# Criar pastas se modo teste
if ENV == 'teste':
    try:
        BASE_PATH.mkdir(parents=True, exist_ok=True)
        PASTA_CLIENTES.mkdir(parents=True, exist_ok=True)
        print(f"✅ Pastas de teste criadas/verificadas")
    except Exception as e:
        print(f"❌ Erro ao criar pastas: {e}")


def verificar_arquivos():
    """Verifica se todos os arquivos necessários estão acessíveis"""
    print(f"\n🔍 VERIFICAÇÃO DETALHADA DE ARQUIVOS:")
    print(f"{'='*70}")

    arquivos = [
        ('CLIENTES',     ARQUIVO_CLIENTES),
        ('FORNECEDORES', ARQUIVO_FORNECEDORES),
        ('MODELO',       ARQUIVO_MODELO),
        ('CONTROLE',     ARQUIVO_CONTROLE),
    ]

    erros = []

    for nome, arquivo in arquivos:
        existe = arquivo.exists()
        print(f"{'✅' if existe else '❌'} {nome}: {arquivo}")

        if existe:
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
print(f"{'='*70}\n")
