"""
Ferramenta de diagnóstico do Sistema de Gestão Financeira
Este script verifica as configurações e relata problemas comuns.
"""
import os
import sys
import platform
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, scrolledtext
from datetime import datetime

# Adicionar o diretório raiz ao path para encontrar módulos
def add_project_root():
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent if "src" in str(current_dir) else current_dir
    if str(project_root) not in sys.path:
        sys.path.append(str(project_root))

add_project_root()

# Carregar variáveis de ambiente
try:
    from dotenv import load_dotenv
    load_dotenv()
    env_loaded = True
except ImportError:
    env_loaded = False

class VerificadorSistema:
    def __init__(self):
        self.resultados = []
        self.root = tk.Tk()
        self.root.title("Verificador do Sistema de Gestão Financeira")
        self.root.geometry("700x500")
        self.criar_interface()
        
    def criar_interface(self):
        # Frame superior para botões
        frame_botoes = tk.Frame(self.root)
        frame_botoes.pack(fill=tk.X, padx=10, pady=10)
        
        # Botão para executar verificação
        btn_verificar = tk.Button(frame_botoes, text="Verificar Sistema", 
                               command=self.executar_verificacao, 
                               bg="#4CAF50", fg="white", padx=10, pady=5)
        btn_verificar.pack(side=tk.LEFT, padx=5)
        
        # Botão para salvar relatório
        btn_salvar = tk.Button(frame_botoes, text="Salvar Relatório", 
                            command=self.salvar_relatorio,
                            bg="#2196F3", fg="white", padx=10, pady=5)
        btn_salvar.pack(side=tk.LEFT, padx=5)
        
        # Área de texto para resultados
        self.txt_resultado = scrolledtext.ScrolledText(self.root, wrap=tk.WORD)
        self.txt_resultado.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Mensagem inicial
        self.adicionar_linha("Bem-vindo ao Verificador do Sistema de Gestão Financeira")
        self.adicionar_linha("Clique em 'Verificar Sistema' para iniciar o diagnóstico")
        
    def adicionar_linha(self, texto):
        """Adiciona uma linha no relatório"""
        self.txt_resultado.insert(tk.END, texto + "\n")
        self.txt_resultado.see(tk.END)
        # Atualizar a interface
        self.root.update_idletasks()
        self.resultados.append(texto)
        
    def verificar_ambiente(self):
        """Verifica configurações do ambiente"""
        self.adicionar_linha("\n=== CONFIGURAÇÕES DE AMBIENTE ===")
        
        # Verificar variável de ambiente
        env = os.getenv('SISTEMA_AMBIENTE', 'Não definido')
        self.adicionar_linha(f"SISTEMA_AMBIENTE: {env}")
        
        if env == 'Não definido':
            self.adicionar_linha("⚠️ AVISO: Variável SISTEMA_AMBIENTE não está definida!")
            self.adicionar_linha("   - Isso pode causar problemas no acesso aos arquivos")
            self.adicionar_linha("   - Verifique se o arquivo .env está presente na pasta raiz")
            self.adicionar_linha("   - Confirme que o conteúdo do arquivo .env está correto (SISTEMA_AMBIENTE=producao)")
        
        # Verificar se dotenv foi carregado
        self.adicionar_linha(f"Dotenv carregado: {'Sim' if env_loaded else 'Não'}")
        if not env_loaded:
            self.adicionar_linha("⚠️ AVISO: Não foi possível carregar o módulo dotenv!")
            self.adicionar_linha("   - Verifique se python-dotenv está instalado")
        
        # Sistema operacional
        sistema = platform.system()
        self.adicionar_linha(f"Sistema Operacional: {sistema} ({platform.release()})")
        
    def verificar_google_drive(self):
        """Verifica caminhos do Google Drive"""
        self.adicionar_linha("\n=== GOOGLE DRIVE ===")
        
        # Lista de possíveis caminhos do Google Drive
        caminhos_windows = [
            Path("H:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios"),
            Path(os.path.expanduser("~")) / "Google Drive",
            Path(os.path.expanduser("~")) / "AppData/Local/Google/Drive/shared_drives"
        ]
        
        caminhos_mac = [
            Path("/Users/emiliamargareth/Library/CloudStorage/GoogleDrive-emilia.mga@gmail.com/Meu Drive"),
            Path(os.path.expanduser("~")) / "Library/CloudStorage/GoogleDrive-emilia.mga@gmail.com/Meu Drive",
            Path(os.path.expanduser("~")) / "Google Drive"
        ]
        
        self.adicionar_linha("Procurando Google Drive...")
        
        encontrado = False
        is_windows = platform.system() == 'Windows'
        
        # Verificar caminhos baseado no sistema operacional
        caminhos = caminhos_windows if is_windows else caminhos_mac
        
        for caminho in caminhos:
            self.adicionar_linha(f"Verificando: {caminho}")
            if caminho.exists():
                self.adicionar_linha(f"✅ Google Drive encontrado em: {caminho}")
                
                # Verificar pastas específicas
                pasta_financeiro = caminho / "Financeiro"
                if pasta_financeiro.exists():
                    self.adicionar_linha(f"✅ Pasta Financeiro encontrada: {pasta_financeiro}")
                    
                    pasta_base = pasta_financeiro / "Planilhas_Base"
                    if pasta_base.exists():
                        self.adicionar_linha(f"✅ Pasta Planilhas_Base encontrada: {pasta_base}")
                    else:
                        self.adicionar_linha(f"❌ Pasta Planilhas_Base NÃO encontrada em: {pasta_base}")
                    
                    pasta_clientes = pasta_financeiro / "Clientes"
                    if pasta_clientes.exists():
                        self.adicionar_linha(f"✅ Pasta Clientes encontrada: {pasta_clientes}")
                    else:
                        self.adicionar_linha(f"❌ Pasta Clientes NÃO encontrada em: {pasta_clientes}")
                else:
                    self.adicionar_linha(f"❌ Pasta Financeiro NÃO encontrada em: {pasta_financeiro}")
                
                encontrado = True
                break
        
        if not encontrado:
            self.adicionar_linha("❌ ERRO: Google Drive não encontrado em nenhum local padrão!")
            self.adicionar_linha("   - Verifique se o Google Drive está instalado e sincronizado")
            self.adicionar_linha("   - Verifique se o caminho correto está configurado no sistema")
    
    def verificar_arquivos_locais(self):
        """Verifica caminhos de arquivos locais"""
        self.adicionar_linha("\n=== ARQUIVOS LOCAIS ===")
        
        caminho_base = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Planilhas_Base')
        self.adicionar_linha(f"Verificando caminho local: {caminho_base}")
        
        if caminho_base.exists():
            self.adicionar_linha(f"✅ Caminho base local existe")
            
            # Verificar arquivos específicos
            arquivo_clientes = caminho_base / "clientes.xlsx"
            if arquivo_clientes.exists():
                self.adicionar_linha(f"✅ Arquivo clientes.xlsx encontrado")
            else:
                self.adicionar_linha(f"❌ Arquivo clientes.xlsx NÃO encontrado")
                
            arquivo_modelo = caminho_base / "MODELO.xlsx"
            if arquivo_modelo.exists():
                self.adicionar_linha(f"✅ Arquivo MODELO.xlsx encontrado")
            else:
                self.adicionar_linha(f"❌ Arquivo MODELO.xlsx NÃO encontrado")
        else:
            self.adicionar_linha(f"❌ Caminho base local NÃO existe")
    
    def verificar_importacao_config(self):
        """Tenta importar config e verifica seus valores"""
        self.adicionar_linha("\n=== TESTE DE IMPORTAÇÃO ===")
        try:
            # Primeiro tenta importar diretamente
            try:
                self.adicionar_linha("Tentando importar config diretamente...")
                import config
                config_importado = True
            except ImportError:
                # Depois tenta importar via src
                try:
                    self.adicionar_linha("Tentando importar de src.config...")
                    from src import config
                    config_importado = True
                except ImportError:
                    # Por último, tenta importar via config.config
                    try:
                        self.adicionar_linha("Tentando importar de config.config...")
                        from config import config
                        config_importado = True
                    except ImportError:
                        config_importado = False
                        self.adicionar_linha("❌ ERRO: Não foi possível importar o módulo config!")
            
            if config_importado:
                self.adicionar_linha("✅ Módulo config importado com sucesso")
                self.adicionar_linha(f"ENV = {config.ENV}")
                self.adicionar_linha(f"BASE_PATH = {config.BASE_PATH}")
                self.adicionar_linha(f"PASTA_CLIENTES = {config.PASTA_CLIENTES}")
                
                # Verificar se as pastas existem
                self.adicionar_linha(f"BASE_PATH existe? {config.BASE_PATH.exists()}")
                self.adicionar_linha(f"PASTA_CLIENTES existe? {config.PASTA_CLIENTES.exists()}")
        except Exception as e:
            self.adicionar_linha(f"❌ ERRO ao testar config: {str(e)}")
    
    def executar_verificacao(self):
        """Executa todas as verificações"""
        self.txt_resultado.delete(1.0, tk.END)
        self.resultados = []
        
        # Cabeçalho do relatório
        self.adicionar_linha("===== RELATÓRIO DE VERIFICAÇÃO DO SISTEMA =====")
        self.adicionar_linha(f"Data e hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        self.adicionar_linha(f"Usuário: {os.getenv('USERNAME', 'Desconhecido')}")
        
        # Executar verificações
        self.verificar_ambiente()
        self.verificar_google_drive()
        self.verificar_arquivos_locais()
        self.verificar_importacao_config()
        
        # Conclusão
        self.adicionar_linha("\n===== CONCLUSÃO =====")
        if "❌ ERRO" in "\n".join(self.resultados):
            self.adicionar_linha("❌ Foram encontrados ERROS que podem impedir o funcionamento correto do sistema.")
        elif "⚠️ AVISO" in "\n".join(self.resultados):
            self.adicionar_linha("⚠️ Foram encontrados AVISOS que podem afetar o funcionamento do sistema.")
        else:
            self.adicionar_linha("✅ Todas as verificações foram concluídas sem erros críticos.")
            
        self.adicionar_linha("\nSe necessário, compartilhe este relatório com o suporte técnico.")
    
    def salvar_relatorio(self):
        """Salva o relatório em um arquivo de texto"""
        if not self.resultados:
            messagebox.showinfo("Aviso", "Execute a verificação primeiro para gerar um relatório.")
            return
        
        # Criar pasta de logs se não existir
        logs_dir = Path("logs")
        logs_dir.mkdir(exist_ok=True)
        
        # Nome do arquivo com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = logs_dir / f"verificacao_sistema_{timestamp}.txt"
        
        try:
            with open(nome_arquivo, 'w', encoding='utf-8') as f:
                for linha in self.resultados:
                    f.write(linha + "\n")
            
            messagebox.showinfo("Sucesso", f"Relatório salvo em:\n{nome_arquivo}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar relatório: {str(e)}")
    
    def run(self):
        """Inicia a aplicação"""
        self.root.mainloop()

if __name__ == "__main__":
    app = VerificadorSistema()
    app.run()