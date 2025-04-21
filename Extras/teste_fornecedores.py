# teste_fornecedores.py
try:
    from relatorio_fornecedores import RelatorioFornecedores
    print("Módulo importado com sucesso")
    
    app = RelatorioFornecedores()
    print("Classe instanciada com sucesso")
    app.root.mainloop()
except Exception as e:
    import traceback
    print(f"Erro: {str(e)}")
    traceback.print_exc()