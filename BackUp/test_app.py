# test_app.py
import tkinter as tk
from tkinter import ttk, messagebox

def main():
    root = tk.Tk()
    root.title("Teste")
    
    ttk.Label(root, text="Se você está vendo esta mensagem, o tkinter está funcionando!").pack(pady=20)
    ttk.Button(root, text="Fechar", command=root.destroy).pack()
    
    # Mostrar uma mensagem popup
    messagebox.showinfo("Teste", "Aplicação de teste iniciada com sucesso!")
    
    root.mainloop()

if __name__ == "__main__":
    main()