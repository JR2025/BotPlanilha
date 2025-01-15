import tkinter as tk
from tkinter import filedialog, messagebox
from extrairDados import iniciar_extracao  # Certifique-se de importar corretamente

def selecionar_pasta():
    pasta_selecionada = filedialog.askdirectory()
    if pasta_selecionada:
        entrada_pasta.delete(0, tk.END)
        entrada_pasta.insert(0, pasta_selecionada)

def carregar_pasta():
    caminho_pasta = entrada_pasta.get()  # Pega o caminho da pasta que foi inserido na entrada
    if caminho_pasta:
        caminho_planilha = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Salvar planilha como"
        )

        if caminho_planilha:
            try:
                # Passa o caminho da pasta e o caminho da planilha para iniciar a extração
                iniciar_extracao(caminho_pasta, caminho_planilha)
                messagebox.showinfo("Sucesso", "Processamento concluído com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento: {e}")
        else:
            messagebox.showwarning("Aviso", "Nenhum caminho para salvar a planilha foi selecionado!")
    else:
        messagebox.showwarning("Aviso", "Nenhuma pasta foi selecionada!")

def limpar_pasta():
    entrada_pasta.delete(0, tk.END)
    messagebox.showinfo("Limpar", "O campo foi limpo. Nenhuma pasta está selecionada.")

# Configuração da janela principal
janela = tk.Tk()
janela.title("Extração de Dados de PDFs")
janela.geometry("500x200")

# Rótulo e entrada para o caminho da pasta
rotulo_pasta = tk.Label(janela, text="Pasta com arquivos PDF:")
rotulo_pasta.pack(pady=10)

entrada_pasta = tk.Entry(janela, width=50)
entrada_pasta.pack(pady=5)

# Botão "Selecionar"
botao_selecionar = tk.Button(janela, text="Selecionar", command=selecionar_pasta)
botao_selecionar.pack(pady=5)

# Frame para os botões "Carregar" e "Limpar"
frame_botoes = tk.Frame(janela)
frame_botoes.pack(pady=5)

botao_carregar = tk.Button(frame_botoes, text="Carregar", command=carregar_pasta)
botao_carregar.pack(side=tk.LEFT, padx=5)

botao_limpar = tk.Button(frame_botoes, text="Limpar", command=limpar_pasta)
botao_limpar.pack(side=tk.LEFT, padx=5)

# Loop principal da aplicação
janela.mainloop()
