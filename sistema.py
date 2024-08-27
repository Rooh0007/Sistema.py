import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image
import os


class ClientManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerenciador de Clientes")
        self.root.geometry("800x600")
        self.root.configure(bg="#f5f5f5")

        # DataFrame inicial
        self.client_df = pd.DataFrame(columns=["Nome", "Código", "CNPJ", "Status"])

        self.create_widgets()

    def create_widgets(self):
        # Frame de Informações do Cliente
        info_frame = tk.Frame(self.root, bg="#ffffff", padx=10, pady=10)
        info_frame.pack(pady=10, fill=tk.X)

        tk.Label(info_frame, text="Nome:", bg="#ffffff").grid(row=0, column=0, padx=5, pady=5)
        tk.Label(info_frame, text="Código:", bg="#ffffff").grid(row=1, column=0, padx=5, pady=5)
        tk.Label(info_frame, text="CNPJ:", bg="#ffffff").grid(row=2, column=0, padx=5, pady=5)
        tk.Label(info_frame, text="Status:", bg="#ffffff").grid(row=3, column=0, padx=5, pady=5)

        self.name_var = tk.StringVar()
        self.code_var = tk.StringVar()
        self.cnpj_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ativo")

        tk.Entry(info_frame, textvariable=self.name_var).grid(row=0, column=1, padx=5, pady=5)
        tk.Entry(info_frame, textvariable=self.code_var).grid(row=1, column=1, padx=5, pady=5)
        tk.Entry(info_frame, textvariable=self.cnpj_var).grid(row=2, column=1, padx=5, pady=5)
        ttk.Combobox(info_frame, textvariable=self.status_var, values=["Ativo", "Não Ativo"]).grid(row=3, column=1,
                                                                                                   padx=5, pady=5)

        # Frame de Botões
        button_frame = tk.Frame(self.root, bg="#ffffff", padx=10, pady=10)
        button_frame.pack(pady=10, fill=tk.X)

        tk.Button(button_frame, text="Incluir", command=self.add_client).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Excluir", command=self.delete_client).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Atualizar Planilha", command=self.update_spreadsheet).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Salvar Tabela", command=self.save_table_as_image).pack(side=tk.LEFT, padx=5)

        # Tabela de Clientes
        self.tree = ttk.Treeview(self.root, columns=("Nome", "Código", "CNPJ", "Status"), show="headings")
        self.tree.heading("Nome", text="Nome")
        self.tree.heading("Código", text="Código")
        self.tree.heading("CNPJ", text="CNPJ")
        self.tree.heading("Status", text="Status")
        self.tree.pack(pady=10, fill=tk.BOTH, expand=True)

    def add_client(self):
        name = self.name_var.get()
        code = self.code_var.get()
        cnpj = self.cnpj_var.get()
        status = self.status_var.get()

        if not name or not code or not cnpj or not status:
            messagebox.showerror("Erro", "Todos os campos devem ser preenchidos")
            return

        new_data = pd.DataFrame([[name, code, cnpj, status]], columns=self.client_df.columns)
        self.client_df = pd.concat([self.client_df, new_data], ignore_index=True)
        self.update_treeview()

        self.name_var.set("")
        self.code_var.set("")
        self.cnpj_var.set("")
        self.status_var.set("Ativo")

    def delete_client(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("Erro", "Selecione um cliente para excluir")
            return

        for item in selected_item:
            values = self.tree.item(item, "values")
            self.client_df = self.client_df[~self.client_df.apply(lambda row: row.tolist() == list(values), axis=1)]

        self.update_treeview()

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for index, row in self.client_df.iterrows():
            self.tree.insert("", tk.END, values=row.tolist())

    def update_spreadsheet(self):
        try:
            self.client_df.to_excel("clientes_atualizados.xlsx", index=False)
            messagebox.showinfo("Sucesso", "Planilha atualizada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível atualizar a planilha: {e}")

    def save_table_as_image(self):
        if self.client_df.empty:
            messagebox.showerror("Erro", "A tabela está vazia. Adicione clientes antes de salvar a imagem.")
            return

        fig, ax = plt.subplots(figsize=(8, 6))
        ax.axis('off')

        # Criar uma tabela
        table = ax.table(cellText=self.client_df.values,
                         colLabels=self.client_df.columns,
                         cellLoc='center',
                         loc='center',
                         bbox=[0, 0, 1, 1])

        table.auto_set_font_size(False)
        table.set_fontsize(10)

        # Abrir caixa de diálogo para salvar a imagem
        file_path = filedialog.asksaveasfilename(defaultextension=".png",
                                                 filetypes=[("PNG files", "*.png"), ("All files", "*.*")],
                                                 title="Salvar imagem como")
        if file_path:
            plt.savefig(file_path, bbox_inches='tight', pad_inches=0.1)
            plt.close(fig)
            messagebox.showinfo("Sucesso", f"Tabela salva como imagem com sucesso!\nArquivo salvo em: {file_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ClientManagerApp(root)
    root.mainloop()
