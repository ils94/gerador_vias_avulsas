import tkinter as tk
from tkinter import ttk
import os
import gerarTermo
import centralizarJanela


def add_item():
    entry_values = [entry.get() for entry in entries]

    if entry_values[2] and entry_values[3] and entry_values[4]:
        # Insert values into treeview
        tree.insert("", "end", values=(entry_values[2], entry_values[4].upper(), entry_values[3]))


def remove_item():
    selected_item = tree.selection()
    for item in selected_item:
        tree.delete(item)


def edit_item():
    entry_values = [entry.get() for entry in entries]

    selected_item = tree.selection()
    for item in selected_item:
        tree.item(item, values=(entry_values[2], entry_values[4].upper(), entry_values[3]))


def gerar_saida():
    gerarTermo.criar_documento("Autorização de Saída de Material", "PORTADOR", "PORTARIA", entries, tree)


def gerar_transferencia():
    gerarTermo.criar_documento("Termo de Transferência Interna", "CEDENTE", "RECEBEDOR", entries, tree)


root = tk.Tk()
root.title("Gerador de Vias Avulsas")
root.resizable(False, False)
root.geometry("500x500")

if os.path.isfile("documento.ico"):
    root.iconbitmap("documento.ico")

# Frames for organizing widgets
labels_entries_frame = tk.Frame(root)
labels_entries_frame.pack(fill="x")
labels_entries_frame.grid_columnconfigure(1, weight=1)

buttons_frame = tk.Frame(root)
buttons_frame.pack(fill="x")

# Label texts and respective entries using grid
label_texts = ["Origem:", "Destino:", "Item:", "Quantidade:", "Descrição:"]
entries = []

for i, label_text in enumerate(label_texts):
    label = tk.Label(labels_entries_frame, text=label_text)
    label.grid(row=i, column=0, padx=5, pady=5, sticky="w")

    entry = tk.Entry(labels_entries_frame)
    entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
    entries.append(entry)

# Buttons using pack
button_texts = ["Adicionar", "Remover", "Editar"]
commands = [add_item, remove_item, edit_item]

for button_text, command in zip(button_texts, commands):
    button = tk.Button(buttons_frame, text=button_text, command=command)
    button.pack(side="left", padx=5, pady=5)

# Treeview
tree = ttk.Treeview(root, columns=("Item", "Descrição", "QTD"))
tree.heading("Item", text="Item")
tree.heading("Descrição", text="Descrição")
tree.heading("QTD", text="QTD")

# Set different column widths
tree.column("Item", width=50)
tree.column("Descrição", width=350)
tree.column("QTD", width=50)
tree.column("#0", width=1)  # Hide the ID column

tree.pack(fill="both", expand=True, padx=5, pady=5)

# Menu
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

gerar_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Menu", menu=gerar_menu)
gerar_menu.add_command(label="Gerar Autorização de Saída", command=gerar_saida)
gerar_menu.add_command(label="Gerar Transferência Interna", command=gerar_transferencia)

centralizarJanela.center_window(root, 500, 500)

root.mainloop()
