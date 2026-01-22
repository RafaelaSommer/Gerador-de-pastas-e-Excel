import os
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk

import openpyxl
from openpyxl.styles import Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------- TEMA ---------------------------- #
BG = "#1e1e1e"
CARD = "#2b2b2b"
INPUT_BG = "#3a3a3a"
TEXT = "#f0f0f0"
BTN_BG = "#6C63FF"
BTN_FG = "white"

# ---------------------------- BOTÃO ---------------------------- #
def styled_button(master, text, command):
    return tk.Button(
        master,
        text=text,
        command=command,
        bg=BTN_BG,
        fg=BTN_FG,
        relief="flat",
        bd=0,
        font=("Segoe UI", 11, "bold"),
        activebackground="#827CFF",
        activeforeground="white"
    )

# ---------------------------- UTIL ---------------------------- #
def centralizar_janela(janela, largura, altura):
    janela.update_idletasks()
    ws = janela.winfo_screenwidth()
    hs = janela.winfo_screenheight()
    x = (ws - largura) // 2
    y = (hs - altura) // 2
    janela.geometry(f"{largura}x{altura}+{x}+{y}")

entradas_celulas = []

# ---------------------------- POPUP SALVAR ---------------------------- #
def popup_salvar_arquivo(parent):
    popup = tk.Toplevel(parent)
    popup.title("Salvar planilha Excel")
    popup.configure(bg=CARD)
    popup.resizable(False, False)

    largura = 420
    altura = 200

    x = (popup.winfo_screenwidth() - largura) // 2
    y = (popup.winfo_screenheight() - altura) // 2
    popup.geometry(f"{largura}x{altura}+{x}+{y}")

    resultado = {"nome": None}

    tk.Label(
        popup,
        text="Digite o nome do arquivo Excel",
        bg=CARD,
        fg=TEXT,
        font=("Segoe UI", 12, "bold")
    ).pack(pady=(25, 10))

    entrada = tk.Entry(
        popup,
        font=("Segoe UI", 12),
        justify="center",
        width=30
    )
    entrada.pack(ipady=6)
    entrada.focus()

    frame_botoes = tk.Frame(popup, bg=CARD)
    frame_botoes.pack(pady=25)

    def confirmar():
        nome = entrada.get().strip()
        if nome:
            resultado["nome"] = nome
            popup.destroy()

    def cancelar():
        popup.destroy()

    tk.Button(
        frame_botoes,
        text="OK",
        width=12,
        bg=BTN_BG,
        fg="white",
        relief="flat",
        command=confirmar
    ).pack(side="left", padx=10)

    tk.Button(
        frame_botoes,
        text="Cancelar",
        width=12,
        bg="#555555",
        fg="white",
        relief="flat",
        command=cancelar
    ).pack(side="left", padx=10)

    popup.transient(parent)
    popup.grab_set()
    parent.wait_window(popup)

    return resultado["nome"]

# ---------------------------- GRADE ---------------------------- #
def criar_grade():
    global entradas_celulas
    for w in frame_grade_inner.winfo_children():
        w.destroy()
    entradas_celulas = []

    try:
        n_colunas = int(entry_colunas.get())
        n_linhas = int(entry_linhas.get())
        if n_colunas <= 0 or n_linhas <= 0:
            raise ValueError
    except:
        messagebox.showerror("Erro", "Digite números válidos.")
        return

    centralizar_janela(root, 1100, 650)

    for r in range(n_linhas + 1):
        linha = []
        for c in range(n_colunas):
            e = tk.Entry(
                frame_grade_inner,
                width=30,
                justify="center",
                bg=INPUT_BG,
                fg=TEXT,
                relief="flat",
                font=("Segoe UI", 11),
                insertbackground="white"
            )
            e.grid(
                row=r,
                column=c,
                padx=15,
                pady=10,
                ipadx=20,
                ipady=12,
                sticky="nsew"
            )

            if r == 0:
                e.config(font=("Segoe UI", 11, "bold"), bg="#4b4b4b")

            linha.append(e)
        entradas_celulas.append(linha)

    for c in range(n_colunas):
        frame_grade_inner.grid_columnconfigure(c, weight=1, minsize=260)

    canvas_grade.configure(scrollregion=canvas_grade.bbox("all"))
    centralizar_grade()

# ---------------------------- EXCEL ---------------------------- #
def ajustar_colunas(ws, dados):
    for col_idx in range(len(dados[0])):
        max_len = max(len(str(row[col_idx])) if row[col_idx] else 0 for row in dados)
        letra = get_column_letter(col_idx + 1)
        ws.column_dimensions[letra].width = min(max_len + 2, 45)

        for r in range(1, len(dados) + 1):
            ws[f"{letra}{r}"].alignment = Alignment(
                wrap_text=True,
                vertical="center"
            )

def gerar_excel():
    if not entradas_celulas:
        messagebox.showerror("Erro", "Crie a grade primeiro.")
        return

    dados = []
    for r, linha in enumerate(entradas_celulas):
        valores = []
        for e in linha:
            valor = e.get().strip()
            if r == 0 and not valor:
                messagebox.showerror("Erro", "Preencha todos os cabeçalhos.")
                return
            valores.append(valor)
        dados.append(valores)

    nome = popup_salvar_arquivo(root)
    if not nome:
        return

    caminho = os.path.join(os.path.expanduser("~"), "Desktop", f"{nome}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dados"

    for linha in dados:
        ws.append(linha)

    ws.freeze_panes = "A2"

    ultima_col = get_column_letter(len(dados[0]))
    ultima_linha = len(dados)

    tabela = Table(
        displayName="TabelaDados",
        ref=f"A1:{ultima_col}{ultima_linha}"
    )

    tabela.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showRowStripes=True
    )

    ws.add_table(tabela)

    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for r in range(1, ultima_linha + 1):
        for c in range(1, len(dados[0]) + 1):
            ws.cell(row=r, column=c).border = thin

    ajustar_colunas(ws, dados)
    wb.save(caminho)

    messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{caminho}")

# ---------------------------- INTERFACE ---------------------------- #
root = tk.Tk()
root.title("Gerador de Excel")
root.configure(bg=BG)
centralizar_janela(root, 520, 360)

# Logo
try:
    logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
    if os.path.exists(logo_path):
        img = Image.open(logo_path).resize((72, 72))
        logo = ImageTk.PhotoImage(img)
        tk.Label(root, image=logo, bg=BG).pack(pady=8)
        root.iconphoto(True, logo)
except:
    pass

frame_config = tk.Frame(root, bg=CARD)
frame_config.pack(pady=10)

topo = tk.Frame(frame_config, bg=CARD)
topo.pack(padx=25, pady=15)

tk.Label(topo, text="Colunas", bg=CARD, fg=TEXT).grid(row=0, column=0, padx=15)
entry_colunas = tk.Entry(topo, bg=INPUT_BG, fg=TEXT, width=10, justify="center")
entry_colunas.grid(row=1, column=0, padx=15)

tk.Label(topo, text="Linhas", bg=CARD, fg=TEXT).grid(row=0, column=1, padx=15)
entry_linhas = tk.Entry(topo, bg=INPUT_BG, fg=TEXT, width=10, justify="center")
entry_linhas.grid(row=1, column=1, padx=15)

frame_btn = tk.Frame(frame_config, bg=CARD)
frame_btn.pack(pady=10)

styled_button(frame_btn, "Criar Grade", criar_grade).pack(side="left", padx=10)
styled_button(frame_btn, "Gerar Excel", gerar_excel).pack(side="left", padx=10)

# ---------------------------- GRADE ---------------------------- #
container_grade = tk.Frame(root, bg=BG)
container_grade.pack(fill="both", expand=True, padx=20, pady=10)

canvas_grade = tk.Canvas(container_grade, bg=BG, highlightthickness=0)
canvas_grade.pack(side="left", fill="both", expand=True)

scroll = tk.Scrollbar(container_grade, orient="vertical", command=canvas_grade.yview)
scroll.pack(side="right", fill="y")
canvas_grade.configure(yscrollcommand=scroll.set)

frame_grade_inner = tk.Frame(canvas_grade, bg=BG)
canvas_window = canvas_grade.create_window((0, 0), window=frame_grade_inner, anchor="n")

def centralizar_grade(event=None):
    canvas_width = canvas_grade.winfo_width()
    frame_width = frame_grade_inner.winfo_reqwidth()
    x = max((canvas_width - frame_width) // 2, 0)
    canvas_grade.coords(canvas_window, x, 0)

frame_grade_inner.bind("<Configure>", lambda e: (
    canvas_grade.configure(scrollregion=canvas_grade.bbox("all")),
    centralizar_grade()
))

canvas_grade.bind("<Configure>", centralizar_grade)

root.mainloop()