# 📁 Gerador de Pastas & 📊 Gerador de Excel

Este projeto reúne duas ferramentas desktop desenvolvidas em Python com interface gráfica para auxiliar na organização de arquivos e criação de planilhas personalizadas.

---

## 🚀 Recursos Principais

### ✔ Gerador de Pastas
- Criação automática de diretórios
- Permite definir:
  - Pasta base
  - Pasta principal
  - Lista de nomes
  - Subpastas gerais
  - Subpastas secundárias
- Interface moderna em **dark mode**

### ✔ Gerador de Excel
- Cria planilhas `.xlsx` dinamicamente
- Definição de número de colunas e linhas
- Preenchimento direto na interface
- Ajuste automático de largura das colunas
- Salva o arquivo diretamente na **Área de Trabalho**

---

## 📂 Estrutura Recomendada do Projeto

📦 Projeto
│
│─ README.md ← (este arquivo)
│─ requirements.txt
│
└─ app/
│─ Gerador de Pastas.py
│─ Gerador Excel.py
│─ logo.png (opcional)
│─ logo.ico (opcional)
│─ README.md (interno - explicação da pasta)

---

## 🛠️ Dependências

As bibliotecas necessárias estão listadas em **requirements.txt**.  
Para instalar:

```bash
pip install -r requirements.txt

---

Bibliotecas utilizadas:

tk / tkinter

Pillow

openpyxl

▶ Execução

Entre na pasta app:

cd app


Execute o programa desejado:

python "Gerador de Pastas.py"


ou

python "Gerador Excel.py"

📌 Observações

Os arquivos logo.ico e logo.png são opcionais. Caso existam, o programa usará automaticamente na interface.

Os dois scripts possuem janela gráfica e podem ser executados em Windows sem terminal aberto.

Interface 100% offline — não depende da internet.

📦 Futuras Melhorias (sugestões)

Gerar executável .exe com PyInstaller

Salvar e carregar modelos de planilha

Tema claro e escuro selecionável pelo usuário

Idioma configurável

Projeto desenvolvido em Python utilizando Tkinter, Pillow e OpenPyXL.
Sinta-se à vontade para modificar, distribuir e melhorar.