# 📁✨ Gerador de Pastas • 📊 Gerador de Excel • 🖼️ Conversor ICO  
### Suite de Ferramentas Desktop em Python para Automação e Produtividade

Este repositório reúne **três aplicativos desktop com interface Tkinter**, criados para aumentar sua produtividade no dia a dia com automação, organização e conversão de arquivos — tudo **100% offline** e compatível com **Windows**.

---

## 🚀 Funcionalidades Principais

### 🗂️ Gerador de Pastas
Ferramenta completa para criar estruturas de diretórios automaticamente:

- Seleção da **pasta base**
- Criação automática da **pasta principal**
- Lista de subpastas (um nome por linha)
- Criação de subpastas gerais e secundárias
- Interface moderna com **Dark Mode**
- 100% offline

---

### 📊 Gerador de Excel
Gera arquivos Excel sem precisar abrir o programa:

- Define colunas e linhas diretamente na interface
- Preenchimento instantâneo via Tkinter
- **Ajuste automático** de largura das colunas
- Exporta para `.xlsx`
- Salva automaticamente na **Área de Trabalho**

---

### 🖼️ Conversor ICO
Converta imagens comuns para ícones `.ico`:

- Suporta `.png`, `.jpg`, `.jpeg`
- Interface simples de seleção de imagem
- Converte para vários tamanhos de ícone
- Ideal para projetos Tkinter ou atalhos personalizados

---

## 📂 Estrutura Recomendada

📦 Projeto
│
├── README.md
│
└── Gerador de Pastas e Excel/
├── requirements.txt
│
├── Gerador de Pastas.py
├── Gerador Excel.py
├── conversor_ico.py
│
├── logo.png (opcional)
├── logo.ico (opcional)
│
└── README.md


---

## 🛠️ Instalação das Dependências

```bash
pip install -r "Gerador de Pastas e Excel/requirements.txt"

Bibliotecas utilizadas

tkinter

Pillow

openpyxl

os

shutil

▶ Como Executar

Acesse a pasta do projeto:

cd "Gerador de Pastas e Excel"

🗂️ Gerador de Pastas
python "Gerador de Pastas.py"

📊 Gerador de Excel
python "Gerador Excel.py"

🖼️ Conversor ICO
python "conversor_ico.py"

📌 Observações

logo.png e logo.ico são opcionais.
Se existirem, são carregados automaticamente.

Funcionamento completamente offline.

Compatível com Python 3.8+.

Programas Tkinter: não é necessário terminal após abrir.

🚧 Melhorias Futuras

Criar executáveis .exe com PyInstaller

Interface modernizada com ttkbootstrap

Alternância entre tema claro/escuro

Suporte a vários idiomas (PT/EN/ES)

Salvamento de modelos de pastas

Salvamento de modelos de planilhas

Criar instalador .exe para Windows

⭐ Contribuições

Sinta-se à vontade para enviar sugestões, melhorias e abrir PRs!
Ferramentas desenvolvidas para facilitar seu fluxo de trabalho e evoluir continuamente.