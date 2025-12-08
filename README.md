📁 Gerador de Pastas • 📊 Gerador de Excel • 🖼️ Conversor ICO
Ferramentas desktop em Python para automação e produtividade

Este projeto reúne três aplicativos desktop com interface gráfica (Tkinter), desenvolvidos para facilitar tarefas de organização, criação de planilhas e conversão de imagens em ícones .ico.

🚀 Recursos Principais
✔ Gerador de Pastas

Ferramenta para criação automática de estruturas de diretórios:

Define pasta base

Cria pasta principal

Aceita lista de nomes (um por linha)

Cria subpastas gerais e subpastas secundárias

Interface moderna Dark Mode

Funcionamento 100% offline

✔ Gerador de Excel

Gera planilhas Excel sem precisar abrir o Excel:

Define colunas e linhas

Preenchimento direto na interface

Ajuste automático de largura

Exporta .xlsx

Salva diretamente na Área de Trabalho

✔ Conversor ICO (conversor_ico.py)

Converte qualquer imagem .png/.jpg/.jpeg em .ico:

Interface simples e direta

Seleção de imagem

Suporte a múltiplos tamanhos

Ideal para ícones de aplicações Tkinter

📂 Estrutura Recomendada do Projeto
📦 Projeto
│
├── README.md
│
└── Gerador de Pastas e Excel/
    ├── requirements.txt   ← (fica aqui!)
    │
    ├── Gerador de Pastas.py
    ├── Gerador Excel.py
    ├── conversor_ico.py
    │
    ├── logo.png (opcional)
    ├── logo.ico (opcional)
    │
    └── README.md (explicação interna da pasta)

🛠️ Dependências

O arquivo requirements.txt está dentro da pasta “Gerador de Pastas e Excel”.

Instale executando:

pip install -r "Gerador de Pastas e Excel/requirements.txt"

Bibliotecas utilizadas:

tkinter

Pillow

openpyxl

os / shutil

▶ Como Executar

Entre na pasta onde os scripts estão:

cd "Gerador de Pastas e Excel"

🗂️ Gerador de Pastas
python "Gerador de Pastas.py"

📊 Gerador de Excel
python "Gerador Excel.py"

🖼️ Conversor ICO
python "conversor_ico.py"

📌 Observações

logo.png e logo.ico são opcionais.
Se existirem, serão carregados automaticamente.

Tudo funciona sem internet.

Projetos feitos em Python 3.8+.

Softwares executam por janelas TK, sem necessidade de terminal após iniciados.

📦 Melhorias Futuras (sugestões)

Criar .exe com PyInstaller

Interface com ttkbootstrap

Tema claro/escuro configurável

Idioma selecionável

Salvar modelos de planilhas e estruturas de pastas

Criar instalador para Windows