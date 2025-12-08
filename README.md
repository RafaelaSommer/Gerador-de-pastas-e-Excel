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

Preenchimento direto na interface Tkinter

Ajuste automático de largura

Exporta arquivo .xlsx

Salva automaticamente na Área de Trabalho

✔ Conversor ICO (conversor_ico.py)

Converte qualquer imagem .png, .jpg ou .jpeg em .ico:

Interface simples e direta

Seleção de imagem via Tkinter

Suporte a múltiplos tamanhos de ícone

Ideal para projetos Tkinter que utilizam ícones .ico

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
    ├── logo.png   (opcional)
    ├── logo.ico   (opcional)
    │
    └── README.md  (explicação interna da pasta)

🛠️ Dependências

O arquivo requirements.txt está dentro da pasta “Gerador de Pastas e Excel”.

Para instalar:

pip install -r "Gerador de Pastas e Excel/requirements.txt"

Bibliotecas utilizadas:

tkinter

Pillow

openpyxl

os / shutil

▶ Como Executar

Acesse a pasta onde os scripts estão:

cd "Gerador de Pastas e Excel"

🗂️ Gerador de Pastas
python "Gerador de Pastas.py"

📊 Gerador de Excel
python "Gerador Excel.py"

🖼️ Conversor ICO
python "conversor_ico.py"

📌 Observações

Os arquivos logo.png e logo.ico são opcionais.
Caso existam, são carregados automaticamente na interface.

Todos os programas funcionam sem internet.

Compatíveis com Python 3.8+.

As aplicações são janelas Tkinter — não é necessário usar o terminal após abrir.

📦 Melhorias Futuras (sugestões)

Criar executáveis .exe com PyInstaller

Interface modernizada com ttkbootstrap

Tema claro/escuro configurável

Múltiplos idiomas (PT/EN/ES)

Salvamento de modelos de pastas

Salvamento de modelos de planilhas

Criar instalador para Windows (.exe instalável)