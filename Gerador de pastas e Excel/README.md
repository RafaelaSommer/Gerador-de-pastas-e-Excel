📦 Aplicações Python – Gerador de Pastas, Gerador de Excel & Conversor de Ícone

Este repositório contém três ferramentas Python com interface gráfica (Tkinter) desenvolvidas para automatizar tarefas comuns do dia a dia: criação de pastas, geração de planilhas Excel e conversão de imagens para ícones .ico.

As aplicações são simples, leves, funcionam em qualquer computador com Python instalado e foram projetadas para facilitar o fluxo de trabalho de usuários iniciantes ou avançados.

🧩 Conteúdo da Pasta
Arquivo	Função
Gerador de Pastas.py	Cria automaticamente estruturas completas de diretórios em poucos cliques.
Gerador Excel.py	Gera planilhas Excel personalizadas, com cabeçalhos e ajuste automático.
conversor_ico.py	Converte qualquer imagem .png, .jpg etc. para arquivo .ico.
logo.ico / logo.png (opcional)	Ícones usados na interface gráfica (Tkinter).
⚙️ Como instalar e executar
1️⃣ Instale as dependências

O arquivo requirements.txt está na pasta raiz do projeto. Execute:

pip install -r ../requirements.txt

2️⃣ Rode o aplicativo desejado
🗂️ Gerador de Pastas
python "Gerador de Pastas.py"

📊 Gerador de Excel
python "Gerador Excel.py"

🖼️ Conversor de Ícone
python "conversor_ico.py"

🖥️ Interfaces Gráficas (GUI)

Todos os programas utilizam Tkinter, abrindo janelas intuitivas e fáceis de usar.
Nenhum conhecimento de terminal é necessário após a execução.

As aplicações funcionam com ou sem os arquivos de logo.

🗂️ Gerador de Pastas – Como Funciona

O Gerador de Pastas permite criar estruturas completas automaticamente.

✨ Funcionalidades:

Seleção da pasta base onde tudo será criado

Criação de uma pasta principal com nome personalizado

Área para inserir vários nomes (um por linha)

Criação de subpastas padrão para cada nome

Subpastas secundárias opcionais

Interface moderna em Dark Mode

Avisos e validações automáticas

🧠 Fluxo de uso:

Escolha a pasta base onde tudo será criado

Digite o nome da pasta principal

Adicione a lista de nomes (um por linha)

Informe as subpastas gerais e secundárias

Clique em Gerar

O programa cria automaticamente:

Pasta Principal/
    Nome 1/
        Subpasta 1/
        Subpasta 2/
    Nome 2/
        Subpasta 1/
        Subpasta 2/
    ...

📊 Gerador de Excel – Como Funciona

O Gerador permite criar planilhas completas em poucos segundos.

✨ Funcionalidades:

Número de linhas e colunas definidas pelo usuário

Preenchimento dos valores diretamente na interface Tkinter

Cabeçalhos na primeira linha são obrigatórios

Ajuste automático da largura das colunas

Exportação automática para .xlsx

Arquivo salvo diretamente na Área de Trabalho

🧠 Fluxo de uso:

Defina o número de colunas e linhas

Preencha os dados na interface

Clique em Salvar Excel

O arquivo é gerado automaticamente e salvo na sua área de trabalho

🖼️ conversor_ico.py – Conversor de PNG/JPG para ICO

Ferramenta simples e prática que converte qualquer imagem em ícone .ico.

✨ Funcionalidades:

Seleção de arquivo .png, .jpg, .jpeg etc.

Escolha do local de salvamento

Conversão rápida via biblioteca Pillow

Suporte a múltiplos tamanhos de ícone

Ideal para projetos Python com Tkinter

🧠 Fluxo de uso:

Abra o programa

Clique em Selecionar Imagem

Escolha onde salvar o .ico

Pronto! O arquivo será criado instantaneamente

📌 Observações importantes

Todos os programas funcionam mesmo sem os arquivos logo.ico ou logo.png.

O Gerador de Excel sempre salva o arquivo diretamente na Área de Trabalho.

Recomendado usar Python 3.10+.

🧪 Tecnologias utilizadas

Python 3.x

Tkinter – interface gráfica

Pillow – usada no conversor_ico.py

openpyxl – criação de arquivos Excel

os / shutil – manipulação de diretórios

🤝 Suporte & Personalizações

Se precisar de melhorias ou versões avançadas, posso criar:

✔ Versão em .exe (compatível com Windows)
✔ Salvamento e carregamento automático de configurações
✔ Banco de dados para histórico
✔ Interface moderna (Tkinter + ttkbootstrap)
✔ Versão multilíngue
✔ Tema claro/escuro
✔ Recursos extras para Excel
✔ Instalação automática (Setup Installer)

É só pedir! 😎🚀