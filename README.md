# 📦 Aplicações Python – Gerador de Pastas • Gerador de Excel • Conversor de Ícone  
### Automação fácil e rápida com interfaces Tkinter

Este repositório reúne **três ferramentas Python com interface gráfica (Tkinter)** desenvolvidas para automatizar tarefas comuns do dia a dia:  
📁 criação de pastas,  
📊 geração de planilhas Excel e  
🖼️ conversão de imagens para ícones `.ico`.

As aplicações são **leves, intuitivas, funcionam 100% offline** e podem ser usadas por qualquer pessoa — desde iniciantes em Python até usuários avançados que precisam agilizar processos.

---

## 🧩 Conteúdo da Pasta

| Arquivo | Função |
|--------|--------|
| **Gerador de Pastas.py** | Cria automaticamente estruturas completas de diretórios. |
| **Gerador Excel.py** | Gera planilhas Excel com cabeçalhos e ajuste automático. |
| **conversor_ico.py** | Converte imagens `.png`, `.jpg`, etc. para `.ico`. |
| **logo.ico / logo.png** *(opcional)* | Ícones exibidos na interface Tkinter. |

---

## ⚙️ Instalação e Execução

### 1️⃣ Instale as dependências  
O arquivo `requirements.txt` está na pasta raiz.

```bash
pip install -r ../requirements.txt

2️⃣ Rode o aplicativo desejado
🗂️ Gerador de Pastas
python "Gerador de Pastas.py"

📊 Gerador de Excel
python "Gerador Excel.py"

🖼️ Conversor de Ícone
python "conversor_ico.py"

🖥️ Interfaces Gráficas (GUI)

✔ Todas as aplicações utilizam Tkinter
✔ Janelas simples, diretas e intuitivas
✔ Não é preciso usar o terminal após abrir
✔ Funcionam com ou sem os logos opcionais
✔ Totalmente offline

🗂️ Gerador de Pastas – Como Funciona

O Gerador de Pastas cria estruturas completas em poucos cliques.

✨ Funcionalidades:

Seleção da pasta base

Criação da pasta principal

Campo para inserir múltiplos nomes (um por linha)

Criação de subpastas padrão

Subpastas secundárias opcionais

Interface moderna em Dark Mode

Validações automáticas e avisos amigáveis

🧠 Fluxo de uso:

Escolha a pasta base

Insira o nome da pasta principal

Adicione a lista de nomes (um por linha)

Informe as subpastas gerais e secundárias

Clique em Gerar

A estrutura gerada será algo como:

Pasta Principal/
    Nome 1/
        Subpasta 1/
        Subpasta 2/
    Nome 2/
        Subpasta 1/
        Subpasta 2/
    ...

📊 Gerador de Excel – Como Funciona

Crie planilhas completas sem abrir o Excel, diretamente via Tkinter.

✨ Funcionalidades:

Definição de número de linhas e colunas

Preenchimento dos dados direto na interface

Cabeçalhos obrigatórios na primeira linha

Ajuste automático da largura das colunas

Exportação para .xlsx

Arquivo salvo automaticamente na Área de Trabalho

🧠 Fluxo de uso:

Defina o número de colunas e linhas

Preencha os dados exibidos na janela

Clique em Salvar Excel

O arquivo será criado automaticamente no desktop do usuário.

🖼️ Conversor de Ícone – PNG/JPG para ICO

Ferramenta rápida para transformar imagens em ícones .ico.

✨ Funcionalidades:

Suporte a .png, .jpg, .jpeg e outros formatos

Escolha do local de salvamento

Conversão instantânea usando Pillow

Vários tamanhos de ícone disponíveis

Ideal para projetos Tkinter ou atalhos personalizados

🧠 Fluxo de uso:

Abra o aplicativo

Clique em Selecionar Imagem

Escolha onde salvar

Pronto — o ícone é gerado na hora!

📌 Observações Importantes

As logos são opcionais — o programa funciona sem elas.

O Gerador de Excel sempre salva na Área de Trabalho.

Recomendado usar Python 3.10+.

Funciona em qualquer sistema com Python instalado.

🧪 Tecnologias Utilizadas

Python 3.x

Tkinter – interface gráfica

Pillow – manipulação de imagens (conversor ICO)

openpyxl – criação de arquivos Excel

os / shutil – manipulação de diretórios

🤝 Suporte & Personalizações

Posso criar versões personalizadas com:

✔ Arquivos .exe para Windows
✔ Interface moderna com ttkbootstrap
✔ Tema claro/escuro
✔ Histórico com banco de dados
✔ Configurações salvas automaticamente
✔ Versão multilíngue
✔ Recursos extras para Excel
✔ Instalador completo (.exe Installer)

Se quiser evoluir este projeto, é só pedir! 😎🚀