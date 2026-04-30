# GestorPrefeitura

Projeto de automação web focado no sistema Gestor Prefeitura utilizando a biblioteca [Playwright](https://playwright.dev/python/) com Python.

## Pré-requisitos

- Python 3.8+
- Google Chrome instalado

## Instalação

1. Clone o repositório:
```bash
git clone https://github.com/Luzemir/GestorPrefeitura.git
cd GestorPrefeitura
```

2. Crie e ative um ambiente virtual (recomendado):
```bash
python -m venv venv
# No Windows:
venv\Scripts\activate
# No Linux/Mac:
source venv/bin/activate
```

3. Instale as dependências:
```bash
pip install -r requirements.txt
```

4. Instale os navegadores do Playwright (opcional, dependendo do uso):
```bash
playwright install
```

## Como Executar

Os scripts deste projeto exigem que o Google Chrome seja iniciado previamente com uma porta de depuração aberta. Isso permite que o Playwright se conecte a uma sessão já existente.

Para garantir que a automação não interfira no seu uso diário e que abra um Chrome "limpo" (sem cache, avisos ou perfis que atrapalhem o script), utilize **exclusivamente o comando abaixo no PowerShell**:

1. Feche **totalmente** as janelas de teste antigas do Chrome (o Chrome de uso diário pode ficar aberto).
2. Abra o **PowerShell** e cole o seguinte comando:

```powershell
Start-Process "C:\Program Files\Google\Chrome\Application\chrome.exe" -ArgumentList "--remote-debugging-port=9222","--user-data-dir=$env:TMP\chrome_test_profile","--no-first-run","--no-default-browser-check"
```

3. Com a aba amarela aberta (porta 9222), faça login no Gestor Prefeitura.
4. Agora sim, no terminal da pasta do projeto, execute os scripts:

```bash
python testes\test_click.py
python testes\test_export.py
python testes\test_fill.py
```

## Arquivos Principais

* `test_click.py`: Script para navegar pelos menus (Gerenciar NFSe -> Livro Fiscal).
* `test_export.py`: Exemplos de preenchimento de datas e clique para exportar planilhas (XLS).
* `test_fill.py`: Scripts similares misturando Playwright e execução de scripts Javascript direto na página.
* `find_pages.py`: Utilitário para listar as abas abertas e encontrar inputs.

## Regra de Organização (PC Casa / PC Escritório)
Para mantermos o projeto limpo em diferentes computadores, foi criado uma **Regra Local**. Sempre que o projeto começar a acumular arquivos soltos na pasta principal, execute:

```bash
python organize.py
```
Esse script automaticamente varre o diretório e move de forma inteligente arquivos `.py` para as suas pastas focais (`scripts` ou `testes`), e arquivos de dados e imagens temporárias para a pasta `data/`.
