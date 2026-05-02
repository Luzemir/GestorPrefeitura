# GestorSefaz - Design & Stack

## Stack Tecnológica
- **Linguagem Principal**: Python 3.10+
- **Automação Web**: Playwright (responsável pelo controle do browser, navegação e extração de tabelas HTML e manipulação de IFrames, caso existam).
- **Interface Gráfica (GUI)**: CustomTkinter (seguindo o padrão visual já implementado no GestorPrefeitura).
- **Manipulação de Dados**: Pandas (essencial para montagem dos DataFrames, formatação de valores financeiros, verificação de sequência e agrupamento de totais).
- **Geração de Relatórios**: OpenPyXL (integrado ao Pandas, para exportar os dados estruturados em múltiplas abas no Excel).

## Arquitetura e Módulos
Seguindo o modus operandi do ecossistema atual, o GestorSefaz será estruturado modularmente:

1. **Módulo de Interface e Integração (`src/app_gui.py` ou nova aba CustomTkinter)**:
   - Aba/Frame dedicada ao GestorSefaz contendo:
     - Componente para selecionar o arquivo Excel listando as empresas (`Código Domínio`, `Nome Empresa`, `CNPJ`, `Inscrição Estadual`).
     - Campo(s) para informar ou selecionar a Competência (`mmaaaa`).
     - Botão `Iniciar Captura Sefaz`.
     - Área de TextBox dedicada para os logs (status visual do processamento).

2. **Motor Core de Automação (`src/sefaz_bot.py`)**:
   - Controla o Playwright injetando o script no portal da Sefaz.
   - `start()`: Inicia o browser de forma semi-headless/visível, aguardando o usuário autenticar o certificado.
   - `validate_certificate()`: Lê o texto do DOM contendo o perfil autenticado e compara com "LUZEMIR MARTINS BARBOSA".
   - `select_company(ie)`: Fluxo para selecionar "Procurador" e injetar a IE no campo.
   - `extract_nfe_emitidas()`, `extract_nfe_recebidas()`, `extract_cte_emitente()`, `extract_cte_tomador()`: Funções orquestradoras para inserir datas, buscar e iterar nas paginações (150/150).

3. **Módulo de Processamento Excel e Negócio (`src/sefaz_exporter.py` ou integrado ao bot)**:
   - Responsável por receber listas de dicionários ou DataFrames das funções de extração.
   - Identifica falhas em sequência numérica nas NF-e emitidas (ex: se encontrou nota 10 e 12, avisa que a 11 não foi achada/foi pulada).
   - Calcula os somatórios (soma de valores e contagem de notas).
   - Salva o arquivo final `.xlsx` definindo corretamente os nomes de aba. Caso a tabela venha vazia, gera o DataFrame single-row indicando "A Consulta não Retornou Registros".

4. **Gerenciador de Estado e Erros**:
   - Tratativas com `try/except` robustas em cada passo para que a falha de uma empresa não interrompa todo o lote, passando para a próxima e registrando o alerta no log.
