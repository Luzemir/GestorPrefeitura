# Plano de Implementação: Automação NFSe Campo Grande

**Objetivo:** Criar uma aplicação capaz de acessar o portal da Nota Fiscal de Serviços de Campo Grande, autenticar via certificado digital, iterar sobre uma lista de empresas, extrair dados financeiros (total de notas, qtde, impostos, etc.) no mês e salvar em uma planilha eletrônica.

## Mapeamento de Interface (Playwright / Selenium)
Com base na análise do HTML (`NFS-e- seleciona cadastro.html` e `NFS-e-livro fiscal.html`) fornecido em ./capturas:

1. **Login e Ponto de Partida:** 
   O robô abrirá o navegador e você fará o login manualmente usando o certificado digital (aguardo de `input` do usuário no terminal). Em seguida, o robô assume quando a tela estiver na posição de "Selecionar Cadastro".
2. **Seleção da Empresa:**
   - *Input CNPJ:* `<input id="frmDados:j_idt98:idCpfCnpj:idInputMask">`
   - *Botão Pesquisar:* Clique no link/botão contendo "Pesquisar".
   - *Tabela e Seleção:* Validar na linha gerada o CNPJ e clicar no check/botão na lateral esquerda.
3. **Navegação de Menus:** 
   - Procurar "Gerenciar NFSE" no menu lateral e ir para "Livro Fiscal".
4. **Campos do Livro Fiscal e Extração:**
   - *Data Inicial:* `<input id="frmRelatorio:j_idt101:j_idt104:idStart">`
   - *Data Final:* `<input id="frmRelatorio...:idEnd">` (ou id correspondente)
   - *Botão Excel:* Existirá um botão do tipo Submit com a ação de Exportar / Gerar Excel. O Playwright irá aguardar o "Download" terminar na pasta padrão do Windows.
5. **Renomeação Inteligente e Processamento:**
   - O arquivo `LivroFiscal.xlsx` será movido da pasta de Downloads do Windows para a pasta `./livros` do projeto.
   - Será renomeado para `{MMAAAA}_{Nome_da_Empresa}.xlsx`.
   - Lê-se essa planilha, faz-se o cálculo das colunas (Valor NF, INSS, IR, Líquido, etc), gravando direto em um `consolidado.xlsx`.
   - O robô navega novamente para a listagem da próxima empresa.
6. **Mecanismo de Retomada (Checkpointing / Resiliência):**
   - O script verificará a Planilha Mestre (`consolidado.xlsx`) ou a pasta `./livros/` na inicialização para identificar quais CNPJs já foram processados na competência atual.
   - Em caso de travamento do site, erro de rede ou timeout, o robô poderá ser reiniciado. Ele pulará automaticamente a etapa das empresas concluídas e **retomará a extração exatamente do CNPJ onde parou**.

### Fase 3: Prova de Conceito (Autenticação e Leitura Simples)
- Script abrirá o navegador visível para você ("headed mode").
- O script vai parar temporariamente para você selecionar o Certificado Digital e colocar o PIN (se houver).
- O script fará a extração do valor de apenas 1 empresa como teste.

### Fase 4: Loop de Empresas, Extração e Relatório Consolidado
- Adicionar o laço de repetição (`for` ou `while`) para trocar de empresa na interface do portal da prefeitura.
- O robô irá navegar até a seção de relatórios e baixar o arquivo `LivroFiscal.xlsx` da empresa alvo no período desejado.
- **Processamento de Dados**: Utilizando a biblioteca `pandas`, o script irá ler cada `LivroFiscal.xlsx` baixado (ignorando o cabeçalho inicial de 5 linhas), e calcular:
  - Total de Notas Ativas e Canceladas.
  - Soma do `VALOR NF` (Faturamento total bruto).
  - Soma dos Tributos Federais Retidos (`VALOR COFINS`, `CSLL`, `INSS`, `IR`, `PIS`).
  - Soma do `VALOR ISS` retido/devido.
  - Soma do `VALOR LÍQUIDO`.
- **Exportação Consolidadora**: O script irá alimentar uma **Planilha Mestre Consolidada** adicionando uma linha com o Resumo Financeiro de cada empresa conforme o robô avança a lista. Os arquivos gerados originalmente serão salvos numa pasta de "Arquivos Originais".

---
## User Review Required
> [!IMPORTANT]
> A forma como vamos capturar a demonstração do seu uso precisa ser definida. (Vide opções passadas no chat).
