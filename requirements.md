# GestorSefaz - Requirements

## Visão Geral
Módulo complementar ao ecossistema (baseado no GestorPrefeitura) denominado **GestorSefaz**. O objetivo é acessar o portal e-Serviços da Sefaz/MS (`https://eservicos.sefaz.ms.gov.br/`), validar o acesso e realizar a coleta de Notas Fiscais Eletrônicas (NF-e) e Conhecimentos de Transporte Eletrônicos (CT-e), tanto emitidos quanto recebidos/tomados, com base em uma lista de empresas. Os resultados deverão ser estruturados e consolidados em arquivos Excel individuais para cada empresa.

## Regras de Negócio (Core)
1. **Autenticação Híbrida**: O usuário selecionará o Certificado Digital manualmente. O bot assumirá após a tela inicial carregar.
2. **Validação de Segurança**: O perfil logado DEVE ser obrigatoriamente "LUZEMIR MARTINS BARBOSA". Se for diferente, abortar o processo e exibir um alerta: "Verifique o certificado digital e tente novamente".
3. **Mapeamento de Empresas**: A lista de entrada conterá as colunas obrigatórias: `Código Domínio`, `Nome Empresa`, `CNPJ` e `Inscrição Estadual`.
4. **Navegação e Extração**:
   - Fechar pop-ups iniciais da home (clicando em OK).
   - Selecionar Perfil "Procurador" e inserir a Inscrição Estadual (com máscara, ex: 28.333.170-4) em "Representando(a)", clicando no item da lista.
   - **NF-e (Emitidas)**: Ir em "Informações Fiscais" -> "NF-e". Preencher período (dia 01 ao último dia do mês). Consultar por "Emitente". Extrair dados de todas as páginas (150 itens/pág). Adicionar linha de totais (valores, e contagem de notas ativas/canceladas). Informar falhas na sequência de notas emitidas.
   - **NF-e (Recebidas)**: Refazer período, consultar por "Destinatário". Extrair dados e calcular totais e quantidades.
   - **CT-e (Emitente)**: Ir em opções -> "Documentos Fiscais Eletronicos" -> "Conhecimento de Transporte Eletronico". Consultar "Emitente". Copiar dados.
   - **CT-e (Tomador)**: Consultar "Tomador". Copiar dados.
5. **Geração de Relatórios**: Para cada empresa processada, gerar um arquivo Excel: `{mmaaaa}_sefaz_{nome_empresa}.xlsx` (ex: `102023_sefaz_empresa.xlsx`). O arquivo deve conter obrigatoriamente 4 abas (Emitidas, Recebidas, CTE Emitente, CTE Tomador).
   - *Exceção:* Caso não haja registros para alguma das consultas, a aba deve ser criada e preenchida com a mensagem literal "A Consulta não Retornou Registros".
6. **Log e Acompanhamento**: Manter arquivo de log em texto para registro e controle de erros, usando a estrutura estabelecida.

## O que o sistema NÃO fará
- Não utilizará banco de dados (o armazenamento será estritamente em arquivos Excel gerados dinamicamente).
- Não fará a digitação de senhas de certificado (processo de login será obrigatoriamente híbrido).
- Não implementará roles/permissões internas do app (o controle e permissionamento baseiam-se unicamente no certificado logado).
