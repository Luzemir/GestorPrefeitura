# GestorSefaz - Tasks List

## Fase 1: Setup da Interface e Estrutura
- [x] Task 1.1: Criar/Adaptar a Interface Gráfica (CustomTkinter) no `app_gui.py` adicionando a aba/função GestorSefaz (seleção de planilha de empresas e mês/ano de competência).
- [x] Task 1.2: Implementar a leitura do arquivo base de empresas garantindo as colunas `Código Domínio`, `Nome Empresa`, `CNPJ` e `Inscrição Estadual`.

## Fase 2: Automação Core - Autenticação e Segurança
- [x] Task 2.1: Criar script base do Playwright (`src/sefaz_bot.py`), inicializar browser em modo visível, navegar para `https://eservicos.sefaz.ms.gov.br/` e aguardar seleção manual de certificado.
- [x] Task 2.2: Adicionar rotina para aguardar e fechar popup inicial de aviso, se houver.
- [x] Task 2.3: Implementar validação de segurança verificando se o elemento com nome do usuário é igual a "LUZEMIR MARTINS BARBOSA". Rejeitar e mostrar popup do CTk caso contrário.

## Fase 3: Navegação do Bot - Mapeamento e Iteração
- [x] Task 3.1: Construir a lógica no loop principal para varrer cada empresa da lista lida na Task 1.2.
- [x] Task 3.2: Implementar a seleção "Tipo de Perfil -> Procurador" e digitar a "Inscrição Estadual" no campo "Representando(a)", disparando o click de seleção da dropdown correspondente.

## Fase 4: Extração de NF-e (Emitidas e Recebidas)
- [x] Task 4.1: Automatizar acesso ao submenu "Informações Fiscais" -> clique no botão popup "NF-e".
- [x] Task 4.2: Inserir datas (Primeiro ao Último dia da competência).
- [x] Task 4.3: Executar busca "Emitente". Extrair dados tabulares iterando pelas paginações ("Próximo"). 
- [x] Task 4.4: Executar busca "Destinatário" usando mesmo período. Extrair dados tabulares e paginar.

## Fase 5: Extração de CT-e (Emitente e Tomador)
- [ ] Task 5.1: Automatizar acesso ao menu -> "Documentos Fiscais Eletronicos" -> "Conhecimento de Transporte Eletronico".
- [ ] Task 5.2: Executar busca "Emitente" e extrair dados, iterando paginação.
- [ ] Task 5.3: Executar busca "Tomador" e extrair dados, iterando paginação.

## Fase 6: Processamento de Dados (Pandas)
- [ ] Task 6.1: Desenvolver função de auditoria: Formatar dados financeiros e calcular a linha extra com totais (somas e contagem de ativas/canceladas).
- [ ] Task 6.2: Desenvolver rotina de verificação de sequência lógica no número da NF-e Emitidas (verificar falhas/saltos de número) e adicionar alerta no resultado.

## Fase 7: Consolidação do Relatório (Excel)
- [ ] Task 7.1: Criar o gravador de Excel gerando o arquivo `{mmaaaa}_sefaz_{nome_empresa}.xlsx` com as abas `Emitidas`, `Recebidas`, `CTE Emitente` e `CTE Tomador`.
- [ ] Task 7.2: Adicionar fallback no gravador: caso a extração retorne vazio, inserir na aba a linha "A Consulta não Retornou Registros".
- [ ] Task 7.3: Conectar callbacks para atualizar os logs da GUI para que o usuário acompanhe o progresso de cada empresa em tempo real e testes End-to-End.
