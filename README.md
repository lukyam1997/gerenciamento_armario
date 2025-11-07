# LockNAC - Gerenciamento de Armários (Google Sheets)

## Visão geral
LockNAC é um sistema de controle de armários físicos voltado para visitantes e acompanhantes em unidades de saúde. A solução combina uma interface web moderna (HTML, CSS e JavaScript) hospedada como um Web App do Google Apps Script com uma base de dados inteiramente gerenciada em planilhas do Google Sheets. Todo o estado do sistema (cadastro de armários, movimentações, unidades, usuários, termos de responsabilidade e logs) é persistido no arquivo **Armários - NAC.xlsx**, permitindo operação simplificada, versionamento por planilha e auditoria diretamente no Google Drive.

## Arquitetura do sistema
- **Google Sheets**: funciona como banco de dados do sistema. Cada aba mantém um domínio de informação (Visitantes, Acompanhantes, Cadastro Armários, Usuários, Unidades, Cadastro de setores, Termos, Movimentações, Logs e Notificações). Todas as leituras e escritas de dados partem desse arquivo.
- **Apps Script (`Code.gs`)**: camada de aplicação e API. Expõe as rotas do Web App (`doGet` e `doPost`), controla cache com `CacheService`, registra eventos com `registrarLog`, trata autenticação e encapsula as regras de negócio de armários, usuários, unidades, termos e movimentações.
- **Frontend (`index.html`)**: interface de usuário responsiva, com navegação lateral, dashboards e componentes interativos (SweetAlert para feedback, FontAwesome para ícones, layout responsivo com CSS customizado). Comunicação com a API é feita via `google.script.run` e requisições POST que enviam a propriedade `action` para o Apps Script.

## Fluxo de funcionamento
1. O usuário acessa a URL publicada do Web App (`doGet`) e recebe o conteúdo de `index.html` com scripts incorporados.
2. A interface carrega dados iniciais chamando `verificarInicializacao`. Caso a planilha esteja vazia, `inicializarPlanilha()` cria as abas básicas e popula registros padrão.
3. Cada ação do usuário (ex.: liberar armário, registrar termo, cadastrar unidade ou usuário) aciona `google.script.run` ou `fetch` com um `action` específico. O Apps Script direciona a solicitação em `handlePost`.
4. As funções de domínio (`getArmarios`, `cadastrarArmario`, `getUsuarios`, `salvarTermoCompleto`, etc.) leem ou escrevem na planilha, atualizam caches e registram logs operacionais.
5. As respostas JSON retornadas pelo Apps Script atualizam a UI em tempo real.

## Estrutura recomendada da planilha
Crie/valide as abas abaixo no arquivo Google Sheets. Os cabeçalhos são lidos dinamicamente por `obterEstruturaPlanilha`, mas seguir os nomes sugeridos garante compatibilidade.

| Aba | Colunas principais | Observações |
| --- | --- | --- |
| **Visitantes** | `id`, `número`, `status`, `nome visitante`, `nome paciente`, `leito`, `volumes`, `hora início`, `hora prevista`, `data registro`, `unidade`, `termo aplicado`, `whatsapp` | Controla armários destinados a visitantes. `status` aceita valores `livre`, `em-uso`, `próximo`, `vencido`. |
| **Acompanhantes** | `id`, `número`, `status`, `nome acompanhante`, `nome paciente`, `leito`, `volumes`, `hora início`, `data registro`, `whatsapp`, `unidade`, `termo aplicado` | Estrutura similar à aba de visitantes, sem `hora prevista`. |
| **Cadastro Armários** | `id`, `número`, `tipo`, `unidade`, `localização`, `status`, `data cadastro` | Mantém o catálogo físico. Função `cadastrarArmarioFisico` atualiza esta aba. |
| **Usuários** | `id`, `nome`, `login`, `perfil`, `ativo`, `podeGerenciar`, `dataCadastro`, `status`, `senha`, `unidades` | Utilizada por `autenticarUsuario`, `cadastrarUsuario`, `atualizarUsuario` e `excluirUsuario`. |
| **Unidades** | `id`, `nome`, `status`, `dataCadastro` | Manipulada por `getUnidades`, `cadastrarUnidade` e `alternarStatusUnidade`. |
| **Cadastro** | `setor` | Lista de setores exibidos em combos pelo front-end (`getSetores`). |
| **Termos** | Campos estruturados para responsaveis, datas, status, links de PDF e assinaturas | Preenchida por `salvarTermoCompleto`, `finalizarTermo`, `getTermo`. |
| **Movimentações** | `id`, `armarioId`, `tipoMovimentacao`, `responsavel`, `horario`, `observacoes`, etc. | Registrada via `salvarMovimentacao` e consultada por `getMovimentacoes`. |
| **Logs** | `timestamp`, `tipo`, `mensagem`, `detalhes` | Recebe todos os eventos de `registrarLog`, útil para auditoria.
| **Notificações** | Estrutura flexível com campos `titulo`, `descricao`, `data`, `lido` | Controlada por `getNotificacoes`.

> 💡 A função `adicionarDadosIniciais` popula as abas **Cadastro Armários**, **Usuários** e **Unidades** com registros padrões caso estejam vazias, facilitando o primeiro uso.

## API do Apps Script (`handlePost`)
Todas as chamadas POST devem enviar os parâmetros `action=<nomeDaAcao>` e outros campos esperados. Principais ações disponíveis:

- `getArmarios` – Lista armários por tipo (`visitante`, `acompanhante`, `admin`, `ambos`).
- `cadastrarArmario` / `liberarArmario` – Registra nova ocupação ou libera armário existente.
- `getUsuarios`, `cadastrarUsuario`, `atualizarUsuario`, `excluirUsuario`, `autenticarUsuario` – Gestão completa de usuários e perfis.
- `getLogs` – Consulta trilhas de auditoria.
- `getNotificacoes` – Recupera avisos operacionais armazenados em planilha.
- `getEstatisticasDashboard` – Calcula indicadores consolidados para os cards do dashboard.
- `getHistorico` – Obtém histórico de uso por tipo de armário.
- `getCadastroArmarios`, `cadastrarArmarioFisico` – Mantém o inventário físico de armários.
- `getUnidades`, `getSetores`, `cadastrarUnidade`, `alternarStatusUnidade` – Administração de unidades/setores.
- `salvarTermoCompleto`, `finalizarTermo`, `getTermo` – Fluxo de termos de responsabilidade, incluindo geração de PDFs e controle de assinaturas.
- `getMovimentacoes`, `salvarMovimentacao` – Registro detalhado de movimentações associadas a cada armário.
- `verificarInicializacao`, `inicializarPlanilha` – Utilidades para preparar a base quando o sistema é publicado pela primeira vez.

Qualquer ação não reconhecida retorna `{ success: false, error: 'Ação não reconhecida: <action>' }`, permitindo validação no front-end.

## Frontend (`index.html`)
- **Layout responsivo** com sidebar fixa, cards analíticos e tabelas com indicadores de status (cores e badges).
- **Componentização**: blocos de UI reusáveis para dashboard, listagem de armários, histórico, gerenciamento de usuários e unidades.
- **Feedback**: uso de SweetAlert 2 para diálogos de confirmação/erro e notificações sutis.
- **Acessibilidade**: contraste alto, suporte a teclado (focus rings), tipografia padrão `Inter`.
- **Integrações**: scripts fazem chamadas assíncronas ao Apps Script para atualizar dados sem recarregar a página.

## Pré-requisitos e permissões
1. Conta Google com acesso ao Google Drive e Apps Script.
2. Planilha `Armários - NAC.xlsx` armazenada no mesmo Drive do projeto.
3. Permissões do Web App publicadas para "Qualquer pessoa com o link" (ou restrição desejada) com execução como "Você (proprietário)" para garantir acesso às abas protegidas.
4. Serviços avançados não são obrigatórios, apenas `SpreadsheetApp`, `HtmlService`, `ContentService`, `CacheService`, `Session` e `UrlFetchApp` (para integrações opcionais) presentes no Apps Script padrão.

## Passo a passo de implantação
1. **Criar projeto Apps Script**: abra o Google Drive, crie um novo Apps Script e conecte-o à planilha `Armários - NAC.xlsx` (Arquivo → Gerenciar versões → Vincular à planilha existente).
2. **Importar código**:
   - Substitua o conteúdo padrão do arquivo `Code.gs` pelo script deste repositório.
   - Adicione um arquivo HTML chamado `index` e cole o conteúdo completo de `index.html`.
3. **Salvar e testar**: execute `verificarInicializacao` ou `inicializarPlanilha` no editor do Apps Script para validar permissões e criar estruturas básicas.
4. **Publicar Web App**: em "Implantar" → "Implantações" → "Nova implantação", escolha "Aplicativo da web", defina "Executar como" = proprietário e selecione quem pode acessar. Salve a URL gerada.
5. **Configurar permissões de planilha**: garanta que as abas críticas estejam protegidas contra edição manual inadvertida. Ajuste filtros/validações conforme necessidade operacional.
6. **Distribuir acesso**: compartilhe a URL apenas com os colaboradores autorizados e configure acessos de edição na planilha de acordo com os perfis cadastrados no sistema.

## Manutenção e operação
- Utilize `limparCacheArmarios`, `limparCacheUnidades`, `limparCacheTermos` (já previstas no script) após alterações massivas para forçar recarregamento de dados na UI.
- Consulte a aba **Logs** regularmente para identificar falhas (`ERRO`, `AVISO_CACHE`, etc.).
- Periodicamente exporte a planilha ou utilize versões do Google Sheets como backup.
- Atualize o front-end via Apps Script para entregar novas funcionalidades sem alterar a URL publicada (basta criar nova implantação com versão mais recente).

## Testes recomendados antes de liberar
1. **Inicialização**: executar `verificarInicializacao` e confirmar que abas obrigatórias são criadas.
2. **Cadastro de armário**: adicionar armário visitante e acompanhante e verificar atualização imediata na interface.
3. **Liberação**: ocupar armário e em seguida liberar, garantindo registro na aba de movimentações.
4. **Gestão de usuários**: criar, editar e desativar usuário; testar login com perfis diferentes.
5. **Termo de responsabilidade**: gerar, salvar e finalizar um termo, confirmando o status e eventual link de PDF.
6. **Unidades/setores**: cadastrar nova unidade, alternar status e checar se aparece nas listas.
7. **Logs**: validar se cada operação gera registro com `success: true` ou mensagem de erro apropriada.

## Segurança e privacidade
- Execute o Web App com o menor conjunto possível de colaboradores. As credenciais dos usuários ficam na aba **Usuários**; limite o acesso à planilha.
- O script sanitiza entradas (`normalizarTextoBasico`, `converterParaBoolean`, validações de JSON) para mitigar dados inconsistentes, mas mantenha validações no front-end.
- Não armazene informações sensíveis (ex.: documentos pessoais) sem consentimento expresso. Caso necessário, aplique criptografia ou remova dados após o uso.
- Ative auditoria: mantenha a aba **Logs** protegida e monitore alterações suspeitas.

Com este README, qualquer administrador consegue entender, implantar e manter o sistema LockNAC integrado ao Google Sheets de ponta a ponta.
