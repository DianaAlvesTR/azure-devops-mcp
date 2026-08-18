# Alterações locais sobre o `microsoft/azure-devops-mcp` (branch `main`, upstream intacto)

Resumo de tudo que foi adicionado neste fork local, para servir de base ao PR contra o
repositório oficial. Nenhuma ferramenta existente foi removida ou teve seu comportamento
anterior quebrado — todas as mudanças são aditivas (novas actions em enums já existentes, ou
uma ferramenta nova).

Arquivos alterados: `src/tools/work-items.ts`, `src/tools/test-plans.ts`,
`test/src/tools/work-items.test.ts`, `test/src/tools/test-plan.test.ts`, `docs/TOOLSET.md`.

Status: `npx tsc --noEmit` limpo, `npx eslint` limpo, `npm run build` limpo,
`npx jest` → 1243/1243 testes passando.

---

## 1. Nova ferramenta `wit_query_write` (`src/tools/work-items.ts`)

Motivação: o MCP oficial só tinha `wit_query` (leitura). Não havia forma de criar/editar/apagar
uma **Shared Query** (pasta `Shared Queries`) — necessário para o fluxo de regressão, que usa
Shared Queries num Dashboard.

Actions:

- `create` — `WorkItemTrackingApi.createQuery(postedQuery, project, queryPath, validateWiqlOnly)`.
  Parâmetros: `query` (caminho da pasta pai), `name`, `wiql` (obrigatório a menos que
  `isFolder: true`), `isFolder`, `validateWiqlOnly`.
- `update` — `WorkItemTrackingApi.updateQuery(queryUpdate, project, queryPath, undeleteDescendants)`.
  Aceita `wiql` (nova WIQL), `newName` (rename), `newPath` (mover de pasta). Exige pelo menos um
  desses três — senão retorna erro de validação amigável.
- `delete` — `WorkItemTrackingApi.deleteQuery(project, queryPath)`.

Todas as 3 actions resolvem o projeto via o padrão `elicitProject` já usado no resto do arquivo
(prompt interativo se `project` não for passado), e mapeiam erros por action num
`Record<string, string>`.

## 2. `testplan` (leitura) — nova action `get_suite` (`src/tools/test-plans.ts`)

Motivação: nenhuma action existente devolvia o campo `queryString` de uma suíte — que é onde
fica embutida a WIQL de uma suíte dinâmica/baseada-em-requisito (ex: filtro de Target Release).
Sem isso não dá pra ler o valor atual antes de decidir o que trocar.

- Nova action `get_suite`, exige `planId` + `suiteId`.
- Chama diretamente `GET .../_apis/testplan/Plans/{planId}/Suites/{suiteId}` via `fetch` (o SDK
  `TestPlanApi` não expõe um "get single suite" que devolva o objeto completo com `queryString`).
- Erro tratado com o mesmo padrão das outras actions (`Error getting test suite: ...`).

## 3. `testplan_test_plan_write` — novas actions `clone` e `get_clone_status`

Motivação: o processo manual do time era sempre "copiar o Test Plan anterior (sem duplicar Test
Cases) e trocar o Target Release nas queries" — usando o botão **Copy test plan** da UI do ADO.
Não havia equivalente programático. A action `create` continua existindo e inalterada
(cria plano vazio).

### `clone`

Parâmetros novos: `sourcePlanId` (obrigatório), `sourceSuiteIds` (opcional — ver abaixo),
`duplicateTestCases` (default `false`), além dos campos já existentes de destino
(`name`, `iteration`, `description`, `startDate`, `endDate`, `areaPath`).

Chama `TestPlanApi.cloneTestPlan(cloneRequestBody, project, duplicateTestCases)` com:

```ts
cloneOptions: { copyAllSuites: false, copyAncestorHierarchy: true }
```

**Por que `copyAllSuites` é sempre `false`, nunca exposto como opção:** testado empiricamente —
`copyAllSuites: true` pode devolver o plano de destino com **apenas a suíte raiz e zero Test
Cases**, sem erro nenhum (`state: "succeeded"` reportado depois, mas `totalTestCasesCount: 0`).
O que funciona de forma confiável é listar explicitamente cada `suiteId` a clonar em
`sourceTestPlan.suiteIds`.

**Descoberta automática de suítes (`fetchDescendantSuiteIds`)**: a API de clone **não recursiona**
— se você passar o ID de uma suíte-pai em `suiteIds`, os filhos dela não são incluídos
automaticamente; é preciso listar cada suíte aninhada individualmente, senão a suíte-pai é clonada
vazia (sem filhos, sem Test Cases). Passar isso manualmente é frágil e fácil de errar, então:

- Se `sourceSuiteIds` for omitido (recomendado), uma função auxiliar
  (`fetchDescendantSuiteIds`) pagina `GET .../testplan/Plans/{planId}/Suites` (respeitando
  `continuationToken`) e retorna o ID de toda suíte que tenha `parentSuite` (ou seja, todas menos
  a raiz do próprio plano).
- Se `sourceSuiteIds` for passado explicitamente, é usado como está (documentado na descrição do
  campo que o chamador é responsável por incluir toda a árvore).

`duplicateTestCases` mapeia para o parâmetro `deepClone` do SDK — `false` corresponde a
"Reference existing test cases" na UI (comportamento manual padrão do time), `true` a
"Duplicate existing test cases".

### `get_clone_status`

Parâmetro: `cloneOperationId` (aceita `0`, que é um ID de operação válido — o schema usa
`.min(0)` e o guard checa `=== undefined`, não falsy, exatamente por causa disso).

**Por que não usa `TestPlanApi.getCloneInformation()`:** esse método do SDK chama um endpoint da
área `testplan` que devolve `"The requested resource does not support http method 'GET'"` — não
funciona para status de clone de **plano** (só de suíte, aparentemente). O endpoint que realmente
funciona está na área mais antiga `test`:
`GET .../_apis/test/cloneoperation/{id}?$includeDetails=true&api-version=5.0-preview.2`
(confirmado contra a documentação da Microsoft Learn e testado contra o ADO real). A ferramenta
chama isso diretamente via `connection.rest.get(...)`, contornando o método quebrado do SDK.

## 4. `testplan_test_suite_write` — nova action `update`, `create` estendida

### `create` (estendida, comportamento antigo preservado)

Novo parâmetro opcional `queryString`. Se fornecido, `suiteType` passa de `2` (StaticTestSuite)
para `1` (DynamicTestSuite) e a query é enviada junto na criação.

**Limitação conhecida, documentada e não perseguida:** criar uma suíte dinâmica nova do zero por
essa rota falhou contra o ADO real com `"The value for the QueryString property is not within the
permissible values for it."` — parece ser uma limitação real da API pública do ADO para criação
(distinta de update, que funciona). Não é um bug deste código; ficou documentado como limitação
conhecida em vez de continuar sendo investigado.

### `update` (nova)

Parâmetros: `planId`, `suiteId` (obrigatórios), `name` (opcional — rename) e/ou `queryString`
(opcional — trocar a WIQL da suíte dinâmica). Exige pelo menos um dos dois.

Chama `TestPlanApi.updateTestSuite(testSuiteUpdateParams, project, planId, suiteId)`. A API do
ADO exige `name` em `TestSuiteUpdateParams` mesmo quando só se quer trocar a query — então, se
`name` não for passado, a ferramenta primeiro busca o nome atual via
`TestPlanApi.getTestSuiteById()` e o reenvia, para não renomear a suíte sem querer.

**Validado contra o ADO real** neste teste: as 6 suítes dinâmicas de um Test Plan clonado
tiveram o Target Release da `queryString` trocado com sucesso (ex.: `26.2.1` → `26.2.5`), e uma
suíte teve a query substituída por um filtro que nunca retorna resultado
(`Source.[System.Id] in (-1)`) para "esvaziá-la" sem apagar a suíte.

---

## Cobertura de testes adicionada

- `test/src/tools/work-items.test.ts`: bloco `describe("wit_query_write tool")` com 5 casos
  (create, update, validação "pelo menos um campo", delete, erro em create).
- `test/src/tools/test-plan.test.ts`: blocos novos/reescritos —
  `update_test_suite tool`, `clone_test_plan tool` (auto-descoberta de suítes, bypass com IDs
  explícitos, paginação via `continuationToken`, header `User-Agent`, tratamento de erro),
  `get_clone_status tool` (mock de `connection.rest.get`), `get_suite tool` (sucesso, `planId`
  ausente, `suiteId` ausente, erro de API).

## Documentação

`docs/TOOLSET.md` atualizado com as novas linhas de tabela para `wit_query_write`,
`testplan` `get_suite`, `testplan_test_plan_write` `clone`/`get_clone_status`, e
`testplan_test_suite_write` `update`.
