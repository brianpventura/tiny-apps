# tiny-apps

Ferramentas em Python sobre a **API v3 do [Tiny ERP](https://tiny.com.br)**:
automação de cadastro (custo, preço, categoria, estoque, contatos), relatórios e
um **pipeline ETL** que carrega os pedidos de venda num data warehouse MySQL
para análise.

Multi-empresa: cada empresa (tenant) tem as suas próprias credenciais de API, e
todos os scripts trabalham com uma, com várias ou com todas.

> [!WARNING]
> Vários scripts **escrevem no ERP de produção**, alguns de forma
> irreversível (um lançamento de estoque, por exemplo, só se corrige com outro
> lançamento). O padrão é sempre o modo seguro (`--dry-run`, ou a ausência de
> `--commit` / `--executar`): rode assim primeiro e confira o preview antes de
> gravar.

## Requisitos

- Python **3.14**
- Uma conta no Tiny com permissão para criar aplicativos de API (uma por empresa)
- Para o pipeline ETL: **MySQL 8**

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Rode sempre a partir da **raiz** do repositório. No console do Windows, defina
`$env:PYTHONUTF8="1"` para evitar erro de encoding com acentos.

```powershell
$env:PYTHONUTF8="1"
.venv\Scripts\python.exe src\apps\<grupo>\<script>.py [args]
```

## Configuração

Dois tipos de arquivo, em lugares diferentes, nenhum deles versionado aqui:

- **Segredos e estado** (credenciais, tokens, caches) ficam sempre em `data/`.
- **Configuração de negócio** (as regras de cada empresa) fica em `data/` ou na
  pasta apontada pela variável de ambiente **`TINY_APPS_CONFIG`**, que pode ser a
  pasta `config/` de um repositório privado da empresa (ver
  [Estendendo para a sua empresa](#estendendo-para-a-sua-empresa)).

O repositório traz os modelos `data/*.exemplo.json`: copie cada um sem o
`.exemplo` e preencha.

| Arquivo | Onde | Para quê | Obrigatório |
|---|---|---|---|
| `tenants.json` | `data/` | Credenciais OAuth e dados de cada empresa | sim |
| `db.json` | `data/` | Conexão com o MySQL do warehouse (ou variáveis `TINY_DB_*`) | só para o `pipeline_etl` |
| `regras.json` | config | Perfil de regras de venda da empresa (ver [Regras de negócio](#regras-de-negócio)) | não: sem ele vale o padrão genérico |
| `verificacao.json` | config | Números de regressão e casos aceitos da fase 5 do pipeline | não: sem ele a regressão é pulada |
| `linhas_produto.json` | config | Linhas de produto do `relatorio_linhas.py` | só para esse relatório |

### 1. Criar o aplicativo no Tiny

Em cada conta do Tiny: **Configurações → Aplicativos → criar aplicativo**, com
URL de redirecionamento `http://localhost:8765/callback`. Anote o
`client_id` e o `client_secret`.

### 2. Preencher `data/tenants.json`

```json
{
  "empresa_a": {
    "nome": "Empresa A",
    "client_id": "SEU_CLIENT_ID",
    "client_secret": "SEU_CLIENT_SECRET",
    "redirect_uri": "http://localhost:8765/callback",
    "razao_social": "EMPRESA A COMERCIO LTDA",
    "clientes_excluidos": ["EMPRESA B COMERCIO LTDA"]
  }
}
```

- A **chave** (`empresa_a`) é o *tenant_id* que você passa na linha de comando.
- `nome` é só o rótulo de exibição.
- `razao_social` e `clientes_excluidos` servem para tirar do faturamento as
  vendas entre empresas do mesmo grupo (transferências internas). O match é
  por prefixo, porque o Tiny trunca a razão social.

### 3. Autenticar

```powershell
.venv\Scripts\python.exe src\apps\auth\_reauth.py empresa_a
```

O script abre o navegador na tela de login do Tiny e captura o código de
autorização num servidor local em `localhost:8765/callback`. Os tokens são
gravados em `data/tokens_<tenant>.json`.

- O `access_token` se renova sozinho. O **`refresh_token` expira com ~24h sem
  uso**: quando isso acontece, os scripts param com `TokenExpiredError` e basta
  rodar o `_reauth.py` de novo.
- **Cuidado com a conta errada.** O Tiny grava o token mesmo se você logar na
  empresa errada, e o token não diz de qual empresa é. O `_reauth.py` força a
  tela de login (`prompt=login`) justamente para não herdar uma sessão aberta.
  Depois de autenticar, confira um dado conhecido da empresa, como a
  quantidade de depósitos.

### 4. (Opcional) MySQL para o pipeline

Crie um usuário próprio para o warehouse, sem usar o `root`:

```sql
CREATE USER 'etl'@'localhost' IDENTIFIED BY 'troque-esta-senha';
CREATE DATABASE tiny_dw CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci;
GRANT ALL PRIVILEGES ON tiny_dw.* TO 'etl'@'localhost';
```

Preencha `data/db.json` a partir de `data/db.exemplo.json` e crie o esquema:

```powershell
.venv\Scripts\python.exe src\apps\pipeline_etl\pipeline_pedidos.py --fases 0
```

## Apps

Cada script documenta o próprio uso e as regras que aplica na docstring do
topo (seção `Uso:`).

### Warehouse de pedidos: `src/apps/pipeline_etl/`

Carrega `GET /pedidos/{id}` num MySQL local (`tiny_dw`). Guarda o JSON cru de
cada pedido (`pedido_raw`) e as tabelas normalizadas (`pedido`, `pedido_item`,
`dim_cliente`, `dim_produto`...). As regras de negócio viram **colunas e views**
em vez de filtro na carga, então mudar uma regra custa zero requisição.

| Script | O que faz |
|---|---|
| `pipeline_pedidos.py` | Orquestrador. Fases: 0 esquema · 1 listagem · 2 detalhe · 3 catálogo · 4 transform+load (offline) · 5 verificação (offline) |
| `relatorio_produtos_canal.py` | Top produtos de uma integração/canal numa janela, com diagnóstico |
| `relatorio_linhas.py` | Vendas por SKU × mês e clientes por linha de produto |
| `relatorio_recorrencia_produtos.py` | Recorrência de compra por SKU, com estoque atual (lê a API só para o estoque) |
| `_teste_offline.py` | Autoteste sem banco e sem API |

```powershell
# carga de um período
.venv\Scripts\python.exe src\apps\pipeline_etl\pipeline_pedidos.py --inicio 2026-01-01 --fim 2026-01-31
# atualização do dia a dia: re-pagina os últimos N dias e só detalha o que mudou
.venv\Scripts\python.exe src\apps\pipeline_etl\pipeline_pedidos.py --incremental --dias-retro 15
# conferência (zero requisição)
.venv\Scripts\python.exe src\apps\pipeline_etl\pipeline_pedidos.py --so-verificar
```

Códigos de saída: `0` ok · `1` token inválido na largada · `2` token expirou
no meio (retome com `--fases 2,3,4,5`) · `3` outro processo do mesmo tenant
rodando · `4` verificação reprovou · `5` o perfil de regras mudou e faltou
`--aplicar-regras`.

As views para consulta direta são `v_pedido` e `v_venda_item` (venda válida:
situação válida, natureza de venda, sem transferência interna), além de
`v_resumo_canal`, `v_resumo_integracao` e as de reconciliação `v_recon_*`, que
devem voltar vazias.

### Escrita no ERP

Todos seguem o mesmo fluxo: lê a planilha → faz GET do objeto completo → altera
**só** o campo alvo → mostra o preview → pede confirmação → PUT/POST → grava
um log CSV em `reports/` (que serve de backup para reverter).

| Script | O que faz |
|---|---|
| `gerenciar_produtos/atualizar_custo.py` | Preço de custo (e opcionalmente fornecedor) por SKU |
| `gerenciar_produtos/renomear_produtos.py` | Descrição do produto por SKU |
| `gerenciar_produtos/inativar_excluir.py` | Inativa ou exclui produtos |
| `precos/atualizar_precos.py` | Preço promocional ou de venda a partir de qualquer lista de preços |
| `custo_kits/custo_kits.py` | Custo do kit = soma do custo dos componentes |
| `padronizar_categoria/padronizar_categoria.py` | Categoria dos produtos |
| `estoque/balanco_estoque.py` | Balanço de estoque com lote/validade na observação (**irreversível**; modelo em `estoque/modelo_balanco_estoque.xlsx`) |
| `contatos/corrigir_telefones.py` | Normaliza o telefone dos contatos |
| `contatos/excluir_ocultos.py` | Exclusão lógica de contatos anonimizados por LGPD |

### Relatórios e listas (API)

| Script | O que faz |
|---|---|
| `relatorios/precos_promocionais.py` | Snapshot de preço × promocional de todo o catálogo (~1 req por 100 produtos) |
| `relatorios/relatorio_produtos.py` | SKU, descrição, NCM e EAN de todos os produtos |
| `contatos/enriquecer_lista_contatos.py` | Anexa email e telefone a qualquer planilha com coluna de CPF/CNPJ |

## Regras de negócio

O que conta como **venda válida** e em que **canal** cada pedido cai muda de uma
empresa para outra. Por isso o código (`src/core/regras_vendas.py`) só tem o
**mecanismo**, e a **política** vem de um perfil:

1. `config/regras.padrao.json`: o padrão genérico, versionado aqui:
   - **situações válidas** {Faturada, Aprovada, Preparando envio, Pronto para
     envio, Enviada, Entregue};
   - **naturezas de venda**: toda natureza cujo nome começa com "venda" (os
     nomes são criados em cada conta do Tiny);
   - **canais**: Site, Marketplace, Vendas internas, Feiras/Eventos e Outros,
     reconhecidos pelo nome da integração e pelo vendedor.
2. `regras.json` da empresa, na pasta de configuração, com **só o que muda**.
   Ele aceita ajustes por empresa (`tenants.<id>`) e recortes por relatório
   (`relatorios.<nome>`). Modelo comentado em `data/regras.exemplo.json`.

Todos os apps leem o mesmo perfil; nenhum define regra própria. No warehouse,
as regras viram colunas (`eh_situacao_valida`, `eh_natureza_alvo`, `canal`)
calculadas sobre o JSON cru, que é guardado inteiro. Mudar a regra, portanto,
não exige recoletar nada. Quando o perfil muda, a fase 4 do pipeline **para**
e mostra a diferença; `--aplicar-regras` reprocessa o warehouse offline. Isso
evita que uma troca acidental (a variável `TINY_APPS_CONFIG` esquecida, por
exemplo) reescreva os dados com a regra errada.

## Estendendo para a sua empresa

O que é específico de uma empresa não entra neste repositório. Nele fica o
núcleo genérico. A empresa mantém um **repositório privado** ao lado, com a
configuração dela e as extensões sob medida:

```
tiny-apps/              este repositório (núcleo)
tiny-apps-<empresa>/    privado
  config/               regras.json, linhas_produto.json, verificacao.json
  extensoes/<grupo>/    scripts próprios, que importam core/ deste repositório
```

- Aponte `TINY_APPS_CONFIG` para `tiny-apps-<empresa>/config`.
- Nas extensões, localize o núcleo pela variável `TINY_APPS_SRC` ou pela pasta
  irmã `../tiny-apps/src`, e use `core.regras_vendas.perfil(tenant, "<relatorio>")`
  para as regras.
- Se uma extensão se provar útil para qualquer empresa, generalize e traga
  para cá.

## Estrutura

```
src/core/          camada compartilhada: auth OAuth, cliente HTTP, rate limiter,
                   mecanismo das regras de venda (regras_vendas.py), config.py
config/            padrões genéricos (regras.padrao.json)
src/apps/<grupo>/  scripts por domínio
data/              configuração, tokens e caches -- NÃO versionado (só *.exemplo.json)
reports/           saída dos scripts (xlsx/csv) -- NÃO versionado
scripts/           ferramentas do repositório (verificar_sensiveis.py)
```

- **Sem instalação de pacote.** Cada script sobe o `sys.path` a partir da
  própria pasta até achar o diretório que contém `core/`, e só então importa.
- **Rate limit.** Todo HTTP passa por `core/tiny_client.py`, que respeita
  2 req/s por processo (`core/rate_limiter.py`) e refaz em 429/5xx. É o
  limiter que controla a taxa, não o tamanho do pool de threads. Empresas
  diferentes podem rodar em processos paralelos.
- **Regras de venda num lugar só.** Ver [Regras de negócio](#regras-de-negócio).

## Comportamentos da API v3 que o código já trata

Descobertos na prática. Vale conhecer antes de mexer:

- **Filtro de data em `/pedidos`**: `dataInicial`/`dataFinal` funcionam no
  servidor; `dataInicio`/`dataFim` são **aceitos e ignorados** (a API responde
  200 com tudo). Pagine por janela mensal: offset profundo (dezenas de
  milhares) devolve `400 "Sua consulta levou muito tempo"`. Esse mesmo 400
  aparece às vezes em offset raso, de forma intermitente: tente de novo antes
  de concluir qualquer coisa.
- **Não há filtro por produto em `/pedidos`.** Saber quem comprou um SKU exige
  o detalhe de todos os pedidos do período, e é por isso que o warehouse
  existe.
- **Natureza de operação e itens só vêm no detalhe** `GET /pedidos/{id}`; a
  listagem traz situação, valor, vendedor, integração e o `cliente` completo.
- **Kits já vêm explodidos** nos itens do pedido. O `tipo` do item (`P`) não
  serve para detectar kit: use o `tipo == "K"` do cadastro.
- **Não existe desconto por item**: `valorUnitario` é preço bruto, às vezes
  com 4 casas por causa de desconto rateado pelo integrador.
- **Produtos**: a listagem devolve `codigo=null` mas `sku` preenchido; o mesmo
  SKU costuma existir duas vezes (vivo + gêmeo `situacao="E"`). Variação só
  aceita `PUT /produtos/{pai}/variacoes/{id}`. O PUT exige
  `fornecedores[].padrao`, que o GET não devolve
  (`produtos_utils.sanitizar_payload_produto`).
- **Contatos**: não há DELETE, só exclusão lógica (`situacao="E"`). O PUT exige
  `nome`, e `tipos` precisa ser lista de ids inteiros. Na listagem, `celular`
  vem `null`; ele só aparece no detalhe.
- **Estoque**: `POST /estoque/{id}` só aceita data `YYYY-MM-DD HH:MM:SS`;
  **HTTP 204 = nada foi gravado**; `observacoes` é cortado em 100 caracteres
  sem aviso. Não existe histórico de estoque na API, só o saldo atual.
- **Marketplaces não repassam o contato do comprador** ao Tiny: nome e
  endereço vêm, email e telefone não.

## Segurança e dados pessoais

- `data/`, `reports/`, planilhas, CSVs e PDFs estão no `.gitignore`. Eles
  guardam segredos e dados pessoais (LGPD) e **não devem ir para o
  repositório**.
- Antes de cada commit, rode o verificador:

  ```powershell
  .venv\Scripts\python.exe scripts\verificar_sensiveis.py
  ```

  Ele procura CPF, CNPJ, telefone, email, JWT, `client_secret` preenchido,
  caminho de usuário do Windows, valores em R$ e ids longos nos arquivos que
  iriam para o git. Também procura termos proibidos de uma lista local em
  `_privado/termos_proibidos.txt` (nomes de empresas, pessoas etc., um por
  linha; pasta ignorada pelo git). Para rodar automaticamente, instale como
  hook em `.git/hooks/pre-commit`:

  ```sh
  #!/bin/sh
  exec .venv/Scripts/python.exe scripts/verificar_sensiveis.py
  ```

## Convenções

- Código, docstrings e textos de CLI em português; identificadores e saída de
  console em ASCII sem acento.
- Cada script leva a própria documentação na docstring do topo. Ao adicionar
  uma flag ou mudar uma regra, atualize a docstring junto.
- Saída em `reports/`; configuração, caches e tokens em `data/`.
