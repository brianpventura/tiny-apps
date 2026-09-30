-- ==================================================================== --
--  tiny_dw -- warehouse local dos pedidos da API v3 do Tiny            --
--  MySQL 8.0.45 / InnoDB / utf8mb4                                     --
-- ==================================================================== --
--
-- Decisoes que valem para o arquivo inteiro:
--
--  * DECIMAL para dinheiro, nunca FLOAT. O teste de regressao fecha ao
--    centavo; com FLOAT nunca fecharia.
--  * DATETIME, nunca TIMESTAMP. TIMESTAMP converte para UTC conforme a
--    time_zone da SESSAO, entao a mesma linha leria diferente em sessoes
--    diferentes. A API do Tiny manda hora local sem offset.
--  * tenant e VARCHAR(32) com o tenant_id de data/tenants.json. Compoe a PK
--    de tudo, e a PK do InnoDB vai dentro de todo indice secundario: manter
--    os ids curtos. (Um ENUM economizaria bytes, mas amarraria o schema aos
--    nomes das empresas; se optar por ENUM, NUNCA reordenar a lista -- o
--    armazenamento e o indice ordinal. Acrescentar valor NO FIM e barato.)
--  * Collation _0900_ai_ci em quase tudo (ignora acento e caixa, o que da de
--    graca o _norm() no match de cliente do grupo). EXCECAO: sku usa
--    _0900_as_cs, senao SKUs que diferem so por caixa colidiriam.
--  * Todo write do pipeline e INSERT ... ON DUPLICATE KEY UPDATE. REPLACE
--    INTO e PROIBIDO: faz DELETE+INSERT e, com ON DELETE CASCADE, apagaria
--    silenciosamente todos os itens do pedido.
--  * Cada tabela upsertada tem UMA unica chave unica. Com mais de um indice
--    unico a ordem de verificacao do ODKU e indefinida e o UPDATE pode cair
--    na linha errada. Por isso NAO existe UNIQUE (tenant, numero_pedido) em
--    pedido -- fica indice nao-unico + view de reconciliacao.

-- -------------------------------------------------------------------- --
--  Controle de migracao                                                 --
-- -------------------------------------------------------------------- --

CREATE TABLE IF NOT EXISTS schema_migracao (
  arquivo      VARCHAR(120)  NOT NULL,
  sha256       CHAR(64)      NOT NULL COLLATE ascii_bin,
  aplicado_em  DATETIME(3)   NOT NULL,
  PRIMARY KEY (arquivo)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='Arquivos sql/ ja aplicados. O sha256 denuncia arquivo editado depois de aplicado.';

-- -------------------------------------------------------------------- --
--  Configuracao (semeada em Python a partir de data/tenants.json e das  --
--  constantes de core/regras_vendas.py -- ver etl_esquema.semear)       --
-- -------------------------------------------------------------------- --

CREATE TABLE IF NOT EXISTS cfg_tenant (
  tenant             VARCHAR(32) NOT NULL,
  empresa            VARCHAR(80)  NOT NULL,
  razao_social       VARCHAR(255) NOT NULL DEFAULT '',
  PRIMARY KEY (tenant)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE IF NOT EXISTS cfg_situacao (
  codigo     TINYINT     NOT NULL,
  nome       VARCHAR(32) NOT NULL,
  eh_valida  TINYINT(1)  NOT NULL DEFAULT 0,
  PRIMARY KEY (codigo)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='eh_valida = vendas.situacoes_validas do perfil de regras da empresa (core/regras_vendas.py)';

CREATE TABLE IF NOT EXISTS cfg_cliente_excluido (
  tenant     VARCHAR(32) NOT NULL,
  nome_norm  VARCHAR(160) NOT NULL,
  PRIMARY KEY (tenant, nome_norm)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='clientes_excluidos do tenants.json MAIS a razao_social do proprio tenant (senao a auto-venda entra como faturamento).';

-- -------------------------------------------------------------------- --
--  Staging append-only                                                  --
-- -------------------------------------------------------------------- --

CREATE TABLE IF NOT EXISTS pedido_raw (
  tenant        VARCHAR(32) NOT NULL,
  id_pedido     BIGINT       NOT NULL,
  fonte         ENUM('listagem','detalhe') NOT NULL,
  payload_hash  CHAR(64)     NOT NULL COLLATE ascii_bin,
  payload       JSON         NOT NULL,
  extraido_em   DATETIME(3)  NOT NULL,
  visto_em      DATETIME(3)  NOT NULL,
  -- Coluna JSON nao e indexavel, entao estas duas existem para poder indexar.
  --
  -- situacao_raw e GENERATED e isso e seguro: o guarda JSON_TYPE='INTEGER'
  -- garante que o CAST so roda sobre um inteiro de verdade, nunca sobre texto.
  --
  -- data_raw NAO e generated, e essa e a licao mais cara deste schema. Medido
  -- em 21/09/2026: STR_TO_DATE levanta erro 1411 e CAST(... AS DATE) levanta
  -- 1292 quando a expressao roda DENTRO de uma coluna gerada num INSERT --
  -- num SELECT o mesmo CAST so devolveria NULL, mas STRICT_TRANS_TABLES
  -- promove o warning a erro em statement que grava. Nao existe expressao SQL
  -- segura para isto: qualquer REGEXP ainda deixaria passar '2026-02-31'. E
  -- uma data torta num unico pedido derrubaria o INSERT do lote inteiro de
  -- 200, repetidamente. Quem converte e o Python (etl_transform.data), que ja
  -- trata ISO, DD/MM/YYYY, '', 'null', lixo e data impossivel devolvendo None.
  situacao_raw  SMALLINT GENERATED ALWAYS AS (
      CASE WHEN JSON_TYPE(JSON_EXTRACT(payload,'$.situacao')) = 'INTEGER'
           THEN CAST(payload->>'$.situacao' AS SIGNED) END) STORED,
  data_raw      DATE         NULL,
  PRIMARY KEY (tenant, id_pedido, fonte, payload_hash),
  KEY ix_raw_janela (tenant, fonte, data_raw),
  KEY ix_raw_versao (tenant, id_pedido, fonte, extraido_em)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='JSON cru e integro, append-only e versionado. Reprocessar o transform nao custa request. Hash SHA-256 canonico calculado em Python (o MySQL nao preserva ordem de chave em JSON).';

-- -------------------------------------------------------------------- --
--  Dimensoes (SCD tipo 1 -- o valor historico esta preservado no raw)   --
-- -------------------------------------------------------------------- --

CREATE TABLE IF NOT EXISTS dim_cliente (
  tenant              VARCHAR(32) NOT NULL,
  id_cliente          BIGINT       NOT NULL,
  nome                VARCHAR(255) NOT NULL DEFAULT '',
  nome_norm           VARCHAR(255) NOT NULL DEFAULT '',
  codigo              VARCHAR(60)  NULL,
  fantasia            VARCHAR(255) NULL,
  tipo_pessoa         CHAR(1)      NULL COMMENT 'J/F/E/X',
  cpf_cnpj            VARCHAR(24)  NULL,
  inscricao_estadual  VARCHAR(40)  NULL,
  rg                  VARCHAR(40)  NULL,
  telefone            VARCHAR(40)  NULL,
  celular             VARCHAR(40)  NULL,
  email               VARCHAR(255) NULL,
  endereco            VARCHAR(255) NULL,
  numero              VARCHAR(30)  NULL,
  complemento         VARCHAR(160) NULL,
  bairro              VARCHAR(160) NULL,
  municipio           VARCHAR(160) NULL,
  cep                 VARCHAR(12)  NULL,
  uf                  CHAR(2)      NULL,
  pais                VARCHAR(60)  NULL,
  eh_cliente_grupo    TINYINT(1)   NOT NULL DEFAULT 0,
  primeiro_visto      DATETIME(3)  NOT NULL,
  ultimo_visto        DATETIME(3)  NOT NULL,
  PRIMARY KEY (tenant, id_cliente),
  KEY ix_cli_nome (tenant, nome_norm(64)),
  KEY ix_cli_doc  (tenant, cpf_cnpj)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='Cliente de marketplace vem sem email e sem telefone -- nao e bug, o marketplace nao repassa. O endereco vem completo.';

CREATE TABLE IF NOT EXISTS dim_produto (
  tenant             VARCHAR(32) NOT NULL,
  id_produto         BIGINT       NOT NULL,
  sku                VARCHAR(80)  COLLATE utf8mb4_0900_as_cs NOT NULL DEFAULT '',
  descricao          VARCHAR(255) NOT NULL DEFAULT '',
  tipo_cadastro      CHAR(1)      NULL COMMENT 'S/V/K/M/F do cadastro -- unica fonte de kit',
  situacao           CHAR(1)      NULL COMMENT 'E = gemeo excluido logicamente',
  categoria_caminho  VARCHAR(500) NULL,
  eh_kit             TINYINT(1) GENERATED ALWAYS AS (tipo_cadastro = 'K') VIRTUAL,
  atualizado_em      DATETIME(3)  NOT NULL,
  PRIMARY KEY (tenant, id_produto),
  KEY ix_prod_sku (tenant, sku)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='Preenchida pela fase 3. O produto.tipo do ITEM do pedido usa outro vocabulario (P) e NAO serve para detectar kit.';

CREATE TABLE IF NOT EXISTS dim_vendedor (
  tenant     VARCHAR(32) NOT NULL,
  id_vendedor BIGINT      NOT NULL,
  nome       VARCHAR(255) NOT NULL DEFAULT '',
  eh_pessoa  TINYINT(1)   NOT NULL DEFAULT 1,
  ultimo_visto DATETIME(3) NOT NULL,
  PRIMARY KEY (tenant, id_vendedor)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE IF NOT EXISTS dim_ecommerce (
  tenant       VARCHAR(32) NOT NULL,
  id_ecommerce BIGINT       NOT NULL,
  nome         VARCHAR(160) NOT NULL DEFAULT '',
  canal_venda  VARCHAR(160) NULL,
  ultimo_visto DATETIME(3)  NOT NULL,
  PRIMARY KEY (tenant, id_ecommerce)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='canalVenda vem vazio em praticamente tudo -- serve de coluna, nunca de chave de agrupamento. Agrupar por nome.';

CREATE TABLE IF NOT EXISTS dim_natureza_operacao (
  tenant           VARCHAR(32) NOT NULL,
  id_natureza      BIGINT       NOT NULL,
  nome             VARCHAR(160) NOT NULL DEFAULT '',
  nome_norm        VARCHAR(160) NOT NULL DEFAULT '',
  eh_natureza_alvo TINYINT(1)   NOT NULL DEFAULT 0,
  ultimo_visto     DATETIME(3)  NOT NULL,
  PRIMARY KEY (tenant, id_natureza)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE IF NOT EXISTS dim_transportador (
  tenant          VARCHAR(32) NOT NULL,
  id_transportador BIGINT      NOT NULL,
  nome            VARCHAR(255) NOT NULL DEFAULT '',
  ultimo_visto    DATETIME(3)  NOT NULL,
  PRIMARY KEY (tenant, id_transportador)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE IF NOT EXISTS dim_deposito (
  tenant       VARCHAR(32) NOT NULL,
  id_deposito  BIGINT       NOT NULL,
  nome         VARCHAR(160) NOT NULL DEFAULT '',
  ultimo_visto DATETIME(3)  NOT NULL,
  PRIMARY KEY (tenant, id_deposito)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE IF NOT EXISTS dim_intermediador (
  tenant           VARCHAR(32) NOT NULL,
  id_intermediador BIGINT       NOT NULL,
  nome             VARCHAR(255) NOT NULL DEFAULT '',
  cnpj             VARCHAR(24)  NULL,
  cnpj_instituicao VARCHAR(24)  NULL,
  ultimo_visto     DATETIME(3)  NOT NULL,
  PRIMARY KEY (tenant, id_intermediador)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE IF NOT EXISTS dim_referencia (
  tenant       VARCHAR(32) NOT NULL,
  dominio      VARCHAR(40)  NOT NULL COMMENT 'forma_recebimento|meio_pagamento|categoria_financeira|forma_envio|forma_frete|lista_preco',
  id_ref       BIGINT       NOT NULL,
  nome         VARCHAR(255) NOT NULL DEFAULT '',
  ultimo_visto DATETIME(3)  NOT NULL,
  PRIMARY KEY (tenant, dominio, id_ref)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='Lookups pequenos do bloco de pagamento e frete numa tabela so; os nomes tambem ficam denormalizados em pedido.';

CREATE TABLE IF NOT EXISTS dim_calendario (
  dia           DATE      NOT NULL,
  ano           SMALLINT  NOT NULL,
  mes           TINYINT   NOT NULL,
  dia_mes       TINYINT   NOT NULL,
  trimestre     TINYINT   NOT NULL,
  ano_mes       CHAR(7)   NOT NULL,
  dia_semana    TINYINT   NOT NULL COMMENT '0=segunda .. 6=domingo',
  nome_dia      CHAR(3)   NOT NULL,
  eh_fim_semana TINYINT(1) NOT NULL,
  PRIMARY KEY (dia),
  KEY ix_cal_ano_mes (ano_mes)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- -------------------------------------------------------------------- --
--  Fato: pedido e filhas                                                --
-- -------------------------------------------------------------------- --

CREATE TABLE IF NOT EXISTS pedido (
  tenant                  VARCHAR(32) NOT NULL,
  id_pedido               BIGINT       NOT NULL,
  numero_pedido           BIGINT       NULL,
  id_nota_fiscal          BIGINT       NULL,
  origem_pedido           TINYINT      NULL COMMENT '0=Venda 1=PDV',

  data_pedido             DATE         NULL,
  data_prevista           DATE         NULL,
  data_entrega            DATE         NULL,
  data_faturamento        DATE         NULL,
  data_envio              DATETIME     NULL,

  situacao                SMALLINT     NULL,
  situacao_nome           VARCHAR(32)  NULL,

  valor_total_produtos    DECIMAL(15,2) NOT NULL DEFAULT 0,
  valor_total_pedido      DECIMAL(15,2) NOT NULL DEFAULT 0,
  valor_desconto          DECIMAL(15,2) NOT NULL DEFAULT 0,
  valor_frete             DECIMAL(15,2) NOT NULL DEFAULT 0,
  valor_outras_despesas   DECIMAL(15,2) NOT NULL DEFAULT 0,

  numero_ordem_compra     VARCHAR(80)  NULL,
  observacoes             TEXT         NULL,
  observacoes_internas    TEXT         NULL,

  id_cliente              BIGINT       NULL,
  cliente_nome            VARCHAR(255) NULL,
  cliente_cpf_cnpj        VARCHAR(24)  NULL,
  cliente_tipo_pessoa     CHAR(1)      NULL,
  cliente_email           VARCHAR(255) NULL,
  cliente_telefone        VARCHAR(40)  NULL,
  cliente_celular         VARCHAR(40)  NULL,
  cliente_uf              CHAR(2)      NULL,
  cliente_municipio       VARCHAR(160) NULL,

  id_natureza             BIGINT       NULL,
  natureza_nome           VARCHAR(160) NULL,
  natureza_norm           VARCHAR(160) NULL,

  id_vendedor             BIGINT       NULL,
  vendedor_nome           VARCHAR(255) NULL,

  id_ecommerce            BIGINT       NULL,
  ecommerce_nome          VARCHAR(160) NULL,
  ecommerce_canal         VARCHAR(160) NULL,
  numero_pedido_ecommerce VARCHAR(80)  NULL,
  numero_pedido_canal     VARCHAR(80)  NULL,

  id_transportador        BIGINT       NULL,
  transportador_nome      VARCHAR(255) NULL,
  frete_por_conta         CHAR(1)      NULL COMMENT 'R/D/T/3/4/S',
  codigo_rastreamento     VARCHAR(160) NULL,
  url_rastreamento        VARCHAR(500) NULL,
  id_forma_envio          BIGINT       NULL,
  forma_envio_nome        VARCHAR(160) NULL,
  id_forma_frete          BIGINT       NULL,
  forma_frete_nome        VARCHAR(160) NULL,

  id_deposito             BIGINT       NULL,
  deposito_nome           VARCHAR(160) NULL,

  id_intermediador        BIGINT       NULL,
  intermediador_nome      VARCHAR(255) NULL,

  id_lista_preco          BIGINT       NULL,
  lista_preco_nome        VARCHAR(160) NULL,
  lista_preco_acrescimo   DECIMAL(15,4) NULL,

  id_forma_recebimento    BIGINT       NULL,
  forma_recebimento_nome  VARCHAR(160) NULL,
  id_meio_pagamento       BIGINT       NULL,
  meio_pagamento_nome     VARCHAR(160) NULL,
  condicao_pagamento      VARCHAR(255) NULL,
  id_categoria_financeira BIGINT       NULL,
  categoria_financeira    VARCHAR(255) NULL,

  -- Flags derivadas das regras de negocio. Materializadas na carga (nao sao
  -- GENERATED) porque dependem de tabelas cfg_* e de classificar_canal(), que
  -- e procedural. Mudou a regra: sobe REGRAS_VERSAO e roda
  -- UPDATE ... WHERE regras_versao < N -- offline, zero request.
  eh_situacao_valida      TINYINT(1)   NOT NULL DEFAULT 0,
  eh_natureza_alvo        TINYINT(1)   NOT NULL DEFAULT 0,
  cliente_excluido        TINYINT(1)   NOT NULL DEFAULT 0,
  canal                   VARCHAR(60)  NULL COMMENT 'grupo de canal definido em canais.grupos do perfil de regras',
  conflito_canal          TINYINT(1)   NOT NULL DEFAULT 0 COMMENT 'vendedor pessoa venceu uma integracao de marketplace',
  tem_detalhe             TINYINT(1)   NOT NULL DEFAULT 0,
  regras_versao           SMALLINT     NOT NULL DEFAULT 0,

  qtd_itens               SMALLINT     NOT NULL DEFAULT 0,
  total_unidades          DECIMAL(15,4) NOT NULL DEFAULT 0,

  criado_em               DATETIME(3)  NOT NULL,
  atualizado_em           DATETIME(3)  NOT NULL,

  PRIMARY KEY (tenant, id_pedido),
  KEY ix_ped_data (tenant, data_pedido),
  KEY ix_ped_numero (tenant, numero_pedido),
  KEY ix_ped_situacao (tenant, situacao, data_pedido),
  KEY ix_ped_cliente (tenant, id_cliente),
  KEY ix_ped_ecom (tenant, id_ecommerce, data_pedido),
  KEY ix_ped_valida (tenant, eh_situacao_valida, eh_natureza_alvo, cliente_excluido),
  KEY ix_ped_regras (regras_versao)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='Cabecalho do pedido, TUDO CRU. Cancelado, transferencia interna e natureza fora de venda entram -- o filtro e coluna, nunca WHERE na ingestao.';

CREATE TABLE IF NOT EXISTS pedido_item (
  tenant          VARCHAR(32) NOT NULL,
  id_pedido       BIGINT       NOT NULL,
  seq             SMALLINT     NOT NULL,
  id_produto      BIGINT       NULL,
  sku             VARCHAR(80)  COLLATE utf8mb4_0900_as_cs NOT NULL DEFAULT '',
  descricao_item  VARCHAR(255) NOT NULL DEFAULT '',
  tipo_item       VARCHAR(8)   NULL COMMENT 'vocabulario do ITEM (P), diferente do cadastro',
  quantidade      DECIMAL(15,4) NOT NULL DEFAULT 0,
  valor_unitario  DECIMAL(15,4) NOT NULL DEFAULT 0 COMMENT 'sempre BRUTO: nao existe desconto em nivel de item',
  valor_bruto     DECIMAL(19,4) GENERATED ALWAYS AS (quantidade * valor_unitario) STORED,
  info_adicional  TEXT         NULL,
  eh_brinde       TINYINT(1) GENERATED ALWAYS AS (valor_unitario = 0) VIRTUAL,
  qtd_zero        TINYINT(1) GENERATED ALWAYS AS (quantidade <= 0) VIRTUAL,
  PRIMARY KEY (tenant, id_pedido, seq),
  KEY ix_item_sku (tenant, sku),
  KEY ix_item_prod (tenant, id_produto),
  CONSTRAINT fk_item_pedido FOREIGN KEY (tenant, id_pedido)
    REFERENCES pedido (tenant, id_pedido) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='Kits ja vem explodidos em componentes pelo Tiny -- NAO explodir de novo. Item de quantidade zero NAO e descartado (so marcado), senao a checagem contra JSON_LENGTH nao fecha.';

CREATE TABLE IF NOT EXISTS pedido_endereco (
  tenant             VARCHAR(32) NOT NULL,
  id_pedido          BIGINT       NOT NULL,
  tipo               ENUM('cobranca','entrega') NOT NULL,
  nome_destinatario  VARCHAR(255) NULL,
  cpf_cnpj           VARCHAR(24)  NULL,
  tipo_pessoa        CHAR(1)      NULL,
  telefone           VARCHAR(40)  NULL,
  inscricao_estadual VARCHAR(40)  NULL,
  endereco           VARCHAR(255) NULL,
  numero             VARCHAR(30)  NULL,
  complemento        VARCHAR(160) NULL,
  bairro             VARCHAR(160) NULL,
  municipio          VARCHAR(160) NULL,
  cep                VARCHAR(12)  NULL,
  uf                 CHAR(2)      NULL,
  pais               VARCHAR(60)  NULL,
  PRIMARY KEY (tenant, id_pedido, tipo),
  KEY ix_end_uf (tenant, tipo, uf),
  CONSTRAINT fk_end_pedido FOREIGN KEY (tenant, id_pedido)
    REFERENCES pedido (tenant, id_pedido) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='cobranca = cliente.endereco; entrega = enderecoEntrega. Mala direta fisica e a unica campanha possivel para comprador de marketplace.';

CREATE TABLE IF NOT EXISTS pedido_parcela (
  tenant                 VARCHAR(32) NOT NULL,
  id_pedido              BIGINT        NOT NULL,
  seq                    SMALLINT      NOT NULL,
  dias                   INT           NULL,
  data_vencimento        DATE          NULL,
  valor                  DECIMAL(15,2) NOT NULL DEFAULT 0,
  observacoes            VARCHAR(500)  NULL,
  id_forma_recebimento   BIGINT        NULL,
  forma_recebimento_nome VARCHAR(160)  NULL,
  id_meio_pagamento      BIGINT        NULL,
  meio_pagamento_nome    VARCHAR(160)  NULL,
  PRIMARY KEY (tenant, id_pedido, seq),
  CONSTRAINT fk_parc_pedido FOREIGN KEY (tenant, id_pedido)
    REFERENCES pedido (tenant, id_pedido) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE IF NOT EXISTS pedido_pagamento_integrado (
  tenant             VARCHAR(32) NOT NULL,
  id_pedido          BIGINT        NOT NULL,
  seq                SMALLINT      NOT NULL,
  valor              DECIMAL(15,2) NOT NULL DEFAULT 0,
  tipo_pagamento     INT           NULL,
  cnpj_intermediador VARCHAR(24)   NULL,
  codigo_autorizacao VARCHAR(160)  NULL,
  codigo_bandeira    INT           NULL,
  PRIMARY KEY (tenant, id_pedido, seq),
  CONSTRAINT fk_pgi_pedido FOREIGN KEY (tenant, id_pedido)
    REFERENCES pedido (tenant, id_pedido) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE IF NOT EXISTS pedido_marcador (
  tenant    VARCHAR(32) NOT NULL,
  id_pedido BIGINT       NOT NULL,
  marcador  VARCHAR(120) NOT NULL,
  PRIMARY KEY (tenant, id_pedido, marcador),
  KEY ix_marc (tenant, marcador),
  CONSTRAINT fk_marc_pedido FOREIGN KEY (tenant, id_pedido)
    REFERENCES pedido (tenant, id_pedido) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='A doc da API nao lista marcadores na resposta do detalhe, so como filtro da listagem. Fica vazia ate a sonda provar o contrario.';

-- -------------------------------------------------------------------- --
--  Controle do pipeline                                                 --
-- -------------------------------------------------------------------- --

CREATE TABLE IF NOT EXISTS etl_pedido_estado (
  tenant          VARCHAR(32) NOT NULL,
  id_pedido       BIGINT      NOT NULL,
  estado          ENUM('listado','detalhado','transformado','erro') NOT NULL,
  data_ref        DATE        NULL,
  numero_pedido   BIGINT      NULL,
  situacao        SMALLINT    NULL,
  hash_lista      CHAR(64)    NULL COLLATE ascii_bin,
  hash_detalhe    CHAR(64)    NULL COLLATE ascii_bin,
  tentativas      SMALLINT    NOT NULL DEFAULT 0,
  erro_msg        VARCHAR(500) NULL,
  listado_em      DATETIME(3) NULL,
  detalhado_em    DATETIME(3) NULL,
  transformado_em DATETIME(3) NULL,
  atualizado_em   DATETIME(3) NOT NULL,
  PRIMARY KEY (tenant, id_pedido),
  KEY ix_estado (tenant, estado, data_ref)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='Estado mutavel, separado do raw que e imutavel. hash_detalhe aponta a versao corrente em pedido_raw. O indice (tenant,estado,data_ref) faz a retomada custar milissegundos.';

CREATE TABLE IF NOT EXISTS etl_janela (
  tenant        VARCHAR(32) NOT NULL,
  ini           DATE        NOT NULL,
  fim           DATE        NOT NULL,
  total_api     INT         NULL COMMENT 'paginacao.total que o servidor respondeu -- unico numero de origem externa no banco',
  total_gravado INT         NOT NULL DEFAULT 0,
  status        ENUM('pendente','ok','erro') NOT NULL DEFAULT 'pendente',
  mensagem      VARCHAR(500) NULL,
  atualizado_em DATETIME(3) NOT NULL,
  PRIMARY KEY (tenant, ini, fim)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE IF NOT EXISTS etl_run (
  id                 BIGINT AUTO_INCREMENT NOT NULL,
  tenant             VARCHAR(32) NOT NULL,
  comando            TEXT        NULL,
  fases              VARCHAR(32) NULL,
  janela_ini         DATE        NULL,
  janela_fim         DATE        NULL,
  inicio             DATETIME(3) NOT NULL,
  fim                DATETIME(3) NULL,
  status             ENUM('rodando','ok','erro','parado_token','interrompido') NOT NULL DEFAULT 'rodando',
  pedidos_listados   INT NOT NULL DEFAULT 0,
  pedidos_detalhados INT NOT NULL DEFAULT 0,
  pedidos_carregados INT NOT NULL DEFAULT 0,
  erros              INT NOT NULL DEFAULT 0,
  mensagem           VARCHAR(500) NULL,
  PRIMARY KEY (id),
  KEY ix_run_tenant (tenant, inicio)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE IF NOT EXISTS etl_erro (
  id        BIGINT AUTO_INCREMENT NOT NULL,
  tenant    VARCHAR(32) NOT NULL,
  id_pedido BIGINT      NULL,
  etapa     VARCHAR(40) NOT NULL,
  tabela    VARCHAR(64) NULL,
  codigo    INT         NULL,
  erro      TEXT        NULL,
  criado_em DATETIME(3) NOT NULL,
  PRIMARY KEY (id),
  KEY ix_erro_tenant (tenant, criado_em)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE IF NOT EXISTS etl_sonda_filtro (
  tenant     VARCHAR(32) NOT NULL,
  filtro     VARCHAR(40) NOT NULL,
  suportado  TINYINT(1)  NOT NULL DEFAULT 0,
  semantica  VARCHAR(20) NULL COMMENT 'maior_igual | igual | ignorado | indefinido',
  evidencia  JSON        NULL,
  testado_em DATETIME(3) NOT NULL,
  PRIMARY KEY (tenant, filtro)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='A API ja provou ACEITAR E IGNORAR filtro (dataInicio/dataFim). Um filtro ignorado responde 200 e nao dispara alarme nenhum, entao dataAtualizacao so e usado depois de aprovado aqui.';
