-- ==================================================================== --
--  005 -- perfil de regras versionado por tenant                        --
-- ==================================================================== --
--
-- A politica de venda (situacoes validas, naturezas, canais) saiu do codigo
-- e virou configuracao (core/regras_vendas.py). cfg_regras guarda, por
-- tenant, o hash do perfil efetivo e a versao gravada em pedido.regras_versao:
-- quando o hash muda, a fase 4 so reprocessa com --aplicar-regras (ver
-- etl_regras.versao_regras). O JSON do perfil fica aqui para auditoria e para
-- o pipeline mostrar O QUE mudou.
--
-- cfg_natureza_alvo deixa de existir: com o modo 'prefixo' nao ha lista
-- fechada de nomes, e a tabela nunca foi lida por view nenhuma.
--
-- pedido.canal passa de 20 para 60 caracteres, para caber grupo de canal
-- personalizado. 60 (e nao 64) de proposito: em utf8mb4, ate 63 caracteres
-- o prefixo de tamanho continua com 1 byte e o MySQL altera a coluna no
-- lugar (INPLACE), sem reconstruir a tabela.

CREATE TABLE IF NOT EXISTS cfg_regras (
  tenant         VARCHAR(32)  NOT NULL,
  versao         SMALLINT     NOT NULL,
  hash           VARCHAR(80)  NOT NULL,
  perfil         JSON         NULL,
  atualizado_em  DATETIME     NOT NULL,
  PRIMARY KEY (tenant)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
  COMMENT='Versao do perfil de regras por tenant -- ver etl_regras.versao_regras';

DROP TABLE IF EXISTS cfg_natureza_alvo;

ALTER TABLE pedido MODIFY canal VARCHAR(60) NULL
  COMMENT 'grupo de canal definido em canais.grupos do perfil de regras',
  ALGORITHM=INPLACE, LOCK=NONE;
