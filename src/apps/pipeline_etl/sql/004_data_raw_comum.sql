-- ==================================================================== --
--  004 -- pedido_raw.data_raw deixa de ser GENERATED                   --
-- ==================================================================== --
--
-- Sequencia do 003, que ainda estava errado. Medido em 21/09/2026, com
-- sql_mode = ONLY_FULL_GROUP_BY,STRICT_TRANS_TABLES,NO_ZERO_IN_DATE,
-- NO_ZERO_DATE,ERROR_FOR_DIVISION_BY_ZERO:
--
--   SELECT CAST('31/12/2026' AS DATE)            -> NULL + warning
--   INSERT com isso dentro de coluna GENERATED   -> ERRO 1292
--
-- Em statement que grava, o STRICT promove o warning a erro. STR_TO_DATE da
-- 1411 pelo mesmo motivo. E nao adianta filtrar por REGEXP antes: '2026-02-31'
-- casa com qualquer padrao ISO e continua sendo data invalida.
--
-- Consequencia pratica se ficasse como estava: UM pedido com data fora do
-- padrao derrubaria o INSERT do lote inteiro de 200 pedidos, e o lote seria
-- refeito linha a linha ate isolar o culpado -- em cada lote, o backfill
-- inteiro.
--
-- Agora data_raw e escrita pelo Python (etl_transform.data), que devolve None
-- em qualquer entrada ruim. Perde-se a garantia de que a coluna nao pode
-- divergir do payload; ganha-se um INSERT que nao tem como falhar por data.
-- A view v_recon_fidelidade continua conferindo o resto contra o JSON.

ALTER TABLE pedido_raw DROP INDEX ix_raw_janela;

ALTER TABLE pedido_raw DROP COLUMN data_raw;

ALTER TABLE pedido_raw ADD COLUMN data_raw DATE NULL;

ALTER TABLE pedido_raw ADD INDEX ix_raw_janela (tenant, fonte, data_raw);
