-- ==================================================================== --
--  003 -- pedido_raw.data_raw: STR_TO_DATE -> CAST AS DATE             --
-- ==================================================================== --
--
-- Migracao para bancos criados antes de 21/09/2026. Instalacao nova ja nasce
-- com a definicao certa em 001_schema.sql; aqui o DROP+ADD so a repete.
--
-- MOTIVO: STR_TO_DATE LEVANTA erro 1411 "Incorrect datetime value" sob
-- sql_mode estrito, em vez de devolver NULL. Como data_raw e uma coluna
-- GENERATED STORED, o erro acontece no INSERT -- e um unico pedido com data
-- fora do padrao ISO derrubaria o lote inteiro de 200 pedidos, repetidamente,
-- sem jeito de isolar o culpado a nao ser linha a linha.
--
-- CAST(... AS DATE) devolve NULL em '31/12/2026', '', 'null', 'abc',
-- '2026-02-31' e '2026-13-45' (todos medidos), e aceita o ISO com 'T'.
--
-- Alterar a expressao de uma coluna STORED exige DROP + ADD com rebuild da
-- tabela. E barato agora (tabela vazia) e caro depois -- por isso a regra do
-- schema: STORED so onde a formula e definicao, nunca conveniencia.
--
-- ORDEM IMPORTA: o DROP COLUMN **nao** remove o indice que contem a coluna --
-- ele apenas encolhe o indice, deixando ix_raw_janela como (tenant, fonte).
-- Dai o DROP INDEX explicito na frente. Assim os quatro comandos sao
-- idempotentes juntos: rodar de novo derruba o indice e a coluna que este
-- mesmo arquivo criou, e refaz os dois.

ALTER TABLE pedido_raw DROP INDEX ix_raw_janela;

ALTER TABLE pedido_raw DROP COLUMN data_raw;

ALTER TABLE pedido_raw ADD COLUMN data_raw DATE GENERATED ALWAYS AS (
    CAST(LEFT(COALESCE(
        NULLIF(NULLIF(payload->>'$.dataCriacao',''),'null'),
        NULLIF(NULLIF(payload->>'$.data',''),'null')), 10) AS DATE)) STORED;

ALTER TABLE pedido_raw ADD INDEX ix_raw_janela (tenant, fonte, data_raw);
