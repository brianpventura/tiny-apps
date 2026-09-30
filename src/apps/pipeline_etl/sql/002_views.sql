-- ==================================================================== --
--  tiny_dw -- views de leitura e de reconciliacao                      --
-- ==================================================================== --
--
-- As views existem para que nenhum relatorio futuro reescreva as regras de
-- negocio em SQL. Quem quiser "vendas validas" usa v_venda_item e pronto.
--
-- CREATE OR REPLACE VIEW e idempotente, entao este arquivo pode ser
-- reaplicado sempre que uma regra mudar.

-- -------------------------------------------------------------------- --
--  v_pedido -- cabecalho com o motivo do descarte ja resolvido          --
-- -------------------------------------------------------------------- --
-- motivo_descarte usa CASE em cascata, o que garante que cada pedido caia em
-- EXATAMENTE UM motivo. E essa propriedade que faz o censo do Diagnostico
-- fechar: listados = contados + descartados, sem sobreposicao.

CREATE OR REPLACE VIEW v_pedido AS
SELECT
    p.tenant,
    t.empresa,
    p.id_pedido,
    p.numero_pedido,
    p.data_pedido,
    c.ano_mes,
    c.nome_dia,
    c.eh_fim_semana,
    p.situacao,
    p.situacao_nome,
    p.natureza_nome,
    p.canal,
    p.conflito_canal,
    p.ecommerce_nome,
    p.ecommerce_canal,
    p.vendedor_nome,
    p.id_cliente,
    p.cliente_nome,
    p.cliente_uf,
    p.valor_total_produtos,
    p.valor_total_pedido,
    p.valor_desconto,
    p.valor_frete,
    p.qtd_itens,
    p.total_unidades,
    p.tem_detalhe,
    p.regras_versao,
    (p.eh_situacao_valida = 1
     AND p.cliente_excluido = 0
     AND p.tem_detalhe = 1
     AND p.eh_natureza_alvo = 1) AS eh_venda_valida,
    CASE
      WHEN p.eh_situacao_valida = 0 THEN 'situacao'
      WHEN p.cliente_excluido   = 1 THEN 'transferencia_interna'
      WHEN p.tem_detalhe        = 0 THEN 'sem_detalhe'
      WHEN p.eh_natureza_alvo   = 0 THEN 'natureza_fora'
      ELSE NULL
    END AS motivo_descarte
FROM pedido p
LEFT JOIN cfg_tenant t     ON t.tenant = p.tenant
LEFT JOIN dim_calendario c ON c.dia = p.data_pedido;

-- -------------------------------------------------------------------- --
--  v_venda_item -- o grao de trabalho de quase todo relatorio           --
-- -------------------------------------------------------------------- --
-- Ja filtrado por eh_venda_valida. Usa a descricao do CADASTRO com fallback
-- para a do item, para que produto renomeado nao vire duas linhas.
-- ATENCAO: inclui linha de brinde (valor_unitario = 0). Brinde saiu do
-- estoque e e unidade vendida com receita zero -- cortar e decisao de quem
-- consulta, nao default.

CREATE OR REPLACE VIEW v_venda_item AS
SELECT
    vp.tenant,
    vp.empresa,
    vp.id_pedido,
    vp.numero_pedido,
    vp.data_pedido,
    vp.ano_mes,
    vp.situacao,
    vp.canal,
    vp.ecommerce_nome,
    vp.vendedor_nome,
    vp.id_cliente,
    vp.cliente_nome,
    i.seq,
    i.id_produto,
    i.sku,
    COALESCE(NULLIF(d.descricao, ''), i.descricao_item) AS produto,
    d.tipo_cadastro,
    d.eh_kit,
    d.categoria_caminho,
    i.quantidade,
    i.valor_unitario,
    i.valor_bruto,
    i.eh_brinde,
    i.qtd_zero
FROM v_pedido vp
JOIN pedido_item i
  ON i.tenant = vp.tenant AND i.id_pedido = vp.id_pedido
LEFT JOIN dim_produto d
  ON d.tenant = i.tenant AND d.id_produto = i.id_produto
WHERE vp.eh_venda_valida = 1;

-- -------------------------------------------------------------------- --
--  v_pedido_situacao_hist -- historico 5->6 SEM tabela de historico     --
-- -------------------------------------------------------------------- --
-- Sai de graca do pedido_raw versionado: cada re-extracao com conteudo
-- diferente vira uma linha nova, e LAG() compara com a anterior. Responde
-- "quanto tempo do Enviada ate o Entregue?", que nenhum relatorio do repo
-- consegue hoje.

CREATE OR REPLACE VIEW v_pedido_situacao_hist AS
SELECT tenant, id_pedido, situacao_ant, situacao, mudou_em,
       TIMESTAMPDIFF(HOUR, visto_ant, mudou_em) AS horas_na_situacao_ant
FROM (
    SELECT r.tenant,
           r.id_pedido,
           r.situacao_raw AS situacao,
           LAG(r.situacao_raw) OVER w AS situacao_ant,
           LAG(r.extraido_em)  OVER w AS visto_ant,
           r.extraido_em AS mudou_em
    FROM pedido_raw r
    WHERE r.situacao_raw IS NOT NULL
    WINDOW w AS (PARTITION BY r.tenant, r.id_pedido
                 ORDER BY r.extraido_em, r.fonte)
) x
WHERE situacao_ant IS NOT NULL AND situacao_ant <> situacao;

-- ==================================================================== --
--  Reconciliacao -- toda view abaixo DEVE voltar vazia (ou status ok)  --
-- ==================================================================== --

-- 1. Completude: o que o servidor disse que existia contra o que gravamos.
--    total_api veio do paginacao.total, e e o unico numero de origem
--    externa no banco -- da para conferir a carga sem gastar request.
CREATE OR REPLACE VIEW v_recon_janela AS
SELECT j.tenant, j.ini, j.fim, j.status, j.total_api, j.total_gravado,
       (j.total_gravado - COALESCE(j.total_api, j.total_gravado)) AS delta
FROM etl_janela j
WHERE j.status <> 'ok'
   OR j.total_api IS NULL
   OR j.total_gravado <> j.total_api;

-- 2. Pendencias por estado. 'listado' = fase 2 incompleta; 'erro' sai nomeado.
CREATE OR REPLACE VIEW v_recon_estado AS
SELECT tenant, estado, COUNT(*) AS pedidos,
       MIN(data_ref) AS de, MAX(data_ref) AS ate
FROM etl_pedido_estado
GROUP BY tenant, estado;

-- 3. Orfaos nos dois sentidos.
CREATE OR REPLACE VIEW v_recon_raw_sem_pedido AS
SELECT e.tenant, e.id_pedido, e.estado, e.data_ref
FROM etl_pedido_estado e
LEFT JOIN pedido p ON p.tenant = e.tenant AND p.id_pedido = e.id_pedido
WHERE e.estado = 'transformado' AND p.id_pedido IS NULL;

CREATE OR REPLACE VIEW v_recon_pedido_sem_raw AS
SELECT p.tenant, p.id_pedido, p.data_pedido
FROM pedido p
LEFT JOIN etl_pedido_estado e ON e.tenant = p.tenant AND e.id_pedido = p.id_pedido
WHERE e.id_pedido IS NULL;

-- 4. Pedido detalhado que deveria ter item e nao tem. E o sintoma de um
--    DELETE+INSERT interrompido no meio -- erro que some em qualquer SUM().
CREATE OR REPLACE VIEW v_recon_pedido_sem_item AS
SELECT p.tenant, p.id_pedido, p.data_pedido, p.qtd_itens, p.valor_total_produtos
FROM pedido p
LEFT JOIN pedido_item i ON i.tenant = p.tenant AND i.id_pedido = p.id_pedido
WHERE p.tem_detalhe = 1 AND p.qtd_itens > 0 AND i.id_pedido IS NULL;

-- 5. Aritmetica: soma(qtd * valorUnitario) contra valorTotalProdutos.
--    Serve para pegar FLOAT vazado no transform, que daria erro grande e
--    erratico. NAO serve para pegar centavo, e o limite de 0,05 explica por que:
--
--    Quando houve desconto rateado pelo integrador, os valorUnitario vem com
--    QUATRO casas decimais (ex.: 82.6132, 108.0346, 57.1918, soma 394,0128)
--    enquanto o proprio Tiny declara valorTotalProdutos = 394,00. O delta de
--    R$ 0,01 e arredondamento DA API, nao erro de carga.
--
--    Isso tambem prova o rateio de desconto: o cabecalho traz valorDesconto e
--    os itens trazem unitarios quebrados. Como a API nao tem desconto em nivel
--    de item, o integrador distribuiu o desconto no unitario antes de criar o
--    pedido. Por isso valorUnitario continua sendo preco BRUTO por definicao,
--    mas nem sempre e o preco de tabela.
--    A tolerancia NAO e um numero fixo, e isso foi medido, nao arbitrado. Em
--    dezenas de milhares de pedidos o delta de arredondamento cresce por dois
--    caminhos independentes:
--
--      a) por LINHA -- cada unitario e arredondado em 4 casas, e o erro se
--         acumula linha a linha: ~R$ 0,002 por linha (ex.: 38 linhas deram
--         delta de R$ 0,11).
--      b) por QUANTIDADE -- a truncagem de 4 casas multiplica pela quantidade.
--         Ex.: UMA linha de 30.100 unidades a R$ 0,8808, unitario real
--         0,88081694...: 30.100 x 0,0000169 = R$ 0,51 de delta, sem acumulo.
--
--    Um limiar fixo nao cobre os dois: o que perdoa (b) cega (a). Dai o
--    GREATEST. O que sobra acusado NAO e arredondamento: ja apareceu pedido
--    cujo cabecalho da API soma um item a mais do que os itens listados
--    (provavelmente item removido sem recalculo do total). Inconsistencia da
--    ORIGEM, nao da carga -- esses casos, depois de conferidos no payload,
--    entram na lista de aceitos de data/verificacao.json.
CREATE OR REPLACE VIEW v_recon_aritmetica AS
SELECT p.tenant, p.id_pedido, p.data_pedido,
       p.valor_total_produtos,
       ROUND(SUM(i.valor_bruto), 2) AS soma_itens,
       ROUND(SUM(i.valor_bruto) - p.valor_total_produtos, 2) AS delta,
       COUNT(*) AS linhas,
       SUM(i.quantidade) AS unidades,
       ROUND(GREATEST(0.05, 0.005 * COUNT(*),
                      0.0001 * SUM(i.quantidade)), 4) AS tolerancia
FROM pedido p
JOIN pedido_item i ON i.tenant = p.tenant AND i.id_pedido = p.id_pedido
WHERE p.tem_detalhe = 1
GROUP BY p.tenant, p.id_pedido, p.data_pedido, p.valor_total_produtos
HAVING ABS(SUM(i.valor_bruto) - p.valor_total_produtos)
       > GREATEST(0.05, 0.005 * COUNT(*), 0.0001 * SUM(i.quantidade));

-- 5b. O mesmo delta na casa do centavo, so para OLHAR. Nao reprova nada: e o
--     arredondamento da API. Se esta contagem explodir de repente, ai sim vale
--     investigar o integrador de origem.
CREATE OR REPLACE VIEW v_resumo_arredondamento AS
SELECT p.tenant, COUNT(*) AS pedidos_com_delta,
       ROUND(SUM(ABS(ROUND(x.soma, 2) - p.valor_total_produtos)), 2) AS delta_absoluto,
       ROUND(MAX(ABS(ROUND(x.soma, 2) - p.valor_total_produtos)), 2) AS maior_delta
FROM pedido p
JOIN (SELECT tenant, id_pedido, SUM(valor_bruto) AS soma
        FROM pedido_item GROUP BY tenant, id_pedido) x
  ON x.tenant = p.tenant AND x.id_pedido = p.id_pedido
WHERE p.tem_detalhe = 1
  AND ABS(ROUND(x.soma, 2) - p.valor_total_produtos) > 0.005
GROUP BY p.tenant;

-- 6. Fidelidade raw -> normalizado, lida direto do JSON, sem reprocessar.
--    Prova que o transform nao inventou nem perdeu nada.
--
--    CUIDADO com JSON_LENGTH em `"itens": null`: ele devolve 1, nao NULL, porque
--    JSON null e um ESCALAR e escalar tem comprimento 1. O COALESCE que estava
--    aqui portanto nunca disparava, e a view acusava falso positivo -- pedido
--    corretamente carregado com 0 itens contra "1 item no raw". A API manda
--    itens = null em vez de [] em alguns pedidos cancelados (raro: ~1 em 20
--    mil). Por isso o comprimento so e lido quando JSON_TYPE confirma que e
--    ARRAY.
CREATE OR REPLACE VIEW v_recon_fidelidade AS
SELECT p.tenant, p.id_pedido,
       p.situacao AS situacao_carregada,
       CAST(r.payload->>'$.situacao' AS SIGNED) AS situacao_raw,
       p.qtd_itens AS itens_carregados,
       CASE WHEN JSON_TYPE(JSON_EXTRACT(r.payload, '$.itens')) = 'ARRAY'
            THEN JSON_LENGTH(r.payload, '$.itens') ELSE 0 END AS itens_raw
FROM pedido p
JOIN etl_pedido_estado e
  ON e.tenant = p.tenant AND e.id_pedido = p.id_pedido
JOIN pedido_raw r
  ON r.tenant = e.tenant AND r.id_pedido = e.id_pedido
 AND r.fonte = 'detalhe' AND r.payload_hash = e.hash_detalhe
WHERE p.situacao <> CAST(r.payload->>'$.situacao' AS SIGNED)
   OR p.qtd_itens <> CASE
        WHEN JSON_TYPE(JSON_EXTRACT(r.payload, '$.itens')) = 'ARRAY'
        THEN JSON_LENGTH(r.payload, '$.itens') ELSE 0 END;

-- 7. numero_pedido repetido. NAO existe UNIQUE sobre ele de proposito (o ODKU
--    exige uma unica chave unica por tabela), entao a conferencia vem aqui.
CREATE OR REPLACE VIEW v_recon_numero_duplicado AS
SELECT tenant, numero_pedido, COUNT(*) AS ocorrencias,
       GROUP_CONCAT(id_pedido ORDER BY id_pedido) AS ids
FROM pedido
WHERE numero_pedido IS NOT NULL
GROUP BY tenant, numero_pedido
HAVING COUNT(*) > 1;

-- 8. Sanidade de negocio -- nao reprova a carga, so mostra o retrato.
CREATE OR REPLACE VIEW v_resumo_canal AS
SELECT tenant, empresa, ano_mes, canal,
       COUNT(*) AS pedidos,
       ROUND(SUM(valor_total_pedido), 2) AS faturamento
FROM v_pedido
WHERE eh_venda_valida = 1
GROUP BY tenant, empresa, ano_mes, canal;

CREATE OR REPLACE VIEW v_resumo_integracao AS
SELECT tenant, empresa, ano_mes,
       COALESCE(NULLIF(ecommerce_nome, ''), '(sem integracao)') AS integracao,
       COUNT(*) AS pedidos,
       ROUND(SUM(valor_total_pedido), 2) AS faturamento
FROM v_pedido
WHERE eh_venda_valida = 1
GROUP BY tenant, empresa, ano_mes,
         COALESCE(NULLIF(ecommerce_nome, ''), '(sem integracao)');
