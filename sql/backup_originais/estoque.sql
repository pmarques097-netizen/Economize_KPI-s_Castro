-- Script validado anteriormente para estoque mínimo / faceamento.
-- Permanece editável no menu Banco de Dados > Scripts SQL.
SELECT
    uni.codigo AS loja,
    CASE
        WHEN prod.status = 'A' THEN 'ATIVO'
        WHEN prod.status = 'I' THEN 'INATIVO'
        ELSE COALESCE(prod.status::text, 'SEM STATUS')
    END AS status_produto,
    prod.codigo AS cod_interno,
    emb.etiqueta,
    emb.codigobarras,
    emb.descricao,
    clas.caminho AS classificacao_3_nivel,
    estmi.estoqueminimo::numeric(18, 0) AS estoque_minimo,
    estmi.datahora::date AS data_cadastro
FROM produto AS prod
JOIN embalagem AS emb
    ON emb.produtoid = prod.id
JOIN estoqueminimoprodutounidadenegocio AS estmi
    ON estmi.produtoid = prod.id
JOIN unidadenegocio AS uni
    ON uni.id = estmi.unidadenegocioid
LEFT JOIN classificacaoproduto AS cp
    ON cp.produtoid = prod.id
LEFT JOIN classificacao AS clas
    ON clas.id = cp.classificacaoid
   AND clas.principal = TRUE
WHERE estmi.estoqueminimo >= 1
  AND emb.codigobarras IS NOT NULL
  AND uni.codigo NOT IN (
      '14-2', '24-2', '26', '40', '41', 'BKP', 'CLOUD', 'ESC'
  );
