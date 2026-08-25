SELECT
    CASE
        WHEN p.status = 'A' THEN 'ATIVO'
        WHEN p.status = 'I' THEN 'INATIVO'
    END AS status_produto,
    uni.codigo AS num_loja,
    p.codigo AS cod_int,
    emb.etiqueta,
    emb.codigobarras AS cod_barras,
    emb_contida.codigobarras AS embalagem_filha,
    emb.quantidadeporembalagem AS qtd_por_embalagem,
    emb.padraofornecedores AS principal,
    p.descricao,
    CASE
        WHEN c.caminho LIKE 'PRINCIPAL > 1%' THEN '1-ETICO'
        WHEN c.caminho LIKE 'PRINCIPAL > 2%' THEN '2-GENERICOS'
        WHEN c.caminho LIKE 'PRINCIPAL > 3%' THEN '3-SIMILARES'
        WHEN c.caminho LIKE 'PRINCIPAL > 4%' THEN '4-DIAMANTES'
        WHEN c.caminho LIKE 'PRINCIPAL > 5%' THEN '5-PERFUMARIA'
        WHEN c.caminho LIKE 'PRINCIPAL > 6%' THEN '6-CONVENIENCIA'
        WHEN c.caminho LIKE 'PRINCIPAL > 7%' THEN '7-LEITES E FRALDAS'
        WHEN c.caminho LIKE 'PRINCIPAL > 8%' THEN '8-SERVICOS'
        WHEN c.caminho LIKE 'PRINCIPAL > 9%' THEN '9-USO CONSUMO'
        ELSE 'Outro'
    END AS classificao,
    TRANSLATE(
        c.caminho,
        'áàãâäéèêëíìîïóòõôöúùûüçÁÀÃÂÄÉÈÊËÍÌÎÏÓÒÕÔÖÚÙÛÜÇ',
        'aaaaaeeeeiiiiooooouuuucAAAAAEEEEIIIIOOOOOUUUUC'
    ) AS classificacao_geral,
    pe.razaosocial AS fabricante,
    cabc.nome AS curva,
    custo_p.custo::numeric(18,4) AS custo_unit_atual,
    custo_p.customedio::numeric(18,4) AS custo_medio_atual,
    est.estoque AS estoque
FROM produto AS p
LEFT JOIN curvaabcproduto AS cpabc
    ON cpabc.produtoid = p.id
LEFT JOIN curvaabc AS cabc
    ON cabc.id = cpabc.curvaabcvalorid
LEFT JOIN embalagem AS emb
    ON emb.produtoid = p.id
LEFT JOIN embalagem AS emb_contida
    ON emb_contida.id = emb.embalagemcontidaid
LEFT JOIN estoque AS est
    ON est.embalagemid = emb.id
LEFT JOIN unidadenegocio AS uni
    ON uni.id = est.unidadenegocioid
LEFT JOIN fabricante AS f
    ON f.id = p.fabricanteid
LEFT JOIN pessoa AS pe
    ON pe.id = f.pessoaid
LEFT JOIN classificacaoproduto AS cp
    ON cp.produtoid = p.id
LEFT JOIN classificacao AS c
    ON c.id = cp.classificacaoid
LEFT JOIN (
    SELECT DISTINCT ON (c.produtoid, c.unidadenegocioid)
        c.produtoid,
        c.unidadenegocioid,
        c.custo,
        c.customedio
    FROM custoproduto AS c
    ORDER BY
        c.produtoid,
        c.unidadenegocioid,
        c.id DESC
) AS custo_p
    ON custo_p.produtoid = p.id
   AND custo_p.unidadenegocioid = uni.id
WHERE uni.codigo NOT IN (
    '14-2', '20', '21', '25', '26', '40', '41', 'BKP', 'CLOUD', 'ESC'
)
  AND c.caminho ILIKE '%PRINCIPAL%'
  AND c.caminho NOT IN (
      'PRINCIPAL > 8 - SERVICOS > SERVICOS',
      'PRINCIPAL > 8 - SERVICOS > TAXA DE ENTREGA',
      'PRINCIPAL > 9 - USO CONSUMO > USO CONSUMO',
      'PRINCIPAL > 9 - USO CONSUMO > USO E CONSUMO'
  )
GROUP BY
    p.status,
    uni.codigo,
    p.codigo,
    emb.etiqueta,
    emb.codigobarras,
    emb_contida.codigobarras,
    p.descricao,
    pe.razaosocial,
    cabc.nome,
    c.caminho,
    est.estoque,
    custo_p.custo,
    custo_p.customedio,
    emb.quantidadeporembalagem,
    emb.padraofornecedores;
