WITH cte_razaosocial AS (
    SELECT
        e.codigobarras,
        e.descricao AS descricao_embalagem_cte,
        pe.razaosocial
    FROM embalagem AS e
    LEFT JOIN produto AS p
        ON p.id = e.produtoid
    LEFT JOIN fabricante AS f
        ON f.id = p.fabricanteid
    LEFT JOIN pessoa AS pe
        ON pe.id = f.pessoaid
)
SELECT DISTINCT
    uni.codigo AS numero_loja,
    prod.codigo AS cod_interno,
    inf.codigobarras,
    emb.descricao AS descricao_embalagem,
    cte.razaosocial AS laboratorio,
    clas.caminho AS classificacao_3_nivel,
    emb.quantidadeporembalagem AS quant_embalagem,
    (
        inf.valorunitario /
        NULLIF(emb.quantidadeporembalagem, 0)
    )::numeric(18, 2) AS custo_unit_r,
    CASE
        WHEN (
            (inf.valorunitario - inf.custo) /
            NULLIF(inf.valorunitario, 0)
        ) * 100 < 0
        THEN 0
        ELSE (
            (
                inf.valorunitario -
                (
                    inf.custo -
                    (
                        inf.acrescimo /
                        NULLIF(inf.quantidade, 0)
                    )
                )
            ) /
            NULLIF(inf.valorunitario, 0)
        )::numeric(18, 5)
    END AS desc_unit_percentual,
    (
        (
            inf.acrescimo /
            NULLIF(inf.quantidade, 0)
        ) /
        NULLIF(emb.quantidadeporembalagem, 0)
    )::numeric(18, 5) AS imposto_r,
    (
        inf.custo /
        NULLIF(emb.quantidadeporembalagem, 0)
    )::numeric(18, 2) AS custo_final_r,
    inf.quantidade::numeric(18, 0) AS quantidade_por_produto,
    (
        inf.custo * inf.quantidade
    )::numeric(18, 2) AS entrada_custo_total,
    nf.numero AS numero_nf,
    inf.cfop,
    pes.nome AS fornecedor,
    nf.datahoraemissao::date AS data_emissao,
    nf.datahoraentrada::date AS data_entrada,
    (
        COALESCE(nf.datahoraentrada::date, CURRENT_DATE) -
        nf.datahoraemissao::date
    ) AS dias_entre_datas,
    nf.status,
    CASE
        WHEN nf.status = 'C' THEN 'Conferido'
        WHEN nf.status = 'I' THEN 'Inicial'
        WHEN nf.status = 'A' THEN 'Cancelado'
        ELSE 'Recebido'
    END AS descricao_status
FROM notafiscal AS nf
LEFT JOIN recebimentofisiconotafiscal AS rnt
    ON rnt.notafiscalid = nf.id
LEFT JOIN recebimentofisico AS rcf
    ON rcf.id = rnt.recebimentofisicoid
LEFT JOIN unidadenegocio AS uni
    ON uni.id = nf.unidadenegocioid
LEFT JOIN fornecedor AS forn
    ON forn.id = nf.fornecedorid
LEFT JOIN pessoa AS pes
    ON pes.id = forn.pessoaid
LEFT JOIN itemnotafiscal AS inf
    ON inf.notafiscalid = nf.id
LEFT JOIN embalagem AS emb
    ON emb.id = inf.embalagemid
LEFT JOIN produto AS prod
    ON prod.id = emb.produtoid
LEFT JOIN classificacaoproduto AS cp
    ON cp.produtoid = prod.id
LEFT JOIN classificacao AS clas
    ON clas.id = cp.classificacaoid
   AND clas.principal = TRUE
LEFT JOIN cte_razaosocial AS cte
    ON cte.codigobarras = emb.codigobarras
WHERE nf.datahoraentrada BETWEEN :data_inicio AND :data_fim
  AND nf.status = 'C'
  AND uni.codigo NOT IN (
      '14-2', '24-2', '26', '40', '41', 'BKP', 'CLOUD', 'ESC'
  )
ORDER BY uni.codigo;
