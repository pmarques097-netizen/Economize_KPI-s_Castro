SELECT
    unidadenegocio.nome AS loja,
    unidadenegocio.codigo AS numero_loja,
    us_orc.apelido AS usuario_orcamento,
    venda.id AS vendaid,
    venda.status AS status_venda,
    embalagem.descricao,
    embalagem.etiqueta,
    embalagem.codigobarras,
    produto.codigo AS cod_interno,
    classificacao.caminho AS classificacao_3_nivel,
    CASE
        WHEN classificacao.caminho LIKE '%>%- %'
            THEN SPLIT_PART(classificacao.caminho, '>', 2)
        ELSE classificacao.caminho
    END AS classificacao_resumida,
    itemvenda.quantidade,
    itemvenda.valorunitario,
    itemorcamento.precovenda,
    COALESCE(me.custo, 0) AS custo,
    COALESCE(me.quantidade, itemvenda.quantidade) AS qtd_mov,
    itemvenda.desconto,
    itemvenda.valortotal,
    (
        itemvenda.valortotal
        - COALESCE(me.custo, 0) * itemvenda.quantidade
    ) AS lucro,
    CASE
        WHEN itemvenda.valortotal <> 0
        THEN (
            itemvenda.valortotal
            - COALESCE(me.custo, 0) * itemvenda.quantidade
        ) / itemvenda.valortotal
        ELSE 0
    END AS lucro_perc,
    venda.datahorafechamento AS datahora_venda_final,
    COALESCE(me.datahora, itemvenda.datahora) AS datahora,
    venda.cpfcnpj,
    itemvenda.status AS status_item_venda,
    pbm_primario.nome AS programa_pbm,
    pbm_secundario.nome AS programa_pbm_secundario,
    CASE
        WHEN 'PEC - FIDELIDADE' IN (
            pbm_primario.nome,
            pbm_secundario.nome
        )
        THEN 'PEC'
        ELSE 'NAO-PEC'
    END AS pec,
    venda.coo
FROM itemvenda
JOIN venda
    ON venda.id = itemvenda.vendaid
   AND venda.status = 'F'
JOIN embalagem
    ON itemvenda.embalagemid = embalagem.id
LEFT JOIN produto
    ON embalagem.produtoid = produto.id
LEFT JOIN LATERAL (
    SELECT c.caminho
    FROM classificacaoproduto cp
    JOIN classificacao c
      ON c.id = cp.classificacaoid
     AND c.principal = TRUE
    WHERE cp.produtoid = produto.id
    ORDER BY c.id
    LIMIT 1
) AS classificacao ON TRUE
LEFT JOIN itemorcamento
    ON itemvenda.itemorcamentoid = itemorcamento.id
LEFT JOIN orcamento
    ON itemorcamento.orcamentoid = orcamento.id
LEFT JOIN usuario us_orc
    ON us_orc.id = itemorcamento.usuarioid
LEFT JOIN unidadenegocio
    ON unidadenegocio.id = itemvenda.unidadenegocioid
LEFT JOIN orcamentopbm
    ON orcamento.id = orcamentopbm.orcamentoid
LEFT JOIN pbm pbm_primario
    ON orcamentopbm.pbmid = pbm_primario.id
LEFT JOIN pbm pbm_secundario
    ON pbm_secundario.id = orcamentopbm.pbmsecundariaid
LEFT JOIN LATERAL (
    SELECT
        mov.custo,
        mov.quantidade,
        mov.datahora
    FROM movimentacaoestoque mov
    WHERE mov.id = itemvenda.movimentacaoestoqueid
      AND mov.unidadenegocioid = itemvenda.unidadenegocioid
      AND mov.tipomovimentacaoestoqueid = 26
    ORDER BY mov.datahora DESC
    LIMIT 1
) AS me ON TRUE
WHERE itemvenda.status = 'F'
  AND venda.datahorafechamento BETWEEN :data_inicio AND :data_fim
  -- Não excluir USO CONSUMO: o relatório A7 considera todos os itens finalizados.
  AND unidadenegocio.codigo NOT IN (
      '14-2', '24-2', '26', '40', '41', 'BKP', 'CLOUD', 'ESC'
  );
