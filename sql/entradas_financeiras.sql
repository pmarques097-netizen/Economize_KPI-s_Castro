-- ENTRADAS FINANCEIRAS OFICIAIS — REGRA A7
-- O relatório Análise de Notas Fiscais de Entrada usa o valor do cabeçalho.
-- Uma nota fiscal aparece uma única vez, independentemente da quantidade de itens.
SELECT DISTINCT
    nf.id AS nota_fiscal_id,
    uni.codigo AS numero_loja,
    nf.numero AS numero_nf,
    pes.nome AS fornecedor,
    nf.valortotal::numeric(18, 2) AS valor_nf_total,
    nf.datahoraemissao::date AS data_emissao,
    nf.datahoraentrada::date AS data_entrada,
    nf.status,
    CASE
        WHEN nf.status = 'C' THEN 'Conferido'
        WHEN nf.status = 'I' THEN 'Inicial'
        WHEN nf.status = 'A' THEN 'Cancelado'
        ELSE 'Recebido'
    END AS descricao_status
FROM notafiscal AS nf
LEFT JOIN unidadenegocio AS uni
    ON uni.id = nf.unidadenegocioid
LEFT JOIN fornecedor AS forn
    ON forn.id = nf.fornecedorid
LEFT JOIN pessoa AS pes
    ON pes.id = forn.pessoaid
WHERE nf.datahoraentrada BETWEEN :data_inicio AND :data_fim
  AND nf.status = 'C'
  AND uni.codigo NOT IN (
      '14-2', '24-2', '26', '40', '41', 'BKP', 'CLOUD', 'ESC'
  )
ORDER BY uni.codigo, nf.numero;
