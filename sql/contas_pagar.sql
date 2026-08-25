/*
===============================================================================
CONTAS A PAGAR — V15 FINAL PARA DBEAVER
PostgreSQL | Somente leitura | Não cria VIEW | Não altera o banco

Altere somente as datas no bloco parametros.
===============================================================================
*/

WITH
parametros AS (
    SELECT
        DATE '2026-07-01' AS data_inicial,
        DATE '2026-07-10' AS data_final
),

pagamentos AS (
    SELECT
        ip.contapagarid,
        MAX(pg.datahora) AS data_pagamento,
        MAX(pg.codigo) AS codigo_pagamento,
        SUM(COALESCE(ip.valorpagamento, 0)) AS valor_pago
    FROM itempgtocontapagar ip
    INNER JOIN pgtocontapagar pg
        ON pg.id = ip.pgtocontapagarid
    GROUP BY ip.contapagarid
),

retiradas AS (
    SELECT
        ip.contapagarid,
        STRING_AGG(
            DISTINCT CONCAT(
                CASE rp.tipo
                    WHEN 'A' THEN 'CONTA CORRENTE'
                    WHEN 'B' THEN 'CHEQUE PRÓPRIO'
                    WHEN 'C' THEN 'CHEQUE DE TERCEIRO'
                    ELSE 'RETIRADA'
                END,
                CASE
                    WHEN rp.contacorrenteid IS NOT NULL
                    THEN CONCAT(' - CONTA ID: ', rp.contacorrenteid)
                    ELSE ''
                END,
                ' (R$ ',
                REPLACE(
                    TO_CHAR(COALESCE(rp.valor, 0), 'FM999999999990.00'),
                    '.',
                    ','
                ),
                ')'
            ),
            ', '
        ) AS retiradas_pagamento
    FROM retiradapgtocontapagar rp
    INNER JOIN itempgtocontapagar ip
        ON ip.id = rp.itempgtocontapagarid
    GROUP BY ip.contapagarid
),

notas AS (
    SELECT
        nfc.contapagarid,
        STRING_AGG(
            DISTINCT COALESCE(nf.numero::text, nfc.notafiscalid::text),
            ', '
        ) AS notas_fiscais
    FROM notafiscalcontapagar nfc
    LEFT JOIN notafiscal nf
        ON nf.id = nfc.notafiscalid
    GROUP BY nfc.contapagarid
),

cancelamentos AS (
    SELECT
        ccp.contapagarid,
        MAX(ccp.motivo) AS motivo
    FROM cancelamentocontapagar ccp
    GROUP BY ccp.contapagarid
),

pagamento_eletronico AS (
    SELECT DISTINCT ON (pec.contapagarid)
        pec.contapagarid,
        pec.status,
        pec.datapagamento,
        pec.datahora
    FROM pagamentoeletronicoconta pec
    ORDER BY
        pec.contapagarid,
        pec.datahora DESC NULLS LAST
),

dda_conta AS (
    SELECT
        d.contapagarid,
        TRUE AS conciliada_dda
    FROM dda d
    WHERE d.contapagarid IS NOT NULL
    GROUP BY d.contapagarid
),

base AS (
    SELECT
        cp.id AS contapagar_id,

        CASE cp.origem
            WHEN 'A' THEN 'CT-e'
            WHEN 'M' THEN 'Manual'
            WHEN 'N' THEN 'Nota Fiscal'
            WHEN 'B' THEN 'Unificação'
            WHEN 'C' THEN 'DDA'
            ELSE 'Não Especificada'
        END AS origem,

        CASE cp.status
            WHEN 'A' THEN 'Pendente'
            WHEN 'B' THEN 'Paga'
            WHEN 'C' THEN 'Cancelada'
            ELSE COALESCE(cp.status::text, 'Não Especificado')
        END AS status,

        CASE pe.status
            WHEN 'A' THEN 'Pendente'
            WHEN 'B' THEN 'Em remessa'
            WHEN 'C' THEN 'Agendado'
            WHEN 'D' THEN 'Efetuado'
            WHEN 'E' THEN 'Cancelado'
            WHEN 'F' THEN 'Rejeitado'
            WHEN 'G' THEN 'Estornado'
            ELSE NULL
        END AS status_eletronico,

        cp.datavencimento::date AS data_vencimento,
        cp.numerodocumento AS numero_documento,
        COALESCE(cp.valor, 0) AS valor,

        COALESCE(
            NULLIF(TRIM(pb.razaosocial), ''),
            NULLIF(TRIM(pb.nome), ''),
            NULLIF(TRIM(pc.razaosocial), ''),
            NULLIF(TRIM(pc.nome), '')
        ) AS credor,

        un.codigo AS unidade,
        NULLIF(TRIM(un.nome), '') AS apelido_unidade,

        COALESCE(cp.valordocumento, 0) AS valor_documento,
        COALESCE(cp.desconto, 0) AS desconto,
        COALESCE(cp.descontocredito, 0) AS desconto_credito,
        COALESCE(cp.acrescimo, 0) AS acrescimo,
        COALESCE(cp.multa, 0) AS multa,

        cp.dataemissao::date AS data_emissao,
        cp.datavencimentoutil::date AS data_util_vencimento,

        COALESCE(
            pg.data_pagamento,
            pg_direto.datahora,
            pe.datapagamento
        )::timestamp AS data_pagamento,

        COALESCE(pg.codigo_pagamento, pg_direto.codigo) AS codigo_pagamento,
        rt.retiradas_pagamento,

        CASE
            WHEN COALESCE(pg.data_pagamento, pg_direto.datahora, pe.datapagamento) IS NOT NULL
            THEN COALESCE(pg.data_pagamento, pg_direto.datahora, pe.datapagamento)::date
                 - cp.datavencimento::date
            WHEN cp.status = 'A'
            THEN GREATEST(CURRENT_DATE - cp.datavencimento::date, 0)
            ELSE NULL
        END AS dias_atraso,

        CASE
            WHEN cp.numeroparcela IS NULL AND cp.totalparcela IS NULL THEN NULL
            ELSE CONCAT(
                COALESCE(cp.numeroparcela, 1),
                ' / ',
                COALESCE(cp.totalparcela, 1)
            )
        END AS parcela,

        COALESCE(cp.aguardandodocumento, FALSE) AS aguardando_documento,
        cp.descricao,
        plc.caminho AS plano_contas,

        CASE
            WHEN ul.id IS NULL THEN NULL
            WHEN COALESCE(NULLIF(TRIM(pul.nome), ''), NULLIF(TRIM(pul.razaosocial), '')) IS NOT NULL
            THEN CONCAT(
                COALESCE(NULLIF(TRIM(ul.login), ''), TRIM(ul.apelido)),
                '-',
                COALESCE(NULLIF(TRIM(pul.nome), ''), NULLIF(TRIM(pul.razaosocial), ''))
            )
            ELSE COALESCE(NULLIF(TRIM(ul.login), ''), TRIM(ul.apelido))
        END AS usuario,

        cp.codigobarras AS codigo_barras,
        COALESCE(cp.restringir, FALSE) AS restringir,
        nt.notas_fiscais,

        CASE cp.origemcancelamento
            WHEN 'A' THEN 'Manual'
            WHEN 'B' THEN 'Unificação'
            ELSE 'Não Especificada'
        END AS origem_cancelamento,

        canc.motivo AS motivo_cancelamento,

        CASE
            WHEN ua.id IS NULL THEN NULL
            WHEN COALESCE(NULLIF(TRIM(pua.nome), ''), NULLIF(TRIM(pua.razaosocial), '')) IS NOT NULL
            THEN CONCAT(
                COALESCE(NULLIF(TRIM(ua.login), ''), TRIM(ua.apelido)),
                '-',
                COALESCE(NULLIF(TRIM(pua.nome), ''), NULLIF(TRIM(pua.razaosocial), ''))
            )
            ELSE COALESCE(NULLIF(TRIM(ua.login), ''), TRIM(ua.apelido))
        END AS usuario_autorizacao_pagamento,

        cp.datahoraautorizacaopagamento AS data_hora_autorizacao_pagamento,
        COALESCE(ddac.conciliada_dda, FALSE) AS conciliada_dda,
        sped.descricao AS contribuicao_credito_sped,
        COALESCE(cp.confirmacaopagamentoautomatico, FALSE) AS confirmacao_pagamento_automatico,

        CASE
            WHEN LENGTH(REGEXP_REPLACE(
                COALESCE(NULLIF(pb.cnpj, ''), NULLIF(pc.cnpj, ''), NULLIF(pb.cpf, ''), NULLIF(pc.cpf, '')),
                '[^0-9]', '', 'g'
            )) = 14
            THEN REGEXP_REPLACE(
                REGEXP_REPLACE(
                    COALESCE(NULLIF(pb.cnpj, ''), NULLIF(pc.cnpj, ''), NULLIF(pb.cpf, ''), NULLIF(pc.cpf, '')),
                    '[^0-9]', '', 'g'
                ),
                '^([0-9]{2})([0-9]{3})([0-9]{3})([0-9]{4})([0-9]{2})$',
                '\1.\2.\3/\4-\5'
            )
            WHEN LENGTH(REGEXP_REPLACE(
                COALESCE(NULLIF(pb.cpf, ''), NULLIF(pc.cpf, ''), NULLIF(pb.cnpj, ''), NULLIF(pc.cnpj, '')),
                '[^0-9]', '', 'g'
            )) = 11
            THEN REGEXP_REPLACE(
                REGEXP_REPLACE(
                    COALESCE(NULLIF(pb.cpf, ''), NULLIF(pc.cpf, ''), NULLIF(pb.cnpj, ''), NULLIF(pc.cnpj, '')),
                    '[^0-9]', '', 'g'
                ),
                '^([0-9]{3})([0-9]{3})([0-9]{3})([0-9]{2})$',
                '\1.\2.\3-\4'
            )
            ELSE COALESCE(NULLIF(pb.cnpj, ''), NULLIF(pc.cnpj, ''), NULLIF(pb.cpf, ''), NULLIF(pc.cpf, ''))
        END AS documento_credor

    FROM contapagar cp
    CROSS JOIN parametros prm
    LEFT JOIN pessoa pc ON pc.id = cp.pessoaid
    LEFT JOIN pessoa pb ON pb.id = cp.pessoabeneficiarioid
    LEFT JOIN unidadenegocio un ON un.id = cp.unidadenegocioid
    LEFT JOIN planocontas plc ON plc.id = cp.planocontasid
    LEFT JOIN usuario ul ON ul.id = cp.usuarioid
    LEFT JOIN pessoa pul ON pul.id = ul.pessoaid
    LEFT JOIN usuario ua ON ua.id = cp.usuarioautorizacaopagamentoid
    LEFT JOIN pessoa pua ON pua.id = ua.pessoaid
    LEFT JOIN pagamentos pg ON pg.contapagarid = cp.id
    LEFT JOIN pgtocontapagar pg_direto ON pg_direto.id = cp.pgtocontapagarid
    LEFT JOIN retiradas rt ON rt.contapagarid = cp.id
    LEFT JOIN notas nt ON nt.contapagarid = cp.id
    LEFT JOIN cancelamentos canc ON canc.contapagarid = cp.id
    LEFT JOIN pagamento_eletronico pe ON pe.contapagarid = cp.id
    LEFT JOIN dda_conta ddac ON ddac.contapagarid = cp.id
    LEFT JOIN spedcadastrogeracaocontribuicaocredito sped
        ON sped.id = cp.spedcadastrogeracaocontribuicaocreditoid

    WHERE cp.datavencimento >= prm.data_inicial
      AND cp.datavencimento < prm.data_final + INTERVAL '1 day'
)

SELECT
    origem AS "Origem",
    status AS "Status",
    status_eletronico AS "Status Eletrônico",
    TO_CHAR(data_vencimento, 'DD/MM/YYYY') AS "Data de Vencimento",
    numero_documento AS "Número Documento",
    REPLACE(TO_CHAR(valor, 'FM999999999990.00'), '.', ',') AS "Valor",
    credor AS "Credor",
    unidade AS "Unidade",
    apelido_unidade AS "Apelido Un. Neg.",
    REPLACE(TO_CHAR(valor_documento, 'FM999999999990.00'), '.', ',') AS "Valor Documento",
    REPLACE(TO_CHAR(desconto, 'FM999999999990.00'), '.', ',') AS "Desconto",
    REPLACE(TO_CHAR(desconto_credito, 'FM999999999990.00'), '.', ',') AS "Desconto Crédito",
    REPLACE(TO_CHAR(acrescimo, 'FM999999999990.00'), '.', ',') AS "Acréscimo",
    REPLACE(TO_CHAR(multa, 'FM999999999990.00'), '.', ',') AS "Multa",
    TO_CHAR(data_emissao, 'DD/MM/YYYY') AS "Data Emissão",
    TO_CHAR(data_util_vencimento, 'DD/MM/YYYY') AS "Data Útil de Vencimento",
    TO_CHAR(data_pagamento, 'DD/MM/YYYY') AS "Data Pagamento",
    codigo_pagamento AS "Código Pagamento",
    retiradas_pagamento AS "Retiradas Pagamento",
    dias_atraso AS "Dias Atraso",
    parcela AS "Parcela",
    CASE WHEN aguardando_documento THEN 'Sim' ELSE 'Não' END AS "Aguardando Documento",
    descricao AS "Descrição",
    plano_contas AS "Plano de Contas",
    usuario AS "Usuário",
    codigo_barras AS "Código Barras",
    CASE WHEN restringir THEN 'Sim' ELSE 'Não' END AS "Restringir",
    notas_fiscais AS "Notas Fiscais",
    origem_cancelamento AS "Origem Cancelamento",
    motivo_cancelamento AS "Motivo Cancelamento",
    usuario_autorizacao_pagamento AS "Usuário Autorização Pgto.",
    TO_CHAR(data_hora_autorizacao_pagamento, 'DD/MM/YYYY HH24:MI:SS') AS "Data Hora Autorização Pgto.",
    CASE WHEN conciliada_dda THEN 'Sim' ELSE 'Não' END AS "Conciliada com DDA",
    contribuicao_credito_sped AS "Contribuição e Crédito SPED",
    CASE WHEN confirmacao_pagamento_automatico THEN 'Sim' ELSE 'Não' END AS "Confirmação Pagamento Automático",
    documento_credor AS "Documento do Credor"
FROM base
ORDER BY
    data_vencimento,
    unidade,
    credor,
    numero_documento;
