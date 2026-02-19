SELECT
    ent.codi_emp AS EMPRESA,
    ent.codi_nat AS CFOP,
    ent.vcon_ent AS VALOR,
    (ent.vcon_ent * CAST(ISNULL(aliq.aliq_pis, 0) AS FLOAT)) / 100 AS PIS,
    (ent.vcon_ent * CAST(ISNULL(aliq.aliq_cofins, 0) AS FLOAT)) / 100 AS COFINS,
    CAST(0 AS FLOAT) AS ICMS,
    'ENTRADA' AS TIPO

FROM bethadba.efentradas ent

OUTER APPLY (
    SELECT TOP 1
        CASE 
            WHEN fov.simples_nacional = 1 THEN 0.65
            ELSE 1.65
        END AS aliq_pis,
        CASE 
            WHEN fov.simples_nacional = 1 THEN 3.0
            ELSE 7.6
        END AS aliq_cofins,
        efm.aliicms_mep AS aliq_icms
    FROM bethadba.FOVIGENCIA_REGIME fov
    JOIN bethadba.efmvepro efm ON efm.codi_emp = fov.codi_emp
    WHERE fov.codi_emp = ent.codi_emp
      AND efm.data_mep <= ent.dent_ent
    ORDER BY efm.data_mep DESC
) aliq

WHERE ent.codi_emp IN ({}) 
  AND ent.dent_ent BETWEEN CONVERT(DATE, ?, 103) AND CONVERT(DATE, ?, 103)

UNION ALL

SELECT
    sai.codi_emp AS EMPRESA,
    sai.codi_nat AS CFOP,
    sai.vcon_sai AS VALOR,
    (sai.vcon_sai * CAST(ISNULL(aliq.aliq_pis, 0) AS FLOAT)) / 100 AS PIS,
    (sai.vcon_sai * CAST(ISNULL(aliq.aliq_cofins, 0) AS FLOAT)) / 100 AS COFINS,
    CAST(0 AS FLOAT) AS ICMS,
    'SAIDA' AS TIPO

FROM bethadba.efsaidas sai

OUTER APPLY (
    SELECT TOP 1
        CASE 
            WHEN fov.simples_nacional = 1 THEN 0.65
            ELSE 1.65
        END AS aliq_pis,
        CASE 
            WHEN fov.simples_nacional = 1 THEN 3.0
            ELSE 7.6
        END AS aliq_cofins,
        efm.aliicms_msp AS aliq_icms
    FROM bethadba.FOVIGENCIA_REGIME fov
    JOIN bethadba.efmvspro efm ON efm.codi_emp = fov.codi_emp
    WHERE fov.codi_emp = sai.codi_emp
      AND efm.data_msp <= sai.dsai_sai
    ORDER BY efm.data_msp DESC
) aliq

WHERE sai.codi_emp IN ({})
  AND sai.dsai_sai BETWEEN CONVERT(DATE, ?, 103) AND CONVERT(DATE, ?, 103)

UNION ALL

SELECT
    codi_emp AS EMPRESA,
    cfop_mep AS CFOP,
    CAST(0 AS FLOAT) AS VALOR,
    CAST(0 AS FLOAT) AS PIS,
    CAST(0 AS FLOAT) AS COFINS,
    valor_icms_mep AS ICMS,
    'ENTRADA' AS TIPO
FROM bethadba.efmvepro
WHERE codi_emp IN ({})
AND data_mep BETWEEN CONVERT(DATE, ?, 103) AND CONVERT(DATE, ?, 103)
AND valor_icms_mep <> 0

UNION ALL

SELECT
    codi_emp AS EMPRESA,
    cfop_msp AS CFOP,
    CAST(0 AS FLOAT) AS VALOR,
    CAST(0 AS FLOAT) AS PIS,
    CAST(0 AS FLOAT) AS COFINS,
    valor_icms_msp AS ICMS,
    'SAIDA' AS TIPO
FROM bethadba.efmvspro
WHERE codi_emp IN ({})
AND data_msp BETWEEN CONVERT(DATE, ?, 103) AND CONVERT(DATE, ?, 103)
AND valor_icms_msp <> 0