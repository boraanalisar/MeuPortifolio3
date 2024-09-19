import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog

# Definir o caminho do arquivo SPED Contribuições
#arquivo_sped = 'c:/temp/PISCOFINS202409090951.txt'

colunas_sped  = {
    "Bloco 0": {
      "0000": "'REG', 'COD_VER', 'TIPO_ESCRIT', 'DT_INI', 'DT_FIN', 'IND_NAT_PJ', 'IND_ATIV', 'NOME', 'CNPJ', 'UF', 'COD_MUN', 'SUFRAMA', 'IND_PERFIL', 'IND_ATIV'",
        "0001": "'REG', 'IND_MOV'",  # Abertura do Bloco 0
        "0035": "'REG', 'NOME_SCP', 'CNPJ_SCP'",  # Identificação de SCP
        "0100": "'REG', 'NOME', 'CPF', 'CRC', 'CNPJ', 'CEP', 'END', 'NUM', 'COMPL', 'BAIRRO', 'FONE', 'FAX', 'EMAIL', 'COD_MUN'",  # Dados do Contabilista
        "0110": "'REG', 'COD_INC_TRIB', 'IND_APRO_CRED', 'COD_TIPO_CONT', 'IND_REG_CUM'",  # Regimes de Apuração da Contribuição
        "0111": "'REG', 'REC_BRU_NCUM_TRIB_MI', 'REC_BRU_NCUM_NT_MI', 'REC_BRU_NCUM_EXP', 'REC_BRU_CUM', 'REC_BRU_TOTAL'",  # Receita Bruta Mensal
        "0140": "'REG', 'COD_EST', 'NOME', 'CNPJ', 'UF', 'IE', 'COD_MUN', 'IM', 'SUFRAMA'",  # Cadastro de Estabelecimentos
        "0150": "'REG', 'COD_PART', 'NOME', 'COD_PAIS', 'CNPJ', 'CPF', 'IE', 'COD_MUN', 'SUFRAMA', 'END', 'NUM', 'COMPL', 'BAIRRO'",  # Cadastro de Participantes
        "0190": "'REG', 'UNID', 'DESCR'",  # Identificação das Unidades de Medida
        "0200": "'REG', 'COD_ITEM', 'DESCR_ITEM', 'COD_BARRA', 'COD_ANT_ITEM', 'UNID_INV', 'TIPO_ITEM', 'COD_NCM', 'EX_IPI', 'COD_GEN', 'COD_LST', 'ALIQ_ICMS'",  # Identificação dos Itens (Produtos e Serviços)
        "0400": "'REG', 'COD_NAT', 'DESCR_NAT'",  # Natureza da Operação/Prestação
        "0450": "'REG', 'COD_INF', 'TXT'",  # Tabela de Informação Complementar do Documento Fiscal
        "0500": "'REG', 'DT_ALT', 'COD_NAT_CC', 'IND_CTA', 'NIVEL', 'COD_CTA', 'NOME_CTA', 'COD_CTA_REF', 'CNPJ_EST'",  # Plano de Contas Contábeis
        "0600": "'REG', 'DT_ALT', 'COD_CCUS', 'CCUS'",  # Centro de Custos
        "0900": "'REG', 'REC_TOTAL_BLOCO_A', 'REC_NRB_BLOCO_A', 'REC_TOTAL_BLOCO_C', 'REC_NRB_BLOCO_C', 'REC_TOTAL_BLOCO_D', 'REC_NRB_BLOCO_D', 'REC_TOTAL_BLOCO_F', 'REC_NRB_BLOCO_F', 'REC_TOTAL_BLOCO_I', 'REC_NRB_BLOCO_I', 'REC_TOTAL_BLOCO_1', 'REC_NRB_BLOCO_1', 'REC_TOTAL_PERIODO', 'REC_TOTAL_NRB_PERIODO'",  # Composição das Receitas do Período
        "0990": "'REG', 'QTD_LIN_0'"  # Encerramento do Bloco 0
    },
    "Bloco A": {
        "A001": "'REG', 'IND_MOV'",  # Abertura do Bloco A
        "A010": "'REG', 'CNPJ'",  # Identificação do Estabelecimento
        "A100": "'REG', 'IND_OPER', 'IND_EMIT', 'COD_PART', 'COD_MOD', 'COD_SIT', 'SER', 'SUB', 'NUM_DOC', 'CHV_NFE', 'DT_DOC', 'DT_EXE_SERV', 'VL_DOC', 'IND_PGTO', 'VL_DESC', 'VL_BC_PIS', 'VL_PIS', 'VL_BC_COFINS', 'VL_COFINS', 'VL_PIS_RET', 'VL_COFINS_RET', 'VL_ISS'",  # Documento - Nota Fiscal de Serviço
        "A110": "'REG', 'COD_INF', 'TXT_COMPL'",  # Complemento do Documento - Informação Complementar da NF
        "A170": "'REG', 'NUM_ITEM', 'COD_ITEM', 'DESCR_COMPL', 'VL_ITEM', 'VL_DESC', 'NAT_BC_CRED', 'IND_ORIG_CRED', 'CST_PIS', 'VL_BC_PIS', 'ALIQ_PIS', 'VL_PIS', 'CST_COFINS', 'VL_BC_COFINS', 'ALIQ_COFINS', 'VL_COFINS', 'COD_CTA', 'COD_CCUS'",  # Complemento do Documento - Itens do Documento
        "A990": "'REG', 'QTD_LIN_A'"  # Encerramento do Bloco A
    },
    "Bloco C": {
        "C001": "'REG', 'IND_MOV'",  # Abertura do Bloco C
        "C010": "'REG', 'CNPJ', 'IND_ESCRI'",  # Identificação do Estabelecimento
        "C100": "'REG', 'IND_OPER', 'IND_EMIT', 'COD_PART', 'COD_MOD', 'COD_SIT', 'SER', 'SUB', 'NUM_DOC', 'CHV_NFE', 'DT_DOC', 'DT_E_S', 'VL_DOC', 'IND_PGTO', 'VL_DESC', 'VL_ABAT_NT', 'VL_MERC', 'VL_FRT', 'VL_SEG', 'VL_OUT_DA', 'VL_BC_ICMS', 'VL_ICMS', 'VL_BC_ICMS_ST', 'VL_ICMS_ST', 'VL_IPI', 'VL_PIS', 'VL_COFINS', 'VL_PIS_ST', 'VL_COFINS_ST'",  # Documento - Nota Fiscal
        "C110": "'REG', 'COD_INF', 'TXT_COMPL'",  # Informação Complementar da Nota Fiscal
        "C170": "'REG', 'NUM_ITEM', 'COD_ITEM', 'DESCR_COMPL', 'QTD', 'UNID', 'VL_ITEM', 'VL_DESC', 'IND_MOV_FISICO', 'CST_ICMS', 'CFOP', 'COD_NAT', 'VL_BC_ICMS', 'ALIQ_ICMS', 'VL_ICMS', 'VL_BC_ICMS_ST', 'ALIQ_ST', 'VL_ICMS_ST', 'IND_APUR', 'CST_IPI', 'COD_ENQ', 'VL_BC_IPI', 'ALIQ_IPI', 'VL_IPI', 'CST_PIS', 'VL_BC_PIS', 'ALIQ_PIS', 'QUANT_BC_PIS', 'ALIQ_PIS_QUANT', 'VL_PIS', 'CST_COFINS', 'VL_BC_COFINS', 'ALIQ_COFINS', 'QUANT_BC_COFINS', 'ALIQ_COFINS_QUANT', 'VL_COFINS', 'COD_CTA'",  # Complemento - Itens do Documento
        "C120": "'REG', 'COD_DOC_IMP', 'NUM_DOC_IMP', 'VL_PIS_IMP', 'VL_COFINS_IMP', 'NUM_ACDRAW'", #Registro C120: Complemento do Documento - Operações de Importação (Código 01)
        "C190": "'REG', 'CST_ICMS', 'CFOP', 'ALIQ_ICMS', 'VL_OPR', 'VL_BC_ICMS', 'VL_ICMS', 'VL_BC_ICMS_ST', 'VL_ICMS_ST', 'VL_RED_BC', 'VL_IPI'",  # Resumo do Documento
        "C395": "'REG', 'COD_MOD', 'COD_PART', 'SER', 'SUB_SER', 'NUM_DOC', 'DT_DOC', 'VL_DOC'",  # Notas Fiscais de Venda a Consumidor - Aquisições/Entradas com Crédito
        "C400": "'REG', 'COD_MOD', 'ECF_MOD', 'ECF_FAB', 'ECF_CX'",  # Equipamento ECF
        "C405": "'REG', 'DT_DOC', 'CRO', 'CRZ', 'NUM_COO_FIN', 'GT_FIN', 'VL_BRT'",  # Redução Z
        "C481": "'REG', 'CST_PIS', 'VL_ITEM', 'VL_BC_PIS', 'ALIQ_PIS', 'QUANT_BC_PIS', 'ALIQ_PIS_QUANT', 'VL_PIS', 'COD_ITEM', 'COD_CTA'",  # Resumo Diário de Documentos Emitidos por ECF – PIS/Pasep
        "C485": "'REG', 'CST_COFINS', 'VL_ITEM', 'VL_BC_COFINS', 'ALIQ_COFINS', 'QUANT_BC_COFINS', 'ALIQ_COFINS_QUANT', 'VL_COFINS', 'COD_ITEM', 'COD_CTA'",  # Resumo Diário de Documentos Emitidos por ECF – Cofins
        "C500": "'REG', 'COD_PART', 'COD_MOD', 'COD_SIT', 'SER', 'SUB', 'NUM_DOC', 'DT_DOC', 'DT_ENT', 'VL_DOC', 'VL_ICMS', 'COD_INF', 'VL_PIS', 'VL_COFINS', 'CHV_DOCe'",  # Nota Fiscal de Energia Elétrica
        "C501": "'REG', 'CST_PIS', 'VL_ITEM', 'NAT_BC_CRED', 'VL_BC_PIS', 'ALIQ_PIS', 'VL_PIS', 'COD_CTA'", #Complemento da Operação (Códigos 06, 28 e 29) – PIS/Pasep
        "C505": "'REG', 'CST_COFINS', 'VL_ITEM', 'NAT_BC_CRED', 'VL_BC_COFINS', 'ALIQ_COFINS', 'VL_COFINS', 'COD_CTA'", #C505: Complemento da Operação (Códigos 06, 28 e 29) – Cofins
        "C600": "'REG', 'COD_MOD', 'COD_PART', 'SER', 'NUM_DOC', 'DT_DOC', 'VL_DOC', 'VL_DESC', 'VL_PIS', 'VL_COFINS'",  # Nota Fiscal de Serviço de Comunicação/Telecomunicação
        "C601": "'REG', 'CST_PIS', 'VL_ITEM', 'VL_BC_PIS', 'ALIQ_PIS', 'VL_PIS'",  # Detalhamento PIS
        "C605": "'REG', 'CST_COFINS', 'VL_ITEM', 'VL_BC_COFINS', 'ALIQ_COFINS', 'VL_COFINS'",  # Detalhamento COFINS
        "C800": "'REG', 'IND_OPER', 'COD_PART', 'COD_MOD', 'COD_SIT', 'SER', 'NUM_DOC', 'CHV_DOCe', 'DT_DOC', 'VL_DOC', 'VL_DESC', 'VL_SERV'",  # Nota Fiscal de Serviço de Transporte
        "C810": "'REG', 'COD_MOD', 'SER', 'NUM_DOC', 'DT_DOC', 'VL_DOC'",  # Resumo Diário de Nota Fiscal de Serviço de Transporte
        "C860": "'REG', 'COD_MOD', 'SER', 'NUM_DOC', 'DT_DOC', 'VL_DOC'",  # Resumo de Documentos Emitidos por SAT
        "C870": "'REG', 'COD_MOD', 'SER', 'NUM_DOC', 'DT_DOC', 'VL_DOC'",  # Resumo Diário de Documentos Emitidos por SAT
        "C990": "'REG', 'QTD_LIN_C'"  # Encerramento do Bloco C
    },
    "Bloco D": {
        "D001": "'REG', 'IND_MOV'",  # Abertura do Bloco D
        "D010": "'REG', 'CNPJ'",  # Identificação do Estabelecimento
        "D100": "'REG', 'IND_OPER', 'IND_EMIT', 'COD_PART', 'COD_MOD', 'COD_SIT', 'SER', 'SUB', 'NUM_DOC', 'CHV_CTE', 'DT_DOC', 'DT_A_P', 'TP_CT-e', 'CHV_CTE_REF', 'VL_DOC', 'VL_DESC', 'IND_FRT', 'VL_SERV', 'VL_BC_ICMS', 'VL_ICMS', 'VL_NT', 'COD_INF', 'COD_CTA'",  # Aquisição de Serviços de Transporte
        "D101": "'REG', 'IND_NAT_FRT', 'VL_ITEM', 'CST_PIS', 'NAT_BC_CRED', 'VL_BC_PIS', 'ALIQ_PIS', 'VL_PIS', 'COD_CTA'",  # Complemento do Documento de Transporte - PIS/PASEP
        "D105": "'REG', 'IND_NAT_FRT', 'VL_ITEM', 'CST_COFINS', 'NAT_BC_CRED', 'VL_BC_COFINS', 'ALIQ_COFINS', 'VL_COFINS', 'COD_CTA'",  # Complemento do Documento de Transporte - COFINS
        "D111": "'REG', 'NUM_PROC', 'IND_PROC'",  # Processo Referenciado
        "D200": "'REG', 'COD_MOD', 'SER', 'SUB', 'NUM_DOC_INI', 'NUM_DOC_FIN', 'CFOP', 'DT_REF', 'VL_DOC', 'VL_DESC'",  # Resumo da Escrituração Diária – Serviços de Transporte
        "D201": "'REG', 'CST_PIS', 'VL_ITEM', 'VL_BC_PIS', 'ALIQ_PIS', 'VL_PIS', 'COD_CTA'",  # Totalização do Resumo Diário – PIS/PASEP
        "D205": "'REG', 'CST_COFINS', 'VL_ITEM', 'VL_BC_COFINS', 'ALIQ_COFINS', 'VL_COFINS', 'COD_CTA'",  # Totalização do Resumo Diário – COFINS
        "D300": "'REG', 'COD_MOD', 'SER', 'SUB', 'NUM_DOC_INI', 'NUM_DOC_FIN', 'CFOP', 'DT_REF', 'VL_DOC', 'VL_DESC'",  # Resumo da Escrituração Diária – Serviços de Transporte
        "D500": "'REG', 'IND_OPER', 'IND_EMIT', 'COD_PART', 'COD_MOD', 'COD_SIT', 'SER', 'SUB', 'NUM_DOC', 'DT_DOC', 'DT_A_P', 'VL_DOC', 'VL_DESC', 'VL_SERV', 'VL_SERV_NT', 'VL_TERC', 'VL_DA', 'VL_BC_ICMS', 'VL_ICMS', 'COD_INF', 'VL_PIS', 'VL_COFINS'", # Nota Fiscal de Serviço de Comunicação e Telecomunicação
        "D501": "'REG', 'CST_PIS', 'VL_ITEM', 'NAT_BC_CRED', 'VL_BC_PIS', 'ALIQ_PIS', 'VL_PIS', 'COD_CTA'",
        "D505": "'REG', 'CST_COFINS', 'VL_ITEM', 'NAT_BC_CRED', 'VL_BC_COFINS', 'ALIQ_COFINS', 'VL_COFINS', 'COD_CTA'",
        "D990": "'REG', 'QTD_LIN_D'"  # Encerramento do Bloco D
    },
    "Bloco F": {
        "F001": "'REG', 'IND_MOV'",  # Abertura do Bloco F
        "F010": "'REG', 'CNPJ'",  # Identificação do Estabelecimento
        "F100": "'REG', 'IND_OPER', 'COD_PART', 'COD_ITEM', 'DT_OPER', 'VL_OPER', 'CST_PIS', 'VL_BC_PIS', 'ALIQ_PIS', 'VL_PIS', 'CST_COFINS', 'VL_BC_COFINS', 'ALIQ_COFINS', 'VL_COFINS', 'NAT_BC_CRED', 'IND_ORIG_CRED', 'COD_CTA', 'COD_CCUS', 'DESC_DOC_OPER'",  # Demais Documentos e Operações Geradoras de Contribuição e Créditos
        "F111": "'REG', 'NUM_PROC', 'IND_PROC'",  # Processo Referenciado
        "F120": "'REG', 'NAT_BC_CRED', 'IDENT_BEM_IMOB', 'IND_ORIG_CRED', 'IND_UTIL_BEM_IMOB', 'VL_OPER_DEP', 'PARC_OPER_NAO_BC_CRED', 'CST_PIS', 'VL_BC_PIS', 'ALIQ_PIS', 'VL_PIS', 'CST_COFINS', 'VL_BC_COFINS', 'ALIQ_COFINS', 'VL_COFINS', 'COD_CTA', 'COD_CCUS', 'DESC_BEM_IMOB'",  # Bens Incorporados ao Ativo Imobilizado - Depreciação e Amortização
        "F129": "'REG', 'NUM_PROC', 'IND_PROC'",  # Processo Referenciado
        "F130": "'REG', 'NAT_BC_CRED', 'IDENT_BEM_IMOB', 'IND_ORIG_CRED', 'IND_UTIL_BEM_IMOB', 'MES_OPER_AQUIS', 'VL_OPER_AQUIS', 'PARC_OPER_NAO_BC_CRED', 'VL_BC_CRED', 'IND_NR_PARC', 'CST_PIS', 'VL_BC_PIS', 'ALIQ_PIS', 'VL_PIS', 'CST_COFINS', 'VL_BC_COFINS', 'ALIQ_COFINS', 'VL_COFINS', 'COD_CTA', 'COD_CCUS', 'DESC_BEM_IMOB'",  # Créditos sobre Aquisição de Bens
        "F139": "'REG', 'NUM_PROC', 'IND_PROC'",  # Processo Referenciado
        "F150": "'REG', 'NAT_BC_CRED', 'VL_OPER_ESTOQUE'",  # Crédito Presumido sobre Estoque de Abertura
        "F200": "'REG', 'IND_OPER', 'VL_OPER', 'VL_BC_OPER', 'VL_CRED_OPER'",  # Operações da Atividade Imobiliária
        "F205": "'REG', 'VL_OPER_INC'",  # Custo Incorrido da Unidade Imobiliária
        "F210": "'REG', 'VL_OPER_ORC'",  # Custo Orçado da Unidade Imobiliária
        "F211": "'REG', 'NUM_PROC', 'IND_PROC'",  # Processo Referenciado
        "F500": "'REG', 'IND_OPER', 'VL_OPER', 'VL_BC_PIS', 'VL_PIS', 'VL_BC_COFINS', 'VL_COFINS'",  # Consolidação de Operações
        "F509": "'REG', 'NUM_PROC', 'IND_PROC'",  # Processo Referenciado
        "F510": "'REG', 'COD_MOD', 'SER', 'NUM_DOC', 'DT_OPER', 'VL_DOC', 'VL_DESC', 'VL_BC_PIS', 'VL_PIS', 'VL_BC_COFINS', 'VL_COFINS'",  # Apuração por Unidade de Medida
        "F559": "'REG', 'NUM_PROC', 'IND_PROC'",  # Processo Referenciado
        "F600": "'REG', 'COD_RET', 'VL_RET'",  # Contribuição Retida na Fonte
        "F700": "'REG', 'VL_DEDUCAO'",  # Deduções Diversas
        "F800": "'REG', 'COD_EVENTO_SUCESSAO', 'DT_EVENTO_SUCESSAO', 'CNPJ_SUCESSOR', 'PER_APUR_CRED', 'COD_TIPO_CREDITO', 'VL_CREDITO_PIS_SUCESSAO', 'VL_CREDITO_COFINS_SUCESSAO'",  # Créditos por Incorporação/Fusão/Cisão
        "F990": "'REG', 'QTD_LIN_F'"  # Encerramento do Bloco F
    },
    "Bloco I": {
        "I001": "'REG', 'IND_MOV'",  # Abertura do Bloco I
        "I010": "'REG', 'CNPJ', 'IND_ATIV', 'INFO_COMPL'",  # Identificação da Pessoa Jurídica/Estabelecimento
        "I100": "'REG', 'VL_REC', 'CST_PIS_COFINS', 'VL_TOT_DED_GER', 'VL_TOT_DED_ESP', 'VL_BC_PIS', 'ALIQ_PIS', 'VL_PIS', 'VL_BC_COFINS', 'ALIQ_COFINS', 'VL_COFINS', 'INFO_COMPL'",  # Consolidação das Operações do Período
        "I199": "'REG', 'NUM_PROC', 'IND_PROC'",  # Processo Referenciado
        "I200": "'REG', 'NUM_CAMPO', 'COD_DET', 'DET_VALOR', 'COD_CTA', 'INFO_COMPL'",  # Composição das Receitas, Deduções e/ou Exclusões do Período
        "I299": "'REG', 'NUM_PROC', 'IND_PROC'",  # Processo Referenciado
        "I300": "'REG', 'COD_COMP', 'DET_VALOR', 'COD_CTA', 'INFO_COMPL'",  # Complemento das Operações - Detalhamento das Receitas, Deduções e/ou Exclusões do Período
        "I399": "'REG', 'NUM_PROC', 'IND_PROC'",  # Processo Referenciado
        "I990": "'REG', 'QTD_LIN_I'"  # Encerramento do Bloco I
    },
    "Bloco M": {
        "M001": "'REG', ' IND_MOV'",
        # Registro M100 - Apuração do Crédito de PIS/Pasep no Período
        "M100": "'REG', 'COD_CRED', 'IND_CRED_ORI', 'VL_BC_PIS', 'ALIQ_PIS', 'QUANT_BC_PIS', 'ALIQ_PIS_QUANT', 'VL_CRED', 'VL_AJUS_ACRES', 'VL_AJUS_REDUC', 'VL_CRED_DIF', 'VL_CRED_DISP', 'IND_DESC_CRED', 'VL_CRED_DESC', 'SLD_CRED'",

        # Registro M105 - Detalhamento da Base de Cálculo do Crédito Apurado no Período
        "M105": "'REG', 'NAT_BC_CRED', 'CST_PIS', 'VL_BC_PIS_TOT', 'VL_BC_PIS_CUM', 'VL_BC_PIS_NC', 'VL_BC_PIS', 'QUANT_BC_PIS_TOT', 'QUANT_BC_PIS', 'DESC_CRED'",

        # Registro M110 - Ajustes de Acréscimo e de Redução da Contribuição Apurada
        "M110": "'REG', 'IND_AJ', 'VL_AJ', 'COD_AJ', 'NUM_DOC', 'DESCR_AJ', 'DT_REF'",

        # Registro M115 - Detalhamento dos Ajustes de Acréscimo e de Redução da Contribuição Apurada
        "M115": "'REG', 'DET_VALOR_AJ', 'CST_PIS', 'DET_BC_CRED', 'DET_ALIQ', 'DT_OPER_AJ', 'DESC_AJ', 'COD_CTA', 'INFO_COMPL'",

        # Registro M200 - Consolidação da Contribuição para o PIS/Pasep do Período
        "M200": "'REG', 'VL_TOT_CONT_NC_PER', 'VL_TOT_CRED_DESC', 'VL_TOT_CRED_DESC_ANT', 'VL_TOT_CONT_NC_DEV', 'VL_RET_NC', 'VL_OUT_DED_NC', 'VL_CONT_NC_REC', 'VL_TOT_CONT_CUM_PER', 'VL_RET_CUM', 'VL_OUT_DED_CUM', 'VL_CONT_CUM_REC', 'VL_TOT_CONT_REC'",

        # Registro M205 - Contribuição para o PIS/Pasep a Recolher
        "M205": "'REG', 'NUM_CAMPO', 'COD_REC', 'VL_DEBITO'",

        # Registro M210 - Apuração da Contribuição para o PIS/Pasep a Recolher – Detalhamento por Código de Receita
        "M210": "'REG', 'COD_CONT', 'VL_REC_BRT', 'VL_BC_CONT', 'VL_AJUS_ACRES_BC_PIS', 'VL_AJUS_REDUC_BC_PIS', 'VL_BC_CONT_AJUS', 'ALIQ_PIS', 'QUANT_BC_PIS', 'ALIQ_PIS_QUANT', 'VL_CONT_APUR', 'VL_AJUS_ACRES', 'VL_AJUS_REDUC', 'VL_CONT_DIFER', 'VL_CONT_DIFER_ANT', 'VL_CONT_PER'",

        # Registro M220 - Ajustes de Redução e de Acréscimo da Contribuição Apurada
        "M220": "'REG', 'IND_AJ', 'VL_AJ', 'COD_AJ', 'NUM_DOC', 'DESCR_AJ', 'DT_REF'",

        # Registro M225 - Detalhamento dos Ajustes de Redução e de Acréscimo da Contribuição Apurada
        "M225": "'REG', 'DET_VALOR_AJ', 'CST_PIS', 'DET_BC_CRED', 'DET_ALIQ', 'DT_OPER_AJ', 'DESC_AJ', 'COD_CTA', 'INFO_COMPL'",

        # Registro M400 - Consolidação da Contribuição para o PIS/Pasep – Regime Cumulativo
        "M400": "'REG', 'CST_PIS', 'VL_TOT_REC', 'COD_CTA', 'DESC_COMPL'",
        "M410": "'REG', 'NAT_REC', 'VL_REC', 'COD_CTA', 'DESC_COMPL'", 

        # Registro M500 - Apuração da Contribuição de Cofins no Período
        "M500": "'REG', 'COD_CRED', 'IND_CRED_ORI', 'VL_BC_COFINS', 'ALIQ_COFINS', 'QUANT_BC_COFINS', 'ALIQ_COFINS_QUANT', 'VL_CRED', 'VL_AJUS_ACRES', 'VL_AJUS_REDUC', 'VL_CRED_DIF', 'VL_CRED_DISP', 'IND_DESC_CRED', 'VL_CRED_DESC', 'SLD_CRED'",

        # Registro M505 - Detalhamento da Base de Cálculo do Crédito Apurado no Período
        "M505": "'REG', 'NAT_BC_CRED', 'CST_COFINS', 'VL_BC_COFINS_TOT', 'VL_BC_COFINS_CUM', 'VL_BC_COFINS_NC', 'VL_BC_COFINS', 'QUANT_BC_COFINS_TOT', 'QUANT_BC_COFINS', 'DESC_CRED'",

        # Registro M510 - Ajustes de Redução e de Acréscimo da Contribuição Apurada
        "M510": "'REG', 'IND_AJ', 'VL_AJ', 'COD_AJ', 'NUM_DOC', 'DESCR_AJ', 'DT_REF'",

        # Registro M515 - Detalhamento dos Ajustes de Redução e de Acréscimo da Contribuição Apurada
        "M515": "'REG', 'DET_VALOR_AJ', 'CST_COFINS', 'DET_BC_CRED', 'DET_ALIQ', 'DT_OPER_AJ', 'DESC_AJ', 'COD_CTA', 'INFO_COMPL'",

        # Registro M600 - Consolidação da Contribuição para a Cofins do Período
        "M600": "'REG', 'VL_TOT_CONT_NC_PER', 'VL_TOT_CRED_DESC', 'VL_TOT_CRED_DESC_ANT', 'VL_TOT_CONT_NC_DEV', 'VL_RET_NC', 'VL_OUT_DED_NC', 'VL_CONT_NC_REC', 'VL_TOT_CONT_CUM_PER', 'VL_RET_CUM', 'VL_OUT_DED_CUM', 'VL_CONT_CUM_REC', 'VL_TOT_CONT_REC'",

        "M610": "'REG', 'COD_CONT', 'VL_REC_BRT', 'VL_BC_CONT', 'ALIQ_COFINS', 'QUANT_BC_COFINS', 'ALIQ_COFINS_QUANT', 'VL_CONT_APUR', 'VL_AJUS_ACRES', 'VL_AJUS_REDUC', 'VL_CONT_DIFER', 'VL_CONT_DIFER_ANT', 'VL_CONT_PER'",

        # Registro M605 - Contribuição para a Cofins a Recolher
        "M605": "'REG', 'NUM_CAMPO', 'COD_REC', 'VL_DEBITO'",

        # Registro M610 - Apuração da Contribuição para a Cofins a Recolher – Detalhamento por Código de Receita
        "M610": "'REG', 'COD_CONT', 'VL_REC_BRT', 'VL_BC_CONT', 'VL_AJUS_ACRES_BC_COFINS', 'VL_AJUS_REDUC_BC_COFINS', 'VL_BC_CONT_AJUS', 'ALIQ_COFINS', 'QUANT_BC_COFINS', 'ALIQ_COFINS_QUANT', 'VL_CONT_APUR', 'VL_AJUS_ACRES', 'VL_AJUS_REDUC', 'VL_CONT_DIFER', 'VL_CONT_DIFER_ANT', 'VL_CONT_PER'",

        # Registro M620 - Ajustes de Redução e de Acréscimo da Contribuição Apurada
        "M620": "'REG', 'IND_AJ', 'VL_AJ', 'COD_AJ', 'NUM_DOC', 'DESCR_AJ', 'DT_REF'",

        # Registro M625 - Detalhamento dos Ajustes de Redução e de Acréscimo da Contribuição Apurada
        "M625": "'REG', 'DET_VALOR_AJ', 'CST_COFINS', 'DET_BC_CRED', 'DET_ALIQ', 'DT_OPER_AJ', 'DESC_AJ', 'COD_CTA', 'INFO_COMPL'",
        "M800": "'REG', 'CST_COFINS', 'VL_TOT_REC', 'COD_CTA', 'DESC_COMPL'",
        "M810": "'REG', 'NAT_REC', 'VL_REC', 'COD_CTA', 'DESC_COMPL'",
        "M990": "'REG', 'QTD_LIN_M'",

    },
    "Bloco P": {
        "P001": "'REG', 'IND_MOV'",  # Abertura do Bloco P
        "P010": "'REG', 'CNPJ'",  # Identificação do Estabelecimento
        "P100": "'REG', 'DT_INI', 'DT_FIN', 'VL_REC_TOT_EST', 'COD_ATIV_ECON', 'VL_REC_ATIV_ESTAB', 'VL_EXC', 'VL_BC_CONT', 'ALIQ_CONT', 'VL_CONT_APU', 'COD_CTA', 'INFO_COMPL'",  # Contribuição Previdenciária sobre a Receita Bruta
        "P110": "'REG', 'NUM_CAMPO', 'COD_DET', 'DET_VALOR', 'INF_COMPL'",  # Complemento da Escrituração – Detalhamento da Apuração da Contribuição
        "P199": "'REG', 'NUM_PROC', 'IND_PROC'",  # Processo Referenciado
        "P200": "'REG', 'PER_REF', 'VL_TOT_CONT_APU', 'VL_TOT_AJ_REDUC', 'VL_TOT_AJ_ACRES', 'VL_TOT_CONT_DEV', 'COD_REC'",  # Consolidação da Contribuição Previdenciária Sobre a Receita Bruta
        "P210": "'REG', 'IND_AJ', 'VL_AJ', 'COD_AJ', 'NUM_DOC', 'DESCR_AJ', 'DT_REF'",  # Ajuste à Contribuição Previdenciária Apurada
        "P990": "'REG', 'QTD_LIN_P'"  # Encerramento do Bloco P
    },
    "Bloco 1": {
        "1001": "'REG', 'IND_MOV'",  # Abertura do Bloco 1
        "1010": "'REG', 'NUM_PROC', 'ID_SEC_JUD', 'ID_VARA', 'IND_NAT_ACAO'",  # Processo Referenciado – Ação Judicial
        "1011": "'REG', 'VL_PIS_SUSP', 'VL_COFINS_SUSP'",  # Detalhamento das Contribuições com Exigibilidade Suspensa
        "1020": "'REG', 'NUM_PROC', 'IND_NAT_ACAO', 'DT_DEC_ADM'",  # Processo Referenciado – Processo Administrativo
        "1050": "'REG', 'DT_REF', 'IND_AJ_BC', 'CNPJ', 'VL_AJ_TOT', 'VL_AJ_CST01', 'VL_AJ_CST02', 'VL_AJ_CST03', 'VL_AJ_CST04', 'VL_AJ_CST05', 'VL_AJ_CST06', 'VL_AJ_CST07', 'VL_AJ_CST08', 'VL_AJ_CST09', 'VL_AJ_CST49', 'VL_AJ_CST99', 'IND_APROP', 'NUM_REC','INFO_COMPL'",  # Detalhamento de Ajustes de Base de Cálculo – Valores Extra Apuração
        "1100": "'REG', 'PER_APU_CRED', 'ORIG_CRED', 'CNPJ_SUC', 'COD_CRED', 'VL_CRED_APU', 'VL_CRED_EXT_APU', 'VL_TOT_CRED_APU', 'VL_CRED_DESC_PA_ANT', 'VL_CRED_PER_PA_ANT', 'VL_CRED_DCOMP_PA_ANT', 'SD_CRED_DISP_EFD', 'VL_CRED_DESC_EFD', 'VL_CRED_PER_EFD', 'VL_CRED_DCOMP_EFD', 'VL_CRED_TRANS', 'VL_CRED_OUT', 'SLD_CRED_FIM'",  # Controle de Créditos Fiscais – PIS/Pasep
        "1500": "'REG', 'PER_APU_CRED', 'ORIG_CRED', 'CNPJ_SUC', 'COD_CRED', 'VL_CRED_APU', 'VL_CRED_EXT_APU', 'VL_TOT_CRED_APU', 'VL_CRED_DESC_PA_ANT', 'VL_CRED_PER_PA_ANT', 'VL_CRED_DCOMP_PA_ANT', 'SD_CRED_DISP_EFD', 'VL_CRED_DESC_EFD', 'VL_CRED_PER_EFD', 'VL_CRED_DCOMP_EFD', 'VL_CRED_TRANS', 'VL_CRED_OUT', 'SLD_CRED_FIM'",  # Controle de Créditos Fiscais – COFINS
        "1990": "'REG', 'QTD_LIN_1'"  # Encerramento do Bloco 1
    },
    "Bloco 9": {
        "9001": "'REG', 'IND_MOV'",  # Abertura do Bloco 9
        "9900": "'REG', 'REG_BLC', 'QTD_REG_BLC'",  # Registros do Arquivo
        "9990": "'REG', 'QTD_LIN_9'",  # Encerramento do Bloco 9
        "9999": "'REG', 'QTD_LIN'"  # Encerramento do Arquivo Digital
    }
}

# Função para importar o arquivo
def importar_arquivo():
    root = tk.Tk()
    root.withdraw()  # Ocultar a janela principal
    arquivo_sped = filedialog.askopenfilename(
        title="Selecione o arquivo SPED Contribuições",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
    )
    return arquivo_sped

# Chamar a função para importar o arquivo
arquivo_sped = importar_arquivo()

# Abrir e ler o arquivo linha por linha
with open(arquivo_sped, 'r', encoding='cp1252') as f:
    linhas = f.readlines()

# Transforma o conteúdo lido em uma lista de strings, preservando os campos vazios
arq1 = [linha.strip() for linha in linhas]

# Inicializar um dicionário para armazenar os registros por tipo
registros = {}

# Processar cada linha do arquivo
for linha in arq1:
    campos = linha.split('|')

    # Preservar os campos vazios entre os pipes e ignorar a última entrada vazia (após o último pipe)
    campos = campos[:-1] if campos[-1] == '' else campos
    
    if len(campos) < 2:
        continue

    reg = campos[1]  # O segundo campo é o tipo de registro
    if reg not in registros:
        registros[reg] = []
    registros[reg].append(campos[1:])  # Preserva todos os campos incluindo os vazios

# Criar um objeto ExcelWriter para escrever no Excel usando openpyxl
with pd.ExcelWriter('c:/temp/sped_contribuicoes.xlsx', engine='openpyxl') as writer:
    for reg, dados in registros.items():
        # Criar um DataFrame para cada tipo de registro
        df_reg = pd.DataFrame(dados)

        # Obter o número de colunas
        num_cols = df_reg.shape[1]

        # Verificar se há um mapeamento de colunas para este tipo de registro no dicionário 'colunas_sped'
        colunas_definidas = None
        for bloco, registros_dicionario in colunas_sped.items():
            if reg in registros_dicionario:
                colunas_definidas = registros_dicionario[reg].replace("'", "").split(", ")
                break
       
        # Verificar se as colunas definidas no dicionário correspondem ao número de colunas do registro
        if colunas_definidas and len(colunas_definidas) == num_cols:
            df_reg.columns = colunas_definidas
        else:
            # Se não houver correspondência exata, usar nomes genéricos
            print(f"Aviso: O número de colunas no arquivo ({num_cols}) não corresponde ao mapeamento para o registro {reg} ({len(colunas_definidas) if colunas_definidas else 0} colunas). Usando nomes genéricos.")
            df_reg.columns = [f'Campo_{i}' for i in range(1, num_cols + 1)]

        # Gravar o DataFrame no arquivo Excel
        sheet_name = f'REG_{reg}'[:31]  # Garantir que o nome da aba tenha no máximo 31 caracteres
        df_reg.to_excel(writer, sheet_name=sheet_name, index=False)

print('Conversão concluída com sucesso! O arquivo sped_contribuicoes.xlsx foi gerado.')