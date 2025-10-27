import pdfplumber
import re
import os
import uuid
import tempfile
# LINHA MODIFICADA: ADICIONAMOS 'flash' aqui
from flask import Flask, request, render_template, redirect, url_for, session, send_file, flash
from werkzeug.utils import secure_filename
import json
import openpyxl
from datetime import datetime, timedelta
import zipfile
import shutil

# --- Import para manipulação de DOCX ---
from docx import Document
from docx.shared import Pt
from docx.text.run import Run # Import para manipular runs diretamente

# --- Import para manipulação de imagens e PDF ---
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader

# --- NOVIDADE AQUI: Importa e carrega variáveis de ambiente do .env ---
from dotenv import load_dotenv
load_dotenv()
# ---------------------------------------------------------------------

# Somente importa win32com se estiver no Windows
if os.name == 'nt':
    try:
        import win32com.client as win32
        import pythoncom
    except ImportError:
        print("Aviso: 'pywin32' ou 'pythoncom' não está instalado. Os cálculos de fórmulas do Excel não serão automáticos.")
        win32 = None
else:
    win32 = None

# --- FUNÇÃO AUXILIAR GLOBAL PARA FORMATAR VALORES PARA EXIBIÇÃO ---
# Esta função garante que números usem vírgula como separador decimal e datas usem DD/MM/YYYY.
def format_value_for_display(value, is_date=False, is_numeric=False, date_separator='/', numeric_decimal_separator=','):
    if value is None or value == 'Não informado' or value == 'Não calculado' or str(value).strip() == '':
        return 'Não informado'

    s_value = str(value).strip()

    if is_date:
        # Tenta analisar YYYY-MM-DD (do input HTML type='date')
        try:
            dt_obj = datetime.strptime(s_value, '%Y-%m-%d')
            return dt_obj.strftime(f'%d{date_separator}%m{date_separator}%Y')
        except ValueError:
            # Tenta analisar DD/MM/YYYY (de cálculos internos ou fallback)
            try:
                dt_obj = datetime.strptime(s_value, '%d/%m/%Y')
                return dt_obj.strftime(f'%d{date_separator}%m{date_separator}%Y')
            except ValueError:
                # Tenta analisar DD-MM-YYYY (de cálculos internos ou fallback)
                try:
                    dt_obj = datetime.strptime(s_value, '%d-%m-%Y')
                    return dt_obj.strftime(f'%d{date_separator}%m{date_separator}%Y')
                except ValueError:
                    return s_value # Se não for possível analisar como data, retorna o valor original
    
    if is_numeric:
        try:
            # Converte para float, aceitando tanto vírgula quanto ponto como separador decimal
            f_value = float(s_value.replace(',', '.'))
            
            # Formata para uma string, garantindo vírgula e removendo zeros e ponto/vírgula finais desnecessários
            # Ex: 10.0 -> '10'
            # Ex: 12.53500 -> '12.535' -> '12,535'
            # Ex: 0.545 -> '0.545' -> '0,545'
            if f_value == int(f_value): # Se for um número inteiro (ex: 10.0), exibe como inteiro
                return str(int(f_value))
            else:
                # Limita a 3 casas decimais (ou mais se necessário), remove zeros à direita e o ponto/vírgula final
                formatted_str = f"{f_value:.3f}".rstrip('0')
                if formatted_str.endswith('.'):
                    formatted_str = formatted_str.rstrip('.')
                return formatted_str.replace('.', numeric_decimal_separator)
        except ValueError:
            return s_value # Se não for um número válido, retorna o valor original

    return s_value # Por padrão, retorna o valor como string


# --- Funções Auxiliares de Extração para Layouts Específicos ---

def _extrair_dados_layout_adriano_style(texto, caminho_pdf):
    """
    Extrai dados de faturas com o layout 'Adriano Sinigaglia.pdf' (DANF3E, CPF Completo).
    Foco nos campos essenciais e genéricos para este layout.
    """
    dados_extraidos = {
        'Nome_Razao_Social': 'Não encontrado',
        'Endereco_Rua_Numero': 'Não encontrado',
        'Bairro': 'Não encontrado',
        'CNPJ_CPF': 'Não encontrado',
        'Cidade': 'Não encontrado',
        'CEP': 'Não encontrado',
        'UC': 'Não encontrado',
        'Grupo_Tarifario': 'Não encontrado',
        'Tensao_Nominal_V': 'Não encontrado',
        'Classe_Tarifaria': 'Não encontrado',
        'Estado': 'Não encontrado',
    }

    try:
        # TENSÃO (Tensão Nominal em Volts)
        match_tensao = re.search(r'TENSÃO NOMINAL EM VOLTS\s*Disp\.:\s*(\d+)', texto)
        if match_tensao:
            dados_extraidos['Tensao_Nominal_V'] = int(match_tensao.group(1))

        # NOME/RAZÃO SOCIAL: Abaixo do CNPJ da distribuidora (RGE)
        match_nome = re.search(r'Inscrição no CNPJ: \d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}\n+([A-Z\s,.]+)\n', texto)
        customer_name_found = None
        if match_nome:
            customer_name_found = match_nome.group(1).strip()
            dados_extraidos['Nome_Razao_Social'] = customer_name_found

        # ENDEREÇO (Rua e Número), BAIRRO, CEP, CIDADE, ESTADO
        if customer_name_found and customer_name_found != 'Não encontrado':
            street_and_number_pattern = r'((?:R|AV|EST|ROD|AL|TV|PR|TR|VD|RUA|VL|PRC|PCA)\s+[A-Z\s,.-]+?\s*\d+\s*(?:[A-Z0-9\s,.-]+)?)'

            address_block_full_regex = (
                re.escape(customer_name_found) + r'.*?' 
                + street_and_number_pattern + r'\n' 
                + r'([A-Z\s,.-]+)\n' 
                + r'(\d{5}-\d{3})\s+([A-Z\s,.-]+)\s+(RS)'
            )

            match_endereco_bloco = re.search(address_block_full_regex, texto, re.DOTALL)

            if match_endereco_bloco:
                dados_extraidos['Endereco_Rua_Numero'] = match_endereco_bloco.group(1).strip()
                dados_extraidos['Bairro'] = match_endereco_bloco.group(2).strip()
                dados_extraidos['CEP'] = match_endereco_bloco.group(3).strip()
                dados_extraidos['Cidade'] = match_endereco_bloco.group(4).strip()
                dados_extraidos['Estado'] = match_endereco_bloco.group(5).strip()

        # CNPJ/CPF (prioriza CPF, depois CNPJ)
        match_cpf = re.search(r'CPF:\s*(\d{3}\.\d{3}\.\d{3}-\d{2})', texto)
        if match_cpf:
            dados_extraidos['CNPJ_CPF'] = match_cpf.group(1)
        else:
            match_cnpj = re.search(r'CNPJ:\s*(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})', texto)
            if match_cnpj:
                dados_extraidos['CNPJ_CPF'] = match_cnpj.group(1)


        # UC: Tenta encontrar 'UC: ' explicitamente primeiro, senão busca após 'Lim. máx.:'
        match_uc = re.search(r'UC:\s*(\d{10})', texto)
        if match_uc:
            dados_extraidos['UC'] = match_uc.group(1)
        else:
            match_uc_alt = re.search(r'Lim.\s*máx.:\s*\d+\s*(\d{10})', texto)
            if match_uc_alt:
                dados_extraidos['UC'] = match_uc_alt.group(1)

        # GRUPO e CLASSE: Da linha de Classificação (regex mais flexível)
        match_classificacao = re.search(r'Classificaç(?:ão|ao):\s*([^\n]+)', texto, re.IGNORECASE)
        if match_classificacao:
            classif = match_classificacao.group(1).strip()
            classif = re.sub(r'\s*Tipo de Fornecimento:\s*$', '', classif)

            match_grupo = re.search(r'(B[1-4]|A)', classif)
            if match_grupo:
                dados_extraidos['Grupo_Tarifario'] = match_grupo.group(1)

            match_classe = re.search(r'(Residencial|Comercial|Industrial|Rural|Poder Público|Iluminação Pública)', classif, re.IGNORECASE) # Adicionado mais classes
            if match_classe:
                dados_extraidos['Classe_Tarifaria'] = match_classe.group(1)

    except Exception as e:
        print(f"Erro ao extrair dados do layout 'Adriano Style': {e}")
        pass

    return dados_extraidos


def _extrair_dados_layout_adroaldo_style(texto, caminho_pdf):
    """
    Extrai dados de faturas com o layout 'Adroaldo.pdf' e 'Aire.pdf' (DANF3E, CPF Mascarado).
    Este é um layout híbrido.
    """
    dados_extraidos = {
        'Nome_Razao_Social': 'Não encontrado',
        'Endereco_Rua_Numero': 'Não encontrado',
        'Bairro': 'Não encontrado',
        'CNPJ_CPF': 'Não encontrado',
        'Cidade': 'Não encontrado',
        'CEP': 'Não encontrado',
        'UC': 'Não encontrado',
        'Grupo_Tarifario': 'Não encontrado',
        'Tensao_Nominal_V': 'Não encontrado',
        'Classe_Tarifaria': 'Não encontrado',
        'Estado': 'Não encontrado',
    }

    try:
        # TENSÃO (Tensão Nominal em Volts)
        match_tensao = re.search(r'TENSÃO NOMINAL EM VOLTS\s*Disp\.:\s*(\d+)', texto)
        if match_tensao:
            dados_extraidos['Tensao_Nominal_V'] = int(match_tensao.group(1))

        # NOME/RAZÃO SOCIAL: Abaixo do CNPJ da distribuidora (RGE)
        match_nome = re.search(r'Inscrição no CNPJ: \d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}\n+([A-Z\s,.]+)\n', texto)
        customer_name_found = None
        if match_nome:
            customer_name_found = match_nome.group(1).strip()
            dados_extraidos['Nome_Razao_Social'] = customer_name_found

        # ENDEREÇO (Rua e Número), BAIRRO, CEP, CIDADE, ESTADO
        if customer_name_found and customer_name_found != 'Não encontrado':
            street_and_number_pattern = r'((?:R|AV|EST|ROD|AL|TV|PR|TR|VD|RUA|VL|PRC|PCA)\s+[A-Z\s,.-]+?\s*\d+\s*(?:[A-Z0-9\s,.-]+)?)'

            address_block_full_regex = (
                re.escape(customer_name_found) + r'.*?' 
                + street_and_number_pattern + r'\n' 
                + r'([A-Z\s,.-]+)\n' 
                + r'(\d{5}-\d{3})\s+([A-Z\s,.-]+)\s+(RS)'
            )

            match_endereco_bloco = re.search(address_block_full_regex, texto, re.DOTALL)

            if match_endereco_bloco:
                dados_extraidos['Endereco_Rua_Numero'] = match_endereco_bloco.group(1).strip()
                dados_extraidos['Bairro'] = match_endereco_bloco.group(2).strip()
                dados_extraidos['CEP'] = match_endereco_bloco.group(3).strip()
                dados_extraidos['Cidade'] = match_endereco_bloco.group(4).strip()
                dados_extraidos['Estado'] = match_endereco_bloco.group(5).strip()

        # CNPJ/CPF (prioriza CPF mascarado, depois CNPJ completo)
        match_cpf_masked = re.search(r'CPF:\s*(((\*)?){6}\.\d{3}-(((\*)?){2}))', texto) # Corrigido regex para CPF mascarado
        if match_cpf_masked:
            dados_extraidos['CNPJ_CPF'] = match_cpf_masked.group(1)
        else:
            match_cnpj = re.search(r'CNPJ:\s*(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})', texto)
            if match_cnpj:
                dados_extraidos['CNPJ_CPF'] = match_cnpj.group(1)


        # UC: No Adroaldo.pdf e Aire.pdf, a UC está após 'Lim. máx.:' (10 dígitos)
        match_uc = re.search(r'Lim.\s*máx.:\s*\d+\s*(\d{10})', texto)
        if match_uc:
            dados_extraidos['UC'] = match_uc.group(1)

        # GRUPO e CLASSE: Da linha de Classificação (regex mais flexível)
        match_classificacao = re.search(r'Classificaç(?:ão|ao):\s*([^\n]+)', texto, re.IGNORECASE)
        if match_classificacao:
            classif = match_classificacao.group(1).strip()
            classif = re.sub(r'\s*Tipo de Fornecimento:\s*$', '', classif) # Remover a parte "Tipo de Fornecimento" se estiver junto

            match_grupo = re.search(r'(B[1-4]|A)', classif)
            if match_grupo:
                dados_extraidos['Grupo_Tarifario'] = match_grupo.group(1)

            match_classe = re.search(r'(Residencial|Comercial|Industrial|Rural|Poder Público|Iluminação Pública)', classif, re.IGNORECASE)
            if match_classe:
                dados_extraidos['Classe_Tarifaria'] = match_classe.group(1)

    except Exception as e:
        print(f"Erro ao extrair dados do layout 'Adroaldo Style': {e}")
        pass

    return dados_extraidos


def _extrair_dados_layout_arcindo_style(texto, caminho_pdf):
    """
    Extrai dados de faturas com o layout 'conta-completa.pdf' (DANFE, CPF Mascarado).
    Foco nos campos essenciais e genéricos para este layout.
    """
    dados_extraidos = {
        'Nome_Razao_Social': 'Não encontrado',
        'Endereco_Rua_Numero': 'Não encontrado',
        'Bairro': 'Não encontrado',
        'CNPJ_CPF': 'Não encontrado',
        'Cidade': 'Não encontrado',
        'CEP': 'Não encontrado',
        'UC': 'Não encontrado',
        'Grupo_Tarifario': 'Não encontrado',
        'Tensao_Nominal_V': 'Não encontrado',
        'Classe_Tarifaria': 'Não encontrado',
        'Estado': 'Não encontrado',
    }

    try:
        # TENSÃO (Tensão Nominal em Volts)
        match_tensao = re.search(r'TENSÃO NOMINAL EM VOLTS\s*Disp\.:\s*(\d+)', texto)
        if match_tensao:
            dados_extraidos['Tensao_Nominal_V'] = int(match_tensao.group(1))

        # NOME/RAZÃO SOCIAL: Abaixo de "CÓDIGO DA UNIDADE CONSUMIDORA:"
        match_nome = re.search(r'CÓDIGO DA UNIDADE CONSUMIDORA:\s*\d+\n([A-Z\s]+)\n', texto)
        customer_name_found = None
        if match_nome:
            customer_name_found = match_nome.group(1).strip()
            dados_extraidos['Nome_Razao_Social'] = customer_name_found

        # ENDEREÇO (Rua e Número), BAIRRO, CEP, CIDADE, ESTADO
        if customer_name_found and customer_name_found != 'Não encontrado':
            street_and_number_pattern = r'((?:R|AV|EST|ROD|AL|TV|PR|TR|VD|RUA|VL|PRC|PCA)\s+[A-Z\s,.-]+?\s*\d+\s*(?:[A-Z0-9\s,.-]+)?)'

            # Arcindo tem "-" entre cidade e estado
            address_block_full_regex = (
                re.escape(customer_name_found) + r'.*?' 
                + street_and_number_pattern + r'\n' 
                + r'([A-Z\s,.-]+)\n' 
                + r'(\d{5}-\d{3})\s+([A-Z\s,.-]+)\s*-\s*(RS)'
            )

            match_endereco_bloco = re.search(address_block_full_regex, texto, re.DOTALL)

            if match_endereco_bloco:
                dados_extraidos['Endereco_Rua_Numero'] = match_endereco_bloco.group(1).strip()
                dados_extraidos['Bairro'] = match_endereco_bloco.group(2).strip()
                dados_extraidos['CEP'] = match_endereco_bloco.group(3).strip()
                dados_extraidos['Cidade'] = match_endereco_bloco.group(4).strip()
                dados_extraidos['Estado'] = match_endereco_bloco.group(5).strip()

        # CNPJ/CPF (prioriza CPF mascarado, depois CNPJ completo)
        match_cpf_masked = re.search(r'CPF:\s*(((\*)?){6}\.\d{3}-(((\*)?){2}))', texto) # Corrigido regex para CPF mascarado
        if match_cpf_masked:
            dados_extraidos['CNPJ_CPF'] = match_cpf_masked.group(1)
        else:
            match_cnpj = re.search(r'CNPJ:\s*(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})', texto)
            if match_cnpj:
                dados_extraidos['CNPJ_CPF'] = match_cnpj.group(1)

        # UC: Após "CÓDIGO DA UNIDADE CONSUMIDORA:" ou antes de "1/2"
        match_uc = re.search(r'CÓDIGO DA UNIDADE CONSUMIDORA:\s*(\d{10})', texto)
        if match_uc:
            dados_extraidos['UC'] = match_uc.group(1)
        else: # Fallback para o padrão antes de 1/2
            match_uc_alt = re.search(r'(\d{10})\n1/2', texto)
            if match_uc_alt:
                dados_extraidos['UC'] = match_uc_alt.group(1)

        # GRUPO e CLASSE: Da linha de Classificação (regex mais flexível)
        match_classificacao = re.search(r'Classificaç(?:ão|ao):\s*([^\n]+)', texto, re.IGNORECASE)
        if match_classificacao:
            classif = match_classificacao.group(1).strip()

            match_grupo = re.search(r'(B[1-4]|A)', classif)
            if match_grupo:
                dados_extraidos['Grupo_Tarifario'] = match_grupo.group(1)

            match_classe = re.search(r'(Residencial|Comercial|Industrial|Rural|Poder Público|Iluminação Pública)', classif, re.IGNORECASE)
            if match_classe:
                dados_extraidos['Classe_Tarifaria'] = match_classe.group(1)

    except Exception as e:
        print(f"Erro ao extrair dados do layout 'Arcindo Style': {e}")
        pass

    return dados_extraidos

# --- Funções para Sub-layouts da Cooperluz ---

def _extrair_dados_layout_cooperluz_sublayout_com_cod_ua(texto, caminho_pdf):
    """
    Extrai dados de faturas da Cooperluz com o padrão 'COD UA'. (Ex: Fatura cooper.pdf, Fatura2.pdf, Conta LUZ.pdf)
    """
    print(f"DEBUG: Entrando em _extrair_dados_layout_cooperluz_sublayout_com_cod_ua para {os.path.basename(caminho_pdf)}")
    dados_extraidos = {
        'Nome_Razao_Social': 'Não encontrado',
        'Endereco_Rua_Numero': 'Não encontrado',
        'Bairro': 'Não encontrado',
        'CNPJ_CPF': 'Não encontrado',
        'Cidade': 'Não encontrado',
        'CEP': 'Não encontrado',
        'UC': 'Não encontrado',
        'Grupo_Tarifario': 'Não encontrado',
        'Tensao_Nominal_V': 'Não encontrado',
        'Classe_Tarifaria': 'Não encontrado',
        'Estado': 'Não encontrado',
    }
    
    try:
        # --- TIPO DE FORNECIMENTO (para Tensão Nominal) ---
        print("DEBUG: Extraindo Tipo de Fornecimento...")
        tipo_fornecimento_match = re.search(
            r'Tipo de Fornecimento:\s*(?:[\s\S]*?)(Monofásico|Bifásico|Trifásico)', # Ajustado para ser mais flexível e seguro
            texto,
            re.IGNORECASE | re.DOTALL
        )
        tipo_fornecimento_extraido = ''
        if tipo_fornecimento_match:
            tipo_fornecimento_extraido = (tipo_fornecimento_match.group(1) or '').strip()
            if tipo_fornecimento_extraido:
                if 'Bifásico' in tipo_fornecimento_extraido or 'Monofásico' in tipo_fornecimento_extraido:
                    dados_extraidos['Tensao_Nominal_V'] = 220
                elif 'Trifásico' in tipo_fornecimento_extraido:
                    dados_extraidos['Tensao_Nominal_V'] = 380
        print(f"DEBUG: Tipo de Fornecimento: {dados_extraidos.get('Tensao_Nominal_V')}")


        # --- GRUPO E CLASSE TARIFÁRIA ---
        print("DEBUG: Extraindo Grupo e Classe Tarifária...")
        classificacao_line_match = re.search(
            r'Classificaç(?:ão|ao):\s*(.*?)(?:(?=\nTipo de Fornecimento)|\n|$)', # Mais robusto
            texto,
            re.DOTALL | re.IGNORECASE
        )
        classif_line_content = ''
        if classificacao_line_match:
            classif_line_content = classificacao_line_match.group(1).strip()
            
            match_grupo = re.search(r'(B[1-4]|A)', classif_line_content)
            if match_grupo:
                dados_extraidos['Grupo_Tarifario'] = match_grupo.group(1)

            match_classe = re.search(r'(Residencial|Comercial|Industrial|Rural|Poder Público|Iluminação Pública)', classif_line_content, re.IGNORECASE)
            if match_classe:
                dados_extraidos['Classe_Tarifaria'] = match_classe.group(1)
        print(f"DEBUG: Grupo_Tarifario: {dados_extraidos.get('Grupo_Tarifario')}, Classe_Tarifaria: {dados_extraidos.get('Classe_Tarifaria')}")

        # --- NOME/RAZÃO SOCIAL ---
        print("DEBUG: Extraindo Nome/Razão Social...")
        nome_match = re.search(
            r'(?:Monofásico|Bifásico|Trifásico)\s*\n+([A-Z\s,.-]+)\s*\n+(?:Leitura anterior|DATAS DE|COD UA)',
            texto,
            re.DOTALL | re.IGNORECASE
        )
        if nome_match:
            dados_extraidos['Nome_Razao_Social'] = nome_match.group(1).strip()
        print(f"DEBUG: Nome_Razao_Social: {dados_extraidos.get('Nome_Razao_Social')}")


        # --- ENDEREÇO (Rua e Número) ---
        print("DEBUG: Extraindo Endereço Rua e Número...")
        endereco_rua_match = re.search(
            r'Proxima Leitura\s*\n+([^\n]+) DATAS DE', 
            texto,
            re.DOTALL
        )
        if endereco_rua_match:
            dados_extraidos['Endereco_Rua_Numero'] = endereco_rua_match.group(1).strip()
        print(f"DEBUG: Endereco_Rua_Numero: {dados_extraidos.get('Endereco_Rua_Numero')}")


        # --- BAIRRO, CIDADE, ESTADO ---
        print("DEBUG: Extraindo Bairro, Cidade, Estado...")
        interior_line_match = re.search(
            r'COD UA \d+ LEITURAS.*?\n\s*(INTERIOR / ([A-Za-zÀ-ÖØ-öø-ÿ\s,.-]+))-([A-Z]{2})', # Suporte a acentos no nome da cidade
            texto,
            re.DOTALL
        )
        if interior_line_match:
            # bairro_cidade_raw = interior_line_match.group(1).strip() # Ex: "INTERIOR / Giruá"
            cidade_capturada = interior_line_match.group(2).strip() # Ex: "Giruá"
            estado = interior_line_match.group(3).strip() # Ex: "RS"
            
            dados_extraidos['Bairro'] = 'INTERIOR'
            dados_extraidos['Cidade'] = cidade_capturada
            dados_extraidos['Estado'] = estado
        print(f"DEBUG: Bairro: {dados_extraidos.get('Bairro')}, Cidade: {dados_extraidos.get('Cidade')}, Estado: {dados_extraidos.get('Estado')}")


        # --- CNPJ/CPF ---
        print("DEBUG: Extraindo CNPJ/CPF...")
        match_cpf_cnpj = re.search(r'CPF/CNPJ:\s*([\d*]{3}\.[\d*]{3}\.[\d*]{3}-\d{2}|\d{2}\.[\d*]{3}\.[\d*]{3}\/\d{4}-\d{2})', texto)
        if match_cpf_cnpj:
            dados_extraidos['CNPJ_CPF'] = match_cpf_cnpj.group(1)
        print(f"DEBUG: CNPJ_CPF: {dados_extraidos.get('CNPJ_CPF')}")


        # --- CEP ---
        print("DEBUG: Extraindo CEP...")
        match_cep = re.search(r'CEP:\s*(\d{2}\s*\d{3}-\d{3})', texto)
        if match_cep:
            dados_extraidos['CEP'] = match_cep.group(1)
        print(f"DEBUG: CEP: {dados_extraidos.get('CEP')}")


        # --- UC (Unidade Consumidora) ---
        print("DEBUG: Extraindo UC...")
        uc_match = re.search(r'CEP:\s*\d{2}\s*\d{3}-\d{3}\s*([\d-]+)', texto)
        if uc_match:
            dados_extraidos['UC'] = uc_match.group(1)
        print(f"DEBUG: UC: {dados_extraidos.get('UC')}")

    except Exception as e:
        print(f"Erro ao extrair dados do sub-layout 'Cooperluz (com COD UA)': {e}")
        pass
    print(f"DEBUG: Saindo de _extrair_dados_layout_cooperluz_sublayout_com_cod_ua com dados extraídos.")
    return dados_extraidos

def _extrair_dados_layout_cooperluz_sublayout_sem_cod_ua(texto, caminho_pdf):
    """
    Extrai dados de faturas da Cooperluz sem o padrão 'COD UA'. (Ex: Fatura de energia Roque Wesner.pdf, FATURA1.pdf, Fatura3.pdf)
    """
    print(f"DEBUG: Entrando em _extrair_dados_layout_cooperluz_sublayout_sem_cod_ua para {os.path.basename(caminho_pdf)}")
    dados_extraidos = {
        'Nome_Razao_Social': 'Não encontrado',
        'Endereco_Rua_Numero': 'Não encontrado',
        'Bairro': 'Não encontrado',
        'CNPJ_CPF': 'Não encontrado',
        'Cidade': 'Não encontrado',
        'CEP': 'Não encontrado',
        'UC': 'Não encontrado',
        'Grupo_Tarifario': 'Não encontrado',
        'Tensao_Nominal_V': 'Não encontrado',
        'Classe_Tarifaria': 'Não encontrado',
        'Estado': 'Não encontrado',
    }

    try:
        # --- TIPO DE FORNECIMENTO (para Tensão Nominal) ---
        print("DEBUG: Extraindo Tipo de Fornecimento...")
        tipo_fornecimento_match = re.search(
            r'Tipo de Fornecimento:\s*(?:[\s\S]*?)(Monofásico|Bifásico|Trifásico)',
            texto,
            re.IGNORECASE | re.DOTALL
        )
        tipo_fornecimento_extraido = ''
        if tipo_fornecimento_match:
            tipo_fornecimento_extraido = (tipo_fornecimento_match.group(1) or '').strip()
            if tipo_fornecimento_extraido:
                if 'Bifásico' in tipo_fornecimento_extraido or 'Monofásico' in tipo_fornecimento_extraido:
                    dados_extraidos['Tensao_Nominal_V'] = 220
                elif 'Trifásico' in tipo_fornecimento_extraido:
                    dados_extraidos['Tensao_Nominal_V'] = 380
        print(f"DEBUG: Tipo de Fornecimento: {dados_extraidos.get('Tensao_Nominal_V')}")

        # --- GRUPO E CLASSE TARIFÁRIA ---
        print("DEBUG: Extraindo Grupo e Classe Tarifária...")
        classificacao_line_match = re.search(
            r'Classificaç(?:ão|ao):\s*(.*?)(?:(?=\nTipo de Fornecimento)|\n|$)',
            texto,
            re.DOTALL | re.IGNORECASE
        )
        classif_line_content = ''
        if classificacao_line_match:
            classif_line_content = classificacao_line_match.group(1).strip()
            
            match_grupo = re.search(r'(B[1-4]|A)', classif_line_content)
            if match_grupo:
                dados_extraidos['Grupo_Tarifario'] = match_grupo.group(1)

            match_classe = re.search(r'(Residencial|Comercial|Industrial|Rural|Poder Público|Iluminação Pública)', classif_line_content, re.IGNORECASE)
            if match_classe:
                dados_extraidos['Classe_Tarifaria'] = match_classe.group(1)
        print(f"DEBUG: Grupo_Tarifario: {dados_extraidos.get('Grupo_Tarifario')}, Classe_Tarifaria: {dados_extraidos.get('Classe_Tarifaria')}")

        # --- NOME/RAZÃO SOCIAL ---
        print("DEBUG: Extraindo Nome/Razão Social...")
        nome_match = re.search(
            r'(?:Monofásico|Bifásico|Trifásico)\s*\n+([A-Z\s,.-]+)\s*\n+Leitura anterior',
            texto,
            re.DOTALL | re.IGNORECASE
        )
        if nome_match:
            dados_extraidos['Nome_Razao_Social'] = nome_match.group(1).strip()
        print(f"DEBUG: Nome_Razao_Social: {dados_extraidos.get('Nome_Razao_Social')}")


        # --- ENDEREÇO (Rua e Número) ---
        print("DEBUG: Extraindo Endereço Rua e Número...")
        endereco_rua_match = re.search(
            r'Proxima Leitura\s*\n+([^\n]+) DATAS DE',
            texto,
            re.DOTALL
        )
        if endereco_rua_match:
            dados_extraidos['Endereco_Rua_Numero'] = endereco_rua_match.group(1).strip()
        print(f"DEBUG: Endereco_Rua_Numero: {dados_extraidos.get('Endereco_Rua_Numero')}")


        # --- BAIRRO, CIDADE, ESTADO ---
        print("DEBUG: Extraindo Bairro, Cidade, Estado...")
        interior_line_match = re.search(
            r'LEITURAS.*?\n\s*(INTERIOR / ([A-Za-zÀ-ÖØ-öø-ÿ\s,.-]+))-([A-Z]{2})', # Suporte a acentos no nome da cidade
            texto,
            re.DOTALL
        )
        if interior_line_match:
            # bairro_cidade_raw = interior_line_match.group(1).strip() # Ex: "INTERIOR / Giruá"
            cidade_capturada = interior_line_match.group(2).strip() # Ex: "Giruá"
            estado = interior_line_match.group(3).strip() # Ex: "RS"
            
            dados_extraidos['Bairro'] = 'INTERIOR'
            dados_extraidos['Cidade'] = cidade_capturada
            dados_extraidos['Estado'] = estado
        print(f"DEBUG: Bairro: {dados_extraidos.get('Bairro')}, Cidade: {dados_extraidos.get('Cidade')}, Estado: {dados_extraidos.get('Estado')}")


        # --- CNPJ/CPF ---
        print("DEBUG: Extraindo CNPJ/CPF...")
        match_cpf_cnpj = re.search(r'CPF/CNPJ:\s*([\d*]{3}\.[\d*]{3}\.[\d*]{3}-\d{2}|\d{2}\.[\d*]{3}\.[\d*]{3}\/\d{4}-\d{2})', texto)
        if match_cpf_cnpj:
            dados_extraidos['CNPJ_CPF'] = match_cpf_cnpj.group(1)
        print(f"DEBUG: CNPJ_CPF: {dados_extraidos.get('CNPJ_CPF')}")


        # --- CEP ---
        print("DEBUG: Extraindo CEP...")
        match_cep = re.search(r'CEP:\s*(\d{2}\s*\d{3}-\d{3})', texto)
        if match_cep:
            dados_extraidos['CEP'] = match_cep.group(1)
        print(f"DEBUG: CEP: {dados_extraidos.get('CEP')}")


        # --- UC (Unidade Consumidora) ---
        print("DEBUG: Extraindo UC...")
        uc_match = re.search(
            r'UNIDADE CONSUMIDORA\s*\n+Rota:\s*\d+,\s*Sequência:\s*\d+\s*([\d-]+)',
            texto,
            re.DOTALL
        )
        if uc_match:
            dados_extraidos['UC'] = uc_match.group(1).strip()
        print(f"DEBUG: UC: {dados_extraidos.get('UC')}")

    except Exception as e:
        print(f"Erro ao extrair dados do sub-layout 'Cooperluz (sem COD UA)': {e}")
        pass
    print(f"DEBUG: Saindo de _extrair_dados_layout_cooperluz_sublayout_sem_cod_ua com dados extraídos.")
    return dados_extraidos


def _extrair_dados_layout_cooperluz_style(texto, caminho_pdf):
    """
    Extrai dados de faturas com o layout da Cooperluz, usando um dispatcher para sub-layouts.
    """
    print(f"DEBUG: Entrando em _extrair_dados_layout_cooperluz_style para {os.path.basename(caminho_pdf)}")
    if not re.search(r'COOPERLUZ.*?NOROESTE', texto, re.IGNORECASE | re.DOTALL):
        # AQUI FOI ALTERADO: Permite que CERTHIL e CERMISSOES usem o mesmo extractor base se o texto NÃO CONTIVER Cooperluz
        # Mas para garantir que a Cooperluz sempre tente seu próprio extractor, mantemos a verificação inicial.
        # Esta função `_extrair_dados_layout_cooperluz_style` é agora mais específica para Cooperluz.
        pass # A verificação para distribuidora já é feita mais acima.

    # Heurística para distinguir entre os sub-layouts da Cooperluz
    print(f"DEBUG: Verificando sub-layout de Cooperluz para {os.path.basename(caminho_pdf)}")
    if re.search(r'COD UA \d+', texto):
        print(f"Detectado sub-layout 'Cooperluz (com COD UA)' para {os.path.basename(caminho_pdf)}.")
        dados = _extrair_dados_layout_cooperluz_sublayout_com_cod_ua(texto, caminho_pdf) # Moved call
        print(f"DEBUG extrair_dados_fatura: Retorno de _extrair_dados_layout_cooperluz_sublayout_com_cod_ua. Keys: {dados.keys()}") # Adicionado
        return dados
    else:
        print(f"Detectado sub-layout 'Cooperluz (sem COD UA)' para {os.path.basename(caminho_pdf)}.")
        dados = _extrair_dados_layout_cooperluz_sublayout_sem_cod_ua(texto, caminho_pdf) # Moved call
        print(f"DEBUG extrair_dados_fatura: Retorno de _extrair_dados_layout_cooperluz_sublayout_sem_cod_ua. Keys: {dados.keys()}") # Adicionado
        return dados

# --- NOVAS Funções para Distribuidoras Similares à Cooperluz (Certhil e Cermissões) ---
def _extrair_dados_layout_coop_similar_style(texto, caminho_pdf, distributor_name):
    """
    Extrai dados de faturas com layout similar ao da Cooperluz (sem o padrão 'COD UA').
    Utilizado para Certhil e Cermissões, que seguem um padrão parecido.
    """
    print(f"DEBUG: Entrando em _extrair_dados_layout_coop_similar_style para {distributor_name} - {os.path.basename(caminho_pdf)}")
    dados_extraidos = {
        'Nome_Razao_Social': 'Não encontrado',
        'Endereco_Rua_Numero': 'Não encontrado',
        'Bairro': 'Não encontrado',
        'CNPJ_CPF': 'Não encontrado',
        'Cidade': 'Não encontrado',
        'CEP': 'Não encontrado',
        'UC': 'Não encontrado',
        'Grupo_Tarifario': 'Não encontrado',
        'Tensao_Nominal_V': 'Não encontrado',
        'Classe_Tarifaria': 'Não encontrado',
        'Estado': 'Não encontrado',
    }

    try:
        # --- TIPO DE FORNECIMENTO (para Tensão Nominal) ---
        tipo_fornecimento_match = re.search(
            r'Tipo de Fornecimento:\s*(?:[\s\S]*?)(Monofásico|Bifásico|Trifásico)',
            texto,
            re.IGNORECASE | re.DOTALL
        )
        if tipo_fornecimento_match:
            tipo_fornecimento_extraido = (tipo_fornecimento_match.group(1) or '').strip()
            if tipo_fornecimento_extraido:
                if 'Bifásico' in tipo_fornecimento_extraido or 'Monofásico' in tipo_fornecimento_extraido:
                    dados_extraidos['Tensao_Nominal_V'] = 220
                elif 'Trifásico' in tipo_fornecimento_extraido:
                    dados_extraidos['Tensao_Nominal_V'] = 380

        # --- GRUPO E CLASSE TARIFÁRIA ---
        classificacao_line_match = re.search(
            r'Classificaç(?:ão|ao):\s*(.*?)(?:(?=\nTipo de Fornecimento)|\n|$)',
            texto,
            re.DOTALL | re.IGNORECASE
        )
        classif_line_content = ''
        if classificacao_line_match:
            classif_line_content = classificacao_line_match.group(1).strip()
            
            match_grupo = re.search(r'(B[1-4]|A)', classif_line_content)
            if match_grupo:
                dados_extraidos['Grupo_Tarifario'] = match_grupo.group(1)

            match_classe = re.search(r'(Residencial|Comercial|Industrial|Rural|Poder Público|Iluminação Pública)', classif_line_content, re.IGNORECASE)
            if match_classe:
                dados_extraidos['Classe_Tarifaria'] = match_classe.group(1)

        # --- NOME/RAZÃO SOCIAL ---
        nome_match = re.search(
            r'(?:Monofásico|Bifásico|Trifásico)\s*\n+([A-Z\s,.-]+)\s*\n+(?:Leitura anterior|DATAS DE)',
            texto,
            re.DOTALL | re.IGNORECASE
        )
        if nome_match:
            dados_extraidos['Nome_Razao_Social'] = nome_match.group(1).strip()


        # --- ENDEREÇO (Rua e Número) ---
        endereco_rua_match = re.search(
            r'Proxima Leitura\s*\n+([^\n]+) DATAS DE',
            texto,
            re.DOTALL
        )
        if endereco_rua_match:
            dados_extraidos['Endereco_Rua_Numero'] = endereco_rua_match.group(1).strip()


        # --- BAIRRO, CIDADE, ESTADO ---
        # Alterado para ser mais flexível, procurando "RURAL / Cidade-RS" que aparece nas faturas da Certhil e Cermissões.
        # Original: r'LEITURAS.*?\n\s*(INTERIOR / ([A-Za-zÀ-ÖØ-öø-ÿ\s,.-]+))-([A-Z]{2})'
        interior_line_match = re.search(
            r'(?:LEITURAS|UNIDADE CONSUMIDORA).*?\n\s*(RURAL|INTERIOR)\s*/\s*([A-Za-zÀ-ÖØ-öø-ÿ\s,.-]+)-([A-Z]{2})',
            texto,
            re.DOTALL
        )
        if interior_line_match:
            bairro_capturado = interior_line_match.group(1).strip() # "RURAL" ou "INTERIOR"
            cidade_capturada = interior_line_match.group(2).strip()
            estado = interior_line_match.group(3).strip()
            
            dados_extraidos['Bairro'] = bairro_capturado
            dados_extraidos['Cidade'] = cidade_capturada
            dados_extraidos['Estado'] = estado


        # --- CNPJ/CPF ---
        match_cpf_cnpj = re.search(r'CPF/CNPJ:\s*([\d*]{3}\.[\d*]{3}\.[\d*]{3}-\d{2}|\d{2}\.[\d*]{3}\.[\d*]{3}\/\d{4}-\d{2})', texto)
        if match_cpf_cnpj:
            dados_extraidos['CNPJ_CPF'] = match_cpf_cnpj.group(1)

        # --- CEP ---
        match_cep = re.search(r'CEP:\s*(\d{2}\s*\d{3}-\d{3})', texto)
        if match_cep:
            dados_extraidos['CEP'] = match_cep.group(1)

        # --- UC (Unidade Consumidora) ---
        # Priorizar a busca por "UC: " que é bem explícita
        uc_match_explicit = re.search(r'UC:\s*([\d]+)[- ]', texto) # Captura dígitos até um traço ou espaço
        if uc_match_explicit:
            dados_extraidos['UC'] = uc_match_explicit.group(1).strip()
        else:
            # Fallback para o padrão próximo a "UNIDADE CONSUMIDORA" e "Rota"
            uc_match_rota = re.search(
                r'UNIDADE CONSUMIDORA\s*\n+Rota:\s*\d+,\s*Sequência:\s*\d+\s*([\d]+)',
                texto,
                re.DOTALL
            )
            if uc_match_rota:
                dados_extraidos['UC'] = uc_match_rota.group(1).strip()
            else:
                # Outro fallback, se houver um "CÓDIGO DO CLIENTE" seguido de número
                uc_match_codigo_cliente = re.search(r'CÓDIGO DO CLIENTE\s*\n*([\d]+)', texto)
                if uc_match_codigo_cliente:
                    dados_extraidos['UC'] = uc_match_codigo_cliente.group(1).strip()

    except Exception as e:
        print(f"Erro ao extrair dados do sub-layout '{distributor_name} (similar Cooperluz)': {e}")
        pass
    print(f"DEBUG: Saindo de _extrair_dados_layout_coop_similar_style com dados extraídos para {distributor_name}.")
    return dados_extraidos


def extrair_dados_fatura(caminho_pdf, distributor_type):
    """
    Extrai dados de uma fatura de energia em formato PDF,
    baseado no tipo de distribuidora fornecido pelo usuário.
    """
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            primeira_pagina = pdf.pages[0]
            texto = primeira_pagina.extract_text()
            print(f"--- Texto extraído de {os.path.basename(caminho_pdf)} ({distributor_type}) ---")
            # print(texto) # Comentado para não poluir muito o log
            print("--- Fim do texto extraído ---")

            if distributor_type == 'RGE':
                # --- Lógica de Detecção de Layout da RGE Existente ---
                is_layout_adriano_style = re.search(r'Inscrição no CNPJ: \d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}\n+([A-Z\s,.]+)\n.*?Pelo CPF:\s*\d{3}\.\d{3}\.\d{3}-\d{2}', texto, re.DOTALL)
                
                is_layout_adroaldo_aire_style = re.search(r'Inscrição no CNPJ: \d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}\n+([A-Z\s,.]+)\n.*?CPF:\s*(((\*)?){6}\.\d{3}-(((\*)?){2}))', texto, re.DOTALL) # Corrigido regex para CPF mascarado
                
                is_layout_arcindo_style = re.search(r'DANFE - DOCUMENTO AUXILIAR DA NOTA FISCAL ELETRÔNICA', texto) and re.search(r'CÓDIGO DA UNIDADE CONSUMIDORA:', texto)

                if is_layout_adriano_style:
                    print(f"Detectado layout 'Adriano Style' para {os.path.basename(caminho_pdf)} (RGE)")
                    return _extrair_dados_layout_adriano_style(texto, caminho_pdf)
                elif is_layout_adroaldo_aire_style:
                    print(f"Detectado layout 'Adroaldo/Aire Style' para {os.path.basename(caminho_pdf)} (RGE)")
                    return _extrair_dados_layout_adroaldo_style(texto, caminho_pdf)
                elif is_layout_arcindo_style:
                    print(f"Detectado layout 'Arcindo Style' para {os.path.basename(caminho_pdf)} (RGE)")
                    return _extrair_dados_layout_arcindo_style(texto, caminho_pdf)
                else:
                    print(f"Aviso: Não foi possível identificar o layout primário RGE para '{os.path.basename(caminho_pdf)}'. Tentando fallbacks RGE...")

                    # Fallbacks para RGE
                    extraction_functions = [
                        (_extrair_dados_layout_adriano_style, 'Adriano Style Fallback'),
                        (_extrair_dados_layout_adroaldo_style, 'Adroaldo/Aire Style Fallback'),
                        (_extrair_dados_layout_arcindo_style, 'Arcindo Style Fallback')
                    ]

                    best_match_data = None
                    max_found_fields = 0

                    for func, layout_name in extraction_functions:
                        current_data = func(texto, caminho_pdf)
                        current_found_fields = sum(1 for k, v in current_data.items() if v != 'Não encontrado' and k not in ['error'])

                        if current_found_fields > max_found_fields:
                            max_found_fields = current_found_fields
                            best_match_data = current_data

                    if best_match_data and max_found_fields > 0:
                         print(f"Fallback RGE bem-sucedido: {max_found_fields} campos encontrados para '{os.path.basename(caminho_pdf)}'.")
                         return best_match_data
                    
                    return {'error': f"Erro: Não foi possível identificar o layout da fatura RGE '{os.path.basename(caminho_pdf)}' e os fallbacks falharam. Layout desconhecido ou estrutura muito diferente."}

            elif distributor_type == 'COOPERLUZ':
                print(f"Tentando extrair dados como Cooperluz para {os.path.basename(caminho_pdf)}...")
                dados = _extrair_dados_layout_cooperluz_style(texto, caminho_pdf)
                
                # Para Cooperluz, se o Nome/Razao Social não for encontrado
                if dados.get('Nome_Razao_Social') == 'Não encontrado':
                    return {'error': f"Erro: A fatura da Cooperluz '{os.path.basename(caminho_pdf)}' não pôde ser extraída. O Nome/Razão Social não foi encontrado, indicando um problema com o layout ou a legibilidade do PDF."}
                
                return dados
            elif distributor_type == 'CERTHIL':
                print(f"Tentando extrair dados como Certhil para {os.path.basename(caminho_pdf)}...")
                dados = _extrair_dados_layout_coop_similar_style(texto, caminho_pdf, 'CERTHIL')
                if dados.get('Nome_Razao_Social') == 'Não encontrado':
                     return {'error': f"Erro: A fatura da Certhil '{os.path.basename(caminho_pdf)}' não pôde ser extraída. O Nome/Razão Social não foi encontrado, indicando um problema com o layout ou a legibilidade do PDF."}
                return dados
            elif distributor_type == 'CERMISSOES':
                print(f"Tentando extrair dados como Cermissões para {os.path.basename(caminho_pdf)}...")
                dados = _extrair_dados_layout_coop_similar_style(texto, caminho_pdf, 'CERMISSOES')
                if dados.get('Nome_Razao_Social') == 'Não encontrado':
                     return {'error': f"Erro: A fatura da Cermissões '{os.path.basename(caminho_pdf)}' não pôde ser extraída. O Nome/Razão Social não foi encontrado, indicando um problema com o layout ou a legibilidade do PDF."}
                return dados
            else:
                return {'error': f"Erro: Tipo de distribuidora '{distributor_type}' desconhecido. Por favor, selecione uma das opções válidas."}

    except pdfplumber.pdfminer.pdfdocument.PDFSyntaxError:
        return {'error': f"Erro: O arquivo '{os.path.basename(caminho_pdf)}' não é um PDF válido ou está corrompido."}
    except FileNotFoundError:
        return {'error': f"Erro: O arquivo '{os.path.basename(caminho_pdf)}' não foi encontrado. (Isso não deveria acontecer se o upload foi bem sucedido)."}
    except Exception as e:
        print(f"Erro inesperado durante a abertura/leitura do PDF ou com o PDF: {e}") # Adicionado print do erro
        return {'error': f"Erro inesperado durante a abertura/leitura do PDF ou com o PDF: {e}"}

# --- Função auxiliar para parsear endereço para o Excel ---
def parse_address_for_excel(full_address):
    street = full_address
    number = ''

    # Tenta encontrar um número no final do endereço (ex: "RUA A 123" ou "AV B, 45 C")
    # Captura a parte numérica (e opcionalmente uma letra) no final da string
    # Considera separadores comuns como espaço ou vírgula antes do número
    match_number = re.search(r'(,\s*|\s*)(\d+[A-Za-z]?)\s*$', full_address)

    if match_number:
        number = match_number.group(2).strip()
        street = full_address[:match_number.start(1)].strip() # Remove o número e o separador (vírgula ou espaço)
    elif "S/N" in full_address.upper(): # Caso específico de "S/N" para "sem número"
        number = "S/N"
        street = full_address.upper().replace("S/N", "").replace(",", "").strip()


    return street, number

# --- Função para converter graus decimais para GMS ---
def decimal_to_dms(decimal_degrees, is_latitude=True):
    """
    Converte graus decimais para o formato Graus, Minutos, Segundos (DMS).

    Args:
        decimal_degrees (float): O valor em graus decimais.
        is_latitude (bool): True se for Latitude (para determinar N/S), False se for Longitude (para E/W).

    Returns:
        str: A string formatada em DMS (ex: "23° 30' 45.12" S").
        Retorna 'N/A' se a entrada não for um número válido.
    """
    if not isinstance(decimal_degrees, (int, float)):
        try:
            decimal_degrees = float(str(decimal_degrees).replace(',', '.'))
        except (ValueError, TypeError):
            return "N/A"

    # Determinar o sinal e a direção
    direction = ""
    if is_latitude:
        if decimal_degrees >= 0:
            direction = "N"
        else:
            direction = "S"
    else: # Longitude
        if decimal_degrees >= 0:
            direction = "E"
        else:
            direction = "W"

    # Tornar o valor positivo para os cálculos
    abs_degrees = abs(decimal_degrees)

    degrees = int(abs_degrees)
    minutes = int((abs_degrees - degrees) * 60)
    seconds = ((abs_degrees - degrees) * 60 - minutes) * 60

    # Arredondar segundos para 2 casas decimais
    seconds = round(seconds, 2)

    # Corrigido: Usando aspas triplas para evitar problemas de escape
    return f"""{degrees}° {minutes}' {seconds:.2f}" {direction}"""


# --- Mapeamento de campos internos para células do Excel (Planilla Projetos FV) ---
EXCEL_PROJETO_FV_CELL_MAPPING = {
    'Nome_Razao_Social': 'B2',
    'Endereco_Rua': 'C2',
    'Numero_Endereco': 'D2',
    'Bairro': 'E2',
    'E-MAIL': 'F2',
    'TELEFONE': 'G2',
    'CNPJ_CPF': 'H2',
    'Cidade': 'I2',
    'CEP': 'J2',
    'LATITUDE': 'K2',
    'LONGITUDE': 'L2',

    'UC': 'A5',
    'Grupo_Tarifario': 'B5',
    'Classe_Tarifaria': 'C5',
    'Tensao_Nominal_V': 'D5',
    'CARGA_INSTALADA': 'E5',
    'CATEGORIA': 'F5',
    'TIPO_DE_ATENDIMENTO': 'G5',

    'TIPO_DE_CAIXA': 'A8',
    'ISOLACAO': 'I8',

    'POTENCIA_MODULO_MANUAL': 'A27',
    'FABRICANTE_MODULO_MANUAL': 'C27',
    'QUANTIDADE_PLACAS_MANUAL': 'D27',

    'POTENCIA_INVERSOR_MANUAL': 'A29',
    'FABRICANTE_INVERSOR_MANUAL': 'C29',
    'QUANTIDADE_INVERSOR_MANUAL': 'D29',
}

# --- Mapeamento das CÉLULAS CALCULADAS na Planilha Projetos FV para leitura ---
EXCEL_PROJETO_FV_CALCULATED_CELLS = {
    'NUMERO_FASES_CALCULADO': 'B8',
    'RAMAL_ENTRADA_CALCULADO': 'D8',
    'DISJUNTOR_CALCULADO': 'E11',
    'MODELO_MODULO_CALCULADO': 'B27',
    'AREA_ARRANJOS_CALCULADO': 'G27',
    'POTENCIA_PICO_MODULOS': 'F27',
    'MODELO_INVERSOR_CALCULADO': 'B29',
    'POTENCIA_TOTAL_SISTEMA_KWP': 'G29',
    'INMETRO': 'J29',
    'ISOLACAO_CA': 'H44',
    'CABO_CA': 'F44',
    'DISJUNTOR_CA': 'A44',
    'DISJ_CA_TENS': 'C44',
    'DISJ_CA_INTR': 'I44',
    'DISJ_CA_ATEN': 'J44',
}

# --- Mapeamento dos textos (placeholders) no Anexo F para as variáveis do Python ---
ANEXO_F_PLACEHOLDER_TO_PYTHON_VAR = {
    'NOME/RAZÃO SOCIAL': 'Nome_Razao_Social',
    'CNPJ/CPF': 'CNPJ_CPF',
    'UC': 'UC',
    'ENDEREÇO + N° + BAIRRO': 'FULL_ADDRESS_COMPOSED', # Isto é uma variável composta que precisa ser montada
    'CEP': 'CEP',
    'CIDADE': 'Cidade',
    'TELEFONE': 'TELEFONE',
    'E-MAIL': 'E-MAIL',
    'CATEGORIA': 'CATEGORIA',
    'TIPO DE ATENDIMENTO': 'TIPO_DE_ATENDIMENTO',
    'TIPO DE CAIXA': 'TIPO_DE_CAIXA',
    'CARGA INSTALADA': 'CARGA_INSTALADA',
    'DISJUNTOR SIMPL.': 'DISJUNTOR_CALCULADO',
    'QUANTIDADE PLACAS': 'QUANTIDADE_PLACAS_MANUAL',
    'FABRICANTE_MODULO': 'FABRICANTE_MODULO_MANUAL',
    'MODELO_MODULO': 'MODELO_MODULO_CALCULADO',
    'ÁREA': 'AREA_ARRANJOS_CALCULADO',
    'QUANTIDADE INVERSOR': 'QUANTIDADE_INVERSOR_MANUAL',
    'FABRICANTE_INVERSOR': 'FABRICANTE_INVERSOR_MANUAL',
    'MODELO_INVERSOR': 'MODELO_INVERSOR_CALCULADO',
    'POTÊNCIA (ANEXO I)': 'POTENCIA_TOTAL_SISTEMA_KWP',
    'POTÊNCIA NOMINAL': 'POTENCIA_INVERSOR_MANUAL',
    'DATA OPERAÇÃO': 'DATA_OPERACAO_PREVISTA',
    'N° DE FASES': 'NUMERO_FASES_CALCULADO',
    'RAMAL DE ENTRADA': 'RAMAL_ENTRADA_CALCULADO',
    'POTENCIA_PICO_MODULOS': 'POTENCIA_PICO_MODULOS', # Placeholder do Excel para a variável Python
    'LATITUDE': 'LATITUDE',
    'LONGITUDE': 'LONGITUDE',
    'LATITUDE_GMS': 'LATITUDE_GMS',
    'LONGITUDE_GMS': 'LONGITUDE_GMS',
}

# Mapeamento dos textos (placeholders) no Anexo I para as variáveis do Python ---
ANEXO_I_PLACEHOLDER_TO_PYTHON_VAR = {
    'UC': 'UC',
    'GRUPO_TARIFARIO': 'Grupo_Tarifario',
    'CLASSE_TARIFARIA': 'Classe_Tarifaria',
    'ENDERECO_RUA_NUMERO': 'Endereco_Rua_Numero',
    'ENDERECO_RUA_NUMERO_BAIRRO': 'Endereco_Rua_Numero_Bairro',
    'BAIRRO': 'Bairro',
    'CIDADE': 'Cidade',
    'ESTADO': 'Estado',
    'CIDADE_ESTADO': 'CIDADE_ESTADO',
    'CEP': 'CEP',
    'LATITUDE_GMS': 'LATITUDE_GMS',
    'LONGITUDE_GMS': 'LONGITUDE_GMS',
    'NOME_RAZAO_SOCIAL': 'Nome_Razao_Social',
    'CNPJ_CPF': 'CNPJ_CPF',
    'TELEFONE': 'TELEFONE',
    'EMAIL': 'E-MAIL',
    'CARGA_INSTALADA': 'CARGA_INSTALADA',
    'TENSAO_NOMINAL_V': 'Tensao_Nominal_V',
    'TIPO_DE_ATENDIMENTO': 'TIPO_DE_ATENDIMENTO',
    'DISJUNTOR_CALCULADO': 'DISJUNTOR_CALCULADO',
    'RAMAL_ENTRADA_CALCULADO': 'RAMAL_ENTRADA_CALCULADO',
    'QUANTIDADE_PLACAS_MANUAL': 'QUANTIDADE_PLACAS_MANUAL',
    'POTENCIA_MODULO_MANUAL': 'POTENCIA_MODULO_MANUAL',
    'POTENCIA_PICO_MODULOS': 'POTENCIA_PICO_MODULOS',
    'FABRICANTE_MODULO_MANUAL': 'FABRICANTE_MODULO_MANUAL',
    'MODELO_MODULO_CALCULADO': 'MODELO_MODULO_CALCULADO',
    'QUANTIDADE_INVERSOR_MANUAL': 'QUANTIDADE_INVERSOR_MANUAL',
    'POTENCIA_INVERSOR_MANUAL': 'POTENCIA_INVERSOR_MANUAL',
    'POTENCIA_TOTAL_SISTEMA_KWP': 'POTENCIA_TOTAL_SISTEMA_KWP',
    'FABRICANTE_INVERSOR_MANUAL': 'FABRICANTE_INVERSOR_MANUAL',
    'MODELO_INVERSOR_CALCULADO': 'MODELO_INVERSOR_CALCULADO',
    'AREA_ARRANJOS_CALCULADO': 'AREA_ARRANJOS_CALCULADO',
    'ART': 'ART',
    'DATA_ATUAL': 'DATA_ATUAL',
    'DATA_OPERACAO_PREVISTA': 'DATA_OPERACAO_PREVISTA',
    'POTENCIA_MODULO_MANUAL_KWP': 'POTENCIA_MODULO_MANUAL_KWP',
    'NUMERO_FASES_CALCULADO': 'NUMERO_FASES_CALCULADO',
}


# --- Configuração do Flask ---
app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
# LINHA MODIFICADA: Aumentado para 20MB
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024
# LINHA MODIFICADA: Substituído 'supersecretkey' por os.urandom(24) para segurança
app.secret_key = os.urandom(24)

ALLOWED_EXTENSIONS = {'pdf'}
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg'} # Adicionado para validação de imagens

EXCEL_PROJETO_FV_TEMPLATE_FILENAME = 'Planilha Projetos FV.xlsx'
EXCEL_PROJETO_FV_TEMPLATE_PATH = os.path.join(app.root_path, 'templates', EXCEL_PROJETO_FV_TEMPLATE_FILENAME)

ANEXO_F_TEMPLATE_FILENAME = 'Formulário Anexo F - GED 15303.xlsx'
ANEXO_F_TEMPLATE_PATH = os.path.join(app.root_path, 'templates', ANEXO_F_TEMPLATE_FILENAME)
ANEXO_F_SHEET_NAME = 'Anexo F - GED 15303'

ANEXO_I_TEMPLATE_FILENAME = 'Anexo I.xlsx'
ANEXO_I_TEMPLATE_PATH = os.path.join(app.root_path, 'templates', ANEXO_I_TEMPLATE_FILENAME)
ANEXO_I_SHEET_NAME = 'Plan1'

ANEXO_E_TEMPLATE_FILENAME = 'Formulário Anexo E - GED.docx'
ANEXO_E_TEMPLATE_PATH = os.path.join(app.root_path, 'templates', ANEXO_E_TEMPLATE_FILENAME)

TERMO_ACEITE_TEMPLATE_FILENAME = 'Termo de Aceite.docx'
TERMO_ACEITE_TEMPLATE_PATH = os.path.join(app.root_path, 'templates', TERMO_ACEITE_TEMPLATE_FILENAME)

# Constantes para a Procuração
PROCURACAO_TEMPLATE_FILENAME = 'Procuracao.docx'
PROCURACAO_TEMPLATE_PATH = os.path.join(app.root_path, 'templates', PROCURACAO_TEMPLATE_FILENAME)

# Constantes para o Termo de Aceite Inciso III
TERMO_ACEITE_INCISO_III_TEMPLATE_FILENAME = 'Termo de Aceite Inciso III.docx'
TERMO_ACEITE_INCISO_III_TEMPLATE_PATH = os.path.join(app.root_path, 'templates', TERMO_ACEITE_INCISO_III_TEMPLATE_FILENAME)

# Constantes para a Responsabilidade Tecnica
RESPONSABILIDADE_TECNICA_TEMPLATE_FILENAME = 'Responsabilidade Tecnica.docx'
RESPONSABILIDADE_TECNICA_TEMPLATE_PATH = os.path.join(app.root_path, 'templates', RESPONSABILIDADE_TECNICA_TEMPLATE_FILENAME)

# Constantes para Dados para GD de UFV
DADOS_GD_UFV_TEMPLATE_FILENAME = 'Dados para GD de UFV.docx'
DADOS_GD_UFV_TEMPLATE_PATH = os.path.join(app.root_path, 'templates', DADOS_GD_UFV_TEMPLATE_FILENAME)

# NOVO: Constante para Memorial Descritivo
MEMORIAL_DESCRITIVO_TEMPLATE_FILENAME = 'Memorial Descritivo Padrão Fecoergs.docx'
MEMORIAL_DESCRITIVO_TEMPLATE_PATH = os.path.join(app.root_path, 'templates', MEMORIAL_DESCRITIVO_TEMPLATE_FILENAME)


# --- NOVAS CONSTANTES PARA ARQUIVOS ESTÁTICOS E PLACEHOLDERS ---
CERTIDAO_REGISTRO_PROFISSIONAL_FILENAME = 'Certidao de Registro Profissional.pdf'
CERTIDAO_REGISTRO_PROFISSIONAL_PATH = os.path.join(app.root_path, 'templates', CERTIDAO_REGISTRO_PROFISSIONAL_FILENAME)
# Verifica se o arquivo existe, se não, cria um placeholder vazio temporariamente
if not os.path.exists(CERTIDAO_REGISTRO_PROFISSIONAL_PATH):
    # Conteúdo mínimo de um PDF vazio para evitar erro de File Not Found
    with open(CERTIDAO_REGISTRO_PROFISSIONAL_PATH, 'wb') as f:
        f.write(b'%PDF-1.4\n%\xc2\xa5\xc2\xb1\xc2\xae\xc2\xbb\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Count 0>>endobj\nxref\n0 3\n0000000000 65535 f\n0000000009 00000 n\n0000000074 00000 n\ntrailer<</Size 3/Root 1 0 R>>startxref\n120\n%%EOF')
    print(f"ATENÇÃO: Arquivo de template '{CERTIDAO_REGISTRO_PROFISSIONAL_FILENAME}' não encontrado na pasta 'templates'. Um placeholder PDF vazio foi criado.")


def allowed_file(filename, allowed_extensions):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in allowed_extensions

# Função para forçar o recálculo de fórmulas do Excel usando win32com
def calculate_excel_formulas(filepath):
    if not win32:
        print("Aviso: 'pywin32' não está disponível. Não foi possível forçar o recálculo das fórmulas do Excel.")
        return False

    excel = None
    try:
        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(filepath)
        workbook.RefreshAll()
        excel.CalculateFullRebuild()
        workbook.Save()
        workbook.Close()
        print(f"Fórmulas recalculadas e salvas para: {filepath}")
        return True
    except Exception as e:
        print(f"Erro ao calcular fórmulas no Excel via COM: {e}")
        return False
    finally:
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()

# --- FUNÇÃO ATUALIZADA PARA SUBSTITUIR PLACEHOLDERS NO DOCX COM PRESERVAÇÃO DE FORMATO ---
def replace_docx_placeholders(doc_path, replacements):
    document = Document(doc_path)

    # List of keys that should always be bolded when replaced
    bold_keys = ['NOME_RAZAO_SOCIAL', 'CPF_CNPJ', 'UC']

    # Helper to process a single paragraph or cell text
    def process_text_block(text_block_container):
        for p in text_block_container.paragraphs:
            if not p.text.strip():
                continue

            original_full_text = p.text
            
            # Keep track of the original formatting from the first run in the paragraph
            # This will be used as a default for all new runs in this paragraph.
            # This simplifies logic, but means fine-grained original formatting (e.g., italic for one word, bold for another)
            # will be lost within the paragraph if it contains placeholders.
            default_run_format = {
                'bold': None, 'italic': None, 'font_name': None, 'font_size': None, 'underline': None
            }
            if p.runs:
                first_run = p.runs[0]
                default_run_format['bold'] = first_run.bold
                default_run_format['italic'] = first_run.italic
                default_run_format['font_name'] = first_run.font.name
                default_run_format['font_size'] = first_run.font.size
                default_run_format['underline'] = first_run.underline

            # Perform replacements and collect information about where the replaced values are
            final_segments_for_runs = []
            
            # Use regex to find all placeholders in the original_full_text
            # Create a combined regex pattern for all placeholders
            placeholder_patterns = [re.escape('{{' + key + '}}') for key in replacements.keys()]
            combined_pattern = '|'.join(placeholder_patterns)
            
            last_idx = 0
            # If there are placeholders, iterate and split the text
            if combined_pattern:
                for match in re.finditer(combined_pattern, original_full_text):
                    # Add text before the placeholder (if any)
                    if match.start() > last_idx:
                        final_segments_for_runs.append((original_full_text[last_idx:match.start()], False, None))
                    
                    # Get the placeholder key from the matched string, e.g., '{{NOME_RAZAO_SOCIAL}}' -> 'NOME_RAZAO_SOCIAL'
                    matched_placeholder = match.group(0)
                    placeholder_key = matched_placeholder.strip('{}') # Remove {{ and }}
                    
                    # Add the replaced value
                    value_to_insert = replacements.get(placeholder_key, matched_placeholder) # Use replacement, fallback to original placeholder if not found
                    final_segments_for_runs.append((str(value_to_insert), True, placeholder_key))
                    
                    last_idx = match.end()
            
            # Add any remaining text after the last placeholder or if no placeholders were found at all
            if last_idx < len(original_full_text):
                final_segments_for_runs.append((original_full_text[last_idx:], False, None))

            # If no actual replacements were made AND no placeholders were found in the combined pattern,
            # we can skip clearing and re-adding runs for efficiency.
            if not (final_segments_for_runs and any(s[1] for s in final_segments_for_runs)) and not combined_pattern:
                return 

            # Clear existing runs
            for i in range(len(p.runs) - 1, -1, -1):
                p.runs[i]._element.getparent().remove(p.runs[i]._element)

            # Add new runs based on segments and apply formatting
            for text, is_replaced, placeholder_key in final_segments_for_runs:
                if text: # Only add non-empty segments as runs
                    new_run = p.add_run(text)
                    
                    # Apply default formatting
                    new_run.bold = default_run_format['bold']
                    new_run.italic = default_run_format['italic']
                    new_run.font.name = default_run_format['font_name']
                    new_run.font.size = default_run_format['font_size']
                    new_run.underline = default_run_format['underline']

                    # Apply conditional bolding if it's a replaced segment and the key is in bold_keys
                    if is_replaced and placeholder_key in bold_keys:
                        new_run.bold = True
    
    # Process paragraphs directly
    process_text_block(document)

    # Process tables (cells also contain paragraphs)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                process_text_block(cell) # Use the same processing logic for cells

    return document

# --- FUNÇÃO PARA GERAR CONTEÚDO DO TXT DA ART ---
def generate_art_txt_content(data):
    # Formata CPF/CNPJ removendo caracteres não numéricos para a ART
    cnpj_cpf_numeric = re.sub(r'[^0-9]', '', str(data.get('CNPJ_CPF', '')))
    cep_numeric = re.sub(r'[^0-9]', '', str(data.get('CEP', '')))
    telefone_numeric = re.sub(r'[^0-9]', '', str(data.get('TELEFONE', '')))

    # Tenta converter POTENCIA_TOTAL_SISTEMA_KWP para float para exibir sem o " kWp" ou " kW"
    # E aplica format_value_for_display para o separador decimal
    potencia_total_kwp_clean = format_value_for_display(data.get('POTENCIA_TOTAL_SISTEMA_KWP', '0'), is_numeric=True)

    return f"""Criar nova ART:
- No site do CREA acessar: ART WEB vá até: NOVA ART

Em nova ART:

-Professional:
Empresa executante da obra/serviço. Selecione: QUASAT SOLAR- EQUIPAMENTOS DE ENERGIA

-ART:
Tipo de ART: Obra ou Serviço
Motivo da ART: Normal

- Contratante:
Busque o contratante (flecha amarela)
Clique em CADASTRAR
Contratante: {data.get('Nome_Razao_Social', 'Não informado')}
CPF/CNPJ: {cnpj_cpf_numeric}
E-mail: {data.get('E-MAIL', 'Não informado')}
CEP: {cep_numeric}
Telefone: {telefone_numeric}
CONFIRME

- Obra/Serviço:
Proprietário: clique em "Buscar dados do Contratante"
Finalidade: Outras Finalidades
Valor Contrato(R$): 2000
Honorários(R$): 200
Data Início: {format_value_for_display(data.get('DATA_ART'), is_date=True, date_separator='-')}
Data Previsão de Fim: {format_value_for_display(data.get('DATA_OPERACAO_PREVISTA'), is_date=True, date_separator='-')}

- Atividades
Ativ. Téc.(flecha amarela): Projeto e Execução
Qtd.: {potencia_total_kwp_clean}
Unidade(flecha amarela): Quilowatt

CONFIRMAR - FINALIZAR

LATITUDE: {format_value_for_display(data.get('LATITUDE'), is_numeric=True, numeric_decimal_separator='.')}
LONGITUDE: {format_value_for_display(data.get('LONGITUDE'), is_numeric=True, numeric_decimal_separator='.')}
LATITUDE_GMS: {data.get('LATITUDE_GMS', 'Não informado')}
LONGITUDE_GMS: {data.get('LONGITUDE_GMS', 'Não informado')}
"""

# --- FUNÇÃO PARA GERAR CONTEÚDO DO TXT DA POSTAGEM (GENÉRICA) ---
def generate_postagem_txt_content(data, distributor_name):
    # Formata potências e outros números usando a função format_value_for_display
    potencia_total_kwp_clean = format_value_for_display(data.get('POTENCIA_TOTAL_SISTEMA_KWP', '0'), is_numeric=True)
    potencia_pico_modulos_clean = format_value_for_display(data.get('POTENCIA_PICO_MODULOS', '0'), is_numeric=True)
    potencia_inversor_manual_clean = format_value_for_display(data.get('POTENCIA_INVERSOR_MANUAL', '0'), is_numeric=True)
    
    return f"""Instalação: {data.get('UC', 'Não informado')}
Título do projeto: FV {data.get('Nome_Razao_Social', 'Não informado')}
Potência instalada do Gerador: {potencia_total_kwp_clean}
Carga atual do cliente: {format_value_for_display(data.get('CARGA_INSTALADA'), is_numeric=True)}
Latitude: {format_value_for_display(data.get('LATITUDE'), is_numeric=True, numeric_decimal_separator='.')}
Longitude: {format_value_for_display(data.get('LONGITUDE'), is_numeric=True, numeric_decimal_separator='.')}
Data prevista para ligação: {format_value_for_display(data.get('DATA_OPERACAO_PREVISTA'), is_date=True)}

Tipo de solicitação: SOLICITAÇÃO DE ACESSO
Quantidade de Beneficiárias: 0
Telefone do Titular (DDD + Número): {data.get('TELEFONE', 'Não informado')}
E-mail do Titular: {data.get('E-MAIL', 'Não informado')}
Latitude: {format_value_for_display(data.get('LATITUDE'), is_numeric=True, numeric_decimal_separator='.')}
Longitude: {format_value_for_display(data.get('LONGITUDE'), is_numeric=True, numeric_decimal_separator='.')}
Potência Total dos Módulos (kw): {potencia_pico_modulos_clean}
Quantidade de Módulos: {format_value_for_display(data.get('QUANTIDADE_PLACAS_MANUAL'), is_numeric=True)}
Potência Total dos Inversores (kw): {potencia_inversor_manual_clean}
Quantidade de Inversores: {format_value_for_display(data.get('QUANTIDADE_INVERSOR_MANUAL'), is_numeric=True)}
Fabricante(s) dos Módulos: {format_value_for_display(data.get('FABRICANTE_MODULO_MANUAL'))}
Modelo(s) dos Módulos: {format_value_for_display(data.get('MODELO_MODULO_CALCULADO'))}
Fabricante(s) dos Inversores: {format_value_for_display(data.get('FABRICANTE_INVERSOR_MANUAL'))}
Modelo(s) dos Inversores: {format_value_for_display(data.get('MODELO_INVERSOR_CALCULADO'))}
Área Total dos Arranjos (m²): {format_value_for_display(data.get('AREA_ARRANJOS_CALCULADO'), is_numeric=True)}

E-mail: {data.get('E-MAIL', 'Não informado')}
Telefone: {data.get('TELEFONE', 'Não informado')}
Celular: {data.get('TELEFONE', 'Não informado')}
Documento de Responsabilidade Técnica:
Data do Documento: {format_value_for_display(data.get('DATA_ART'), is_date=True, date_separator='-')}
Email: projetos@quasatservices.com

LATITUDE_GMS: {data.get('LATITUDE_GMS', 'Não informado')}
LONGITUDE_GMS: {data.get('LONGITUDE_GMS', 'Não informado')}
"""

def generate_images_pdf(image_data_list, output_pdf_path):
    """
    Gera um PDF a partir de uma lista de imagens, com título para cada imagem.
    image_data_list: [{'path': 'caminho/para/imagem.jpg', 'title': 'Título da Imagem'}]
    """
    c = canvas.Canvas(output_pdf_path, pagesize=portrait(A4))
    width, height = portrait(A4)
    margin = 2 * cm

    for item in image_data_list:
        img_path = item['path']
        title = item['title']

        if not os.path.exists(img_path):
            print(f"Aviso: Imagem não encontrada para PDF: {img_path}")
            # Adiciona uma página em branco com um aviso se a imagem não for encontrada
            c.setFont('Helvetica-Bold', 14)
            c.drawCentredString(width / 2.0, height - 2 * cm, title)
            c.setFont('Helvetica', 12)
            c.drawCentredString(width / 2.0, height / 2.0, f"Imagem não encontrada: {title}")
            c.showPage()
            continue

        try:
            # Adiciona o título no topo da página
            c.setFont('Helvetica-Bold', 14)
            c.drawCentredString(width / 2.0, height - cm, title)

            # Abre a imagem com Pillow para obter dimensões e manipulá-la
            pil_img = Image.open(img_path)

            # Redimensiona a imagem para caber na página com margens, mantendo a proporção
            max_img_width = width - 2 * margin
            max_img_height = height - 3.5 * cm # Espaço para título e margens

            img_width_orig, img_height_orig = pil_img.size
            aspect_ratio = img_width_orig / img_height_orig

            if img_width_orig > max_img_width or img_height_orig > max_img_height:
                # Reduzir imagem
                scale_width = max_img_width / img_width_orig
                scale_height = max_img_height / img_height_orig
                scale_factor = min(scale_width, scale_height)

                img_width = img_width_orig * scale_factor
                img_height = img_height_orig * scale_factor
            else:
                # Não precisa reduzir, usa o tamanho original
                img_width = img_width_orig
                img_height = img_height_orig

            # Centraliza a imagem na página abaixo do título
            x_pos = (width - img_width) / 2
            y_pos = (height - img_height) / 2 - (1 * cm) # Ajusta ligeiramente para baixo do centro, considerando título

            c.drawImage(ImageReader(img_path), x_pos, y_pos, width=img_width, height=img_height, preserveAspectRatio=True)
            c.showPage() # Inicia uma nova página para a próxima imagem

        except Exception as e:
            print(f"Erro ao adicionar imagem {img_path} ao PDF: {e}")
            c.setFont('Helvetica-Bold', 14)
            c.drawCentredString(width / 2.0, height - 2 * cm, title)
            c.setFont('Helvetica', 12)
            c.drawCentredString(width / 2.0, height / 2.0, f"Erro ao processar imagem: {title}")
            c.showPage()

    c.save()


@app.route('/', methods=['GET'])
def upload_form():
    session.clear() # Always clear session when starting a new upload
    return render_template('index.html')

# Nova rota para processar_data que pode receber dados pré-preenchidos ou vazio
@app.route('/process_data', methods=['GET'])
def show_process_data_form():
    # Prioriza os dados completos de uma correção, senão os dados extraídos do PDF
    dados = session.pop('all_input_data_for_correction', None) or session.pop('extracted_data_from_pdf', {})
    
    # Se 'dados' ainda for vazio e não houver dados anteriores para correção,
    # significa que o usuário talvez acessou a rota diretamente ou a sessão expirou.
    if not dados and not session.get('extracted_data_from_pdf'):
        return redirect(url_for('upload_form')) # Redireciona para o início se não houver dados.

    # Salva os dados para que possam ser pré-preenchidos no formulário
    session['current_process_data_form_data'] = dados 
    return render_template('process_data.html', dados=dados)

# Nova rota para limpar a sessão e redirecionar para o upload inicial
@app.route('/clear_session_and_redirect_to_upload')
def clear_session_and_redirect_to_upload():
    session.clear()
    return redirect(url_for('upload_form'))


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        print("DEBUG UPLOAD: Nenhum arquivo enviado.")
        return render_template('index.html', error='Nenhum arquivo enviado.'), 400

    file = request.files['file']
    distributor_type = request.form.get('distribuidora')

    if not distributor_type:
        print("DEBUG UPLOAD: Distribuidora não selecionada.")
        return render_template('index.html', error='Por favor, selecione a distribuidora.'), 400

    if file.filename == '':
        print("DEBUG UPLOAD: Nome de arquivo vazio.")
        return render_template('index.html', error='Nenhum arquivo selecionado.'), 400

    if file and allowed_file(file.filename, ALLOWED_EXTENSIONS):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        print("DEBUG UPLOAD: Arquivo salvo temporariamente. Chamando extrair_dados_fatura...")
        dados_fatura = extrair_dados_fatura(filepath, distributor_type)
        print(f"DEBUG UPLOAD: extrair_dados_fatura retornou. Dados: {dados_fatura}")

        if 'error' in dados_fatura or dados_fatura.get('Nome_Razao_Social') == 'Não encontrado':
            os.remove(filepath)
            error_message = dados_fatura.get('error', f'Erro desconhecido na extração da fatura para {distributor_type}. O Nome/Razão Social não foi encontrado, indicando um problema com o layout ou a legibilidade do PDF.')
            print(f"DEBUG UPLOAD: Erro de extração detectado: {error_message}")
            return render_template('index.html', error=error_message), 500

        os.remove(filepath)
        print("DEBUG UPLOAD: Arquivo de upload temporário removido.")

        dados_fatura['distributor_type'] = distributor_type
        session['extracted_data_from_pdf'] = dados_fatura # Armazena os dados extraídos do PDF
        
        print("DEBUG UPLOAD: Tentando redirecionar para process_data.html...")
        return redirect(url_for('show_process_data_form')) # Redireciona para a rota GET
    else:
        print("DEBUG UPLOAD: Tipo de arquivo não permitido.")
        return render_template('index.html', error='Tipo de arquivo não permitido. Por favor, envie um PDF.'), 400

# Conjunto de chaves de `all_input_data` que contêm valores numéricos e precisam de formatação com vírgula.
NUMERIC_KEYS_FOR_FORMATTING = {
    'LATITUDE', 'LONGITUDE', 'POTENCIA_MODULO_MANUAL', 'POTENCIA_INVERSOR_MANUAL',
    'CARGA_INSTALADA', 'POTENCIA_PICO_MODULOS', 'POTENCIA_TOTAL_SISTEMA_KWP',
    'POTENCIA_MODULO_MANUAL_KWP', 'AREA_ARRANJOS_CALCULADO', 'Tensao_Nominal_V',
    'NUMERO_FASES_CALCULADO', 'DISJUNTOR_CALCULADO', 'DISJUNTOR_CA',
    'DISJ_CA_TENS', 'DISJ_CA_INTR', 'DISJ_CA_ATEN', 'QUANTIDADE_PLACAS_MANUAL',
    'QUANTIDADE_INVERSOR_MANUAL'
}

# Conjunto de chaves de `all_input_data` que contêm valores de data e precisam de formatação.
DATE_KEYS_FOR_FORMATTING = {
    'DATA_ATUAL', 'DATA_OPERACAO_PREVISTA', 'DATA_ART'
}

# Função auxiliar para obter valor formatado de forma consistente
def get_formatted_value_for_doc(key, data_dict):
    value = data_dict.get(key)
    if key in NUMERIC_KEYS_FOR_FORMATTING:
        return format_value_for_display(value, is_numeric=True)
    elif key in DATE_KEYS_FOR_FORMATTING:
        return format_value_for_display(value, is_date=True)
    return str(value) if value is not None else ''


@app.route('/process_and_save', methods=['POST'])
def process_and_save():
    # Retrieve data from the session for fields that might not be explicitly in the form
    # and to preserve the distributor_type
    all_input_data = session.pop('current_process_data_form_data', {})
    if not all_input_data:
        return render_template('index.html', error='Dados da sessão expirados ou não encontrados. Por favor, reinicie o processo.'), 400

    # The distributor_type is now explicitly submitted via a hidden input, or could still be in session.
    # Prioritize form submission for distributor_type if available, otherwise use session.
    distributor_type = request.form.get('distributor_type', all_input_data.get('distributor_type', 'RGE'))
    all_input_data['distributor_type'] = distributor_type # Ensure it's in all_input_data

    # Temporary storage for image files
    image_files = {}

    # Update all_input_data with values from the form, differentiating between
    # 'extracted_' fields and '_manual' fields (and the hidden distributor_type)
    for key, value in request.form.items():
        processed_value = value.strip() if isinstance(value, str) else value

        if key.startswith('extracted_'):
            original_key = key[len('extracted_'):]
            all_input_data[original_key] = processed_value
            print(f"DEBUG SAVE: Campo 'extracted_' {original_key} atualizado para: {all_input_data[original_key]}") # DEBUG
            # Handle specific type conversions for extracted fields
            if original_key == 'Tensao_Nominal_V' and processed_value.isdigit():
                all_input_data[original_key] = int(processed_value)
        elif key.endswith('_manual') or key == 'CARGA_INSTALADA' or key == 'TIPO_DE_ATENDIMENTO' or key == 'TIPO_DE_CAIXA': # Adicionado CARGA_INSTALADA etc. para serem tratados como "_manual"
            original_key = key.replace('_manual', '') # Remove _manual apenas se existir
            all_input_data[original_key] = processed_value
            print(f"DEBUG SAVE: Campo 'manual' {original_key} atualizado para: {all_input_data[original_key]}") # DEBUG
            # Handle specific type conversions for manual fields (e.g., numeric ones)
            if original_key in ['CARGA_INSTALADA', 'POTENCIA_MODULO_MANUAL', 'POTENCIA_INVERSOR_MANUAL', 'LATITUDE', 'LONGITUDE'] and processed_value:
                try:
                    # Replace comma with dot for float conversion
                    clean_value = processed_value.replace(',', '.')
                    all_input_data[original_key] = float(clean_value)
                except ValueError:
                    all_input_data[original_key] = processed_value # Keep original string if conversion fails
            elif original_key in ['QUANTIDADE_PLACAS_MANUAL', 'QUANTIDADE_INVERSOR_MANUAL'] and processed_value.isdigit():
                try:
                    all_input_data[original_key] = int(processed_value)
                except ValueError:
                    all_input_data[original_key] = processed_value # Keep original string if conversion fails
        elif key in ['ART', 'DATA_ART', 'CATEGORIA']: # These are top-level keys from manual input directly, INMETRO is now calculated
            all_input_data[key] = processed_value
            print(f"DEBUG SAVE: Campo 'top-level' {key} atualizado para: {all_input_data[key]}") # DEBUG
        # Adicionei um else final para pegar os outros campos que não seguem o padrão extracted_ ou _manual
        else:
            all_input_data[key] = processed_value
            print(f"DEBUG SAVE: Campo 'geral' {key} atualizado para: {all_input_data[key]}") # DEBUG
    
    # Process image files (these are in request.files, not request.form)
    image_files = {
        'foto_disjuntor': request.files.get('foto_disjuntor'),
        'foto_fachada': request.files.get('foto_fachada'),
        'foto_entrada_energia': request.files.get('foto_entrada_energia')
    }

    # --- Adiciona variável combinada de Bairro e Cidade ---
    bairro = all_input_data.get('Bairro', '').strip()
    cidade = all_input_data.get('Cidade', '').strip()

    if bairro and cidade:
        all_input_data['BAIRRO_CIDADE_COMBINADO'] = f"{bairro} - {cidade}"
    elif bairro:
        all_input_data['BAIRRO_CIDADE_COMBINADO'] = bairro
    elif cidade:
        all_input_data['BAIRRO_CIDADE_COMBINADO'] = cidade
    else:
        all_input_data['BAIRRO_CIDADE_COMBINADO'] = '' # Vazio se nenhum for encontrado

    # --- Adiciona variável combinada de Cidade e Estado ---
    cidade_val = all_input_data.get('Cidade', '').strip() # Renomeado para evitar conflito
    estado = all_input_data.get('Estado', '').strip()

    if cidade_val and estado:
        all_input_data['CIDADE_ESTADO'] = f"{cidade_val} - {estado}"
    elif cidade_val:
        all_input_data['CIDADE_ESTADO'] = cidade_val
    elif estado:
        all_input_data['CIDADE_ESTADO'] = estado
    else:
        all_input_data['CIDADE_ESTADO'] = '' # Vazio se nenhum for encontrado

    # --- Adiciona variável combinada de Endereço, Número e Bairro ---
    endereco_rua_numero = all_input_data.get('Endereco_Rua_Numero', '').strip()
    bairro_val = all_input_data.get('Bairro', '').strip() # Renomeado para evitar conflito com 'bairro' acima

    if endereco_rua_numero and bairro_val:
        all_input_data['Endereco_Rua_Numero_Bairro'] = f"{endereco_rua_numero} - {bairro_val}"
    elif endereco_rua_numero:
        all_input_data['Endereco_Rua_Numero_Bairro'] = endereco_rua_numero
    elif bairro_val:
        all_input_data['Endereco_Rua_Numero_Bairro'] = bairro_val
    else:
        all_input_data['Endereco_Rua_Numero_Bairro'] = '' # Vazio se nenhum for encontrado

    # Processar Latitude e Longitude para GMS
    try:
        lat_dec = float(str(all_input_data.get('LATITUDE', '0')).replace(',', '.'))
        lon_dec = float(str(all_input_data.get('LONGITUDE', '0')).replace(',', '.'))
        all_input_data['LATITUDE_GMS'] = decimal_to_dms(lat_dec, is_latitude=True)
        all_input_data['LONGITUDE_GMS'] = decimal_to_dms(lon_dec, is_latitude=False)
    except (ValueError, TypeError):
        all_input_data['LATITUDE_GMS'] = 'N/A'
        all_input_data['LONGITUDE_GMS'] = 'N/A'


    data_hoje = datetime.now()
    all_input_data['DATA_ATUAL'] = data_hoje.strftime('%d/%m/%Y')
    
    # Se DATA_ART não for fornecida manualmente, usa DATA_ATUAL como fallback
    if all_input_data.get('DATA_ART') == 'Não informado' or not all_input_data.get('DATA_ART'):
        all_input_data['DATA_ART'] = all_input_data['DATA_ATUAL']

    data_operacao_prevista = data_hoje + timedelta(days=90)
    all_input_data['DATA_OPERACAO_PREVISTA'] = data_operacao_prevista.strftime('%d/%m/%Y')

    try:
        potencia_wp_str = str(all_input_data.get('POTENCIA_MODULO_MANUAL', '0')).replace('Wp', '').strip().replace(',', '.')
        potencia_wp = float(potencia_wp_str)
        all_input_data['POTENCIA_MODULO_MANUAL_KWP'] = potencia_wp / 1000
    except ValueError:
        all_input_data['POTENCIA_MODULO_MANUAL_KWP'] = 'Não informado'

    # --- Inicializa o mapa de arquivos temporários e a lista para o ZIP na sessão ---
    session['temp_files_to_zip'] = []
    temp_zip_dir = tempfile.mkdtemp()

    # Limpa o Nome_Razao_Social para ser usado no nome do arquivo/pasta (removendo caracteres especiais)
    nome_razao_social_clean = re.sub(r'[\/:*?"<>|]', '', all_input_data.get('Nome_Razao_Social', 'Cliente')).strip()
    if not nome_razao_social_clean:
        nome_razao_social_clean = 'Cliente'
    session['nome_razao_social_zip_folder'] = nome_razao_social_clean


    current_excel_fv_template_path = EXCEL_PROJETO_FV_TEMPLATE_PATH
    current_certidao_registro_profissional_path = CERTIDAO_REGISTRO_PROFISSIONAL_PATH
    
    # --- Processar Planilha Projetos FV.xlsx (COMUM PARA AMBOS RGE E COOPERLUZ) ---
    temp_excel_proj_fv_filename_unique = f"dados_do_projeto_{uuid.uuid4().hex}.xlsx"
    temp_excel_proj_fv_filepath = os.path.join(temp_zip_dir, temp_excel_proj_fv_filename_unique)

    try:
        workbook_proj_fv = openpyxl.load_workbook(current_excel_fv_template_path)

        if 'DADOS' not in workbook_proj_fv.sheetnames:
            print("Erro: A aba 'DADOS' não foi encontrada no arquivo Excel 'Planilha Projetos FV.xlsx'.")
            return render_template('index.html', error="Erro: A aba 'DADOS' não foi encontrada no arquivo Excel 'Planilha Projetos FV.xlsx'.")

        sheet_dados_proj_fv = workbook_proj_fv['DADOS']

        for key, cell_address in EXCEL_PROJETO_FV_CELL_MAPPING.items():
            value_to_write = all_input_data.get(key, 'Não encontrado')

            if key in ['POTENCIA_MODULO_MANUAL', 'POTENCIA_INVERSOR_MANUAL']:
                cleaned_value = str(value_to_write).replace('Wp', '').strip().replace(',', '.')
                try:
                    num_value = float(cleaned_value)
                    sheet_dados_proj_fv[cell_address] = num_value
                except ValueError:
                    sheet_dados_proj_fv[cell_address] = ''
                continue

            # Nao escreve o INMETRO da entrada manual no excel, pois ele sera calculado.
            # O INMETRO no excel template eh uma formula em J29
            if key == 'INMETRO':
                continue

            if value_to_write == 'Não encontrado' or value_to_write == 'Não informado' or str(value_to_write).strip() == '':
                value_to_write = ''

            if key == 'Endereco_Rua' or key == 'Numero_Endereco':
                full_address = all_input_data.get('Endereco_Rua_Numero', '')
                street, number = parse_address_for_excel(str(full_address))
                if key == 'Endereco_Rua':
                    sheet_dados_proj_fv[cell_address] = street
                elif key == 'Numero_Endereco':
                    sheet_dados_proj_fv[cell_address] = number
            else:
                # Escreve o valor já formatado para o Excel
                if key in NUMERIC_KEYS_FOR_FORMATTING:
                    sheet_dados_proj_fv[cell_address] = format_value_for_display(value_to_write, is_numeric=True)
                else:
                    sheet_dados_proj_fv[cell_address] = value_to_write

        workbook_proj_fv.save(temp_excel_proj_fv_filepath)

        if win32:
            if not calculate_excel_formulas(temp_excel_proj_fv_filepath):
                print("Aviso: O recálculo das fórmulas do Excel falhou. Os valores calculados podem estar incorretos.")
        else:
            print("Aviso: 'pywin32' não está disponível. Os valores calculados do Excel podem não ser atualizados automaticamente.")

        workbook_proj_fv_calculated = openpyxl.load_workbook(temp_excel_proj_fv_filepath, data_only=True)
        sheet_dados_proj_fv_calculated = workbook_proj_fv_calculated['DADOS']

        for key_calc, cell_calc in EXCEL_PROJETO_FV_CALCULATED_CELLS.items():
            calculated_value = sheet_dados_proj_fv_calculated[cell_calc].value
            if calculated_value is not None and not (isinstance(calculated_value, str) and (str(calculated_value).startswith('#') or str(calculated_value) == '')): # Convert to str for startswith check
                processed_value_calc = str(calculated_value)
                # Apply specific cleaning for spacing in mm²
                if key_calc in ['RAMAL_ENTRADA_CALCULADO', 'CABO_CA']:
                    processed_value_calc = re.sub(r'\s*mm²', 'mm²', processed_value_calc) # Remove space before mm²
                all_input_data[key_calc] = processed_value_calc
            else:
                all_input_data[key_calc] = 'Não calculado' # Default if calculation fails or is empty

        session['temp_files_to_zip'].append({
            'path': temp_excel_proj_fv_filepath,
            'zip_filename': 'Dados do Projeto.xlsx'
        })
        print(f"Dados escritos e calculados no Excel de Projeto FV temporário: {temp_excel_proj_fv_filepath}.")

    except FileNotFoundError:
        print(f"Erro: Arquivo Excel de template '{EXCEL_PROJETO_FV_TEMPLATE_FILENAME}' não encontrado.")
        return render_template('index.html', error=f"Erro: Arquivo Excel de template '{EXCEL_PROJETO_FV_TEMPLATE_FILENAME}' não encontrado.")
    except Exception as e:
        print(f"Erro ao processar/salvar dados no Excel de Projeto FV ou ler valores calculados: {e}")
        return render_template('index.html', error=f"Erro ao processar/salvar dados no Excel de Projeto FV ou ler valores calculados: {e}")


    # --- Processar ART.txt (COMUM PARA AMBOS RGE E COOPERLUZ) ---
    temp_art_filename_unique = f"art_preenchida_{uuid.uuid4().hex}.txt"
    temp_art_filepath = os.path.join(temp_zip_dir, temp_art_filename_unique)
    try:
        art_content = generate_art_txt_content(all_input_data)
        with open(temp_art_filepath, 'w', encoding='utf-8') as f:
            f.write(art_content)
        session['temp_files_to_zip'].append({
            'path': temp_art_filepath,
            'zip_filename': '2. ART - PREENCHIDO.txt'
        })
        print(f"Conteúdo da ART gerado em: {temp_art_filepath}.")
    except Exception as e:
        print(f"Erro ao gerar ART.txt: {e}")
        return render_template('index.html', error=f"Erro ao gerar ART.txt: {e}")
    
    # --- Processar Upload de Imagens e Gerar PDF (COMUM PARA AMBOS RGE E COOPERLUZ) ---
    image_files = {}
    image_files = {
        'foto_disjuntor': request.files.get('foto_disjuntor'),
        'foto_fachada': request.files.get('foto_fachada'),
        'foto_entrada_energia': request.files.get('foto_entrada_energia')
    }

    image_paths_for_pdf = []
    titles_for_pdf = {
        'foto_disjuntor': 'Foto do Disjuntor',
        'foto_fachada': 'Foto da Fachada',
        'foto_entrada_energia': 'Foto da Entrada de Energia'
    }

    for key, file_obj in image_files.items():
        if file_obj and file_obj.filename != '' and allowed_file(file_obj.filename, ALLOWED_IMAGE_EXTENSIONS):
            image_filename = secure_filename(file_obj.filename)
            temp_image_filepath = os.path.join(temp_zip_dir, image_filename)
            file_obj.save(temp_image_filepath)
            image_paths_for_pdf.append({'path': temp_image_filepath, 'title': titles_for_pdf[key]})
        else:
            print(f"Aviso: Imagem para '{titles_for_pdf.get(key, key)}' não foi fornecida ou é de tipo inválido.")


    temp_images_pdf_filename_unique = f"fotos_entrada_fachada_{uuid.uuid4().hex}.pdf"
    temp_images_pdf_filepath = os.path.join(temp_zip_dir, temp_images_pdf_filename_unique)

    if image_paths_for_pdf:
        try:
            generate_images_pdf(image_paths_for_pdf, temp_images_pdf_filepath)
            session['temp_files_to_zip'].append({
                'path': temp_images_pdf_filepath,
                'zip_filename': '6. 7. Foto Geral da Entrada de Energia - Fachada.pdf'
            })
            print(f"PDF de imagens gerado em: {temp_images_pdf_filepath}.")
        except Exception as e:
            print(f"Erro ao gerar PDF de imagens: {e}")
            temp_error_txt_filepath = os.path.join(temp_zip_dir, f"erro_geracao_fotos_{uuid.uuid4().hex}.txt")
            with open(temp_error_txt_filepath, 'w', encoding='utf-8') as f:
                f.write(f"Erro ao gerar o PDF de imagens: {e}. Por favor, verifique os arquivos enviados.")
            session['temp_files_to_zip'].append({
                'path': temp_error_txt_filepath,
                'zip_filename': '6. 7. Foto Geral da Entrada de Energia - Fachada - ERRO.txt'
            })
    else:
        temp_empty_image_pdf_placeholder = os.path.join(temp_zip_dir, f"placeholder_fotos_vazio_{uuid.uuid4().hex}.txt")
        with open(temp_empty_image_pdf_placeholder, 'w', encoding='utf-8') as f:
            f.write("Nenhuma imagem de disjuntor, fachada ou entrada de energia foi fornecida. Este arquivo é um placeholder.")
        session['temp_files_to_zip'].append({
            'path': temp_empty_image_pdf_placeholder,
            'zip_filename': '6. 7. Foto Geral da Entrada de Energia - Fachada.txt'
        })

    # --- Adicionar Certidão de Registro Profissional (COMUM PARA AMBOS RGE E COOPERLUZ) ---
    session['temp_files_to_zip'].append({
        'path': current_certidao_registro_profissional_path,
        'zip_filename': '1. Certidao de Registro Profissional.pdf'
    })

    # --- Adicionar placeholders de documentos comuns (COMUM PARA AMBOS RGE E COOPERLUZ) ---
    empty_txt_files_common = {
        '9. 10. Documento de Identidade.txt': 'Conteúdo de placeholder para o documento de identidade.',
        '5. Certificado Inmetro Inversor Solar.txt': 'Conteúdo de placeholder para o certificado Inmetro do inversor solar.',
    }
    for zip_name, content in empty_txt_files_common.items():
        temp_empty_filepath = os.path.join(temp_zip_dir, f"empty_placeholder_{uuid.uuid4().hex}.txt")
        with open(temp_empty_filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        session['temp_files_to_zip'].append({
            'path': temp_empty_filepath,
            'zip_filename': zip_name
        })

    # --- Processar ANEXOS ESPECÍFICOS POR DISTRIBUIDORA ---
    if distributor_type == 'RGE':
        # --- Gerar Postagem do Projeto (ESPECÍFICO PARA RGE) ---
        temp_postagem_filename_unique = f"postagem_{distributor_type.lower()}_{uuid.uuid4().hex}.txt"
        temp_postagem_filepath = os.path.join(temp_zip_dir, temp_postagem_filename_unique)
        try:
            postagem_content = generate_postagem_txt_content(all_input_data, distributor_type)
            with open(temp_postagem_filepath, 'w', encoding='utf-8') as f:
                f.write(postagem_content)
            session['temp_files_to_zip'].append({
                'path': temp_postagem_filepath,
                'zip_filename': f'Postagem do projeto no site da {distributor_type}.txt'
            })
            print(f"Conteúdo da Postagem {distributor_type} gerado em: {temp_postagem_filepath}.")
        except Exception as e:
            print(f"Erro ao gerar Postagem_{distributor_type}.txt: {e}")
            return render_template('index.html', error=f"Erro ao gerar Postagem_{distributor_type}.txt: {e}")

        # --- Templates para RGE ---
        current_anexo_f_template_path = ANEXO_F_TEMPLATE_PATH
        current_anexo_e_template_path = ANEXO_E_TEMPLATE_PATH
        current_termo_aceite_template_path = TERMO_ACEITE_TEMPLATE_PATH

        # --- Processar Formulário Anexo F - GED 15303.xlsx ---
        temp_excel_anexo_f_filename_unique = f"anexo_f_preenchido_{uuid.uuid4().hex}.xlsx"
        temp_excel_anexo_f_filepath = os.path.join(temp_zip_dir, temp_excel_anexo_f_filename_unique)

        try:
            workbook_anexo_f = openpyxl.load_workbook(current_anexo_f_template_path)

            if ANEXO_F_SHEET_NAME not in workbook_anexo_f.sheetnames:
                print(f"Erro: A aba '{ANEXO_F_SHEET_NAME}' não foi encontrada no arquivo Excel '{ANEXO_F_TEMPLATE_FILENAME}'.")
                return render_template('index.html', error=f"Erro: A aba '{ANEXO_F_SHEET_NAME}' não foi encontrada no arquivo Excel '{ANEXO_F_TEMPLATE_FILENAME}'.")

            sheet_anexo_f = workbook_anexo_f[ANEXO_F_SHEET_NAME]

            for row in sheet_anexo_f.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        cell_text_standardized = cell.value.strip().upper()

                        if cell_text_standardized in ANEXO_F_PLACEHOLDER_TO_PYTHON_VAR:
                            python_var_name = ANEXO_F_PLACEHOLDER_TO_PYTHON_VAR[cell_text_standardized]
                            value_to_write = all_input_data.get(python_var_name)

                            if python_var_name == 'FULL_ADDRESS_COMPOSED':
                                rua_num_completo = all_input_data.get('Endereco_Rua_Numero', '')
                                bairro_comp = all_input_data.get('Bairro', '')

                                rua_parsed, numero_parsed = parse_address_for_excel(str(rua_num_completo))

                                full_address_str = f"{rua_parsed}"
                                if numero_parsed:
                                    full_address_str += f", {numero_parsed}"
                                if bairro_comp:
                                    full_address_str += f" - {bairro_comp}"
                                cell.value = full_address_str.strip(' ,-')
                            elif python_var_name in DATE_KEYS_FOR_FORMATTING:
                                if value_to_write:
                                    try:
                                        dt_obj = datetime.strptime(value_to_write, '%d/%m/%Y')
                                        cell.number_format = 'DD/MM/YYYY'
                                        cell.value = dt_obj
                                    except ValueError:
                                        cell.value = str(value_to_write)
                                else:
                                    cell.value = ''
                            elif python_var_name in NUMERIC_KEYS_FOR_FORMATTING:
                                cell.value = format_value_for_display(value_to_write, is_numeric=True)
                            else:
                                cell.value = str(value_to_write) if value_to_write is not None else ''

                            original_font = cell.font
                            cell.font = openpyxl.styles.Font(color='00000000', name=original_font.name, size=original_font.size)

            workbook_anexo_f.save(temp_excel_anexo_f_filepath)
            session['temp_files_to_zip'].append({
                'path': temp_excel_anexo_f_filepath,
                'zip_filename': '3. Formulário Anexo F - PREENCHIDO.xlsx'
            })
            print(f"Dados escritos no Anexo F temporário: {temp_excel_anexo_f_filepath}.")

        except FileNotFoundError:
            print(f"Erro: Arquivo Excel de template '{ANEXO_F_TEMPLATE_FILENAME}' não encontrado.")
            return render_template('index.html', error=f"Erro: Arquivo Excel de template '{ANEXO_F_TEMPLATE_FILENAME}' não encontrado.")
        except Exception as e:
            print(f"Erro ao processar/salvar dados no Anexo F: {e}")
            return render_template('index.html', error=f"Erro ao processar/salvar dados no Anexo F: {e}")


        # --- Processar Formulário Anexo E - GED.docx ---
        temp_anexo_e_filename_unique = f"anexo_e_preenchido_{uuid.uuid4().hex}.docx"
        temp_anexo_e_filepath = os.path.join(temp_zip_dir, temp_anexo_e_filename_unique)

        try:
            docx_replacements_anexo_e = {
                'UC': get_formatted_value_for_doc('UC', all_input_data),
                'ENDERECO_RUA_NUMERO': get_formatted_value_for_doc('Endereco_Rua_Numero', all_input_data),
                'Bairro': get_formatted_value_for_doc('Bairro', all_input_data),
                'Cidade': get_formatted_value_for_doc('Cidade', all_input_data),
                'ESTADO': get_formatted_value_for_doc('Estado', all_input_data),
                'CEP': get_formatted_value_for_doc('CEP', all_input_data),
                'TELEFONE': get_formatted_value_for_doc('TELEFONE', all_input_data),
                'E-MAIL': get_formatted_value_for_doc('E-MAIL', all_input_data),
                'CARGA_INSTALADA': get_formatted_value_for_doc('CARGA_INSTALADA', all_input_data),
                'CATEGORIA': get_formatted_value_for_doc('CATEGORIA', all_input_data),
                'CLASSE_TARIFARIA': get_formatted_value_for_doc('Classe_Tarifaria', all_input_data),
                'GRUPO_TARIFARIO': get_formatted_value_for_doc('Grupo_Tarifario', all_input_data),
                'TENSAO_NOMINAL_V': get_formatted_value_for_doc('Tensao_Nominal_V', all_input_data),
                'NUMERO_FASES_CALCULADO': get_formatted_value_for_doc('NUMERO_FASES_CALCULADO', all_input_data),
                'POTENCIA_TOTAL_SISTEMA_KWP': get_formatted_value_for_doc('POTENCIA_TOTAL_SISTEMA_KWP', all_input_data),
                'FABRICANTE_INVERSOR_MANUAL': get_formatted_value_for_doc('FABRICANTE_INVERSOR_MANUAL', all_input_data),
                'MODELO_INVERSOR_CALCULADO': get_formatted_value_for_doc('MODELO_INVERSOR_CALCULADO', all_input_data),
                'QUANTIDADE_INVERSOR_MANUAL': get_formatted_value_for_doc('QUANTIDADE_INVERSOR_MANUAL', all_input_data),
                'POTENCIA_INVERSOR_MANUAL': get_formatted_value_for_doc('POTENCIA_INVERSOR_MANUAL', all_input_data),
                'Nome_Razao_Social': get_formatted_value_for_doc('Nome_Razao_Social', all_input_data),
                'DATA_ATUAL': get_formatted_value_for_doc('DATA_ATUAL', all_input_data),
                'LATITUDE': get_formatted_value_for_doc('LATITUDE', all_input_data),
                'LONGITUDE': get_formatted_value_for_doc('LONGITUDE', all_input_data),
                'LATITUDE_GMS': get_formatted_value_for_doc('LATITUDE_GMS', all_input_data),
                'LONGITUDE_GMS': get_formatted_value_for_doc('LONGITUDE_GMS', all_input_data),
            }

            # DEBUG para o Anexo E
            print(f"DEBUG DOCX Anexo E: Cidade: {all_input_data.get('Cidade')}")
            print(f"DEBUG DOCX Anexo E: Estado: {all_input_data.get('Estado')}")
            print(f"DEBUG DOCX Anexo E: DATA_ATUAL: {all_input_data.get('DATA_ATUAL')}")

            # Note: No loop abaixo, 'value' já é uma string formatada se get_formatted_value_for_doc foi usado corretamente
            # Não é mais necessário o loop `for key, value in docx_replacements_anexo_e.items(): docx_replacements_anexo_e[key] = str(value)`
            # pois get_formatted_value_for_doc já retorna string.
            
            doc_preenchido = replace_docx_placeholders(current_anexo_e_template_path, docx_replacements_anexo_e)

            doc_preenchido.save(temp_anexo_e_filepath)
            session['temp_files_to_zip'].append({
                'path': temp_anexo_e_filepath,
                'zip_filename': '8. Formulário Anexo E - PREENCHIDO.docx'
            })
            print(f"Dados escritos no Anexo E temporário: {temp_anexo_e_filepath}.")

        except FileNotFoundError:
            print(f"Erro: Arquivo DOCX de template '{ANEXO_E_TEMPLATE_FILENAME}' não encontrado.")
            return render_template('index.html', error=f"Erro: Arquivo DOCX de template '{ANEXO_E_TEMPLATE_FILENAME}' não encontrado.")
        except Exception as e:
            print(f"Erro ao processar/salvar dados no Anexo E: {e}")
            return render_template('index.html', error=f"Erro ao processar/salvar dados no Anexo E: {e}")

        # --- Processar Termo de Aceite.docx ---
        temp_termo_aceite_filename_unique = f"termo_aceite_preenchido_{uuid.uuid4().hex}.docx"
        temp_termo_aceite_filepath = os.path.join(temp_zip_dir, temp_termo_aceite_filename_unique)

        try:
            docx_replacements_termo_aceite = {
                'UC': get_formatted_value_for_doc('UC', all_input_data),
                'CIDADE': get_formatted_value_for_doc('Cidade', all_input_data),
                'ESTADO': get_formatted_value_for_doc('Estado', all_input_data),
                'DATA_ATUAL': get_formatted_value_for_doc('DATA_ATUAL', all_input_data),
                'Nome_Razao_Social': get_formatted_value_for_doc('Nome_Razao_Social', all_input_data),
                'CPF_CNPJ': get_formatted_value_for_doc('CNPJ_CPF', all_input_data),
                'LATITUDE': get_formatted_value_for_doc('LATITUDE', all_input_data),
                'LONGITUDE': get_formatted_value_for_doc('LONGITUDE', all_input_data),
                'LATITUDE_GMS': get_formatted_value_for_doc('LATITUDE_GMS', all_input_data),
                'LONGITUDE_GMS': get_formatted_value_for_doc('LONGITUDE_GMS', all_input_data),
            }

            doc_termo_aceite_preenchido = replace_docx_placeholders(current_termo_aceite_template_path, docx_replacements_termo_aceite)

            doc_termo_aceite_preenchido.save(temp_termo_aceite_filepath)
            session['temp_files_to_zip'].append({
                'path': temp_termo_aceite_filepath,
                'zip_filename': '11. Termo de Aceite - PREENCHIDO.docx'
            })
            print(f"Dados escritos no Termo de Aceite temporário: {temp_termo_aceite_filepath}.")

        except FileNotFoundError:
            print(f"Erro: Arquivo DOCX de template '{TERMO_ACEITE_TEMPLATE_FILENAME}' não encontrado.")
            return render_template('index.html', error=f"Erro: Arquivo DOCX de template '{TERMO_ACEITE_TEMPLATE_FILENAME}' não encontrado.")
        except Exception as e:
            print(f"Erro ao processar/salvar dados no Termo de Aceite: {e}")
            return render_template('index.html', error=f"Erro ao processar/salvar dados no Termo de Aceite: {e}")

        # Se for RGE, e não processamos um Anexo I, então precisamos do Projeto.txt
        temp_empty_filepath = os.path.join(temp_zip_dir, f"empty_placeholder_projeto_{uuid.uuid4().hex}.txt")
        with open(temp_empty_filepath, 'w', encoding='utf-8') as f:
            f.write('Conteúdo de placeholder para o projeto.')
        session['temp_files_to_zip'].append({
            'path': temp_empty_filepath,
            'zip_filename': '4. Projeto.txt'
        })

    elif distributor_type in ['COOPERLUZ', 'CERTHIL', 'CERMISSOES']: # ATUALIZADO: Inclui Certhil e Cermissões
        # --- Templates para Cooperluz/Certhil/Cermissões ---
        current_anexo_i_template_path = ANEXO_I_TEMPLATE_PATH
        current_procuracao_template_path = PROCURACAO_TEMPLATE_PATH
        current_termo_aceite_inciso_iii_template_path = TERMO_ACEITE_INCISO_III_TEMPLATE_PATH
        current_responsabilidade_tecnica_template_path = RESPONSABILIDADE_TECNICA_TEMPLATE_PATH
        current_dados_gd_ufv_template_path = DADOS_GD_UFV_TEMPLATE_PATH
        current_memorial_descritivo_template_path = MEMORIAL_DESCRITIVO_TEMPLATE_PATH # NOVO

        # --- Processar Anexo I.xlsx ---
        temp_anexo_i_filename_unique = f"anexo_i_preenchido_{uuid.uuid4().hex}.xlsx"
        temp_anexo_i_filepath = os.path.join(temp_zip_dir, temp_anexo_i_filename_unique)

        try:
            workbook_anexo_i = openpyxl.load_workbook(current_anexo_i_template_path)

            if ANEXO_I_SHEET_NAME not in workbook_anexo_i.sheetnames:
                print(f"Erro: A aba '{ANEXO_I_SHEET_NAME}' não foi encontrada no arquivo Excel '{ANEXO_I_TEMPLATE_FILENAME}'.")
                # Fallback para um placeholder se o template não for encontrado
                temp_error_txt_filepath = os.path.join(temp_zip_dir, f"erro_anexo_i_template_nao_encontrado_{uuid.uuid4().hex}.txt")
                with open(temp_error_txt_filepath, 'w', encoding='utf-8') as f:
                    f.write(f"Erro ao gerar Anexo I: Template '{ANEXO_I_TEMPLATE_FILENAME}' não encontrado.")
                session['temp_files_to_zip'].append({
                    'path': temp_error_txt_filepath,
                    'zip_filename': '4. Anexo I - ERRO (Template não encontrado).txt'
                })
            else:
                sheet_anexo_i = workbook_anexo_i[ANEXO_I_SHEET_NAME]

                for row in sheet_anexo_i.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            cell_text_standardized = cell.value.strip().upper() # Padroniza texto da célula para comparação

                            # Verifica se o texto da célula corresponde a um placeholder mapeado
                            if cell_text_standardized in ANEXO_I_PLACEHOLDER_TO_PYTHON_VAR:
                                python_var_name = ANEXO_I_PLACEHOLDER_TO_PYTHON_VAR[cell_text_standardized]
                                value_to_write = all_input_data.get(python_var_name)

                                # Formatação específica para datas, se houver placeholders de data
                                if python_var_name in DATE_KEYS_FOR_FORMATTING:
                                    if value_to_write:
                                        try:
                                            dt_obj = datetime.strptime(value_to_write, '%d/%m/%Y')
                                            cell.number_format = 'DD/MM/YYYY' # Garante o formato de data
                                            cell.value = dt_obj
                                        except ValueError:
                                            cell.value = str(value_to_write) # Se falhar, escreve como texto
                                    else:
                                        cell.value = ''
                                elif python_var_name in NUMERIC_KEYS_FOR_FORMATTING:
                                    # Aplica a formatação numérica para exibição no padrão brasileiro
                                    cell.value = format_value_for_display(value_to_write, is_numeric=True)
                                else:
                                    cell.value = str(value_to_write) if value_to_write is not None else ''

                                # Opcional: manter a cor da fonte original (ou definir para preto)
                                original_font = cell.font
                                cell.font = openpyxl.styles.Font(color='00000000', name=original_font.name, size=original_font.size) # Define cor preta
                
                workbook_anexo_i.save(temp_anexo_i_filepath)
                session['temp_files_to_zip'].append({
                    'path': temp_anexo_i_filepath,
                    'zip_filename': '4. Anexo I - PREENCHIDO.xlsx' # Número 4 para o Anexo I
                })
                print(f"Dados escritos no Anexo I temporário: {temp_anexo_i_filepath}.")

        except FileNotFoundError:
            print(f"Erro: Arquivo Excel de template '{ANEXO_I_TEMPLATE_FILENAME}' não encontrado.")
            return render_template('index.html', error=f"Erro: Arquivo Excel de template '{ANEXO_I_TEMPLATE_FILENAME}' não encontrado.")
        except Exception as e:
            print(f"Erro ao processar/salvar dados no Anexo I: {e}")
            temp_error_txt_filepath = os.path.join(temp_zip_dir, f"erro_anexo_i_{uuid.uuid4().hex}.txt")
            with open(temp_error_txt_filepath, 'w', encoding='utf-8') as f:
                f.write(f"Erro ao gerar Anexo I: {e}")
            session['temp_files_to_zip'].append({
                'path': temp_error_txt_filepath,
                'zip_filename': '4. Anexo I - ERRO.txt'
            })
        
        # --- Processar Procuracao.docx (ESPECÍFICO PARA COOPERLUZ/CERTHIL/CERMISSOES) ---
        temp_procuracao_filename_unique = f"procuracao_preenchida_{uuid.uuid4().hex}.docx"
        temp_procuracao_filepath = os.path.join(temp_zip_dir, temp_procuracao_filename_unique)

        # 1. Compor ENDERECO_COMPLETO_OUTORGANTE (já estava correto)
        full_address_outorgante_parts = []
        rua_num_out = all_input_data.get('Endereco_Rua_Numero', '').strip()
        bairro_out = all_input_data.get('Bairro', '').strip()
        cidade_est_out = all_input_data.get('CIDADE_ESTADO', '').strip()
        cep_out = all_input_data.get('CEP', '').strip() # Incluir CEP também

        if rua_num_out:
            full_address_outorgante_parts.append(rua_num_out)
        if bairro_out:
            full_address_outorgante_parts.append(f"– {bairro_out}") # Formato "Rua, Num – Bairro"
        if cidade_est_out:
            full_address_outorgante_parts.append(f"na cidade de {cidade_est_out}") # Formato "..., na cidade de Santa Rosa - RS"
        if cep_out:
             full_address_outorgante_parts.append(f"CEP {cep_out}") # Formato "CEP 98797-899"

        all_input_data['ENDERECO_COMPLETO_OUTORGANTE'] = ", ".join(filter(None, full_address_outorgante_parts)) # Une, ignorando vazios

        # 2. Compor CIDADE_ESTADO_DATA_ASSINATURA (já estava correto)
        cidade_estado_data_str = f"{format_value_for_display(all_input_data.get('Cidade', '')).upper()} – {format_value_for_display(all_input_data.get('Estado', '')).upper()}, {get_formatted_value_for_doc('DATA_ATUAL', all_input_data)}" # Use a função auxiliar para DATA_ATUAL
        all_input_data['CIDADE_ESTADO_DATA_ASSINATURA'] = cidade_estado_data_str


        try:
            docx_replacements_procuracao = {
                'NOME_RAZAO_SOCIAL': get_formatted_value_for_doc('Nome_Razao_Social', all_input_data),
                'CPF_CNPJ': get_formatted_value_for_doc('CNPJ_CPF', all_input_data),
                'ENDERECO_RUA_NUMERO': get_formatted_value_for_doc('Endereco_Rua_Numero', all_input_data),
                'BAIRRO': get_formatted_value_for_doc('Bairro', all_input_data),
                'CIDADE': get_formatted_value_for_doc('Cidade', all_input_data),
                'ESTADO': get_formatted_value_for_doc('Estado', all_input_data),
                'UC': get_formatted_value_for_doc('UC', all_input_data),
                'DATA_ATUAL': get_formatted_value_for_doc('DATA_ATUAL', all_input_data),
                'ENDERECO_COMPLETO_OUTORGANTE': get_formatted_value_for_doc('ENDERECO_COMPLETO_OUTORGANTE', all_input_data),
                'CIDADE_ESTADO_DATA_ASSINATURA': get_formatted_value_for_doc('CIDADE_ESTADO_DATA_ASSINATURA', all_input_data),
            }

            doc_procuracao_preenchida = replace_docx_placeholders(current_procuracao_template_path, docx_replacements_procuracao)
            doc_procuracao_preenchida.save(temp_procuracao_filepath)
            session['temp_files_to_zip'].append({
                'path': temp_procuracao_filepath,
                'zip_filename': '12. Procuração - PREENCHIDO.docx'
            })
            print(f"Dados escritos na Procuração temporária: {temp_procuracao_filepath}.")

        except FileNotFoundError:
            print(f"Erro: Arquivo DOCX de template '{PROCURACAO_TEMPLATE_FILENAME}' não encontrado.")
            temp_error_txt_filepath = os.path.join(temp_zip_dir, f"erro_procuracao_template_nao_encontrado_{uuid.uuid4().hex}.txt")
            with open(temp_error_txt_filepath, 'w', encoding='utf-8') as f:
                f.write(f"Erro ao gerar Procuração: Template '{PROCURACAO_TEMPLATE_FILENAME}' não encontrado.")
            session['temp_files_to_zip'].append({
                'path': temp_error_txt_filepath,
                'zip_filename': '12. Procuração - ERRO (Template não encontrado).txt'
            })
        except Exception as e:
            print(f"Erro ao processar/salvar dados na Procuração: {e}")
            temp_error_txt_filepath = os.path.join(temp_zip_dir, f"erro_procuracao_{uuid.uuid4().hex}.txt")
            with open(temp_error_txt_filepath, 'w', encoding='utf-8') as f:
                f.write(f"Erro ao gerar Procuração: {e}")
            session['temp_files_to_zip'].append({
                'path': temp_error_txt_filepath,
                'zip_filename': '12. Procuração - ERRO.txt'
            })

        # --- Processar Termo de Aceite Inciso III.docx (ESPECÍFICO PARA COOPERLUZ/CERTHIL/CERMISSOES) ---
        temp_termo_aceite_inciso_iii_filename_unique = f"termo_aceite_inciso_iii_preenchido_{uuid.uuid4().hex}.docx"
        temp_termo_aceite_inciso_iii_filepath = os.path.join(temp_zip_dir, temp_termo_aceite_inciso_iii_filename_unique)

        try:
            docx_replacements_termo_aceite_inciso_iii = {
                'NOME_RAZAO_SOCIAL': get_formatted_value_for_doc('Nome_Razao_Social', all_input_data),
                'CPF_CNPJ': get_formatted_value_for_doc('CNPJ_CPF', all_input_data),
                'UC': get_formatted_value_for_doc('UC', all_input_data),
                'DATA_ATUAL': get_formatted_value_for_doc('DATA_ATUAL', all_input_data),
            }

            doc_termo_aceite_inciso_iii_preenchido = replace_docx_placeholders(current_termo_aceite_inciso_iii_template_path, docx_replacements_termo_aceite_inciso_iii)
            doc_termo_aceite_inciso_iii_preenchido.save(temp_termo_aceite_inciso_iii_filepath)
            session['temp_files_to_zip'].append({
                'path': temp_termo_aceite_inciso_iii_filepath,
                'zip_filename': '13. Termo de Aceite Inciso III - PREENCHIDO.docx'
            })
            print(f"Dados escritos no Termo de Aceite Inciso III temporário: {temp_termo_aceite_inciso_iii_filepath}.")

        except FileNotFoundError:
            print(f"Erro: Arquivo DOCX de template '{TERMO_ACEITE_INCISO_III_TEMPLATE_FILENAME}' não encontrado.")
            temp_error_txt_filepath = os.path.join(temp_zip_dir, f"erro_termo_aceite_inciso_iii_template_nao_encontrado_{uuid.uuid4().hex}.txt")
            with open(temp_error_txt_filepath, 'w', encoding='utf-8') as f:
                f.write(f"Erro ao gerar Termo de Aceite Inciso III: Template '{TERMO_ACEITE_INCISO_III_TEMPLATE_FILENAME}' não encontrado.")
            session['temp_files_to_zip'].append({
                'path': temp_error_txt_filepath,
                'zip_filename': '13. Termo de Aceite Inciso III - ERRO (Template não encontrado).txt'
            })
        except Exception as e:
            print(f"Erro ao processar/salvar dados no Termo de Aceite Inciso III: {e}")
            temp_error_txt_filepath = os.path.join(temp_zip_dir, f"erro_termo_aceite_inciso_iii_{uuid.uuid4().hex}.txt")
            with open(temp_error_txt_filepath, 'w', encoding='utf-8') as f:
                f.write(f"Erro ao gerar Termo de Aceite Inciso III: {e}")
            session['temp_files_to_zip'].append({
                'path': temp_error_txt_filepath,
                'zip_filename': '13. Termo de Aceite Inciso III - ERRO.txt'
            })
        
        # --- Processar Responsabilidade Tecnica.docx (ESPECÍFICO PARA COOPERLUZ/CERTHIL/CERMISSOES) ---
        temp_responsabilidade_tecnica_filename_unique = f"responsabilidade_tecnica_preenchido_{uuid.uuid4().hex}.docx"
        temp_responsabilidade_tecnica_filepath = os.path.join(temp_zip_dir, temp_responsabilidade_tecnica_filename_unique)

        try:
            docx_replacements_responsabilidade_tecnica = {
                'POTENCIA_TOTAL_SISTEMA_KWP': get_formatted_value_for_doc('POTENCIA_TOTAL_SISTEMA_KWP', all_input_data),
                'Tensao_Nominal_V': get_formatted_value_for_doc('Tensao_Nominal_V', all_input_data),
                'NUMERO_FASES_CALCULADO': get_formatted_value_for_doc('NUMERO_FASES_CALCULADO', all_input_data),
                'UC': get_formatted_value_for_doc('UC', all_input_data),
                'NOME_RAZAO_SOCIAL': get_formatted_value_for_doc('Nome_Razao_Social', all_input_data),
                'ENDERECO_RUA_NUMERO': get_formatted_value_for_doc('Endereco_Rua_Numero', all_input_data),
                'BAIRRO': get_formatted_value_for_doc('Bairro', all_input_data),
                'CIDADE': get_formatted_value_for_doc('Cidade', all_input_data),
                'ESTADO': get_formatted_value_for_doc('Estado', all_input_data),
                'ART': get_formatted_value_for_doc('ART', all_input_data),
                'MODELO_INVERSOR_CALCULADO': get_formatted_value_for_doc('MODELO_INVERSOR_CALCULADO', all_input_data),
                'POTENCIA_INVERSOR_MANUAL': get_formatted_value_for_doc('POTENCIA_INVERSOR_MANUAL', all_input_data),
                'DATA_ATUAL': get_formatted_value_for_doc('DATA_ATUAL', all_input_data),
                'INMETRO': get_formatted_value_for_doc('INMETRO', all_input_data), # Adicionado INMETRO
            }

            print(f"DEBUG: docx_replacements_responsabilidade_tecnica content: {docx_replacements_responsabilidade_tecnica}")

            doc_responsabilidade_tecnica_preenchido = replace_docx_placeholders(current_responsabilidade_tecnica_template_path, docx_replacements_responsabilidade_tecnica)
            doc_responsabilidade_tecnica_preenchido.save(temp_responsabilidade_tecnica_filepath)
            session['temp_files_to_zip'].append({
                'path': temp_responsabilidade_tecnica_filepath,
                'zip_filename': '14. Responsabilidade Tecnica - PREENCHIDO.docx'
            })
            print(f"Dados escritos na Responsabilidade Tecnica temporária: {temp_responsabilidade_tecnica_filepath}.")

        except FileNotFoundError:
            print(f"Erro: Arquivo DOCX de template '{RESPONSABILIDADE_TECNICA_TEMPLATE_FILENAME}' não encontrado.")
            temp_error_txt_filepath = os.path.join(temp_zip_dir, f"erro_responsabilidade_tecnica_template_nao_encontrado_{uuid.uuid4().hex}.txt")
            with open(temp_error_txt_filepath, 'w', encoding='utf-8') as f:
                f.write(f"Erro ao gerar Responsabilidade Tecnica: Template '{RESPONSABILIDADE_TECNICA_TEMPLATE_FILENAME}' não encontrado.")
            session['temp_files_to_zip'].append({
                'path': temp_error_txt_filepath,
                'zip_filename': '14. Responsabilidade Tecnica - ERRO (Template não encontrado).txt'
            })
        except Exception as e:
            print(f"Erro ao processar/salvar dados na Responsabilidade Tecnica: {e}")
            temp_error_txt_filepath = os.path.join(temp_zip_dir, f"erro_responsabilidade_tecnica_{uuid.uuid4().hex}.txt")
            with open(temp_error_txt_filepath, 'w', encoding='utf-8') as f:
                f.write(f"Erro ao gerar Responsabilidade Tecnica: {e}")
            session['temp_files_to_zip'].append({
                'path': temp_error_txt_filepath,
                'zip_filename': '14. Responsabilidade Tecnica - ERRO.txt'
            })

        # --- Processar Dados para GD de UFV.docx (ESPECÍFICO PARA COOPERLUZ/CERTHIL/CERMISSOES) ---
        temp_dados_gd_ufv_filename_unique = f"dados_gd_ufv_preenchido_{uuid.uuid4().hex}.docx"
        temp_dados_gd_ufv_filepath = os.path.join(temp_zip_dir, temp_dados_gd_ufv_filename_unique)

        try:
            docx_replacements_dados_gd_ufv = {
                'Classe_Tarifaria': get_formatted_value_for_doc('Classe_Tarifaria', all_input_data),
                'Grupo_Tarifario': get_formatted_value_for_doc('Grupo_Tarifario', all_input_data),
                'Cidade': get_formatted_value_for_doc('Cidade', all_input_data),
                'Estado': get_formatted_value_for_doc('Estado', all_input_data),
                'Endereco_Rua_Numero': get_formatted_value_for_doc('Endereco_Rua_Numero', all_input_data),
                'Bairro': get_formatted_value_for_doc('Bairro', all_input_data),
                'CEP': get_formatted_value_for_doc('CEP', all_input_data),
                'LATITUDE_GMS': get_formatted_value_for_doc('LATITUDE_GMS', all_input_data),
                'LONGITUDE_GMS': get_formatted_value_for_doc('LONGITUDE_GMS', all_input_data),
                'CNPJ_CPF': get_formatted_value_for_doc('CNPJ_CPF', all_input_data),
                'Nome_Razao_Social': get_formatted_value_for_doc('Nome_Razao_Social', all_input_data),
                'TELEFONE': get_formatted_value_for_doc('TELEFONE', all_input_data),
                'E-MAIL': get_formatted_value_for_doc('E-MAIL', all_input_data),
                'POTENCIA_TOTAL_SISTEMA_KWP': get_formatted_value_for_doc('POTENCIA_TOTAL_SISTEMA_KWP', all_input_data),
                'QUANTIDADE_PLACAS_MANUAL': get_formatted_value_for_doc('QUANTIDADE_PLACAS_MANUAL', all_input_data),
                'FABRICANTE_MODULO_MANUAL': get_formatted_value_for_doc('FABRICANTE_MODULO_MANUAL', all_input_data),
                'MODELO_MODULO_CALCULADO': get_formatted_value_for_doc('MODELO_MODULO_CALCULADO', all_input_data),
                'POTENCIA_INVERSOR_MANUAL': get_formatted_value_for_doc('POTENCIA_INVERSOR_MANUAL', all_input_data),
                'QUANTIDADE_INVERSOR_MANUAL': get_formatted_value_for_doc('QUANTIDADE_INVERSOR_MANUAL', all_input_data),
                'FABRICANTE_INVERSOR_MANUAL': get_formatted_value_for_doc('FABRICANTE_INVERSOR_MANUAL', all_input_data),
                'MODELO_INVERSOR_CALCULADO': get_formatted_value_for_doc('MODELO_INVERSOR_CALCULADO', all_input_data),
                'AREA': get_formatted_value_for_doc('AREA_ARRANJOS_CALCULADO', all_input_data),
                'DATA_OPERACAO_PREVISTA': get_formatted_value_for_doc('DATA_OPERACAO_PREVISTA', all_input_data),
                'POTENCIA_PICO_MODULOS': get_formatted_value_for_doc('POTENCIA_PICO_MODULOS', all_input_data), # Adicionado POTENCIA_PICO_MODULOS
                'POTENCIA_MODULO_MANUAL_KWP': get_formatted_value_for_doc('POTENCIA_MODULO_MANUAL_KWP', all_input_data), # Adicionado POTENCIA_MODULO_MANUAL_KWP
            }

            print(f"DEBUG: docx_replacements_dados_gd_ufv content: {docx_replacements_dados_gd_ufv}")

            doc_dados_gd_ufv_preenchido = replace_docx_placeholders(current_dados_gd_ufv_template_path, docx_replacements_dados_gd_ufv)
            doc_dados_gd_ufv_preenchido.save(temp_dados_gd_ufv_filepath)
            session['temp_files_to_zip'].append({
                'path': temp_dados_gd_ufv_filepath,
                'zip_filename': '15. Dados para GD de UFV - PREENCHIDO.docx'
            })
            print(f"Dados escritos no Dados para GD de UFV temporário: {temp_dados_gd_ufv_filepath}.")

        except FileNotFoundError:
            print(f"Erro: Arquivo DOCX de template '{DADOS_GD_UFV_TEMPLATE_FILENAME}' não encontrado.")
            temp_error_txt_filepath = os.path.join(temp_zip_dir, f"erro_dados_gd_ufv_template_nao_encontrado_{uuid.uuid4().hex}.txt")
            with open(temp_error_txt_filepath, 'w', encoding='utf-8') as f:
                f.write(f"Erro ao gerar Dados para GD de UFV: Template '{DADOS_GD_UFV_TEMPLATE_FILENAME}' não encontrado.")
            session['temp_files_to_zip'].append({
                'path': temp_error_txt_filepath,
                'zip_filename': '15. Dados para GD de UFV - ERRO (Template não encontrado).txt'
            })
        except Exception as e:
            print(f"Erro ao processar/salvar dados no Dados para GD de UFV: {e}")
            temp_error_txt_filepath = os.path.join(temp_zip_dir, f"erro_dados_gd_ufv_{uuid.uuid4().hex}.txt")
            with open(temp_error_txt_filepath, 'w', encoding='utf-8') as f:
                f.write(f"Erro ao gerar Dados para GD de UFV: {e}")
            session['temp_files_to_zip'].append({
                'path': temp_error_txt_filepath,
                'zip_filename': '15. Dados para GD de UFV - ERRO.txt'
            })

        # --- NOVO: Processar Memorial Descritivo Padrão Fecoergs.docx (ESPECÍFICO PARA COOPERLUZ/CERTHIL/CERMISSOES) ---
        temp_memorial_descritivo_filename_unique = f"memorial_descritivo_preenchido_{uuid.uuid4().hex}.docx"
        temp_memorial_descritivo_filepath = os.path.join(temp_zip_dir, temp_memorial_descritivo_filename_unique)

        try:
            docx_replacements_memorial_descritivo = {
                'Nome_Razao_Social': get_formatted_value_for_doc('Nome_Razao_Social', all_input_data),
                'UC': get_formatted_value_for_doc('UC', all_input_data),
                'Endereco_Rua_Numero': get_formatted_value_for_doc('Endereco_Rua_Numero', all_input_data),
                'Bairro': get_formatted_value_for_doc('Bairro', all_input_data),
                'Cidade': get_formatted_value_for_doc('Cidade', all_input_data),
                'Estado': get_formatted_value_for_doc('Estado', all_input_data),
                'LATITUDE_GMS': get_formatted_value_for_doc('LATITUDE_GMS', all_input_data),
                'LONGITUDE_GMS': get_formatted_value_for_doc('LONGITUDE_GMS', all_input_data),
                'ART': get_formatted_value_for_doc('ART', all_input_data),
                'QUANTIDADE_PLACAS_MANUAL': get_formatted_value_for_doc('QUANTIDADE_PLACAS_MANUAL', all_input_data),
                'FABRICANTE_MODULO_MANUAL': get_formatted_value_for_doc('FABRICANTE_MODULO_MANUAL', all_input_data),
                'MODELO_MODULO_CALCULADO': get_formatted_value_for_doc('MODELO_MODULO_CALCULADO', all_input_data),
                'POTENCIA_TOTAL_SISTEMA_KWP': get_formatted_value_for_doc('POTENCIA_TOTAL_SISTEMA_KWP', all_input_data),
                'QUANTIDADE_INVERSOR_MANUAL': get_formatted_value_for_doc('QUANTIDADE_INVERSOR_MANUAL', all_input_data),
                'FABRICANTE_INVERSOR_MANUAL': get_formatted_value_for_doc('FABRICANTE_INVERSOR_MANUAL', all_input_data),
                'MODELO_INVERSOR_CALCULADO': get_formatted_value_for_doc('MODELO_INVERSOR_CALCULADO', all_input_data),
                'POTENCIA_INVERSOR_MANUAL': get_formatted_value_for_doc('POTENCIA_INVERSOR_MANUAL', all_input_data),
                'DISJUNTOR_CA': get_formatted_value_for_doc('DISJUNTOR_CA', all_input_data),
                'DISJ_CA_INTR': get_formatted_value_for_doc('DISJ_CA_INTR', all_input_data),
                'DISJ_CA_TENS': get_formatted_value_for_doc('DISJ_CA_TENS', all_input_data),
                'DISJ_CA_ATEN': get_formatted_value_for_doc('DISJ_CA_ATEN', all_input_data),
                'ISOLACAO_CA': get_formatted_value_for_doc('ISOLACAO_CA', all_input_data),
                'CABO_CA': get_formatted_value_for_doc('CABO_CA', all_input_data),
                'DATA_ATUAL': get_formatted_value_for_doc('DATA_ATUAL', all_input_data),
                'POTENCIA_PICO_MODULOS': get_formatted_value_for_doc('POTENCIA_PICO_MODULOS', all_input_data), # Adicionado POTENCIA_PICO_MODULOS
                'POTENCIA_MODULO_MANUAL_KWP': get_formatted_value_for_doc('POTENCIA_MODULO_MANUAL_KWP', all_input_data), # Adicionado POTENCIA_MODULO_MANUAL_KWP
            }

            print(f"DEBUG: docx_replacements_memorial_descritivo content: {docx_replacements_memorial_descritivo}")

            doc_memorial_descritivo_preenchido = replace_docx_placeholders(current_memorial_descritivo_template_path, docx_replacements_memorial_descritivo)
            doc_memorial_descritivo_preenchido.save(temp_memorial_descritivo_filepath)
            session['temp_files_to_zip'].append({
                'path': temp_memorial_descritivo_filepath,
                'zip_filename': '16. Memorial Descritivo Padrão Fecoergs - PREENCHIDO.docx'
            })
            print(f"Dados escritos no Memorial Descritivo temporário: {temp_memorial_descritivo_filepath}.")

        except FileNotFoundError:
            print(f"Erro: Arquivo DOCX de template '{MEMORIAL_DESCRITIVO_TEMPLATE_FILENAME}' não encontrado.")
            temp_error_txt_filepath = os.path.join(temp_zip_dir, f"erro_memorial_descritivo_template_nao_encontrado_{uuid.uuid4().hex}.txt")
            with open(temp_error_txt_filepath, 'w', encoding='utf-8') as f:
                f.write(f"Erro ao gerar Memorial Descritivo: Template '{MEMORIAL_DESCRITIVO_TEMPLATE_FILENAME}' não encontrado.")
            session['temp_files_to_zip'].append({
                'path': temp_error_txt_filepath,
                'zip_filename': '16. Memorial Descritivo - ERRO (Template não encontrado).txt'
            })
        except Exception as e:
            print(f"Erro ao processar/salvar dados no Memorial Descritivo: {e}")
            temp_error_txt_filepath = os.path.join(temp_zip_dir, f"erro_memorial_descritivo_{uuid.uuid4().hex}.txt")
            with open(temp_error_txt_filepath, 'w', encoding='utf-8') as f:
                f.write(f"Erro ao gerar Memorial Descritivo: {e}")
            session['temp_files_to_zip'].append({
                'path': temp_error_txt_filepath,
                'zip_filename': '16. Memorial Descritivo - ERRO.txt'
            })

        # --- Adicionar os 3 novos arquivos TXT placeholders para COOPERLUZ/CERTHIL/CERMISSOES ---
        empty_txt_files_cooperluz_specific = {
            'Diagrama Unifilar indicando desde o ponto de conexão com a Cooperluz.txt': 'Conteúdo de placeholder para Diagrama Unifilar.',
            'Planta de Localizacao e Situacao.txt': 'Conteúdo de placeholder para Planta de Localização e Situação.',
            'Diagrama Multifilar indicando desde o ponto de conexão com a Cooperluz.txt': 'Conteúdo de placeholder para Diagrama Multifilar.',
        }
        for zip_name, content in empty_txt_files_cooperluz_specific.items():
            temp_empty_filepath = os.path.join(temp_zip_dir, f"empty_placeholder_{uuid.uuid4().hex}.txt")
            with open(temp_empty_filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            session['temp_files_to_zip'].append({
                'path': temp_empty_filepath,
                'zip_filename': zip_name
            })
        print(f"Placeholders de documentos TXT adicionados para {distributor_type}.")

    
    # Armazena o caminho do diretório temporário para futura limpeza
    session['temp_zip_dir'] = temp_zip_dir
    session['all_input_data_for_correction'] = all_input_data # Salva todos os dados para a página de sucesso e correção

    # Gerar um nome único para o arquivo ZIP final (com o nome do cliente)
    final_zip_filename_unique = f"Projeto de {nome_razao_social_clean}_{uuid.uuid4().hex}.zip"

    # A extracted_data_from_pdf será limpa ao voltar para a página de upload, ou quando show_process_data_form a pegar
    # current_process_data_form_data é limpa no início desta função process_and_save

    # Specific handling for Latitude and Longitude to use '.' and 6 decimal places for display in success.html
    # This was already done correctly, but ensure it uses the formatted value for display
    lat_val = all_input_data.get('LATITUDE')
    lon_val = all_input_data.get('LONGITUDE')
    
    display_latitude = 'Não informado'
    if lat_val is not None and str(lat_val).strip() != '':
        try:
            display_latitude = f"{float(str(lat_val).replace(',', '.')):.6f}"
        except ValueError:
            display_latitude = str(lat_val) # If it's a non-numeric string like "abc"
    
    display_longitude = 'Não informado'
    if lon_val is not None and str(lon_val).strip() != '':
        try:
            display_longitude = f"{float(str(lon_val).replace(',', '.')):.6f}"
        except ValueError:
            display_longitude = str(lon_val)


    # Prepare dados_cliente para success.html
    dados_cliente_for_display = {
        'Nome/Razão Social': format_value_for_display(all_input_data.get('Nome_Razao_Social')),
        'CNPJ/CPF': format_value_for_display(all_input_data.get('CNPJ_CPF')),
        'E-Mail': format_value_for_display(all_input_data.get('E-MAIL')),
        'Telefone': format_value_for_display(all_input_data.get('TELEFONE')),
        'Endereço': format_value_for_display(all_input_data.get('Endereco_Rua_Numero')),
        'Bairro': format_value_for_display(all_input_data.get('Bairro')),
        'Cidade': format_value_for_display(all_input_data.get('Cidade')),
        'Estado': format_value_for_display(all_input_data.get('Estado')),
        'CEP': format_value_for_display(all_input_data.get('CEP')),
        'UC': format_value_for_display(all_input_data.get('UC')),
        'Classe Tarifária': format_value_for_display(all_input_data.get('Classe_Tarifaria')),
        'Grupo Tarifário': format_value_for_display(all_input_data.get('Grupo_Tarifario')),
        'Tensão Nominal (V)': format_value_for_display(all_input_data.get('Tensao_Nominal_V'), is_numeric=True), 
        'Latitude': display_latitude,
        'Longitude': display_longitude,
    }

    # Prepare dados_entrada para success.html
    carga_instalada_display = format_value_for_display(all_input_data.get('CARGA_INSTALADA'), is_numeric=True)
    if carga_instalada_display != 'Não informado':
        carga_instalada_display += 'kW'

    dados_entrada_for_display = {
        'Categoria': format_value_for_display(all_input_data.get('CATEGORIA')),
        'Carga Instalada': carga_instalada_display,
        'Tipo de Atendimento': format_value_for_display(all_input_data.get('TIPO_DE_ATENDIMENTO')),
        'Tipo de Caixa': format_value_for_display(all_input_data.get('TIPO_DE_CAIXA')),
        'Isolação': format_value_for_display(all_input_data.get('ISOLACAO')),
        'Número de Fases': format_value_for_display(all_input_data.get('NUMERO_FASES_CALCULADO'), is_numeric=True),
        'Ramal de Entrada': format_value_for_display(all_input_data.get('RAMAL_ENTRADA_CALCULADO')),
        'Disjuntor': format_value_for_display(all_input_data.get('DISJUNTOR_CALCULADO'), is_numeric=True),
    }

    # Prepare dados_sistema_fv para success.html
    dados_sistema_fv_for_display = {
        'ART': format_value_for_display(all_input_data.get('ART')),
        'Data da ART': format_value_for_display(all_input_data.get('DATA_ART'), is_date=True, date_separator='-'),
        'INMETRO do Inversor': format_value_for_display(all_input_data.get('INMETRO')),
        'Quantidade de Placas': format_value_for_display(all_input_data.get('QUANTIDADE_PLACAS_MANUAL'), is_numeric=True),
        'Potência do Módulo (Wp)': format_value_for_display(all_input_data.get('POTENCIA_MODULO_MANUAL'), is_numeric=True),
        'Fabricante do Módulo': format_value_for_display(all_input_data.get('FABRICANTE_MODULO_MANUAL')),
        'Modelo do Módulo': format_value_for_display(all_input_data.get('MODELO_MODULO_CALCULADO')),
        'Potência Pico dos Módulos (kW)': format_value_for_display(all_input_data.get('POTENCIA_PICO_MODULOS'), is_numeric=True),
        'Quantidade de Inversores': format_value_for_display(all_input_data.get('QUANTIDADE_INVERSOR_MANUAL'), is_numeric=True),
        'Potência do Inversor (kW)': format_value_for_display(all_input_data.get('POTENCIA_INVERSOR_MANUAL'), is_numeric=True),
        'Fabricante do Inversor': format_value_for_display(all_input_data.get('FABRICANTE_INVERSOR_MANUAL')),
        'Modelo do Inversor': format_value_for_display(all_input_data.get('MODELO_INVERSOR_CALCULADO')),
        'Potência Total do Sistema (kWp)': format_value_for_display(all_input_data.get('POTENCIA_TOTAL_SISTEMA_KWP'), is_numeric=True),
        'Isolação CA': format_value_for_display(all_input_data.get('ISOLACAO_CA')),
        'Cabo CA': format_value_for_display(all_input_data.get('CABO_CA')),
        'Disjuntor CA': format_value_for_display(all_input_data.get('DISJUNTOR_CA'), is_numeric=True),
    }

    return render_template('success.html',
                           message="Todos os documentos foram processados e estão prontos para download em um único arquivo ZIP!",
                           zip_filename=final_zip_filename_unique,
                           distributor_type=distributor_type,
                           dados_cliente=dados_cliente_for_display,
                           dados_entrada=dados_entrada_for_display,
                           dados_sistema_fv=dados_sistema_fv_for_display)


@app.route('/download_zip/<zip_filename>', methods=['GET'])
def download_zip(zip_filename):
    files_to_zip_info = session.pop('temp_files_to_zip', [])
    temp_zip_dir = session.pop('temp_zip_dir', None)
    nome_razao_social_zip_folder = session.pop('nome_razao_social_zip_folder', 'Projeto')
    # all_input_data_for_correction is explicitly kept for the 'Corrigir Dados' button functionality

    if not files_to_zip_info or not temp_zip_dir:
        print(f"Erro: Informações do ZIP não encontradas na sessão para '{zip_filename}'.")
        return "Erro: Arquivo ZIP não encontrado ou sessão expirada. Por favor, tente novamente.", 404

    zip_filepath = os.path.join(tempfile.gettempdir(), zip_filename)

    try:
        with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zf:
            for file_info in files_to_zip_info:
                original_path = file_info['path']
                zip_name = file_info['zip_filename']
                zf.write(original_path, arcname=f"{nome_razao_social_zip_folder}/{zip_name}")
        print(f"Arquivo ZIP criado temporariamente em: {zip_filepath}")

        response = send_file(zip_filepath,
                             mimetype='application/zip',
                             as_attachment=True,
                             download_name=f"Projeto de {nome_razao_social_zip_folder}.zip")

        @response.call_on_close
        def cleanup():
            try:
                if os.path.exists(zip_filepath):
                    os.remove(zip_filepath)
                    print(f"Arquivo ZIP temporário removido: {zip_filepath}")
                if temp_zip_dir and os.path.exists(temp_zip_dir):
                    shutil.rmtree(temp_zip_dir)
                    print(f"Diretório temporário e seus conteúdos removidos: {temp_zip_dir}")
            except Exception as e:
                print(f"Erro durante a limpeza de arquivos temporários: {e}")

        return response

    except Exception as e:
        print(f"Erro ao criar e servir o arquivo ZIP: {e}")
        return f"Erro ao gerar o arquivo ZIP: {e}", 500

# BLOCO MODIFICADO: Agora usa a variável de ambiente FLASK_DEBUG
if __name__ == '__main__':
    # Usa a variável de ambiente FLASK_DEBUG. Se não estiver definida, ou for diferente de '1', será False.
    debug_mode = os.environ.get('FLASK_DEBUG') == '1'
    app.run(host='0.0.0.0', port=5000, debug=debug_mode)