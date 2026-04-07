"""
Aplicação Web - Gerador de Lote de Importação
Flask app para PythonAnywhere
"""

import io
import os
import sys
from datetime import datetime

import pandas as pd
from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, url_for

# Garante que o módulo gerador seja encontrado na mesma pasta
sys.path.insert(0, os.path.dirname(__file__))
from gerador_planilha_completa import GeradorPlanilhaCompleta

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'gerador_lote_importacao_2026')


def _erro(msg, is_fetch):
    """Retorna erro como JSON (fetch) ou flash+redirect (form clássico)."""
    if is_fetch:
        return jsonify({'error': msg}), 400
    flash(msg, 'error')
    return redirect(url_for('index'))


@app.route('/')
def home():
    return render_template('facilitadoresPMT.html')


@app.route('/gerador', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        is_fetch = request.headers.get('X-Fetch') == '1'
        # Validar e capturar entradas
        try:
            num_pf_str = request.form.get('num_pf', '').strip()
            num_pj_str = request.form.get('num_pj', '').strip()
            num_pj_alfa_str = request.form.get('num_pj_alfa', '').strip()
            num_pf = int(num_pf_str) if num_pf_str else 0
            num_pj = int(num_pj_str) if num_pj_str else 0

            if num_pf < 0 or num_pj < 0:
                return _erro('Quantidades não podem ser negativas.', is_fetch)

            if num_pf == 0 and num_pj == 0:
                return _erro('Informe ao menos 1 registro de PF ou PJ.', is_fetch)

            if num_pf + num_pj > 1000:
                return _erro('Limite máximo de 1000 registros por geração.', is_fetch)

            # CNPJ alfanumérico
            usar_cnpj_alfa = request.form.get('usar_cnpj_alfa') == 'on'
            num_pj_alfa = 0
            if usar_cnpj_alfa:
                if num_pj_alfa_str == '':
                    num_pj_alfa = None
                else:
                    num_pj_alfa = int(num_pj_alfa_str)
                    if num_pj_alfa < 1 or num_pj_alfa > num_pj:
                        return _erro(f'CNPJ alfanumérico deve ser entre 1 e {num_pj}.', is_fetch)

            # Beneficiários PF
            usar_beneficiario = request.form.get('usar_beneficiario') == 'on'
            num_beneficiarios_pf = None
            if usar_beneficiario:
                num_ben_str = request.form.get('num_beneficiarios_pf', '').strip()
                if num_ben_str:
                    num_beneficiarios_pf = int(num_ben_str)
                    if num_beneficiarios_pf < 1 or num_beneficiarios_pf > 5:
                        return _erro('Quantidade de beneficiários deve ser entre 1 e 5.', is_fetch)
            else:
                num_beneficiarios_pf = 0

            # Nome social segurado PF
            usar_nome_social_pf = request.form.get('usar_nome_social_pf') == 'on'
            num_pf_nome_social = 0
            if usar_nome_social_pf:
                ns_pf_str = request.form.get('num_nome_social_pf', '').strip()
                if ns_pf_str == '':
                    num_pf_nome_social = None
                else:
                    num_pf_nome_social = int(ns_pf_str)
                    if num_pf_nome_social < 1 or num_pf_nome_social > num_pf:
                        return _erro(f'Nome social do segurado PF deve ser entre 1 e {num_pf}.', is_fetch)

            # Nome social beneficiário PF
            usar_nome_social_ben = request.form.get('usar_nome_social_ben') == 'on'
            num_ben_nome_social = 0
            if usar_nome_social_ben:
                ns_ben_str = request.form.get('num_nome_social_ben', '').strip()
                if ns_ben_str == '':
                    num_ben_nome_social = None
                else:
                    num_ben_nome_social = int(ns_ben_str)
                    if num_ben_nome_social < 1 or num_ben_nome_social > 5:
                        return _erro('Nome social do beneficiário deve ser entre 1 e 5.', is_fetch)

            # Nome social sócio PJ
            usar_nome_social_socio = request.form.get('usar_nome_social_socio') == 'on'
            num_socio_nome_social = 0
            if usar_nome_social_socio:
                ns_socio_str = request.form.get('num_nome_social_socio', '').strip()
                if ns_socio_str == '':
                    num_socio_nome_social = None
                else:
                    num_socio_nome_social = int(ns_socio_str)
                    if num_socio_nome_social < 1 or num_socio_nome_social > 15:
                        return _erro('Nome social do sócio deve ser entre 1 e 15.', is_fetch)

            # Sócios PJ
            usar_socio = request.form.get('usar_socio') == 'on'
            num_socios_pj = None
            if usar_socio:
                num_soc_str = request.form.get('num_socios_pj', '').strip()
                if num_soc_str:
                    num_socios_pj = int(num_soc_str)
                    if num_socios_pj < 1 or num_socios_pj > 15:
                        return _erro('Quantidade de sócios deve ser entre 1 e 15.', is_fetch)
            else:
                num_socios_pj = 1

            # Capital (obrigatório)
            qtd_str = request.form.get('qtd_capital', '').strip()
            if not qtd_str:
                return _erro('Informe a quantidade de campos de capital a preencher (1 a 20).', is_fetch)
            qtd_capital = int(qtd_str)
            if qtd_capital < 1 or qtd_capital > 20:
                return _erro('A quantidade de campos de capital deve ser entre 1 e 20.', is_fetch)
            vl_str = request.form.get('vl_capital', '').strip()
            vl_capital = int(vl_str) if vl_str else None

            # Informações Adicionais
            cd_produto_str = request.form.get('cd_produto', '').strip()
            cd_produto = int(cd_produto_str) if cd_produto_str else None

            def _conv_data(s):
                if not s:
                    return None
                try:
                    return datetime.strptime(s, '%Y-%m-%d').strftime('%d/%m/%Y')
                except ValueError:
                    return None

            dt_adesao = _conv_data(request.form.get('dt_adesao', ''))
            dt_cancelamento = _conv_data(request.form.get('dt_cancelamento', ''))
            dt_proposta_val = _conv_data(request.form.get('dt_proposta', ''))
            inicio_vig = _conv_data(request.form.get('inicio_vig', ''))
            fim_vig = _conv_data(request.form.get('fim_vig', ''))

            inicio_vig_endosso = _conv_data(request.form.get('inicio_vig_endosso', ''))
            fim_vig_endosso = _conv_data(request.form.get('fim_vig_endosso', ''))
            vencimento_str = _conv_data(request.form.get('vencimento', ''))

            forma_pag_str = request.form.get('forma_pagamento', '').strip()
            forma_pagamento = int(forma_pag_str) if forma_pag_str else None

            nro_apolice = request.form.get('nro_apolice', '').strip() or None
            nro_sub = request.form.get('nro_sub', '').strip() or None

            tpmovto_raw = request.form.get('tpmovto', '').strip()
            if tpmovto_raw:
                tpmovto = int(tpmovto_raw) if tpmovto_raw.isdigit() else tpmovto_raw
            else:
                tpmovto = None

            # Informações de Consórcio
            nr_grupo_str = request.form.get('nr_grupo', '').strip()
            nr_grupo = int(nr_grupo_str) if nr_grupo_str else None

            nr_cota_str = request.form.get('nr_cota', '').strip()
            nr_cota = int(nr_cota_str) if nr_cota_str else None

            nm_bem = request.form.get('nm_bem', '').strip() or None

            dt_inicio_consorcio = _conv_data(request.form.get('dt_inicio_consorcio', ''))

            nr_meses_str = request.form.get('nr_meses_financiamento', '').strip()
            nr_meses_financiamento = int(nr_meses_str) if nr_meses_str else None

            vl_saldo_str = request.form.get('vl_saldo_devedor', '').strip()
            vl_saldo_devedor = float(vl_saldo_str) if vl_saldo_str else None

            # Layout da planilha
            layout_tipo = 'completo' if request.form.get('layout_completo') == 'on' else 'flexivel'

        except ValueError:
            return _erro('Valores inválidos. Use apenas números inteiros.', is_fetch)

        # Gerar planilha em memória
        gerador = GeradorPlanilhaCompleta()
        df = gerador.gerar_planilha(num_pf, num_pj, num_pj_alfa, num_beneficiarios_pf=num_beneficiarios_pf, num_socios_pj=num_socios_pj, qtd_capital=qtd_capital, vl_capital=vl_capital, num_pf_nome_social=num_pf_nome_social, num_ben_nome_social=num_ben_nome_social, num_socio_nome_social=num_socio_nome_social, cd_produto=cd_produto, dt_adesao=dt_adesao, dt_cancelamento=dt_cancelamento, dt_proposta_val=dt_proposta_val, inicio_vig=inicio_vig, fim_vig=fim_vig, inicio_vig_endosso=inicio_vig_endosso, fim_vig_endosso=fim_vig_endosso, forma_pagamento=forma_pagamento, nro_apolice=nro_apolice, nro_sub=nro_sub, tpmovto=tpmovto, vencimento_str=vencimento_str, nr_grupo=nr_grupo, nr_cota=nr_cota, nm_bem=nm_bem, dt_inicio_consorcio=dt_inicio_consorcio, nr_meses_financiamento=nr_meses_financiamento, vl_saldo_devedor=vl_saldo_devedor, layout=layout_tipo)

        # Colunas que devem ser gravadas como texto explícito no xlsx.
        # Inclui datas (DD/MM/YYYY como string) e identificadores com zeros à esquerda.
        # O formato '@' é aplicado SOMENTE em células com valor para não gerar
        # células vazias com estilo que alguns sistemas interpretam como dado.
        COLUNAS_TEXTO = {
            'CPF', 'CEP', 'NR_MATRICULA', 'NRO_APOLICE', 'NRO_SUB',
            'NR_GRUPO', 'NR_COTA',
            'DTNASC',
            'DT_ADESAO', 'DT_CANCELAMENTO', 'DT_PROPOSTA',
            'INICIO_VIG', 'FIM_VIG', 'INICIO_VIG_ENDOSSO', 'FIM_VIG_ENDOSSO',
            'VENCIMENTO', 'DT_INICIO_CONSORCIO',
        }
        for i in range(1, 6):
            COLUNAS_TEXTO.add(f'CPF_BENEFICIARIO_{i}')
            COLUNAS_TEXTO.add(f'DT_NASC_BENEFICIARIO_{i}')
        for i in range(1, 16):
            COLUNAS_TEXTO.add(f'CNPJ_CPF_SOCIO_{i}')
            COLUNAS_TEXTO.add(f'DATA_NASCIMENTO_SOCIO_{i}')

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Planilha1', index=False)
            ws = writer.sheets['Planilha1']
            col_indices = {
                cell.value: cell.column
                for cell in ws[1]
                if cell.value in COLUNAS_TEXTO
            }
            for col_idx in col_indices.values():
                for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        if cell.value is not None:
                            cell.number_format = '@'
        output.seek(0)

        prefixo_raw = request.form.get('prefixo_arquivo_form', '').strip()
        # Manter apenas caracteres seguros para nome de arquivo
        prefixo = ''.join(c for c in prefixo_raw if c.isalnum() or c in '-_ ').strip().replace(' ', '_')
        if not prefixo:
            prefixo = 'lote_importacao'
        filename = f'{prefixo}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

        return send_file(
            output,
            download_name=filename,
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    return render_template('geradorLoteColetivoPFPJ.html')


if __name__ == '__main__':
    app.run(debug=True)
