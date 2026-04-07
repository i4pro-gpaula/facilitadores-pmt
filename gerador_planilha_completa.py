"""
Gerador de Planilha Completa - PF e PJ
Preenche todas as 133 colunas da planilha com dados realistas
"""

import pandas as pd
import random
from datetime import datetime, timedelta
from faker import Faker
import os

fake = Faker('pt_BR')

class GeradorPlanilhaCompleta:
    def __init__(self):
        self.produtos = [1, 2, 3, 4]
        self.formas_pagamento = ['DEBITO', 'BOLETO', 'CARTAO']
        self.estados = ['SP', 'RJ', 'MG', 'PR', 'SC', 'RS', 'BA', 'PE', 'CE', 'GO', 'DF']
        
    def gerar_cpf(self):
        """Gera CPF válido"""
        cpf = [random.randint(0, 9) for _ in range(9)]
        
        # Primeiro dígito verificador
        soma = sum((10 - i) * cpf[i] for i in range(9))
        digito1 = (soma * 10 % 11) % 10
        cpf.append(digito1)
        
        # Segundo dígito verificador
        soma = sum((11 - i) * cpf[i] for i in range(10))
        digito2 = (soma * 10 % 11) % 10
        cpf.append(digito2)
        
        return ''.join(map(str, cpf))
    
    def gerar_cnpj(self):
        """Gera CNPJ válido (numérico tradicional)"""
        cnpj = [random.randint(0, 9) for _ in range(8)]
        cnpj.extend([0, 0, 0, 1])  # filial 0001
        
        # Primeiro dígito
        pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        soma = sum(cnpj[i] * pesos1[i] for i in range(12))
        digito1 = 11 - (soma % 11) if soma % 11 >= 2 else 0
        cnpj.append(digito1)
        
        # Segundo dígito
        pesos2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        soma = sum(cnpj[i] * pesos2[i] for i in range(13))
        digito2 = 11 - (soma % 11) if soma % 11 >= 2 else 0
        cnpj.append(digito2)
        
        return ''.join(map(str, cnpj))
    
    def gerar_cnpj_alfanumerico(self):
        """Gera CNPJ alfanumérico válido (14 caracteres)
        
        Formato SERPRO: 12 caracteres alfanuméricos + 2 dígitos verificadores
        Exemplo: 12ABC34501DE35
        
        Algoritmo oficial SERPRO:
        - Usa valor ASCII dos caracteres menos 48
        - Pesos cíclicos: 2,3,4,5,6,7,8,9 (repetindo)
        - Ambos DVs calculados com módulo 11
        """
        caracteres = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
        
        # Gera 12 primeiros caracteres alfanuméricos
        cnpj_base = ''.join(random.choice(caracteres) for _ in range(12))
        
        # Calcular primeiro dígito verificador
        dv1 = self._calcular_dv_serpro(cnpj_base)
        
        # Calcular segundo dígito verificador (incluindo o primeiro)
        dv2 = self._calcular_dv_serpro(cnpj_base + str(dv1))
        
        # Montar CNPJ completo
        return cnpj_base + str(dv1) + str(dv2)
    
    def _calcular_dv_serpro(self, cnpj_parcial):
        """Calcula dígito verificador usando algoritmo SERPRO
        
        Algoritmo:
        - Converte caracteres para ASCII - 48
        - Aplica pesos cíclicos 2-9 (reverso)
        - Módulo 11: se resto < 2, DV=0, senão DV=11-resto
        """
        from math import ceil
        
        # Gerar pesos cíclicos (2,3,4,5,6,7,8,9)
        tamanho = len(cnpj_parcial)
        num_ciclos = ceil(tamanho / 8)
        pesos = []
        for _ in range(num_ciclos):
            pesos.extend(range(2, 10))
        pesos = pesos[:tamanho]
        pesos.reverse()
        
        # Calcular soma ponderada usando ASCII - 48
        soma = 0
        for i, char in enumerate(cnpj_parcial):
            valor_ascii = ord(char.upper()) - 48
            soma += valor_ascii * pesos[i]
        
        # Calcular DV
        resto = soma % 11
        if resto < 2:
            return 0
        else:
            return 11 - resto
    
    def gerar_data_nascimento_pf(self):
        """Gera data de nascimento para pessoa física (18 a 80 anos)"""
        anos_atras = random.randint(18, 80)
        data = datetime.now() - timedelta(days=anos_atras*365 + random.randint(0, 364))
        return data.strftime('%d/%m/%Y')
    
    def gerar_data_nascimento_socio(self):
        """Gera data de nascimento para sócio (25 a 70 anos)"""
        anos_atras = random.randint(25, 70)
        data = datetime.now() - timedelta(days=anos_atras*365 + random.randint(0, 364))
        return data.strftime('%d/%m/%Y')
    
    def gerar_linha_pf(self, dt_proposta=None, vencimento=None, num_beneficiarios_fixo=None, qtd_capital=None, vl_capital=None,
                       usar_nome_social_seg=False, num_nome_social_ben=0,
                       cd_produto=None, dt_adesao=None, dt_cancelamento=None, dt_proposta_val=None, inicio_vig=None, fim_vig=None,
                       inicio_vig_endosso=None, fim_vig_endosso=None, forma_pagamento=None, nro_apolice=None, nro_sub=None,
                       tpmovto=None, vencimento_str=None,
                       nr_grupo=None, nr_cota=None, nm_bem=None, dt_inicio_consorcio=None, nr_meses_financiamento=None, vl_saldo_devedor=None):
        """Gera uma linha completa para Pessoa Física
        
        Args:
            dt_proposta: Data da proposta (compartilhada entre todos os registros)
            vencimento: Data de vencimento (compartilhada entre todos os registros)
            num_beneficiarios_fixo: Quantidade fixa de beneficiários (None = aleatório 0-5)
            cd_produto: Código do produto (None = em branco)
            dt_adesao: Data de adesão no formato DD/MM/YYYY (None = em branco)
            dt_cancelamento: Data de cancelamento no formato DD/MM/YYYY (None = em branco)
            dt_proposta_val: Data da proposta no formato DD/MM/YYYY (None = em branco)
            inicio_vig: Data de início de vigência no formato DD/MM/YYYY (None = em branco)
            fim_vig: Data de fim de vigência no formato DD/MM/YYYY (None = em branco)
            inicio_vig_endosso: Data de início de vigência do endosso DD/MM/YYYY (None = em branco)
            fim_vig_endosso: Data de fim de vigência do endosso DD/MM/YYYY (None = em branco)
            forma_pagamento: Código numérico da forma de pagamento (None = em branco)
            nro_apolice: Número da apólice (None = em branco)
            nro_sub: Número do subestipulante (None = em branco)
            tpmovto: Tipo de movimento – número ou sigla (None = usa 1)
            vencimento_str: Data de vencimento da parcela DD/MM/YYYY (None = calculado automaticamente)
        """
        nome = fake.name()
        cpf = self.gerar_cpf()
        sexo = random.choice(['M', 'F'])
        dt_nascimento = self.gerar_data_nascimento_pf()
        
        # Se não foram fornecidas datas compartilhadas, gera aleatoriamente
        if vencimento is None:
            _dt_base = datetime.now() - timedelta(days=random.randint(30, 365))
            _fim_calc = _dt_base + timedelta(days=365)
            vencimento = _fim_calc
        
        # Número de beneficiários
        if num_beneficiarios_fixo is not None:
            num_beneficiarios = num_beneficiarios_fixo
        else:
            num_beneficiarios = random.randint(0, 5)
        
        dados = {
            # Dados principais PF
            'NOME': nome,
            'NOME_SOCIAL': fake.name() if usar_nome_social_seg else None,
            'CD_PRODUTO': cd_produto,
            'DT_ADESAO': dt_adesao,
            'DT_CANCELAMENTO': dt_cancelamento,
            'DT_PROPOSTA': dt_proposta_val,
            'FIM_VIG': fim_vig,
            'FIM_VIG_ENDOSSO': fim_vig_endosso,
            'FORMA_PAGAMENTO': forma_pagamento,
            'INICIO_VIG': inicio_vig,
            'INICIO_VIG_ENDOSSO': inicio_vig_endosso,
            'NRO_APOLICE': nro_apolice,
            'NRO_SUB': nro_sub,
            'TPMOVTO': tpmovto,
            'VENCIMENTO': vencimento_str,
            'TIPO_PESSOA_SEGURADO': 'F',
            'CPF': cpf,
            'SEXO': sexo,
            'ENDEREÇO': fake.street_address(),
            'Cidade': fake.city(),
            'Estado': random.choice(self.estados),
            'CEP': fake.postcode().replace('-', ''),
            'DTNASC': dt_nascimento,
            'NR_MATRICULA': str(random.randint(100000, 999999)),
        }
        
        # Calcular participação por beneficiário (distribuição igual com ajuste no último)
        if num_beneficiarios > 0:
            base = round(100 / num_beneficiarios, 2)
            participacoes_ben = [base] * (num_beneficiarios - 1) + [round(100 - base * (num_beneficiarios - 1), 2)]
        else:
            participacoes_ben = []

        # Beneficiários (até 5)
        for i in range(1, 6):
            if i <= num_beneficiarios:
                dados[f'CPF_BENEFICIARIO_{i}'] = self.gerar_cpf()
                dados[f'BENEFICIARIO_{i}'] = fake.name()
                dados[f'DT_NASC_BENEFICIARIO_{i}'] = self.gerar_data_nascimento_pf()
                dados[f'PARTICIPACAO_BENEFICIARIO_{i}'] = participacoes_ben[i - 1]
                if num_nome_social_ben is None:
                    dados[f'NOME_SOCIAL_BENEFICIARIO_{i}'] = fake.name() if random.random() < 0.5 else None
                elif num_nome_social_ben > 0 and i <= num_nome_social_ben:
                    dados[f'NOME_SOCIAL_BENEFICIARIO_{i}'] = fake.name()
                else:
                    dados[f'NOME_SOCIAL_BENEFICIARIO_{i}'] = None
            else:
                dados[f'CPF_BENEFICIARIO_{i}'] = None
                dados[f'BENEFICIARIO_{i}'] = None
                dados[f'DT_NASC_BENEFICIARIO_{i}'] = None
                dados[f'PARTICIPACAO_BENEFICIARIO_{i}'] = None
                dados[f'NOME_SOCIAL_BENEFICIARIO_{i}'] = None

        # Campos adicionais
        for i in range(1, 21):
            if qtd_capital is not None and i <= qtd_capital:
                dados[f'CAPITAL_{i}'] = vl_capital if vl_capital is not None else random.randint(1000, 1000000)
            else:
                dados[f'CAPITAL_{i}'] = None
        dado_premio = None
        dados['PREMIO'] = dado_premio
        for i in range(1, 16):
            dados[f'NOME_SOCIO_{i}'] = None
            dados[f'NOME_SOCIAL_SOCIO_{i}'] = None
            dados[f'CNPJ_CPF_SOCIO_{i}'] = None
            dados[f'SEXO_SOCIO_{i}'] = None
            dados[f'DATA_NASCIMENTO_SOCIO_{i}'] = None
            dados[f'%_SOCIO_{i}'] = None

        # Consórcio
        dados['NR_GRUPO'] = nr_grupo
        dados['NR_COTA'] = nr_cota
        dados['NM_BEM'] = nm_bem
        dados['DT_INICIO_CONSORCIO'] = dt_inicio_consorcio
        dados['NR_MESES_FINANCIAMENTO'] = nr_meses_financiamento
        dados['VL_SALDO_DEVEDOR'] = vl_saldo_devedor
        
        return dados
    
    def gerar_linha_pj(self, usar_cnpj_alfanumerico=None, dt_proposta=None, vencimento=None, num_socios_fixo=None, qtd_capital=None, vl_capital=None,
                       num_nome_social_socio=0,
                       cd_produto=None, dt_adesao=None, dt_cancelamento=None, dt_proposta_val=None, inicio_vig=None, fim_vig=None,
                       inicio_vig_endosso=None, fim_vig_endosso=None, forma_pagamento=None, nro_apolice=None, nro_sub=None,
                       tpmovto=None, vencimento_str=None,
                       nr_grupo=None, nr_cota=None, nm_bem=None, dt_inicio_consorcio=None, nr_meses_financiamento=None, vl_saldo_devedor=None):
        """Gera uma linha completa para Pessoa Jurídica
        
        Args:
            usar_cnpj_alfanumerico: Se True, usa CNPJ alfanumérico; se False, usa numérico; se None, decide aleatoriamente (50/50)
            dt_proposta: Data da proposta (compartilhada entre todos os registros)
            vencimento: Data de vencimento (compartilhada entre todos os registros)
            num_socios_fixo: Quantidade fixa de sócios (None = aleatório 1-15, int = fixo)
            cd_produto: Código do produto (None = em branco)
            dt_adesao: Data de adesão no formato DD/MM/YYYY (None = em branco)
            dt_cancelamento: Data de cancelamento no formato DD/MM/YYYY (None = em branco)
            dt_proposta_val: Data da proposta no formato DD/MM/YYYY (None = em branco)
            inicio_vig: Data de início de vigência no formato DD/MM/YYYY (None = em branco)
            fim_vig: Data de fim de vigência no formato DD/MM/YYYY (None = em branco)
            inicio_vig_endosso: Data de início de vigência do endosso DD/MM/YYYY (None = em branco)
            fim_vig_endosso: Data de fim de vigência do endosso DD/MM/YYYY (None = em branco)
            forma_pagamento: Código numérico da forma de pagamento (None = em branco)
            nro_apolice: Número da apólice (None = em branco)
            nro_sub: Número do subestipulante (None = em branco)
            tpmovto: Tipo de movimento – número ou sigla (None = usa 1)
            vencimento_str: Data de vencimento da parcela DD/MM/YYYY (None = calculado automaticamente)
        """
        razao_social = fake.company()
        # Usar CNPJ alfanumérico ou numérico
        if usar_cnpj_alfanumerico is None:
            # Decisão aleatória (50% de chance)
            cnpj = self.gerar_cnpj_alfanumerico() if random.random() < 0.5 else self.gerar_cnpj()
        elif usar_cnpj_alfanumerico:
            cnpj = self.gerar_cnpj_alfanumerico()
        else:
            cnpj = self.gerar_cnpj()
        
        # Se não foram fornecidas datas compartilhadas, gera aleatoriamente
        if vencimento is None:
            _dt_base = datetime.now() - timedelta(days=random.randint(30, 365))
            _fim_calc = _dt_base + timedelta(days=365)
            vencimento = _fim_calc
        
        # Número de sócios
        if num_socios_fixo is not None:
            num_socios = num_socios_fixo
        else:
            num_socios = random.randint(1, 15)

        # Distribuição igual de participação entre sócios
        base = round(100 / num_socios, 2)
        participacoes = [base] * (num_socios - 1) + [round(100 - base * (num_socios - 1), 2)]
        
        dados = {
            # Dados principais PJ
            'NOME': razao_social,
            'NOME_SOCIAL': None,
            'CD_PRODUTO': cd_produto,
            'DT_ADESAO': dt_adesao,
            'DT_CANCELAMENTO': dt_cancelamento,
            'DT_PROPOSTA': dt_proposta_val,
            'FIM_VIG': fim_vig,
            'FIM_VIG_ENDOSSO': fim_vig_endosso,
            'FORMA_PAGAMENTO': forma_pagamento,
            'INICIO_VIG': inicio_vig,
            'INICIO_VIG_ENDOSSO': inicio_vig_endosso,
            'NRO_APOLICE': nro_apolice,
            'NRO_SUB': nro_sub,
            'TPMOVTO': tpmovto,
            'VENCIMENTO': vencimento_str,
            'TIPO_PESSOA_SEGURADO': 'J',
            'CPF': cnpj,
            'SEXO': None,
            'ENDEREÇO': fake.street_address(),
            'Cidade': fake.city(),
            'Estado': random.choice(self.estados),
            'CEP': fake.postcode().replace('-', ''),
            'DTNASC': None,
            'NR_MATRICULA': str(random.randint(100000, 999999)),
        }
        
        # Beneficiários (vazios para PJ)
        for i in range(1, 6):
            dados[f'CPF_BENEFICIARIO_{i}'] = None
            dados[f'BENEFICIARIO_{i}'] = None
            dados[f'DT_NASC_BENEFICIARIO_{i}'] = None
            dados[f'PARTICIPACAO_BENEFICIARIO_{i}'] = None
        for i in range(1, 6):
            dados[f'NOME_SOCIAL_BENEFICIARIO_{i}'] = None

        # Campos adicionais PJ
        for i in range(1, 21):
            if qtd_capital is not None and i <= qtd_capital:
                dados[f'CAPITAL_{i}'] = vl_capital if vl_capital is not None else random.randint(1000, 1000000)
            else:
                dados[f'CAPITAL_{i}'] = None
        dados['PREMIO'] = None  # Não preencher
        
        # Sócios - preencher conforme num_socios, demais vazios
        for i in range(1, 16):
            if i <= num_socios:
                dados[f'NOME_SOCIO_{i}'] = fake.name()
                if num_nome_social_socio is None:
                    dados[f'NOME_SOCIAL_SOCIO_{i}'] = fake.name() if random.random() < 0.5 else None
                elif num_nome_social_socio > 0 and i <= num_nome_social_socio:
                    dados[f'NOME_SOCIAL_SOCIO_{i}'] = fake.name()
                else:
                    dados[f'NOME_SOCIAL_SOCIO_{i}'] = None
                dados[f'CNPJ_CPF_SOCIO_{i}'] = self.gerar_cpf()
                dados[f'SEXO_SOCIO_{i}'] = random.choice(['M', 'F'])
                dados[f'DATA_NASCIMENTO_SOCIO_{i}'] = self.gerar_data_nascimento_socio()
                dados[f'%_SOCIO_{i}'] = participacoes[i - 1]
            else:
                dados[f'NOME_SOCIO_{i}'] = None
                dados[f'NOME_SOCIAL_SOCIO_{i}'] = None
                dados[f'CNPJ_CPF_SOCIO_{i}'] = None
                dados[f'SEXO_SOCIO_{i}'] = None
                dados[f'DATA_NASCIMENTO_SOCIO_{i}'] = None
                dados[f'%_SOCIO_{i}'] = None

        # Consórcio
        dados['NR_GRUPO'] = nr_grupo
        dados['NR_COTA'] = nr_cota
        dados['NM_BEM'] = nm_bem
        dados['DT_INICIO_CONSORCIO'] = dt_inicio_consorcio
        dados['NR_MESES_FINANCIAMENTO'] = nr_meses_financiamento
        dados['VL_SALDO_DEVEDOR'] = vl_saldo_devedor
        
        return dados
    
    def gerar_planilha(self, num_pf=10, num_pj=10, num_pj_cnpj_alfa=None, num_beneficiarios_pf=None, num_socios_pj=None, qtd_capital=None, vl_capital=None,
                       num_pf_nome_social=0, num_ben_nome_social=0, num_socio_nome_social=0,
                       cd_produto=None, dt_adesao=None, dt_cancelamento=None, dt_proposta_val=None, inicio_vig=None, fim_vig=None,
                       inicio_vig_endosso=None, fim_vig_endosso=None, forma_pagamento=None, nro_apolice=None, nro_sub=None,
                       tpmovto=None, vencimento_str=None,
                       nr_grupo=None, nr_cota=None, nm_bem=None, dt_inicio_consorcio=None, nr_meses_financiamento=None, vl_saldo_devedor=None,
                       layout='completo'):
        """
        Gera planilha completa com dados de PF e PJ
        
        Args:
            num_pf: Número de registros de Pessoa Física
            num_pj: Número de registros de Pessoa Jurídica
            num_pj_cnpj_alfa: Número de PJ que terão CNPJ alfanumérico (None = aleatório 50/50)
            num_beneficiarios_pf: Quantidade fixa de beneficiários por PF (None = aleatório 0-5, 0 = nenhum)
            num_socios_pj: Quantidade fixa de sócios por PJ (None = aleatório 1-15, int = fixo)
            cd_produto: Código do produto (None = em branco)
            dt_adesao: Data de adesão no formato DD/MM/YYYY (None = em branco)
            dt_cancelamento: Data de cancelamento no formato DD/MM/YYYY (None = em branco)
            dt_proposta_val: Data da proposta no formato DD/MM/YYYY (None = em branco)
            inicio_vig: Data de início de vigência no formato DD/MM/YYYY (None = em branco)
            fim_vig: Data de fim de vigência no formato DD/MM/YYYY (None = em branco)
            inicio_vig_endosso: Data de início de vigência do endosso DD/MM/YYYY (None = em branco)
            fim_vig_endosso: Data de fim de vigência do endosso DD/MM/YYYY (None = em branco)
            forma_pagamento: Código numérico da forma de pagamento (None = em branco)
            nro_apolice: Número da apólice (None = em branco)
            nro_sub: Número do subestipulante (None = em branco)
            tpmovto: Tipo de movimento – número ou sigla (None = usa 1)
            vencimento_str: Data de vencimento da parcela DD/MM/YYYY (None = calculado automaticamente)
        """
        dados_lista = []
        
        # Gerar datas compartilhadas para todo o arquivo
        dt_proposta_comum = datetime.now() - timedelta(days=random.randint(30, 180))
        vencimento_comum = dt_proposta_comum + timedelta(days=365)
        
        print(f"Datas comuns para o arquivo:")
        print(f"  Data da Proposta: {dt_proposta_comum.strftime('%d/%m/%Y')}")
        print(f"  Vencimento: {vencimento_comum.strftime('%d/%m/%Y')}")
        print()
        
        # Determinar quais PFs terão nome social do segurado
        if num_pf_nome_social is None:
            indices_nome_social_pf = None  # 50/50 por registro
        elif num_pf_nome_social > 0 and num_pf > 0:
            indices_nome_social_pf = set(random.sample(range(num_pf), min(num_pf_nome_social, num_pf)))
        else:
            indices_nome_social_pf = set()

        print(f"Gerando {num_pf} registros de Pessoa Física...")
        for i in range(num_pf):
            usar_ns_seg = (random.random() < 0.5) if indices_nome_social_pf is None else (i in indices_nome_social_pf)
            dados_lista.append(self.gerar_linha_pf(dt_proposta=dt_proposta_comum, vencimento=vencimento_comum, num_beneficiarios_fixo=num_beneficiarios_pf, qtd_capital=qtd_capital, vl_capital=vl_capital, usar_nome_social_seg=usar_ns_seg, num_nome_social_ben=num_ben_nome_social, cd_produto=cd_produto, dt_adesao=dt_adesao, dt_cancelamento=dt_cancelamento, dt_proposta_val=dt_proposta_val, inicio_vig=inicio_vig, fim_vig=fim_vig, inicio_vig_endosso=inicio_vig_endosso, fim_vig_endosso=fim_vig_endosso, forma_pagamento=forma_pagamento, nro_apolice=nro_apolice, nro_sub=nro_sub, tpmovto=tpmovto, vencimento_str=vencimento_str, nr_grupo=nr_grupo, nr_cota=nr_cota, nm_bem=nm_bem, dt_inicio_consorcio=dt_inicio_consorcio, nr_meses_financiamento=nr_meses_financiamento, vl_saldo_devedor=vl_saldo_devedor))
            if (i + 1) % 10 == 0:
                print(f"  {i + 1}/{num_pf} PF gerados")
        
        print(f"\nGerando {num_pj} registros de Pessoa Jurídica...")
        
        # Determinar quais PJs terão CNPJ alfanumérico
        if num_pj_cnpj_alfa is not None:
            # Criar lista de índices que terão CNPJ alfanumérico
            indices_alfa = random.sample(range(num_pj), min(num_pj_cnpj_alfa, num_pj))
            print(f"  {len(indices_alfa)} PJ com CNPJ alfanumérico")
            print(f"  {num_pj - len(indices_alfa)} PJ com CNPJ numérico")
        else:
            indices_alfa = None
        
        for i in range(num_pj):
            if indices_alfa is not None:
                usar_alfa = i in indices_alfa
                dados_lista.append(self.gerar_linha_pj(usar_cnpj_alfanumerico=usar_alfa, dt_proposta=dt_proposta_comum, vencimento=vencimento_comum, num_socios_fixo=num_socios_pj, qtd_capital=qtd_capital, vl_capital=vl_capital, num_nome_social_socio=num_socio_nome_social, cd_produto=cd_produto, dt_adesao=dt_adesao, dt_cancelamento=dt_cancelamento, dt_proposta_val=dt_proposta_val, inicio_vig=inicio_vig, fim_vig=fim_vig, inicio_vig_endosso=inicio_vig_endosso, fim_vig_endosso=fim_vig_endosso, forma_pagamento=forma_pagamento, nro_apolice=nro_apolice, nro_sub=nro_sub, tpmovto=tpmovto, vencimento_str=vencimento_str, nr_grupo=nr_grupo, nr_cota=nr_cota, nm_bem=nm_bem, dt_inicio_consorcio=dt_inicio_consorcio, nr_meses_financiamento=nr_meses_financiamento, vl_saldo_devedor=vl_saldo_devedor))
            else:
                dados_lista.append(self.gerar_linha_pj(dt_proposta=dt_proposta_comum, vencimento=vencimento_comum, num_socios_fixo=num_socios_pj, qtd_capital=qtd_capital, vl_capital=vl_capital, num_nome_social_socio=num_socio_nome_social, cd_produto=cd_produto, dt_adesao=dt_adesao, dt_cancelamento=dt_cancelamento, dt_proposta_val=dt_proposta_val, inicio_vig=inicio_vig, fim_vig=fim_vig, inicio_vig_endosso=inicio_vig_endosso, fim_vig_endosso=fim_vig_endosso, forma_pagamento=forma_pagamento, nro_apolice=nro_apolice, nro_sub=nro_sub, tpmovto=tpmovto, vencimento_str=vencimento_str, nr_grupo=nr_grupo, nr_cota=nr_cota, nm_bem=nm_bem, dt_inicio_consorcio=dt_inicio_consorcio, nr_meses_financiamento=nr_meses_financiamento, vl_saldo_devedor=vl_saldo_devedor))
            if (i + 1) % 10 == 0:
                print(f"  {i + 1}/{num_pj} PJ gerados")
        
        # Criar DataFrame
        df = pd.DataFrame(dados_lista)
        
        # Ordenar colunas conforme a planilha original
        colunas_ordenadas = [
            'NOME', 'NOME_SOCIAL', 'CD_PRODUTO', 'DT_ADESAO', 'DT_CANCELAMENTO', 'DT_PROPOSTA',
            'FIM_VIG', 'FIM_VIG_ENDOSSO', 'FORMA_PAGAMENTO', 'INICIO_VIG', 'INICIO_VIG_ENDOSSO',
            'NRO_APOLICE', 'NRO_SUB', 'TPMOVTO', 'VENCIMENTO', 'TIPO_PESSOA_SEGURADO',
            'CPF', 'SEXO', 'ENDEREÇO', 'Cidade', 'Estado', 'CEP', 'DTNASC', 'NR_MATRICULA'
        ]
        
        # Adicionar colunas de beneficiários
        for i in range(1, 6):
            colunas_ordenadas.append(f'CPF_BENEFICIARIO_{i}')
        for i in range(1, 6):
            colunas_ordenadas.append(f'BENEFICIARIO_{i}')
        for i in range(1, 6):
            colunas_ordenadas.append(f'NOME_SOCIAL_BENEFICIARIO_{i}')
        for i in range(1, 6):
            colunas_ordenadas.append(f'DT_NASC_BENEFICIARIO_{i}')
        for i in range(1, 6):
            colunas_ordenadas.append(f'PARTICIPACAO_BENEFICIARIO_{i}')
        
        # Adicionar campos adicionais
        colunas_ordenadas.extend([f'CAPITAL_{i}' for i in range(1, 21)] + ['PREMIO'])
        
        # Adicionar sócios
        for i in range(1, 16):
            colunas_ordenadas.extend([
                f'NOME_SOCIO_{i}', f'NOME_SOCIAL_SOCIO_{i}', f'CNPJ_CPF_SOCIO_{i}', f'SEXO_SOCIO_{i}',
                f'DATA_NASCIMENTO_SOCIO_{i}', f'%_SOCIO_{i}'
            ])

        # Colunas de consórcio
        colunas_ordenadas.extend([
            'NR_GRUPO', 'NR_COTA', 'NM_BEM',
            'NR_MESES_FINANCIAMENTO', 'DT_INICIO_CONSORCIO', 'VL_SALDO_DEVEDOR'
        ])

        df = df[colunas_ordenadas]

        # Layout flexível: remove colunas completamente vazias
        if layout == 'flexivel':
            df = df.dropna(axis=1, how='all')

        return df
    
    def salvar_planilha(self, df, nome_arquivo=None):
        """Salva a planilha gerada em Excel"""
        if nome_arquivo is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            nome_arquivo = f'planilha_completa_{timestamp}.xlsx'
        
        # Criar pasta de saída se não existir
        pasta_saida = 'planilhas_geradas'
        os.makedirs(pasta_saida, exist_ok=True)
        
        caminho_completo = os.path.join(pasta_saida, nome_arquivo)
        
        # Salvar com formatação
        with pd.ExcelWriter(caminho_completo, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Planilha1', index=False)
        
        print(f"\n✓ Planilha salva com sucesso!")
        print(f"  Arquivo: {caminho_completo}")
        print(f"  Registros: {len(df)}")
        print(f"  Colunas: {len(df.columns)}")
        
        return caminho_completo


def main():
    """Função principal"""
    print("=" * 80)
    print("GERADOR DE PLANILHA COMPLETA - PF e PJ")
    print("=" * 80)
    print()
    
    gerador = GeradorPlanilhaCompleta()
    
    # Perguntar quantos registros gerar
    try:
        print("Quantos registros deseja gerar?")
        num_pf = int(input("  Pessoa Física (PF): ") or "10")
        num_pj = int(input("  Pessoa Jurídica (PJ): ") or "10")
        
        if num_pj > 0:
            print()
            print("Quantas Pessoas Jurídicas terão CNPJ alfanumérico?")
            print(f"  (0 a {num_pj}, deixe em branco para 50% aleatório)")
            entrada_alfa = input(f"  CNPJ alfanumérico: ").strip()
            
            if entrada_alfa == "":
                num_pj_cnpj_alfa = None
                print("  → Usando distribuição aleatória (50/50)")
            else:
                num_pj_cnpj_alfa = int(entrada_alfa)
                if num_pj_cnpj_alfa > num_pj:
                    print(f"  ⚠ Valor ajustado para {num_pj} (máximo possível)")
                    num_pj_cnpj_alfa = num_pj
                elif num_pj_cnpj_alfa < 0:
                    print("  ⚠ Valor ajustado para 0 (mínimo possível)")
                    num_pj_cnpj_alfa = 0
        else:
            num_pj_cnpj_alfa = None
            
    except ValueError:
        print("Entrada inválida. Usando valores padrão: 10 PF e 10 PJ (50% CNPJ alfanumérico)")
        num_pf = 10
        num_pj = 10
        num_pj_cnpj_alfa = None
    
    print()
    print("=" * 80)
    
    # Gerar planilha
    df = gerador.gerar_planilha(num_pf, num_pj, num_pj_cnpj_alfa)
    
    # Salvar
    caminho = gerador.salvar_planilha(df)
    
    print()
    print("=" * 80)
    print("RESUMO DA PLANILHA GERADA")
    print("=" * 80)
    print(f"Total de Pessoa Física: {len(df[df['TIPO_PESSOA_SEGURADO'] == 'F'])}")
    print(f"Total de Pessoa Jurídica: {len(df[df['TIPO_PESSOA_SEGURADO'] == 'J'])}")
    print(f"Total de registros: {len(df)}")
    print()
    
    # Mostrar amostra
    print("Primeiras 3 linhas (colunas principais):")
    colunas_amostra = ['NOME', 'TIPO_PESSOA_SEGURADO', 'CPF', 'Cidade', 'Estado', 'CAPITAL_1', 'PREMIO']
    print(df[colunas_amostra].head(3).to_string(index=False))
    

if __name__ == '__main__':
    main()
