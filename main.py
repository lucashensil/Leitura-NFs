import xmltodict
import pandas as pd
import os

def ler_xml_DANFE(NF):
    """Funcao de leitura de DANFE

    Esta funcao realiza leitura de uma nota fiscal modelo DANFE

    Args:
        NF (string): Nome do arquivo 

    Returns:
        dict: dicionario contendo todos as informacoes retiradas
    """
    with open(f'NFs Exemplo\{NF}', 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)

    dic_nfe = documento['nfeProc']['NFe']['infNFe']
    dic_produtos = dic_nfe['det']
    dic_dest = dic_nfe['dest']
    dic_emit = dic_nfe['emit']
    dic_transportador = dic_nfe['transp']['transporta']
    dic_valor = dic_nfe['total']['ICMSTot']

    lista_produtos = []
    for produto in dic_produtos:
        valor_produto = produto['prod']['vProd']
        nome_produto = produto['prod']['xProd']
        lista_produtos.append((nome_produto, valor_produto))


    info_natureza = dic_nfe['ide']['natOp']
    info_emissao_data = dic_nfe['ide']['dhEmi']
    info_data_vencimento = dic_nfe['cobr']['dup']['dVenc']
    info_valor_nota = dic_valor['vNF']
    info_valor_prod = dic_valor['vProd']
    info_valor_frete = dic_valor['vFrete']
    info_valor_desc = dic_valor['vDesc']
    info_valor_totTrib = dic_valor['vTotTrib']

    info_cnpj_vendedor = dic_emit['CNPJ']
    info_nome_vendedor = dic_emit['xNome']
    info_rua_vend =  dic_emit['enderEmit']['xLgr']
    info_bairro_vend = dic_emit['enderEmit']['xBairro']
    info_endereco_vend = info_rua_vend + ', ' + info_bairro_vend
    info_municipio_vend = dic_emit['enderEmit']['xMun']
    info_uf_vend = dic_emit['enderEmit']['UF']
    info_fone_vend = dic_emit['enderEmit']['fone']

    info_cnpj_transp = dic_transportador['CNPJ']
    info_nome_transp = dic_transportador['xNome']
    info_endereco_transp = dic_transportador['xEnder']
    info_municipio_transp = dic_transportador['xMun']
    info_uf_transp = dic_transportador['UF']

    info_cpf_comprador = dic_dest['CPF']
    info_nome_comprador = dic_dest['xNome']
    info_rua = dic_dest['enderDest']['xLgr']
    info_bairro = dic_dest['enderDest']['xBairro']
    info_endereco_compr = info_rua + ', ' + info_bairro
    info_municipio_compr = dic_dest['enderDest']['xMun']
    info_uf_compr = dic_dest['enderDest']['UF']
    info_fone_compr = dic_dest['enderDest']['fone']
    info_email_compr = dic_dest['email']


    dic_respostas = {
        'natureza_operacao': [info_natureza],
        'data_emissao': [info_emissao_data],
        'data_vencimento': [info_data_vencimento],
        'valor_total_prod': info_valor_prod,
        'valor_total_nota': [info_valor_nota],
        'valor_desconto': [info_valor_desc],
        'valor_frete': [info_valor_frete],
        'valor_total_tributos': [info_valor_totTrib],
        
        'cnpj_vendedor': [info_cnpj_vendedor],
        'nome_vendedor': [info_nome_vendedor],
        'endereco_vendedor': [info_endereco_vend],
        'municipio_vendedor': [info_municipio_vend],
        'estado_vendedor': [info_uf_vend],
        'fone_comprador': [info_fone_vend],

        'cnpj_transportador': [info_cnpj_transp],
        'nome_transportador': [info_nome_transp],
        'endereco_transportador': [info_endereco_transp],
        'municipio_transportador': [info_municipio_transp],
        'estado_transportador': [info_uf_transp],

        'endereo_comprador': [info_endereco_compr],
        'municipio_comprador': [info_municipio_compr],
        'estado_comprador': [info_uf_compr],
        'cpf_comprador': [info_cpf_comprador],
        'nome_comprador': [info_nome_comprador],
        'fone_comprador': [info_fone_compr],
        'email_comprador': [info_email_compr],

        'produtos': [lista_produtos]
    }

    return dic_respostas
def ler_xml_servico(NF):
    """Funcao para ler notas de servico

    Args:
        NF (string): Nome do arquivo 

    Returns:
        dict: dicionario contendo todos as informacoes retiradas
    """
    with open(f'NFs Exemplo\{NF}', 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)

    dic_nfe = documento['ConsultarNfseResposta']['ListaNfse']['CompNfse']['Nfse']['InfNfse']
    dic_prestador = dic_nfe['PrestadorServico']
    dic_tomador = dic_nfe['TomadorServico']

    info_numero = dic_nfe['Numero']
    info_codigo_verificacao = dic_nfe['CodigoVerificacao']
    info_data_emissao = dic_nfe['DataEmissao']


    info_valor = dic_nfe['Servico']['Valores']['ValorServicos']
    info_servicos = dic_nfe['Servico']['Discriminacao']

    info_cnpj_vendedor = dic_prestador['IdentificacaoPrestador']['Cnpj']
    info_inscricao_municipal = dic_prestador['IdentificacaoPrestador']['InscricaoMunicipal']
    info_nome_vendedor = dic_prestador['NomeFantasia']
    info_razao_social = dic_prestador['RazaoSocial']
    info_endereco_vendedor = dic_prestador['Endereco']['Endereco'] + ' - ' + dic_prestador['Endereco']['Bairro'] + ', ' + dic_prestador['Endereco']['Uf']
    info_cep_vendedor = dic_prestador['Endereco']['Cep']    
    info_tel_vendedor = dic_prestador['Contato']['Telefone']
    info_email_vendedor = dic_prestador['Contato']['Email']

    info_cod_municipio = dic_prestador['Endereco']['CodigoMunicipio']


    try:
        info_cpfCnpj_comprador = dic_tomador['IdentificacaoTomador']['CpfCnpj']['Cnpj']
    except:
        info_cpfCnpj_comprador = dic_tomador['IdentificacaoTomador']['CpfCnpj']['Cpf']
    info_nome_comprador = dic_tomador['RazaoSocial']
    info_endereco_tomador = dic_tomador['Endereco']['Endereco'] + ' - ' + dic_tomador['Endereco']['Bairro'] + ', ' + dic_tomador['Endereco']['Uf']
    info_cep_tomador = dic_tomador['Contato']['Telefone']
    info_tel_tomador = dic_tomador['Contato']['Telefone']
    info_email_tomador = dic_tomador['Contato']['Email']


    dic_respostas = {
        'numero_nota': [info_numero],
        'codigo_verificacao': [info_codigo_verificacao],
        'codigo_municipal': [info_cod_municipio],
        'data_emissao': [info_data_emissao],
        'valor_total': [info_valor],

        'cnpj_vendedor': [info_cnpj_vendedor],
        "inscricao_municiapl": [info_inscricao_municipal],
        'nome_vendedor': [info_nome_vendedor],
        'razao_social_vendedor': [info_razao_social],
        'endereco_vendedor': [info_endereco_vendedor],
        'cep_vendedor': [info_cep_vendedor],
        'tel_vendedor': [info_tel_vendedor],
        'email_vendedor': [info_email_vendedor],

        'CpfCnpj_comprador': [info_cpfCnpj_comprador],
        'nome_comprador': [info_nome_comprador],
        'servicos': [info_servicos],
        'endereco_tomador': [info_endereco_tomador],
        'cep_tomador': [info_cep_tomador],
        'tel_tomador': [info_tel_tomador],
        'email_tomador': [info_email_tomador]
    }
    return dic_respostas

def identificar_nf(NF):
    with open(f'NFs Exemplo\{NF}', 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
    
    if 'nfeProc' in documento:
        return 'DANFE'
    elif 'ConsultarNfseResposta' in documento:
        return 'Xml Servico'
    else:
        return 'Modelo de NF não suportado'   
def planilhar_unico(NF):
    modelo = identificar_nf(NF)
    
    if modelo == 'DANFE':
        resposta = ler_xml_DANFE(NF)
    elif modelo == 'Xml Servico':
        resposta = ler_xml_servico(NF)

    
    tabela = pd.DataFrame.from_dict(resposta)
    tabela.to_excel(f'NF_{NF}.xlsm')
def planilhar_todos():
    notas = os.listdir('NFs Exemplo')
    
    for nota in notas:
        if '.xml' in nota:
            planilhar_unico(nota)



