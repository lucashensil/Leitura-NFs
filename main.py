import xmltodict
import pandas as pd


def ler_xml_DANFE(NF):
    """Funcao de leitura de DANFE

    Esta funcao realiza leitura de uma nota fiscal modelo DANFE

    Args:
        NF (string): Nome do arquivo 

    Returns:
        dict: dicionario contendo todos as informacoes retiradas
    """
    with open(f'Notas Fiscais\{NF}', 'rb') as arquivo:
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
        'valor_total_prod': [info_valor_prod],
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
    with open(f'Notas Fiscais\{NF}', 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)

    dic_nfe = documento['ConsultarNfseResposta']['ListaNfse']['CompNfse']['Nfse']['InfNfse']

    info_valor = dic_nfe['Servico']['Valores']['ValorServicos']
    info_servicos = dic_nfe['Servico']['Discriminacao']

    info_cnpj_vendedor = dic_nfe['PrestadorServico']['IdentificacaoPrestador']['Cnpj']
    info_nome_vendedor = dic_nfe['PrestadorServico']['NomeFantasia']

    try:
        info_cpfCnpj_comprador = dic_nfe['TomadorServico']['IdentificacaoTomador']['CpfCnpj']['Cnpj']
    except:
        info_cpfCnpj_comprador = dic_nfe['TomadorServico']['IdentificacaoTomador']['CpfCnpj']['Cpf']
    info_nome_comprador = dic_nfe['TomadorServico']['RazaoSocial']


    dic_respostas = {
        'valor_total': [info_valor],
        'cnpj_vendedor': [info_cnpj_vendedor],
        'nome_vendedor': [info_nome_vendedor],
        'CpfCnpj_comprador': [info_cpfCnpj_comprador],
        'nome_comprador': [info_nome_comprador],
        'servicos': [info_servicos]
    }
    return dic_respostas

def planilhar(NF):
    if 'DANFE' in NF:
        resposta = ler_xml_DANFE(NF)
        print('NF DANFE')
    else:
        resposta = ler_xml_servico(NF)
    
    tabela = pd.DataFrame.from_dict(resposta)
    tabela.to_excel('NF.xlsm')

