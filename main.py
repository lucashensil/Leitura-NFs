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

    lista_produtos = []
    for produto in dic_produtos:
        valor_produto = produto['prod']['vProd']
        nome_produto = produto['prod']['xProd']
        lista_produtos.append((nome_produto, valor_produto))


    inf_valor = dic_nfe['total']['ICMSTot']['vNF']

    inf_cnpj_vendedor = dic_nfe['emit']['CNPJ']
    inf_nome_vendedor = dic_nfe['emit']['xNome']

    inf_cpf_comprador = dic_dest['CPF']
    inf_nome_comprador = dic_dest['xNome']
    inf_rua = dic_dest['enderDest']['xLgr']
    inf_bairro = dic_dest['enderDest']['xBairro']
    info_endereco = inf_rua + ', ' + inf_bairro
    info_municipio = dic_dest['enderDest']['xMun']
    info_uf = dic_dest['enderDest']['UF']
    info_fone = dic_dest['enderDest']['fone']
    info_email = dic_dest['email']


    dic_respostas = {
        'valor_total': [inf_valor],
        'cnpj_vendedor': [inf_cnpj_vendedor],
        'nome_vendedor': [inf_nome_vendedor],

        'endereo_comprador': [info_endereco],
        'municipio_comprador': [info_municipio],
        'estado_comprador': [info_uf],
        'cpf_comprador': [inf_cpf_comprador],
        'nome_comprador': [inf_nome_comprador],
        'fone_comprador': [info_fone],
        'email_comprador': [info_email],
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

    inf_valor = dic_nfe['Servico']['Valores']['ValorServicos']
    inf_servicos = dic_nfe['Servico']['Discriminacao']

    inf_cnpj_vendedor = dic_nfe['PrestadorServico']['IdentificacaoPrestador']['Cnpj']
    inf_nome_vendedor = dic_nfe['PrestadorServico']['NomeFantasia']

    try:
        inf_cpfCnpj_comprador = dic_nfe['TomadorServico']['IdentificacaoTomador']['CpfCnpj']['Cnpj']
    except:
        inf_cpfCnpj_comprador = dic_nfe['TomadorServico']['IdentificacaoTomador']['CpfCnpj']['Cpf']
    inf_nome_comprador = dic_nfe['TomadorServico']['RazaoSocial']


    dic_respostas = {
        'valor_total': [inf_valor],
        'cnpj_vendedor': [inf_cnpj_vendedor],
        'nome_vendedor': [inf_nome_vendedor],
        'CpfCnpj_comprador': [inf_cpfCnpj_comprador],
        'nome_comprador': [inf_nome_comprador],
        'servicos': [inf_servicos]
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

