import xmltodict
import pandas as pd


def ler_xml_DANFE(NF):
    with open(f'Notas Fiscais\{NF}', 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)

    dic_nfe = documento['nfeProc']['NFe']['infNFe']
    dic_produtos = dic_nfe['det']

    lista_produtos = []
    for produto in dic_produtos:
        valor_produto = produto['prod']['vProd']
        nome_produto = produto['prod']['xProd']
        lista_produtos.append((nome_produto, valor_produto))


    inf_valor = dic_nfe['total']['ICMSTot']['vNF']

    inf_cnpj_vendedor = dic_nfe['emit']['CNPJ']
    inf_nome_vendedor = dic_nfe['emit']['xNome']

    inf_cpf_comprador = dic_nfe['dest']['CPF']
    inf_nome_comprador = dic_nfe['dest']['xNome']


    dic_respostas = {
        'valor_total': [inf_valor],
        'cnpj_vendedor': [inf_cnpj_vendedor],
        'nome_vendedor': [inf_nome_vendedor],
        'cpf_comprador': [inf_cpf_comprador],
        'nome_vendedor': [inf_nome_comprador],
        'produtos': [lista_produtos]
    }

    return dic_respostas
def ler_xml_servico(NF):
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

def planilhar(Nota):
    if 'DANFE' in Nota:
        resposta = ler_xml_DANFE(Nota)
        print('Nota DANFE')
    else:
        resposta = ler_xml_servico(Nota)
    
    tabela = pd.DataFrame.from_dict(resposta)
    tabela.to_excel('NF.xlsm')

