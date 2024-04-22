import xmltodict
import pandas as pd
import os

def ler_xml_NFe(NF, caminho='Notas fiscais'):
    """Realiza a leitura de uma nota fiscal eletrônica (NFe).

    Esta função lê o arquivo XML de uma nota fiscal no formato NFe e extrai suas informações.

    Args:
        NF (str): Nome do arquivo da nota fiscal no formato NFe.
        caminho (str, opcional): Caminho para o diretório onde o arquivo está localizado. O padrão é 'Notas fiscais'.

    Returns:
        dict: Um dicionário contendo todas as informações extraídas da nota fiscal.
    """
    caminho_completo = os.path.join(caminho, NF)

    # Carregar o arquivo XML
    with open(caminho_completo, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)

    # Acesso aos dados da nota
    nfe = documento['nfeProc']['NFe']['infNFe']
    emitente = nfe['emit']
    destinatario = nfe['dest']
    transporte = nfe['transp']['transporta']
    total = nfe['total']['ICMSTot']
    produtos = [(p['prod']['xProd'], p['prod']['vProd']) for p in nfe['det']]

    # dicionário contendo as respostas
    dic_respostas = {
        'natureza_operacao': nfe['ide']['natOp'],
        'data_emissao': nfe['ide']['dhEmi'],
        'data_vencimento': nfe['cobr']['dup']['dVenc'],
        'valor_total_prod': total['vProd'],
        'valor_total_nota': total['vNF'],
        'valor_desconto': total['vDesc'],
        'valor_frete': total['vFrete'],
        'valor_total_tributos': total['vTotTrib'],
        'cnpj_vendedor': emitente['CNPJ'],
        'nome_vendedor': emitente['xNome'],
        'endereco_vendedor': emitente['enderEmit']['xLgr'] + ', ' + emitente['enderEmit']['xBairro'],
        'municipio_vendedor': emitente['enderEmit']['xMun'],
        'estado_vendedor': emitente['enderEmit']['UF'],
        'fone_vendedor': emitente['enderEmit']['fone'],
        'cnpj_transportador': transporte['CNPJ'],
        'nome_transportador': transporte['xNome'],
        'endereco_transportador': transporte['xEnder'],
        'municipio_transportador': transporte['xMun'],
        'estado_transportador': transporte['UF'],
        'endereco_destinatario': destinatario['enderDest']['xLgr'] + ', ' + destinatario['enderDest']['xBairro'],
        'municipio_destinatario': destinatario['enderDest']['xMun'],
        'estado_destinatario': destinatario['enderDest']['UF'],
        'cpf_comprador': destinatario['CPF'],
        'nome_comprador': destinatario['xNome'],
        'fone_comprador': destinatario['enderDest']['fone'],
        'email_comprador': destinatario['email'],
        'produtos': produtos
    }

    return dic_respostas
    
def ler_xml_servico(NF, caminho='Notas fiscais'):
    """Realiza a leitura de notas fiscais de serviço.

    Esta função lê o arquivo XML de uma nota fiscal de serviço eletronico e extrai suas informações.

    Args:
        NF (str): Nome do arquivo da nota fiscal.
        caminho (str, opcional): Caminho para o diretório onde o arquivo está localizado. O padrão é 'Notas fiscais'.

    Returns:
        dict: Um dicionário contendo todas as informações extraídas da nota fiscal.
    """
    caminho_completo = os.path.join(caminho, NF)

    # Abrindo o arquivo XML e parseando o documento
    with open(caminho_completo, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)

    # Acessando os dados relevantes do XML
    nfse = documento['ConsultarNfseResposta']['ListaNfse']['CompNfse']['Nfse']['InfNfse']
    prestador = nfse['PrestadorServico']
    tomador = nfse['TomadorServico']

    # Tratando a exceção para diferentes tipos de identificação (CPF ou CNPJ) do tomador
    try:
        cpf_cnpj_comprador = tomador['IdentificacaoTomador']['CpfCnpj']['Cnpj']
    except KeyError:
        cpf_cnpj_comprador = tomador['IdentificacaoTomador']['CpfCnpj']['Cpf']

    # Organizando os dados de retorno em um dicionário
    dic_respostas = {
        'numero_nota': nfse['Numero'],
        'codigo_verificacao': nfse['CodigoVerificacao'],
        'codigo_municipal': prestador['Endereco']['CodigoMunicipio'],
        'data_emissao': nfse['DataEmissao'],
        'valor_total': nfse['Servico']['Valores']['ValorServicos'],
        'cnpj_vendedor': prestador['IdentificacaoPrestador']['Cnpj'],
        'inscricao_municipal': prestador['IdentificacaoPrestador']['InscricaoMunicipal'],
        'nome_vendedor': prestador['NomeFantasia'],
        'razao_social_vendedor': prestador['RazaoSocial'],
        'endereco_vendedor': prestador['Endereco']['Endereco'] + ' - ' + prestador['Endereco']['Bairro'] + ', ' + prestador['Endereco']['Uf'],
        'cep_vendedor': prestador['Endereco']['Cep'],
        'tel_vendedor': prestador['Contato']['Telefone'],
        'email_vendedor': prestador['Contato']['Email'],
        'cpf_cnpj_comprador': cpf_cnpj_comprador,
        'nome_comprador': tomador['RazaoSocial'],
        'servicos': nfse['Servico']['Discriminacao'],
        'endereco_tomador': tomador['Endereco']['Endereco'] + ' - ' + tomador['Endereco']['Bairro'] + ', ' + tomador['Endereco']['Uf'],
        'cep_tomador': tomador['Endereco']['Cep'],
        'tel_tomador': tomador['Contato']['Telefone'],
        'email_tomador': tomador['Contato']['Email']
    }

    return dic_respostas

def identificar_nf(NF, caminho='Notas fiscais'):
    """Identifica o modelo da nota fiscal.

    Esta função identifica o modelo da nota fiscal. Seu uso serve para auxiliar funções que necessitam do modelo.

    Args:
        NF (str): Nome do arquivo da nota fiscal.
        caminho (str, opcional): Caminho para o diretório onde o arquivo está localizado. O padrão é 'Notas fiscais'.

    Returns:
        str: O modelo da nota fiscal, podendo ser 'NFe', 'Xml Servico' ou uma mensagem informando que não foi possível identificar.
    """
    try:
        with open(f'{caminho}\{NF}', 'rb') as arquivo:
            documento = xmltodict.parse(arquivo)
        
        if 'nfeProc' in documento:
            return 'NFe'
        elif 'ConsultarNfseResposta' in documento:
            return 'Xml Servico'
    except:
        return 'Modelo nao suportado/Erro com o arquivo'
  
def planilhar_arquivo(*NF, planilha='separada', caminho='Notas fiscais'):
    """Transforma informações das notas fiscais em planilhas Excel.

    Esta função recebe como entrada uma ou mais notas fiscais e gera planilhas Excel com suas informações. 
    Se desejar planilhar todos os arquivos da pasta, utilize a função planilhar_pasta().

    Args:
        *NF (str): Uma ou mais notas fiscais que deseja planilhar.
        planilha (str, opcional): Modo de planilhamento. 
            - 'separada': gera uma planilha para cada nota. 
            - 'unica': gera uma única planilha com todas as notas. O padrão é 'separada'.
            Obs: Para utilizar 'unica', todas as notas devem ser do mesmo modelo.
        caminho (str, opcional): Caminho para o diretório onde os arquivos das notas fiscais estão localizados. O padrão é 'Notas fiscais'.

    Raises:
        ValueError: Erro caso o modelo não seja atendido ou haja um problema no arquivo XML da nota.
    """
    notas = list(NF)
    dfs = []

    if planilha == 'separada':
        for nota in notas:
            modelo = identificar_nf(nota, caminho=caminho)
            
            if modelo == 'NFe':
                resposta = ler_xml_NFe(nota, caminho=caminho)
            elif modelo == 'Xml Servico':
                resposta = ler_xml_servico(nota, caminho=caminho)
            else:
                raise ValueError('Modelo nao suportado / Erro com a nota')


            tabela = pd.DataFrame.from_dict(resposta)
            tabela.to_excel(f'NF_{nota}.xlsm')

    elif planilha == 'unica':
        for nota in notas:
            modelo = identificar_nf(nota, caminho=caminho)
            
            if modelo == 'NFe':
                resposta = ler_xml_NFe(nota, caminho=caminho)
                resposta_limpa = {chave: valor if valor else None for chave, valor in resposta.items()}
                df = pd.DataFrame(resposta_limpa)
                dfs.append(df)

                tabela = pd.concat(dfs, ignore_index=True)
                tabela.to_excel('NFs.xlsm')

            elif modelo == 'Xml Servico':
                resposta = ler_xml_servico(nota, caminho=caminho)
                resposta_limpa = {chave: valor if valor else None for chave, valor in resposta.items()}
                df = pd.DataFrame(resposta_limpa)
                dfs.append(df)

                tabela = pd.concat(dfs, ignore_index=True)
                tabela.to_excel('NFs.xlsm')
            else:
                raise ValueError('Modelo nao suportado / Erro com a nota')

def planilhar_pasta(caminho='Notas fiscais'):
    """Gera planilhas Excel com informações de todas as notas fiscais em uma pasta.

    Esta função processa todas as notas fiscais em uma determinada pasta e gera planilhas Excel com suas informações.

    Args:
        caminho (str, opcional): Caminho para a pasta que contém as notas fiscais. O padrão é 'Notas fiscais'.

    Raises:
        ValueError: Erro caso não seja possível acessar a pasta especificada
    """
    notas = os.listdir(caminho)
    
    for nota in notas:
        if '.xml' in nota:
            planilhar_arquivo(nota)

