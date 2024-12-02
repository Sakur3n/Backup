import requests
import json
import datetime
import time

    
linha_lista = []
def lista_produtos (pagina):
    url = 'https://app.omie.com.br/api/v1/geral/produtos/'
    app_key = 'Chave de acesso'
    app_secret = 'Chave secreta'

    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }

    payload = {
    "call": "ListarProdutos",
    "app_key": app_key,
    "app_secret": app_secret,
    "param": [
        {
  "pagina": pagina,
  "registros_por_pagina": 50,
  "apenas_importado_api": "N",
  "filtrar_apenas_omiepdv": "N"
    }
    ]
}   
    response = requests.post(url, headers=headers, data=json.dumps(payload))

    produto = response.json()
    produtos = produto['produto_servico_cadastro']

    return produtos

def estoque_produto(codigo):
    from datetime import date
    data = date.today().strftime('%d/%m/%Y')

    url1 = 'https://app.omie.com.br/api/v1/estoque/resumo/'
    app_key = '3931984068012'
    app_secret = 'b9769b1862544c8b2bbaa5b0ba88d675'

    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }

    payloadd = {
    "call": "ObterEstoqueProduto",
    "app_key": app_key,
    "app_secret": app_secret,
    "param": [{
                "cCodigo": codigo,
                "dDia": data
            }]}
    
    responsee = requests.post(url1, headers=headers, data=json.dumps(payloadd))

    prod = responsee.json()
    lista_estoque = prod['listaEstoque']
    return lista_estoque

        
def planilha (dados):
    from datetime import date
    import pandas as pd

    df = pd.DataFrame(dados)

    data = date.today().strftime('%d%m%Y')
    backup_do_dia = f"Backup{data}.xlsx"
    df.to_excel(backup_do_dia, index=False)

    return backup_do_dia

def envia_email():
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    import os
    from datetime import date
    data = date.today().strftime('%d%m%Y')
    data1 = date.today().strftime('%d/%m/%Y')

    corpo_email = f'''

    <p>Backup realizado com sucesso!!</p>
    '''
    
    assunto = f"Backup do dia {data1}"

    msg = MIMEMultipart()
    msg['Subject'] = assunto
    msg['From'] = 'backup.graonatural@gmail.com'
    msg['To'] = 'backup.graonatural@gmail.com'
    senha = 'dlwdrumpabufhovi'
    msg.attach(MIMEText(corpo_email, 'html'))

    
    arquivo_anexo = f"Backup{data}.xlsx"

    with open(arquivo_anexo, "rb") as anexo:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(anexo.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(arquivo_anexo)}')
        msg.attach(part)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    
    s.starttls()
    s.login(msg['From'], senha)

    s.sendmail(msg['From'], [msg['To']], msg.as_string())
    s.quit()

    print(f'Email enviado {arquivo_anexo}')

    os.remove(arquivo_anexo)

def inicio():
    try:
        total = 35
        
        for pagina in range (1, total +1):
            a = lista_produtos(pagina)

            for item in a:
                codigo = item.get('codigo')
                descricao = item.get('descricao')
                familia = item.get('descricao_familia')
            
                if codigo != None:
                    b = estoque_produto(codigo)
                    
                    for iten in b:
                        estoque_fisico = iten.get('fisico')
                        valor_cmc = iten.get('nCMC')
                        valor_venda = iten.get('nPrecoUnitario')

                        dicionario = {
                        'codigo': codigo,
                        'descricao': descricao,
                        'estoque': estoque_fisico,
                        'CMC': valor_cmc,
                        'Valor de venda': valor_venda,
                        'familia do produto': familia,
                        }

                        linha_lista.append(dicionario)
                        
                        print (" Código:", codigo, '\n',"Descrição:", descricao,'\n',"Estoque:", estoque_fisico,'\n',"Valor CMC:", valor_cmc,'\n',"Valor de Venda:", valor_venda,'\n ~~~~~~~~')

        planilha(linha_lista)
        linha_lista.clear()
        


    except RuntimeError as e:
        a = estoque_produto()
        
        print(f"Erro de execução: {e}")
        print(a)

    envia_email()
    
def on():
    entrada = '21:21:00'
    print(entrada)
    a = True

    while a:
    
        t_agora = datetime.datetime.now()
        O = t_agora.strftime("%H:%M:%S")
        
        print(O)

        if O == entrada:
            inicio()
            print("deu certo")
            
        time.sleep(1)

inicio()
