import pandas as pd
import urllib.parse

def formatoTelefone(telefone):
    caracteres = ('(', ')', '-', ' ', '+')
    telefoneFormatado = ''
    telefone = str(telefone)
    if any(caractere in telefone for caractere in caracteres):
        for numero in telefone:
            if numero in caracteres:
                continue
            else: 
                telefoneFormatado += numero
    return telefoneFormatado

def formatarCNPJ(cnpj):
    # tira a formatacao do cnpj
    caracteres = ('.', '/', '-')
    cnpjFormatado = ''
    cnpj = str(cnpj)
    if any(caractere in cnpj for caractere in caracteres):
        for numero in cnpj:
            if numero in caracteres:
                continue
            else: 
                cnpjFormatado += numero
    else:
        cnpjFormatado = cnpj
    while len(cnpjFormatado) < 14:
            cnpjFormatado = "0" + cnpjFormatado
    return cnpjFormatado

# fonte = str(input("Digite o path da planilha fonte: "))
# notasSalvar = str(input("Digite o path de onde deseja salvar: "))
fonte = "c:\\Users\\Samuel\\Downloads\\TESTE_LINKS.xlsx"
notasSalvar = "c:\\Users\\Samuel\\Downloads"
fonte = pd.read_excel(fonte)
fonte = pd.DataFrame(fonte)

for n, id in enumerate (fonte["ID"]):
        
    try:
        linhas = fonte.loc[n, "LINHAS"]
    except:
        None
    try:
        gestor = fonte.loc[n, "GESTOR"]
    except:
        None
    try:
        telefone = fonte.loc[n, "TELEFONE"]
        telefone = formatoTelefone(telefone)
    except:
        None
    try:
        email = fonte.loc[n, "EMAIL"]
    except:
        None
    try:
        cnpj = fonte.loc[n, "CNPJ"]
        cnpj = formatarCNPJ(cnpj)
    except:
        None
    try:
        nomeEmpresa = fonte.loc[n, "EMPRESA"]
    except:
        None

    mensagem = f"Olá!\nMeu nome é Samuel e sou Gerente de Negócios Vivo Empresas, sou responsável por atender a {nomeEmpresa}, CNPJ {cnpj}.\nQuero destacar que temos *LINK DEDICADO* disponível para o seu endereço.\n\nQue tal agendarmos um horário para que eu possa apresentar um orçamento sem compromisso? Se preferir, posso enviar a proposta por aqui mesmo.\n\nAlém disso, temos excelentes opções para aluguel de equipamentos!"

    if len(telefone) == 10 or len(telefone) == 12:
        continue
    else:
        if telefone[0:2] == "55":
            link = f'https://api.whatsapp.com/send?phone={telefone}&text={mensagem}'
        else:
            link = f'https://api.whatsapp.com/send?phone=55{telefone}&text={mensagem}'

    link = urllib.parse.quote(link, safe=':/?&=')
    # with open(f"{notasSalvar}\\Links e gestores.txt", 'a', encoding='utf-8') as notas:
    #     notas.write(cnpj + "\n" + nomeEmpresa + "\n" + telefone + "\n" +  link + "\n\n\n")
    # notas.close()

    with open(f"{notasSalvar}\\Links e gestores.txt", 'a', encoding='utf-8') as notas:
        notas.write(link + "\n\n\n")
    notas.close()