
import win32com.client as win32

import pandas as pd

df = pd.read_excel('lista_email.xlsx')

print(f'Esse são os emails cadastrados:\
 {df}')

novo_email = str(input('Desenha acrescer um novo contato a sua lista de e-mail ? '))


while novo_email == 'Sim' or novo_email == 'SIM' or novo_email == 'S' or novo_email == 's' or novo_email == 'sim' or novo_email == 'desejo' or novo_email == 'Desejo' :
        nome = str(input('Digite o nome da pessoa: '))
        email = str(input('Digite o email : '))
        nome_vazio = []
        nome_vazio.append(nome)
        email_vazio = []
        email_vazio.append(email)
        suporte = {'nome':nome_vazio, 'email': email_vazio}
        df_1 = pd.DataFrame(suporte)
        df_concat = pd.concat([df,df_1])
        df_concat.to_excel('lista_email.xlsx', index=False)
        print('Salvo na base de dados')
        novo_email =  str(input('Desenha acrescer um novo contato a sua lista de e-mail ? '))


df = pd.read_excel('lista_email.xlsx')

assunto = str(input('Digite o assunto do e-emai:'))
menagem= str(input('Digite sua mensagem de e-mail: '))


# criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')




def devolver_nome (email):  #essa funçao devolve o nome atrelado ao email 
    df['is_true']= df['email'] == email  # aqui faz a verificacao se o email informado é igual o da coluna e cria uma nova coluna boliana
    a = df[df['is_true'] == True] # a é um datafram ou a coluna boliana é verdadeira, ou seja, uma unica linha atrelado ao email informado
    a = a['nome'].values.tolist()  #transforma a coluna nome em em uma lista 
    return a[0] # retorna a lista a na primeira posiçao, ou seja, onde o e-mail informado é correspondente ao nome 





# for i in df['email'] :
#         a = devolver_nome(i) # a aqui vai assumir o email de cada i no datafram 
#         outlook = win32.Dispatch('outlook.application')
#         # criar um email
#         email = outlook.CreateItem(0)
#         # configurar as informações do seu e-mail
#         email.To = i
#         email.Subject = "Convite Especial"
#         email.HTMLBody =  f"""
#         <p>Olá {a}  você está  convidado para a Smart</p>


#         <p>Smart </p>

#         """

#         # anexo = "C://Users/joaop/Downloads/arquivo.xlsx"
#         # email.Attachments.Add(anexo)

#         email.Send()
#         print("Email Enviado")


for i, j in zip(df['email'], df['nome']):
        
        outlook = win32.Dispatch('outlook.application')
        # criar um email
        email = outlook.CreateItem(0)
        # configurar as informações do seu e-mail
        email.To = i
        email.Subject = assunto
        email.HTMLBody =  f"""
        <p>Olá {j},</p>
        
        
        <p>  {assunto} </p>


        <p>Matheus Brito Silva </p>

        """

        # anexo = "C://Users/joaop/Downloads/arquivo.xlsx"
        # email.Attachments.Add(anexo)

        email.Send()
        print("Email Enviado")

