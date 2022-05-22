#Bibliotecas que serão utilizadas nesse projeto
from pandas_datareader import data as web
import pandas as pd
import matplotlib.pyplot as plt
from datetime import date
import win32com.client as win32

#lista das empresas a serem cotadas (essa lista futuramente será dinâmica)
lista_empresa = ["ITSA4","CIEL3","XPCI11","ALZR11","HSML11","TAEE11","FBOK34","BTLG11","TSLA34"]

#pegar a data atual e tratar ela no padrão necessário
today = date.today()
today = today.strftime("%d-%m-%Y")

#lista vazia para armazenar os valores de cotação
lista_cotacao = []

#laço de repetição para buscar a cotação de cada uma das empresas listadas
for empresa in lista_empresa:
    #print(empresa)
    #buscar a cotação pelo código da empresa na Bolsa
    cotacao = web.DataReader(f'{empresa}.SA',data_source='yahoo',start="20-04-2022", end=today)
    #plotar um gráfico com as cotações dia a dia
    cotacao["Adj Close"].plot()

    #encontrar a média de preço dessas ações no período selecionado
    media = cotacao["Adj Close"].mean()

    #pegar o valor atual do cotação
    valor_atual = round(cotacao['Adj Close'].iloc[-1],2)

    #guardar o valor de cotação na lista
    lista_cotacao.append(valor_atual)

    #alterando o titulo do gráfico
    title = str(empresa)+" Média mês R$ "+str(float("{:.2f}".format(media)))
    plt.title(title)

    #salvar a imagem com o gráfico de cada empresa
    plt.savefig(f'{empresa}.png')

#criando um DataFrame com as empresa x cotações atuais capturadas no laço
cotacao_real = {'Ação':lista_empresa,
                'Valor Atual':lista_cotacao}
dfcotacao = pd.DataFrame(cotacao_real)


#envio de e-mail com a tabelas das cotações atualizadas do dia
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'lucasdelegredo@gmail.com'
mail.Subject = 'Cotação de ações diárias'
mail.HTMLBody = f'''

    <h3>Segue o relatório referente aos valores diários atuais das ações listadas:</h3>
    <br>
    <h4>Data: {today}</h4>
    <br>
    <h2>Cotações:</h2>
    {dfcotacao.to_html()}

'''
mail.Send()

