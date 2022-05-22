# Cotação de Ações Atualizadas em Python
A ideia desse projeto seria pegar algumas cotações de empresas na bolsa de valores, capturar os valores, gerar e armazenar gráficos e enviar automaticamente no e-mail desejado

Esse projeto foi desenvolvido todo em python e as bibliotecas utilizadas foram:

![image](https://user-images.githubusercontent.com/74476423/169717209-93d0ff50-49ea-4b15-b19e-514283d161ea.png)

Utilizei de uma lista fixa de empresas a serem cotadas na bolsa, porém em um próximo upgrade a ideia é deixar essa lista de empresas dinâmicas conforme o usuário desejar alterar

![image](https://user-images.githubusercontent.com/74476423/169717257-1d8e8327-8a17-451d-9b82-7893a936f65d.png)

Todo o código estpa disponibilizado em Python e comentado detalhando o passo a passo utilizado

O comando utilizado para buscar as ações descritas durante um período de tempo selecionado vem através do servidor do yahoo, dessa forma:

![image](https://user-images.githubusercontent.com/74476423/169717321-d2732af2-9088-4017-9cb3-87b2ae597677.png)

Com essas informações é possível gerar um gráfico das cotações atualizadas por dia, utilizando o comando:

![image](https://user-images.githubusercontent.com/74476423/169717341-11009853-fb27-43bc-811c-6492f77995c0.png)

Depois disso foi criado um dataframe com as informações vindas das listas de "ações" e "cotações" que criei anteriormente

![image](https://user-images.githubusercontent.com/74476423/169717373-99f7d2b1-7426-4a0d-8ee7-def1d3bc739e.png)


Depois disso utilizei de mail = outlook.CreateItem(0) que vem da biblioteca win32com.client
Criando então a instrução para enviar os emails com informações do dataframe:

![image](https://user-images.githubusercontent.com/74476423/169717423-9eb20e70-95aa-402a-bbc3-a3ddadba9453.png)



