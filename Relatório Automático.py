# Esse código foi feito para ser executado no google colab, ele faz um relatório da planilha do excel "TABELA ACORDOS". 
# A planilha tem as colunas: "VENC" datetime; "VALOR " float; "ITEM " object; "Nº DOCUMENTO" object; "DATA PAG" float; "VALOR PAG" float. 
# O objetivo era exibir um relatório de acordos ainda não pagos, mostrando os que estão atrasados, os que vencem no dia atual e os que vencem nos próximos 30 dias. 
# O código também deveria informar quanto seria gasto com acordos atrasados, com os que se tem data de vencimento para o dia de hoje e o quanto deve ser gasto com a soma dos acordos seguintes.


#Imports básicos para o código funcionar
import pandas as pd
from openpyxl import *

wb = load_workbook(filename = "/content/TABELA ACORDOS.xlsx")
#Salva a lista de nomes na variável nomes
nomes = wb.sheetnames

#Lê a tabela 1 e mostra a parte desejada dela
planilha1 = pd.read_excel(r"/content/TABELA ACORDOS.xlsx", engine = "openpyxl", sheet_name=nomes[0], usecols="D:I", skiprows=4)
planilha1.head()

#Mostrar apenas os valores que não são vazios na coluna desejada
filtro1 = planilha1["VALOR "].notnull()
planilha1 = planilha1[filtro1]

#Resetar o valor do índice
planilha1 = planilha1.reset_index(drop = True)
planilha1.head()

#Guardar o tamanho da tabela em uma variável
tamanho = planilha1["VALOR "].notnull().sum()

#Verificar os tipos de dado de cada coluna
planilha1.dtypes

#Ordenar as datas válidas
planilha1 = planilha1.sort_values(by = ["VENC"])
#Resetar o índice
planilha1 = planilha1.reset_index(drop = True)

#Receber os dados do dia de hoje para poder alterar para o nosso fuso
hoje = pd.Timestamp.today(tz = "America/Sao_Paulo")
hoje.normalize()

#Fazer com que a variável hoje só tenha os valores de ano, mes e dia
hoje = pd.Timestamp(year=hoje.year, month=hoje.month, day=hoje.day)

#Criar uma variável que guarda os acordos que ainda não foram pagos
n_pagos = planilha1
filtro2 = n_pagos["VALOR PAG"].isnull()
n_pagos = n_pagos[filtro2]
tamanho = n_pagos["VALOR PAG"].isnull().sum()

#Criar uma variável para verificar se há acordos que já passaram do prazo de vencimento
vencidos = n_pagos
#Fazer uma variável que guarda os que vencem hoje
venc_hoje = n_pagos
#Guardar os valores que vencem após o dia de hoje
pag_futuro = n_pagos

i = 0

#Se não houverem valores nulos há uma verificação para remoção de dados mediante filtros
if (n_pagos["VENC"].notnull().sum() == tamanho):
  while (i < tamanho):
    if (hoje <= n_pagos["VENC"][i]):
      vencidos = vencidos.drop(i)
    if (hoje != n_pagos["VENC"][i]):
      venc_hoje = venc_hoje.drop(i)
    if (hoje > n_pagos["VENC"][i]):
      pag_futuro = pag_futuro.drop(i)
    i = i + 1

#Caso haja algum valor vazio dentro dos vencimentos terá uma outra camada de verificação
else:
  while (i < tamanho):
    if (type(vencidos["VENC"][i]) != pd._libs.tslibs.timestamps.Timestamp):
      print("Vencimento ", str(n_pagos["VENC"][i]), " desconsiderado")
      print("Item: ", str(n_pagos["ITEM"][i]), "\nNº Documento: ", str(n_pagos["Nº DOCUMENTO	"][i]))
      vencidos = vencidos.drop(i)
      venc_hoje = venc_hoje.drop(i)
      pag_futuro = pag_futuro.drop(i)

    #vencidos não recebe valores que vem em dias seguintes 
    if (hoje <= vencidos["VENC"][i]):
      vencidos = vencidos.drop(i)
    # venc_hoje não recebe valores que vem em dias diferentes de hoje
    if (hoje != vencidos["VENC"][i]):
      venc_hoje = venc_hoje.drop(i)
    #pag_futuro não recebe valores que vem em dias antes de hoje
    if (hoje > n_pagos["VENC"][i]):
      pag_futuro = pag_futuro.drop(i)
    i = i + 1

vencidos = vencidos.reset_index(drop = True)
venc_hoje = venc_hoje.reset_index(drop = True)
pag_futuro = pag_futuro.reset_index(drop = True)

i = 0

#Intervalo de 30 dias no futuro em relação a hoje
mes = pd.Timedelta(-30, "d")

#Separar os que vencem nos próximos 30 dias (não contando hoje)
mes_atual = pag_futuro
t_mes = mes_atual["VENC"].notnull().sum()

#Remover a linha de posição i quando a diferença for de mais de 30 dias
while(i < t_mes):
  if ((hoje - mes_atual["VENC"][i]) < mes):
    mes_atual = mes_atual.drop(i)
  else:
    pag_futuro = pag_futuro.drop(i)
    
  i = i + 1
#Reseta o index
mes_atual = mes_atual.reset_index(drop = True)
pag_futuro = pag_futuro.reset_index(drop = True)
mes_atual.head()

t_pag_futuro = pag_futuro["VENC"].notnull().sum()
t_atual = mes_atual["VENC"].notnull().sum()
if(t_atual > 0):
  prox_mes = mes_atual["VENC"][0]
  prox_item = mes_atual["ITEM "][0]
  prox_valor = mes_atual["VALOR "][0]

t_vencidos = vencidos["VENC"].notnull().sum()
t_hoje = venc_hoje["VENC"].notnull().sum()

print("                                  Relatório\n\n")

print("----------------------------------Atrasados----------------------------------\n")
if (t_vencidos > 0):

  print("Existe(m) ", t_vencidos, " acordo(s) com pagamento atrasado, sendo ele(s): \n")

  print(vencidos)


  venc_divida = str(vencidos["VALOR "].sum())
  venc_divida = venc_divida.replace(".", ",")

  print("No total há uma dívida de R$" + venc_divida + " referente as contas atrasadas\n")
  print("-----------------------------------------------------------------------------")

else:

  print("Não existem acordos com pagamento atrasados!\n")
  print("-----------------------------------------------------------------------------")


print("----------------------------------Hoje---------------------------------------\n")
if (t_hoje > 0):

  print("Existe(m) ", t_hoje, " acordo(s) que vencem hoje, sendo ele(s): \n")
  print(venc_hoje)
  pag_hoje = str(venc_hoje["VALOR "].sum())
  pag_hoje = pag_hoje.replace(".", ",")

  print("Hoje você deve pagar um total de R$" + pag_hoje)
  print("-----------------------------------------------------------------------------\n")
else:
  print("Nenhum acordo vence hoje! \n")
  print("-----------------------------------------------------------------------------\n")


print("------------------------------Próximos 30 dias-------------------------------\n")
if (t_mes > 0):
  print("No mês atual você ainda tem os seguintes pagamentos a fazer: \n")
  print(mes_atual)
  print("\nO próximo pagamento a ser feito é o seguinte: \n", prox_mes.day, "/", prox_mes.month, "/", prox_mes.year, "----", prox_item,"---- R$", prox_valor)
  
  pag_30_dias = str(mes_atual["VALOR "].sum())
  pag_30_dias = pag_30_dias.replace(".", ",")
  print("\nNo total isso vai te custar: R$" + pag_30_dias + "\n")
  print("-----------------------------------------------------------------------------\n")

else:
    print("Não há pagamentos para os próximos 30 dias")
    print("-----------------------------------------------------------------------------\n")


print("----------------------------------Futuro-------------------------------------\n")
if(t_pag_futuro > 0):
  print("Mais para frente ainda tem ", str(t_pag_futuro), " contas a serem pagas. Sendo elas: \n")
  print(pag_futuro)
  print("\nO primeiro pagamento fora dos 30 dias será daqui ", str((pag_futuro["VENC"][0]-hoje).days), " dias.")
  
  print("R$", pag_futuro["VALOR "][0], "----", pag_futuro["ITEM "][0])
  print("-----------------------------------------------------------------------------")

else:
  print("Não há pagamentos pendentes após 30 dias")
  print("-----------------------------------------------------------------------------")

# Link de acesso colab: https://colab.research.google.com/drive/1IX-jqrNOEOHv8NARwC3Kpb7B-RLPTU_H?usp=sharing
