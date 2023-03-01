#Este programa converte arquivos .csv para .xlsx sem usar o Excel com o objetivo de automatizar o processo de conversão de arquivos sem tirar o "0" 
#das celulas que iniciam com "0", ideal para sistemas que precisam de códigos de barras com "0" no início.

#openpyxl (Biblioteca para manipulação de arquivos .xlsx)
#csv (Biblioteca para manipulação de arquivos .csv)
#os (Biblioteca para manipulação de arquivos)
#time (Biblioteca para manipulação de tempo)
from openpyxl import Workbook
import csv
import os
import time


#Lista de meses para tratamento de erro
meses = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]
meses2 = ["1", "2", "3", "4", "5", "6", "7", "8", "9"]

#Input para o nome do operador e tratamento de erro
while True: #O nome pode ser composto por letras e espaços
    nome = input("Digite o nome do operador: ").upper()
    if all(c.isalpha() or c.isspace() for c in nome):
        break
    else:
        print("Nome inválido, digite novamente. (Apenas letras e espaços.)")

#Input para o mês e tratamento de erro (o input deve ser numérico)
while True: #O mês pode ser digitado com ou sem o "0" no início
    mes = input("Digite o mês: ")
    if mes in meses2:
        mes = "0" + mes
    if mes in meses:
        break
    else:
        print("Mês inválido, digite novamente. (Apenas números de 1 a 12.)")

#Input para o ano e tratamento de erro
while True: #O ano deve ser digitado com 4 dígitos e estar entre 2021 e 2100
    ano = input("Digite o ano: ")
    try:
        if len(ano) == 4 and int(ano) >= 2021 and int(ano) <= 2100:
            break
        else:
            print("Ano inválido, digite novamente. (Apenas números de 4 dígitos e maiores que 2020.)")
    except:
        print("Ano inválido, digite novamente. (Apenas números de 4 dígitos e maiores que 2020.)")

#Variaveis para controle de loop
x = 1
c = 1
l = 0
y = 1

#Loop que lê os arquivos .csv e escreve em um arquivo .csv com delimitador ",".
while True:
    try:
        #Caminho para os arquivos .csv
        reader = csv.reader(open(f"Insira o diretório dos arquivos csv aqui\\{nome}_{ano}-MES-{mes}_{c}.csv", "r"), delimiter=';')
    except:
        if l == 0: #Tratamento de erro para caso o arquivo não seja encontrado
            print("ARQUIVOS NÃO ENCONTRADOS. Confirme que os dados digitados estão corretos e tente novamente.")
            sair = input("Pressione ENTER para sair...")
            exit()
        print("Convertendo arquivos...")
        break
    with open(f"delimitado{c}.csv", 'w') as file:
        writer = csv.writer(file, delimiter=',')
        writer.writerows(reader)
    l = l + 1
    c = c + 1

#Loop que lê os arquivos .csv com delimitador "," e escreve em um arquivo .xlsx.
while x < c:
    wb = Workbook()
    ws = wb.active
    try:
        with open(f'delimitado{x}.csv', 'r') as f:
            for row in csv.reader(f):
                ws.append(row)
    except: #O loop é interrompido quando não há mais arquivos .csv para ler
        break

    try:    #Caminho para os arquivos .xlsx
        wb.save(f"Insira o diretório dos arquivos xlsx aqui\\{nome}_{ano}-MES-{mes}_{x}.xlsx")
    except:
        print("Caminho para salvar os arquivos .xlsx não encontrado. Confirme que os dados digitados estão corretos, que a extensão dos arquivos é correta e tente novamente")
        sair = input("Pressione ENTER para sair...")
        exit()
    n = (f"{nome}_{ano}-MES-{mes}_{x}.xlsx")
    print(f"\033[1;32m Arquivo {n} convertido com sucesso. \033[0;0m")
    x = x + 1

#Loop que deleta os arquivos temporários .csv
print("Deletando arquivos temporários...")
while y < c:
    try:
        os.remove(f"delimitado{y}.csv")
        time.sleep(0.5)  #Tempo de espera para evitar erro de permissão
    except Exception as e: #Tratamento de erro caso o arquivo não seja apagado
        print(f"Error deleting delimitado{y}.csv: {e}")
    y = y + 1

#Mensagem de sucesso
print("Tudo certo! Todos os arquivos foram convertidos com sucesso e salvos na pasta ArquivosConvertidos.")
sair = input("Pressione ENTER para sair...")
exit()