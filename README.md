Conversor CSV para XLSX
Este script em Python converte arquivos .csv em arquivos .xlsx sem o uso do Excel, com o objetivo de automatizar o processo de conversão de arquivos sem remover os "0"s iniciais das células. Isso é ideal para sistemas que exigem códigos de barras com "0"s iniciais.

Pré-requisitos
Este script requer as seguintes bibliotecas do Python para serem instaladas:

openpyxl (para manipulação de arquivos .xlsx)
csv (para manipulação de arquivos .csv)
os (para manipulação de arquivos)
time (para manipulação de tempo)

Uso
Execute o script em seu ambiente Python ou clicando duas vezes no arquivo .py.
Digite o nome do operador. O nome só pode conter letras e espaços.
Digite o mês. A entrada pode ser numérica com ou sem um "0" inicial.
Digite o ano. O ano deve ser um número de quatro dígitos entre 2021 e 2100.
Confirme que o diretório dos arquivos .csv está correto.
Confirme que o diretório para salvar os arquivos .xlsx está correto.
Aguarde o script terminar de converter os arquivos.
Confirme que os arquivos .xlsx foram salvos no diretório correto.
O script irá automaticamente excluir os arquivos .csv temporários.

Observações
O script usa loops para ler e escrever os arquivos. O script irá automaticamente parar se não houver mais arquivos para ler.
Se o script encontrar um erro, ele exibirá uma mensagem de erro e irá parar de ser executado.
O script usa o prefixo delimitado para criar arquivos .csv temporários. Este prefixo pode ser alterado editando o script.
O script usa o ponto e vírgula ";" como delimitador padrão para os arquivos .csv. Se o delimitador for diferente, o script pode ser modificado de acordo.
Exemplo de arquivo: Carlos_2023-MES-01_1.csv
