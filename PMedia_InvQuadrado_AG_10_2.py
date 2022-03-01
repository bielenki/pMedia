# ---------------------------------------------------------------------------
# Gera��o da Precipitacao M�dia - Inverso Quadrado da dist�ncia 
# Cl�udio Bielenki J�nior
# Especialista em Geoprocessamento - SUM - ANA
# Novembro 2013
# V 0.0.2   21112013
# ---------------------------------------------------------------------------

# Importando os m�dulos
import win32com, sys, string, os, xlrd, arcpy
from win32com.client import Dispatch

# Licen�a Arc Info necess�ria
arcpy.SetProduct("ArcInfo")

# Recebendo as vari�veis de entrada
work=arcpy.GetParameterAsText(0) # Armazena o endere�o do workspace
fc_estacao=arcpy.GetParameterAsText(1) # Definindo a feature class das esta��es
fc_bacia=arcpy.GetParameterAsText(2) # Definindo a feature class da �rea da bacia
FieldID =arcpy.GetParameterAsText(3) # Campo da feature class bacia que a individualiza
file = arcpy.GetParameterAsText(4)# Arquivo de entrada de dados de pluviometria
saida_xls = arcpy.GetParameterAsText(5) # Arquivo de saida de dados de pluviometria com ponderadores

arcpy.env.workspace=work # Definindo o local de armazenamento
arcpy.gp.overwriteOutput="True" # Permite a sobreposi��o de resultados
# Definindo o extent do projeto
ext_x=[] # Vetor para armazenar os valores das ordenadas da �rea de trabalho
ext_y=[] # Vetor para armazenar os valores das abcissas da �rea de trabalho
dsc_estacao=arcpy.Describe(fc_estacao) # Vari�vel recebe o descritor da feature class das esta��es
dsc_bacia=arcpy.Describe(fc_bacia) # Vari�vel recebe o descritor da feature class da bacia
# Recebendo do descritor os valores de limite de cada feature class
ext_x.append(dsc_estacao.Extent.XMax)
ext_x.append(dsc_estacao.Extent.XMin)
ext_x.append(dsc_bacia.Extent.XMax)
ext_x.append(dsc_bacia.Extent.XMin)
ext_y.append(dsc_estacao.Extent.YMax)
ext_y.append(dsc_estacao.Extent.YMin)
ext_y.append(dsc_bacia.Extent.YMax)
ext_y.append(dsc_bacia.Extent.YMin)
# Definindo os limites da �rea de trabalho
x_maximo=max(ext_x)+10000
x_minimo=min(ext_x)-10000
y_maximo=max(ext_y)+10000
y_minimo=min(ext_y)-10000
arcpy.env.extent= arcpy.Extent(x_minimo,y_minimo,x_maximo,y_maximo)

# Lendo os dados de entrada do xls
xls = xlrd.open_workbook(file) # Abre o arquivo xls de dados de entrada
dados=xls.sheets()[0] # Vari�vel dados aponta para a 1� planilha do xls
ncol=dados.ncols # N�mero de colunas
nrow=dados.nrows # N�mero de linhas
linha=dados.row_values # Vari�vel dados recebe os valores da planilha
estacao=linha(0) # Vetor estacao recebe a 1� linha da planilha

# Prepara o arquivo de dados para grava��o da PM�dia
nome_plans=[] # Vetor que receber� os identificadores das bacias para nomear as planilhas
pbacias = arcpy.SearchCursor(fc_bacia) # Posciciona o cursor na fei��o de bacias
pbacia = pbacias.next() # Inicia o looping pelas rows
while pbacia:
   nome_plans.append(pbacia.getValue(FieldID)) # Captura os valores dos identificadores
   pbacia=pbacias.next() # Incrementa o looping

xlApp = Dispatch("Excel.Application") # Inicia o Excel
wb = xlApp.Workbooks.Open(saida_xls) # Abre o arquivo de sa�da
for i in range(wb.Sheets.Count, 1, -1): # Looping pelas planilhas
    wb.Sheets(i).Delete() # Deleta as planilhas exceto a 1�
for i in range(0, len(nome_plans)): # Looping pelos identificadores das bacias
    planilha = xlApp.ActiveWorkbook.Worksheets.Add() # Adiciona uma planilha
    planilha.Name=nome_plans[i] # Nomeia a planilha de acordo com os identificadores

# Loop pelas linhas do arquivo de dados pluviometricos
num=1 # Contador
for j in range(1,nrow):
    sel =['']*ncol # Inicializa a vari�vel sel
    expressao=''# Inicializa a vari�vel espressao
    aux=''# Inicializa a vari�vel auxiliar
    est_vet=[]# Inicializa a vari�vel est_ve
    prec_vet=[] # Inicializa a vari�vel prec_ve
    # Loop pelos postos pluviom�tricos
    for i in range(1,ncol):
        if linha(j)[i]!= '': # Verifica se existe dado medido, caso sim copia o c�digo para montar a express��o SQL
            est_valor=(linha(0)[i]) # L� o c�digo da esta��o na primeira linha da planilha
            est_vet.append(est_valor) # Armazena o c�digo no vetor est_vet
            prec_valor=(linha(j)[i])# L� o valor de precipita��o
            prec_vet.append(prec_valor) # Armazena o valor de precipita��o no vetor prec_vet
            sel[i] = '\"Codigo\" =' + ' ' + str(int(est_valor)) + ' ' + 'OR ' # Monta a express�o SQL para a referida esta��o
        aux=aux+sel[i] # Monta a express�o SQL
    # Apaga o �ltimo OR da express�o SQL
    for i in range(len(aux)-4):
        expressao=expressao+aux[i]

    # Nomes aos arquivos auxiliares
    selecao="selecao_"+str(num)
    raster="raster_"+str(num)
    table="table_"+str(num)


    # Processo: Selececionando...
    arcpy.Select_analysis(fc_estacao, "in_memory\\"+str(selecao), expressao) # Seleciona as esta��es de acordo com a espress�o SQL
    pRows = arcpy.UpdateCursor("in_memory\\"+str(selecao)) # Posciciona o cursor na fei��o de esta��es selecionadas
    pRow = pRows.next() # Vai para a primeira linha da feature class
    while pRow: # Loop pelas linhas
        cod = pRow.codigo # Armazena o c�digo da esta��o
        prec_valor=prec_vet[(est_vet.index(cod))] # Copia o valor da precipita��o do vetor prec_vet de acordo do index de posi��o do c�digo da esta��o para o vetor prec_valor
        pRow.prec = prec_valor # Copia o valor da precipita��o para a feature class de esta��es selecionadas
        pRows.updateRow(pRow) # Grava os valores
        pRow=pRows.next() # Passa para a nova linha

    # Processo: cria raster...
    raster=arcpy.sa.Idw("in_memory\\"+str(selecao),"prec") # Realiza a interpola��o pelo IDW
    arcpy.sa.ZonalStatisticsAsTable(fc_bacia,FieldID,raster,"in_memory\\"+str(table),'TRUE','MEAN') # Calcula a m�dia para a �rea das bacias
    pTables = arcpy.SearchCursor("in_memory\\"+str(table)) # Posciciona o cursor na tabela do zonal
    pTable = pTables.next() # Vai para a primeira linha
    while pTable: # Looping pelas linhas da tabela
        plan=str(pTable.getValue(FieldID)) # Copia o valor do campo
        media=pTable.MEAN # Copia da tabela a m�dia do IDW
        planilha=wb.Worksheets(plan) # Ativa a planilha no Excel de acordo com o valor do campo
        planilha.Cells(j,1).Value = media # Escreve o valor da m�dia na planilha
        pTable = pTables.next() # Incrementa o looping

    num=num+1 # Incrementa a vari�vel num

wb.save # Salva o Excel

# Terminar aplica��o
xlApp.Quit()

# Limpar a mem�ria
del xlApp
