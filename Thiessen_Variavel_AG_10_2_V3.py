# ---------------------------------------------------------------------------
# Gera��o de Pol�gonos de Thiessen Vari�veis  - ArcGis 10.2
# Cl�udio Bielenki J�nior
# Especialista em Geoprocessamento -  SUM  -  ANA
# Novembro 2013     21112013
# ---------------------------------------------------------------------------

# Importando os m�dulos
import win32com, sys, string, os
from win32com.client import Dispatch
import xlrd
import arcpy
# Recebendo as variaveis de entrada
work=arcpy.GetParameterAsText(0) # armazena o endereco do workspace
fc_estacao=arcpy.GetParameterAsText(1) # definindo a feature class das estacoes
fc_bacia=arcpy.GetParameterAsText(2) # definindo a feature class da �rea da bacia
file = arcpy.GetParameterAsText(3)# arquivo de entrada de dados de pluviometria
saida_xls = arcpy.GetParameterAsText(4) # arquivo de saida de dados de pluviometria com ponderadores

arcpy.env.workspace=work #definindo o local de armazenamento

#define o extent para os poligonos de thiessen
ext_x=[]
ext_y=[]
dsc_estacao=arcpy.Describe(fc_estacao)
dsc_bacia=arcpy.Describe(fc_bacia)
ext_x.append(dsc_estacao.Extent.XMax)
ext_x.append(dsc_estacao.Extent.XMin)
ext_x.append(dsc_bacia.Extent.XMax)
ext_x.append(dsc_bacia.Extent.XMin)
ext_y.append(dsc_estacao.Extent.YMax)
ext_y.append(dsc_estacao.Extent.YMin)
ext_y.append(dsc_bacia.Extent.YMax)
ext_y.append(dsc_bacia.Extent.YMin)
x_maximo=max(ext_x)+10000
x_minimo=min(ext_x)-10000
y_maximo=max(ext_y)+10000
y_minimo=min(ext_y)-10000
arcpy.env.extent= arcpy.Extent(x_minimo,y_minimo,x_maximo,y_maximo)

arcpy.gp.overwriteOutput="True" #permite a sobreposicao

xls = xlrd.open_workbook(file) #abre o arquivo xls de dados de entrada
dados=xls.sheets()[0] #vari�vel dados aponta para a 1� planilha do xls
ncol=dados.ncols #n�mero de colunas
nrow=dados.nrows #n�mero de linhas
linhas=dados.row_values #vari?vel dados recebe os valores da planilha
estacao=linhas(0) #vetor estacao recebe a 1� linha da planilha
linha=[0]*(nrow)#cria o vetor linha

# looping para ler cada linha de dados
for i in range(1,nrow):
	linha[i]=linhas(i)

combinacao=[0]*(nrow)#vari�vel combinacao para receber o valor da combina��o soma(2^posi��o) para cada data
combinacao_sel=[0]#vari�vel que armazena os valores de combina��o sem repeti��o
comb_est=[0]#vari�vel que armazena a lista de esta��es para cada combina��o selecao
comb_pond=[0]#vari�vel que armazena a lista de ponderadores para cada combina��o
comb_est_clip=[0]
for j in range(1,nrow): #looping pelas linhas
	combinacao[j]=0 #inicializa a vari�vel combinacao
	est_aux=[] #inicializa a vari�vel est_aux para capturar as esta��es com dados em cada data
	for k in range(1,ncol): #looping pelas colunas
		if linha[j][k]!= '': #verifica se existe dado medido para a esta��o numa data
			combinacao[j]=combinacao[j] + 2**(k-1) #calcula o valor da combina��o
			est_aux.append(estacao[k]) #seleciona os c?dios das esta��es com dados
	if not(combinacao[j] in combinacao_sel) : #verifica se a combina��o j� existe
		combinacao_sel.append(combinacao[j]) #insere na lista de combina��es uma nova combina��o
		comb_est.append(est_aux)#armazena a lista de esta��es para uma nova combina??o

consulta=[0]*len(combinacao_sel)#inicializa a vari�vel consulta

for x in range(1,len(combinacao_sel)): #looping para cada combina��o
	aux='' #inicializa a vari�vel aux
	expressao='' #inicializa a vari�vel expressao
	for y in range(0,len(comb_est[x])): #looping pelas esta��es selecionadas para uma dada combina��o
		sel= '\"codigo\" =' + ' ' + str(int(comb_est[x][y])) + ' ' + 'OR ' #montagem da express�o SQL
		aux=aux+sel
	for i in range(len(aux)-4): #looping para eliminar o �ltimo OR da express�o
		expressao=expressao+aux[i]
	consulta[x]=expressao #armazena as respectivas consultas SQL para cada combina��o

#captura o valor da area da bacia
SRows=arcpy.SearchCursor(fc_bacia)
pRow=SRows.next()
soma=pRow.AREAM2

#funcoes de geoprocessamento para selecionar, criar os poligonos e clipar
for x in range(1,len(consulta)):
	saida_sel="Sel_combinacao_" + str(combinacao_sel[x])#nome para os arquivos de selecao de saida
	saida_thiessen="Thiessen_combinacao_" + str(combinacao_sel[x])#nome para os arquivos de poligonos de saida
	saida_thiessen_Clip="Thiessen_combinacao_" + str(combinacao_sel[x])+"_Clip" #nome para os arquivos de poligonos clipados de saida
	arcpy.MakeFeatureLayer_management(fc_estacao,"in_memory\\"+str(saida_sel), consulta[x])

	arcpy.analysis.CreateThiessenPolygons("in_memory\\"+str(saida_sel),"in_memory\\"+str(saida_thiessen),"ALL")# Processo de geracao dos poligonos de thiessen
	arcpy.analysis.Clip("in_memory\\"+str(saida_thiessen), fc_bacia, saida_thiessen_Clip,"")# Processo de clipagem
	arcpy.Delete_management("in_memory\\"+str(saida_sel))
	arcpy.Delete_management("in_memory\\"+str(saida_thiessen))
	pRows_Thiessen = arcpy.UpdateCursor(saida_thiessen_Clip) # Posciciona um cursor na fei��o de thiessen clipado
	pRow_Thiessen = pRows_Thiessen.next() # Cursor vai para a primeira linha desta feature class
	pond_aux=[] #inicializa a variavel auxiliar para armazenar os ponderadores
	est_clip_aux=[]
	while pRow_Thiessen: # Loop pelas linhas da feature class poligonos thiessen clipados
		area = pRow_Thiessen.Shape_Area
		pRow_Thiessen.area = area
		pRow_Thiessen.pond = (area/soma) #Calcula o ponderador
		est_clip_aux.append(pRow_Thiessen.Codigo)
		pond_aux.append(pRow_Thiessen.pond) # Inclui o ponderador calculado no fim do vetor auxiliar
		pRows_Thiessen.updateRow(pRow_Thiessen) #Atualiza a row
		pRow_Thiessen=pRows_Thiessen.next() #posiciona o cursor na proxima linha
	comb_pond.append(pond_aux) #armazena em um vetor os ponderadores
	comb_est_clip.append(est_clip_aux)
# Prepara o arquivo de dados para grava��o dos ponderadores
xlApp = Dispatch("Excel.Application")
docxls = xlApp.Workbooks.Open(saida_xls)
docxls.Sheets('dados').Select()
planilha = docxls.ActiveSheet

# Gravacao dos ponderadores no xls de saida
for data in range(1,len(combinacao)): # Looping pelas datas
	comb=combinacao[data] #Vari�vel auxiliar para armazenar o valor da combinacao
	index_comb=combinacao_sel.index(comb) # Index para recuperar no vetor de estacoes e ponderadores
	est=comb_est_clip[index_comb] # Recupera as estacoes
	pond=comb_pond[index_comb] # Recupera o ponderador
	for col in range(0,len(pond)): # Looping pelos ponderadores calculados para a data determinada
		codigo_est=est[col] # Recupera a esta��o
		ponderador=pond[col] # Recupera o ponderador
		index_est=estacao.index(codigo_est) # Recupera a posicao da esta��o
		planilha.Cells(data+1,(index_est+ncol+1)).Value = ponderador # Grava o ponderador na planilha de acordo com a posicao recuperada da esta��o

# Salva o xls
xlApp.Visible=1
docxls.save
# Termina a aplica??o xls
xlApp.Quit()
# Limpa a mem?ria
del xlApp

print "Fim da Rotina"