**ToolBox PMedia – ArcGis 10**

A *ToolBox* PMedia contém dois scripts escritos em Python para o cálculo da
precipitação média em uma determinada região, dados como entrada um arquivo
Excel xls com as alturas de precipitação a cada época para cada posto
pluviométrico disponível.

O script Thiessen Variáveis calcula a série média de precipitação levando-se em
conta os ponderadores de Thiessen calculados para cada época considerando apenas
os postos pluviométricos com dados disponíveis. Já o script IDW calcula a série
média de precipitação com base na média de pixels dentro das zonas consideradas
(a bacia pode estar subdividida em subacias) para cada época utilizando apenas
os postos pluviométricos com dados disponíveis na interpolação IDW.

**Adicionando a ToolBox ao ArcGis:**

Para adicionar a *ToolBox* ao Arc Gis 10 clique com o botão direito do mouse em
**ArcToolbox** e selecione **Add Toolbox**, em seguida navegue até o diretório
que contém a *ToolBox* , selecione-a e clique em **Open**.

![fig1](https://github.com/bielenki/pMedia/blob/main/Fig/fig1.png?raw=true)

Figura 1: Adicionando a *ToolBox*

Clique com o botão direito do mouse em um dos *scripts* e selecione
**Properties...**

![fig1](https://github.com/bielenki/pMedia/blob/main/Fig/fig2.png?raw=true)

Figura 2: Abrindo as propriedades do script

Na caixa de diálogo **Properties** selecione a guia *Source* para a opção
**Script file:** clique no botão de navegação e aponte para o arquivo
**Thiessen_Variaveis_AG_10_2.py** (no diretório onde foram salvos os arquivos
py) e clique em **OK**.

![fig1](https://github.com/bielenki/pMedia/blob/main/Fig/fig3.png?raw=true)

Figura 3: Propriedades do *script* – Guia *Source*

Repita a operação para o outro script.

**Script Thiessen Variáveis**

Para executar a ferramenta de geração de polígonos de Thiessen variáveis são
necessárias 3 *feature* *class* em um *geodatabase* (gdb) e 2 arquivos xls.

O arquivo de entrada xls com os dados pluviométricos deve conter na primeira
linha os códigos das estações; na primeira coluna têm-se as datas e seguem-se
nas demais colunas as medições de cada um dos postos pluviométricos, sendo que
os meses com falha devem estar sem nenhum valor.

A planilha com os dados deve estar na primeira posição das planilhas na pasta
Excel.

O arquivo de saída que conterá os resultados deve ser uma cópia do arquivo de
entrada apenas com um nome diferente (“saída”, por exemplo).

![fig1](https://github.com/bielenki/pMedia/blob/main/Fig/fig4.png?raw=true)

Figura 4: Exemplo de arquivo xls com os dados de entrada

O *geodatabase* deve conter uma *feature* *class* de pontos com os postos
pluviométricos que na tabela de atributos, além do campo dos códigos das
estações, deve possuir um campo denominado “area” (do tipo “*double*”) para
armazenar as áreas dos polígonos gerados e um campo denominado “pond” (do tipo
“*double*”) para armazenar o valor dos ponderadores a serem calculados. Esta
*feature class* deve estar em uma projeção cartográfica métrica.

![fig1](https://github.com/bielenki/pMedia/blob/main/Fig/fig5.png?raw=true)

Figura 5: Tabela de atributos da *feature class* das estações

O *geodatabase* também devará conter uma *feature class* de polígono que
delimita a área de drenagem da área de estudo em que se deseja calcular a
precipitação média. A tabela de atributos deve conter um campo denominado
“AREAM2” com a área calculada em metros quadrados. Esta *feature class* deve
estar uma projeção cartográfica métrica.

![fig1](https://github.com/bielenki/pMedia/blob/main/Fig/fig6.png?raw=true)

Figura 6: Tabela de atributos da *feature* *class* da bacia

Para executar o script clique duas vezes sobre o ícone do script na toolbox
adicionada. Uma caixa de diálogo para seleção dos arquivos de entrada e saída e
das *feature* *class* é aberta. Clique nos botões de navegação para selecionar
cada um dos parâmetros de entrada e clique em **OK** para rodar a ferramenta.

![fig1](https://github.com/bielenki/pMedia/blob/main/Fig/fig7.png?raw=true)

Figura 7: Janela do script de Thiessen Variáveis

Os ponderadores são então calculados e gravados no arquivo XLS de saída.

![fig1](https://github.com/bielenki/pMedia/blob/main/Fig/fig8.png?raw=true)

Figura 8: Exemplo de arquivo xls com os resultados

**Script IDW**

Os dados necessários para a utilização deste script são os mesmos utilizados no
script de Thiessen Variáveis, com as mesmas considerações.

Apenas devemos atentar para o fato de que se deve acrescentar um campo do tipo
“*double*” e chamado de “prec” à *feature* *class* das estações pluviométricas.

![fig1](https://github.com/bielenki/pMedia/blob/main/Fig/fig9.png?raw=true)

Figura 9: Campo “prec” adicionado à tabela de atributos da *feature* *class* das
estações

Para rodar a aplicação procede-se de forma análoga a outra ferramenta, atentando
para o fato que agora também se deve indicar um campo da tabela de atributos da
*feature* *class* que delimita as zonas para o cálculo da média que seja um
identificador único destas zonas.

![fig1](https://github.com/bielenki/pMedia/blob/main/Fig/fig10.png?raw=true)

Figura 10: Janela do *script* de Interpolação IDW

Os resultados da série de precipitação média serão gravados no arquivo xls
indicado como arquivo de saída, sendo que para cada zona terá sido criada uma
planilha na pasta Excel (com o nome do identificador da zona).

![fig1](https://github.com/bielenki/pMedia/blob/main/Fig/fig11.png?raw=true)

Figura 11: Exemplo de resultado da interpolação IDW para bacia dividida em duas
zonas

Quaisquer dúvidas, sugestões ou problemas na utilização da *ToolBox* por favor
contatar bielenki@ufscar.br

**Cláudio Bielenki Júnior**

**Especialista em Geoprocessamento – DHB – UFSCar**
