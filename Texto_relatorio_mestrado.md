**2.1 Considerações
sobre o funcionamento da Global Terrorism Database:**

O objetivo
deste capítulo não é avaliar os parâmetros elencados pelo Global Terrorism
Database (GTD) para possíveis definições de ataque terrorista. Assim, apesar de
os ter em “valor facial”, expõem dados importantes para entender os resultados
do GTD, muitas vezes díspares das pesquisas feitas nos arquivos de OCS. Tal
como foi o caso dos artigos mencionados na revisão bibliográfica. A explicação
parte do *Codebook: Methodology, Inclusion, Criteria and Variables*, o
manual desta base de dados.

De acordo
com o *Codebook,* estão registados no GTD todos os ataques terroristas
perpetrados a nível mundial desde 1970 até junho de 2020, num total de 209.705
incidentes registados, com ou sem sucesso. 
Os autores reconhecem que os investigadores deste fenómeno e outros
tipos de público podem ter diferentes definições de "terrorismo". A
abordagem escolhida foi recolher e estruturar os dados de forma que fossem
úteis à audiência mais alargada possível. A seleção inclui mecanismos de
filtragem para os utilizadores escolham os parâmetros que melhor se adequem ao
seu trabalho.

Um ataque
terrorista, no GTD, é definido como a ameaça ou o uso efetivo de força ilegal e
violência por um “ator não estatal” para alcançar um objetivo político,
económico-sistémico, religioso e social, através do medo, coerção ou
intimidação. Os ataques incluídos têm de verificar três condições. A primeira
condição é que o incidente é intencional, resultado de um “cálculo consciente”
do perpetrador. A segunda, é que deve conter algum grau de violência, incluindo
pessoas ou propriedade. Por último, os atacantes têm de ser atores
“subnacionais”, a base de dados exclui atos de “terrorismo Estatal”.

Para além
destas características, o ataque tem de obedecer a pelo menos dois de três
critérios. O objetivo final do ato tem de ser político, económico, religioso ou
social. Em segundo, têm de existir evidências da intenção de coagir, intimidar
ou transmitir uma mensagem a uma audiência mais alargada do que as vítimas
diretas. É a totalidade do ato que é tida em conta e o ato é considerado
intencional se pelo menos um dos planeadores teve as intenções já descritas.
Por último, deve ser realizado fora das “atividades legitimas de guerra”,
excluído do direito internacional humanitário, na medida em que alveja pessoas
não-combatentes. 

As
variáveis categóricas permitem ao utilizador registar quais os critérios de
inclusão para além dos necessários, filtrando os incidentes que não cumprem o
seu critério. É possível adicionar parâmetros que ajudem na pesquisa, que
incluem ataques em que restam dúvidas quanto aos critérios anteriores. Existe
ainda o mecanismo de filtragem adicional “Dúvida sobre Terrorismo Propriamente
Dito?”. Este é usado quando há sobreposições de definições entre terrorismo e
outros crimes e violência política, como insurgências, crimes de ódio e crime
organizado, nas fontes analisadas pelo GTD. 

O GTD
exclui planos ou conspirações que não foram tentados ou concretizados. No
entanto, inclui ataques falhados, o sucesso de um acidente é determinado de
acordo com algumas variáveis. Os incidentes que ocorreram no mesmo lugar e
intervalo temporal são considerados um único ataque, se o sítio ou o tempo de
ocorrência for descontinuo, é coliderado ataques separados. O GTD dá como
exemplo quatro camiões-bomba explodirem quase simultaneamente em diferentes
lugares de uma grande cidade, o que resulta em quatro incidentes.

Todos os
incidentes são descritos, incluindo “quando, onde, quem, o quê, como e porquê”.
Alguns dos filtros são interdependentes e assentam em múltiplas variáveis. Por
exemplo, o sucesso de um ataque depende da arma e das vítimas causadas. O GTD
elenca os ataques incluindo assassinato, sequestro, bombardeamento/explosão,
ataque armado, ataque desarmado, entre outros. Um ataque desarmado causa danos
físicos ou morte através de meios que não sejam explosivos, armas de fogo e
armas brancas, mas pode envolver armas químicas, biológicas ou
radiológicas/bombas sujas. 

A
informação sobre a vítima/alvo inclui o tipo de vítima, nacionalidade e nome da
entidade atacada. Entre os 22 parâmetros estão os negócios, governo, (clínicas
de) aborto, instituição de ensino, (estruturas de) fornecimento de comida ou
água, jornalistas e comunicação social, cidadãos e propriedades privadas, alvos
religiosos, turistas, serviços públicos e partidos políticos violentos. Por sua
vez, cada variável tem certas especificidades. O GTD contabiliza ainda alvos
humanos e propriedade militar. E "militares envolvidos em funções de
policiamento" e operações de manutenção de paz.

Sobre a
autoria do ataque, são registadas informações até três perpetradores, incluindo
nome do grupo, alegações de responsabilidade e o número de terroristas
participantes. Visto que os dados são provenientes de OCS de código aberto, não
implica uma afirmação legal de culpabilidade. Se as fontes dos dados não
identificam um grupo, são referidas informações sobre a identidade genérica do
atacante contextualizando o incidente. Estes dados não representam entidades
distintas, não são mutuamente exclusivas e não caracterizam o comportamento de
uma população ou movimento ideológico, adverte o manual.

**2.2 Método
de recolha de dados do arquivo da Lusa e do GTD:**

O arquivo
da Lusa está disponível no software e no website da agência, desenvolvidos pela
empresa *News Asset*. Os ficheiros foram transferidos através do *website*,
em pesquisas por palavras-chave no parâmetro “Texto”. Foram recolhidas todos as
notícias publicadas no campo “Tema” com “terrorismo” e todas as publicações com
as palavras da família de “terrorismo”, “extremismo” e “radicalismo” no corpo
da peça. No entanto, apesar de referir o número total de notícias em que estas
palavras-chave existem, o foco incide nos títulos, a única informação passível
de extrair de forma massiva.

O programa
de análise foi criado com recurso ao *Chat GPT Plus*. As instruções foram
escritas com os objetivos em consideração, mas sem revelar diretamente de que
se tratava de um trabalho académico, em formato de relatório ou tese. No
entanto, o modelo de IA generativo rapidamente "induziu" os
objetivos. Também foram dados ficheiros CVS do arquivo da Lusa e do GTD para a
análise das palavras necessárias para escrever o código o programa, como foi o
caso das listas descritas. 

Neste
caso, o objetivo foi criar um programa com todas as variáveis de pesquisa
necessárias, capaz de cruzar a informação dos títulos com os dados do GTD.
Igualmente, capaz de produzir gráficos ilustrativos dos cruzamentos e de
exportar as tabelas dos resultados em formato CVS, compatíveis com o Excel. O
programa foi codificado *Power Shell*, em “linguagem” *Python***,**
com a inclusão do *DuckDB* para processar os ficheiros CVS. E com o *layout*,
organização e exposição no formato da *app* *Streamlit*, por sugestão
do Chat GPT. 

Os dados
provenientes do GTD e os títulos e informação adjacente do arquivo da Lusa
foram transferidos no formato original, em CVS. O CVS do GTD inclui 209.706
linhas, cada uma corresponde a um incidente registado. Já o programa **DuckDB** é
utilizado para “contabilizar” cerca de 129.456 linhas, cada uma correspondendo
a uma notícia, que pode estar repetida, por ter sido publicada mais do que uma
vez no arquivo da Lusa. 

O programa
funciona em três camadas, todas concebidas com recurso ao Chat GPT. Na camada
mais simples estão as duas bases de dados já processados com recurso ao **Duck
Db**. Este é um sistema de processamento de bases de dados em
"relações", ou seja, tabelas. Na segunda camada encontra-se o
"script", ou seja, o conjunto de instruções que transforma os
ficheiros originais CVS em tabelas analíticas. Por último, está a camada
visível, a *interface* **Streamlit**, que permite filtrar, pesquisar e
visualizar os resultados, que permite pesquisar, filtrar e visualizar os dados.

De forma a
obter resultados diversos e o mais rigorosos possível, as ferramentas de
pesquisa e filtros têm diferentes variáveis. Entre os diferentes parâmetros
está o intervalo temporal, desde janeiro de 1987 até à primeira quinzena de
dezembro de 2025, excetuando alguns meses em três períodos diferentes, alturas
em que o sistema operativo do arquivo da Lusa foi alterado. É capaz de filtrar
por país e tipo de ação, sendo possível, também, pesquisar palavras soltas.
Contém a variável de pesquisa das palavras-chave apenas no título, ou no título
e nas palavras-chave extraídas do texto. 

O filtro
“violência” permite ver apenas os títulos que referem violência, excluir
violência. Também é possível filtrar por resultado "falhado",
"consumado" e "ameaça". Assim como incluir um ou mais tipo
de ataque violento, "bomba/explosão", "tiroteio",
"ataque armado", "esfaqueamento",
"sequestro/rapto", "atropelamento/veículo",
"incêndio/arson", "granada/morteiro" e "ataque
químico". Outra variável de pesquisa é a “Política” com os parâmetros
"nacional", "externa", "só política" e
"excluir política".

Já que as
famílias lexicais das palavras extremis\*, radical\* e terroris\* foram tratadas
na política nacional e internacional, foi criada uma ferramenta capaz de
analisar este tema separadamente, ou não. O código do programa deteta se uma
notícia é sobre política, através da presença de certas palavras nos títulos.
Para a política nacional foram listados os partidos políticos e com palavras
ligadas à política portuguesa, como governo, legislativas,
presidenciais, autárquicas, assembleia da república, ministro / ministra, entre
outras. 

Para a
política internacional a lista contém palavras ligadas às relações
internacionais, como “diplomacia/diplomático”, “política externa”,
“negociação”, “cimeira”, “sanções”, “embaixada”, “consulado”, “acordo”,
“tratado”, “cessar-fogo”, “processo de paz” e “mediação”. A lista de
organizações e instituições internacionais inclui “União Europeia”, “ONU/Nações
Unidas”, “NATO/OTAN”, “G7”, “G20”, “BRICS”. Se uma das palavras listadas surgir
no título, o programa seleciona-a como podendo ser sobre política portuguesa
e/ou externa. 

O filtro
de “autor” classifica o sujeito do título como, "judicial",
"Estado-governo", "Estado-forças armadas",
"Estado-forças de segurança", "grupo não Estatal",
"ONG/sociedade civil", "universidades/instituições
académicas", "indivíduo" e "autor omitido/evento".
Identificação é semelhante à anterior, cada categoria tem uma lista com
sujeitos institucionais ou profissionais. Inclui listados os termos “ministério
público”, “força aérea”, “GNR”, “*Human Rights Watch”,* “Al-Qaeda”,
“organização terrorista” e “grupo armado”. Para indivíduos, o filtro conta com
“suspeito”, “atacante” e “autor”.

No
programa desenvolvido, as famílias lexicais são automaticamente identificadas.
Conta individualmente as palavras “terror”, “terrorismo”, “terrorista”,
“terroristas”, “radicalismo”, “radical”, “radicalização”, “extremismo”,
“extremista” e “extremistas”. O cruzamento dos títulos da Lusa com os eventos
do GTD é feito através do país e não da cidade presente no título, já que nem
sempre são mencionadas. A Lusa e GTD usam formatos ISO diferentes,
identificador de países. Assim foi preciso compatibilizar esta secção do
código, para cruzar corretamente os dados.

O processo
é consistido por três etapas distintas. Nos títulos da Lusa o país não é
extraído diretamente de uma coluna, é automaticamente retirado do texto. Desta
maneira, foi incluído no programa um *script –* secção de código – que
identifica os países no título e cria um registo ISO. Por exemplo, existe um
título em que o país não é diretamente identificado, “Explosão em Bagdade mata
10 pessoas”, mas que referencia a capital do Iraque. Em segundo, o país é
convertido para ISO-2, como Portugal- PT, Iraque-IQ. 

Por
último, foi criado um dicionário de conversão do sistema ISO-2 para o ISO-3, o
código utilizado pelo GTD. Desta forma, 56 países coincidem nas duas bases e o
resto das correspondências possíveis são dia e país. Algumas grafias dos países
foram traduzidos ou normalizadas. 
Igualmente foram tratados casos de Estados como a Alemanha Ocidental,
Jugoslávia, Kosovo, West Bank/Cisjordânia, diretamente pelo Chat GPT. O
cruzamento entre os dois grupos de dados é feito através da data e país, sendo
que cada linha da tabela representa um evento num determinado dia e país.

**2.3 Visualização:**

O programa
mostra em tabela o número e os títulos em que certa palavra aparece. O gráfico
temporal é interativo, mostrando os resultados do cruzamento. As comparações
dos dados da Lusa e do GTD são visíveis nas tabelas e gráficos. As tabelas
mostram o número de notícias, de eventos no GTD, mortos e feridos, sucesso do
ataque e a identificação dos perpetradores. A *interface* reúne o
intervalo temporal, campo de escrita livre, listas de palavras-chave
pré-definidas, escolha de *AND / OR,* pesquisa apenas no título, e
pesquisa com inclusão ou exclusão de termos das famílias lexicais definidas.