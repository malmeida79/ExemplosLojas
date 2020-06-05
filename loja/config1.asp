<body bgcolor="#FFFFFF"><%
'       =================================================================
'        C O N F I G U R A � � E S   G E R A I S   D A   S U A   L O J A 
'                    Altere todos os dados que desejar                   
'																		 
'							I N D I C E									 
'																		 
'		1.	CONFIGURA��O DE BANCO DE DADOS E TIPO DE CONEX�O			 
'		2.	CONFIGURA��O DE FRETE										 
'		3.	CONFIGURA��ES GERAIS										 
'	 	4.	CONFIGURA��ES DO BOLETO BANC�RIO							 
'		5.	CONFIGURA��ES DAS CORES DA SUA LOJA							 
'																		 
'       =================================================================
VersaoDb = 2

'StringdeConexao = "compuware"

'#########################################################################################################
'	2.	CONFIGURA��O DE FRETE																			  
'#########################################################################################################


'***	2.1  ENTREGA NACIONAL OU ESTADUAL (VIA CORREIOS)


'## Colocar a Forma de Entrega dos pedidos via Sedex?
entrega_sedex="Sim"	'Deixe "Sim" ou "Nao"

seu_cep = "09607-020"		'Digite o CEP de sua loja usando somente numeros

'## Voc� deseja testar Offline e sem Conexao Internet?
sem_calculo_online_SiteCorreios = "Sim" 
'IMPORTANTE: Deixe "Sim" caso voce N�O esteja conectado na Internet (para nao travar Calc. de Frete), 
'pois o c�lculo vem do site dos Correios, mas setando para "Sim" ser� colocado um valor fict�cio de R$ 10,00

'## Servi�o Adicional dos Correios para entrega do pedido SOMENTE para o Cliente, digite "S" para ativar ou "N" para n�o ativar
mao_propria = "N"			

'## Servi�o Adicional dos Correios de Aviso de Recebimento da encomenda na casa do Cliente, digite "S" para ativar ou "N" para n�o ativar
aviso_recebimento = "N"		

'## Servi�o de Seguro do pr�prio Correios da encomenda do Cliente, digite  "S" para ativar ou "N" para n�o ativar
cobrar_seguro_produto = "N"		
'Obs 1: Quando o cliente opta por Sedex a Cobrar, este valor � cobrado obrigatoriamente pelos Correios
'Obs 2: O seguro � 1% do valor da compra

'## Colocar a Op��o de Pagamento do pedido via Sedex a Cobrar?
entrega_sedex_a_cobrar="Nao"	'Deixe "Sim" ou "Nao"



'***	2.2  ENTREGA REGIONAL (Via Motoboy, Tele-Entrega [com Carro pr�prio] ou similares que s�o cobrados Taxas Fixas)

'## Colocar a Forma de Entrega dos pedidos via Motoboy?
entrega_motoboy="Nao"	'Deixe "Sim" ou "Nao"

'## Descricao da Forma de Entrega dos pedidos a Domicilio 
descricao_forma_entrega_motoboy = "Delivery" '(Por padr�o est� como Motoboy, mas pode ser outro como "Tele-Entrega" , etc...)

'## Tarifa para Entrega dos pedidos via Motoboy (se valor com CENTAVOS, digite COM PONTO - Exemplo: 4.50 )
tarifa_entrega_motoboy = 20.00

'## Descricao da Regiao da Entrega dos pedidos via Motoboy
descricao_regiao_entrega_motoboy = "Entrega via Delivery em somente em S�o Jos� dos Campos"



'***	2.3  ENTREGA PESSOAL (Via Download, Pessoalmente ou similares que N�O s�o cobrados)

'## Colocar a Forma de Entrega dos pedidos via Download?
entrega_download="Nao"	'Deixe "Sim" ou "Nao" 
'IMPORTANTE: Caso deixe Sim, por questoes de logica DESATIVE as formas de entrega anteriores ( entrega_motoboy="Nao" e entrega_sedex="Nao")

'## Descricao da Forma de Entrega dos pedidos a Domicilio 
descricao_forma_entrega_download = "Download" '(Por padr�o est� como Download, mas pode ser outro como "Pessoalmente" , etc...)

'## Tarifa para Entrega dos pedidos via Download
tarifa_entrega_download = 0

'## Descricao da Regiao da Entrega dos pedidos via Download
descricao_regiao_entrega_download = "Voc� poder� fazer o Download direto do site!"


'---> ENCOMENDA NORMAL IMPLEMENTADO POR MATEUS MORAES (jusivansjc@yahoo.com.br)

'***	2.4  ENCOMENDA NORMAL

'## Colocar a Forma de Entrega dos pedidos via Encomenda Normal?
entrega_encomenda="Nao"	'Deixe "Sim" ou "Nao" 

'## Descri��o da Forma de Entrega dos pedidos a Domicilio 
descricao_forma_entrega_encomenda = "Encomenda Normal"

'## Configura��o das Regi�es para frete via Encomenda Normal (Modelo Padr�o configurado para o PARAN� - CAPITAL)

' 1. Configure o Estado em que sua loja se encontra para ZERO
' 2. Configure os demais Estados de acordo com o arq em anexo (ENCOMENDA_NORMAL.xls), ou pelas tabelas disponibilizadas
' pelos Correios: http://www.correios.com.br/servicos/precos_tarifas/nacionais/tar_encomenda_normal.cfm

' 3. D� uma olhada em http://www.correios.com.br/servicos/precos_tarifas/nacionais/pdf/ENC_NORMAL_PR.pdf
' para entender melhor a configura��o desse tipo de frete. No caso do Paran�, por ser LOCAL (conforme o PDF acima),
' ele leva o valor 0 no valor abaixo, os da 2a coluna leva o valor 1  e assim por diante

estado_sp = 0

estado_mg = 1
estado_pr = 1
estado_rj = 1
estado_sc = 1


estado_df = 2
estado_es = 2
estado_go = 2
estado_ms = 2
estado_rs = 2

estado_ba = 3
estado_mt = 3
estado_to = 3

estado_al = 4
estado_se = 4

estado_ce = 5
estado_ma = 5
estado_pb = 5
estado_pe = 5
estado_pi = 5
estado_rn = 5
estado_ro = 5

estado_ac = 6
estado_ap = 6
estado_pa = 6

estado_am = 7

estado_rr = 8
'###################################################################################
'	3.	CONFIGURA��ES GERAIS	
'###################################################################################

'## Escreva o nome do administrador da loja para suporte via Administrador
Seu_nome = "Jo�o Paulo Moreira"

'## Voc� deseja Usar Criptografia de 64 BITS em sua loja?
Cripto_Ativa = "Sim" 'Deixe "Sim" para usar ou ent�o escreva "Nao"

'## Voc� deseja mostra 4, 6 ou 8 produtos no inicio da sua loja?
mostrar_produtos_fachada=6		'Digite 4 ou 6 ou 8

'## Voc� deseja colocar Destaque de Produto na p�gina inicial? (produto aleat�rio)
mostrar_produto_destaque_fachada="Sim" 'Deixe "Sim" para usar ou ent�o escreva "Nao"

'## Voc� deseja mostra qtos produtos por p�gina (paginas Secoes e Pesquisa)?
mostrar_produtos_por_pagina_na_secao=8

'## Voc� deseja mostrar o aviso de estoque "Esgotado" automaticamente , qdo acabar o estoque do Produto?
'Se Sim, o produto ficar� "Vis�vel", mas aparecer� um aviso de estoque "Esgotado". Este aviso n�o aparecer� se o produto j� estava com o status de "N�o Vis�vel"
'Se Nao, o produto ficar� com o status de "N�o Vis�vel" ao p�blico, qdo acabar o estoque do Produto. 
mostrar_produto_esgotado="Sim" 'Deixe "Sim" ou "Nao"

'## Voc� deseja mostrar as Op��es de Cart�es de Creditos como Forma de Pgto?
mostrar_pgtos_com_cartoescredito="Sim" 'Deixe "Sim" para usar ou ent�o escreva "Nao"

'## Voc� deseja limitar a compra de um produto conforme a qtidade M�xima fixada na �rea administrativa?
mostrar_limite_max_prod_compra="Nao" 'Deixe "Sim" ou "Nao"
'Em caso de Nao, pode acontecer de o cliente comprar mais do que tem no estoque, portanto poder� aparecer para o Administrador estoques "negativos"

'## Voc� deseja mostrar o Formul�rio para o Cliente solicitar por email produtos que est�o como "Esgotados"?
mostrar_form_email_prod_esgotado="Sim" 'Deixe "Sim" para usar ou ent�o escreva "Nao"

'## Voc� deseja mostrar o link de Pol�tica de Trocas no Quadro Atendimento?
mostrar_politica_de_trocas="Sim" 'Deixe "Sim" ou "Nao"

'## Nome da imagem da Logo da Loja (coloque dentro da Pasta /Imagens/ de cada linguagem)
arquivo_logo_loja="logo1.gif"

'## Dados da loja	
Nome_da_sua_loja = "Compuware Technologies5555"
CNPJ_da_sua_empresa = "00.000.000/0001-00"
Razao_Social = "Compuware Technologies"
Email_da_sua_loja = "contato@compuwaretecno.com.br"
Slogan_da_sua_loja = "SLOGAN DA SUA EMPRESA"
Endereco_do_site = "www.compuwaretecno.com.br" 
         'ATEN��O: Escreva o endere�o virtual da sua loja sem o Http:// e sem o / no final

Endereco_da_sua_loja = "Av. Onze de Agosto , N�79 "
Bairro_da_sua_loja = "Jardim Silvestre"
Cidade_da_sua_loja = "S�o Bernardo do Campo "
Estado_da_sua_loja = "S�o Paulo"
Pais_da_sua_loja = "Brasil"
Telefone_da_sua_loja = "(011) 4367-1415"

'## Dados bancarios - Banco 1
Seu_banco = "Seu Banco"
Sua_agencia = "Sua Agenia"
Sua_conta_poupanca = "Sua conta"
Nome_do_cedente = "Seu Nome"

'## Dados bancarios - Banco 2		( 'Se nao for usar , deixe os 4 campos abaixo como "" )
Seu_banco2 = "" 
Sua_agencia2 = ""
Sua_conta_corrente2 = ""
Nome_do_cedente2 = ""

'## Confirura��es de e-mail
Host_da_loja = "smtp.compuwaretecno.com.br" ' --> Se o componente usado for CDONTS deixe esse campo assim: "-"

'## Escolha o componente e desmarque o coment�rio
Componente_usado = "AspEmail"
'Componente_usado = "AspMail"
'Componente_usado = "AspQmail"
'Componente_usado = "CDONTS"
'Componente_usado = "JMail"

'## Banners da Loja
Banner_AdMentor="Sim"	'--> Coloque "Sim" se for usar Banners via Sistema Admentor, da Area Administrativa, ou deixe como ""
Banner_Fixo="" '--> Para usar um banner fixo, coloque o nome da imagem entre as aspas, coloque a imagem na pasta Banners e deixe Banner_AdMentor=""

'## Fonte da loja
Fontes_da_loja = "tahoma,verdana,arial,helvetica"

'## Dias para entrega dos pedidos
Dias_para_entrega = 5

'## Mostrar Quadro Linguagens na Loja
mostrar_quadro_linguagens="Nao"	'--> 'Deixe "Sim" ou "Nao"

'## Mostrar Contador na Loja
mostrar_contador="Nao<strong></strong>"	'--> Coloque "Sim" se quiser mostrar

'## Tipo da Moeda Corrente
strLgMoeda = "R$"  ' REAL - R$ - Brasil
strLgMoedax = "Reais"  ' Nome da moeda no plural




'###################################################################################
'	4.	CONFIGURA��ES DO BOLETO BANC�RIO
'###################################################################################

'## Voc� deseja colocar a op��o Boleto como Forma de Pagamento?
mostrar_pagamento_por_boleto="Sim" 'Deixe "Sim" para usar ou ent�o escreva "Nao"

'caminho_boleto = "boleto_hsbc.asp" ' use este caminho para utilizar o boleto do HSBC
caminho_boleto = "boleto_bb1.asp" ' use este caminho para utilizar o boleto do BANCO DO BRASIL
'caminho_boleto = "boleto_bradesco.asp" ' use este caminho para utilizar o boleto do BRADESCO
'caminho_boleto = "boleto_itau.asp" ' use este caminho para utilizar o boleto do ITA�

'## Informa��es banc�rias referente ao seu Boleto  ( Alguns campos s�o opcionais )
bol_nr_cedente="2338-5/159225-0" 'numero do seu codigo cedente (n�o � necess�rio para alguns bancos)
bol_agencia = "2578" 'agencia
bol_dagencia = "5" 'digito da agencia (n�o � necess�rio para alguns bancos)
bol_conta = "153385" 'conta (n�o � necess�rio para alguns bancos)
bol_dconta = "0" 'digito da conta (n�o � necess�rio para alguns bancos)
bol_carteira = "17" 'c�digo da carteira
bol_convenio = "1047849" 'numero do seu convenio (n�o � necess�rio para alguns bancos)
bol_cedente = "bondshop - Uma nova op��o em Loja Virtual na Web!" 'Geralmente � o nome da loja e o nome do cedente no banco

'## Observa��es no Boleto  ( Campos opcionas )
' Voce pode colocar alguma instru��o para o cliente, por exemplo "Obrigado por comprar em nossa loja"

obs_linha1 = "Para agilizar o seu atendimento, envie-nos o seu email ap�s o seu pagamento para "&Email_da_sua_loja ' Observa��o Linha 1 
obs_linha2 = "informando apenas o n�mero do seu pedido e o banco usado para pagamento." ' Observa��o Linha 2
obs_linha3 = "A "& Nome_da_sua_loja &" agradece o seu pedido." ' Observa��o Linha 3
obs_linha4 = "" ' Observa��o Linha 4
obs_linha5 = "" ' Observa��o Linha 5


'## Instru��es ao Caixa do Banco (Por exemplo "N�o Receber ap�s o vencimento" ou "Cobrar multa de 2 % ap�s o vencimento")
instr_linha1 = "PAG�VEL EM QUALQUER BANCO"  ' Instrucoes Linha 1
instr_linha2 = "N�O ACEITAR AP�S VENCIMENTO" ' Instrucoes Linha 2
instr_linha3 = "" ' Instrucoes Linha 3
instr_linha4 = ""  ' Instrucoes Linha 4
instr_linha5 = ""  ' Instrucoes Linha 5


'###################################################################################
'	5.	CONFIGURA��ES DAS CORES DA SUA LOJA
'###################################################################################

'Dica: Acesse o arquivo tabela_de_cores.asp para ver a Tabela de Cores e os respectivos c�digos hexadecimais


'Cor do fundo na resolu��o 1024x768
Cor_de_fundo = "#ebefef"

'Cor principal da loja
Cor_principal = "#ffffff"

'Cor das linhas
Cor_das_linhas = "#cccccc"

'Cor da borda do contador de visitas
Cor_do_Contador = "#fcc701"

'Cor da barra de menu e copyright
Cor_do_menu = "#c0c0c0"

'Cor principal dos links
Cor_dos_links ="#000000"

'Cor secundaria dos links
Cor_dos_links_secundarios ="#808080"

'Cor do texto dos menus
Texto_dos_menus = "#ffffff"

'Cor do texto da loja
Texto_da_loja = "#000000"

'Cor da borda do carrinho de compras
Cor_da_borda_carrinho_de_compras = "#000000"

'Cor do texto do menu de navegacao de fechamento de compras
Cor_texto_menu_fechamento_compras = "#ff4500" 'Exemplo: #0000ff ou #ff4500

'##################################################################



'####################################################################
'																	 
'     NAO ALTERE ESSA PARTE, VOC� PODER� ESTRAGAR SUA LOJA !		 
'																	 
'####################################################################


    'Inicia as variaveis da Aplica��o 
 	application("razaoloja") = Razao_Social
    application("nomeloja") = Nome_da_sua_loja
    application("emailloja") = Email_da_sua_loja
   	application("slogan") = Slogan_da_sua_loja
   	application("urlloja") = Endereco_do_site
   	application("endereco11") = Endereco_da_sua_loja
   	application("bairro11") = Bairro_da_sua_loja
   	application("cidade11") = Cidade_da_sua_loja
   	application("estado11") = Estado_da_sua_loja
   	application("pais11") = Pais_da_sua_loja
   	application("fone11") = Telefone_da_sua_loja
    application("CORREIOSseucep11") = seu_cep 
    application("CORREIOSmaop11") = mao_propria 
    application("CORREIOSaviso11") = aviso_recebimento 
   	application("bancopag") = Seu_banco
   	application("contapag") = Sua_conta_poupanca
   	application("pagpara") = Nome_do_cedente
	application("agpag") = Sua_agencia
   	application("bancopag2") = Seu_banco2
   	application("contapag2") = Sua_conta_corrente2
   	application("pagpara2") = Nome_do_cedente2
	application("agpag2") = Sua_agencia2
	application("fonte") = Fontes_da_loja
	application("maopropria") = Preco_da_mao_propria
	application("diasentrega") = Dias_para_entrega
	application("corfundo") = Cor_de_fundo
	application("cor1") = Cor_das_linhas
	application("cor2") = Cor_do_menu
	application("cor3") = Cor_principal
	application("cor4") = Cor_dos_links
	application("cor5") = Cor_dos_links_secundarios
    application("cor6") = Cor_do_Contador
	application("fontebranca") = Texto_dos_menus
	application("fontepreta") = Texto_da_loja
	application("corborda") = Cor_da_borda_carrinho_de_compras
    application("nomecontato") = Seu_nome
	application("HostLoja") = Host_da_loja
    application("ComponenteLoja") = Componente_usado
	application("Cripto_Ativa") = Cripto_Ativa


'>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<

    Server.ScriptTimeOut = 60		'*** Evite usar mais do que 60 segundos
	application("vsversao") = "OPEN 3.0 - Pack 5"
	application("ultimaatualizacao")  = "Sexta, 07 de Maio de 2006"
	application("link_comunidade")  = "<a class=menu  href=""http://www.studio7seven.com"" target=_blank>Desenvolvimento de Web Site</a>"
	set abredb = Server.CreateObject("ADODB.Connection")
	abredb.Open(StringdeConexao)

'>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<

'Executa a fun��o GLOBAL

set rsglobal = abredb.execute("SELECT * FROM compras Order by idcompra desc")
if rsglobal.eof then
Session("orderID") = ""
end if
rsglobal.close
set rsglobal = nothing

if Session("orderID") = "" then
set rs = abredb.execute("SELECT * FROM compras Order by idcompra desc")
	if rs.eof then
	Session("orderID") = 1
	abredb.Execute("INSERT INTO compras (idcompra, datacompra, status) values ('" &session("orderID") & "', '" & date & "', 'Compra em Aberto')")
	else
	Session("orderID") = rs("idcompra") + 1
	abredb.Execute("INSERT INTO compras (idcompra, datacompra, status) values (" &session("orderID") & ", '" & date & "','Compra em Aberto')")
	end if
session.timeout = 120
else
end if

'###################################################################################
'## ARQUIVOS DE LINGUAGEM DA LOJA
'###################################################################################

'###################################################################################
'### Mostra a linguagem padr�o dependendo do navegador
'###################################################################################

IF Escolha <> "" then
 Response.Cookies("VirtuaStore").Expires 	=  DateAdd("d",365,date()) 'VENCE NO PROXIMO ANO
 Response.Cookies("VirtuaStore")("Linguagem")	= Escolha
End IF

IF Request("Lan") <> "" THEN
	Response.cookies("VirtuaStore")("Linguagem")= ""
	Response.cookies("VirtuaStore")("Linguagem")= Request("Lan")
END IF

strReporLingua = "?"&replace(replace(replace(replace(replace(Request.ServerVariables("Query_String") ,"&Lan=Espanol",""),"&Lan=Portugal",""),"&Lan=Brasil",""),"&Lan=English","")," ","")
strReporLingua = "?"&replace(replace(replace(replace(replace(Request.ServerVariables("Query_String") ,"Lan=Espanol",""),"Lan=Portugal",""),"Lan=Brasil",""),"Lan=English","")," ","")

SELECT CASE Request.cookies("VirtuaStore")("Linguagem")

	CASE "Brasil"
	dirlingua = "linguagens/portuguesbr"
	OpcaoLingua = "<a href='"& strReporLingua&"&Lan=English'><img src=linguagens/english.gif border=0></a>"&_
		      "<a href='"& strReporLingua&"&Lan=Portugal'><img src=linguagens/portugal.gif border=0></a>"&_
		      "<a href='"& strReporLingua&"&Lan=Espanol'><img src=linguagens/espanha.gif border=0></a>"
	%><!--#include file="linguagens/portuguesbr/linguagem.asp"--><%
	
	CASE "English"
	dirlingua = "linguagens/ingles"
	OpcaoLingua = "<a href='"& strReporLingua&"&Lan=Brasil'><img src=linguagens/brasil.gif border=0></a>"&_
		      "<a href='"& strReporLingua&"&Lan=Portugal'><img src=linguagens/portugal.gif border=0></a>"&_
		      "<a href='"& strReporLingua&"&Lan=Espanol'><img src=linguagens/espanha.gif border=0></a>"
	%><!--#include file="linguagens/ingles/linguagem.asp"--><%
	
	CASE "Portugal"
	dirlingua = "linguagens/portuguespt"
	OpcaoLingua = "<a href='"& strReporLingua&"&Lan=Brasil'><img src=linguagens/brasil.gif border=0></a>"&_
		      "<a href='"& strReporLingua &"&Lan=English'><img src=linguagens/english.gif border=0></a>"&_
		      "<a href='"& strReporLingua &"&Lan=Espanol'><img src=linguagens/espanha.gif border=0></a>"
	%><!--#include file="linguagens/portuguespt/linguagem.asp"--><%
	
	CASE "Espanol"
	dirlingua = "linguagens/espanhol"
	OpcaoLingua = "<a href='"& strReporLingua&"&Lan=Brasil'><img src=linguagens/brasil.gif border=0></a>"&_
		      "<a href='"& strReporLingua&"&Lan=Portugal'><img src=linguagens/portugal.gif border=0></a>"&_
		      "<a href='"& strReporLingua&"&Lan=English'><img src=linguagens/english.gif border=0></a>"
	%><!--#include file="linguagens/espanhol/linguagem.asp"--><%
	
	CASE ELSE
	
		'-----------------------------------------------------------------
		'### BRASIL
		'-----------------------------------------------------------------
		IF Request.ServerVariables("HTTP_ACCEPT_LANGUAGE") = "pt-br" THEN
		dirlingua = "linguagens/portuguesbr"
		OpcaoLingua = "<a href='"& strReporLingua&"&Lan=English'><img src=linguagens/english.gif border=0></a>"&_
			      "<a href='"& strReporLingua&"&Lan=Portugal'><img src=linguagens/portugal.gif border=0></a>"&_
			      "<a href='"& strReporLingua&"&Lan=Espanol'><img src=linguagens/espanha.gif border=0></a>"
				%><!--#include file="linguagens/portuguesbr/linguagem.asp"--><%
				
		'-----------------------------------------------------------------
		'### PORTUGAL
		'-----------------------------------------------------------------
		ELSEIF Request.ServerVariables("HTTP_ACCEPT_LANGUAGE") = "pt" THEN
		dirlingua = "linguagens/portuguespt"
		OpcaoLingua = "<a href='"& strReporLingua&"&Lan=Brasil'><img src=linguagens/brasil.gif border=0></a>"&_
			      "<a href='"& strReporLingua&"&Lan=English'><img src=linguagens/english.gif border=0></a>"&_
			      "<a href='"& strReporLingua&"&Lan=Espanol'><img src=linguagens/espanha.gif border=0></a>"
				%><!--#include file="linguagens/portuguespt/linguagem.asp"--><%
				
		'-----------------------------------------------------------------
		'### ESPANHOL
		'-----------------------------------------------------------------
		ELSEIF Request.ServerVariables("HTTP_ACCEPT_LANGUAGE") = "es" THEN
		dirlingua = "linguagens/espanhol"
		OpcaoLingua = "<a href='"& strReporLingua&"&Lan=Brasil'><img src=linguagens/brasil.gif border=0></a>"&_
			      "<a href='"& strReporLingua&"&Lan=English'><img src=linguagens/english.gif border=0></a>"&_
			      "<a href='"& strReporLingua&"&Lan=Portugal'><img src=linguagens/portugal.gif border=0></a>"
				%><!--#include file="linguagens/portuguespt/linguagem.asp"--><%
				
		'-----------------------------------------------------------------
		'### INGLES TB PODE SER US
		'-----------------------------------------------------------------
		ELSE
		dirlingua = "linguagens/portuguesbr"
		OpcaoLingua = "<a href='"& strReporLingua&"&Lan=English'><img src=linguagens/english.gif border=0></a>"&_
			     "<a href='"& strReporLingua&"&Lan=Portugal'><img src=linguagens/portugal.gif border=0></a>"&_
			     "<a href='"& strReporLingua&"&Lan=Espanol'><img src=linguagens/espanha.gif border=0></a>"
				%><!--#include file="linguagens/portuguesbr/linguagem.asp"--><%
		END IF
END SELECT

OpcaoLingua=replace(OpcaoLingua,"?&","?")
OpcaoLingua=replace(OpcaoLingua,"&&","&")

'###################################################################################

'if aaaa="" then   '*** Ativar   
if aaaa<>"" then   '*** Desativar

response.write "<font size=1 face=arial>"
response.write "<br> Testando as sessoes...<br> "

response.write "<br> modo_entrega: "&session("modo_entrega")
response.write "<br> usuario: "&session("usuario")
response.write "<br> orderID: "&Session("orderID")
response.write "<br> pedido1: "&Session("pedido1")
response.write "<br> PesoTotalValor: "&session("PesoTotalValor")
response.write "<br> sql: "&session("sql")
response.write "<br> sql_max: "&session("sql_max")
response.write "<br> dados_gravados_compra: "&session("dados_gravados_compra")

response.write "</font>"
end if

%>
