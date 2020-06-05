<%
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################
'#
'#  C�DIGO: VirtuaStore Vers�o OPEN - Copyright 2001-2004 VirtuaStore
'#  URL: http://comunidade.virtuastore.com.br
'#  E-MAIL: comunidade@virtuastore.com.br
'#  AUTORES: Ot�vio Dias(Desenvolvedor)
'# 
'#     Este programa � um software livre; voc� pode redistribu�-lo e/ou 
'#     modific�-lo sob os termos do GNU General Public License como 
'#     publicado pela Free Software Foundation.
'#     � importante lembrar que qualquer altera��o feita no programa 
'#     deve ser informada e enviada para os criadores, atrav�s de e-mail 
'#     ou da p�gina de onde foi baixado o c�digo.
'#
'#  //-------------------------------------------------------------------------------------
'#  // LEIA COM ATEN��O: O software VirtuaStore OPEN deve conter as frases
'#  // "Powered by VirtuaStore OPEN" ou "Software derivado de VirtuaStore 1.2" e 
'#  // o link para o site http://comunidade.virtuastore.com.br no cr�ditos da 
'#  // sua loja virtual para ser utilizado gratuitamente. Se o link e/ou uma das 
'#  // frases n�o estiver presentes ou visiveis este software deixar� de ser
'#  // considerado Open-source(gratuito) e o uso sem autoriza��o ter� 
'#  // penalidades judiciais de acordo com as leis de pirataria de software.
'#  //--------------------------------------------------------------------------------------
'#      
'#     Este programa � distribu�do com a esperan�a de que lhe seja �til,
'#     por�m SEM NENHUMA GARANTIA. Veja a GNU General Public License
'#     abaixo (GNU Licen�a P�blica Geral) para mais detalhes.
'# 
'#     Voc� deve receber a c�pia da Licen�a GNU com este programa, 
'#     caso contr�rio escreva para
'#     Free Software Foundation, Inc., 59 Temple Place, Suite 330, 
'#     Boston, MA  02111-1307  USA
'# 
'#     Para enviar suas d�vidas, sugest�es e/ou contratar a VirtuaWorks 
'#     Internet Design entre em contato atrav�s do e-mail 
'#     contato@virtuaworks.com.br ou pelo endere�o abaixo: 
'#     Rua Ven�ncio Aires, 1001 - Niter�i - Canoas - RS - Brasil. Cep 92110-000.
'#
'#     Para ajuda e suporte acesse um dos sites abaixo:
'#     http://comunidade.virtuastore.com.br
'#     http://br.groups.yahoo.com/group/virtuastore
'#
'#     BOM PROVEITO!          
'#     Equipe VirtuaStore
'#     []'s
'#
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################

'IN�CIO DO C�DIGO---------------------
'Este c�digo est� otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP

'======================================
' ARQUIVO COM AS STRINGS PARA DEFINI��O DA LINGUAGEM
' Portugu�s (BR) - LCID 1046
' Pais : Brasil
' HTML Content-Type : ISO-8859-1 
' Traduzido por : Ot�vio Dias (otavio@virtuaworks.com.br ) 20/08/2002 
' Acr�scimos/Corre��es por : Elizeu Alcantara (elizeu@cristaosite.com.br / elizeu@onda.com.br ) 15/03/2004 
'======================================


nomeloja = application("nomeloja")

strLgDataS1 = "Jan"            													'Nome dos meses abreviados
strLgDataS2 = "Fev"
strLgDataS3 = "Mar"
strLgDataS4 = "Abr"
strLgDataS5 = "Mai"
strLgDataS6 = "Jun"
strLgDataS7 = "Jul"
strLgDataS8 = "Ago"
strLgDataS9 = "Set"
strLgDataS10 = "Out"
strLgDataS11 = "Nov"
strLgDataS12 = "Dez"

strLgData1 = "Janeiro"														 'Nome dos meses completos
strLgData2 = "Fevereiro"
strLgData3 = "Mar�o"
strLgData4 = "Abril"
strLgData5 = "Maio"
strLgData6 = "Junho"
strLgData7 = "Julho"
strLgData8 = "Agosto"
strLgData9 = "Setembro"
strLgData10 = "Outubro"
strLgData11 = "Novembro"
strLgData12 = "Dezembro"


'======================================
'	VALOR DAS STRINGS DA LOJA          
'======================================

strLg1 = "Finalizar Compras"
strLg2 = "Quantidade de itens:"
strLg3 = "Sub-total:"
strLg4 = "Home"
strLg5 = "Registro de Clientes"
strLg6 = "Login"
strLg7 = "Carrinho de Compras de"
strLg8 = "Fechar Pedido"
strLg9 = "Meus Dados"
strLg10 = "Hist�rico de Compras"
strLg11 = "Logout"
strLg12 = "Departamentos"
strLg13 = "Pesquisa"
strLg14 = "Atendimento"
strLg15 = "Todas as categorias"
strLg16 = "Como Comprar?"
strLg17 = "Por e-mail"
strLg18 = "Sobre Seguran�a"
strLg19 = "Seguran�a e Privacidade"
strLg20 = "Quem somos"
strLg21 = "Pesquisar"
strLg22 = "Sobre Seguran�a"
strLg23 = "Por e-mail"
strLg24 = "Atendimento por e-mail"
strLg25 = "para entrega"
strLg26 = "Pronta Entrega" ' ou "Dispon�vel para entrega"
strLg27 = "Esgotado" 'ou "N�o dispon�vel para entrega"
strLg28 = "Estoque:"
strLg29 = "Pre�o:"
strLg30 = "+ Detalhes"
strLg31 = "Fabricante:"
strLg32 = "Adicionar"
strLg33 = "produto(s) no meu carrinho de compras."
strLg34 = "Comprando o"
strLg35 = "na"
strLg36 = "voc� economiza at� " 
strLg37 = "� vista por:"
strLg38 = "por:"
strLg39 = "de:"
strLg40 = "ou ainda:"
strLg41 = "Voltar"
strLg42 = "Pesquisando por" 
strLg43 = "Foram encontrado(s)"
strLg44 = "produto(s) :"
strLg45 = "P�gina"
strLg46 = "de"
strLg47 = "Pr�xima P�gina"
strLg48 = "Produtos neste departamento:"
strLg49 = "N�o h� nenhum item em seu carrinho de compras."
strLg50 = "Se voc� comprou algum produto anteriormente sua sess�o de compras est� expirada!"
strLg51 = "Ver meu carrinho de compras"
strLg52 = "Mais Ofertas:"
strLg53 = "O item foi adicionado em seu carrinho de compras."
strLg54 = "Compra"
strLg55 = "Sua compra"
strLg56 = "Por favor, verifique se sua compra est� correta antes de prosseguir."
strLg57 = "Para modificar a quantidade de algum item digite o novo valor e clique no bot�o atualizar."
strLg58 = "Quantidade"
strLg59 = "Produto"
strLg60 = "Pre�o Unit�rio"
strLg61 = "Pre�o Total"
strLg62 = "Remover"
strLg63 = "Total da compra"
strLg64 = "Valor da entrega em"
strLg65 = "Sua compra"
strLg66 = "Atualizar"
strLg67 = "Prosseguir"
strLg68 = "Continuar Compras"
strLg69 = "Escolha a regi�o de entrega para o calculo do frete!"
strLg70 = "Clique aqui"
strLg71 = "Clique aqui!"
strLg72 = "Dados para entrega do pedido"
strLg73 = "Para modificar sua compra"
strLg74 = "Por favor, verifique se o endere�o para entrega est� correto antes de prosseguir."
strLg75 = "se endere�o para a entrega e o mesmo que o seu!"
strLg76 = "Op��es para entrega"
strLg77 = "Endere�o para entrega"
strLg78 = "Endere�o:"
strLg79 = "Bairro:"
strLg80 = "Cidade:"
strLg81 = "Estado:"
strLg82 = "CEP:"
strLg83 = "Pa�s:"
strLg84 = "Telefone:"
strLg85 = "Voc� deseja que os itens comprados sejam empacotados e enviados como presente?"
strLg86 = "Sim"
strLg87 = "N�o"
strLg88 = "Mensagem escrita no cart�o-presente:"
strLg89 = "Limpar"
strLg90 = "Por favor, preencha todos os dados corretamente antes de prosseguir."
strLg91 = "Forma de Pagamento"
strLg92 = "Por favor, preencha todos os dados corretamente antes de prosseguir."
strLg93 = "Para modificar o endere�o de entrega das compras e/ou outros dados"
strLg94 = "Selecione a forma de pagamento escolhida:"
strLg95 = "Outras formas de pagamento"
strLg96 = "Pagamento via Cart�o de Cr�dito" 
strLg97 = "Selecione seu cart�o:"

strLg98 = "N�mero do cart�o: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;C�digo de Seguran�a:"

strLg99 = "Expira��o do cart�o:" 
strLg100 = "Cancelar"
strLg101 = "Ano"
strLg102 = "Titular do cart�o (Como aparece no Cart�o):" 
strLg103 = "Meus Dados"
strLg104 = "Para modificar seus dados, selecione o campo desejado e escreva o dado atualizado sobre o antigo."
strLg105 = "Para nenhuma modifica��o clique em 'Cancelar'. "
strLg106 = "Dados do Registro"
strLg107 = "Dados Pessoais"
strLg108 = "Seu e-mail:"
strLg109 = "Seu telefone:"
strLg110 = "Seu pa�s:"
strLg111 = "Seu estado:"
strLg112 = "Sua cidade:"
strLg113 = "Seu bairro:"
strLg114 = "Seu endere�o:"
strLg115 = "Seu RG ou IE"
strLg116 = "Seu CPF ou CNPJ"
strLg117 = "Seu nascimento:"
strLg118 = "Seu Nome ou Raz�o Social:"
strLg119 = "Compras efetuadas por"
strLg120 = "Para visualizar os dados da sua compra digite o n�mero do pedido logo abaixo:"
strLg121 = "Pedido n.:"
strLg122 = "Para sua comodidade e seguran�a verifique todos os seus dados antes de prosseguir." 
strLg123 = "Sua Senha:"
strLg124 = "Seu e-mail foi enviado com sucesso!"
strLg125 = "Obrigado"
strLg126 = "em breve voc� receber� a resposta de sua d�vida!"
strLg127 = "Agora desfrute da comodidade e seguran�a de comprar em nossa loja virtual."
strLg128 = "Como Comprar na "
strLg129 = "Seu Nome:"
strLg130 = "Digite o seu e-mail corretamente para que esta mensagem possa ser respondida."
strLg131 = "D�vida sobre:"
strLg132 = "Assunto:"
strLg133 = "Mensagem:"
strLg134 = "Como utilizar o Carrinho de Compras?"
strLg135 = "Como pesquisar produtos na loja?"
strLg136 = "Como � o cadastro de clientes?"
strLg137 = "Como efetuar o login na loja?"
strLg138 = "Como modifico meus dados?"
strLg139 = "D�vidas diversas"
strLg140 = "Enviar"
strLg141 = "Preencha este campo!"
strLg142 = "O CPF deve ter 11 d�gitos!"
strLg143 = "CPF ou CNPJ Inv�lido!"
strLg144 = "Preencha este campo corretamente!"
strLg145 = "E-mail j� existente em nossos registros!"
strLg146 = "Modificar meu e-mail ou senha"
strLg147 = "Dia"
strLg148 = "M�s" 
strLg149 = "Registro Alterado"
strLg150 = "Seu registro foi alterado com sucesso!"
strLg151 = "Todos os seus dados foram modificados e atualizados com sucesso em nossa base de dados."
strLg152 = "Digite seu e-mail corretamente!"
strLg153 = "Fechar esta Janela"
strLg154 = "Seu e-mail foi registrado com sucesso!"
strLg155 = "Obrigado, agora voc� receber� ofertas exclusivas da "
strLg156 = "Este item j� existe no seu carrinho de compras."
strLg157 = "O item foi adicionado em seu carrinho de compras."
strLg158 = "Sua Primeira vez aqui?"
strLg159 = "Cadastre-se para receber ofertas exclusivas"
strLg160 = "Digite seu e-mail"
strLg161 = "Cadastre-se J�!"
strLg162 = "Este e-mail j� pertence a outro cliente cadastrado!"
strLg163 = "A senha digitada n�o confere!"
strLg164 = "Para modifica��o dos dados clique em 'Atualizar', para nenhuma modifica��o clique em 'Cancelar'."
strLg165 = "S� � permitida a altera��o de um item por vez."
strLg166 = "Dados do seu e-mail"
strLg167 = "se voc� quer modificar seu e-mail!"
strLg168 = "Seu e-mail atual:"
strLg169 = "Digite seu novo e-mail:"
strLg170 = "Dados da sua senha"
strLg171 = "se voc� quer modificar sua senha!"
strLg172 = "Digite sua senha atual:"
strLg173 = "Digite sua nova senha:"
strLg174 = "O campo Endere�o n�o pode estar vazio!"
strLg175 = "O campo Bairro n�o pode estar vazio!"
strLg176 = "O campo Cidade n�o pode estar vazio!"
strLg177 = "O campo CEP n�o pode estar vazio!"
strLg178 = "O campo Telefone n�o pode estar vazio!"
strLg179 = "Login na Loja"
strLg180 = "Se � a sua primeira compra nesta loja"
strLg181 = "Se j� � nosso cliente registrado ent�o introduza os dados."
strLg182 = "Voc� esqueceu sua senha?"
strLg183 = "E-mail de usu�rio n�o encontrado!"
strLg184 = "A Senha est� incorreta!"
strLg185 = "Selecione sua regi�o"
strLg186 = "Informa��es sobre o pedido "
strLg187 = "Esta compra n�o est� presente na base de dados da loja"
strLg188 = "ou n�o � referente a este cliente!"
strLg189 = "Compra efetuada por"
strLg190 = "SEDEX � cobrar"
strLg191 = "Dep�sito Identificado"
strLg192 = "Dep�sito Simples"
strLg193 = "Aguardando confirma��o de pagamento"
strLg194 = "Pagamento confirmado e Compra em processamento"
strLg195 = "Compra j� enviada ao cliente"
strLg196 = "Compra j� Entregue"
strLg197 = "Compra Cancelada"
strLg198 = "Compra Negada"
strLg199 = "Outras - Contate o Atendimento"
strLg200 = "Frete"
strLg201 = "Status da Compra"
strLg202 = "Data da Compra"
strLg203 = "Estimativa para entrega"
strLg204 = "Voc� ainda n�o efetuou nenhuma compra na"
strLg205 = "Ver informa��es deste pedido!"
strLg206 = "Departamento Inexistente"
strLg207 = "Departamento inexistente na"
strLg208 = "Nenhum produto presente neste departamento!"
strLg209 = "P�gina Anterior"
strLg210 = " - Cart�o Inv�lido."
strLg211 = " - N�mero Inv�lido."
strLg212 = " - N�mero e Cart�o Inv�lidos."
strLg213 = " - Soma de Controle Inv�lida."
strLg214 = " - Soma de Controle e Cart�o Inv�lidos."
strLg215 = " - Soma de Controle e N�mero Inv�lidos."
strLg216 = " - Soma de Controle, N�mero e Cart�o Inv�lidos"
strLg217 = " - Digite o n�mero do cart�o de cr�dito"
strLg218 = " - Selecione uma data v�lida!"
strLg219 = " - Data do cart�o inv�lida."
strLg220 = " - Digite o nome do titular (como aparece no cart�o)"
strLg221 = "Cart�o de Cr�dito MASTERCARD"
strLg222 = "Cart�o de Cr�dito VISA"
strLg223 = "Cart�o de Cr�dito AMERICAN EXPRESS"
strLg224 = "Cart�o de Cr�dito DINNERS"
strLg225 = "SEDEX � cobrar"
strLg226 = "Dep�sito Identificado"
strLg227 = "Dep�sito Banc�rio"
strLg228 = "Nenhum produto foi encontrado em nossa base de dados. Refa�a sua pesquisa!"
strLg229 = "Produto Inexistente ou Fora de Cat�logo"
strLg230 = "Produto inexistente ou Fora de Cat�logo na "
strLg231 = "Digite seu e-mail:"
strLg232 = "Senha:"
strLg233 = "Incorreto"
strLg234 = "Sua senha deve ter no m�ximo 10 caracteres!"
strLg235 = "Sua senha deve ter no m�nimo 5 caracteres!"
strLg236 = "Seu registro foi feito com sucesso!"
strLg237 = "Obrigado"
strLg238 = "agora desfrute da comodidade e seguran�a de comprar na "
strLg239 = "Compra Segura"
strLg240 = "Criptografia"
strLg241 = "Compras com Cart�o de Cr�dito"
strLg242 = "Utiliza��o de Informa��es"
strLg243 = "D�vidas e Sugest�es"
strLg244 = "E-mail n�o encontrado em nossos registros!"
strLg245 = "Preencha seu e-mail corretamente!"
strLg246 = "Sua Senha"
strLg247 = "Sua senha foi enviada com sucesso para"
strLg248 = "Digite seu e-mail abaixo para que sua senha possa ser enviada."
strLg249 = "E-mail:"
strLg250 = "Enviar Senha!"
strLg251 = "- Preencha os Campos ao lado."
strLg252 = "em breve voc� receber� a resposta de sua solicita��o."
strLg253 = "Equipe M�dica"
strLg254 = "Resultados Reais"
strLg255 = "Sobre"
strLg256  = "Boleto Banc�rio"
strLg257 = "Compre o produto "
strLg258 = "Atendimento On-Line "
strLg259 = "Atendimento Off-Line "
strLg260 = "Versions in "
strLg261 = "Esta p�gina foi carregada em "
strLg262 = " segundo(s)"
strLg263 = "Visitantes at� "
strLg264 = "Visitante "
strLg265 = "Seja Bem Vindo "
strLg266 = "Remover meu email "
strLg267 = "Incluir meu email para receber"
strLg268 = "Seu e-mail foi removido com sucesso!"
strLg269 = "Obrigado, apartir de hoje voc� n�o receber� ofertas exclusivas da "
strLg270 = " M�ximo de "
strLg271 = " produto(s) por compra. "
strLg272 = " unidade(s) "
strLg273 = " Voc� est� em "
strLg274 = " Campe�es de Venda "
strLg275 = " Enviar essa oferta a um amigo "
strLg276 = " Comprar Produto "
strLg277 = "Temos somente"
strLg278 = " produto(s) em nosso estoque"
strLg279 = "Deseja comprar os "
strLg280 = "produtos restantes?"
strLg281 = "O 'CEP' informado est� diferente do seu cadastro. Clique em Atualizar para recalcular"
strLg282 = " Calcular o valor da Entrega"
strLg283 = "Dados para entrega da compra"
strLg284 = "Forma de Pagamento"
strLg285 = "Compra conclu�da!"
strLg286 = "Fechando Pedido..."
strLg287 = "Como Pagar?"
strLg288 = "Como Reimprimir Boleto?"
strLg289 = "Dados para Login na loja"
strLg290 = "Dados Cadastrais"
strLg291 = "Nome do Contato"
strLg292 = " dias"
strLg293 = " ap�s a Confirma��o de pagamento"
strLg294 = "Escolha uma Forma de Entrega"
strLg295 = "Digite seu CEP para o calculo do Frete e clique em Atualizar"
strLg296 = "Pol�tica de Troca"
strLg297 = "Referencia" 
strLg298 = "Tamanho"
strLg299 = "Cor"
strLg300 = "Pagamento via PagSeguro" 
strLg301 = "Cadastrar como Revenda da Compuware?"

'###########################################
'Textos da loja
'###########################################

'Este texto � referente a �rea "Compra Segura"
'This text is used in the "Safety" area
strLgcomprasegura = "&nbsp;&nbsp;&nbsp;&nbsp;O " & nomeloja & " tem um especial compromisso com voc�, nosso cliente, quanto � seguran�a e privacidade de seus dados. O respeito ao cliente e o sigilo de suas informa��es � muito importante para n�s. Por isso, voc�, cliente do " & nomeloja & ", tem a garantia que sua compra � segura.<br>&nbsp;&nbsp;&nbsp;&nbsp;A " & nomeloja & " em parceria com a VirtuaWoks elaborou um conjunto de a��es para garantir a seguran�a de todas as compras feitas em nossa loja, independentemente da forma de pagamento escolhida. Todas as informa��es que voc� fornecer no processo de compra s�o totalmente criptografadas."


'Este texto � referente a �rea "Criptografia"
'This text is used in the "Criptography" area
strLgcriptografia = "&nbsp;&nbsp;&nbsp;&nbsp;Todas as informa��es que passam pelo nosso processo de compra s�o automaticamente codificadas por um sistema de criptografia pr�prio. Assim, seus dados pessoais, a forma de pagamento escolhida e toda e qualquer outra informa��o fornecida ao " & nomeloja & " ser� mantida em sigilo."


'Este texto � referente a �rea "Compras com Cart�o de Cr�dito"
'This text is used in the "Credit Card Sales" area
strLgcomprascc = "&nbsp;&nbsp;&nbsp;&nbsp;Al�m da criptografia, outro fator de seguran�a � a autom�tica destrui��o dos dados relativos ao n�mero do cart�o de cr�dito. A " & nomeloja & " utiliza o n�mero do cart�o somente no processamento da compra e, t�o logo ocorra a confirma��o pela Administradora do cart�o, o n�mero � automaticamente destru�do, n�o sendo, de nenhuma forma, guardado na base de dados da " & nomeloja & "."


'Este texto � referente a �rea "Utiliza��o de Informa��es"
'This text is used in the "Your Information" area
strLginformacoes = "&nbsp;&nbsp;&nbsp;&nbsp;A " & nomeloja & " n�o comercializar� suas informa��es pessoais. Tais informa��es poder�o, entretanto, ser agrupadas conforme determinados crit�rios e utilizadas como estat�sticas gen�ricas objetivando um melhor entendimento do perfil do consumidor."


'Este texto � referente a �rea "D�vidas e sugest�es"
'This text is used in the "faq and sugestions" area
strLgduvsug = "&nbsp;&nbsp;&nbsp;&nbsp;Caso tenha qualquer d�vida ou sugest�o sobre seguran�a e privacidade ou sobre qualquer outro aspecto do nosso processo de compra, n�o deixe de nos contatar."


'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLgquemsomos = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa,Texto referente a informa��es da sua empresa, "

'Este texto � referente a �rea "Restrita"
'This text is used in the "restrict" area
strLgrestrito = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Texto referente a �rea restrita,Texto referente a �rea restrita,Texto referente a �rea restrita,Texto referente a �rea restrita,Texto referente a �rea restrita,Texto referente a �rea restrita,Texto referente a �rea restrita,Texto referente a �rea restrita,Texto referente a �rea restrita,Texto referente a �rea restrita,Texto referente a �rea restrita,Texto referente a �rea restrita,Texto referente a �rea restrita ,"

'Este texto � referente a �rea "Tabela"
'This text is used in the "Tabela" area
strLgtabela = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; EM CONSTRU��O,"
%>