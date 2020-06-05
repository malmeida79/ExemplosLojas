<%
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################
'#
'#  C�DIGO: VirtuaStore Vers�o OPEN - Copyright 2001-2004 VirtuaStore
'#  URL: http://comunidade.virtuastore.com.br
'#  E-MAIL: jusivansjc@yahoo.com.br
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
'#  // "bondhost - Hospedagem" ou "Software derivado de VirtuaStore 1.2" e 
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
'#
'#     Para ajuda e suporte acesse um dos sites abaixo:
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
' Portugu�s (PT) - LCID 2070
' Pais : Portugal/Brasil
' HTML Content-Type : ISO-8859-1 
' Traduzido por : Pedro Magalh�es 27/12/2003 
' Acr�scimos/Corre��es por : Elizeu Alcantara (jusivansjc@yahoo.com.br / jusivansjc@yahoo.com.br ) 15/03/2004 
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
strLg5 = "Registo de Clientes"
strLg6 = "Login"
strLg7 = "Carrinho de Compras de"
strLg8 = "Fechar Pedido"
strLg9 = "Os meus Dados"
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
strLg28 = "Stock:"
strLg29 = "Pre�o:"
strLg30 = "+ Detalhes"
strLg31 = "Fabricante:"
strLg32 = "Adicionar"
strLg33 = "produto(s) no meu carrinho de compras."
strLg34 = "Comprando o"
strLg35 = "na"
strLg36 = "voc� economiza at� " 
strLg37 = "A pronto por:"
strLg38 = "por:"
strLg39 = "de:"
strLg40 = "ou ainda:"
strLg41 = "Voltar"
strLg42 = "Pesquisar por" 
strLg43 = "Foram encontrado(s)"
strLg44 = "produto(s) :"
strLg45 = "P�gina"
strLg46 = "de"
strLg47 = "Pr�xima P�gina"
strLg48 = "Produtos neste departamento:"
strLg49 = "N�o h� nenhum item em seu carrinho de compras."
strLg50 = "Se voc� comprou algum produto anteriormente a sua sess�o de compras est� expirada!"
strLg51 = "Ver o meu carrinho de compras"
strLg52 = "Mais Ofertas:"
strLg53 = "O item foi adicionado ao seu carrinho de compras."
strLg54 = "Compra"
strLg55 = "A sua compra"
strLg56 = "Por favor, verifique se sua compra est� correta antes de prosseguir."
strLg57 = "Para modificar a quantidade de algum item digite o novo valor e clique no bot�o actualizar."
strLg58 = "Quantidade"
strLg59 = "Produto"
strLg60 = "Pre�o Unit�rio"
strLg61 = "Pre�o Total"
strLg62 = "Remover"
strLg63 = "Total da compra"
strLg64 = "Valor da entrega em"
strLg65 = "A sua compra"
strLg66 = "actualizar"
strLg67 = "Prosseguir"
strLg68 = "Continuar as Compras"
strLg69 = "Escolha a regi�o de entrega para o calculo do frete!"
strLg70 = "Clique aqui"
strLg71 = "Clique aqui!"
strLg72 = "Dados para entrega do pedido"
strLg73 = "Para modificar a sua compra"
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
strLg103 = "Os Meus Dados"
strLg104 = "Para modificar os seus dados, selecione o campo desejado e escreva os dados actualizados sobre o antigo."
strLg105 = "Para manter clique em 'Cancelar'. "
strLg106 = "Dados do registo"
strLg107 = "Dados Pessoais"
strLg108 = "O seu e-mail:"
strLg109 = "O seu telefone:"
strLg110 = "O seu pa�s:"
strLg111 = "O seu estado:"
strLg112 = "A sua cidade:"
strLg113 = "O seu bairro:"
strLg114 = "O seu endere�o:"
strLg115 = "O seu RG ou IE"
strLg116 = "O CPF ou CNPJ"
strLg117 = "Data nascimento:"
strLg118 = "O seu Nome ou Raz�o Social:"
strLg119 = "Compras efetuadas por"
strLg120 = "Para visualizar os dados da sua compra digite o n�mero do pedido abaixo:"
strLg121 = "Pedido n.:"
strLg122 = "Para sua comodidade e seguran�a verifique todos os seus dados antes de prosseguir." 
strLg123 = "A sua Senha:"
strLg124 = "O seu e-mail foi enviado com sucesso!"
strLg125 = "Obrigado"
strLg126 = "em breve voc� receber� resposta � sua d�vida!"
strLg127 = "Agora desfrute da comodidade e seguran�a de comprar em nossa loja virtual."
strLg128 = "Como Comprar na "
strLg129 = "O seu Nome:"
strLg130 = "Digite o seu e-mail corretamente para que esta mensagem possa ser respondida."
strLg131 = "D�vida sobre:"
strLg132 = "Assunto:"
strLg133 = "Mensagem:"
strLg134 = "Como utilizar o Carrinho de Compras?"
strLg135 = "Como pesquisar produtos na loja?"
strLg136 = "Como � o registo de clientes?"
strLg137 = "Como efetuar o login na loja?"
strLg138 = "Como modifico meus dados?"
strLg139 = "D�vidas diversas"
strLg140 = "Enviar"
strLg141 = "Preencha este campo!"
strLg142 = "O NIF deve ter 9 d�gitos!"
strLg143 = "NIF Inv�lido!"
strLg144 = "Preencha este campo corretamente!"
strLg145 = "E-mail j� existente em nossos registos!"
strLg146 = "Modificar o meu e-mail ou senha"
strLg147 = "Dia"
strLg148 = "M�s" 
strLg149 = "registo Alterado"
strLg150 = "O seu registo foi alterado com sucesso!"
strLg151 = "Todos os seus dados foram modificados e actualizados com sucesso em nossa base de dados."
strLg152 = "Digite seu e-mail corretamente!"
strLg153 = "Fechar esta Janela"
strLg154 = "O seu e-mail foi registrado com sucesso!"
strLg155 = "Obrigado, agora voc� receber� ofertas exclusivas da "
strLg156 = "Este item j� existe no seu carrinho de compras."
strLg157 = "O item foi adicionado ao seu carrinho de compras."
strLg158 = "Esta � a sua primeira vez aqui?"
strLg159 = "Registe-se e receba ofertas exclusivas"
strLg160 = "Digite o seu e-mail"
strLg161 = "Registe-se J�!"
strLg162 = "Este e-mail j� pertence a outro cliente registado!"
strLg163 = "A senha digitada n�o confere!"
strLg164 = "Para modificar os dados clique em 'actualizar', para manter clique em 'Cancelar'."
strLg165 = "S� � permitida a altera��o de um item de cada vez."
strLg166 = "Dados do seu e-mail"
strLg167 = "se voc� quiser modificar o seu e-mail!"
strLg168 = "O seu e-mail actual:"
strLg169 = "Digite o seu novo e-mail:"
strLg170 = "Dados da sua senha"
strLg171 = "se voc� quer modificar a sua senha!"
strLg172 = "Digite a sua senha actual:"
strLg173 = "Digite sua nova senha:"
strLg174 = "O campo Endere�o n�o pode estar vazio!"
strLg175 = "O campo Bairro n�o pode estar vazio!"
strLg176 = "O campo Cidade n�o pode estar vazio!"
strLg177 = "O campo CEP n�o pode estar vazio!"
strLg178 = "O campo Telefone n�o pode estar vazio!"
strLg179 = "Login na Loja"
strLg180 = "Se � a sua primeira visita a esta loja"
strLg181 = "Se j� � nosso cliente registado ent�o introduza os dados."
strLg182 = "Voc� esqueceu a sua senha?"
strLg183 = "E-mail de utilizador n�o encontrado!"
strLg184 = "A Senha est� incorreta!"
strLg185 = "Selecione a sua regi�o"
strLg186 = "Informa��es sobre o pedido "
strLg187 = "Esta compra n�o est� presente na base de dados da loja"
strLg188 = "ou n�o � referente a este cliente!"
strLg189 = "Compra efetuada por"
strLg190 = "CTT � cobran�a"
strLg191 = "Dep�sito Banc�rio"
strLg192 = "Transfer�ncia Eletr�nica"
strLg193 = "Aguardando confirma��o de pagamento"
strLg194 = "Pagamento confirmado e Compra em processamento"
strLg195 = "Compra j� enviada ao cliente"
strLg196 = "Compra J� Entregue"
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
strLg225 = "CTT � cobran�a"
strLg226 = "Dep�sito Banc�rio"
strLg227 = "Transfer�ncia Eletr�nica"
strLg228 = "N�o foram encontrados PRODUTOS para esta pesquisa!"
strLg229 = "Produto Inexistente ou Fora de Cat�logo"
strLg230 = "Produto inexistente ou Fora de Cat�logo na"
strLg231 = "Digite o seu e-mail:"
strLg232 = "Senha:"
strLg233 = "Incorreto"
strLg234 = "A sua senha deve ter no m�ximo 10 caracteres!"
strLg235 = "A sua senha deve ter no m�nimo 5 caracteres!"
strLg236 = "Seu registo foi feito com sucesso!"
strLg237 = "Obrigado"
strLg238 = "agora desfrute da comodidade e seguran�a de comprar na "
strLg239 = "Compra Segura"
strLg240 = "Criptografia"
strLg241 = "Compras com Cart�o de Cr�dito"
strLg242 = "Utiliza��o de Informa��es"
strLg243 = "D�vidas e Sugest�es"
strLg244 = "E-mail n�o encontrado nos nossos registos!"
strLg245 = "Preencha o seu e-mail corretamente!"
strLg246 = "A sua Senha"
strLg247 = "A sua senha foi enviada com sucesso para"
strLg248 = "Digite o seu e-mail abaixo para que a sua senha possa ser enviada."
strLg249 = "E-mail:"
strLg250 = "Enviar Senha!"
strLg251 = "- Preencha os Campos ao lado."
strLg252 = "Em breve voc� receber� resposta para a sua solicita��o."
strLg253 = "Equipa M�dica"
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
strLg267 = "Incluir meu email "
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
strLg294 = "Escolha uma Forma de Pagamento"
strLg295 = "Digite seu CEP para o calculo do Frete e clique em Atualizar"
strLg296 = "Pol�tica de Troca"

strLg297= "Pagamento"
strLg298= "Dep�sito direto no caixa"
strLg299= "Dep�sito no caixa eletr�nico"
strLg300= "Dep�sito atr�ves de DOC"
strLg301= "Transfer�ncia entre contas"
strLg302= "Boleto banc�rio"
strLg303= "Data: "
strLg304= "Observa��es: "
strLg305= "Hor�rio: "
strLg306= "Ag�ncia Origem: "
strLg307= "Valor do pagamento: "
strLg308= "em breve voc� receber� a confirma��o do seu pagamento!"
strLg309= "Dentro de 1 a 2 dias �teis estaremos confirmando seu pagamento por e-mail.<br>Caso necessite de outras informa��es, entre em contato conosco pelo e-mail "&Email_da_sua_loja
strLg310= "Meu Carrinho"
strLg311= "Confirmar pagamento"
strLg312= "Pol�tica de Venda"
strLg313= "Como Confirmar Pagamento"
strLg314 = "O campo Estado n�o pode estar vazio!"

strLg315 = "Novos PRODUTOS est�o sendo catalogados."
strLg316 = "Se voc� desejar algo que ainda n�o encontrou,"
strLg317 = "e nos informe o que est� procurando."
strLg318 = "Informa��es sobre produtos"

'#############################
'## STRINGS DO TERMO DE USO ##
'#############################
'strLg360 = "1. Aceita��o dos Termos do Contrato"
'strLg361 = "2. Restri��es de Uso"
'strLg362 = "3. Marcas Comerciais"
'strLg363 = "4. Desconhecimento de Garantia"
'strLg364 = "5. Nossas Transmiss�es"
'strLg365 = "6. Revis�es dos Termos de Uso"
'strLg366 = "7. Lei Aplic�vel"
'strLg367 = "Termo de Uso Site"


'###########################################
'	Textos da loja							
'###########################################

'Este texto � referente a �rea "Compra Segura"
'This text is used in the "Safety" area
strLgcomprasegura = "&nbsp;&nbsp;&nbsp;&nbsp;A " & nomeloja & " tem um especial compromisso com voc�, nosso cliente, quanto � seguran�a e privacidade dos seus dados. O respeito pelo cliente e o sigilo das suas informa��es � muito importante para n�s. Por isso, voc�, cliente da " & nomeloja & ", tem a garantia que sua compra � segura.<br>&nbsp;&nbsp;&nbsp;&nbsp;A " & nomeloja & " em parceria com a VirtuaWoks elaborou um conjunto de ac��es para garantir a seguran�a de todas as compras feitas na nossa loja, independentemente da forma de pagamento escolhida. Todas as informa��es que voc� fornecer no processo de compra s�o totalmente criptografadas."

'Este texto � referente a �rea "Criptografia"
'This text is used in the "Criptography" area
strLgcriptografia = "&nbsp;&nbsp;&nbsp;&nbsp;Todas as informa��es que passam pelo nosso processo de compra s�o automaticamente codificadas por um sistema de criptografia pr�prio. Assim, os seus dados pessoais, a forma de pagamento escolhida e toda e qualquer outra informa��o fornecida � " & nomeloja & " ser� mantida em sigilo. O �cone 'cadeado fechado' - na parte inferior do seu monitor - durante o pedido � um s�mbolo da criptografia das informa��es."

'Este texto � referente a �rea "Compras com Cart�o de Cr�dito"
'This text is used in the "Credit Card Sales" area
strLgcomprascc = "&nbsp;&nbsp;&nbsp;&nbsp;Al�m da criptografia, outro factor de seguran�a � a autom�tica destrui��o dos dados relativos ao n�mero do cart�o de cr�dito. A " & nomeloja & " utiliza o n�mero do cart�o somente no processamento da compra e, t�o logo que ocorra a confirma��o pela Concessionaria do cart�o, o n�mero � automaticamente destru�do, n�o sendo, de nenhuma forma, guardado na base de dados da " & nomeloja & "."

'Este texto � referente a �rea "Utiliza��o de Informa��es"
'This text is used in the "Your Information" area
strLginformacoes = "&nbsp;&nbsp;&nbsp;&nbsp;A " & nomeloja & " n�o comercializar� as suas informa��es pessoais. Tais informa��es poder�o, entretanto, ser agrupadas conforme determinados crit�rios e utilizadas como estat�sticas gen�ricas objetivando um melhor entendimento do perfil do consumidor."

'Este texto � referente a �rea "D�vidas e sugest�es"
'This text is used in the "faq and sugestions" area
strLgduvsug = "&nbsp;&nbsp;&nbsp;&nbsp;Caso tenha qualquer d�vida ou sugest�o sobre seguran�a e privacidade ou sobre qualquer outro aspecto do nosso processo de compra, n�o deixe de nos contatar."

'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLgquemsomos = "&nbsp;&nbsp;&nbsp;&nbsp;N�s somos uma loja virtual que vende produtos de qualidade!"
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa = "&nbsp;&nbsp;&nbsp;&nbsp;Os seguintes s�o os termos de um contrato estabelecido entre voc� (cliente) e o " & nomeloja & ". Ao acessar, navegar e/ou usar este site (Site), voc� admite ter lido e entendido estes termos e cumprir� com todas as leis e regras aplic�veis."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa1 = "&nbsp;&nbsp;&nbsp;&nbsp;O " & nomeloja & " n�o se responsabiliza se o material deste Site for apropriado, ou esteja dispon�vel para uso em outros lugares, estando proibido o seu acesso em territ�rios onde seu conte�do seja ilegal. Aqueles que decidam acessar este site em outros lugares o far�o por iniciativa pr�pria e � sua responsabilidade se sujeitar as leis locais aplic�veis."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa2 = "&nbsp;&nbsp;&nbsp;&nbsp;Os direitos de autor de todo o material apresentado neste Site s�o propriedade do " & nomeloja & ", ou do criador original do material. Com excess�o do expressamente estipulado neste documento, este material ou nenhuma parte do mesmo pode ser copiado, reproduzido, distribuido, re-publicado, apresentado, anunciado ou transmitido de nenhuma maneira ou por nenhum meio, incluindo, mas n�o limitado a, meios eletr�nicos, mec�nicos, de fotoc�pia, de gravac�o ou de qualquer outra �ndole, sem a permiss�o previa por escrito do " & nomeloja & " ou do titular dos direitos de autor."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa3 = "&nbsp;&nbsp;&nbsp;&nbsp;Concede-se a autoriza��o a permiss�o para apresentar no ecran, copiar e distribuir o material neste site unicamente para uso pessoal, n�o comercial, com a condi��o de que n�o se modifique o material e que se conserve todas as legendas de direitos de autor e de outros tipos de propriedade inclu�dos neste material. Esta permiss�o termina automaticamente se voc� violar qualquer destes termos e condi��es. Ao seu t�rmino, se dever� destruir qualquer material reproduzido ou impresso. Tampouco � permitido, sem a permiss�o do " & nomeloja & ", replicar em outro servidor qualquer material inclu�do neste Site."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa4 = "&nbsp;&nbsp;&nbsp;&nbsp;Qualquer uso n�o autorizado de qualquer material inclu�do neste site pode constituir uma viola��o das leis de direitos de autor, das leis de marcas comerciais, das leis de privacidade e publicidade e das leis e regras de comunica��es."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa5 = "&nbsp;&nbsp;&nbsp;&nbsp;As marcas comerciais, as marcas de servi�o e os logotipos (as Marcas Comerciais) usados e apresentados neste Site s�o Marcas Comerciais registradas ou n�o registradas do " & nomeloja & " e outros. Nada neste Site dever� ser interpretado como concess�o, por implicitude, por exclus�o ou de nenhuma outra maneira, de nenhuma licen�a ou de direito de uso de qualquer Marca Comercial apresentada neste Site, sem a permiss�o por escrito do titular dos direitos da Marca Comercial."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa6 = "&nbsp;&nbsp;&nbsp;&nbsp;O " & nomeloja & " faz cumprir de maneira ativa os seus direitos de propriedade intelectual em toda a extens�o da lei. O nome " & nomeloja & " e/ou o logotipo do " & nomeloja & " n�o pode ser usado de maneira alguma, incluindo an�ncios ou publicidade relacionados a distribui��o do material neste Site, sem permiss�o previa por escrito."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa7 = "&nbsp;&nbsp;&nbsp;&nbsp;O material neste site � apresentado ASSIM COMO EST�,sem haver inclu�do garantias de qualquer tipo, expl�citas ou impl�citas. em cumprimento com a lei aplic�vel, na maior medida poss�vel, o " & nomeloja & " desconhece todas as garantias expl�citas ou impl�citas, incluindo, mas n�o limitada a, garantias impl�citas de comercializa��o, adequa��o para um prop�sito particular, n�o contraven��o ou outra viola��o de direitos."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa8 = "&nbsp;&nbsp;&nbsp;&nbsp;O " & nomeloja & " n�o garante nem faz representa��o alguma em rela��o ao uso, validez, precis�o ou confiabilidade de, ou dos resultados do uso de, ou qualquer outra situa��o relacionada ao material neste site ou qualquer site com um link com este site. responsabilidade limitada sobre nenhuma hip�tese, incluindo, mas n�o limitada a, neglig�ncia, o " & nomeloja & " ser� respons�vel por qualquer dano directo, indirecto, especial, incidental, consequente, incluindo, mas n�o limitado a, perda de informa��o ou utilidades, resultado de uso ou a incapacidade de usar o material neste site, mesmo no caso de que o " & nomeloja & " ou um representante autorizado da " & nomeloja & " tenha consci�ncia da possibilidade de tal dano."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa9 = "&nbsp;&nbsp;&nbsp;&nbsp;No caso de que o uso do material neste site por sua parte resulte em necessidade de servi�o, reparo ou corre��o de equipamento ou informa��o, voc� assume qualquer custo derivado dele. alguns pa�ses n�o permitem a exclus�o ou limita��o de danos incidentais ou consequentes, assim que as limita��es e exclus�es mencionadas acima podem n�o ser aplic�veis em seu caso."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa10 = "&nbsp;&nbsp;&nbsp;&nbsp;Qualquer material, informac�o ou id�ia que voc� transmita ou coloque neste Site por qualquer meio ser� tratado como confidencial e somente pode ser disseminado ou usado pelo " & nomeloja & " ou seus parceiros."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa11 = "&nbsp;&nbsp;&nbsp;&nbsp;Voc� est� proibido de colocar em, ou transmitir deste Site qualquer material ilegal, amea�ador, caluniador, difamat�rio, obsceno, escandaloso, provocante, pornogr�fico ou profano ou qualquer outro material que possa dar oportunidade a qualquer responsabilidade civil ou penal nos termos da lei."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa12 = "&nbsp;&nbsp;&nbsp;&nbsp;O Shopping VirtuaStore pode em qualquer momento revisar estes Termos de Uso por meio da actualiza��o deste an�ncio. Ao usar este Site, voc� concorda em cumprir quaisquer revis�es, devendo ent�o visitar periodicamente esta p�gina para determinar os termos de uso vigentes nesse momento, com os quais voc� deve cumprir."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa13 = "&nbsp;&nbsp;&nbsp;&nbsp; Este Termo de Uso � regido pelas leis brasileiras."
'Este texto � referente a �rea "Quem Somos"
'This text is used in the "About Us" area
strLglicensa14 = "&nbsp;&nbsp;&nbsp;&nbsp;� Copyright 2004 " & nomeloja & ".Proibida sua reprodu��o total ou parcial."
%>
