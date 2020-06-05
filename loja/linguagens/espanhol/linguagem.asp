<%
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################
'#
'#  C�DIGO: VirtuaStore Vers�o OPEN - Copyright 2001-2004 VirtuaStore
'#  E-MAIL: corpsjc@megapaper.com.br
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
'#     Internet Design entre em contato atrav�s do e-mail 
'#     jusivansjc@yahoo.com.br ou pelo endere�o abaixo: 
'#     Rua Ven�ncio Aires, 1001 - Niter�i - Canoas - RS - Brasil. Cep 92110-000.
'#
'#     Para ajuda e suporte acesse um dos sites abaixo:
'#     http://comunidade.virtuastore.com.br
'#     http://www.bondhost.com.br
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
' Espanhol
' Pais : Brasil
' HTML Content-Type : ISO-8859-1 
' Traduzido por : Henrique Metasupri (corpsjc@megapaper.com.br) 09/01/2004
' Acr�scimos/Corre��es por : Elizeu Alcantara (corpsjc@megapaper.com.br / jusivansjc@yahoo.com.br ) 15/03/2004 
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

strLg1 = "Acabar compras" 
strLg2 = "Cantidad de itens:" 
strLg3 = "Subtotal:" 
strLg4 = "Principal" 
strLg5 = "Registro de clientes" 
strLg6 = "Conexi�n" 
strLg7 = "Soporte de compras de" 
strLg8 = "Conexi�n" 
strLg9 = "mis datos" 
strLg10 = "hist�rico de compras" 
strLg11 = "registro de estado de la m�quina" 
strLg12 = "departamentos" 
strLg13 = "investigaci�n" 
strLg14 = "atenci�n" 
strLg15 = "todas las categor�as" 
strLg16 = "en cuanto a compra" 
strLg17 = "por el email" 
strLg18 = "en seguridad" 
strLg19 = "seguridad y aislamiento" 
strLg20 = "qui�nes somos" 
strLg21 = "buscar" 
strLg22 = "en seguridad" 
strLg23 = "por el email" 
strLg24 = "atenci�n por el email" 
strLg25 = "para la entrega" 
strLg26 = "Disponible para la entrega" 
strLg27 = "No disponible para la entrega" 
strLg28 = "fuente:" 
strLg29 = "precio:" 
strLg30 = "+ Detalla" 
strLg31 = "Fabricante:" 
strLg32 = "agregar" 
strLg33 = "product(s) en mi soporte de compras." 
strLg34 = "comprando este producto" 
strLg35 = "en" 
strLg36 = "usted incluso excepto" 
strLg37 = "a la vista para:" 
strLg38 = "para:" 
strLg39 = "de:" 
strLg40 = "o a�n:" 
strLg41 = "volverse" 
strLg42 = "buscando para" 
strLg43 = "hab�a sido encontrado(s)" 
strLg44 = "produto(s):" 
strLg45 = "p�gina" 
strLg46 = "de" 
strLg47 = "p�gina siguiente" 
strLg48 = "productos en este departamento:" 
strLg49 = "no tiene ning�n art�culo en su soporte de compras." 
strLg50 = "si le su sesi�n de las compras compradas alg�n producto exhalan previamente" 
strLg51 = "ver mi soporte de compras" 
strLg52 = "m�s ofertas:"
strLg53 = "el art�culo fue agregado en su soporte de compras." 
strLg54 = "compra" 
strLg55 = "su compra" 
strLg56 = "por favor, verifica si su compra est� correcta antes de continuar." 
strLg57 = "modificar la cantidad de un cierto art�culo mecanograf�a el nuevos valor y tecleo en el bot�n para traer actualizado." 
strLg58 = "cantidad" 
strLg59 = "producto" 
strLg60 = "precio unitario" 
strLg61 = "precio total" 
strLg62 = "quitar" 
strLg63 = "total de la compra" 
strLg64 = "valor de la entrega en" 
strLg65 = "su compra" 
strLg66 = "traer actualizado" 
strLg67 = "continuar" 
strLg68 = "continuar compras" 
strLg69 = "opci�n la regi�n de la entrega para m� la calculo de la carga" 
strLg70 = "chasco aqu�" 
strLg71 = "chasco aqu�" 
strLg72 = "dado para la entrega de la orden" 
strLg73 = "modificar su compra" 
strLg74 = "por favor, verifica si la direcci�n para la entrega est� correcta antes de continuar." 
strLg75 = "si direcci�n para la entrega e igual que su" 
strLg76 = "opciones para la entrega" 
strLg77 = "direcci�n para la entrega" 
strLg78 = "direcci�n:" 
strLg79 = "cuarto:" 
strLg80 = "ciudad:" 
strLg81 = "sido:" 
strLg82 = "CEP:" 
strLg83 = "pa�s:" 
strLg84 = "tel�fono:" 
strLg85 = "usted desea que los itens comprados est�n embalados y enviado como presente" 
strLg86 = "s�" 
strLg87 = "no" 
strLg88 = "mensaje escrito en el tarjeta-regalo:" 
strLg89 = "limpiar" 
strLg90 = "por favor, terraplenes todos correctamente los datos antes de continuar." 
strLg91 = "modo del pago" 
strLg92 = "por favor, terraplenes todos correctamente los datos antes de continuar." 
strLg93 = "modificar la direcci�n de la entrega de las compras y/o de otros datos" 
strLg94 = "selecciona el modo del pago elegido:" 
strLg95 = "otros modos del pago" 
strLg96 = "tarjeta de cr�dito de la sierra del pago" 
strLg97 = "selecciona su tarjeta:" 
strLg98 = "n�mero de la tarjeta:" 

strLg98 = "n�mero de la tarjeta: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;n�mero de seguranca:"

strLg99 = "expiraci�n de la tarjeta:" 
strLg100 = "cancelar" 
strLg101 = "a�o" 
strLg102 = "titular de la tarjeta (como ella aparece en la tarjeta):" 
strLg103 = "mis datos" 
strLg104 = "modificar sus datos, selecciona el campo deseado y escribe los datos tra�dos actualizado en el viejo." 
strLg105 = "para ning�n tecleo de la modificaci�n en ' cancelar '. "
strLg106 = "dado del registro" 
strLg107 = "datos personales" 
strLg108 = "su email:" 
strLg109 = "su tel�fono:" 
strLg110 = "su pa�s:" 
strLg111 = "su estado:" 
strLg112 = "su ciudad:" 
strLg113 = "su cuarto:" 
strLg114 = "su direcci�n:" 
strLg115 = "su RG / IE" 
strLg116 = "su CPF / CNPJ" 
strLg117 = "su nacimiento:" 
strLg118 = "su nombre completo:" 
strLg119 = "compras efectuadas por" 
strLg120 = "visualizar los datos de su compra mecanograf�a el n�mero de la orden pronto abajo:" 
strLg121 = "pedido el n.:" 
strLg122 = "para su comodidad y seguridad verifica todos sus datos antes de continuar." 
strLg123 = "Su Contrase�a:" 
strLg124 = "su email fue enviado con �xito" 
strLg125 = "obligado" 
strLg126 = "en el informe que usted recibir� la contestaci�n de su duda" 
strLg127 = "ahora goza de la comodidad y de la seguridad para comprar en nuestra memoria virtual." 
strLg128 = "en cuanto a compra en" 
strLg129 = "su nombre:" 
strLg130 = "mecanograf�a su email correctamente para poder contestar este mensaje." 
strLg131 = "duda encendido:" 
strLg132 = "tema:" 
strLg133 = "mensaje:" 
strLg134 = "en cuanto a uso el soporte de compras" 
strLg135 = "en cuanto a productos de la b�squeda en el almac�n" 
strLg136 = "pues �l es yo lo coloca en cadastre de clientes" 
strLg137 = "en cuanto a la conexi�n del efecto en el almac�n" 
strLg138 = "pues modifico mis datos" 
strLg139 = "dudas diversas" 
strLg140 = "enviar" 
strLg141 = "terraplenes este campo" 
strLg142 = "el CPF debo tener 11 d�gitos" 
strLg143 = "CPF ou CNPJ inv�lido" 
strLg144 = "terraplenes este campo correctamente" 
strLg145 = "email existente ya en nuestros registros" 
strLg146 = "modificar mi email o contrase�a" 
strLg147 = "d�a" 
strLg148 = "mes" 
strLg149 = "registro modificado" 
strLg150 = "su registro fue modificado con �xito" 
strLg151 = "todos sus datos hab�a sido modificado y trajo actualizado con �xito en nuestra base de datos." 
strLg152 = "mecanograf�a su email correctamente" 
strLg153 = "cerrar esta ventana" 
strLg154 = "su email fue colocado con �xito" 
strLg155 = "obligado, ahora usted recibir� las ofertas exclusivas de" 
strLg156 = "este art�culo existe ya en su soporte de compras." 
strLg157 = "el art�culo fue agregado en su soporte de compras." 
strLg158 = "su primera vez aqu�" 
strLg159 = "mecanograf�a su email y ofertas recive de las exclusivas." 
strLg160 = "mecanograf�a su email" 
strLg161 = "se coloca en cadastre ya"
strLg162 = "este email pertenece ya a otro registrado en cliente del cadastre" 
strLg163 = "la contrase�a mecanografiada no confiere" 
strLg164 = "para la modificaci�n del tecleo de los datos en ' traer actualizado ', para ning�n tecleo de la modificaci�n en ' cancelar '." 
strLg165 = "no se prohibe solamente la alteraci�n de un art�culo por tiempo." 
strLg166 = "dado de su email" 
strLg167 = "si usted desea modificar su email" 
strLg168 = "su email actual:" 
strLg169 = "mecanograf�a su email nuevo:" 
strLg170 = "dado de su contrase�a" 
strLg171 = "si usted desea modificar su contrase�a" 
strLg172 = "mecanograf�a su corriente de la contrase�a:" 
strLg173 = "mecanograf�a su nueva contrase�a:" 
strLg174 = " el campo Direcci�n no puede ser vac�o" 
strLg175 = " el campo Bairro no puede ser que vac�o" 
strLg176 = " el campo Ciudad no puede ser vac�o" 
strLg177 = " el campo CEP no puede ser vac�o" 
strLg178 = " el campo Tel�fono  no puede ser vac�o" 
strLg179 = "conexi�n en el almac�n" 
strLg180 = "si su primer es visita este almac�n" 
strLg181 = "si est� nuestro cliente registrado despu�s introduce ya los datos." 
strLg182 = "usted se olvid� que su contrase�a" 
strLg183 = "email del usuario no encontrado" 
strLg184 = "la contrase�a es incorrecto" 
strLg185 = "selecciona su regi�n" 
strLg186 = "informaci�n sobre la orden " 
strLg187 = "esta compra no est� presente en la base de datos del almac�n" 
strLg188 = "ni est� refiriendo a este cliente" 
strLg189 = "compra efectuada por" 
strLg190 = "SEDEX a la carga" 
strLg191 = "deposito bancario"
strLg192 = "transfer�ncia eletronica"
strLg193 = "aguardando confirmacion de la paga"
strLg194 = "paga confirmado e compra en el proceso"
strLg195 = "compra enviada ya al cliente" 
strLg196 = "compra entrega ya" 
strLg197 = "compra cancelada" 
strLg198 = "compra negada" 
strLg199 = "otros - la atenci�n entra en contacto con" 
strLg200 = "carga" 
strLg201 = "estado de la compra" 
strLg202 = "fecha de la compra" 
strLg203 = "Estimative para la entrega" 
strLg204 = "usted todav�a no efectu� ninguna compra en" 
strLg205 = "ver la informaci�n de esta orden" 
strLg206 = "departamento inexistente" 
strLg207 = "departamento inexistente en" 
strLg208 = "ning�n actual producto en este departamento" 
strLg209 = "p�gina anterior" 
strLg210 = "- tarjeta inv�lida." 
strLg211 = "- N�mero Inv�lido." 
strLg212 = "- n�mero inv�lido y tarjeta." 
strLg213 = "- adici�n inv�lida del control." 
strLg214 = "- adici�n del control inv�lido y de la tarjeta." 
strLg215 = "- adici�n del control y del n�mero inv�lidos." 
strLg216 = "- adici�n inv�lida del control, del n�mero y de la tarjeta" 
strLg217 = "- el n�mero de los tipos de la tarjeta de cr�dito" 
strLg218 = "- una fecha selecciona v�lido" 
strLg219 = "- fecha de la tarjeta inv�lida." 
strLg220 = "- el nombre del portador mecanograf�a (mientras que aparece en la tarjeta)" 
strLg221 = "MASTERCARD de la tarjeta de cr�dito" 
strLg222 = "tarjeta de cr�dito TIENE COMO OBJETIVO" 
strLg223 = "AMERICANO de la tarjeta de cr�dito EXPRESO" 
strLg224 = "las CENAS de la tarjeta de cr�dito" 
strLg225 = "SEDEX a la carga" 
strLg226 = "deposito bancario"
strLg227 = "transferencia eletronica"
strLg228 = "ning�n producto fue encontrado en nuestra base de datos. "
strLg229 =" producto inexistente "
strLg230 =" producto inexistente en "
strLg231 =" los tipos su email: " 
strLg232 = "contrase�a:" 
strLg233 = "incorrecto" 
strLg234 = "su contrase�a debe tener en los caracteres del m�ximo 10" 
strLg235 = "su contrase�a debe tener 5 caracteres por lo menos" 
strLg236 = "su registro fue hecho con �xito" 
strLg237 = "obligado" 
strLg238 = "ahora goza de la comodidad y la seguridad a comprar en" 
strLg239 = "compra segura" 
strLg240 = "Criptografia" 
strLg241 = "compras con la tarjeta de cr�dito" 
strLg242 = "uso de la informaci�n" 
strLg243 = "dudas y las sugerencias" 
strLg244 = "email no encontrado en nuestros registros" 
strLg245 = "llena su email correctamente" 
strLg246 = "contrase�a" 
strLg247 = "su la contrase�a fue enviada con �xito las paradas "
strLg248 =" los tipos su email abajo para poder enviar su contrase�a." 
strLg249 = "email:" 
strLg250 = "enviar contrase�a"
strLg251 = "- llena los campos al lado." 
strLg252 = "en el informe usted recibir� la contestaci�n Dee su petici�n." 
strLg253 = "equipo m�dico" 
strLg254 = "verdadero resultante"
strLg255 = "on" 
strLg256 = "billete de las actividades bancarias" 
strLg257 = "compra el producto" 
strLg258 = "atenci�n en l�nea" 
strLg259 = "atenci�n fuera de l�nea"
strLg251 = "- llena los campos al lado." 
strLg252 = "en el informe usted recibir� la contestaci�n Dee su petici�n." 
strLg253 = "equipo m�dico" ' 
strLg254 = "verdadero resultante" ' 
strLg255 = "on" ' 
strLg256 = "depositando el billete" ' 
strLg257 = "compra el producto" ' 
strLg258 = "atenci�n en l�nea" ' 
strLg259 = "atenci�n fuera de l�nea"
strLg260 = "Versoes en " ' 
strLg261 = "esta p�gina fue cargado en" ' 
strLg262 = "second(s)" ' 
strLg263 = "visitantes incluso" ' 
strLg264 = "visitante" ' 
strLg265 = "Bienvenido" ' 
strLg266 = "quitar mi email" ' 
strLg267 = "colocar mi email" ' 
strLg268 = "su email fue quitado con �xito!" ' 
strLg269 = "deudor, a partir de hoy en usted no recibir� las ofertas exclusivas de" '
strLg270 = " Maximo "
strLg271 = " producto(s) / cliente "
strLg272 = " unidad "
strLg273 = " Usted est� en "
strLg274 = " M�s Vendidos "
strLg275 = " Email para el amigo "
strLg276 = " Comprar Producto "
strLg277 = "Nos tiemos solamente "
strLg278 = " producto(s) en nuestro almac�n"
strLg279 = "Usted desieja comprar "
strLg280 = "producto en almac�n?"
strLg281 = "El 'CEP' est� diferente do su registro. Tecleo en el Recalcular bot�n"
strLg282 = " Calcular el valor de la entrega"
strLg283 = "Dados para la entrega de la compra"
strLg284 = "Forma de la paga"
strLg285 = "La compra concluida!"
strLg286 = "Cerrando Pedido..."
strLg287 = "Como pagar?"
strLg288 = "Como reimprimir o Boleto?"
strLg289 = "Datos para Conexi�n en el almac�n"
strLg290 = "Datos personales"
strLg291 = "Nombre de su Comprador"
strLg292 = " dias"
strLg293 = " ap�s la Confirmacion de la paga"
strLg294 = "Opci�n Forma de la paga"
strLg295 = "Tecleo su CEP para calculo de la carga y tecleo en el Calcular bot�n"
strLg296 = "Pol�tica de el cambio"

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
strLg314 = " el campo Estado no puede ser vac�o" 

strLg315 = "Novos PRODUTOS est�o sendo catalogados."
strLg316 = "Se voc� desejar algo que ainda n�o encontrou,"
strLg317 = "e nos informe o que est� procurando."
strLg318 = "Informa��es sobre produtos"

'###########################################
'	Textos da loja							
'###########################################

strLgcomprasegura = "&nbsp;&nbsp;&nbsp;&nbsp; El" & nomeloja & "tiene una comisi�n especial con usted, nuestro cliente, cu�nto a la seguridad y a la aislamiento de sus datos. El respecto al cliente y el secreto de su informaci�n son muy importantes para nosotros. Por lo tanto, usted, el cliente " & nomeloja & ", tienen la garant�a que su compra es segura.<br>&nbsp;&nbsp;&nbsp;&nbsp;A " & nomeloja & " en sociedad con el VirtuaWoks un sistema de acci�n para garantizar la seguridad de todas las compras hechas en nuestro almac�n, independientemente del modo elegido del pago. Toda la informaci�n que usted proveer en el total del proceso de la compra es los criptografadas." 

strLgcriptografia = "&nbsp;&nbsp;&nbsp;&nbsp;Todas la informaci�n que pasan para nuestro proceso de la compra autom�ticamente es codificado por un sistema apropiado del criptografia. As�, sus datas personales, el modo elegido del pago y todo y cualquier otra informaci�n proveieron a "& nomeloja & " ser� mantenido secreto. El ' cadeado cerrado ' icono - en la parte inferior de su monitor - durante la orden es un s�mbolo del criptografia de la informaci�n." 

strLgcomprascc = "&nbsp;&nbsp;&nbsp;&nbsp;Voc� podr� solamente a la paga que sus cuentas con la tarjeta de cr�dito ser�n quitar sus compras en nuestro almac�n en Guarulhos-SP, por lo tanto nuestro sitio todav�a no cuenta en el pago en l�nea vio la tarjeta o el debe de cr�dito. Direcci�n del almac�n: Sistema de pesos americano. Guarulhos, n�3900 - Guarulhos - SP - CEP: 07030-001 "

strLginformacoes =" &nbsp;&nbsp;&nbsp;&nbsp;A "&  nomeloja & " no comercializar�n su informaci�n personal. Tal informaci�n se pod�a, sin embargo, agrupar como criterios resueltos y utilizar como los estad�sticos gen�ricos objectifying un acuerdo mejor del perfil del consumidor." 

strLgduvsug = "&nbsp;&nbsp;&nbsp;&nbsp;Caso tiene cualquier duda o la sugerencia en seguridad y aislamiento o ning�n otro aspecto de nuestro proceso de la compra, no licencia en de entrarlas en contacto con. "&  nomeloja & " en sociedad con el VirtuaWoks elabor� un sistema de acci�n para garantizar la seguridad de todas las compras hechas en nuestro almac�n, independientemente del modo elegido del pago. Toda la informaci�n que usted proveer en el total del proceso de la compra es los criptografadas." 

strLgcriptografia = "Todas la informaci�n que pasa para nuestro proceso de la compra autom�ticamente es codificado por un sistema apropiado del criptografia. As�, sus datas personales, el modo elegido del pago y todo y cualquier otra informaci�n proveieron a "&  nomeloja & " ser� mantenido secreto. El ' cadeado cerrado ' icono - en la parte inferior de su monitor - durante la orden es un s�mbolo del criptografia de la informaci�n." 

strLgcomprascc = "usted podr� solamente a la paga que sus cuentas con la tarjeta de cr�dito ser�n quitar sus compras en nuestro almac�n en Guarulhos-SP, por lo tanto nuestro sitio todav�a no cuenta en el pago en l�nea vio la tarjeta o el debe de cr�dito. Direcci�n del almac�n: Sistema de pesos americano. Guarulhos, n�3900 - Guarulhos - SP - CEP: 07030-001 "

strLginformacoes =" "&  nomeloja & " no comercializar�n su informaci�n personal. Tal informaci�n se pod�a, sin embargo, agrupar como criterios resueltos y utilizar como los estad�sticos gen�ricos objectifying un acuerdo mejor del perfil del consumidor."

strLgduvsug = "si hay cualquier duda o sugerencia en seguridad y aislamiento o cualquier otro aspecto de nuestro proceso de la compra, no licencia en de entrarlos en contacto con."

strLgquemsomos = "&nbsp;&nbsp;&nbsp;&nbsp;El Metasupri son una compa��a de la inform�tica dedicada a la buena atenci�n de todos sus clientes <br><br>&nbsp;&nbsp;&nbsp;&nbsp;Atualmente con un almac�n en Guarulhos-SP en el Guarulhos av., n� 3900, el Metasupri hacen uso servicios como venda de computadoras nuevas y de mitad-nuevo y de accesorio usados, celulares, mantenimiento de computadoras, las impresoras, los monitores, telefon�a, las manijas, los adaptadores y muy nosotros trabajo de mais.<br><br>&nbsp;&nbsp;&nbsp;We con los productos de las mejores marcas y modelos con los pocos precios de la regi�n. Hace un presupuesto sin la comisi�n y nosotros financiate de la lata del confira!<br><br>&nbsp;&nbsp;&nbsp;Puedes financiar sus compras para el Banco Real (ABN AMRO) o Panamericano en hasta 24 veces. El " & nomeloja & " en sociedad con el VirtuaWoks un sistema de acci�n para garantizar la seguridad de todas las compras hechas en nuestro almac�n, independientemente del modo elegido del pago. Toda la informaci�n que usted proveer en el total del proceso de la compra es los criptografadas. <br>&nbsp;&nbsp;&nbsp;&nbsp;<br><img align='center' src='linguagens/portuguesbr/imagens/logos.gif'>"
%>