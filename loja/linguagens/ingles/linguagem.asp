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
'#     Internet Design entre em contato atrav�s do e-mail 
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
' Ingl�s
' Pais : Brasil
' HTML Content-Type : ISO-8859-1 
' Traduzido por : Henrique Metasupri (jusivansjc@yahoo.com.br) 09/01/2004
' Acr�scimos/Corre��es por : Elizeu Alcantara (corpsjc@megapaper.com.br / corpsjc@megapaper.com.br ) 15/03/2004 
'======================================

nomeloja = application("nomeloja")

strLgDataS1 = "Jan"            													'Nome dos meses abreviados
strLgDataS2 = "Feb"
strLgDataS3 = "Mar"
strLgDataS4 = "Apr"
strLgDataS5 = "May"
strLgDataS6 = "Jun"
strLgDataS7 = "Jul"
strLgDataS8 = "Aug"
strLgDataS9 = "Set"
strLgDataS10 = "Out"
strLgDataS11 = "Nov"
strLgDataS12 = "Dec"

strLgData1 = "January"														 'Nome dos meses completos
strLgData2 = "February"
strLgData3 = "March"
strLgData4 = "April"
strLgData5 = "May"
strLgData6 = "June"
strLgData7 = "July"
strLgData8 = "August"
strLgData9 = "September"
strLgData10 = "Ouctober"
strLgData11 = "November"
strLgData12 = "December"


'======================================
'	VALOR DAS STRINGS DA LOJA          
'======================================

strLg1 = "To finish Purchases" 
strLg2 = "Amount of itens:" 
strLg3 = "Subtotal:" 
strLg4 = "Home" 
strLg5 = "Register of Customers" 
strLg6 = "Login" 
strLg7 = "Stand of Purchases of"
strLg8 = "To close Asked for" 
strLg9 = "My Data" 
strLg10 = "Historical of Purchases" 
strLg11 = "Logout" 
strLg12 = "Departments"
strLg13 = "Research" 
strLg14 = "Attendance"
strLg15 = "All the categories" 
strLg16 = "As To buy" 
strLg17 = "by email" 
strLg18 = "On Security" 
strLg19 = "Security and Privacy" 
strLg20 = "Who we are" 
strLg21 = "To search"
strLg22 = "On Security" 
strLg23 = "By email" 
strLg24 = "Attendance by email" 
strLg25 = "for delivery" 
strLg26 = "Available for delivery" 
strLg27 = "Not available for delivery" 
strLg28 = "Supply:" 
strLg29 = "Price:" 
strLg30 = "+ Details" 
strLg31 = "Manufacturer:" 
strLg32 = "To add" 
strLg33 = "product(s) in my stand of purchases." 
strLg34 = "Buying this product" 
strLg35 = "in" 
strLg36 = "you even save" 
strLg37 = "To the sight for:" 
strLg38 = "for:" 
strLg39 = "of:"
strLg40 = "or still:" 
strLg41 = "To come back" 
strLg42 = "Searching for" 
strLg43 = "Had been find" 
strLg44 = "product(s):" 
strLg45 = "Page" 
strLg46 = "of" 
strLg47 = "Next Page" 
strLg48 = "Products in this department:" 
strLg49 = "does not have no item in its stand of purchases." 
strLg50 = "If you its session of purchases bought some product previously is exhaled" 
strLg51 = "To see my stand of purchases" 
strLg52 = "More Offers:"
strLg53 = "the item was added in its stand of purchases." 
strLg54 = "Purchase" 
strLg55 = "Its purchase" 
strLg56 = "Please, verifies if its purchase is correct before continuing." 
strLg57 = "to modify the amount of some item types the new value and click in the button to bring up to date." 
strLg58 = "Amount" 
strLg59 = "Product" 
strLg60 = "Unitary Price" 
strLg61 = "Total Price" 
strLg62 = "To remove" 
strLg63 = "Total of the purchase" 
strLg64 = "Value of the delivery in" 
strLg65 = "Its purchase" 
strLg66 = "To bring up to date" 
strLg67 = "To continue" 
strLg68 = "To continue Purchases" 
strLg69 = "Choice the region of delivery for I calculate it of the freight" 
strLg70 = "Click here" 
strLg71 = "Click here" 
strLg72 = "Given for delivery of the order"
strLg73 = "to modify its purchase" 
strLg74 = "Please, verifies if the address for delivery is correct before continuing." 
strLg75 = "if address for the delivery and the same that its" 
strLg76 = "Options for delivery" 
strLg77 = "Address for delivery" 
strLg78 = "Address:" 
strLg79 = "Quarter:" 
strLg80 = "City:" 
strLg81 = "Been:" 
strLg82 = "ZIP:" 
strLg83 = "Country:" 
strLg84 = "Telephone:" 
strLg85 = "You desire that itens bought is packed and envoy as present"
strLg86 = "Yes" 
strLg87 = "Not" 
strLg88 = "Message written in the card-gift:" 
strLg89 = "To clean" 
strLg90 = "Please, fills all correctly the data before continuing." 
strLg91 = "Mode of payment" 
strLg92 = "Please, fills all correctly the data before continuing." 
strLg93 = "to modify the address of delivery of the purchases and/or other data" 
strLg94 = "Selects the mode of payment chosen:" 
strLg95 = "Other modes of payment" 
strLg96 = "Payment saw Credit card" 
strLg97 = "Selects its card:"

strLg98 = "Number of the card: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Secure Code:"

strLg99 = "Expiration of the card:" 
strLg100 = "To cancel" 
strLg101 = "Year" 
strLg102 = "Titular of the card (As it appears in the Card):" 
strLg103 = "My Data" 
strLg104 = "to modify its data, selects the field desired and writes the data brought up to date on the old one." 
strLg105 = "For no modification click in ' Cancelling '. " 
strLg106 = "Given of the Register" 
strLg107 = "Personal data" 
strLg108 = "Your email:" 
strLg109 = "Your telephone:" 
strLg110 = "Your country:" 
strLg111 = "Your state:" 
strLg112 = "Your city:" 
strLg113 = "Your quarter:" 
strLg114 = "Your address:" 
strLg115 = "Your RG / IE" 
strLg116 = "Your CPF or CNPJ" 
strLg117 = "Your birth:" 
strLg118 = "Your full name:" 
strLg119 = "Purchases effected by" 
strLg120 = "to visualize the data of its purchase types the number of the order soon below:" 
strLg121 = "Asked for n.:" 
strLg122 = "For its comfort and security verifies all its data before continuing."
strLg123 = "Your Password:" 
strLg124 = "Your email was sent successfully" 
strLg125 = "Obliged" 
strLg126 = "in briefing you will receive the reply from its doubt" 
strLg127 = "Now enjoys of the comfort and security to buy in our virtual store." 
strLg128 = "As To buy in" 
strLg129 = "Its Name:" 
strLg130 = "Types your email correctly so that this message can be answered." 
strLg131 = "Doubt on:" 
strLg132 = "Subject:" 
strLg133 = "Message:" 
strLg134 = "As to use the Stand of Purchases" 
strLg135 = "As to search products in the store" 
strLg136 = "As he is I register in cadastre it of customers" 
strLg137 = "As to effect login in the store" 
strLg138 = "As I modify my data" 
strLg139 = "diverse Doubts" 
strLg140 = "To send" 
strLg141 = "Fills this field" 
strLg142 = "the CPF must have 11 digits" 
strLg143 = "Invalid CPF / CNPJ" 
strLg144 = "Fills this field correctly" 
strLg145 = "existing email already in our registers" 
strLg146 = "To modify my email or password" 
strLg147 = "Day" 
strLg148 = "Month" 
strLg149 = "Modified Register" 
strLg150 = "Its register was modified successfully" 
strLg151 = "All its data had been modified and brought up to date successfully in our database." 
strLg152 = "Types your email correctly" 
strLg153 = "To close this Window" 
strLg154 = "Its email was registered successfully" 
strLg155 = "Obliged, now you will receive offers exclusive of" 
strLg156 = "This item already exists in its stand of purchases." 
strLg157 = "the item was added in its stand of purchases." 
strLg158 = "Its first time here" 
strLg159 = "Types your email and recive exclusives offers." 
strLg160 = "Types your email" 
strLg161 = "Is registered in cadastre Already"
strLg162 = "This email already belongs to another registered in cadastre customer" 
strLg163 = "the typed password does not confer"
strLg164 = "For modification of the data click in ' Bringing up to date ', for no modification click in ' Cancelling '." 
strLg165 = "is only allowed the alteration of an item for time." 
strLg166 = "Given of its email" 
strLg167 = "if you want to modify its email" 
strLg168 = "Its current email:" 
strLg169 = "Types its new email:" 
strLg170 = "Given of its password" 
strLg171 = "if you want to modify its password" 
strLg172 = "Types its password current:" 
strLg173 = "Types its new password:" 
strLg174 = "the field Address cannot be empty" 
strLg175 = "the field Quarter cannot be empty" 
strLg176 = "the field City cannot be empty" 
strLg177 = "field ZIP cannot be empty" 
strLg178 = "the field Telephone cannot be empty" 
strLg179 = "Login in the Store" 
strLg180 = "If is its first one visits this store" 
strLg181 = "If already is our registered customer then introduces the data." 
strLg182 = "You forgot its password" 
strLg183 = "email of user not found" 
strLg184 = "the Password is incorrect" 
strLg185 = "Selects its region" 
strLg186 = "Information on order " 
strLg187 = "This purchase is not present in the database of the store" 
strLg188 = "or is not referring to this customer" 
strLg189 = "Purchase effected by" 
strLg190 = "SEDEX to charging" 
strLg191 = "Bank Deposit" 
strLg192 = "Eletronic Transfer" 
strLg193 = "Waiting confirmattion payment" 
strLg194 = "Payment OK and purchase in processing" 
strLg195 = "sent Purchase already to the customer" 
strLg196 = "Purchase Already Delivers" 
strLg197 = "Cancelled Purchase" 
strLg198 = "Denied Purchase" 
strLg199 = "Others - the Attendance Contacts" 
strLg200 = "Freight"
strLg201 = "Status of the Purchase" 
strLg202 = "Date of the Purchase" 
strLg203 = "Estimative for delivery" 
strLg204 = "You still did not effect no purchase in" 
strLg205 = "To see information of this order" 
strLg206 = "Inexistent Department" 
strLg207 = "inexistent Department in" 
strLg208 = "No present product in this department" 
strLg209 = "Previous Page" 
strLg210 = "- Invalid Card." 
strLg211 = "- Invalid Number." 
strLg212 = "- Invalid Number and Card." 
strLg213 = "- Invalid Addition of Control." 
strLg214 = "- Addition of Invalid Control and Card." 
strLg215 = "- Addition of Invalid Control and Number." 
strLg216 = "- Invalid Addition of Control, Number and Card" 
strLg217 = "- the number of the credit card Types" 
strLg218 = "- a date Selects valid" 
strLg219 = "- Date of the invalid card." 
strLg220 = "- the name of the bearer Types (as it appears in the card)" 
strLg221 = "Credit card MASTERCARD" 
strLg222 = "Credit card AIMS AT" 
strLg223 = "Credit card AMERICAN EXPRESS" 
strLg224 = "Credit card DINNERS" 
strLg225 = "SEDEX to charging" 
strLg226 = "Bank Deposit" 
strLg227 = "Eletronic Transfer"
strLg228 = "No product was found in this search." 
strLg229 = "Inexistent Product" 
strLg230 = "Inexistent Product in " 
strLg231 = "Types your email:" 
strLg232 = "Password:" 
strLg233 = "Incorrect" 
strLg234 = "Its password must have in the maximum 10 characters" 
strLg235 = "Its password must have 5 characters at the very least" 
strLg236 = "Its register was made successfully" 
strLg237 = "Obliged" 
strLg238 = "now enjoys of the comfort and security to buy in" 
strLg239 = "Safe Purchase" 
strLg240 = "Cryptografy" 
strLg241 = "Purchases with Credit card" 
strLg242 = "Use of Information" 
strLg243 = "Doubts and Suggestions" 
strLg244 = "email not found in our registers" 
strLg245 = "Fills its email correctly" 
strLg246 = "Password" 
strLg247 = "Its password was sent successfully stops" 
strLg248 = "Types your email below so that its password can be sent." 
strLg249 = "email:" 
strLg250 = "To send Password"
strLg251 = "- It fills the Fields to the side." 
strLg252 = "In briefing you will receive reply Dee its request." 
strLg253 = "Medical Team" 
strLg254 = "Resulted Real" 
strLg255 = "On" 
strLg256 = "Banking Billet" 
strLg257 = "Buys the product" 
strLg258 = "On-Line Attendance" 
strLg259 = "Off-Line Attendance" 
strLg260 = "Version in "
strLg261 = "This page was loaded in "
strLg262 = " second(s)"
strLg263 = "Visitors even "
strLg264 = "Visitor "
strLg265 = "Welcome "
strLg266 = "To remove my email "
strLg267 = "To register my email "
strLg268 = "Its email was removed successfully!"
strLg269 = "Debtor, from today on you will not receive offers exclusive of "
strLg270 = " Maximum "
strLg271 = " product(s) / client "
strLg272 = " unit "
strLg273 = " You are in "
strLg274 = " Top Sellers "
strLg275 = " Email to friend "
strLg276 = " Buy Product "
strLg277 = "We have only "
strLg278 = " product(s) in our stock"
strLg279 = "Do you wish to buy "
strLg280 = "product of the stock?"
strLg281 = "The 'ZIP' filled is different of your profile. Click in the Recalcular button"
strLg282 = " Calculate Value of the delivery"
strLg283 = "Information for delivery of the order"
strLg284 = "Mode of payment"
strLg285 = "Order finish!"
strLg286 = "Order in progress..."
strLg287 = "How to pay?"
strLg288 = "How to reprint Billet?"
strLg289 = "Login data"
strLg290 = "Profile data"
strLg291 = "Buyer name"
strLg292 = " days"
strLg293 = " after Confirmattion payment" 
strLg294 = "Choice mode of payment"
strLg295 = "Types your ZIP to calculate Freight and click Atualizar button"
strLg296 = "Changes politics"

strLg297= "Payment"
strLg298= "Bank Deposit"
strLg299= "Bank EletronicDeposit"
strLg300= "Bank Deposit (By DOC)"
strLg301= "Eletronic Transfer"
strLg302= "Banking Billet"
strLg303= "Date: "
strLg304= "Observations: "
strLg305= "Hour: "
strLg306= "Bank Agency: "
strLg307= "Payment value: "
strLg308= "next days ypou will receive Payment confirmation!"
strLg309= "Dentro de 1 a 2 dias �teis estaremos confirmando seu pagamento por e-mail.<br>Caso necessite de outras informa��es, entre em contato conosco pelo e-mail "&Email_da_sua_loja
strLg310= "My Shopcarto"
strLg311= "Confirm payment"
strLg312= "Seller politics"
strLg313= "How to confirm payment"
strLg314 = "the field State cannot be empty"

strLg315 = "Novos PRODUTOS est�o sendo catalogados."
strLg316 = "Se voc� desejar algo que ainda n�o encontrou,"
strLg317 = "e nos informe o que est� procurando."
strLg318 = "Informa��es sobre produtos"

'###########################################
'	Store's Text								
'###########################################

strLgcomprasegura = "&nbsp;&nbsp;&nbsp;&nbsp;The " & nomeloja & " has a special commitment with you, our customer, how much to the security and privacy of its data. The respect to the customer and the secrecy of its information are very important for us. Therefore, you, customer of "& nomeloja &", have the guarantee that its purchase is segura.<br>&nbsp;&nbsp;&nbsp;&nbsp;A "& nomeloja &" in partnership with the VirtuaWoks elaborated a set of action to guarantee the security of all the purchases made in our store, independently of the chosen mode of payment. All the information that you to supply in the purchase process total are criptografadas." 
strLgcriptografia = "&nbsp;&nbsp;&nbsp;&nbsp;Todas the information that pass for our process of purchase automatically is codified by a proper system of criptografia. Thus, its personal datas, the chosen mode of payment and all and any another information supplied to "& nomeloja &" will be kept in secrecy. The closed ' cadeado ' icon - in the inferior part of its monitor - during the order is a symbol of the criptografia of the information." 
strLgcomprascc = "&nbsp;&nbsp;&nbsp;&nbsp;Voc� will only be able to pay its accounts with credit card will be to remove its purchases in our store in Guarulhos-SP, therefore our site still does not count on payment on-line saw credit card or debit. Address of the store: Av. Guarulhos, n�3900 - Guarulhos - SP - CEP: 07030-001 " 
strLginformacoes = "&nbsp;&nbsp;&nbsp;&nbsp;A" & nomeloja & "will not commercialize its personal information. Such information could, however, be grouped as determined criteria and be used as generic statisticians objectifying one better agreement of the profile of the consumer." 
strLgduvsug = "&nbsp;&nbsp;&nbsp;&nbsp;Caso has any doubt or suggestion on security and privacy or any another aspect of our process of purchase, does not leave of in contacting them. A "& nomeloja &" in partnership with the VirtuaWoks elaborated a set of action to guarantee the security of all the purchases made in our store, independently of the chosen mode of payment. All the information that you to supply in the purchase process total are criptografadas."
strLgcriptografia = "    Todas the information that pass for our process of purchase automatically is codified by a proper system of criptografia. Thus, its personal datas, the chosen mode of payment and all and any another information supplied to "& nomeloja &" will be kept in secrecy. The closed ' cadeado ' icon - in the inferior part of its monitor - during the order is a symbol of the criptografia of the information." 
strLgcomprascc = "    You will only be able to pay its accounts with credit card will be to remove its purchases in our store in Guarulhos-SP, therefore our site still does not count on payment on-line saw credit card or debit. Address of the store: Av. Guarulhos, n�3900 - Guarulhos - SP - CEP: 07030-001 " 
strLginformacoes = "    The " & nomeloja & " will not commercialize its personal information. Such information could, however, be grouped as determined criteria and be used as generic statisticians objectifying one better agreement of the profile of the consumer." 

strLgduvsug = "If there is any doubt or suggestion on security and privacy or any another aspect of our process of purchase, does not leave of in contacting them."

strLgquemsomos = "&nbsp;&nbsp;&nbsp;&nbsp;The Metasupri is a Company of dedicated Computer science to the good attendance of all its Customers < br><br>&nbsp;&nbsp;&nbsp;&nbsp;Atualmente with a store in Guarulhos-SP in the Av. Guarulhos, n� 3900, the Metasupri makes use of services as venda of new computers and used, cellular half-new and accessory, maintenance of computers, printers, monitors, telephony, handles, adapters and very we mais.<br><br>&nbsp;&nbsp;&nbsp;We work with products of the best marks and models with the lesser prices of the region. It makes a budget without commitment and we confira!<br><br>&nbsp;&nbsp;&nbsp;We can financiate its purchases for the Real Bank (ABN AMRO) or Panamericano in up to 24 times. We send all for Brazil with the security and rapidity of the Post offices <br><br>&nbsp;&nbsp;&nbsp;Also we are available for Web Design services , Development of systems, Installation and Maintenance of Nets (commercial, residential, lan houses, etc). Contact with us through email or telephone enters in (menu to the side).<br><br> <img align='center' src='linguagens/portuguesbr/imagens/logos.gif'><br>"
%>