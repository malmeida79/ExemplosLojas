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

'IN�CIO DO C�DIGO

'Este c�digo est� otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP

%>
<!-- #include file="topo.asp" -->
	 	 <SCRIPT LANGUAGE="JavaScript">
			  <!--
			  function paises(){
			  var currentState = document.registro1.outropais.checked == false; 
			  var newState = document.registro1.outropais.checked == true;
			  if (newState != currentState){
			  	 document.registro1.cepzz.disabled = newState;
				 document.registro1.estado.disabled = newState;
				 document.registro1.pais.disabled = !newState;
			 if (document.registro1.pais.disabled = !newState) {
			 	document.registro1.pais.value = "Brasil";
				}
			  }
			 }
			 
 function veremail(semail){
   var bvalido = false;
   var theStr = new String(semail.value);
   var index = theStr.indexOf("@");
   if (index > 0)
   {
    var pindex = theStr.indexOf(".",index);
    if ((pindex > index+1) && (theStr.length > pindex+1))
	bvalido = true;
   }
   if (!bvalido){
      alert("Formato de E-Mail Inv�lido !");      
      document.registro1.email.focus();       
   }
   return bvalido;
 }

 function vernome(snome){
  if (snome.value.substring(0,9) == 'Visitante'){
    alert('DIGITE O NOME COMPLETO !');
    return false;    
  }
  else{
  return true;
  }
 }

 function verCPF(sCPF){
  sCPF.value = SoNumeros(sCPF.value)
  if (sCPF.value==''){
    alert('DIGITE O CPF ou CNPJ !');      
    document.registro1.cpf.focus();
    return false;
  }
  else{
    document.registro1.cpf.value = sCPF.value;    
    return true;
  }
  return true;         
 }

 function SoNumeros(strCampo) {
   nTamanho = strCampo.length;
   szCampo = "";
   for (i = nTamanho-1;i>=0;i--){
     if (isDigit(strCampo.charAt(i))){
	szCampo = strCampo.charAt(i) + szCampo;
     }
    }
    return szCampo;
 }

 function isDigit (c){   
    return ((c >= "0") && (c <= "9"))
 }   			 
			// -->
	  </SCRIPT>

<%strNome = ""%>
	  		   <table><tr><td align="left" valign="top">
			   				  <table border="0" cellspacing="4" cellpadding="4" width="100%"><tr><td><font face="<%=fonte%>" style=font-size:11px;color:000000> <a href=default.asp style=text-decoration:none; onMouseOut="window.status='';return true;" onMouseOver="window.status='Home';return true;"><b>Home</b></a> � <%=strLg5%><br><br><div align=left> <table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></div><font style=font-size:10px><%=strLg122%>
							  		 <table cellspacing="1" cellpadding="1" align="center" width=98% font-style:11px><form action=registros.asp method=post name="registro1" onMouseOver="javascript: paises();"><tr><td colspan=2><br><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><b><%=strLg289%></b><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr><tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg231%></td><td><input type=text name=email style=font-family:<%=fonte%>;font-size:11px;color:#000000; size=30 value="<%=strEmail%>" onBlur="veremail(this);"></td></tr>
									 <tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg232%></td><td><input type=password name=senha style=font-family:<%=fonte%>;font-size:11px;color:#000000; value="<%=strSenha%>"></td></tr>
									 <tr><td colspan=2><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><br><b><%=strLg290%><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></b></td></tr>
									 <tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg118%></td><td><input type=text size=40 name=nomecompleto style=font-family:<%=fonte%>;font-size:11px;color:#000000; value="<%=strNome%>" onBlur="vernome(this);" maxlength="30"></td></tr>
									 <tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg117%></td><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><select style="font-family:<%=fonte%>;font-size:11px;color:#000000;" name="nascdia">
									 <option select value="">Dia <option value="01">01<option value="02">02<option value="03">03<option value="04">04<option value="05">05<option value="06">06<option value="07">07<option value="08">08<option value="09">09<option value="10">10<option value="11">11<option value="12">12<option value="13">13<option value="14">14<option value="15">15<option value="16">16 <option value="17">17 <option value="18">18<option value="19">19<option value="20">20 <option value="21">21 <option value="22">22<option value="23">23<option value="24">24 <option value="25">25<option value="26">26<option value="27">27<option value="28">28<option value="29">29<option value="30">30 <option value="31">31</option></select> <b>/</b> <select style="font-family:<%=fonte%>;font-size:11px;color:#000000;" name="nascmes"> <option value=""> M�s <option value="Jan">Jan <option value="Fev">Fev<option value="Mar">Mar<option value="Abr">Abr<option value="Mai">Mai <option value="Jun">Jun <option value="Jul">Jul<option value="Ago">Ago<option value="Set">Set<option value="Out">Out<option value="Nov">Nov<option value="Dez">Dez</option></select> <b>/</b> <input style="font-family:<%=fonte%>;font-size:11px;color:#000000;" maxLength="4" size="5" name="nascano" value="<%=strNascano%>"></font></td></tr>
									 <tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg116%><sup><font size="3" color="#FF0000">*</font></sup></td><td><input type=text size=20 name=cpf style=font-family:<%=fonte%>;font-size:11px;color:#000000; value="<%=strCpf%>" onBlur="verCPF(this);" maxlength="50"></td></tr>
									 <tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg115%><sup><font size="3" color="#FF0000">*</font></sup></td><td><input type=text size=20 name=rg style=font-family:<%=fonte%>;font-size:11px;color:#000000; value="<%=strRg%>" maxlength="50"></td></tr>
									 <tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg114%></td><td><input type=text size=35 name=endereco style=font-family:<%=fonte%>;font-size:11px;color:#000000; value="<%=strEndereco%>" maxlength="100"></td></tr>
									 <tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg113%></td><td><input type=text size=35 name=bairro style=font-family:<%=fonte%>;font-size:11px;color:#000000; value="<%=strBairro%>" maxlength="50"></td></tr>
									 <tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg112%></td><td><input type=text size=35 name=cidade style=font-family:<%=fonte%>;font-size:11px;color:#000000; value="<%=strCidade%>" maxlength="50"></td></tr>
									 <tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg111%></td><td><select style=font-family:<%=fonte%>;font-size:11px;color:#000000; name="estado"> <option value="AC">AC<option value="AL">AL<option value="AM">AM<option value="AP">AP<option value="BA">BA<option value="CE">CE<option value="DF">DF<option value="ES">ES<option value="GO">GO<option value="MA">MA<option value="MG">MG<option value="MS">MS<option value="MT">MT<option value="PA">PA<option value="PB">PB<option value="PE">PE<option value="PI">PI<option value="PR">PR<option value="RJ">RJ<option value="RN">RN<option value="RO">RO<option value="RR">RR<option value="RS">RS<option value="SC">SC<option value="SE">SE<option value="SP">SP<option value="TO">TO</option></select><font face="<%=fonte%>" style=font-family:<%=fonte%>;font-size:11px;color:#000000;>  <input type="checkbox" name="outropais" value="Sim" onclick="javascript: paises();">Outro Pa�s</td></tr>
									 <tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg82%></td><td><input type=text size=13 name=cepzz style=font-family:<%=fonte%>;font-size:11px;color:#000000; value="<%=strCep%>" maxlength="8">&nbsp;&nbsp;&nbsp;

<!-- Link Antigo:  <a href=http://www.correios.com.br/servicos/cep/cep_default.cfm target=_novocep> -->
<a href="busca_cep.asp" target="novocep" onclick="window.open('','novocep','width=400,height=400,resizable=yes,scrollbars=yes')"><font face="<%=fonte%>" style=font-size:11px;color:000000><u>N�o sabe seu CEP?</u></font></a></td></tr>
									 <tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg110%></td><td><select style=font-family:<%=fonte%>;font-size:11px;color:#000000; name="pais">
									 <option value="Brasil">Brasil</option><option value="Afeganist�o">Afeganist�o</option><option value="�frica do Sul">�frica do Sul</option><option value="Aland - Finl�ndia">Aland - Finl�ndia</option><option value="Alb�nia">Alb�nia</option><option value="Alemanha">Alemanha</option><option value="Andorra">Andorra</option><option value="Angola">Angola</option><option value="Anguilla - Reino Unido">Anguilla - Reino Unido</option><option value="Ant�rtida">Ant�rtida</option><option value="Ant�gua e Barbuda">Ant�gua e Barbuda</option><option value="Antilhas Holandesa">Antilhas Holandesas</option><option value="Ar�bia Saudita">Ar�bia Saudita</option><option value="Arg�lia">Arg�lia</option><option value="Argentina">Argentina</option><option value="Arm�nia">Arm�nia</option><option value="Aruba - Holanda">Aruba - Holanda</option><option value="Ascens�o - Reino Unido">Ascens�o - Reino Unido</option><option value="Austr�lia">Austr�lia</option><option value="�ustria">�ustria</option><option value="Azerbaij�o">Azerbaij�o</option><option value="Bahamas">Bahamas</option><option value="Bahrein">Bahrein</option><option value="Bangladesh">Bangladesh</option><option value="Barbados">Barbados</option><option value="Belarus">Belarus</option><option value="B�lgica">B�lgica</option><option value="Belize">Belize</option><option value="Benin">Benin</option><option value="Bermudas - Reino Unido">Bermudas - Reino Unido</option><option value="Bioko - Guin� Equatorial">Bioko - Guin� Equatorial</option>
									 <option value="Bol�via">Bol�via</option><option value="B�snia-Herzeg�vina">B�snia-Herzeg�vina</option><option value="Botsuana">Botsuana</option><option value="Brunei">Brunei</option><option value="Bulg�ria">Bulg�ria</option><option value="Burkina Fasso">Burkina Fasso</option><option value="Burundi">Burundi</option><option value="But�o">But�o</option><option value="Cabo Verde">Cabo Verde</option><option value="Camar�es">Camar�es</option><option value="Camboja">Camboja</option><option value="Canad�">Canad�</option><option value="Cazaquist�o">Cazaquist�o</option><option value="Ceuta - Espanha">Ceuta - Espanha</option><option value="Chade">Chade</option><option value="Chile">Chile</option><option value="China">China</option><option value="Chipre">Chipre</option><option value="Cidade do Vaticano">Cidade do Vaticano</option><option value="Cingapura">Cingapura</option><option value="Col�mbia">Col�mbia</option><option value="Congo">Congo</option><option value="Cor�ia do Norte">Cor�ia do Norte</option><option value="Cor�ia do Sul">Cor�ia do Sul</option><option value="C�rsega - Fran�a">C�rsega - Fran�a</option><option value="Costa do Marfim">Costa do Marfim</option><option value="Costa Rica">Costa Rica</option><option value="Creta - Gr�cia">Creta - Gr�cia</option><option value="Cro�cia">Cro�cia</option><option value="Cuba">Cuba</option><option value="Cura�ao - Holanda">Cura�ao - Holanda</option><option value="Dinamarca">Dinamarca</option>
									 <option value="Djibuti">Djibuti</option><option value="Dominica">Dominica</option><option value="Egito">Egito</option><option value="El Salvador">El Salvador</option><option value="Emirado �rabes Unidos">Emirado �rabes Unidos</option><option value="Equador">Equador</option><option value="Eritr�ia">Eritr�ia</option><option value="Eslov�quia">Eslov�quia</option><option value="Eslov�nia">Eslov�nia</option><option value="Espanha">Espanha</option><option value="Estados Unidos">Estados Unidos</option><option value="Est�nia">Est�nia</option><option value="Eti�pia">Eti�pia</option><option value="Fiji">Fiji</option><option value="Filipinas">Filipinas</option><option value="Finl�ndia">Finl�ndia</option><option value="Fran�a">Fran�a</option><option value="Gab�o">Gab�o</option><option value="G�mbia">G�mbia</option><option value="Gana">Gana</option><option value="Ge�rgia">Ge�rgia</option><option value="Gibraltar - Reino Unido">Gibraltar - Reino Unido</option><option value="Granada">Granada</option><option value="Gr�cia">Gr�cia</option><option value="Groenl�ndia - Dinamarca">Groenl�ndia - Dinamarca</option><option value="Guadalupe - Fran�a">Guadalupe - Fran�a</option><option value="Guam - Estados Unidos">Guam - Estados Unidos</option><option value="Guatemala">Guatemala</option><option value="Guiana">Guiana</option><option value="Guiana Francesa">Guiana Francesa</option><option value="Guin�">Guin�</option><option value="Guin� Equatorial">Guin� Equatorial</option>
									 <option value="Guin�-Bissau">Guin�-Bissau</option><option value="Haiti">Haiti</option><option value="Holanda">Holanda</option><option value="Honduras">Honduras</option><option value="Hong Kong">Hong Kong</option><option value="Hungria">Hungria</option><option value="I�men">I�men</option><option value="IIhas Virgens - Estados Unidos">IIhas Virgens - Estados Unidos</option><option value="Ilha de Man - Reino Unido">Ilha de Man - Reino Unido</option><option value="Ilha Natal - Austr�lia">Ilha Natal - Austr�lia</option><option value="Ilha Norfolk - Austr�lia">Ilha Norfolk - Austr�lia</option><option value="Ilha Pitcairn - Reino Unido">Ilha Pitcairn - Reino Unido</option><option value="Ilha Wrangel - R�ssia">Ilha Wrangel - R�ssia</option><option value="Ilhas Aleutas - Estados Unidos">Ilhas Aleutas - Estados Unidos</option><option value="Ilhas Can�rias - Espanha">Ilhas Can�rias - Espanha</option><option value="Ilhas Cayman - Reino Unido">Ilhas Cayman - Reino Unido</option><option value="Ilhas Comores">Ilhas Comores</option><option value="Ilhas Cook - Nova Zel�ndia">Ilhas Cook - Nova Zel�ndia</option><option value="Ilhas do Canal - Reino Unido">Ilhas do Canal - Reino Unido</option><option value="Ilhas Salom�o">Ilhas Salom�o</option><option value="Ilhas Seychelles">Ilhas Seychelles</option><option value="Ilhas Tokelau - Nova Zel�ndia">Ilhas Tokelau - Nova Zel�ndia</option><option value="Ilhas Turks e Caicos - Reino Unido">Ilhas Turks e Caicos - Reino Unido</option>
									 <option value="Ilhas Virgens - Reino Unido">Ilhas Virgens - Reino Unido</option><option value="Ilhas Wallis e Futuna - Fran�a">Ilhas Wallis e Futuna - Fran�a</option><option value="Ilhsa Cocos - Austr�lia">Ilhsa Cocos - Austr�lia</option><option value="�ndia">�ndia</option><option value="Indon�sia">Indon�sia</option><option value="Ir�">Ir�</option><option value="Iraque">Iraque</option><option value="Irlanda">Irlanda</option><option value="Isl�ndia">Isl�ndia</option><option value="Israel">Israel</option><option value="It�lia">It�lia</option><option value="Iugosl�via">Iugosl�via</option><option value="Jamaica">Jamaica</option><option value="Jan Mayen - Noruega">Jan Mayen - Noruega</option><option value="Jap�o">Jap�o</option><option value="Jord�nia">Jord�nia</option><option value="Kiribati">Kiribati</option><option value="Kuait">Kuait</option><option value="Laos">Laos</option><option value="Lesoto">Lesoto</option><option value="Let�nia">Let�nia</option><option value="L�bano">L�bano</option><option value="Lib�ria">Lib�ria</option><option value="L�bia">L�bia</option><option value="Liechtenstein">Liechtenstein</option><option value="Litu�nia">Litu�nia</option><option value="Luxemburgo">Luxemburgo</option><option value="Macau - Portugal">Macau - Portugal</option><option value="Maced�nia">Maced�nia</option><option value="Madagascar">Madagascar</option><option value="Madeira - Portugal">Madeira - Portugal</option><option value="Mal�sia">Mal�sia</option><option value="Malaui">Malaui</option>
									 <option value="Maldivas">Maldivas</option><option value="Mali">Mali</option><option value="Malta">Malta</option><option value="Marrocos">Marrocos</option><option value="Martinica - Fran�a">Martinica - Fran�a</option><option value="Maur�cio - Reino Unido">Maur�cio - Reino Unido</option><option value="Maurit�nia">Maurit�nia</option><option value="M�xico">M�xico</option><option value="Micron�sia">Micron�sia</option><option value="Mo�ambique">Mo�ambique</option><option value="Moldova">Moldova</option><option value="M�naco">M�naco</option><option value="Mong�lia">Mong�lia</option><option value="MontSerrat - Reino Unido">MontSerrat - Reino Unido</option><option value="Myanma">Myanma</option><option value="Nam�bia">Nam�bia</option><option value="Nauru">Nauru</option><option value="Nepal">Nepal</option><option value="Nicar�gua">Nicar�gua</option><option value="N�ger">N�ger</option><option value="Nig�ria">Nig�ria</option><option value="Niue">Niue</option><option value="Noruega">Noruega</option><option value="Nova Bretanha - Papua-Nova Guin�">Nova Bretanha - Papua-Nova Guin�</option><option value="Nova Caled�nia - Fran�a">Nova Caled�nia - Fran�a</option><option value="Nova Zel�ndia">Nova Zel�ndia</option><option value="Om�">Om�</option><option value="Palau - Estados Unidos">Palau - Estados Unidos</option><option value="Palestina">Palestina</option><option value="Panam�">Panam�</option><option value="Papua-Nova Guin�">Papua-Nova Guin�</option><option value="Paquist�o">Paquist�o</option>
									 <option value="Paraguai">Paraguai</option><option value="Peru">Peru</option><option value="Polin�sia Francesa">Polin�sia Francesa</option><option value="Pol�nia">Pol�nia</option><option value="Porto Rico">Porto Rico</option><option value="Portugal">Portugal</option><option value="Qatar">Qatar</option><option value="Qu�nia">Qu�nia</option><option value="Quirguist�o">Quirguist�o</option><option value="Reino Unido">Reino Unido</option><option value="Rep�blica Centro-Africana">Rep�blica Centro-Africana</option><option value="Rep�blica Dominicana">Rep�blica Dominicana</option><option value="Rep�blica Tcheca">Rep�blica Tcheca</option><option value="Rom�nia">Rom�nia</option><option value="Ruanda">Ruanda</option><option value="R�ssia">R�ssia</option><option value="Samoa Ocidental">Samoa Ocidental</option><option value="San Marino">San Marino</option><option value="Santa Helena - Reino Unido">Santa Helena - Reino Unido</option><option value="Santa L�cia">Santa L�cia</option><option value="S�o Cristov�o e N�vis">S�o Cristov�o e N�vis</option><option value="S�o Tom� e Pr�ncipe">S�o Tom� e Pr�ncipe</option><option value="S�o Vicente e Granadinas">S�o Vicente e Granadinas</option><option value="Sardenha - It�lia">Sardenha - It�lia</option><option value="Senegal">Senegal</option><option value="Serra Leoa">Serra Leoa</option><option value="S�ria">S�ria</option><option value="Som�lia">Som�lia</option><option value="Sri Lanka">Sri Lanka</option><option value="Suazil�ndia">Suazil�ndia</option>
									 <option value="Sud�o">Sud�o</option><option value="Su�cia">Su�cia</option><option value="Su��a">Su��a</option><option value="Suriname">Suriname</option><option value="Tadjiquist�o">Tadjiquist�o</option><option value="Tail�ndia">Tail�ndia</option><option value="Taiti">Taiti</option><option value="Taiwan">Taiwan</option><option value="Tanz�nia">Tanz�nia</option><option value="Terra de Francisco Jos� - R�ssia">Terra de Francisco Jos� - R�ssia</option><option value="Togo">Togo</option><option value="Tonga">Tonga</option><option value="Trinidad e Tobago">Trinidad e Tobago</option><option value="Trist�o da Cunha - Reino Unido">Trist�o da Cunha - Reino Unido</option><option value="Tun�sia">Tun�sia</option><option value="Turcomenist�o">Turcomenist�o</option><option value="Turquia">Turquia</option><option value="Tuvalu">Tuvalu</option><option value="Ucr�nia">Ucr�nia</option><option value="Uganda">Uganda</option><option value="Uruguai">Uruguai</option><option value="Uzbequist�o">Uzbequist�o</option><option value="Vanuatu">Vanuatu</option><option value="Venezuela">Venezuela</option><option value="Vietn�">Vietn�</option><option value="Zaire">Zaire</option><option value="Z�mbia">Z�mbia</option><option value="Zimb�bue">Zimb�bue</option></select></td></tr>
									 <tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg109%></td><td><input type=text size=35 name=fone style=font-family:<%=fonte%>;font-size:11px;color:#000000; value="<%=strFone%>" maxlength="50"> <font style=font-family:<%=fonte%>;font-size:9px;color:#000000;></td></tr>
									 <tr><td><font style=font-family:<%=fonte%>;font-size:11px;color:#000000;><%=strLg291%><sup><font size="3" color="#FF0000">*</font></sup></td><td><input type=text size=20 name=contato style=font-family:<%=fonte%>;font-size:11px;color:#000000; value="<%=strContato%>" maxlength="50"> <font style=font-family:<%=fonte%>;font-size:9px;color:#000000;></td></tr>

									 <tr><td colspan="2"><br><font style=font-family:<%=fonte%>;font-size:11px;color:#FF0000;><sup><font size="3">*</font></sup> tipo de informa��es solicitadas se for pessoa jur�dica</tr>

<tr><td colspan=2><table border=0 cellspacing=0 width=100% cellpadding=0><tr><td height=5></td></tr><tr><td height=1 bgcolor=<%=cor3%>></td></tr><tr><td height=5></td></tr></table></td></tr><tr align=center><td valign=top><div id="layer1" style="position:absolute; left:350px"><input type=image src=<%=dirlingua%>/imagens/prosseguir.gif onMouseOut="window.status='';return true;" onMouseOver="window.status='Prosseguir';return true;"></div></td><td valign=top></form><div id="layer1" style="position:absolute; left:470px"><input type=image src=<%=dirlingua%>/imagens/limpar.gif onMouseOut="window.status='';return true;" onMouseOver="window.status='Limpar';return true;" OnClick="javascript:document.registro1.reset()"></div></td></tr>
									 </table></form></td></tr>
							</table></td></tr>
			  </table>
			  <!-- #include file="baixo.asp" -->