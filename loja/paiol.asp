<%@LANGUAGE="VBSCRIPT"%>
<% Option Explicit %>
<!--#include file="_conexao.asp" -->
<!--#include file="_cookie.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<style type="text/css" title="mystyle" media="screen">
<!--

	body{
		background-image: url(fundo.gif);
		}


#corpo{
		background-image: url(corpo.gif);
		position: absolute;
		width: 800px;
		height: 1103px;
		top: 1px;
		margin-top: 0px;
		margin-left:-400px;
		left:50%;
		margin-bottom: 15px;
		}
		
#topo{
		background-color:;
		position: absolute;
		width: 760px;
		height: 180px;
		margin-top: 10px;
		margin-left: 21px;
		}
		
#menu{
		background-color:;
		position: absolute;
		width: 120px;
		height: 423px;
		margin-top: 191px;
		margin-left: 21px;
		}
		
#imagempaiol{
		background-image: url(imagempaiol.gif);
		position: absolute;
		width: 636px;
		height: 87px;
		margin-top: 195px;
		margin-left: 145px;
		}
		
#dataJs{
		background-color:   ;
		position: absolute;
		width: 217px;
		height: 15px;
		margin-top: 0px;
		margin-left: 413px;
		font-family: Verdana, Arial, Helvetica, sans-serif; color: #333333; font-size: 11px;
		text-align: right;
		line-height: 20px;
		}

#textopaiol{
		background-color:;
	position: absolute;
	width: 590px;
	height: 60px;
	margin-top: 280px;
	margin-left: 165px;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	color: #333333;
	font-size: 12px;
	line-height: 14px;
	text-align: justify;
	left: 9px;
	top: 1px;
		}
		
#caixaform{
		background-color:;
		position:absolute;
		width:350px;
		height:180px;
		margin-top:350px;
		margin-left:297px;
		border:1px;
		border-style:dashed;
		border-color: #333333;
		border-bottom-color: #FF3333;
		border-right-color: #FF3333;
		}
		
#caixalogin{
		background-color:;
		position:absolute;
		width:45px;
		height:20px;
		margin-top:35px;
		margin-left: 10px;
		font-family:Verdana, Arial, Helvetica, sans-serif;color: #FF3333;font-size:12px;
		text-align:left;line-height:21px;
		}
		
#caixasenha{
		background-color:;
		position:absolute;
		width:70px;
		height:15px;
		margin-top:70px;
		margin-left: 10px;
		font-family:Verdana, Arial, Helvetica, sans-serif;color: #FF3333;font-size:12px;
		text-align:left;line-height:21px;
		}
#textlogin{
		background-color:;
	position:absolute;
	width:170px;
	height:20px;
	margin-top:35px;
	margin-left:40px;
	left: 32px;
	top: 1px;
		}
		
#textpassword{
		background-color:;
	position:absolute;
	width:170px;
	height:20px;
	margin-top:70px;
	margin-left:40px;
	left: 40px;
		}
		
#caixabutton{
		background-color:;
		position:absolute;
		width:50px;
		height:25px;
		margin-top:100px;
		margin-left:10px;
		}
		
#imagemnewsletter{
		background-image: url(newsletter.gif);
		position: absolute;
		width: 125px;
		height: 450px;
		margin-left: 21px;
		margin-top: 620px;
		}
		
#barra2{
		background-color:;
		position: absolute;
		width: 759px;
		height: 25px;
		margin-top: 1073px;
		margin-left: 20px;
		}
		
#linksrodape{
		background-color:;
		position: absolute;
		width: 475px;
		height: 20px;
		margin-top: 1020px;
		margin-left: 160px;
		font-family: Verdana, Arial, Helvetica, sans-serif; 
		color: #333333;
		font-size: 12px;
		line-height: 20px;
		text-align: center;
		}

		
		
#rodape1{
		background-color:;
		position: absolute;
		width: 330px;
		height: 15px;
		margin-top: 7px;
		margin-left: 10px;
		font-family: Verdana, Arial, Helvetica, sans-serif; color: #333333; font-size: 10px;
		text-align:left; line-height: 12px;
		}
		
#rodape2{
		background-color:;
		position: absolute;
		width: 320px;
		height: 15px;
		margin-top: 7px;
		margin-left: 435px;
		font-family: Verdana, Arial, Helvetica, sans-serif; color: #333333; font-size: 10px;
		text-align: right; line-height: 12px;
		}
a:link {
	color: #333333;
	text-decoration: none;
}
a:visited {
	color: #333333;
	text-decoration: none;
}
a:hover {
	color: #FF3333;
	text-decoration: none;
}
a:active {
	color: #333333;
	text-decoration: none;
}
.style1 {
	color: #FF0000;
	font-weight: bold;
	font-size: 11px;
}
.style2 {
	font-size: 11px;
	font-weight: bold;
}
#conteudopaiol{
		background-color:;
	position: absolute;
	width: 570px;
	height: 40px;
	margin-top: 280px;
	margin-left: 165px;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	color: #333333;
	font-size: 12px;
	line-height: 14px;
	text-align: justify;
	left: 9px;
	top: 81px;
		}
#conteudopaio2{
		background-color:;
	position: absolute;
	width: 400px;
	height: 20px;
	margin-top: 280px;
	margin-left: 165px;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	color: #333333;
	font-size: 12px;
	line-height: 14px;
	text-align: justify;
	left: 10px;
	top: 148px;
		}
.style3 {
	color: #00FF00;
	font-weight: bold;
}




-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>.::.Rede Jovens!! Avivados .::. Paiol</title>
<script language="javascript" src="menu_flash.js"></script>
<script language="javascript" src="topo_flash.js"></script>
</head>
<body>
<div id="corpo">
		<div id="topo">
		 <script>Topo_Flash();</script>
		</div>
		<div id="menu">
		 <script>Menu_Flash();</script> 
		</div>
		<div id="imagempaiol">
				<div id="dataJs"><script language="Javascript"> 
function Hoje() { 
ContrRelogio = setTimeout ("Hoje()", 1000) 
Hr = new Date() 
dd = Hr.getDate() 
mm = Hr.getMonth() + 1 
aa = Hr.getYear() 
hh = Hr.getHours() 
min = Hr.getMinutes() 
seg = Hr.getSeconds() 
DataAtual = ((dd < 10) ? "0" + dd + "/" : dd + "/") 
DataAtual += ((mm < 10) ? "0" + mm + "/" + aa : mm + "/" + aa) 
document.DataHora.Data.value=DataAtual 
} 
// 
function CriaArray (n) { 
this.length = n } 
// 
NomeDia = new CriaArray(7) 
NomeDia[0] = "Domingo" 
NomeDia[1] = "Segunda" 
NomeDia[2] = "Terça" 
NomeDia[3] = "Quarta" 
NomeDia[4] = "Quinta" 
NomeDia[5] = "Sexta" 
NomeDia[6] = "Sábado" 
// 
NomeMes = new CriaArray(12) 
NomeMes[0] = "Janeiro" 
NomeMes[1] = "Fevereiro" 
NomeMes[2] = "Março" 
NomeMes[3] = "Abril" 
NomeMes[4] = "Maio" 
NomeMes[5] = "Junho" 
NomeMes[6] = "Julho" 
NomeMes[7] = "Agosto" 
NomeMes[8] = "Setembro" 
NomeMes[9] = "Outubro" 
NomeMes[10] = "Novembro" 
NomeMes[11] = "Dezembro" 
// 
Data1 = new Date() 
dia = Data1.getDate() 
dias = Data1.getDay() 
mes = Data1.getMonth() 
ano = Data1.getYear() 
document.write ("Mauá, " + NomeDia[dias] + " " + dia + " de " + 
NomeMes[mes] + " de " + (ano ) ) 
</script>
		  </div>
		</div>
		<div id="textopaiol">
		  <p>"Caro L&iacute;der, este &eacute; um lugar aonde voc&ecirc; pode adquirir materiais para a edifica&ccedil;&atilde;o da sua vida e para o bom desenvolvimento da c&eacute;lula. </p>
		  <p>&nbsp;</p>
		  <p>&nbsp;</p>
		  <p>&nbsp;</p>
		  <p>&nbsp;</p>
		  <p>&nbsp;</p>
		  <p>&nbsp;</p>
		  <p><a href="sair.asp" class="style3">SAIR</a></p>
		</div>
		<div id="conteudopaiol">
		  <p>Os arquivos para downloads que est&atilde;o relacionados abaixo est&atilde;o na exten&ccedil;&atilde;o <strong>.Doc </strong>(Documento do Word).Clique no link abaixo para fazer o download. </p>
  </div>
  <div id="conteudopaio2">
		  <p><a href="relatorio.doc">*Relat&oacute;rio de C&eacute;lula </a></p>
		  <p><a href="ficha.doc">*Ficha de Encontro </a></p>
  </div>
        <div id="linksrodape"><span class="style2">.</span><span class="style1">::</span><span class="style2">.</span> <a href="quemsomos.html">Quem Somos</a> <span class="style2">.</span><span class="style1">:</span>:<span class="style2">.</span> | <span class="style2">.</span><span class="style1">::</span><span class="style2">.</span> <a href="visao.html">Visão</a> <span class="style2">.</span><span class="style1">::</span><span class="style2">.</span></div>
		
		<div id="imagemnewsletter"></div>
		<div id="barra2">
				<div id="rodape1">.::. &copy;Copyright: Rede Jovens Avivados - 2007</div>
				<div id="rodape2">Agencia de Desenvolvimento WSD</div>
  </div>

</div>
</body>
</html>
