<%
'#########################################################################################
'#----------------------------------------------------------------------------------------
'#########################################################################################

'INÍCIO DO CÓDIGO

'Este código está otimizado e roda tanto em Windows 2000/NT/XP/ME/98 quanto em servidores UNIX-LINUX com chilli!ASP
%>
<!-- #include file="topo.asp" -->

<script>
function cadmail()
{
window.open('cadmail.asp?email='+document.emailx.email.value+'&Opcao='+document.emailx.tipo.value,'email','resizable=no,width=270,height=180,scrollbars=no');
		document.emailx.reset()
	}
	function limpa() {
	document.emailx.email.value=''
	}
	
</script>
<style>
SELECT
{
    BACKGROUND-COLOR: #ffffff;
    BORDER-BOTTOM-COLOR: #000000;
    BORDER-BOTTOM-STYLE: #666666;
    BORDER-LEFT-COLOR: #000000;
    BORDER-LEFT-STYLE: #666666;
    BORDER-RIGHT-COLOR: #000000;
    BORDER-RIGHT-STYLE: #666666;
    BORDER-TOP-COLOR: #000000;
    BORDER-TOP-STYLE: #666666;
    COLOR: #000000;
    FONT-FAMILY: Arial, Helvetica, sans-serif;
    FONT-SIZE: xx-small;
}
</style>
		
			<!-- #include file="produto.asp" -->
				<!-- #include file="baixo.asp" -->
 
	<%rs.close%>
	<%set rs = nothing%>
				
	