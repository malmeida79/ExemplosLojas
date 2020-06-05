<%

Set rs9 = Server.CreateObject("ADODB.Recordset")
rs9.Open "SELECT * FROM produtos WHERE destaque='s'" , abredb, adOpenStatic, adLockReadOnly
rmax=rs9.recordcount
	if rs9.eof or rs9.bof then
		mostrar_produto_destaque_fachada="Nao"
	rs9.close
	set rs9 = nothing
	end if

	t=Timer
	Randomize t
	rnum = Int(RND * rmax)
	rs9.move rnum
	set atualizar = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs9("idprod")&" ;")



'Produto 1
set rs = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' AND estoque='s' ;")
if rs.eof or rs.bof then
rs.close
set rs = nothing
set atualizar = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' AND estoque='s' ;")
if rs.eof or rs.bof then
rs.close
set rs = nothing
set atualizar = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' AND estoque='s' ;")
end if
else
set atualizar = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs("idprod")&" AND estoque='s' ;")
end if

'Produto 2
set rs2 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" AND estoque='s' ;")
if rs2.eof or rs2.bof then
rs2.close
set rs2 = nothing
set atualizar2 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs2 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" AND estoque='s' ;")
if rs2.eof or rs2.bof then
rs2.close
set rs2 = nothing
set atualizar2 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs2 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" AND estoque='s' ;")
end if
else
set atualizar2 = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs2("idprod")&" AND estoque='s' ;")
end if

'Produto 3
set rs3 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" AND estoque='s' ;")
if rs3.eof or rs3.bof then
rs3.close
set rs3 = nothing
set atualizar3 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs3 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" AND estoque='s' ;")
if rs3.eof or rs3.bof then
rs3.close
set rs3 = nothing
set atualizar3 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs3 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" AND estoque='s' ;")
end if
else
set atualizar3 = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs3("idprod")&" AND estoque='s' ;")
end if

'Produto 4
set rs4 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" AND estoque='s' ;")
if rs4.eof or rs4.bof then
rs4.close
set rs4 = nothing
set atualizar4 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs4 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" AND estoque='s' ;")
if rs4.eof or rs4.bof then
rs4.close
set rs4 = nothing
set atualizar4 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs4 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> '"&rs("idprod")&"' and idprod <> '"&rs2("idprod")&"' and idprod <> '"&rs3("idprod")&"' AND estoque='s' ;")
end if
else
set atualizar4 = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs4("idprod")&" AND estoque='s' ;")
end if

'Produto 5
set rs5 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" AND estoque='s' ;")
if rs5.eof or rs5.bof then
rs5.close
set rs5 = nothing
set atualizar5 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs5 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" AND estoque='s' ;")
if rs5.eof or rs5.bof then
rs5.close
set rs5 = nothing
set atualizar5 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs5 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" AND estoque='s' ;")
end if
else
set atualizar5 = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs5("idprod")&" AND estoque='s' ;")
end if


'Produto 6
set rs6 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&" AND estoque='s' ;")
if rs6.eof or rs6.bof then
rs6.close
set rs6 = nothing
set atualizar6 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs6 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&" AND estoque='s' ;")
if rs6.eof or rs6.bof then
rs6.close
set rs6 = nothing
set atualizar6 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs6 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&" AND estoque='s' ;")
end if
else
set atualizar6 = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs6("idprod")&" AND estoque='s' ;")
end if


'Produto 7
set rs7 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&" and idprod <> "&rs6("idprod")&" AND estoque='s' ;")
if rs7.eof or rs7.bof then
rs7.close
set rs7 = nothing
set atualizar7 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs7 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&" and idprod <> "&rs6("idprod")&" AND estoque='s' ;")
if rs7.eof or rs7.bof then
rs7.close
set rs7 = nothing
set atualizar7 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs7 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&" and idprod <> "&rs6("idprod")&" AND estoque='s' ;")
end if
else
set atualizar7 = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs7("idprod")&" AND estoque='s' ;")
end if

'Produto 8
set rs8 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&" and idprod <> "&rs6("idprod")&" and idprod <> "&rs7("idprod")&" AND estoque='s' ;")
if rs8.eof or rs8.bof then
rs8.close
set rs8 = nothing
set atualizar8 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs8 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&" and idprod <> "&rs6("idprod")&" and idprod <> "&rs7("idprod")&" AND estoque='s' ;")
if rs8.eof or rs8.bof then
rs8.close
set rs8 = nothing
set atualizar8 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs8 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&" and idprod <> "&rs6("idprod")&" and idprod <> "&rs7("idprod")&" AND estoque='s' ;")
end if
else
set atualizar8 = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs8("idprod")&" AND estoque='s' ;")
end if

'Produto 9
set rs99 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&" and idprod <> "&rs6("idprod")&" and idprod <> "&rs7("idprod")&" and idprod <> "&rs8("idprod")&" AND estoque='s' ;")
if rs99.eof or rs99.bof then
rs99.close
set rs99 = nothing
set atualizar99 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs99 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&" and idprod <> "&rs6("idprod")&" and idprod <> "&rs7("idprod")&" and idprod <> "&rs8("idprod")&" AND estoque='s' ;")
if rs99.eof or rs99.bof then
rs99.close
set rs99 = nothing
set atualizar99 = abredb.Execute("UPDATE produtos SET status = 'nao' WHERE status = 'ok' AND estoque='s' ;")
set rs99 = abredb.execute("SELECT * FROM produtos WHERE status <> 'ok' and idprod <> "&rs("idprod")&" and idprod <> "&rs2("idprod")&" and idprod <> "&rs3("idprod")&" and idprod <> "&rs4("idprod")&" and idprod <> "&rs5("idprod")&" and idprod <> "&rs6("idprod")&" and idprod <> "&rs7("idprod")&" and idprod <> "&rs8("idprod")&" AND estoque='s' ;")
end if
else
set atualizar99 = abredb.Execute("UPDATE produtos SET status = 'ok' WHERE idprod = "&rs99("idprod")&" AND estoque='s' ;")
end if


intProdID1 = rs("idprod")
intProdID2 = rs2("idprod")
intProdID3 = rs3("idprod")
intProdID4 = rs4("idprod")
intProdID5 = rs5("idprod")
intProdID6 = rs6("idprod")
intProdID7 = rs7("idprod")
intProdID8 = rs8("idprod")
intProdID9 = rs9("idprod")
intProdID99 = rs99("idprod")


'Formatação dos preços dos produtos
precito1 = formatNumber(rs("preco"), 2)
precito2 = formatNumber(rs2("preco"), 2)
precito3 = formatNumber(rs3("preco"), 2)
precito4 = formatNumber(rs4("preco"), 2)
precito5 = formatNumber(rs5("preco"), 2)
precito6 = formatNumber(rs6("preco"), 2)
precito7 = formatNumber(rs7("preco"), 2)
precito8 = formatNumber(rs8("preco"), 2)
precito9 = formatNumber(rs9("preco"), 2)
precito99 = formatNumber(rs99("preco"), 2)%>


<!-- #include file="destaque_i.asp" -->
<% End If %>
<div align="center">
  <!-- #include file="produto_i.asp" -->
</div>
