<!--#include file="../include/admentor2.asp"-->
<!--#include file="../include/admentorsecurity.asp"-->
<!--#include file="menu.inc"-->
<%
'''''''''''  (C) Stefan Holmberg 1999  
'''''''''''  Free to use if these sourcecode lines is not deleted 
'''''''''''  Contact me at webmaster@sqlexperts.com
'''''''''''  http://www.sqlexperts.com
'''''''''''  AdMentor homepage at http://www.create-a-webshop.com
If Session("admin") <> True Then
Session.Abandon
Response.Redirect "login.asp?reason=noauth"
End If

    Set oConn = AdMentor_DBOpenConnection2()
    
	If Request.QueryString("del") = "yes" Then
		If g_AdMentor_Demo = False Then
			oConn.Execute "delete from banner where bannerid=" & Request.QueryString("id") 
   	 	End If
		Response.Redirect "adminbanners.asp"
	End if

	
    Set oRS = oConn.Execute("select  bannerid, name from banner where bannerid = " & Request.QueryString("id")  )

%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Administrate Banners</title>
<style type="text/css">
<!--
     body {  font-family: Arial, Geneva, Helvetica, Verdana; font-size: smaller; color: #000000}
     td {  font-family: Arial, Geneva, Helvetica, Verdana; font-size: smaller; color: #000000}
     th {  font-family: Arial, Geneva, Helvetica, Verdana; font-size: smaller; color: #000000}
     A:link {text-decoration: none;}
     A:visited {text-decoration: none;}
     A:hover {text-decoration: underline;}
-->
</style>
</head>

<body bgcolor="#DDDDDD">

<table bgColor="#003399" border="0" cellPadding="3" cellSpacing="0" height="100%" width="100%">
  <tbody>
    <tr>
      <td vAlign="middle" width="50%" height="60">
      <img src="../images/administration.gif" border="0" width="200" height="60">
      </td>
      <td vAlign="baseline" width="468" height="60" align="right">
   		<%=GetAdminPagesBannerCode()%>
      </td>
    </tr>
    <tr>
      <td height="100%" vAlign="top" width="100%" colspan="2">
        <table align="center" bgColor="#ffffff" border="0" cellPadding="0" cellSpacing="0" height="100%" width="100%">
          <tbody>
            <tr>
              <td height="100%" vAlign="top" width="85%">
                <table bgColor="#ffffff" border="0" cellPadding="10" cellSpacing="0" height="100%" width="100%">
                  <tbody>
                    <tr>
                      <td align="left" height="100%" vAlign="top" width="65%">
                        <table border="0" width="100%">
                          <tr>
                            <td width="50%" valign="middle" align="left"><font color="#aa3333" face="Arial, Geneva, Helvetica, Verdana" size="4"><b>Aten��o: Excluir banner
                              <%=oRS("name").Value%> 
                              ??</b></font>
                            </td>
                            <td width="50%" valign="middle" align="right">&nbsp</td>
                          </tr>
                        </table>
                        <hr color="#000066" noShade SIZE="1">
                        
                        <table border="0" width="100%">
                          <tr>
                            <td width="160" valign="top"><%AdAdminWriteMenu%></td>
                            <td>
                        
                        <div align="center">
                        <table cellpadding="0" cellspacing="0" bgcolor="#000000" width="50%"><tr><td>
                        <table cellpadding="0" cellspacing="1" width="100%"><tr><td bgcolor="#AA3333">
                        <p align="center">
                        <b><font color="#FFFFFF">ATEN��O!</font></b></p>
                            </td></tr>
                        <tr><td bgcolor="#ffffff">
                        <p align="center"><br>
                        Voc� solicitou a exclus�o do banner <%=oRS("name").Value%>. Voc� tem certeza? Esta a��o <b>n�o</b> poder� ser desfeita e ser�o exclu�das todas as estat�sticas deste banner.<br>
                        <br>
                        <%WriteFakeButton "deleteabanner.asp?id=" & oRS("bannerid").Value & "&del=yes", "Sim, tenho certeza" %>
                        &nbsp;&nbsp;&nbsp;&nbsp;
                        <%WriteFakeButton "showdetails.asp?id=" & oRS("bannerid").Value , "N�o!" 
                        %>
                        </td></tr></table>
                        </td></tr></table>

                        </div>
                        
                            </td>
                          </tr>
                        </table>
                        
                        <table border="0" width="100%">
                          <tr>
                            <td width="50%"><%
                            If g_AdMentor_Demo = True Then
                            	Response.Write "This is a demo. No changes will be saved"
                            End If
                            %>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </td>
            </tr>
          </tbody>
        </table>
      </td>
    </tr>
  </tbody>
</table>

</body>

