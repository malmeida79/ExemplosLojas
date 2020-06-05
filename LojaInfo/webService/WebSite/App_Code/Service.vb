Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols

<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
     Inherits System.Web.Services.WebService

    <WebMethod()> _
    Public Function HelloWorld(ByVal nome As String) As String
        Return "Hello World," & nome
    End Function

    <WebMethod()> _
    Public Function Soma(ByVal v1 As Short, ByVal v2 As Short)
        soma = v1 + v2
    End Function

End Class
