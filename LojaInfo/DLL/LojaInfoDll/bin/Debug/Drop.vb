Imports System
Imports System.Collections.Generic
Imports System.Text

Public Class Drop

    'constructor
    Public Sub New(ByVal nm As String, ByVal vl As Integer)
        nome = nm
        valor = vl
    End Sub

    Private _Nome As String
    Private _valor As Integer

    Public Property Nome() As String
        Get
            Return _Nome
        End Get
        Set(ByVal value As String)
            _Nome = value
        End Set
    End Property

    Public Property Valor() As Integer
        Get
            Return _valor
        End Get
        Set(ByVal value As Integer)
            _valor = value
        End Set
    End Property
        
End Class

