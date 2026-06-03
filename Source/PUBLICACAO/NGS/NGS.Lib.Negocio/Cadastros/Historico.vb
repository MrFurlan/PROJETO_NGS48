Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListHistorico
    Inherits List(Of Historico)

    Public Sub New(Optional ByVal CarregarHistorico As Boolean = False, Optional ByVal Where As String = "")
        If CarregarHistorico Then
            Dim objBanco As New AcessaBanco
            Dim strSQL As String = "SELECT     Historico_Id, Descricao " & _
                                   "FROM         Historicos "
            If Where.Length > 0 Then
                strSQL &= Where
            End If

            strSQL &= " ORDER BY Historico_Id"

            Dim ds As DataSet = objBanco.ConsultaDataSet(strSQL, "Historicos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim i As New Historico
                i.Codigo = row("Historico_Id").ToString()
                i.Descricao = row("Descricao").ToString()
                Add(i)
            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class Historico
    Implements IBaseEntity

#Region "Variáveis locais"

    Private _Codigo As String
    Private _Descricao As String

#End Region

#Region "Campos públicos"

    Public Erro As Exception

#End Region

#Region "Construtores"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Selecionar(Codigo)
    End Sub

#End Region

#Region "Propriedades"

    Public Property Codigo() As String
        Get
            Return _Codigo
        End Get
        Set(ByVal value As String)
            _Codigo = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

#End Region

#Region "Métodos"

    Public Function Selecionar(ByVal Codigo As Integer) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Historico_Id AS Codigo, Descricao " & _
                                   "FROM Historicos " & _
                                   "WHERE Historico_Id = " & Codigo

            Dim dsHistoricos As DataSet = objBanco.ConsultaDataSet(strSQL, "Historicos")

            With dsHistoricos.Tables(0).Rows
                If .Count > 0 Then
                    Me.Codigo = .Item(0)("Codigo").ToString()
                    Me.Descricao = .Item(0)("Descricao").ToString()

                    Return True
                Else : Return False
                End If
            End With
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

#End Region

End Class