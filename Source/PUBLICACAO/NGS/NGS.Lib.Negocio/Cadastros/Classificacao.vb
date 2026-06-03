Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class Classificacao
    Implements IBaseEntity

    Private _Codigo As Integer
    Private _Descricao As String

    Public Erro As Exception

    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Selecionar(Codigo)
    End Sub

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
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

    Public Function Selecionar(ByVal Codigo As Integer) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Codigo_Id, Descricao " & _
                                   "FROM TabelaDeClassificacoes " & _
                                   "WHERE Codigo_Id = " & Codigo.ToString()

            Dim dsClassificacoes As DataSet = objBanco.ConsultaDataSet(strSQL, "TabelaDeClassificacoes")

            If dsClassificacoes.Tables(0).Rows.Count > 0 Then
                Dim drClassificacao = dsClassificacoes.Tables(0).Rows(0)
                Me.Codigo = Convert.ToInt32(drClassificacao("Codigo_Id"))
                Me.Descricao = drClassificacao("Descricao").ToString()
            End If

            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

End Class

<Serializable()> _
Public Class Classificacoes
    Inherits List(Of Classificacao)

    Public Sub New()
    End Sub

    Public Function Selecionar() As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Codigo_Id, Descricao " & _
                                   "FROM TabelaDeClassificacoes"

            Dim dsClassificacoes As DataSet = objBanco.ConsultaDataSet(strSQL, "TabelaDeClassificacoes")

            For Each drClassificacao As DataRow In dsClassificacoes.Tables(0).Rows
                Dim objClassificacao As New Classificacao()

                objClassificacao.Codigo = Convert.ToInt32(drClassificacao("Codigo_Id"))
                objClassificacao.Descricao = drClassificacao("Descricao").ToString()

                Me.Add(objClassificacao)
            Next

            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

End Class