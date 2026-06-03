Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class Finalidades
    Inherits List(Of Finalidade)

    Public Sub New(Optional ByVal Carregar As Boolean = False)
        If Not Carregar Then Exit Sub
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Finalidade_Id, Descricao " & _
                                   "FROM Finalidades "

            Dim dsFinalidades As DataSet = objBanco.ConsultaDataSet(strSQL, "Finalidades")

            For Each drFinalidade As DataRow In dsFinalidades.Tables(0).Rows
                Dim objFinalidade As New Finalidade
                objFinalidade.Codigo = drFinalidade("Finalidade_Id")
                objFinalidade.Descricao = drFinalidade("Descricao")
                Me.Add(objFinalidade)
            Next

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            objBanco = Nothing
        End Try
    End Sub

End Class

<Serializable()> _
Public Class Finalidade
    Implements IBaseEntity

    Public Erro As Exception

#Region "Fields"
    Private _Codigo As Integer
    Private _Descricao As String
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Finalidade_Id, Descricao " & _
                                   "FROM Finalidades " & _
                                   "WHERE Finalidade_Id = " & Codigo.ToString()

            Dim dsFinalidade As DataSet = objBanco.ConsultaDataSet(strSQL, "Finalidades")

            If dsFinalidade.Tables(0).Rows.Count > 0 Then
                Dim drFinalidade As DataRow = dsFinalidade.Tables(0).Rows(0)

                Me.Codigo = Convert.ToInt32(drFinalidade("Finalidade_Id"))
                Me.Descricao = drFinalidade("Descricao").ToString()
            End If
        Catch ex As Exception
            Me.Erro = ex
        Finally
            objBanco = Nothing
        End Try
    End Sub
#End Region

#Region "Property"
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
#End Region

End Class