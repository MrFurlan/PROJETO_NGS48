Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class Indexador
    Implements IBaseEntity

    Private _Codigo As Integer
    Private _Descricao As String

    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Me.Codigo = Codigo
        Selecionar()
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

    Public Function Selecionar() As Boolean
        Return Selecionar(Me.Codigo)
    End Function

    Public Function Selecionar(ByVal Codigo As Integer) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Indexador_Id, Descricao " & _
                                   "FROM Indexadores " & _
                                   "WHERE Indexador_Id = " & Codigo.ToString()

            Dim dsIndexadores As DataSet = objBanco.ConsultaDataSet(strSQL, "Indexadores")

            If dsIndexadores.Tables(0).Rows.Count > 0 Then
                Dim drIndexador As DataRow = dsIndexadores.Tables(0).Rows(0)

                Me.Codigo = Convert.ToInt32(drIndexador("Indexador_Id"))
                Me.Descricao = drIndexador("Descricao").ToString()
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
Public Class Indexadores
    Inherits List(Of Indexador)

    Public Function Selecionar() As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Indexador_Id, Descricao " & _
                                   "FROM Indexadores"

            Dim dsIndexadores As DataSet = objBanco.ConsultaDataSet(strSQL, "Indexadores")

            For Each drIndexador As DataRow In dsIndexadores.Tables(0).Rows
                Dim objIndexador As New Indexador()

                objIndexador.Codigo = Convert.ToInt32(drIndexador("Indexador_Id"))
                objIndexador.Descricao = drIndexador("Descricao").ToString()

                Me.Add(objIndexador)
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