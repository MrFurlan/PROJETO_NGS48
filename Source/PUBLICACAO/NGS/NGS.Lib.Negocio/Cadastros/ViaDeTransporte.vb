Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ViaDeTransporte
    Implements IBaseEntity

#Region "Variáveis locais"
    Private _IUD As String = ""
    Private _Codigo As Integer
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

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property
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

#Region "Métodos"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = "INSERT Into ViaDeTransportes(Codigo_id, Descricao) " & vbCrLf & _
                      " Values(" & Me.Codigo & ", " & vbCrLf & _
                      "'" & UCase(Me.Descricao) & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE ViaDeTransportes " & vbCrLf & _
                      "    SET Descricao  = '" & Me.Descricao & "'" & vbCrLf & _
                      "  WHERE Codigo_Id  = " & Me.Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE ViaDeTransportes" & vbCrLf & _
                      "  WHERE Codigo_Id = " & Me.Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
    Public Function Selecionar(ByVal Codigo As Integer) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Codigo_Id AS Codigo, Descricao " & _
                                   "FROM ViaDeTransportes " & _
                                   "WHERE Codigo_Id = " & Codigo

            Dim dsViaDeTransportes As DataSet = objBanco.ConsultaDataSet(strSQL, "ViaDeTransportes")

            With dsViaDeTransportes.Tables(0).Rows
                If .Count > 0 Then
                    Me.Codigo = Convert.ToInt32(.Item(0)("Codigo"))
                    Me.Descricao = .Item(0)("Descricao").ToString()

                    Return True
                Else : Return False
                End If
            End With
        Catch ex As Exception
            Me.Erro = ex
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

#End Region

End Class

<Serializable()> _
Public Class ViaDeTransportes
    Inherits List(Of ViaDeTransporte)

    Public Sub New(Optional ByVal CarregarViaDeTransporte As Boolean = False)
        If CarregarViaDeTransporte Then
            Dim objBanco As New AcessaBanco
            Dim strSQL As String = "SELECT     Codigo_Id, Descricao " & _
                                   "FROM         ViaDeTransportes " & _
                                   "ORDER BY Codigo_Id"
            Dim ds As DataSet = objBanco.ConsultaDataSet(strSQL, "ViaDeTransporte")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim via As New ViaDeTransporte
                via.Codigo = row("Codigo_Id")
                via.Descricao = row("Descricao")
                Add(via)
            Next
        End If
    End Sub

End Class