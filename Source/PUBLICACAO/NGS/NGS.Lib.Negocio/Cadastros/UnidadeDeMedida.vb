Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListUnidadeDeMedida
    Inherits List(Of UnidadeDeMedida)

#Region "Construtor"
    Public Sub New()
        Dim Banco As New AcessaBanco

        Try
            Dim strSQL As String = "Select Unidade_Id, Descricao, isnull(UnidadeIndea,'') AS UnidadeIndea FROM UnidadeDeMedida"

            Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "Consulta")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    Dim objUnidade As New UnidadeDeMedida
                    objUnidade.Unidade = dr("Unidade_Id")
                    objUnidade.Descricao = dr("Descricao")
                    objUnidade.UnidadeIndea = dr("UnidadeIndea")
                    Me.Add(objUnidade)
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

End Class

<Serializable()> _
Public Class UnidadeDeMedida
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Unidade As String)
        Dim Banco As New AcessaBanco

        Try
            Dim strSQL As String = "Select Unidade_Id, Descricao, isnull(UnidadeIndea,'') AS UnidadeIndea FROM UnidadeDeMedida " & vbCrLf & _
                                   "WHERE Unidade_Id = '" & Unidade & "'"

            Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "Consulta")

            If ds.Tables(0).Rows.Count > 0 Then
                Dim dr As DataRow = ds.Tables(0).Rows(0)

                Me.Unidade = dr("Unidade_Id")
                Me.Descricao = dr("Descricao")
                Me.UnidadeIndea = dr("UnidadeIndea")
            End If
        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Fields"
    'Public Erro As Exception
    'Private _IUD As String = ""
    'Private _Unidade As String
    'Private _Descricao As String
    'Private _UnidadeIndea As String
#End Region

#Region "Property"
    Public Property Erro As Exception
    Public Property IUD() As String
    Public Property Unidade() As String
    Public Property Descricao() As String
    Public Property UnidadeIndea() As String
#End Region

#Region "Methods"
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
                Sql = " INSERT INTO UnidadeDeMedida(Unidade_Id, Descricao, UnidadeIndea) " & vbCrLf & _
                      " VALUES ('" & Me.Unidade & "','" & Me.Descricao & "','" & Me.UnidadeIndea & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE UnidadeDeMedida SET" & vbCrLf & _
                      "    Descricao = '" & Me.Descricao & "'," & vbCrLf & _
                      "    UnidadeIndea = '" & Me.UnidadeIndea & "'" & vbCrLf & _
                      "  WHERE Unidade_Id    ='" & Me.Unidade & "'"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE UnidadeDeMedida" & vbCrLf & _
                      "  WHERE Unidade_Id    ='" & Me.Unidade & "'"
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class