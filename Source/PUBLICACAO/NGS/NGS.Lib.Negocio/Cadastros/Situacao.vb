Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class Situacao
    Implements IBaseEntity

#Region "Constructors"
    Public Sub New()
    End Sub

    Public Sub New(ByVal pCodigo As Integer)
        Dim objBanco As New AcessaBanco()

        Dim strSQL As String = "SELECT Situacao_Id, Descricao " & _
                                   "FROM Situacoes " & _
                                   "WHERE Situacao_id=" & pCodigo

        Dim dsSituacoes As DataSet = objBanco.ConsultaDataSet(strSQL, "Situacoes")

        For Each drSituacao As DataRow In dsSituacoes.Tables(0).Rows
            Codigo = Convert.ToInt32(drSituacao("Situacao_Id"))
            Descricao = drSituacao("Descricao")
        Next

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer
    Private _Descricao As String
    Public Erro As Exception
#End Region

#Region "Properties"
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
                Sql = " INSERT INTO Situacoes (Situacao_Id, Descricao) " & vbCrLf & _
                      " VALUES ('" & Me.Codigo & "','" & Me.Descricao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Situacoes " & vbCrLf & _
                      "    SET Descricao     = '" & Me.Descricao & "'" & vbCrLf & _
                      "  WHERE Situacao_Id = " & Me.Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Situacoes" & vbCrLf & _
                      "  WHERE Situacao_Id = " & Me.Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Function Selecionar(ByVal Codigo As String) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Situacao_Id, Descricao " & _
                                   "  FROM Situacoes " & _
                                   " WHERE Situacao_Id = '" & Codigo & "'"

            Dim dsSituacao As DataSet = objBanco.ConsultaDataSet(strSQL, "Situacoes")

            With dsSituacao.Tables(0).Rows
                If .Count > 0 Then
                    Me.Codigo = .Item(0)("Situacao_Id").ToString()
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
Public Class ListSituacao
    Inherits List(Of Situacao)

    Public Sub New(ByVal CarregarDados As Boolean)
        If CarregarDados Then
            Dim objBanco As New AcessaBanco()

            Dim strSQL As String = "SELECT Situacao_Id, Descricao " & _
                                       "FROM Situacoes " & _
                                       "ORDER BY Situacao_Id"

            Dim dsSituacoes As DataSet = objBanco.ConsultaDataSet(strSQL, "Situacoes")

            For Each drSituacao As DataRow In dsSituacoes.Tables(0).Rows
                Dim objSituacao As New Situacao()

                objSituacao.Codigo = Convert.ToInt32(drSituacao("Situacao_Id"))
                objSituacao.Descricao = drSituacao("Descricao")

                Me.Add(objSituacao)
            Next
        End If
    End Sub

End Class