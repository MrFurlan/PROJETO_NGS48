Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class Estado
    Implements IBaseEntity

#Region "Variáveis"
    Private _IUD As String = ""
    Private _Codigo As String
    Private _Descricao As String
    Private _Regiao As String
    Private _Selecionado As Boolean = False
    Private _Sequencia As Integer
    Public Erro As Exception

#End Region

#Region "Construtores"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As String)
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

    Public Property Regiao() As String
        Get
            Return _Regiao
        End Get
        Set(ByVal value As String)
            _Regiao = value
        End Set
    End Property

    Public Property Selecionado() As Boolean
        Get
            Return _Selecionado
        End Get
        Set(ByVal value As Boolean)
            _Selecionado = value
        End Set
    End Property

    Public Property Sequencia() As Integer
        Get
            Return _Sequencia
        End Get
        Set(ByVal value As Integer)
            _Sequencia = value
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
                Sql = " INSERT INTO Estados (Estado_Id, Descricao, Regiao) " & vbCrLf & _
                      " VALUES ('" & _Codigo & "','" & _Descricao & "','" & _Regiao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Estados SET" & vbCrLf & _
                      "    Descricao     = '" & _Descricao & "'," & vbCrLf & _
                      "    Regiao     = '" & _Regiao & "'" & vbCrLf & _
                      "  WHERE Estado_Id = '" & _Codigo & "'"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Estados" & vbCrLf & _
                      "  WHERE Estado_Id = '" & _Codigo & "'"
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Function Selecionar(ByVal Codigo As String) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Estado_Id, Descricao, Regiao " & _
                                   "  FROM Estados " & _
                                   " WHERE Estado_Id = '" & Codigo & "'"

            Dim dsEstado As DataSet = objBanco.ConsultaDataSet(strSQL, "Estados")

            With dsEstado.Tables(0).Rows
                If .Count > 0 Then
                    Me.Codigo = .Item(0)("Estado_Id").ToString()
                    Me.Descricao = .Item(0)("Descricao").ToString()
                    Me.Regiao = .Item(0)("Regiao").ToString()

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

    Function VerificaClientes() As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Estado " & _
                                   "  FROM Clientes " & _
                                   " WHERE Estado = '" & Codigo & "'"

            Dim dsEstado As DataSet = objBanco.ConsultaDataSet(strSQL, "Estados")

            If dsEstado.Tables(0).Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Me.Erro = ex
            Return True
        Finally
            objBanco = Nothing
        End Try
    End Function

#End Region

End Class

<Serializable()> _
Public Class Estados
    Inherits List(Of Estado)
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(Optional ByVal CarregarEstados As Boolean = False, Optional ByVal pWhere As String = "")
        If CarregarEstados Then
            Dim objBanco As New AcessaBanco
            Dim strSql As String
            strSql = "SELECT Estado_id, Descricao, Regiao " & _
                     "  FROM Estados "

            If pWhere.Trim.Length = 0 Then
                strSql &= pWhere
            End If

            strSql &= "  ORDER BY Estado_Id "

            Dim ds As DataSet = objBanco.ConsultaDataSet(strSql, "Estados")

            If (ds IsNot Nothing AndAlso ds.Tables.Count > 0) Then
                For Each row As DataRow In ds.Tables(0).Rows
                    Dim est As New Estado
                    est.Codigo = row("Estado_Id")
                    est.Descricao = row("Descricao")
                    est.Regiao = row("Regiao")
                    Add(est)
                Next
            End If
        End If
    End Sub
#End Region

End Class