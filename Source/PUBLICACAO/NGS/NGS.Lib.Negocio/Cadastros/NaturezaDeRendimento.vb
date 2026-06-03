Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()>
Public Class NaturezaDeRendimento
    Implements IBaseEntity

#Region "Construtores"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Selecionar(Codigo)
    End Sub

#End Region

#Region "Variáveis"
    Private _IUD As String = ""
    Private _Codigo As Integer
    Private _Descricao As String
    Private _TipoPessoa As eTipoPessoa
    Private _Selecionado As Boolean = False
    Private _Sequencia As Integer
    Public Erro As Exception

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

    Public Property TipoPessoa As eTipoPessoa
        Get
            Return _TipoPessoa
        End Get
        Set(value As eTipoPessoa)
            _TipoPessoa = value
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

    Public Function temNF() As Boolean
        Dim objBanco As New AcessaBanco()

        Dim Sql As String = "SELECT TOP 1 CodigoNaturezaDeRendimento From NotasFiscais where CodigoNaturezaDeRendimento = " & Me.Codigo

        Dim dsNaturezaDeRendimento As DataSet = objBanco.ConsultaDataSet(Sql, "NaturezaDeRendimento")

        If dsNaturezaDeRendimento.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

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
                Sql = " INSERT INTO NaturezaDeRendimento (Codigo_Id, Descricao, TipoDePessoa) " & vbCrLf &
                      " VALUES (" & _Codigo & ",'" & _Descricao & "'," & _TipoPessoa & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE NaturezaDeRendimento SET" & vbCrLf &
                      "    Descricao     = '" & _Descricao & "'" & vbCrLf &
                      "   ,TipoDePessoa  = " & _TipoPessoa & vbCrLf &
                      "  WHERE Codigo_Id = " & _Codigo
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE NaturezaDeRendimento" & vbCrLf &
                      "  WHERE Codigo_Id = " & _Codigo
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Function Selecionar(ByVal Codigo As Integer) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Codigo_Id, Descricao, TipoDePessoa " &
                                   "  FROM NaturezaDeRendimento " &
                                   " WHERE Codigo_Id = " & Codigo

            Dim dsNaturezaDeRendimento As DataSet = objBanco.ConsultaDataSet(strSQL, "NaturezaDeRendimento")

            With dsNaturezaDeRendimento.Tables(0).Rows
                If .Count > 0 Then
                    Me.Codigo = .Item(0)("Codigo_Id")
                    Me.Descricao = .Item(0)("Descricao")
                    Me.TipoPessoa = .Item(0)("TipoDePessoa")
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


<Serializable()>
Public Class NaturezaDeRendimentos
    Inherits List(Of NaturezaDeRendimento)
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(Optional ByVal pWhere As String = "")
        Dim objBanco As New AcessaBanco
        Dim strSql As String
        strSql = "SELECT Codigo_Id, Descricao, TipoDePessoa " &
                 "  FROM NaturezaDeRendimento "

        If pWhere.Trim.Length = 0 Then
            strSql &= pWhere
        End If

        strSql &= "  ORDER BY Codigo_Id "

        Dim ds As DataSet = objBanco.ConsultaDataSet(strSql, "NaturezaDeRendimento")

        If (ds IsNot Nothing AndAlso ds.Tables.Count > 0) Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim nr As New NaturezaDeRendimento
                nr.Codigo = row("Codigo_Id")
                nr.Descricao = row("Descricao")
                nr.TipoPessoa = row("TipoDePessoa")
                Add(nr)
            Next
        End If
    End Sub
#End Region

End Class
