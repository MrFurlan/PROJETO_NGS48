Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis
Imports System.Web

<Serializable()> _
Public Class ListProdutoXSequenciaDeLote
    Inherits List(Of ProdutoXSequenciaDeLote)

    Public Sub New()

    End Sub

    Public Sub New(ByVal pProduto As Produto)
        Dim sql As String
        sql = "Select ps.Produto_id, ps.Ano_Id, ps.SequenciaDoProduto, ps.Sequencia " & vbCrLf & _
              "  from ProdutoXSequenciaDeLote ps" & vbCrLf & _
              "      inner Join Produtos prd " & vbCrLf & _
              "              on prd.Produto_Id = ps.Produto_Id " & vbCrLf

        If pProduto.Codigo.Length > 0 Then
            sql &= " Where Produto_id ='" & pProduto.Codigo & "'" & vbCrLf
        End If

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "ProdutoXSequenciaDeLote")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ps As New ProdutoXSequenciaDeLote()

            If pProduto.Codigo.Length > 0 Then
                ps.Produto = pProduto
            Else
                Dim objPrd = New Produto(row("Produto_id"))
                ps.Produto = objPrd
            End If

            ps.Ano = row("Ano_Id")

            ps.SequenciaDoProduto = row("SequenciaDoProduto")
            ps.SequenciaDoLote = row("Sequencia")
            Me.Add(ps)
        Next
    End Sub

End Class

<Serializable()> _
Public Class ProdutoXSequenciaDeLote

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pProduto As Produto)
        Me.Produto = pProduto
    End Sub

    Public Sub New(ByVal pProduto As String, ByVal ano As Integer)
        Dim sql As String
        sql = "Select Produto_id, Ano_Id, SequenciaDoProduto, Sequencia " & vbCrLf & _
              "  from ProdutoXSequenciaDeLote" & vbCrLf & _
              " Where Produto_id ='" & pProduto & "'" & vbCrLf & _
              "   and Ano_Id = " & ano

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "ProdutoXSequenciaDeLote")

        If ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow
            row = ds.Tables(0).Rows(0)

            Me.Produto = New Produto(row("Produto_Id"))
            Me.Ano = row("Ano_Id")
            Me.SequenciaDoProduto = row("SequenciaDoProduto")
            Me.SequenciaDoLote = row("Sequencia")
        End If

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Produto As Produto
    Private _SequenciaDoProduto As String
    Private _Ano As Integer
    Private _SequenciaDoLote As Integer

    Private _UsuarioInclusao As String = ""
    Private _DataInclusao As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _DataAlteracao As DateTime

#End Region

#Region "Property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property SequenciaDoProduto() As String
        Get
            Return _SequenciaDoProduto
        End Get
        Set(ByVal value As String)
            _SequenciaDoProduto = value
        End Set
    End Property

    Public Property Ano() As Integer
        Get
            Return _Ano
        End Get
        Set(ByVal value As Integer)
            _Ano = value
        End Set
    End Property

    Public Property SequenciaDoLote() As Integer
        Get
            Return _SequenciaDoLote
        End Get
        Set(ByVal value As Integer)
            _SequenciaDoLote = value
        End Set
    End Property

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property DataInclusao() As DateTime
        Get
            Return _DataInclusao
        End Get
        Set(ByVal value As DateTime)
            _DataInclusao = value
        End Set
    End Property

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property DataAlteracao() As DateTime
        Get
            Return _DataAlteracao
        End Get
        Set(ByVal value As DateTime)
            _DataAlteracao = value
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
                Sql = " INSERT INTO ProdutoXSequenciaDeLote(Produto_Id, Ano_Id, SequenciaDoProduto, Sequencia, UsuarioInclusao, UsuarioInclusaoData) " & vbCrLf & _
                      " VALUES ('" & Produto.Codigo & "', " & Me.Ano & ",'" & Me.SequenciaDoProduto & "'," & Me.SequenciaDoLote & ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', getdate())"
                Sqls.Add(Sql)
            Case "U"
                Sql = " Update ProdutoXSequenciaDeLote set" & vbCrLf & _
                      "   SequenciaDoProduto     = '" & Me.SequenciaDoProduto & "'" & vbCrLf & _
                      "   ,Sequencia             = " & Me.SequenciaDoLote & vbCrLf & _
                      "	   ,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                      "	   ,UsuarioAlteracaoData = getdate() " & vbCrLf & _
                      "  WHERE Produto_Id        ='" & Produto.Codigo & "'" & vbCrLf & _
                      "    AND Ano_Id            = " & Me.Ano
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE ProdutoXSequenciaDeLote" & vbCrLf & _
                      "  WHERE Produto_Id ='" & Produto.Codigo & "'" & vbCrLf & _
                      "    AND Ano_Id            = " & Me.Ano
                Sqls.Add(Sql)
        End Select
    End Sub

#End Region

End Class