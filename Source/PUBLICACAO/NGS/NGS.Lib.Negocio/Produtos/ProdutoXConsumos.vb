Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis
Imports System.Web

<Serializable()> _
Public Class ListProdutoXConsumos
    Inherits List(Of ProdutoXConsumos)

#Region "Construtor"
    Public Sub New()

    End Sub
#End Region

#Region "Methods"
    Public Sub CarregarTudo()
        Carregar(0)
    End Sub

    Public Sub CarregarProduto(ByVal pCodigoProduto As String)
        Carregar(1, pCodigoProduto)
    End Sub

    Public Sub CarregarConsumo(ByVal pCodigoProduto As String, ByVal pCodigoProdutoConsumo As String)
        Carregar(2, pCodigoProduto, pCodigoProdutoConsumo)
    End Sub

    Private Sub Carregar(ByVal pTipo As Integer, Optional ByVal pCodigoProduto As String = "", Optional ByVal pCodigoProdutoConsumo As String = "")
        Dim sql As String = String.Empty

        Select Case pTipo
            Case 0
                sql = "SELECT Produto_Id, ProdutoConsumo_Id, Percentual " & vbCrLf & _
                      "  FROM ProdutoXConsumos " & vbCrLf
            Case 1
                sql = "SELECT Produto_Id, ProdutoConsumo_Id, Percentual " & vbCrLf & _
                      "  FROM ProdutoXConsumos " & vbCrLf & _
                      " Where Produto_Id = '" & pCodigoProduto & "'"
            Case 2
                sql = "SELECT Produto_Id, ProdutoConsumo_Id, Percentual " & vbCrLf & _
                      "  FROM ProdutoXConsumos " & vbCrLf & _
                      " Where Produto_Id        = '" & pCodigoProduto & "'" & vbCrLf & _
                      "   and ProdutoConsumo_Id = '" & pCodigoProdutoConsumo & "'"

        End Select

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "ProdutoXConsumos")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim pXc As New ProdutoXConsumos
            pXc.CodigoProduto = row("Produto_Id")
            pXc.CodigoProdutoConsumo = row("ProdutoConsumo_Id")
            pXc.Percentual = row("Percentual")
            Me.Add(pXc)
        Next

    End Sub

    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        SalvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each pXc As ProdutoXConsumos In Me
            If pXc.IUD = "D" OrElse pXc.IUD = "I" Then pXc.IUD = pXc.IUD
            If pXc.IUD <> "" Then pXc.SalvarSql(Sqls)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ProdutoXConsumos

#Region "Construtor"
    Public Sub New()

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String

    '************* Produção **************
    Private _CodigoProduto As String
    Private _Produto As Produto

    '************* Producao **************
    Private _CodigoProdutoConsumo As String
    Private _ProdutoConsumo As Produto

    Private _Percentual As Decimal

    Private _UsuarioInclusao As String = ""
    Private _DataInclusao As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _DataAlteracao As DateTime
    Private _TemConsumo As Boolean = False

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


    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing And _CodigoProduto.Length > 0 Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property CodigoProdutoConsumo() As String
        Get
            Return _CodigoProdutoConsumo
        End Get
        Set(ByVal value As String)
            _CodigoProdutoConsumo = value
        End Set
    End Property

    Public Property ProdutoConsumo() As Produto
        Get
            If _ProdutoConsumo Is Nothing And _CodigoProdutoConsumo.Length > 0 Then _ProdutoConsumo = New Produto(_CodigoProdutoConsumo)
            Return _ProdutoConsumo
        End Get
        Set(ByVal value As Produto)
            _ProdutoConsumo = value
        End Set
    End Property

    Public Property Percentual As Decimal
        Get
            Return _Percentual
        End Get
        Set(value As Decimal)
            _Percentual = value
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

    Public Property TemConsumo As Boolean
        Get
            If _ProdutoConsumo Is Nothing And _CodigoProdutoConsumo.Length > 0 Then _TemConsumo = verProdutoConsumo(_CodigoProdutoConsumo)
            Return _TemConsumo
        End Get
        Set(ByVal value As Boolean)
            _TemConsumo = value
        End Set
    End Property

#End Region

#Region "Methods"

    Private Function verProdutoConsumo(ByVal codigoProdutoConsumo As String) As Boolean
        Dim sql As String = String.Empty

        sql = "SELECT Produto_Id" & vbCrLf & _
                "  FROM OrdemDeProducaoXConsumo" & vbCrLf & _
                " WHERE Empresa_Id = '" & HttpContext.Current.Session("ssEmpresa") & "'" & vbCrLf & _
                "   AND Produto_id = '" & codigoProdutoConsumo & "'" & vbCrLf

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "OrdemDeProducaoXConsumo")

        If ds.Tables(0).Rows.Count > 0 Then
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
                Sql = " INSERT INTO ProdutoXConsumos(Produto_Id, ProdutoConsumo_Id, Percentual, UsuarioInclusao, UsuarioInclusaoData) " & vbCrLf & _
                      " VALUES ('" & _CodigoProduto & "'," & _CodigoProdutoConsumo & "," & Str(_Percentual) & ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', getdate())"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE ProdutoXConsumos SET " & vbCrLf & _
                      "         Percentual           = " & Str(_Percentual) & vbCrLf & _
                         "	   ,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                         "	   ,UsuarioAlteracaoData = getdate() " & vbCrLf & _
                      "  WHERE Produto_Id        ='" & _CodigoProduto & "'" & vbCrLf & _
                      "    and ProdutoConsumo_Id ='" & _CodigoProdutoConsumo & "'"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE ProdutoXConsumos" & vbCrLf & _
                      "  WHERE Produto_Id        ='" & _CodigoProduto & "'" & vbCrLf & _
                      "    and ProdutoConsumo_Id ='" & _CodigoProdutoConsumo & "'"
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class