<Serializable()> _
Public Class ListProdutosXUnidadeDeComercializacao
    Inherits List(Of ProdutosXUnidadeDeComercializacao)

#Region "Construtor"
    Public Sub New(pCodigoProduto As String)
        Me.Produto = New Produto(pCodigoProduto)
        CarregarLista()
    End Sub

    Public Sub New(pProduto As Produto)
        Me.Produto = pProduto
        CarregarLista()
    End Sub
#End Region

#Region "Fields"
    Private _Produto As Produto
#End Region

#Region "Propertys"
    Public Property Produto As Produto
        Get
            Return _Produto
        End Get
        Set(value As Produto)
            _Produto = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub CarregarLista()
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String
        sql = "SELECT produto_id, Unidade_id, FatorConversao_Id, PesoDaEmbalagem, Ativo" & vbCrLf &
              "  FROM ProdutosxUnidadeDeComercializacao" & vbCrLf &
              " Where Produto_id ='" & Me.Produto.Codigo & "'" & vbCrLf &
              " Order by FatorConversao_Id"

        ds = Banco.ConsultaDataSet(sql, "UnidComercializacao")

        For Each row In ds.Tables(0).Rows
            Dim unid As New ProdutosXUnidadeDeComercializacao()
            unid.CodigoProduto = row("produto_id")
            unid.CodigoUnidade = row("Unidade_id")
            unid.FatorConversao = row("FatorConversao_Id")
            unid.PesoDaEmbalagem = row("PesoDaEmbalagem")
            unid.Ativo = row("Ativo")
            Me.Add(unid)
        Next
    End Sub

    Public Function Existe(pUnidadeComercializacao As String) As Boolean
        Return Me.Exists(Function(s) s.CodigoUnidade = pUnidadeComercializacao)
    End Function

    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        SalvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each unid As ProdutosXUnidadeDeComercializacao In Me
            If Produto.IUD = "D" OrElse Produto.IUD = "I" Then unid.IUD = Produto.IUD
            If unid.IUD <> "" Then unid.salvarSql(Sqls)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ProdutosXUnidadeDeComercializacao
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Sub New(pObjProduto As Produto)
        'TODO: Complete member initialization 
        Produto = pObjProduto
    End Sub

    Public Sub New(pProduto As Integer, pUnidade As String, pFatorConversao_id As Decimal)

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String

        sql = "SELECT FatorConversao_id, ativo, ISNULL(PesoDaEmbalagem, 0.00) AS PesoDaEmbalagem " & vbCrLf &
              "  FROM ProdutosxUnidadeDeComercializacao" & vbCrLf &
              " WHERE Produto_id            = '" & pProduto & "'" & vbCrLf &
              "   AND Unidade_id            = '" & pUnidade & "'" & vbCrLf &
              "   AND FatorConversao_id     = " & Str(pFatorConversao_id) & ";"

        ds = Banco.ConsultaDataSet(sql, "UnidComercializacao")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        Me.CodigoProduto = pProduto
        Me.CodigoUnidade = pUnidade
        Me.FatorConversao = row("FatorConversao_id")
        Me.PesoDaEmbalagem = row("PesoDaEmbalagem")
        Me.Ativo = row("ativo")

    End Sub

#End Region

#Region "Fields"

    Private _IUD As String = ""
    Private _CodigoProduto As String
    Private _Produto As Produto

    Private _CodigoUnidade As String = ""
    Private _FatorConversao As Decimal
    Private _PesoDaEmbalagem As Decimal
    Private _Ativo As Boolean = True

#End Region

#Region "Property"

    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodigoProduto As String
        Get
            Return _CodigoProduto
        End Get
        Set(value As String)
            _CodigoProduto = value
            _Produto = Nothing
        End Set
    End Property

    Public Property Produto As Produto
        Get
            If _Produto Is Nothing And Me.CodigoProduto > 0 Then _Produto = New Produto(CodigoProduto)
            Return _Produto
        End Get
        Set(value As Produto)
            _Produto = value
            _CodigoProduto = _Produto.Codigo
        End Set
    End Property

    Public Property CodigoUnidade As String
        Get
            Return _CodigoUnidade
        End Get
        Set(value As String)
            _CodigoUnidade = value
        End Set
    End Property

    Public Property FatorConversao As Decimal
        Get
            Return _FatorConversao
        End Get
        Set(value As Decimal)
            _FatorConversao = value
        End Set
    End Property

    Public Property Ativo As Boolean
        Get
            Return _Ativo
        End Get
        Set(value As Boolean)
            _Ativo = value
        End Set
    End Property

    Public Property PesoDaEmbalagem As Decimal
        Get
            Return _PesoDaEmbalagem
        End Get
        Set(value As Decimal)
            _PesoDaEmbalagem = value
        End Set
    End Property

#End Region

#Region "Methods"

    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        salvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub salvarSql(ByRef sqls As ArrayList)
        Dim strSql As String = ""

        Select Case IUD
            Case "I"
                strSql = "INSERT INTO ProdutosxUnidadeDeComercializacao(produto_id, Unidade_id, FatorConversao_id, PesoDaEmbalagem, Ativo) " & vbCrLf &
                         "Values('" & Me.CodigoProduto & "','" & Me.CodigoUnidade & "'," & Str(Me.FatorConversao) & "," & Str(Me.PesoDaEmbalagem) & "," & IIf(Me.Ativo, 1, 0) & ")"
            Case "U"
                strSql = "Update ProdutosxUnidadeDeComercializacao Set" & vbCrLf &
                         "   FatorConversao_id = " & Str(Me.FatorConversao) & vbCrLf &
                         "  ,PesoDaEmbalagem = " & Str(Me.PesoDaEmbalagem) & vbCrLf &
                         "  ,Ativo          = " & IIf(Me.Ativo, 1, 0) & vbCrLf &
                         " WHERE Produto_Id = '" & Me.CodigoProduto & "'" & vbCrLf &
                         "   AND Unidade_id = '" & Me.CodigoUnidade & "' " & vbCrLf &
                         "   AND FatorConversao_id = " & Str(Me.FatorConversao) & ";"
            Case "D"
                strSql = " DELETE ProdutosxUnidadeDeComercializacao" & vbCrLf &
                         " WHERE Produto_Id = '" & Me.CodigoProduto & "'" & vbCrLf

                If Not String.IsNullOrWhiteSpace(Me.CodigoUnidade) Then strSql &= " AND Unidade_id = '" & Me.CodigoUnidade & "' AND FatorConversao_id = " & Str(Me.FatorConversao) & " "
        End Select
        sqls.Add(strSql)
    End Sub

    Public Function VerificarRelacionamento() As Boolean

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String

        sql = "SELECT Produto_id " & vbCrLf &
              "  FROM PedidoXItem" & vbCrLf &
              " WHERE Produto_id                = '" & Me.CodigoProduto & "'" & vbCrLf &
              "   AND UnidadeComercializacao    = '" & Me.CodigoUnidade & "'" & vbCrLf &
              "   AND FatorConversao            = " & Str(Me.FatorConversao) & ";"

        ds = Banco.ConsultaDataSet(sql, "UnidComercializacao")

        Return ds.Tables(0).Rows.Count > 0

    End Function

#End Region

End Class
