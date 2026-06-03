Imports System.Web

<Serializable()> _
Public Class ListProdutoXEspecificacao
    Inherits List(Of ProdutoXEspecificacao)

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
        sql = "Select pXe.Produto_Id, pXe.CodigoEspecificacao_Id, pXe.FaixaInicial, pXe.FaixaFinal, pXe.Ativo, EPrd.Ativo AS codigoAtivo" & vbCrLf & _
              "FROM ProdutoXEspecificacao pXe" & vbCrLf & _
              "     INNER JOIN EspecificacaoDoProduto EPrd" & vbCrLf & _
              "             on EPrd.Codigo_Id = pXe.CodigoEspecificacao_Id" & vbCrLf & _
              " Where pXe.Produto_id ='" & Me.Produto.Codigo & "'" & vbCrLf & _
              " Order by pXe.CodigoEspecificacao_Id"

        ds = Banco.ConsultaDataSet(sql, "ProdutoXEspecificacao")

        For Each row In ds.Tables(0).Rows
            Dim prdE As New ProdutoXEspecificacao()
            prdE.CodigoProduto = row("Produto_Id")
            prdE.CodigoEspecificacao = row("CodigoEspecificacao_Id")
            prdE.FaixaInicial = row("FaixaInicial")
            prdE.FaixaFinal = row("FaixaFinal")
            prdE.Ativo = row("Ativo")
            prdE.Existe = row("codigoAtivo")
            Me.Add(prdE)
        Next
    End Sub

    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        SalvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each prdE As ProdutoXEspecificacao In Me
            If Produto.IUD = "D" OrElse Produto.IUD = "I" Then prdE.IUD = Produto.IUD
            If prdE.IUD <> "" Then prdE.salvarSql(Sqls)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ProdutoXEspecificacao
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(pProduto As String, pCodigoEspecificacao As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String

        sql = "Select pXe.Produto_Id, pXe.CodigoEspecificacao_Id, pXe.FaixaInicial, pXe.FaixaFinal, pXe.Ativo, EPrd.Ativo AS codigoAtivo" & vbCrLf & _
              "FROM ProdutoXEspecificacao pXe" & vbCrLf & _
              "     INNER JOIN EspecificacaoDoProduto EPrd" & vbCrLf & _
              "             on EPrd.Codigo_Id = pXe.CodigoEspecificacao_Id" & vbCrLf & _
              " Where pXe.Produto_id             = '" & pProduto & "'" & vbCrLf & _
              "   and pXe.CodigoEspecificacao_Id = " & pCodigoEspecificacao
        ds = Banco.ConsultaDataSet(sql, "UnidComercializacao")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        Me.CodigoProduto = pProduto
        Me.CodigoEspecificacao = pCodigoEspecificacao
        Me.FaixaInicial = row("FaixaInicial")
        Me.FaixaInicial = row("FaixaFinal")
        Me.Ativo = row("Ativo")
        Me.Existe = row("codigoAtivo")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoProduto As String
    Private _Produto As Produto

    Private _CodigoEspecificacao As Integer
    Private _EspecificacaoDoProduto As EspecificacaoDoProduto

    Private _FaixaInicial As Decimal
    Private _FaixaFinal As Decimal
    Private _Ativo As Boolean = False

    Private _Existe As Boolean = False

    Private _UsuarioInclusao As String = ""
    Private _DataInclusao As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _DataAlteracao As DateTime
    Private _UsuarioCancelamento As String = ""
    Private _DataCancelamento As DateTime
#End Region

#Region "Property"
    Sub New(objProduto As Produto)
        'TODO: Complete member initialization 
        Produto = objProduto
    End Sub

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

    Public Property CodigoEspecificacao As Integer
        Get
            Return _CodigoEspecificacao
        End Get
        Set(value As Integer)
            _CodigoEspecificacao = value
        End Set
    End Property

    Public Property EspecificacaoDoProduto As EspecificacaoDoProduto
        Get
            If _EspecificacaoDoProduto Is Nothing And _CodigoEspecificacao > 0 Then _EspecificacaoDoProduto = New EspecificacaoDoProduto(_CodigoEspecificacao)
            Return _EspecificacaoDoProduto
        End Get
        Set(value As EspecificacaoDoProduto)
            _EspecificacaoDoProduto = value
        End Set
    End Property

    Public Property FaixaInicial As Decimal
        Get
            Return _FaixaInicial
        End Get
        Set(value As Decimal)
            _FaixaInicial = value
        End Set
    End Property

    Public Property FaixaFinal As Decimal
        Get
            Return _FaixaFinal
        End Get
        Set(value As Decimal)
            _FaixaFinal = value
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

    Public Property Existe As Boolean
        Get
            Return _Existe
        End Get
        Set(value As Boolean)
            _Existe = value
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

    Public Property UsuarioCancelamento() As String
        Get
            Return _UsuarioCancelamento
        End Get
        Set(ByVal value As String)
            _UsuarioCancelamento = value
        End Set
    End Property

    Public Property DataCancelamento() As DateTime
        Get
            Return _DataCancelamento
        End Get
        Set(ByVal value As DateTime)
            _DataCancelamento = value
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
                strSql = "INSERT INTO ProdutoXEspecificacao(Produto_Id, CodigoEspecificacao_Id, FaixaInicial, FaixaFinal, Ativo, UsuarioInclusao, UsuarioInclusaoData) " & vbCrLf & _
                         "Values('" & Me.CodigoProduto & "'," & Me.CodigoEspecificacao & "," & Str(Me.FaixaInicial) & "," & Str(Me.FaixaFinal) & "," & IIf(Me.Ativo, 1, 0) & ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', getdate())"
            Case "U"
                strSql = "Update ProdutoXEspecificacao Set " & vbCrLf & _
                         "   FaixaInicial         = " & Str(Me.FaixaInicial) & vbCrLf & _
                         "  ,FaixaFinal           = " & Str(Me.FaixaFinal) & vbCrLf & _
                         "  ,Ativo                = " & IIf(Me.Ativo, 1, 0) & vbCrLf & _
                         "	,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                         "	,UsuarioAlteracaoData = getdate() " & vbCrLf & _
                         " Where Produto_Id             = '" & Me.CodigoProduto & "'" & vbCrLf & _
                         "   and CodigoEspecificacao_Id = " & Me.CodigoEspecificacao
            Case "D"
                strSql = "Update ProdutoXEspecificacao Set " & vbCrLf & _
                         "   Ativo                   = " & IIf(Me.Ativo, 1, 0) & vbCrLf & _
                         "	,UsuarioCancelamento     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                         "	,UsuarioCancelamentoData = getdate() " & vbCrLf & _
                         " Where Produto_Id             = '" & Me.CodigoProduto & "'" & vbCrLf & _
                         "   and CodigoEspecificacao_Id = " & Me.CodigoEspecificacao
        End Select

        sqls.Add(strSql)
    End Sub
#End Region

End Class
