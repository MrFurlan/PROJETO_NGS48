Imports System.Web

<Serializable()> _
Public Class ListProdutoXProcedimento
    Inherits List(Of ProdutoXProcedimento)

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

        sql = " SELECT PrdxP.Produto_Id, PrdxP.CodigoProcedimento_Id, PrdxP.Ativo, Pro.Descricao" & vbCrLf & _
              " FROM ProdutoXProcedimentoDeProducao as PrdxP" & vbCrLf & _
              " INNER JOIN ProcedimentoDeProducao Pro" & vbCrLf & _
              " ON Pro.Codigo_Id = PrdxP.CodigoProcedimento_Id" & vbCrLf & _
              " WHERE PrdxP.Produto_Id = " & Me.Produto.Codigo & "" & vbCrLf & _
              " ORDER BY PrdxP.CodigoProcedimento_Id"

        ds = Banco.ConsultaDataSet(sql, "ProdutoXProcedimentoDeProducao")

        For Each row In ds.Tables(0).Rows
            Dim pPro As New ProdutoXProcedimento()
            pPro.CodigoProduto = row("Produto_Id")
            pPro.CodigoProcedimento = row("CodigoProcedimento_Id")
            pPro.Ativo = row("Ativo")
            pPro.Descricao = row("Descricao")
            Me.Add(pPro)
        Next
    End Sub

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each prdP As ProdutoXProcedimento In Me
            If Produto.IUD = "D" OrElse Produto.IUD = "I" Then prdP.IUD = Produto.IUD
            If prdP.IUD <> "" Then prdP.SalvarSql(Sqls)
        Next
    End Sub

#End Region
End Class

<Serializable()> _
Public Class ProdutoXProcedimento
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(pProduto As String, pCodigoProcedimento As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String

        sql = " SELECT PrdxP.Produto_Id, PrdxP.CodigoProcedimento_Id, PrdxP.Ativo," & vbCrLf & _
              " PrdxP.UsuarioInclusao, PrdxP.UsuarioInclusaoData," & vbCrLf & _
              " PrdxP.UsuarioAlteracao, PrdxP.UsuarioAlteracaoData," & vbCrLf & _
              " Pro.Descricao" & vbCrLf & _
              " FROM ProdutoXProcedimentoDeProducao as PrdxP" & vbCrLf & _
              " INNER JOIN ProcedimentoDeProducao Pro" & vbCrLf & _
              " ON Pro.Codigo_Id = PrdxP.CodigoProcedimento_Id" & vbCrLf & _
              " WHERE PrdxP.Produto_Id          = '" & pProduto & "'" & vbCrLf & _
              " AND PrdxP.CodigoProcedimento_Id =  " & pCodigoProcedimento

        ds = Banco.ConsultaDataSet(sql, "ProdutoXProcedimentoDeProducao")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        Me.CodigoProduto = pProduto
        Me.CodigoProcedimento = pCodigoProcedimento
        Me.Ativo = row("Ativo")
        Me.UsuarioInclusao = row("UsuarioInclusao")
        Me.DataInclusao = row("UsuarioInclusaoData")
        Me.UsuarioAlteracao = row("UsuarioAlteracao")
        Me.DataAlteracao = row("UsuarioAlteracaoData")
        Me.Descricao = row("Descricao")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoProduto As String
    Private _Produto As Produto

    Private _CodigoProcedimento As Integer
    Private _ProcedimentoDoProduto As ProcedimentoParaProducao

    Private _Ativo As Boolean = False

    Private _UsuarioInclusao As String = ""
    Private _DataInclusao As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _DataAlteracao As DateTime

    Private _Descricao As String
#End Region

#Region "Property"
    Sub New(objProduto As Produto)
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

    Public Property CodigoProcedimento As Integer
        Get
            Return _CodigoProcedimento
        End Get
        Set(value As Integer)
            _CodigoProcedimento = value
        End Set
    End Property

    Public Property ProcedimentoDoProduto As ProcedimentoParaProducao
        Get
            If _ProcedimentoDoProduto Is Nothing And _CodigoProcedimento > 0 Then _ProcedimentoDoProduto = New ProcedimentoParaProducao(_CodigoProcedimento)
            Return _ProcedimentoDoProduto
        End Get
        Set(value As ProcedimentoParaProducao)
            _ProcedimentoDoProduto = value
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

    Public Property Descricao As String
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
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = " INSERT INTO ProdutoXProcedimentoDeProducao(Produto_Id, CodigoProcedimento_Id, Ativo," & vbCrLf & _
                      " UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                      " VALUES ('" & CodigoProduto & "', '" & CodigoProcedimento & "', '" & Ativo & "','" & HttpContext.Current.Session("ssNomeUsuario") & "', getdate())"
                Sqls.Add(sql)
            Case "U"
                sql = " UPDATE ProdutoXProcedimentoDeProducao " & vbCrLf & _
                      " SET Ativo                   = " & 1 & "," & vbCrLf & _
                      "	UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                      " UsuarioAlteracaoData        = getdate()" & vbCrLf & _
                      "	WHERE Produto_Id            =  " & CodigoProduto & vbCrLf & _
                      " AND   CodigoProcedimento_Id =  " & CodigoProcedimento
                Sqls.Add(sql)
            Case "D"
                sql = " UPDATE ProdutoXProcedimentoDeProducao " & vbCrLf & _
                      " SET Ativo                 = " & 0 & "," & vbCrLf & _
                      "	UsuarioAlteracao          = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                      " UsuarioAlteracaoData      = getdate()" & vbCrLf & _
                      "	WHERE Produto_Id          = " & CodigoProduto & vbCrLf & _
                      " AND CodigoProcedimento_Id = " & CodigoProcedimento
                Sqls.Add(sql)
        End Select
    End Sub
#End Region
End Class
