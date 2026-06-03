Imports System.Web

<Serializable()> _
Public Class ListProdutoXEPI
    Inherits List(Of ProdutoXEPI)

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

        sql = " SELECT PxE.Produto_Id, PxE.CodigoEPI_Id, PxE.Ativo, EPrd.Descricao" & vbCrLf & _
              " FROM ProdutoXEPI as PxE" & vbCrLf & _
              " INNER JOIN EPI EPrd" & vbCrLf & _
              " ON EPrd.Codigo_Id = PxE.CodigoEPI_Id" & vbCrLf & _
              " WHERE PxE.Produto_id = " & Me.Produto.Codigo & "" & vbCrLf & _
              " ORDER BY PxE.CodigoEPI_Id"

        ds = Banco.ConsultaDataSet(sql, "ProdutoXEPI")

        For Each row In ds.Tables(0).Rows
            Dim pE As New ProdutoXEPI()
            pE.CodigoProduto = row("Produto_Id")
            pE.CodigoEPI = row("CodigoEPI_Id")
            pE.Ativo = row("Ativo")
            pE.Descricao = row("Descricao")
            Me.Add(pE)
        Next
    End Sub

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each prdEPI As ProdutoXEPI In Me
            If Produto.IUD = "D" OrElse Produto.IUD = "I" Then prdEPI.IUD = Produto.IUD
            If prdEPI.IUD <> "" Then prdEPI.SalvarSql(Sqls)
        Next
    End Sub

#End Region
End Class

<Serializable()> _
Public Class ProdutoXEPI
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(pProduto As String, pCodigoEPI As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String

        sql = " SELECT PxE.Produto_Id, PxE.CodigoEPI_Id, PxE.Ativo," & vbCrLf & _
              " PxE.UsuarioInclusao, PxE.UsuarioInclusaoData," & vbCrLf & _
              " PxE.UsuarioAlteracao, PxE.UsuarioAlteracaoData," & vbCrLf & _
              " EPrd.Descricao" & vbCrLf & _
              " FROM ProdutoXEPI as PxE" & vbCrLf & _
              " INNER JOIN EPI EPrd" & vbCrLf & _
              " ON EPrd.Codigo_Id = PxE.CodigoEPI_Id" & vbCrLf & _
              " WHERE PxE.Produto_id = '" & pProduto & "'" & vbCrLf & _
              " AND PxE.CodigoEPI_Id =  " & pCodigoEPI
        ds = Banco.ConsultaDataSet(sql, "ProdutoXEPI")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        Me.CodigoProduto = pProduto
        Me.CodigoEPI = pCodigoEPI
        Me.Ativo = row("Ativo")
        Me.UsuarioInclusao = row("UsuarioInclusao")
        Me.DataInclusao = row("DataInclusao")
        Me.UsuarioAlteracao = row("UsuarioAlteracao")
        Me.DataAlteracao = row("DataAlteracao")
        Me.Descricao = row("Descricao")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoProduto As String
    Private _Produto As Produto

    Private _CodigoEPI As Integer
    Private _EPIDoProduto As EPI

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

    Public Property CodigoEPI As Integer
        Get
            Return _CodigoEPI
        End Get
        Set(value As Integer)
            _CodigoEPI = value
        End Set
    End Property

    Public Property EPIDoProduto As EPI
        Get
            If _EPIDoProduto Is Nothing And _CodigoEPI > 0 Then _EPIDoProduto = New EPI(_CodigoEPI)
            Return _EPIDoProduto
        End Get
        Set(value As EPI)
            _EPIDoProduto = value
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
                sql = " INSERT INTO ProdutoXEPI(Produto_id, CodigoEPI_id, Ativo, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                      " VALUES ('" & CodigoProduto & "', '" & CodigoEPI & "', '" & Ativo & "','" & HttpContext.Current.Session("ssNomeUsuario") & "', getdate())"
                Sqls.Add(sql)
            Case "U"
                sql = " UPDATE ProdutoXEPI      " & vbCrLf & _
                      " SET Ativo            = " & 1 & "," & vbCrLf & _
                      "	UsuarioAlteracao     = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                      " UsuarioAlteracaoData = getdate()" & vbCrLf & _
                      "	WHERE Produto_Id     =  " & CodigoProduto & vbCrLf & _
                      " AND   CodigoEPI_id   =  " & CodigoEPI
                Sqls.Add(sql)
            Case "D"
                sql = " UPDATE ProdutoXEPI      " & vbCrLf & _
                      " SET Ativo            =  " & 0 & "," & vbCrLf & _
                      "	UsuarioAlteracao     = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                      " UsuarioAlteracaoData = getdate()" & vbCrLf & _
                      "	WHERE Produto_Id     =  " & CodigoProduto & vbCrLf & _
                      " AND   CodigoEPI_id   =  " & CodigoEPI
                Sqls.Add(sql)
        End Select
    End Sub
#End Region
End Class