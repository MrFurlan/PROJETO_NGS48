Imports System.Web

<Serializable()> _
Public Class ListProdutoXCentroDeCusto
    Inherits List(Of ProdutoXCentroDeCusto)

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

        sql = " SELECT PxC.Produto_Id As Produto_Id, " & vbCrLf &
              " PxC.CentroDeCusto_Id As CentroDeCusto_Id, " & vbCrLf &
              " PxC.Conta_Id As Conta_Id, " & vbCrLf &
              " ISNULL(PxC.UsuarioInclusao, '') As UsuarioInclusao, " & vbCrLf &
              " ISNULL(PxC.UsuarioAlteracao, '') As UsuarioAlteracao " & vbCrLf &
              " FROM ProdutoXCentrosDeCustos as PxC" & vbCrLf &
              " INNER JOIN CentrosDeCustos CC" & vbCrLf &
              " ON CC.CentroDeCusto_Id = PxC.CentroDeCusto_Id" & vbCrLf &
              " WHERE Produto_Id = '" & Produto.Codigo & "'" & vbCrLf &
              " ORDER BY PxC.CentroDeCusto_Id"

        ds = Banco.ConsultaDataSet(sql, "ProdutoXCentroDeCusto")

        For Each row In ds.Tables(0).Rows
            Dim pC As New ProdutoXCentroDeCusto()
            pC.CodigoProduto = row("Produto_Id")
            pC.CodigoCusto = row("CentroDeCusto_Id")
            pC.CodigoConta = row("Conta_Id")
            pC.UsuarioInclusao = row("UsuarioInclusao")
            pC.UsuarioAlteracao = row("UsuarioAlteracao")
            Me.Add(pC)
        Next
    End Sub

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each prdCP As ProdutoXCentroDeCusto In Me
            If Produto.IUD = "D" OrElse Produto.IUD = "I" Then prdCP.IUD = Produto.IUD
            If prdCP.IUD <> "" Then prdCP.SalvarSql(Sqls)
        Next
    End Sub

#End Region
End Class

<Serializable()> _
Public Class ProdutoXCentroDeCusto
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(pProduto As String, pCodigoCusto As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String

        sql = " SELECT PxC.Produto_Id, PxC.CentroDeCusto_Id, PxC.Conta_Id," & vbCrLf & _
              " PxC.UsuarioInclusao, PxC.UsuarioInclusaoData," & vbCrLf & _
              " PxC.UsuarioAlteracao, PxC.UsuarioAlteracaoData" & vbCrLf & _
              " FROM ProdutoXCentroDeCusto as PxC" & vbCrLf & _
              " INNER JOIN CentrosDeCustos CC" & vbCrLf & _
              " ON CC.CentroDeCusto_Id = PxC.CentroDeCusto_Id" & vbCrLf & _
              " WHERE PxC.Produto_id = '" & pProduto & "'" & vbCrLf & _
              " AND PxC.CentroDeCusto_Id = '" & pCodigoCusto & "'"
        ds = Banco.ConsultaDataSet(sql, "ProdutoXCentroDeCusto")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        Me.CodigoProduto = pProduto
        Me.CodigoCusto = pCodigoCusto
        Me.CodigoConta = row("Conta_Id")
        Me.UsuarioInclusao = row("UsuarioInclusao")
        Me.DataInclusao = row("DataInclusao")
        Me.UsuarioAlteracao = row("UsuarioAlteracao")
        Me.DataAlteracao = row("DataAlteracao")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoProduto As String
    Private _Produto As Produto

    Private _CodigoCusto As String = ""
    Private _CentroDeCusto As CentroDeCusto

    Private _CodigoConta As String = ""
    Private _PlanoDeConta As PlanoDeConta

    Private _UsuarioInclusao As String = ""
    Private _DataInclusao As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _DataAlteracao As DateTime

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

    Public Property CodigoCusto As Integer
        Get
            Return _CodigoCusto
        End Get
        Set(value As Integer)
            _CodigoCusto = value
        End Set
    End Property

    Public Property CentroDeCusto As CentroDeCusto
        Get
            If _CentroDeCusto Is Nothing And _CodigoCusto > 0 Then _CentroDeCusto = New CentroDeCusto(_CodigoCusto)
            Return _CentroDeCusto
        End Get
        Set(value As CentroDeCusto)
            _CentroDeCusto = value
        End Set
    End Property

    Public Property CodigoConta As String
        Get
            Return _CodigoConta
        End Get
        Set(ByVal value As String)
            _CodigoConta = value
        End Set
    End Property

    Public Property PlanoDeConta As PlanoDeConta
        Get
            If _PlanoDeConta Is Nothing And _CodigoConta > 0 Then _PlanoDeConta = New PlanoDeConta("", 0, _CodigoConta)
            Return _PlanoDeConta
        End Get
        Set(value As PlanoDeConta)
            _PlanoDeConta = value
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
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = " INSERT INTO ProdutoXCentrosDeCustos (Produto_id, CentroDeCusto_Id, Conta_Id, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf &
                      " VALUES ('" & Produto.Codigo & "', '" & CentroDeCusto.Codigo & "', '" & PlanoDeConta.Conta & "','" & HttpContext.Current.Session("ssNomeUsuario") & "', getdate())"
                Sqls.Add(sql)
            'Case "U"
            '    sql = " UPDATE ProdutoXCentrosDeCustos      " & vbCrLf &
            '          " SET Conta_Id           = '" & PlanoDeConta.Conta & "'," & vbCrLf &
            '          "	UsuarioAlteracao       = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
            '          " UsuarioAlteracaoData   = getdate()" & vbCrLf &
            '          "	WHERE Produto_Id       =  " & Produto.Codigo & vbCrLf &
            '          " AND   CentroDeCusto_Id =  " & CentroDeCusto.Codigo
            '    Sqls.Add(sql)
            Case "D"
                sql = " DELETE ProdutoXCentrosDeCustos " & vbCrLf &
                      "	WHERE Produto_Id       =  " & Produto.Codigo & vbCrLf &
                      " AND   CentroDeCusto_Id =  " & CentroDeCusto.Codigo
                Sqls.Add(sql)
        End Select
    End Sub
#End Region
End Class