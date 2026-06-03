Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListConsumoXProducao
    Inherits List(Of ConsumoXProducao)

#Region "Construtor"
    Public Sub New()

    End Sub
#End Region

#Region "Methods"
    Public Sub CarregarTudo()
        Carregar(0)
    End Sub

    Public Sub CarregarProdutosRealcionadosComOProdutoDeConsumo(ByVal pCodigoProduto As String, ByVal pCodigoCusto As Integer)
        Carregar(1, pCodigoProduto, pCodigoCusto)
    End Sub

    Public Sub CarregarListaDeProdutosDeConsumo()
        Carregar(2)
    End Sub

    Private Sub Carregar(ByVal pTipo As Integer, Optional ByVal pCodigoProduto As String = "", Optional ByVal pCodigoCusto As Integer = 0)
        Dim sql As String = String.Empty

        Select Case pTipo
            Case 0
                sql = "SELECT ProdutoOrigem_Id, CodigoCustoOrigem_Id, ProdutoDestino_Id, CodigoCustoDestino_Id " & vbCrLf & _
                      "  FROM ConsumoXProducao " & vbCrLf
            Case 1
                sql = "SELECT ProdutoOrigem_Id, CodigoCustoOrigem_Id, ProdutoDestino_Id, CodigoCustoDestino_Id " & vbCrLf & _
                      "  FROM ConsumoXProducao " & vbCrLf & _
                      " Where ProdutoOrigem_Id     ='" & pCodigoProduto & "'" & vbCrLf & _
                      "   and CodigoCustoOrigem_Id = " & pCodigoCusto
            Case 2
                sql = "SELECT distinct ProdutoOrigem_Id, CodigoCustoOrigem_Id, '' as ProdutoDestino_Id, 0 as CodigoCustoDestino_Id FROM ConsumoXProducao"
        End Select

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "ConsumoXProducao")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CxP As New ConsumoXProducao
            CxP.CodigoProdutoOrigem = row("ProdutoOrigem_Id")
            CxP.CodigoCustoOrigem = row("CodigoCustoOrigem_Id")
            CxP.CodigoProdutoDestino = row("ProdutoDestino_Id")
            CxP.CodigoCustoDestino = row("CodigoCustoDestino_Id")
            Me.Add(CxP)
        Next

    End Sub

#End Region

End Class

<Serializable()> _
Public Class ConsumoXProducao

#Region "Construtor"
    Public Sub New()

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String

    '************* Consumo **************
    Private _CodigoProdutoOrigem As String
    Private _ProdutoOrigem As Produto
    Private _CodigoCustoOrigem As Integer
    Private _CustoOrigem As PlanoDeCusto

    '************* Producao **************
    Private _CodigoProdutoDestino As String
    Private _ProdutoDestino As Produto
    Private _CodigoCustoDestino As Integer
    Private _CustoDestino As PlanoDeCusto
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

    '************* Consumo **************
    Public Property CodigoProdutoOrigem() As String
        Get
            Return _CodigoProdutoOrigem
        End Get
        Set(ByVal value As String)
            _CodigoProdutoOrigem = value
        End Set
    End Property

    Public Property ProdutoOrigem() As Produto
        Get
            If _ProdutoOrigem Is Nothing And _CodigoProdutoOrigem.Length > 0 Then _ProdutoOrigem = New Produto(_CodigoProdutoOrigem)
            Return _ProdutoOrigem
        End Get
        Set(ByVal value As Produto)
            _ProdutoOrigem = value
        End Set
    End Property

    Public Property CodigoCustoOrigem() As Integer
        Get
            Return _CodigoCustoOrigem
        End Get
        Set(ByVal value As Integer)
            _CodigoCustoOrigem = value
            _CustoOrigem = Nothing
        End Set
    End Property

    Public ReadOnly Property CustoOrigem() As PlanoDeCusto
        Get
            If _CustoOrigem Is Nothing And _CodigoCustoOrigem > 0 Then _CustoOrigem = New PlanoDeCusto(_CodigoCustoOrigem)
            Return _CustoOrigem
        End Get
    End Property

    Public ReadOnly Property DescricaoCustoOrigem() As String
        Get
            If CustoOrigem Is Nothing Then
                Return ""
            Else
                Return CustoOrigem.Descricao
            End If
        End Get
    End Property


    '************* Producao **************
    Public Property CodigoProdutoDestino() As String
        Get
            Return _CodigoProdutoDestino
        End Get
        Set(ByVal value As String)
            _CodigoProdutoDestino = value
        End Set
    End Property

    Public Property ProdutoDestino() As Produto
        Get
            If _ProdutoDestino Is Nothing And _CodigoProdutoDestino.Length > 0 Then _ProdutoDestino = New Produto(_CodigoProdutoDestino)
            Return _ProdutoDestino
        End Get
        Set(ByVal value As Produto)
            _ProdutoDestino = value
        End Set
    End Property

    Public Property CodigoCustoDestino() As Integer
        Get
            Return _CodigoCustoDestino
        End Get
        Set(ByVal value As Integer)
            _CodigoCustoDestino = value
        End Set
    End Property

    Public ReadOnly Property CustoDestino() As PlanoDeCusto
        Get
            If _CustoDestino Is Nothing And _CodigoCustoDestino > 0 Then _CustoDestino = New PlanoDeCusto(_CodigoCustoDestino)
            Return _CustoDestino
        End Get
    End Property

    Public ReadOnly Property DescricaoCustoDestino() As String
        Get
            If CustoDestino Is Nothing Then
                Return ""
            Else
                Return CustoDestino.Descricao
            End If

        End Get
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
                Sql = " INSERT INTO ConsumoXProducao(ProdutoOrigem_Id, CodigoCustoOrigem_Id, ProdutoDestino_Id, CodigoCustoDestino_Id) " & vbCrLf & _
                      " VALUES ('" & _CodigoProdutoOrigem & "'," & _CodigoCustoOrigem & ",'" & _CodigoProdutoDestino & "'," & _CodigoCustoDestino & ")"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE ConsumoXProducao" & vbCrLf & _
                      "  WHERE ProdutoOrigem_Id      ='" & _CodigoProdutoOrigem & "'" & vbCrLf & _
                      "    and CodigoCustoOrigem_Id  = " & _CodigoCustoOrigem & vbCrLf & _
                      "    and ProdutoDestino_Id     ='" & _CodigoProdutoDestino & "'" & vbCrLf & _
                      "    and CodigoCustoDestino_Id = " & _CodigoCustoDestino
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class