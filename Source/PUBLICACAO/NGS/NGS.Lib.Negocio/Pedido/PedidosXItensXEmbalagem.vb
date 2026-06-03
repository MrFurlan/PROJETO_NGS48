Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListPedidosXItensXEmbalagem
    Inherits List(Of PedidosXItensXEmbalagem)

#Region "Field"
    Private _ItemPedido As PedidoXItem
#End Region

#Region "Property"
    Public ReadOnly Property ItemPedido() As PedidoXItem
        Get
            Return _ItemPedido
        End Get
    End Property
#End Region

#Region "Contrutor"
    Public Sub New(ByVal pItemPedido As PedidoXItem)
        Dim sql As String = ""

        sql = " SELECT Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id, QtdeEmbalagem" & vbCrLf & _
              "   FROM PedidosXItensXEmbalagem " & vbCrLf & _
              " Where Empresa_Id             ='" & pItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
              "   and EndEmpresa_Id          = " & pItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
              "   and Pedido_Id              = " & pItemPedido.Pedido.Codigo & vbCrLf & _
              "   and Produto_Id             ='" & pItemPedido.CodigoProduto & "'" & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "PedItemEmb")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim PIE As New PedidosXItensXEmbalagem(pItemPedido)
            PIE.IUD = "U"
            PIE.CodigoEmbalagem = row("Embalagem_Id")
            PIE.CodigoTipoEmbalagem = row("TipoDeEmbalagem_Id")
            PIE.CapacidadeEmbalagem = row("CapacidadeEmbalagem_Id")
            PIE.QtdeEmbalagem = row("QtdeEmbalagem")
            Me.Add(PIE)
        Next
    End Sub
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each PIE As PedidosXItensXEmbalagem In Me
            If Me.ItemPedido.IUD = "D" Or Me.ItemPedido.IUD = "I" Then PIE.IUD = Me.ItemPedido.IUD
            If Me.ItemPedido.IUD <> "" Then
                PIE.SalvarSql(Sqls)
            End If
        Next
    End Sub

    Public Sub CarregarListaComPossibilidades()
        Dim sql As String = ""

        sql = "Select PE.Embalagem_Id, PE.TipoDeEmbalagem_Id, PE.CapacidadeEmbalagem_Id, isnull(PIE.QtdeEmbalagem,0) as QtdeEmbalagem " & vbCrLf & _
              "  from ProdutoxEmbalagem PE" & vbCrLf & _
              "  left join PedidosXItensXEmbalagem PIE" & vbCrLf & _
              "    on PE.Produto_Id             = PIE.Produto_Id" & vbCrLf & _
              "   and PE.Embalagem_Id           = PIE.Embalagem_Id" & vbCrLf & _
              "   and PE.TipoDeEmbalagem_Id     = PIE.TipoDeEmbalagem_Id" & vbCrLf & _
              "   and PE.CapacidadeEmbalagem_Id = PIE.CapacidadeEmbalagem_Id" & vbCrLf & _
              " Where PE.Produto_Id ='" & Me.ItemPedido.CodigoProduto & "'" & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "PedItemEmb")

        Me.Clear()
        For Each row As DataRow In ds.Tables(0).Rows
            Dim PIE As New PedidosXItensXEmbalagem(Me.ItemPedido)
            PIE.IUD = IIf(row("QtdeEmbalagem") > 0, "U", "")
            PIE.CodigoEmbalagem = row("Embalagem_Id")
            PIE.CodigoTipoEmbalagem = row("TipoDeEmbalagem_Id")
            PIE.CapacidadeEmbalagem = row("CapacidadeEmbalagem_Id")
            PIE.QtdeEmbalagem = row("QtdeEmbalagem")
            Me.Add(PIE)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class PedidosXItensXEmbalagem

#Region "Construtor"
    Public Sub New(ByVal pItemPedido As PedidoXItem)
        _ItemPedido = pItemPedido
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _ItemPedido As PedidoXItem
    Private _CodigoEmbalagem As Integer
    Private _CodigoTipoEmbalagem As String = ""
    Private _CapacidadeEmbalagem As Decimal
    Private _QtdeEmbalagem As Integer
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

    Public ReadOnly Property ItemPedido() As PedidoXItem
        Get
            Return _ItemPedido
        End Get
    End Property

    Public Property CodigoEmbalagem() As Integer
        Get
            Return _CodigoEmbalagem
        End Get
        Set(ByVal value As Integer)
            _CodigoEmbalagem = value
        End Set
    End Property

    Public Property CodigoTipoEmbalagem() As String
        Get
            Return _CodigoTipoEmbalagem
        End Get
        Set(ByVal value As String)
            _CodigoTipoEmbalagem = value
        End Set
    End Property

    Public Property CapacidadeEmbalagem() As Decimal
        Get
            Return _CapacidadeEmbalagem
        End Get
        Set(ByVal value As Decimal)
            _CapacidadeEmbalagem = value
        End Set
    End Property

    Public Property QtdeEmbalagem() As Integer
        Get
            Return _QtdeEmbalagem
        End Get
        Set(ByVal value As Integer)
            _QtdeEmbalagem = value
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

        If Me.IUD <> "" And Me.QtdeEmbalagem = 0 Then Me.IUD = "D"

        Select Case Me.IUD
            Case "I"
                Sql = "Insert Into PedidosXItensXEmbalagem(Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id, Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id, QtdeEmbalagem)" & vbCrLf & _
                      " values('" & Me.ItemPedido.Pedido.CodigoEmpresa & "'," & Me.ItemPedido.Pedido.EnderecoEmpresa & "," & Me.ItemPedido.Pedido.Codigo & ",'" & Me.ItemPedido.CodigoProduto & "'," & Me.CodigoEmbalagem & ",'" & Me.CodigoTipoEmbalagem & "'," & Str(Me.CapacidadeEmbalagem) & "," & Me.QtdeEmbalagem & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update PedidosXItensXEmbalagem set" & vbCrLf & _
                      "  QtdeEmbalagem = " & Me.QtdeEmbalagem & vbCrLf & _
                      " Where Empresa_Id             ='" & Me.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id          = " & Me.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                      "   and Pedido_Id              = " & Me.ItemPedido.Pedido.Codigo & vbCrLf & _
                      "   and Produto_Id             ='" & Me.ItemPedido.CodigoProduto & "'" & vbCrLf & _
                      "   and Embalagem_Id           = " & Me.CodigoEmbalagem & vbCrLf & _
                      "   and TipoDeEmbalagem_Id     ='" & Me.CodigoTipoEmbalagem & "'" & vbCrLf & _
                      "   and CapacidadeEmbalagem_Id = " & Me.QtdeEmbalagem & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = "Delete PedidosXItensXEmbalagem" & vbCrLf & _
                      " Where Empresa_Id             ='" & Me.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id          = " & Me.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                      "   and Pedido_Id              = " & Me.ItemPedido.Pedido.Codigo & vbCrLf & _
                      "   and Produto_Id             ='" & Me.ItemPedido.CodigoProduto & "'" & vbCrLf & _
                      "   and Embalagem_Id           = " & Me.CodigoEmbalagem & vbCrLf & _
                      "   and TipoDeEmbalagem_Id     ='" & Me.CodigoTipoEmbalagem & "'" & vbCrLf & _
                      "   and CapacidadeEmbalagem_Id = " & Me.QtdeEmbalagem & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
