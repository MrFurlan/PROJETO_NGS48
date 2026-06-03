Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

'****************************************************************************************************************************************************************
'*****************************************   Classe Lista PedidoXRepresentantexTabelaDeComissao   ***************************************************************
'****************************************************************************************************************************************************************
Public Class ListPedidoXRepresentantesxTabelaDeComissao
    Inherits List(Of PedidoXRepresentantesxTabelaDeComissao)

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByRef pRepresentantePedido As PedidoXRepresentante, Optional ByVal SomenteSalvos As Boolean = True)
        _RepresentantePedido = pRepresentantePedido
        Dim sql As String
        Dim Banco As New AcessaBanco
        sql = "Select PxI.Empresa_Id, PxI.EndEmpresa_Id, PxI.Pedido_id, " & vbCrLf & _
              "       PxI.Produto_Id, Prd.Descricao as NomeProduto," & vbCrLf & _
              "       C.Representante_Id, C.EndRepresentante_Id, isnull(Rep.nome,'') as NomeRepresentante," & vbCrLf & _
              "       isnull(PxT.Tabela_Id,0) as Tabela_Id, isnull(T.Descricao,'') as NomeTabela " & vbCrLf & _
              "  from PedidoxitemXLancamento PxI" & vbCrLf & _
              " Inner Join Produtos Prd" & vbCrLf & _
              "    on Prd.Produto_Id    = PxI.Produto_Id" & vbCrLf & _
              " Inner Join Comissoes C" & vbCrLf & _
              "    on PxI.Empresa_Id    = C.Empresa_Id" & vbCrLf & _
              "   and PxI.EndEmpresa_Id = C.EndEmpresa_Id" & vbCrLf & _
              "   and PxI.Pedido_Id     = C.Pedido_Id" & vbCrLf & _
              " Inner Join Clientes Rep " & vbCrLf & _
              "    on Rep.Cliente_Id  = C.Representante_Id" & vbCrLf & _
              "   and Rep.Endereco_Id = c.EndRepresentante_Id" & vbCrLf & _
              "  Left Join PedidoXTabelaDeComissao PxT" & vbCrLf & _
              "    on PxT.Empresa_Id          = PxI.Empresa_Id" & vbCrLf & _
              "   and PxT.EndEmpresa_Id       = PxI.EndEmpresa_Id" & vbCrLf & _
              "   and PxT.Pedido_Id           = PxI.Pedido_id" & vbCrLf & _
              "   and PxT.Produto_Id          = PxI.Produto_Id" & vbCrLf & _
              "   and PxT.Representante_Id    = C.Representante_Id" & vbCrLf & _
              "   and PxT.EndRepresentante_Id = C.EndRepresentante_Id" & vbCrLf & _
              "  Left Join TabelaDeComissao T" & vbCrLf & _
              "    on T.Codigo_Id = PxT.Tabela_Id" & vbCrLf & _
              " where PxI.Empresa_id          ='" & _RepresentantePedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
              "   and PxI.EndEmpresa_Id       = " & _RepresentantePedido.Pedido.EnderecoEmpresa & vbCrLf & _
              "   and PxI.pedido_id           = " & _RepresentantePedido.Pedido.Codigo & vbCrLf & _
              "   and C.Representante_Id      ='" & _RepresentantePedido.CodigoRepresentante & "'" & vbCrLf & _
              "   and C.EndRepresentante_Id   = " & _RepresentantePedido.CodigoEnderecoRepresentante & vbCrLf & _
              "   and PxI.TipoDeLancamento = 'N'" & vbCrLf
        If SomenteSalvos Then sql &= " and PxT.Tabela_Id is not null"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Pedidos")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim PxT As New PedidoXRepresentantesxTabelaDeComissao(pRepresentantePedido, TabelasSafra)
            PxT.IUD = IIf(row("Tabela_Id") = 0, "I", "")
            PxT.CodigoProduto = row("Produto_Id")
            PxT.NomeProduto = row("NomeProduto")
            PxT.CodigoTabela = row("Tabela_Id")
            PxT.NomeTabela = row("NomeTabela")
            Me.Add(PxT)
        Next

    End Sub
#End Region

#Region "Field"
    Private _TabelasSafra As ListTabelaDeComissao
    Private _RepresentantePedido As PedidoXRepresentante
#End Region

#Region "Property"
    Public ReadOnly Property TabelasSafra(Optional ByVal Atualizada As Boolean = False) As ListTabelaDeComissao
        Get
            If Atualizada Then
                _TabelasSafra = Nothing
            End If

            If _TabelasSafra Is Nothing Then _TabelasSafra = New ListTabelaDeComissao(True, RepresentantePedido.Pedido.CodigoSafra)
            Return _TabelasSafra
        End Get
    End Property

    Public ReadOnly Property RepresentantePedido() As PedidoXRepresentante
        Get
            Return _RepresentantePedido
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub AtualizarLista()
        For Each prod In RepresentantePedido.Pedido.Itens
            Dim codigo As String = prod.CodigoProduto
            If Not Me.Exists(Function(s) s.CodigoProduto = codigo AndAlso s.IUD <> "D") Then
                Dim PxT As New PedidoXRepresentantesxTabelaDeComissao(_RepresentantePedido, TabelasSafra(True))
                PxT.IUD = "I"
                PxT.CodigoProduto = prod.CodigoProduto
                PxT.NomeProduto = prod.Descricao
                Me.Add(PxT)
            End If
        Next
    End Sub

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As PedidoXRepresentantesxTabelaDeComissao In Me
            If RepresentantePedido.IUD = "D" Or RepresentantePedido.IUD = "I" Then item.IUD = RepresentantePedido.IUD
            If RepresentantePedido.PercentualFixo Or RepresentantePedido.IUD = "C" Then item.IUD = "D"
            item.SalvarSql(Sqls)
        Next
    End Sub
#End Region

End Class

'****************************************************************************************************************************************************************
'*****************************************   Classe Base PedidoXRepresentantexTabelaDeComissao   ****************************************************************
'*****************************************************************************************************************************************************************
Public Class PedidoXRepresentantesxTabelaDeComissao

#Region "Construtor"
    Public Sub New(ByRef pRepresentantePedido As PedidoXRepresentante, ByRef pTabelasSafra As ListTabelaDeComissao, Optional ByVal pProduto As String = "", Optional ByVal pTabela As Integer = 0)
        _RepresentantePedido = pRepresentantePedido
        _TabelasSafra = pTabelasSafra

        If pProduto.Length > 0 Then _CodigoProduto = pProduto

        If pTabela > 0 Then _CodigoTabela = pTabela
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _RepresentantePedido As PedidoXRepresentante
    Private _CodigoProduto As Integer
    Private _Produto As Produto
    Private _NomeProduto As String = ""
    Private _CodigoTabela As Integer
    Private _Tabela As TabelaDeComissao
    Private _NomeTabela As String = ""
    Private _TabelasSafra As ListTabelaDeComissao
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

    Public ReadOnly Property RepresentantePedido() As PedidoXRepresentante
        Get
            Return _RepresentantePedido
        End Get
    End Property

    Public Property CodigoProduto() As Integer
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As Integer)
            _CodigoProduto = value
            _Produto = Nothing
            _NomeProduto = ""
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing And _CodigoProduto > 0 Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property NomeProduto() As String
        Get
            If _NomeProduto.Length = 0 AndAlso Not Produto Is Nothing Then _NomeProduto = Produto.Nome
            Return _NomeProduto
        End Get
        Set(ByVal value As String)
            _NomeProduto = value
        End Set
    End Property

    Public Property CodigoTabela() As Integer
        Get
            Return _CodigoTabela
        End Get
        Set(ByVal value As Integer)
            _CodigoTabela = value
            _Tabela = Nothing
            _NomeTabela = ""
        End Set
    End Property

    Public Property Tabela() As TabelaDeComissao
        Get
            If _Tabela Is Nothing And _CodigoTabela > 0 Then _Tabela = New TabelaDeComissao(_CodigoTabela)
            Return _Tabela
        End Get
        Set(ByVal value As TabelaDeComissao)
            _Tabela = value
        End Set
    End Property

    Public Property NomeTabela() As String
        Get
            If _NomeTabela.Length = 0 AndAlso Not Tabela Is Nothing Then _NomeTabela = Tabela.Descricao
            Return _NomeTabela
        End Get
        Set(ByVal value As String)
            _NomeTabela = value
        End Set
    End Property

    Public ReadOnly Property TabelasSafra() As ListTabelaDeComissao
        Get
            Return _TabelasSafra
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
                Sql = "Insert Into PedidoXTabelaDeComissao(Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id, Representante_Id, EndRepresentante_Id, Tabela_Id)" & vbCrLf & _
                      " values('" & RepresentantePedido.Pedido.CodigoEmpresa & "'," & RepresentantePedido.Pedido.EnderecoEmpresa & "," & RepresentantePedido.Pedido.Codigo & ",'" & _CodigoProduto & "','" & RepresentantePedido.CodigoRepresentante & "'," & RepresentantePedido.CodigoEnderecoRepresentante & "," & _CodigoTabela & ")"
                Sqls.Add(Sql)
            Case "D"
                Sql = "Delete PedidoXTabelaDeComissao " & vbCrLf & _
                      " where Empresa_Id          ='" & RepresentantePedido.Pedido.CodigoEmpresa & "'" & _
                      "   and EndEmpresa_Id       = " & RepresentantePedido.Pedido.EnderecoEmpresa & _
                      "   and Pedido_Id           = " & RepresentantePedido.Pedido.Codigo & _
                      "   and Produto_Id          ='" & _CodigoProduto & "'" & _
                      "   and Representante_Id    ='" & RepresentantePedido.CodigoRepresentante & "'" & _
                      "   and EndRepresentante_Id = " & RepresentantePedido.CodigoEnderecoRepresentante
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class