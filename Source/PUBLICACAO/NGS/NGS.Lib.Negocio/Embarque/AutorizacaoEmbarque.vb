Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

'***********************************************************************************************************************************************************************************
'*******************************************  LISTA CLASSE BASE EMBARQUE X ENTREGA X ITEM EMBARQE normal/complemento/estorno *******************************************************
'***********************************************************************************************************************************************************************************
<Serializable()> _
Public Class ListLancamentoProdutoEmbarque
    Inherits List(Of LancamentoProdutoEmbarque)

#Region "Construtor"
    Public Sub New(pEmbarquexEntregaProduto As EmbarqueXEntregaProduto)
        _Parent = pEmbarquexEntregaProduto

        Dim sql As String
        sql = "SELECT AE.Empresa_Id," & vbCrLf & _
              "       AE.EndEmpresa_Id," & vbCrLf & _
              "       AE.Pedido_Id," & vbCrLf & _
              "       AE.Produto_Id," & vbCrLf & _
              "       AE.Lancamento_Id," & vbCrLf & _
              "       AE.Entrega_Id," & vbCrLf & _
              "       AE.EndEntrega_Id," & vbCrLf & _
              "       AE.TipoDeLancamento," & vbCrLf & _
              "       AE.Quantidade," & vbCrLf & _
              "       AE.Movimento," & vbCrLf & _
              "       AE.UsuarioInclusao" & vbCrLf & _
              "  FROM AutEmbarque AE" & vbCrLf & _
              " INNER JOIN Pedidos P" & vbCrLf & _
              "    ON P.Empresa_Id    = AE.Empresa_Id" & vbCrLf & _
              "   AND P.EndEmpresa_Id = AE.EndEmpresa_Id" & vbCrLf & _
              "   AND P.Pedido_id     = AE.Pedido_Id" & vbCrLf & _
              " WHERE AE.Empresa_Id    = " & pEmbarquexEntregaProduto.ParentEntrega.ParentEmbPedido.CodigoEmpresa & vbCrLf & _
              "   AND AE.EndEmpresa_Id = " & pEmbarquexEntregaProduto.ParentEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
              "   AND AE.Pedido_id     = " & pEmbarquexEntregaProduto.ParentEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
              "   AND AE.Entrega_Id    ='" & pEmbarquexEntregaProduto.ParentEntrega.CodigoClienteEntrega & "'" & vbCrLf & _
              "   AND AE.EndEntrega_Id = " & pEmbarquexEntregaProduto.ParentEntrega.EndClienteEntrega & vbCrLf & _
              "   AND AE.Produto_Id    = " & pEmbarquexEntregaProduto.CodigoProduto & vbCrLf & _
              " ORDER BY AE.Empresa_Id," & vbCrLf & _
              "          AE.EndEmpresa_Id," & vbCrLf & _
              "          AE.Pedido_Id," & vbCrLf & _
              "          AE.Produto_Id" & vbCrLf

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "AE")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim LAE As New LancamentoProdutoEmbarque(pEmbarquexEntregaProduto)
            'LAE.CodigoProduto = row("Produto_Id")
            LAE.NrLancamento = row("Lancamento_Id")
            LAE.TipoDeLancamento = row("TipoDeLancamento")
            LAE.Quantidade = row("Quantidade")
            LAE.Movimento = row("Movimento")
            LAE.UsuarioInclusao = row("UsuarioInclusao")
            Me.Add(LAE)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Parent As EmbarqueXEntregaProduto
#End Region

#Region "Property"
    Public ReadOnly Property Parent As EmbarqueXEntregaProduto
        Get
            Return _Parent
        End Get
    End Property
#End Region

End Class

'***********************************************************************************************************************************************************************************
'************************************************** CLASSE BASE EMBARQUE X ENTREGA X ITEM EMBARQE normal/complemento/estorno *******************************************************
'***********************************************************************************************************************************************************************************
<Serializable()> _
Public Class LancamentoProdutoEmbarque
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(pEmbarquexEntregaProduto As EmbarqueXEntregaProduto)
        _ParentEntregaProduto = pEmbarquexEntregaProduto
    End Sub
#End Region

#Region "Fields"
    Private _ParentEntregaProduto As EmbarqueXEntregaProduto
    Private _IUD As String

    Private _CodigoProduto As String = ""
    Private _Produto As Produto

    Private _NrLancamento As Integer
    Private _TipoDeLancamento As String

    Private _Quantidade As Decimal
    Private _Movimento As Date
    Private _UsuarioInclusao As String

    Private _QtdePedido As Decimal
    Private _QtdeCarregada As Decimal

#End Region

#Region "Property"
    Public ReadOnly Property ParentEntregaProduto() As EmbarqueXEntregaProduto
        Get
            Return _ParentEntregaProduto
        End Get
    End Property

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property NrLancamento() As Integer
        Get
            Return _NrLancamento
        End Get
        Set(ByVal value As Integer)
            _NrLancamento = value
        End Set
    End Property

    Public Property TipoDeLancamento() As String
        Get
            Return _TipoDeLancamento
        End Get
        Set(ByVal value As String)
            _TipoDeLancamento = value
        End Set
    End Property

    Public Property Quantidade() As Decimal
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Decimal)
            _Quantidade = value
        End Set
    End Property

    Public Property Movimento() As Date
        Get
            Return _Movimento
        End Get
        Set(ByVal value As Date)
            _Movimento = value
        End Set
    End Property

    Public Property QtdePedido() As Decimal
        Get
            Return _QtdePedido
        End Get
        Set(ByVal value As Decimal)
            _QtdePedido = value
        End Set
    End Property

    Public Property QtdeCarregada() As Decimal
        Get
            Return _QtdeCarregada
        End Get
        Set(ByVal value As Decimal)
            _QtdeCarregada = value
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
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        If Banco.GravaBanco(Sqls) Then
            IUD = ""
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Dim sql As String = ""
        Select Case _IUD
            Case "I"
                sql = "INSERT INTO AutEmbarque (Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id, Lancamento_Id, Entrega_Id, EndEntrega_Id, TipoDeLancamento, Quantidade, Movimento, UsuarioInclusao) " & vbCrLf & _
                      "     VALUES ('" & ParentEntregaProduto.ParentEntrega.ParentEmbPedido.CodigoEmpresa & "', " & ParentEntregaProduto.ParentEntrega.ParentEmbPedido.EndEmpresa & "," & ParentEntregaProduto.ParentEntrega.ParentEmbPedido.CodigoPedido & ",'" & ParentEntregaProduto.CodigoProduto & "' " & vbCrLf & _
                      ", (select isnull(max(Lancamento_Id),0) + 1 from AutEmbarque where Empresa_Id = '" & ParentEntregaProduto.ParentEntrega.ParentEmbPedido.CodigoEmpresa & "' and EndEmpresa_Id = '" & ParentEntregaProduto.ParentEntrega.ParentEmbPedido.EndEmpresa & "' and Pedido_Id = '" & ParentEntregaProduto.ParentEntrega.ParentEmbPedido.CodigoPedido & "' and Produto_Id = '" & ParentEntregaProduto.CodigoProduto & "'), " & vbCrLf & _
                      "'" & ParentEntregaProduto.ParentEntrega.CodigoClienteEntrega & "', " & ParentEntregaProduto.ParentEntrega.EndClienteEntrega & ", '" & Me.TipoDeLancamento & "'," & Str(Me.Quantidade) & ",'" & Me.Movimento.ToString("yyyy-MM-dd") & "','" & Me.UsuarioInclusao & "')"
                Sqls.Add(sql)
            Case "U"
                sql = "UPDATE AutorizacaoDeEmbarque SET" & _
                      "    TipoDeLancamento  = '" & Me.TipoDeLancamento & "'" & vbCrLf & _
                      "   ,Quantidade        =  " & Str(Me.Quantidade) & vbCrLf & _
                      "   ,Movimento         = '" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,UsuarioInclusao   = '" & Me.UsuarioInclusao & "'" & vbCrLf & _
                      " WHERE Empresa_Id     = '" & ParentEntregaProduto.ParentEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id  =  " & ParentEntregaProduto.ParentEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
                      "   AND Pedido_Id      =  " & ParentEntregaProduto.ParentEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
                      "   AND Entrega_Id     = '" & ParentEntregaProduto.ParentEntrega.CodigoClienteEntrega & "'" & vbCrLf & _
                      "   AND EndEntrega_Id  =  " & ParentEntregaProduto.ParentEntrega.EndClienteEntrega & "" & vbCrLf & _
                      "   AND Produto_Id     = '" & ParentEntregaProduto.CodigoProduto & "'" & vbCrLf & _
                      "   AND Lancamento_Id  = '" & Me.NrLancamento & "'"
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE AutorizacaoDeEmbarque " & _
                     " WHERE Empresa_Id     = '" & ParentEntregaProduto.ParentEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id =  " & ParentEntregaProduto.ParentEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
                      "   AND Pedido_Id     =  " & ParentEntregaProduto.ParentEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
                      "   AND Entrega_Id    = '" & ParentEntregaProduto.ParentEntrega.CodigoClienteEntrega & "'" & vbCrLf & _
                      "   AND EndEntrega_Id =  " & ParentEntregaProduto.ParentEntrega.EndClienteEntrega & "" & vbCrLf & _
                      "   AND Produto_Id    = '" & ParentEntregaProduto.CodigoProduto & "'" & vbCrLf & _
                      "   AND Lancamento_Id = '" & Me.NrLancamento & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
