Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

'***************************************************************************************************************************************
'*****************************************  LISTA CLASSE EMBARQUE x ENTREGA PRODUTOS  **************************************************
'***************************************************************************************************************************************
<Serializable()> _
Public Class ListEmbarqueXEntregaProduto
    Inherits List(Of EmbarqueXEntregaProduto)

#Region "Construtor"
    Public Sub New(pEmbarqueEntrega As EmbarqueXEntrega)
        _ParentEntrega = pEmbarqueEntrega

        Dim sql As String = ""
        sql = "     Select pxi.Produto_Id,                                                                " & vbCrLf & _
         "            prd.Nome,                                                                           " & vbCrLf & _
         "            isnull(Rom.Cliente_id,'') as Cliente_id,                                            " & vbCrLf & _
         "            isnull(Rom.EndCliente_Id,0) as EndCliente_Id,                                       " & vbCrLf & _
         "                                                                                                " & vbCrLf & _
         "            pxi.QtdePedido,                                                                     " & vbCrLf & _
         "            isnull(autcar.AutorizadoCarregar,0) as AutorizadoCarregar,                                 " & vbCrLf & _
         "                                                                                                " & vbCrLf & _
         "            SUM(isnull(Case                                                                     " & vbCrLf & _
         "                     when Rom.EntradaSaida = 'S'                                                " & vbCrLf & _
         "                       then Rom.PesoSaida                                                       " & vbCrLf & _
         "                       else Rom.PesoEntrada                                                     " & vbCrLf & _
         "                   end,0)) as QtdeEmbarcado,                                                    " & vbCrLf & _
         "            SUM(isnull(Case                                                                     " & vbCrLf & _
         "                     when Rom.EntradaSaida = 'S'                                                " & vbCrLf & _
         "                       then Rom.PesoEntrada                                                     " & vbCrLf & _
         "                       else Rom.PesoSaida                                                       " & vbCrLf & _
         "     		         end,0)) as QtdeDevolvido                                                   " & vbCrLf & _
         "       from (Select pxi.Empresa_Id,                                                             " & vbCrLf & _
         "                    pxi.EndEmpresa_Id,                                                          " & vbCrLf & _
         "                    pxi.Pedido_Id,                                                              " & vbCrLf & _
         "                    pxi.Produto_Id,                                                             " & vbCrLf & _
         "                    SUM(Case                                                                    " & vbCrLf & _
         "                          when pxi.TipoDeLancamento = 'E'                                       " & vbCrLf & _
         "                            Then pxi.Quantidade * - 1                                           " & vbCrLf & _
         "                            else pxi.Quantidade                                                 " & vbCrLf & _
         "                        end) QtdePedido,                                                        " & vbCrLf & _
         "                    SUM(Case                                                                    " & vbCrLf & _
         "                          when pxi.TipoDeLancamento = 'E'                                       " & vbCrLf & _
         "                            Then isnull(pxi.QuantidadeEstimada,0) * - 1                         " & vbCrLf & _
         "                            else isnull(pxi.QuantidadeEstimada,0)                               " & vbCrLf & _
         "                        end) QtdeEstimada                                                       " & vbCrLf & _
         "               from Pedidos P                                                                   " & vbCrLf & _
         "              inner join PedidoXItemxLancamento pxi                                             " & vbCrLf & _
         "                 on P.Empresa_Id    = PxI.Empresa_Id                                            " & vbCrLf & _
         "                and P.EndEmpresa_Id = PxI.EndEmpresa_Id                                         " & vbCrLf & _
         "                and P.Pedido_Id     = PxI.Pedido_Id                                             " & vbCrLf & _
         "              Group by pxi.Empresa_Id,                                                          " & vbCrLf & _
         "                       pxi.EndEmpresa_Id,                                                       " & vbCrLf & _
         "                       pxi.Pedido_Id,                                                           " & vbCrLf & _
         "                       pxi.Produto_Id                                                           " & vbCrLf & _
         "                  ) PxI                                                                         " & vbCrLf & _
         "      inner join Produtos prd                                                                   " & vbCrLf & _
         "         on pxi.produto_id = prd.produto_id                                                     " & vbCrLf & _
         "       Left Join(                                                                               " & vbCrLf & _
         "                 Select R.Empresa_Id,                                                           " & vbCrLf & _
         "                        R.EndEmpresa_Id,                                                        " & vbCrLf & _
         "                        R.Pedido,                                                               " & vbCrLf & _
         "                        NFxR.cliente_id,                                                        " & vbCrLf & _
         "                        NFxR.EndCliente_id,                                                     " & vbCrLf & _
         "                        R.Produto,                                                              " & vbCrLf & _
         "                        R.EntradaSaida,                                                         " & vbCrLf & _
         "                        SUM(case                                                                " & vbCrLf & _
         "                              when R.EntradaSaida = 'E'                                         " & vbCrLf & _
         "                                then R.PesoLiquido                                              " & vbCrLf & _
         "                                else 0                                                          " & vbCrLf & _
         "                            end) as PesoEntrada,                                                " & vbCrLf & _
         "                        SUM(case                                                                " & vbCrLf & _
         "                              when R.EntradaSaida = 'S'                                         " & vbCrLf & _
         "                                then R.PesoLiquido                                              " & vbCrLf & _
         "                                else 0                                                          " & vbCrLf & _
         "                            end) as PesoSaida                                                   " & vbCrLf & _
         "                   from Romaneios R                                                             " & vbCrLf & _
         "                  inner join NotasfiscaisxRomaneios NFxR                                        " & vbCrLf & _
         "                     on R.Empresa_id    = NFxR.Empresa_id                                       " & vbCrLf & _
         "                    And R.EndEmpresa_id = NFxR.EndEmpresa_id                                    " & vbCrLf & _
         "                    and R.Romaneio_id   = NFxR.Romaneio_Id                                      " & vbCrLf & _
         "                  group by R.Empresa_Id,                                                        " & vbCrLf & _
         "                           R.EndEmpresa_Id,                                                     " & vbCrLf & _
         "                           R.Pedido,                                                            " & vbCrLf & _
         "                           NFxR.cliente_id,                                                     " & vbCrLf & _
         "                           NFxR.EndCliente_id,                                                  " & vbCrLf & _
         "                           R.Produto,                                                           " & vbCrLf & _
         "                           R.EntradaSaida                                                       " & vbCrLf & _
         "                ) Rom                                                                           " & vbCrLf & _
         "           on PxI.Empresa_Id    = Rom.Empresa_Id                                                " & vbCrLf & _
         "          and PxI.EndEmpresa_Id = Rom.EndEmpresa_Id                                             " & vbCrLf & _
         "          and PxI.Pedido_Id     = Rom.Pedido                                                    " & vbCrLf & _
         "          and PxI.Produto_Id    = Rom.Produto                                                   " & vbCrLf & _
         "          and Rom.Cliente_Id    = '" & pEmbarqueEntrega.CodigoClienteEntrega & "'" & vbCrLf & _
         "          and Rom.EndCliente_id = " & pEmbarqueEntrega.EndClienteEntrega & vbCrLf & _
         "         left Join (SELECT ac.Empresa,                                                          " & vbCrLf & _
         "                           ac.EndEmpresa,                                                       " & vbCrLf & _
         "                           ac.Pedido,                                                           " & vbCrLf & _
         "                           aci.Produto_Id,                                                      " & vbCrLf & _
         "                           ISNULL(SUM(ACI.QuantidadeProgramado),0) as AutorizadoCarregar        " & vbCrLf & _
         "     				 FROM AutCarregamento AC                                                    " & vbCrLf & _
         "     				INNER JOIN AutCarregamentoXItens ACI                                        " & vbCrLf & _
         "     				   ON ACI.Empresa_Id     = AC.Empresa                                       " & vbCrLf & _
         "     				  AND AC.Carregamento_Id = ACI.Carregamento_Id                              " & vbCrLf & _
         "     				  AND AC.EndEmpresa      = ACI.EndEmpresa_Id                                " & vbCrLf & _
         "     				  AND AC.Pedido          = ACI.Pedido_Id                                    " & vbCrLf & _
         "     				  AND AC.Entrega         = ACI.Entrega_Id                                   " & vbCrLf & _
         "     				  AND AC.EndEntrega      = ACI.EndEntrega_Id                                " & vbCrLf & _
         "     				Where AC.Situacao    = 1                                                    " & vbCrLf & _
         "     			    group by ac.Empresa,                                                        " & vbCrLf & _
         "                              ac.EndEmpresa,                                                    " & vbCrLf & _
         "                              ac.Pedido,                                                        " & vbCrLf & _
         "                              aci.Produto_Id                                                    " & vbCrLf & _
         "     			) AutCar                                                                        " & vbCrLf & _
         "           on PxI.Empresa_Id    = AutCar.Empresa                                                " & vbCrLf & _
         "          and PxI.EndEmpresa_Id = AutCar.EndEmpresa                                             " & vbCrLf & _
         "          and PxI.Pedido_Id     = AutCar.Pedido                                                 " & vbCrLf & _
         "          and PxI.Produto_Id    = AutCar.Produto_Id                                             " & vbCrLf & _
         "                                                                                                " & vbCrLf & _
         "      WHERE pxi.Empresa_Id  = '" & pEmbarqueEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
       "        and pxi.EndEmpresa_Id =   " & pEmbarqueEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
       "        and pxi.Pedido_Id     =   " & pEmbarqueEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
       "      GROUP BY                                                                                  " & vbCrLf & _
       "          pxi.Produto_Id,                                                                       " & vbCrLf & _
       "          prd.Nome,                                                                             " & vbCrLf & _
       "          autcar.AutorizadoCarregar,                                                           " & vbCrLf & _
       "          pxi.QtdePedido,                                                                     " & vbCrLf & _
       "          isnull(Rom.Cliente_id,''),                                                            " & vbCrLf & _
       "          isnull(Rom.EndCliente_Id, 0)                                                          " & vbCrLf

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "EmbarqueXEntregaProduto")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Me.QtdePedidoPrd = ds.Tables(0).Rows(0)("QtdePedido")

        For Each row In ds.Tables(0).Rows
            Me.QtdeEmbarcadoPrd += ds.Tables(0).Rows(0)("QtdeEmbarcado")
            Me.QtdeDevolvidoPrd += ds.Tables(0).Rows(0)("QtdeDevolvido")

            Dim prd As New EmbarqueXEntregaProduto(pEmbarqueEntrega)
            prd.CodigoProduto = row("Produto_id")
            prd.DescricaoProduto = row("Nome")
            prd.QtdePedido = row("QtdePedido")
            prd.QtdeEmbarcado = row("QtdeEmbarcado")
            prd.QtdeDevolvido = row("QtdeDevolvido")
            prd.QtdeAutorizadoCarregar = row("AutorizadoCarregar")
            Me.Add(prd)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _ParentEntrega As EmbarqueXEntrega
    Private _QtdePedido As Decimal
    Private _QtdeEmbarcado As Decimal
    Private _QtdeDevolvido As Decimal
#End Region

#Region "Property"
    Public ReadOnly Property ParentEntrega As EmbarqueXEntrega
        Get
            Return _ParentEntrega
        End Get
    End Property

    'DO PEDIDO
    Public Property QtdePedidoPrd As Decimal
        Get
            Return _QtdePedido
        End Get
        Set(value As Decimal)
            _QtdePedido = value
        End Set
    End Property

    'DO ROMANEIO
    Public Property QtdeEmbarcadoPrd As Decimal
        Get
            Return _QtdeEmbarcado
        End Get
        Set(value As Decimal)
            _QtdeEmbarcado = value
        End Set
    End Property

    Public Property QtdeDevolvidoPrd As Decimal
        Get
            Return _QtdeDevolvido
        End Get
        Set(value As Decimal)
            _QtdeDevolvido = value
        End Set
    End Property

    'DO PRODUTO
    Public ReadOnly Property QtdeNormalPrd As Decimal
        Get
            Return ParentEntrega.ParentEmbPedido.LocaisDeEntrega.SelectMany(Function(s) s.Produtos).Sum(Function(s) s.QtdeNormal)
        End Get
    End Property

    Public ReadOnly Property QtdeComplementoPrd As Decimal
        Get
            Return ParentEntrega.ParentEmbPedido.LocaisDeEntrega.SelectMany(Function(s) s.Produtos).Sum(Function(s) s.QtdeComplemento)
        End Get
    End Property

    Public ReadOnly Property QtdeEstornoPrd As Decimal
        Get
            Return ParentEntrega.ParentEmbPedido.LocaisDeEntrega.SelectMany(Function(s) s.Produtos).Sum(Function(s) s.QtdeEstorno)
        End Get
    End Property

    Public ReadOnly Property QtdeAutorizadoPrd As Decimal
        Get
            Return ParentEntrega.ParentEmbPedido.LocaisDeEntrega.SelectMany(Function(s) s.Produtos).Sum(Function(s) s.Quantidade)
        End Get
    End Property

    'DO LOCAL DE ENTREGA
    Public ReadOnly Property QtdeNormalLE As Decimal
        Get
            Return Me.Sum(Function(s) s.QtdeNormal)
        End Get
    End Property

    Public ReadOnly Property QtdeEstornoLE As Decimal
        Get
            Return Me.Sum(Function(s) s.QtdeEstorno)
        End Get
    End Property

    Public ReadOnly Property QtdeComplementoLE As Decimal
        Get
            Return Me.Sum(Function(s) s.QtdeComplemento)
        End Get
    End Property

    Public ReadOnly Property QtdeAutorizadoLE As Decimal
        Get
            Return Me.Sum(Function(s) s.Quantidade)
        End Get
    End Property

    Public ReadOnly Property QtdeEmbarcadoLE As Decimal
        Get
            Return Me.Sum(Function(s) s.QtdeEmbarcado)
        End Get
    End Property

    Public ReadOnly Property QtdeDevolvidoLE As Decimal
        Get
            Return Me.Sum(Function(s) s.QtdeDevolvido)
        End Get
    End Property

    'CALCULADO
    Public ReadOnly Property MaximoEstornoLE() As Decimal
        Get
            Return Me.QtdeAutorizadoLE - Me.QtdeEmbarcadoLE
        End Get
    End Property

    Public ReadOnly Property MaximoAutorizarPrd As Decimal
        Get
            If Me.QtdePedidoPrd + Me.QtdeDevolvidoPrd < Me.QtdeAutorizadoPrd Then
                Return Decimal.Zero
            End If
            Return Me.QtdePedidoPrd + Me.QtdeDevolvidoPrd - Me.QtdeAutorizadoPrd
        End Get
    End Property

    Public ReadOnly Property MaximoEstornoPrd() As Decimal
        Get
            Return Me.QtdeAutorizadoPrd - Me.QtdeEmbarcadoPrd
        End Get
    End Property
#End Region

End Class

'***************************************************************************************************************************************
'****************************************  CLASSE BASE EMBARQUE x ENTREGA PRODUTOS  ****************************************************
'***************************************************************************************************************************************
<Serializable()> _
Public Class EmbarqueXEntregaProduto
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(pEmbarqueEntrega As EmbarqueXEntrega)
        _ParentEntrega = pEmbarqueEntrega
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _ParentEntrega As EmbarqueXEntrega
    Private _CodigoProduto As Integer
    Private _DescricaoProduto As String
    Private _Produto As Produto
    Private _QtdePedido As Decimal
    Private _QtdeEmbarcando As Decimal
    Private _QtdeAutorizadoCarregar As Decimal
    Private _QtdeEmbarcado As Decimal
    Private _QtdeDevolvido As Decimal
    Private _QtdeNormal As Decimal
    Private _QtdeEstorno As Decimal
    Private _QtdeComplemento As Decimal
    Private _Lancamentos As ListLancamentoProdutoEmbarque
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

    Public ReadOnly Property ParentEntrega As EmbarqueXEntrega
        Get
            Return _ParentEntrega
        End Get
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

    Public Property DescricaoProduto As String
        Get
            Return _DescricaoProduto
        End Get
        Set(value As String)
            _DescricaoProduto = value
        End Set
    End Property

    Public Property Produto As Produto
        Get
            If _Produto Is Nothing And Me.CodigoProduto > 0 Then _Produto = New Produto(Me.CodigoProduto)
            Return _Produto
        End Get
        Set(value As Produto)
            _Produto = value
        End Set
    End Property

    'Pedido 
    Public Property QtdePedido As Decimal
        Get
            Return _QtdePedido
        End Get
        Set(value As Decimal)
            _QtdePedido = value
        End Set
    End Property

    'Autorizacao de Carregamento
    Public Property QtdeAutorizadoCarregar() As Decimal
        Get
            Return _QtdeAutorizadoCarregar
        End Get
        Set(ByVal value As Decimal)
            _QtdeAutorizadoCarregar = value
        End Set
    End Property

    Public Property QtdeEmbarcando() As Decimal
        Get
            Return _QtdeEmbarcando
        End Get
        Set(ByVal value As Decimal)
            _QtdeEmbarcando = value
        End Set
    End Property


    'Romaneio
    Public Property QtdeEmbarcado As Decimal
        Get
            Return _QtdeEmbarcado
        End Get
        Set(value As Decimal)
            _QtdeEmbarcado = value
        End Set
    End Property

    Public Property QtdeDevolvido As Decimal
        Get
            Return _QtdeDevolvido
        End Get
        Set(value As Decimal)
            _QtdeDevolvido = value
        End Set
    End Property

    'Lancamentos Normal / Complemento / Estorno
    Public ReadOnly Property QtdeNormal As Decimal
        Get
            Return Me.Lancamentos.Where(Function(s) s.TipoDeLancamento = "N").Sum(Function(s) s.Quantidade)
        End Get
    End Property

    Public ReadOnly Property QtdeEstorno As Decimal
        Get
            Return Me.Lancamentos.Where(Function(s) s.TipoDeLancamento = "E").Sum(Function(s) s.Quantidade)
        End Get
    End Property

    Public ReadOnly Property QtdeComplemento As Decimal
        Get
            Return Me.Lancamentos.Where(Function(s) s.TipoDeLancamento = "C").Sum(Function(s) s.Quantidade)
        End Get
    End Property

    Public ReadOnly Property Quantidade As Decimal
        Get
            Return Me.Lancamentos.Sum(Function(s)
                                          Dim Soma As Decimal
                                          If s.TipoDeLancamento = "E" Then
                                              Soma += s.Quantidade * -1
                                          Else
                                              Soma += s.Quantidade
                                          End If
                                          Return Soma - _QtdeAutorizadoCarregar
                                      End Function)

        End Get
    End Property

    'Calculados
    Public ReadOnly Property MaximoAutorizar As Decimal
        Get
            Return Me.QtdePedido + Me.QtdeDevolvido - Me.Quantidade
        End Get
    End Property

    Public ReadOnly Property MaximoEstorno() As Decimal
        Get
            Return Me.Quantidade - Me.QtdeEmbarcado
        End Get
    End Property

    'Lancamento
    Public Property Lancamentos As ListLancamentoProdutoEmbarque
        Get
            If _Lancamentos Is Nothing Then _Lancamentos = New ListLancamentoProdutoEmbarque(Me)
            Return _Lancamentos
        End Get
        Set(value As ListLancamentoProdutoEmbarque)
            _Lancamentos = value
        End Set
    End Property
#End Region

End Class

