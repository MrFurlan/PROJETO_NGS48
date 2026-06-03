Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaPedidos
    Inherits BaseUserControl

    Public Property MyParameters() As Dictionary(Of String, Object)
        Get
            Return CType(Session("MyParameters" & HID.Value), Dictionary(Of String, Object))
        End Get
        Set(ByVal value As Dictionary(Of String, Object))
            Session("MyParameters" & HID.Value) = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra)
            ddlSafra.Parent.Visible = False
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        Session.Remove("_MainUserControl")
        txtConsultaPedido.Text = String.Empty
        GridPedidos.DataSource = New List(Of Object)()
        GridPedidos.DataBind()
        gridPedidos2.DataSource = New List(Of Object)()
        gridPedidos2.DataBind()
    End Sub

    Public Function BindGridView(ByVal parametros As Dictionary(Of String, Object)) As Integer
        ddlSafra.Parent.Visible = False
        MyParameters = parametros
        Select Case Session("ssCampo" & HID.Value)
            Case "Romaneio" 
                Return CargaPedidosRomaneio()
            Case "Pedidos"
                Return ListarPedidos(parametros)
            Case "Laudo"
                ddlSafra.Parent.Visible = True
                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra)
                If parametros IsNot Nothing AndAlso parametros.ContainsKey("safra") Then
                    ddlSafra.SelectedValue = parametros("safra").ToString
                End If
                Return CargaPedidosLaudo(parametros)
            Case "Carregamento"
                Return CargaPedidosCarregamento(parametros)
            Case Else
                Return CargaPedidos(parametros)
        End Select
        Return 0
    End Function

    Private Function CargaPedidosCarregamento(ByVal parametros As Dictionary(Of String, Object)) As Integer
        Try
            Dim pedidos As String = ""
            If parametros IsNot Nothing AndAlso parametros.ContainsKey("pedido") Then
                pedidos = parametros("pedido")
            End If

            Dim sql As String = "SELECT P.Empresa_Id     as CodigoEmpresa," & vbCrLf & _
                                   "       P.EndEmpresa_Id  as EnderecoEmpresa," & vbCrLf & _
                                   "       P.Pedido_Id      as Pedido, " & vbCrLf & _
                                   "       P.Operacao       as CodigoOperacao," & vbCrLf & _
                                   "       P.SubOperacao    as CodigoSubOperacao," & vbCrLf & _
                                   "       SO.Descricao     as DescricaoSubOperacao," & vbCrLf & _
                                   "       CAST(P.Operacao as varchar) + '-' + CAST(P.SubOperacao as varchar) + ' - ' + SO.Descricao as SubOp," & vbCrLf & _
                                   "       C.Nome           as NomeCliente," & vbCrLf & _
                                   "       P.DataPedido," & vbCrLf & _
                                   "       M.Descricao      as Moeda," & vbCrLf & _
                                   "       PxI.Saldo        as Quantidade," & vbCrLf & _
                                   "       CASE WHEN M.Classificacao = 'O'" & vbCrLf & _
                                   "              THEN isnull(Liq.ValorOficial,0)" & vbCrLf & _
                                   "              ELSE isnull(Liq.ValorMoeda,0)" & vbCrLf & _
                                   "       END as Total," & vbCrLf & _
                                   "       [dbo].ProdutosPedido(P.Empresa_Id, P.EndEmpresa_Id, P.Pedido_Id) as Itens" & vbCrLf & _
                                   "  FROM Pedidos P " & vbCrLf & _
                                   " INNER JOIN (" & vbCrLf & _
                                   "             SELECT PED.Empresa_Id,  " & vbCrLf & _
                                   "                    PED.EndEmpresa_Id,  " & vbCrLf & _
                                   "                    PED.Pedido_Id,  " & vbCrLf & _
                                   "                    CASE" & vbCrLf & _
                                   "                      WHEN PD.Agrupar = 'N'" & vbCrLf & _
                                   "                        THEN SUM(CASE" & vbCrLf & _
                                   "                                   WHEN PEDxI.TipoDeLancamento = 'E'" & vbCrLf & _
                                   "                                     THEN PEDxI.Quantidade * - 1" & vbCrLf & _
                                   "                                     ELSE PEDxI.Quantidade" & vbCrLf & _
                                   "                                 END)" & vbCrLf & _
                                   "                        ELSE 0" & vbCrLf & _
                                   "                    END as Saldo" & vbCrLf & _
                                   "               FROM Pedidos PED" & vbCrLf & _
                                   "              INNER JOIN PedidoXItemxLancamento PEDxI" & vbCrLf & _
                                   "                 ON PED.Empresa_Id    = PEDxI.Empresa_Id" & vbCrLf & _
                                   "                AND PED.EndEmpresa_Id = PEDxI.EndEmpresa_Id" & vbCrLf & _
                                   "                AND PED.Pedido_Id     = PEDxI.Pedido_Id" & vbCrLf & _
                                   "              INNER JOIN Produtos PD" & vbCrLf & _
                                   "                 ON PD.Produto_Id = PEDxI.Produto_Id" & vbCrLf

            If parametros IsNot Nothing AndAlso parametros.ContainsKey("granel") Then
                sql &= "          AND PD.Agrupar = 'N'" & vbCrLf
            End If

            sql &= "              GROUP BY PED.Empresa_Id, PED.EndEmpresa_Id, PED.Pedido_Id, PD.Agrupar" & vbCrLf & _
                   "            ) as PxI  " & vbCrLf & _
                   "    ON PxI.Empresa_Id = P.Empresa_Id" & vbCrLf & _
                   "   AND PxI.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
                   "   AND PxI.Pedido_Id = P.Pedido_Id" & vbCrLf & _
                   " INNER JOIN Clientes C" & vbCrLf & _
                   "    ON P.Cliente    = C.cliente_id" & vbCrLf & _
                   "   AND P.EndCliente = C.Endereco_id" & vbCrLf & _
                   " INNER JOIN Moedas M" & vbCrLf & _
                   "    ON P.moeda = M.Moeda_id" & vbCrLf & _
                   "  LEFT JOIN (SELECT Empresa_id," & vbCrLf & _
                   "                    EndEmpresa_id," & vbCrLf & _
                   "                    Pedido_id," & vbCrLf & _
                   "                    Sum(ValorOficial) as ValorOficial," & vbCrLf & _
                   "                    Sum(ValorMoeda) as ValorMoeda" & vbCrLf & _
                   "               FROM PedidosXEncargos" & vbCrLf & _
                   "              WHERE Encargo_id = 'LIQUIDO'" & vbCrLf & _
                   "              GROUP BY Empresa_id, EndEmpresa_id, Pedido_Id" & vbCrLf & _
                   "             ) as Liq" & vbCrLf & _
                   "    ON P.Empresa_Id    = liq.Empresa_Id" & vbCrLf & _
                   "   AND P.EndEmpresa_Id = liq.EndEmpresa_Id" & vbCrLf & _
                   "   AND P.Pedido_Id     = liq.Pedido_Id" & vbCrLf & _
                   " INNER JOIN SubOperacoes SO " & vbCrLf & _
                   "    ON SO.Operacao_Id = P.Operacao " & vbCrLf & _
                   "   AND SO.SubOperacoes_Id = P.SubOperacao " & vbCrLf & _
                   " WHERE P.UnidadeDeNegocio IS NOT NULL " & vbCrLf & _
                   "   AND P.PedidoOrigem = 0 " & vbCrLf

            If Not String.IsNullOrWhiteSpace(pedidos) Then
                sql &= "AND P.Pedido_Id in (" & pedidos & ") " & vbCrLf
            End If

            sql &= "    ORDER BY P.Pedido_Id " & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Pedidos")

            Select Case ds.Tables(0).Rows.Count
                Case 0
                    Session.Remove("objPedidoSelecionado" & HID.Value)
                    Session.Remove("objConsultaPedidos" & HID.Value)
                    MsgBox(Me.Page, "Não foi encontrado nenhum pedido com os parâmetros informados!")
                    Popup.CloseDialog(Me.Page, "divConsultaPedidos", True, 500)
                Case 1
                    Dim row As DataRow = ds.Tables(0).Rows(0)
                    Dim objPedido As [Lib].Negocio.Pedido = New [Lib].Negocio.Pedido(row("CodigoEmpresa"), row("EnderecoEmpresa"), row("Pedido"))
                    Selecionar(objPedido)
                Case Else
                    Session("objConsultaPedidos" & HID.Value) = ds
                    GridPedidos.Visible = False
                    gridPedidos2.Visible = True
                    gridPedidos2.DataSource = ds
                    gridPedidos2.DataBind()
            End Select

            Return ds.Tables(0).Rows.Count
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Function

    Private Function CargaPedidos(ByVal parametros As Dictionary(Of String, Object)) As Integer
        Dim ds As New DataSet
        Dim CodigoProduto As String = ""
        If Not HttpContext.Current.Session("CodigoProduto") Is Nothing Then
            CodigoProduto = HttpContext.Current.Session("CodigoProduto")
            HttpContext.Current.Session.Remove("CodigoProduto")
        ElseIf parametros.ContainsKey("CodigoProduto") Then
            CodigoProduto = parametros("CodigoProduto")
        End If

        Dim CodigoSafra As String = ""
        If Not HttpContext.Current.Session("CodigoSafra") Is Nothing Then
            CodigoSafra = HttpContext.Current.Session("CodigoSafra")
            HttpContext.Current.Session.Remove("CodigoSafra")
        End If

        Dim AnoSafra As Integer
        If Not HttpContext.Current.Session("Ano") Is Nothing Then
            AnoSafra = HttpContext.Current.Session("Ano")
            HttpContext.Current.Session.Remove("Ano")
        End If


        Dim sql As String
        sql = "Select P.Empresa_id as CodigoEmpresa," & vbCrLf &
              "       P.EndEmpresa_id as EnderecoEmpresa," & vbCrLf &
              "       P.Pedido_Id as Pedido," & vbCrLf &
              "       P.DataPedido," & vbCrLf &
              "       P.DataEntrega," & vbCrLf &
              "       P.Operacao," & vbCrLf &
              "       P.SubOperacao," & vbCrLf &
              "       P.Safra," & vbCrLf &
              "       sbPed.Produto_Id as Produto," & vbCrLf &
              "       prd.Nome as NomeDoProduto," & vbCrLf &
              "       isnull(sbPed.QtdNormal,0) AS QtdNormal," & vbCrLf &
              "       case when isnull(sbPed.Total,0) = 0 or isnull(sbPed.QtdNormal,0) = 0 " & vbCrLf &
              "         then 0.00                                                          " & vbCrLf &
              "         else (isnull(sbPed.Total,0) / isnull(sbPed.QtdNormal,0))           " & vbCrLf &
              "       end as Unitario,                                                     " & vbCrLf &
              "       isnull(sbnf.EntregueNota,0) as EntregueNota," & vbCrLf &
              "       isnull(sbnf.DevolucaoNota,0) as DevolucaoNota" & vbCrLf &
              "  Into #Consulta" & vbCrLf &
              "  from Pedidos P" & vbCrLf &
              " inner Join (Select Pxi.Empresa_Id, Pxi.EndEmpresa_Id, Pxi.Pedido_Id, Pxi.Produto_Id, " & vbCrLf &
              "		               sum(case " & vbCrLf &
              "		            		 When Pxi.TipoDeLancamento = 'E' " & vbCrLf &
              "		            		   then Pxi.Quantidade * - 1 " & vbCrLf &
              "		            		   else Pxi.Quantidade " & vbCrLf &
              "		            	   end) as QtdNormal, " & vbCrLf &
              "                    sum(case" & vbCrLf &
              "		            		 When Pxi.TipoDeLancamento = 'E'" & vbCrLf &
              "		            		   then Pxi.TotalOficial * - 1" & vbCrLf &
              "		            		   else Pxi.TotalOficial" & vbCrLf &
              "		            	   end) as Total" & vbCrLf &
              "		          from PedidoxItemxLancamento Pxi" & vbCrLf &
              "		         Group By Pxi.Empresa_Id, Pxi.EndEmpresa_Id, Pxi.Pedido_Id, Pxi.Produto_Id" & vbCrLf &
              "         ) as sbPed " & vbCrLf &
              "		on sbPed.Empresa_Id    = P.Empresa_Id " & vbCrLf &
              "	   and sbPed.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf &
              "	   and sbPed.Pedido_Id     = P.Pedido_Id " & vbCrLf &
              " inner Join Produtos prd" & vbCrLf &
              "    on prd.produto_id =  sbPed.Produto_id" & vbCrLf &
              " left join (" & vbCrLf &
              "                SELECT NF.Empresa_Id," & vbCrLf &
              "                       NF.EndEmpresa_Id," & vbCrLf &
              "                       NF.Pedido," & vbCrLf &
              "                       NFxI.Produto_Id," & vbCrLf &
              "                       sum(case " & vbCrLf &
              "                              when SO.Devolucao = 'N'" & vbCrLf &
              "                                then NFxI.QuantidadeFiscal" & vbCrLf &
              "                                else 0" & vbCrLf &
              "                           end) AS EntregueNota," & vbCrLf &
              "                       sum(case " & vbCrLf &
              "                              when SO.Devolucao = 'S'" & vbCrLf &
              "                                then NFxI.QuantidadeFiscal" & vbCrLf &
              "                                else 0" & vbCrLf &
              "                           end) AS DevolucaoNota" & vbCrLf &
              "                  FROM NotasFiscais NF" & vbCrLf &
              "                 INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf &
              "                    ON NFxI.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
              "                   AND NFxI.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
              "                   AND NFxI.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
              "                   AND NFxI.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
              "                   AND NFxI.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
              "                   AND NFxI.Serie_Id        = NF.Serie_Id" & vbCrLf &
              "                   AND NFxI.Nota_Id         = NF.Nota_Id" & vbCrLf &
              "                  LEFT OUTER JOIN SubOperacoes SO" & vbCrLf &
              "                    ON NF.Operacao    = SO.Operacao_Id" & vbCrLf &
              "                   AND NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
              "                 WHERE NF.Situacao        = 1" & vbCrLf &
              "                   And NF.TipoDeDocumento = 1" & vbCrLf &
              "                   And SO.Classe not in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.COMPLEMENTACOES.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "')" & vbCrLf &
              "                 group by NF.Empresa_Id," & vbCrLf &
              "                          NF.EndEmpresa_Id," & vbCrLf &
              "                          NF.Pedido," & vbCrLf &
              "                          NFxI.Produto_Id" & vbCrLf &
              "             ) sbnf" & vbCrLf &
              "    on sbnf.Empresa_Id    = P.Empresa_Id" & vbCrLf &
              "   and sbnf.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf &
              "   and sbnf.Pedido        = P.Pedido_Id " & vbCrLf &
              "   and sbnf.Produto_Id    = sbPed.Produto_Id" & vbCrLf &
              " where P.Situacao                  = 1 " & vbCrLf &
              "   and isnull(P.FiscalAberto,1)    = 1 " & vbCrLf &
              "   and isnull(P.PedidoBloqueado,0) = 0 " & vbCrLf &
              "   and P.Empresa_Id    ='" & parametros("empresa") & "'" & vbCrLf &
              "   and P.EndEmpresa_Id = " & parametros("enderecoEmpresa")


        If parametros.ContainsKey("pedidoefetivo") Then
            sql &= " and P.PedidoEfetivo = '" & parametros("pedidoefetivo") & "'"
        Else
            sql &= "   and P.Cliente       ='" & parametros("cliente") & "'" & vbCrLf &
                  "   and P.EndCliente    = " & parametros("enderecoCliente")
        End If

        '#NGSVerificar a necessidade FiscalAberto PedidoBloqueado

        If parametros.ContainsKey("dataInicio") AndAlso parametros.ContainsKey("dataFim") Then
            sql &= "   and P.DataPedido BETWEEN '" & parametros("dataInicio") & "' AND '" & parametros("dataFim") & "'  " & vbCrLf
        End If

        If CodigoSafra.Length > 0 Then
            sql &= "   and P.Safra   = '" & CodigoSafra & "'" & vbCrLf
        End If

        If AnoSafra > 0 Then
            sql &= "   and P.Safra in (Select Safra_id from safras where year(Vencimento) >= " & AnoSafra & ")" & vbCrLf
        End If

        If CodigoProduto.Length > 0 Then
            sql &= "   and sbPed.Produto_Id = '" & CodigoProduto & "'" & vbCrLf
        End If

        sql &= " group by P.Empresa_Id," & vbCrLf & _
               "          P.EndEmpresa_Id," & vbCrLf & _
               "          P.Pedido_Id," & vbCrLf & _
               "		  P.DataPedido," & vbCrLf & _
               "		  P.DataEntrega," & vbCrLf & _
               "		  P.Operacao," & vbCrLf & _
               "		  P.SubOperacao," & vbCrLf & _
               "		  P.Safra," & vbCrLf & _
               "          sbPed.Produto_Id," & vbCrLf & _
               "          prd.Nome," & vbCrLf & _
               "          isnull(sbPed.QtdNormal,0)," & vbCrLf & _
               "          isnull(sbPed.Total,0)," & vbCrLf & _
               "          isnull(sbnf.EntregueNota,0)," & vbCrLf & _
               "          isnull(sbnf.DevolucaoNota,0)" & vbCrLf

        sql &= " select C.CodigoEmpresa," & vbCrLf & _
               "        C.EnderecoEmpresa," & vbCrLf & _
               "        C.Pedido," & vbCrLf & _
               "        C.DataPedido," & vbCrLf & _
               "        C.DataEntrega," & vbCrLf & _
               "        C.Operacao," & vbCrLf & _
               "        C.SubOperacao," & vbCrLf & _
               "        C.Safra," & vbCrLf & _
               "        C.Produto, " & vbCrLf & _
               "        C.NomeDoProduto," & vbCrLf & _
               "        C.QtdNormal As Contratada," & vbCrLf & _
               "        C.Unitario," & vbCrLf & _
               "        C.EntregueNota - C.DevolucaoNota as Entregue," & vbCrLf & _
               "        case " & vbCrLf & _
               "             when so.precofixo = 'N'" & vbCrLf & _
               "            then 0" & vbCrLf & _
               "            else C.QtdNormal - (C.EntregueNota - C.DevolucaoNota)" & vbCrLf & _
               "        end as Saldo, SO.Descricao DescricaoOperacao" & vbCrLf & _
               "  from #Consulta C" & vbCrLf & _
               " Inner Join OPeracoes OP" & vbCrLf & _
               "    on op.Operacao_id = c.Operacao" & vbCrLf & _
               " Inner Join SubOperacoes SO" & vbCrLf & _
               "    ON C.Operacao    = SO.Operacao_Id" & vbCrLf & _
               "   AND C.SubOperacao = SO.SubOperacoes_Id" & vbCrLf

        If parametros.ContainsKey("ClassePedido") Then
            sql &= "   and OP.CLASSE in (" & parametros("ClassePedido") & ")" & vbCrLf
        End If

        ds = Banco.ConsultaDataSet(sql, "Pedidos")
        Session("objConsultaPedidos" & HID.Value) = ds
        GridPedidos.DataSource = ds
        GridPedidos.DataBind()
        Return ds.Tables(0).Rows.Count
    End Function

    Private Function CargaPedidosRomaneio() As Integer
        Dim ds As New DataSet
        Dim sql As String
        sql = "SELECT P.Empresa_id as CodigoEmpresa," & vbCrLf & _
              "       P.EndEmpresa_id as EnderecoEmpresa," & vbCrLf & _
              "       P.Pedido_Id AS Pedido," & vbCrLf & _
              "       P.DataPedido," & vbCrLf & _
              "       P.DataEntrega," & vbCrLf & _
              "       P.Operacao," & vbCrLf & _
              "       P.SubOperacao," & vbCrLf & _
              "       ISNULL(Pxi.Produto_Id, '') AS Produto," & vbCrLf & _
              "       ISNULL(prd.Nome, '') AS NomeDoProduto," & vbCrLf & _
              "       ISNULL(lan.Quantidade, 0) AS Contratada, " & vbCrLf & _
              "       Isnull(nf.Entregue,0) as Entregue," & vbCrLf & _
              "       ISNULL(lan.Quantidade, 0) - Isnull(nf.Entregue,0) as Saldo" & vbCrLf & _
              "  FROM Pedidos P  " & vbCrLf & _
              " Inner join PedidoxItem pxi" & vbCrLf & _
              "    on P.Empresa_id    = pxi.Empresa_id" & vbCrLf & _
              "   and P.EndEmpresa_id = pxi.EndEmpresa_id" & vbCrLf & _
              "   and P.Pedido_id     = pxi.Pedido_id" & vbCrLf & _
              " Inner JOIN (Select Empresa_id, EndEmpresa_id, Pedido_id, Produto_Id," & vbCrLf & _
              "                    Sum(Case" & vbCrLf & _
              "                           When TipoDeLancamento = 'E'" & vbCrLf & _
              "                             then Quantidade * -1" & vbCrLf & _
              "                             Else Quantidade" & vbCrLf & _
              "                         end) as Quantidade" & vbCrLf & _
              "               from PedidoXItemXLancamento" & vbCrLf & _
              "              group by Empresa_id, endEmpresa_id, Pedido_id, Produto_Id" & vbCrLf & _
              "             ) Lan" & vbCrLf & _
              "    ON P.Empresa_Id    = pxi.Empresa_Id" & vbCrLf & _
              "   AND P.EndEmpresa_Id = pxi.EndEmpresa_Id" & vbCrLf & _
              "   AND P.Pedido_Id     = pxi.Pedido_Id" & vbCrLf & _
              "   And P.Produto_Id    = pxi.Produto_id" & vbCrLf & _
              " Inner Join Produtos prd" & vbCrLf & _
              "    ON prd.produto_id = pxi.produto_id" & vbCrLf & _
              "  Left Join (" & vbCrLf & _
              "             SELECT NF.Empresa_Id, NF.EndEmpresa_Id,NF.Pedido, NFI.Produto_Id," & vbCrLf & _
              "                    SUM(Case" & vbCrLf & _
              "                          When SO.Devolucao = 'S'" & vbCrLf & _
              "                            then nfi.PesoFiscal * -1" & vbCrLf & _
              "                            else nfi.PesoFiscal" & vbCrLf & _
              "                        end) AS Entregue " & vbCrLf & _
              "               FROM NotasFiscais NF" & vbCrLf & _
              "              INNER JOIN NotasFiscaisXItens NFI" & vbCrLf & _
              "                 ON NF.Empresa_Id      = NFI.Empresa_Id" & vbCrLf & _
              "                AND NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
              "                AND NF.Cliente_Id      = NFI.Cliente_Id " & vbCrLf & _
              "                AND NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
              "                AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf & _
              "                AND NF.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
              "                AND NF.Nota_Id         = NFI.Nota_Id" & vbCrLf & _
              "              INNER JOIN SubOperacoes SO" & vbCrLf & _
              "                 ON NFI.Operacao    = SO.Operacao_Id" & vbCrLf & _
              "                AND NFI.SubOperacao = SO.SubOperacoes_Id " & vbCrLf & _
              "              Where NotasFiscais.Situacao = 1" & vbCrLf & _
              "              group By NF.Empresa_Id, NF.EndEmpresa_Id,NF.Pedido, NFI.Produto_Id" & vbCrLf & _
              "            ) NF" & vbCrLf & _
              "    ON NF.Empresa_Id    = pxi.Empresa_Id" & vbCrLf & _
              "   AND NF.EndEmpresa_Id = pxi.EndEmpresa_Id" & vbCrLf & _
              "   AND NF.Pedido        = Pxi.Pedido_id" & vbCrLf & _
              "   AND NF.Produto_Id    = Pxi.Produto_Id" & vbCrLf & _
              " WHERE P.Empresa_Id    ='" & HttpContext.Current.Session("ssCnpjDaEmpresa") & "'" & vbCrLf & _
              "   AND P.EndEmpresa_Id = " & HttpContext.Current.Session("ssEndDaEmpresa") & vbCrLf & _
              "   AND P.Cliente       ='" & HttpContext.Current.Session("txtCnpjDoCliente") & "'" & vbCrLf & _
              "   AND P.EndCliente    = " & HttpContext.Current.Session("txtEndDoCliente") & vbCrLf & _
              "   AND P.Situacao      = 1" & vbCrLf
        '#NGSVerificar

        ds = Banco.ConsultaDataSet(sql, "Pedidos")
        Session("objConsultaPedidos" & HID.Value) = ds
        GridPedidos.DataSource = ds
        GridPedidos.DataBind()
        Return ds.Tables(0).Rows.Count
    End Function

    Protected Sub GridPedidos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridPedidos2.SelectedIndexChanged, GridPedidos.SelectedIndexChanged
        Dim grd As GridView = CType(sender, GridView)
        Selecionar(grd.SelectedRow.Cells(1).Text & ";" & grd.SelectedRow.Cells(2).Text & ";" & grd.SelectedRow.Cells(3).Text)
    End Sub

    Private Function ListarPedidos(ByVal parametros As Dictionary(Of String, Object)) As Integer
        Try
            Dim numPedido As Integer = 0
            If parametros IsNot Nothing AndAlso parametros.ContainsKey("pedido") Then
                numPedido = Convert.ToInt32(parametros("pedido"))
            End If

            Dim strEfetivo As String = String.Empty
            If parametros IsNot Nothing AndAlso parametros.ContainsKey("efetivo") Then
                strEfetivo = parametros("efetivo").ToString
            End If

            Dim strUnidade As String = String.Empty
            If parametros IsNot Nothing AndAlso parametros.ContainsKey("unidade") Then
                strUnidade = parametros("unidade").ToString
            End If

            Dim strEmpresa As String = String.Empty
            Dim intEndEmpresa As Integer = 0

            If parametros IsNot Nothing AndAlso parametros.ContainsKey("empresa") Then
                strEmpresa = parametros("empresa").ToString
            End If

            If parametros IsNot Nothing AndAlso parametros.ContainsKey("enderecoEmpresa") Then
                intEndEmpresa = Convert.ToInt32(parametros("enderecoEmpresa"))
            End If

            Dim strCliente As String = String.Empty
            Dim intEndCliente As Integer = 0

            If parametros IsNot Nothing AndAlso parametros.ContainsKey("cliente") Then
                strCliente = parametros("cliente").ToString
            End If

            If parametros IsNot Nothing AndAlso parametros.ContainsKey("enderecoCliente") Then
                intEndCliente = Convert.ToInt32(parametros("enderecoCliente"))
            End If

            Dim objSituacao As [Lib].Negocio.eSituacao = Nothing
            If parametros IsNot Nothing AndAlso parametros.ContainsKey("situacao") Then
                objSituacao = [Enum].Parse(GetType([Lib].Negocio.eSituacao), parametros("situacao"))
            End If

            Dim strSafra As String = String.Empty
            If parametros IsNot Nothing AndAlso parametros.ContainsKey("safra") Then
                strSafra = parametros("safra").ToString
            End If

            Dim AnoSafra As Integer = 0
            If parametros IsNot Nothing AndAlso parametros.ContainsKey("Ano") Then
                AnoSafra = Convert.ToInt32(parametros("Ano"))
            End If

            Dim AnoFinalSafra As Integer = 0
            If parametros IsNot Nothing AndAlso parametros.ContainsKey("AnoFinalSafra") Then
                AnoFinalSafra = Convert.ToInt32(parametros("AnoFinalSafra"))
            End If

            Dim intOperacao As Integer = 0
            Dim intSubOperacao As Integer = 0

            If parametros IsNot Nothing AndAlso parametros.ContainsKey("operacao") Then
                intOperacao = Convert.ToInt32(parametros("operacao"))
            End If

            If parametros IsNot Nothing AndAlso parametros.ContainsKey("suboperacao") Then
                intSubOperacao = Convert.ToInt32(parametros("suboperacao"))
            End If

            Dim grupoProdutos As String = String.Empty
            If parametros IsNot Nothing AndAlso parametros.ContainsKey("grupoProduto") Then
                grupoProdutos = parametros("grupoProduto").ToString
            End If

            Dim inProduto As String = String.Empty
            If parametros IsNot Nothing AndAlso parametros.ContainsKey("produto") Then
                inProduto = IIf(Not parametros("produto").ToString.Contains("'"), "'" & parametros("produto") & "'", parametros("produto"))
            End If

            Dim entSai As [Lib].Negocio.eEntradaSaidaNenhum = eEntradaSaidaNenhum.Nenhum

            If parametros IsNot Nothing AndAlso parametros.ContainsKey("es") Then
                If parametros("es").ToString = "E" Then
                    entSai = eEntradaSaidaNenhum.Entrada
                ElseIf parametros("es").ToString = "S" Then
                    entSai = eEntradaSaidaNenhum.Saida
                End If
            End If

            Dim sql As String
            sql = "SELECT P.Empresa_Id     as CodigoEmpresa," & vbCrLf &
                  "       P.EndEmpresa_Id  as EnderecoEmpresa," & vbCrLf &
                  "       P.Pedido_Id      as Pedido, " & vbCrLf &
                  "       P.PedidoEfetivo  as PedidoEfetivo, " & vbCrLf &
                  "       P.Operacao       as CodigoOperacao," & vbCrLf &
                  "       P.SubOperacao    as CodigoSubOperacao," & vbCrLf &
                  "       SO.Descricao     as DescricaoSubOperacao," & vbCrLf &
                  "       CAST(P.Operacao as varchar) + '-' + CAST(P.SubOperacao as varchar) + ' - ' + SO.Descricao as SubOp," & vbCrLf &
                  "       C.Nome           as NomeCliente," & vbCrLf &
                  "       P.DataPedido," & vbCrLf &
                  "       M.Descricao      as Moeda," & vbCrLf &
                  "       PxI.Saldo        as Quantidade," & vbCrLf &
                  "       CASE WHEN M.Classificacao = 'O'" & vbCrLf &
                  "              THEN isnull(Liq.ValorOficial,0)" & vbCrLf &
                  "              ELSE isnull(Liq.ValorMoeda,0)" & vbCrLf &
                  "       END as Total," & vbCrLf &
                  "       ISNULL(P.Embalagem, 0) AS Embalagem," & vbCrLf &
                  "       ISNULL(P.TipoCondicaoEntrega, '') AS TipoCondicaoEntrega," & vbCrLf &
                  "       ISNULL(P.TipoPagamentoPtax, '') AS TipoPagamentoPtax, ISNULL(P.LocalDeEmbarque, '') AS LocalDeEmbarque, " & vbCrLf &
                  "       [dbo].ProdutosPedido(P.Empresa_Id, P.EndEmpresa_Id, P.Pedido_Id) as Itens" & vbCrLf &
                  "  FROM Pedidos P " & vbCrLf &
                  " INNER JOIN (" & vbCrLf &
                  "             SELECT PEDxI.Empresa_Id,  " & vbCrLf &
                  "                    PEDxI.EndEmpresa_Id,  " & vbCrLf &
                  "                    PEDxI.Pedido_Id,  " & vbCrLf &
                  "                    SUM(CASE" & vbCrLf &
                  "                          WHEN PEDxI.TipoDeLancamento = 'E'" & vbCrLf &
                  "                             THEN PEDxI.Quantidade * - 1" & vbCrLf &
                  "                             ELSE PEDxI.Quantidade" & vbCrLf &
                  "                          END) as Saldo" & vbCrLf &
                  "               FROM PedidoXItemxLancamento PEDxI" & vbCrLf &
                  "               WHERE PEDxI.Empresa_Id    = '" & strEmpresa & "'" & vbCrLf &
                  "                 AND PEDxI.EndEmpresa_Id = " & intEndEmpresa.ToString() & vbCrLf

            sql &= "         GROUP BY PEDxI.Empresa_Id, PEDxI.EndEmpresa_Id, PEDxI.Pedido_Id" & vbCrLf & _
                   "        ) as PxI  " & vbCrLf & _
                   "    ON PxI.Empresa_Id = P.Empresa_Id" & vbCrLf & _
                   "   AND PxI.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
                   "   AND PxI.Pedido_Id = P.Pedido_Id" & vbCrLf & _
                   " INNER JOIN Clientes C" & vbCrLf & _
                   "    ON P.Cliente    = C.cliente_id" & vbCrLf & _
                   "   AND P.EndCliente = C.Endereco_id" & vbCrLf & _
                   " INNER JOIN Moedas M" & vbCrLf & _
                   "    ON P.moeda = M.Moeda_id" & vbCrLf & _
                   "  LEFT JOIN (SELECT Empresa_id," & vbCrLf & _
                   "                    EndEmpresa_id," & vbCrLf & _
                   "                    Pedido_id," & vbCrLf & _
                   "                    Sum(ValorOficial) as ValorOficial," & vbCrLf & _
                   "                    Sum(ValorMoeda) as ValorMoeda" & vbCrLf & _
                   "               FROM PedidosXEncargos" & vbCrLf & _
                   "              WHERE Encargo_id = 'LIQUIDO'" & vbCrLf & _
                   "                 AND Empresa_Id = '" & strEmpresa & "'" & vbCrLf &
                   "                 AND EndEmpresa_Id = " & intEndEmpresa.ToString() & vbCrLf & _
                   "              GROUP BY Empresa_id, EndEmpresa_id, Pedido_Id" & vbCrLf & _
                   "             ) as Liq" & vbCrLf & _
                   "    ON P.Empresa_Id    = liq.Empresa_Id" & vbCrLf & _
                   "   AND P.EndEmpresa_Id = liq.EndEmpresa_Id" & vbCrLf & _
                   "   AND P.Pedido_Id     = liq.Pedido_Id" & vbCrLf

            sql &= " INNER JOIN SubOperacoes SO " & vbCrLf & _
                      "    ON SO.Operacao_Id = P.Operacao " & vbCrLf & _
                      "   AND SO.SubOperacoes_Id = P.SubOperacao " & vbCrLf

            sql &= "WHERE P.UnidadeDeNegocio IS NOT NULL " & vbCrLf & _
                      "--AND P.PedidoOrigem = 0 --OQ?" & vbCrLf

            If strUnidade <> "" Then sql &= "AND P.UnidadeDeNegocio = '" & strUnidade & "' " & vbCrLf
            If numPedido > 0 Then sql &= "AND P.Pedido_Id = " & numPedido.ToString() & vbCrLf
            If strEfetivo <> "" Then sql &= "AND P.PedidoEfetivo = '" & strEfetivo & "' " & vbCrLf

            If strEmpresa <> "" Then
                sql &= "AND P.Empresa_Id = '" & strEmpresa & "' " & vbCrLf & _
                          "AND P.EndEmpresa_Id = " & intEndEmpresa.ToString() & vbCrLf
            End If

            If strCliente <> "" Then
                sql &= "AND P.Cliente = '" & strCliente & "' " & vbCrLf & _
                          "AND P.EndCliente = " & intEndCliente.ToString() & vbCrLf
            End If

            If AnoSafra > 0 Then
                sql &= "   and P.Safra in (Select Safra_id from safras where year(Vencimento) >= " & AnoSafra & ")" & vbCrLf
            End If

            If AnoFinalSafra > 0 Then
                sql &= "   and P.Safra in (Select Safra_id from safras where year(Vencimento) <= " & AnoFinalSafra & ")" & vbCrLf
            End If

            If objSituacao <> Nothing Then sql &= "AND P.Situacao = " & Convert.ToInt32(objSituacao).ToString() & vbCrLf
            If strSafra <> "" Then sql &= "AND P.Safra = '" & strSafra & "' " & vbCrLf

            If intOperacao > 0 Then
                sql &= "AND P.Operacao = " & intOperacao.ToString() & vbCrLf
                If intSubOperacao > 0 Then sql &= "AND P.SubOperacao = " & intSubOperacao.ToString() & vbCrLf
            End If

            If entSai <> eEntradaSaidaNenhum.Nenhum Then
                sql &= "AND (SO.EntradaSaida = '" & entSai.ToString().Substring(0, 1) & "' OR SO.Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "') " & vbCrLf
            End If

            If grupoProdutos.Length > 0 Then
                sql &= "    AND EXISTS (SELECT 1" & vbCrLf & _
                          "                  FROM PedidoXItem PIT" & vbCrLf & _
                          "                 WHERE P.Empresa_Id    = PIT.Empresa_Id" & vbCrLf & _
                          "                   AND P.EndEmpresa_Id = PIT.EndEmpresa_Id" & vbCrLf & _
                          "                   AND P.Pedido_Id     = PIT.Pedido_Id" & vbCrLf & _
                          "                   AND left(PIT.Produto_Id,5)  in (" & grupoProdutos & "))" & vbCrLf
            End If

            If inProduto.Length > 0 Then
                sql &= "    AND EXISTS (SELECT 1" & vbCrLf & _
                          "                  FROM PedidoXItem PIT" & vbCrLf & _
                          "                 WHERE P.Empresa_Id    = PIT.Empresa_Id" & vbCrLf & _
                          "                   AND P.EndEmpresa_Id = PIT.EndEmpresa_Id" & vbCrLf & _
                          "                   AND P.Pedido_Id     = PIT.Pedido_Id" & vbCrLf & _
                          "                   AND PIT.Produto_Id  in (" & inProduto & "))" & vbCrLf
            End If

            If parametros.ContainsKey("ContaContabil") Then
                sql &= " AND (SELECT ISNULL(SUM(TC.ValorOficial),0) FROM Titulos T " & vbCrLf & _
                            "       INNER JOIN TitulosxContaContabil TC " & vbCrLf & _
                            "       ON T.Titulo_Id = TC.Titulo_Id " & vbCrLf & _
                            "       WHERE TC.Conta_Id = " & parametros("ContaContabil") & vbCrLf & _
                            "       AND T.Pedido = P.Pedido_Id) " & vbCrLf & _
                            " > " & vbCrLf & _
                            " (SELECT ISNULL(SUM(AxB.ValorOficial),0) FROM Titulos T " & vbCrLf & _
                            "       INNER JOIN vw_AdiantamentosXBaixas AxB " & vbCrLf & _
                            "       ON T.Titulo_Id = AxB.Titulo_Id " & vbCrLf & _
                            "       WHERE T.ContaContabilCliFor = " & parametros("ContaContabil") & vbCrLf & _
                            "       AND T.Pedido = P.Pedido_Id) " & vbCrLf
            End If

            sql &= " ORDER BY P.Pedido_Id " & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Pedidos")

            Select Case ds.Tables(0).Rows.Count
                Case 0
                    Session.Remove("objPedidoSelecionado" & HID.Value)
                    Session.Remove("objConsultaPedidos" & HID.Value)
                    MsgBox(Me.Page, "Não foi encontrado nenhum pedido com os parâmetros informados!")
                    Popup.CloseDialog(Me.Page, "divConsultaPedidos", True, 500)
                Case 1
                    Dim objPedido As [Lib].Negocio.Pedido = New [Lib].Negocio.Pedido(ds.Tables(0).Rows(0)("CodigoEmpresa"), ds.Tables(0).Rows(0)("EnderecoEmpresa"), ds.Tables(0).Rows(0)("Pedido"))
                    Selecionar(objPedido)
                Case Else
                    Session("objConsultaPedidos" & HID.Value) = ds
                    GridPedidos.Visible = False
                    gridPedidos2.Visible = True
                    gridPedidos2.DataSource = ds
                    gridPedidos2.DataBind()
            End Select
            Return ds.Tables(0).Rows.Count
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Function

    Protected Sub gridPedidos2_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gridPedidos2.PageIndex = e.NewPageIndex
        gridPedidos2.DataSource = CType(Session("objConsultaPedidos" & HID.Value), DataSet)
        gridPedidos2.DataBind()
    End Sub

    Private Function CargaPedidosLaudo(ByVal parametros As Dictionary(Of String, Object)) As Integer
        Dim Pesagem As New [Lib].Negocio.Pesagem
        Pesagem.CodigoEmpresa = parametros("CodigoEmpresa")
        Pesagem.EnderecoCliente = parametros("CodigoEndEmpresa")
        Pesagem.CodigoCliente = parametros("CodigoCliente")
        Pesagem.EnderecoCliente = parametros("CodigoEndCliente")
        Pesagem.Consolidado = parametros("Consolidado")
        Dim AnoCorrente = parametros("AnoCorrente")
        Dim ds As DataSet = Pesagem.SaldoDePedidos(Pesagem, AnoCorrente, IIf(ddlSafra.SelectedIndex > 0, ddlSafra.SelectedValue, ""))
        Session("objConsultaPedidos" & HID.Value) = ds
        GridPedidos.DataSource = ds
        GridPedidos.DataBind()
        If (GridPedidos.Rows.Count > 0) Then
            GridPedidos.Rows(0).Cells(0).Focus()
        End If
        Return ds.Tables(0).Rows.Count
    End Function

    Protected Overrides Sub Selecionar(ByVal args As String)
        Try
            Dim strPedido() As String = args.Split(";")
            Dim objPedido As New [Lib].Negocio.Pedido(strPedido(0), strPedido(1), strPedido(2))
            Session(Session("ssTipoRetorno")) = objPedido
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(objPedido)
                Else
                    CType(Me.Page, IBasePage).Carregar(objPedido)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaPedidos")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Overrides Sub Selecionar(ByVal obj As IBaseEntity)
        Try
            Session(Session("ssTipoRetorno")) = obj
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(obj)
                Else
                    CType(Me.Page, IBasePage).Carregar(obj)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaPedidos")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtConsultaPedido_TextChanged(sender As Object, e As EventArgs) Handles txtConsultaPedido.TextChanged
        Dim ds As DataSet = CType(Session("objConsultaPedidos" & HID.Value), DataSet)
        Dim dv As New DataView(ds.Tables(0))
        If Not String.IsNullOrWhiteSpace(txtConsultaPedido.Text) AndAlso IsNumeric(txtConsultaPedido.Text) Then
            dv.RowFilter = "Pedido = " & txtConsultaPedido.Text.Trim()
        End If

        If GridPedidos.Visible Then
            GridPedidos.DataSource = dv
            GridPedidos.DataBind()
        Else
            gridPedidos2.DataSource = dv
            gridPedidos2.DataBind()
        End If
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Popup.CloseDialog(Me.Page, "divConsultaPedidos")
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Dim ds As DataSet = CType(Session("objConsultaPedidos" & HID.Value), DataSet)
        Dim dv As New DataView(ds.Tables(0))
        If GridPedidos.Visible Then
            GridPedidos.DataSource = dv
            GridPedidos.DataBind()
        Else
            gridPedidos2.DataSource = dv
            gridPedidos2.DataBind()
        End If
        txtConsultaPedido.Text = String.Empty
    End Sub

    Protected Sub ddlSafra_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlSafra.SelectedIndexChanged
        Select Case Session("ssCampo" & HID.Value)
            Case "Romaneio"
                CargaPedidosRomaneio()
            Case "Pedidos"
                ListarPedidos(MyParameters)
            Case "Laudo"
                CargaPedidosLaudo(MyParameters)
            Case "Carregamento"
                CargaPedidosCarregamento(MyParameters)
            Case Else
                CargaPedidos(MyParameters)
        End Select
    End Sub

End Class