Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class RelacaoDePedidos
    Inherits BasePage

    Dim SqlArray As New ArrayList
    Dim ds As DataSet
    Dim dr As DataRow

    Private Property Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelacaoDePedidos", "ACESSAR") Then
                    HttpContext.Current.Session("ssRegistros") = ""
                    HttpContext.Current.Session("ssObservacoes") = ""

                    txtPeriodoInicial.Text = Format(Today, "dd/MM/yyyy")
                    txtPeriodoFinal.Text = Format(Today, "dd/MM/yyyy")

                    CargaUnidadeDeNegocioEmpresaCliente()
                    CarregaClasse()
                    BuscarOperacoes()
                    CarregarSafras()
                    BuscarMoedas()
                    ddl.Carregar(ddlMarca, CarregarDDL.Tabela.Marca, "")
                    ddl.Carregar(ddlRepresentante, CarregarDDL.Tabela.ClientesXTipos, "6", True)
                    ddl.Carregar(lstFinalidade, CarregarDDL.Tabela.Finalidade, "", False)

                    HID.Value = Guid.NewGuid().ToString()
                    ucConsultaClientes.SetarHID(HID.Value)

                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeConsultaTitulos.Enabled = False
            DdlEmpresaConsultaTitulos.Enabled = False
        End If
    End Sub

    Private Sub CargaUnidadeDeNegocioEmpresaCliente()
        ddl.Carregar(DdlUnidadeConsultaTitulos, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeConsultaTitulos, DdlEmpresaConsultaTitulos)
    End Sub

    Protected Sub DdlUnidadeConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(DdlEmpresaConsultaTitulos, CarregarDDL.Tabela.Empresas, DdlUnidadeConsultaTitulos.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarSafras()
        ddl.Carregar(cmbSafra, CarregarDDL.Tabela.Safra, "", True)
    End Sub

    Private Sub BuscarMoedas()
        Dim objMoedas As New [Lib].Negocio.Moedas()

        cmbMoeda.Items.Clear()
        Funcoes.InserirLinhaEmBranco(cmbMoeda, 0)
        For Each objMoeda As [Lib].Negocio.Moeda In objMoedas
            cmbMoeda.Items.Add(New ListItem(objMoeda.Codigo.ToString() & "-" & objMoeda.Descricao, objMoeda.Codigo.ToString()))
        Next
    End Sub

    Protected Sub cmdAjuda_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim NomeArquivo As String = "Manual/RelacaoDePedidos.mht"
            Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)

            ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmbOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            BuscarSubOperacoes(cmbOperacao.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregaClasse()
        Dim sql As String = " select Classe_Id from ClassesDeOperacoes Where Operacao = 1"
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Classes")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            lstClasses.Items.Clear()
            For Each row As DataRow In ds.Tables(0).Rows
                lstClasses.Items.Add(New ListItem(row("Classe_Id"), row("Classe_Id")))
            Next
        End If
    End Sub

    Private Sub BuscarOperacoes()
        ddl.Carregar(cmbOperacao, CarregarDDL.Tabela.Operacao, "", True)
    End Sub

    Private Sub BuscarSubOperacoes(ByVal subOperacao As String)
        ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & subOperacao, True)
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objClienteRP" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteRP" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value.Replace("-", ";")
            Session.Remove("objClienteRP" & HID.Value)
        End If
    End Sub

    Protected Sub cmdConsultaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteRP" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function SqlBase(Optional ByRef Parametros As String = "") As String
        Dim Sql As String = ""
        Dim selecionaClasse As Boolean = False
        Dim selecionaGrupo As Boolean = False

        Sql = "SELECT P.Empresa_id," & vbCrLf &
              "       P.EndEmpresa_id ," & vbCrLf &
              "       E.Nome AS EmpresaNome," & vbCrLf &
              "       E.Reduzido AS EmpresaReduzido," & vbCrLf &
              "       E.Cidade AS EmpresaCidade," & vbCrLf &
              "       P.Pedido_Id AS Pedido, " & vbCrLf &
              "       isnull(P.PedidoEfetivo,'') as PedidoEfetivo," & vbCrLf &
              "       P.Moeda, " & vbCrLf &
              "       M.Classificacao," & vbCrLf &
              "       M.Cifrao," & vbCrLf &
              "       isnull(Pg.Descricao,'') as CondicaoPagamento," & vbCrLf &
              "       Clientes.Cliente_Id as Cliente," & vbCrLf &
              "       Clientes.Nome AS NomeCliente," & vbCrLf &
              "       CO.Representante_Id as Representante," & vbCrLf &
              "       CLI.Nome as NomeRepresentante," & vbCrLf &
              "       CASE" & vbCrLf &
              "         WHEN SUBSTRING(isnull(Clientes.Complemento, ''), 1, 1) = '' " & vbCrLf &
              "           THEN (Clientes.Cidade + '/' + Clientes.Estado)" & vbCrLf &
              "           ELSE (Clientes.Complemento + ' - ' + Clientes.Cidade + '/' + Clientes.Estado) " & vbCrLf &
              "       END AS Complemento, " & vbCrLf &
              "       Clientes.Cidade as Cidade," & vbCrLf &
              "       Clientes.Estado as Estado," & vbCrLf &
              "       CONVERT(varchar, P.DataPedido, 103) AS DataPedido," & vbCrLf &
              "       P.DataEntrega," & vbCrLf &
              "       PxI.Produto_Id as CodigoProduto," & vbCrLf &
              "       PxI.Nome AS NomeProduto, " & vbCrLf &
              "       PxI.Embalagem, " & vbCrLf &
              "       REPLICATE('0', 2 - LEN(CAST(P.Operacao AS varchar))) + CAST(P.Operacao AS varchar) AS Operacao, " & vbCrLf &
              "       REPLICATE('0', 2 - LEN(CAST(P.SubOperacao AS varchar))) + CAST(P.SubOperacao AS varchar) AS SubOperacao, " & vbCrLf &
              "       SO.Descricao as DescOperacao," & vbCrLf &
              "       SO.Classe," & vbCrLf &
              "       case" & vbCrLf &
              "          When isnull(PxI.Quantidade,0) > 0" & vbCrLf &
              "            then (PxI.TotalOficial / PxI.Quantidade)" & vbCrLf &
              "            else 0" & vbCrLf &
              "       end as UnitarioOficial," & vbCrLf &
              "       case" & vbCrLf &
              "          When isnull(PxI.Quantidade,0) > 0" & vbCrLf &
              "            then (PxI.TotalMoeda / PxI.Quantidade)" & vbCrLf &
              "            else 0" & vbCrLf &
              "       end as UnitarioMoeda," & vbCrLf &
              "       case" & vbCrLf &
              "          When isnull(PxI.PesoDoItem,0) > 0" & vbCrLf &
              "            then PxI.Quantidade * PxI.PesoDoItem" & vbCrLf &
              "            else PxI.Quantidade" & vbCrLf &
              "       end as QuantidadeItemPedido," & vbCrLf &
              "       Isnull(PxI.Quantidade,0) as QuantidadeSoma," & vbCrLf &
              "       Isnull(PxI.Quantidade,0) as Quantidade," & vbCrLf &
              "       PxI.PesoDoItem," & vbCrLf &
              "       case when isnull(pxe.ValorOficial, 0) > 0 AND isnull(PxI.Quantidade,0) > 0 " & vbCrLf &
              "	        then pxe.ValorOficial / pxi.Quantidade        " & vbCrLf &
              "		    else 0                                        " & vbCrLf &
              "	      end as UnitarioLiquido,                         " & vbCrLf &
              "	      case when isnull(pxe.ValorMoeda, 0) > 0 AND isnull(PxI.Quantidade,0) > 0   " & vbCrLf &
              "	        then pxe.ValorMoeda / pxi.Quantidade          " & vbCrLf &
              "		    else 0                                        " & vbCrLf &
              "	      end as UnitarioLiquidoMoeda,                    " & vbCrLf &
              "	      isnull(pxe.ValorOficial, 0) as ValorLiquido,    " & vbCrLf &
              "	      isnull(pxe.ValorMoeda, 0) as ValorLiquidoMoeda, " & vbCrLf &
              "       PxI.TotalOficial," & vbCrLf &
              "       PxI.TotalMoeda, " & vbCrLf &
              "       isnull(FNE.EntregueGlobal,0) - isnull(FNE.EntregueRemessa,0) as FaturadoNaoEntregue," & vbCrLf &
              "       ISNULL(NFI.EntregueNota, 0) AS EntregueNota," & vbCrLf &
              "       ISNULL(NFI.EntregueNotaValor, 0) AS EntregueNotaValor, " & vbCrLf &
              "       ISNULL(NFI.DevolucaoNota, 0) AS DevolucaoNota, " & vbCrLf &
              "       ISNULL(NFI.DevolucaoNotaValor, 0) AS DevolucaoNotaValor," & vbCrLf &
              "       isnull(P.FreteCIFFOB,'FOB') as FreteCIFFOB," & vbCrLf &
              "       SO.EntradaSaida, Pg.AVista, isnull(SO.Financeiro,'N') as Financeiro, isnull(p.troca,0) Troca, " & vbCrLf &
              "       P.Observacoes, " & vbCrLf &
              "       IIF(P.PedidoBloqueado = 0, 'LIBERADO', 'BLOQUEADO') AS PedidoBloqueado " & vbCrLf &
              "  INTO [#SaldoPedido]" & vbCrLf &
              "  FROM Pedidos AS P WITH(NOLOCK)" & vbCrLf &
              " INNER JOIN SubOperacoes AS SO " & vbCrLf &
              "    ON SO.Operacao_Id     = P.Operacao " & vbCrLf &
              "   AND SO.SubOperacoes_Id = P.SubOperacao " & vbCrLf &
              " INNER JOIN Moedas M" & vbCrLf &
              "    ON P.Moeda = M.Moeda_id" & vbCrLf &
              "  Left Join Pagamentos Pg " & vbCrLf &
              "    on Pg.Pagamento_Id = P.CondicaoPagamento" & vbCrLf &
              " INNER JOIN Clientes " & vbCrLf &
              "    ON P.Cliente = Clientes.Cliente_Id " & vbCrLf &
              "   AND P.EndCliente = Clientes.Endereco_Id" & vbCrLf &
               " INNER JOIN Clientes AS E" & vbCrLf &
              "    ON E.Cliente_Id  = P.Empresa_Id" & vbCrLf &
              "   AND E.Endereco_Id = P.EndEmpresa_Id" & vbCrLf &
              " LEFT JOIN Comissoes AS CO" & vbCrLf &
              "    On CO.Empresa_Id    = P.Empresa_Id" & vbCrLf &
              "   And CO.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf &
              "   And CO.Pedido_Id = P.Pedido_Id" & vbCrLf &
              "   And CO.Principal = 'S'" & vbCrLf &
              " LEFT JOIN Clientes AS CLI" & vbCrLf &
              "    On CLI.Cliente_Id    = CO.Representante_Id" & vbCrLf &
              "   And CLI.Endereco_Id = CO.EndRepresentante_Id" & vbCrLf &
              "  LEFT JOIN (SELECT pxi.Empresa_Id, pxi.EndEmpresa_Id, pxi.Pedido_Id, Pr.Grupo, Pr.Produto_Id, Pr.CodigoProdutoTerceiro, Pr.Nome, Pr.Embalagem, isnull(pr.Marca,0) as Marca, " & vbCrLf &
              "					 CASE " & vbCrLf &
              "							WHEN left(pxi.Empresa_Id,8) = '24450490' and pxi_item.UnidadeComercializacao = 'TON'" & vbCrLf &
              "							  THEN 1 " & vbCrLf &
              "							  ELSE pxi_item.FatorConversao " & vbCrLf &
              "					 END AS PesoDoItem," & vbCrLf &
              "					 SUM(CASE " & vbCrLf &
              "							WHEN pxi.TipoDeLancamento = 'E'" & vbCrLf &
              "							  THEN pxi.Quantidade * - 1 " & vbCrLf &
              "							  ELSE pxi.Quantidade " & vbCrLf &
              "						   END) AS Quantidade," & vbCrLf &
              "					 SUM(case" & vbCrLf &
              "                        When pxi.TipoDeLancamento = 'E'" & vbCrLf &
              "							 THEN pxi.TotalOficial * -1" & vbCrLf &
              "                          ELSE pxi.TotalOficial" & vbCrLf &
              "                      end) AS TotalOficial, " & vbCrLf &
              "					  SUM(case" & vbCrLf &
              "                         When pxi.TipoDeLancamento = 'E'" & vbCrLf &
              "							  THEN pxi.TotalMoeda * -1" & vbCrLf &
              "                           ELSE pxi.TotalMoeda" & vbCrLf &
              "                       end) AS TotalMoeda" & vbCrLf &
              "			    FROM PedidoXItemxLancamento pxi " & vbCrLf &
              "		        INNER JOIN Produtos AS PR" & vbCrLf &
              "			        ON PR.Produto_Id                            = pxi.Produto_Id " & vbCrLf &
              "             INNER JOIN PedidoXItem pxi_item " & vbCrLf &
              "                 ON pxi.empresa_id                           = pxi_item.empresa_id " & vbCrLf &
              "                 AND pxi.endempresa_id                       = pxi_item.endempresa_id " & vbCrLf &
              "                 AND pxi.pedido_id                           = pxi_item.pedido_id " & vbCrLf &
              "                 AND Pr.Produto_Id                           = pxi_item.Produto_Id " & vbCrLf &
              "             INNER JOIN ProdutosxUnidadeDeComercializacao PXU  " & vbCrLf &
              "                 ON PR.Produto_Id                            = PXU.Produto_Id           " & vbCrLf &
              "                 AND pxi_item.UnidadeComercializacao         = PXU.Unidade_id  " & vbCrLf &
              "                 AND pxi_item.FatorConversao                 = PXU.FatorConversao_id  " & vbCrLf &
              "			   GROUP BY pxi.Empresa_Id, pxi.EndEmpresa_Id, pxi.Pedido_Id, Pr.Grupo, Pr.Produto_Id, Pr.CodigoProdutoTerceiro, Pr.Nome, Pr.Embalagem, isnull(pr.Marca,0), pxi_item.UnidadeComercializacao, pxi_item.FatorConversao " & vbCrLf &
              "           ) As PxI " & vbCrLf &
              "      On PxI.Empresa_Id    = P.Empresa_Id " & vbCrLf &
              "     And PxI.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf &
              "     And PxI.Pedido_Id     = P.Pedido_Id " & vbCrLf &
              "    Left Join (Select Empresa_id, EndEmpresa_id, Pedido_id, Produto_Id, sum(ValorOficial) As ValorOficial, sum(ValorMoeda) As ValorMoeda" & vbCrLf &
              "	                from pedidosxencargos pxe" & vbCrLf &
              "                where encargo_id = 'LIQUIDO'" & vbCrLf &
              "			       group by Empresa_id, EndEmpresa_id, Pedido_id, Produto_Id" & vbCrLf &
              "	        )  as pxe" & vbCrLf &
              "	     on pxe.Empresa_id    = PxI.Empresa_Id" & vbCrLf &
              "	    and pxe.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf &
              "	    and pxe.Pedido_Id     = PxI.Pedido_Id" & vbCrLf &
              "	    and pxe.Produto_Id    = PxI.Produto_Id" & vbCrLf &
              "    LEFT JOIN (SELECT NotasFiscais.Empresa_Id," & vbCrLf &
              "                      NotasFiscais.EndEmpresa_Id," & vbCrLf &
              "                      NotasFiscais.Pedido," & vbCrLf &
              "                      NotasFiscaisXItens.Produto_Id, " & vbCrLf &
              "                      SUM(CASE" & vbCrLf &
              "                            WHEN SubOperacoes.Devolucao   = 'N'" & vbCrLf &
              "                              THEN" & vbCrLf &
              "								    CASE" & vbCrLf &
              "									    WHEN SubOperacoes.Classe = 'SERVICOS'" & vbCrLf &
              "                                         THEN ISNULL(NotasFiscaisXItens.QuantidadeFiscal, 0)" & vbCrLf &
              "										     ELSE ISNULL(NotasFiscaisXItens.QuantidadeFisica, 0)" & vbCrLf &
              "									    END" & vbCrLf &
              "                              ELSE 0" & vbCrLf &
              "                          END) AS EntregueNota," & vbCrLf &
              "                      SUM(CASE" & vbCrLf &
              "                            WHEN SubOperacoes.Devolucao   = 'N'" & vbCrLf &
              "                              THEN ISNULL(NotasFiscaisXItens.Valor, 0)" & vbCrLf &
              "                              ELSE 0" & vbCrLf &
              "                          END) AS EntregueNotaValor," & vbCrLf &
              "                      SUM(CASE" & vbCrLf &
              "                            WHEN SubOperacoes.Devolucao   = 'S'" & vbCrLf &
              "                              THEN" & vbCrLf &
              "								    CASE" & vbCrLf &
              "									    WHEN SubOperacoes.Classe = 'SERVICOS'" & vbCrLf &
              "                                         THEN ISNULL(NotasFiscaisXItens.QuantidadeFiscal, 0)" & vbCrLf &
              "										     ELSE ISNULL(NotasFiscaisXItens.QuantidadeFisica, 0)" & vbCrLf &
              "									    END" & vbCrLf &
              "                              ELSE 0" & vbCrLf &
              "                          END) AS DevolucaoNota, " & vbCrLf &
              "                      SUM(CASE" & vbCrLf &
              "                             WHEN  SubOperacoes.Devolucao   = 'S'" & vbCrLf &
              "                               THEN ISNULL(NotasFiscaisXItens.Valor, 0)" & vbCrLf &
              "                               ELSE 0" & vbCrLf &
              "                          END) AS DevolucaoNotaValor" & vbCrLf &
              "                 FROM NotasFiscaisXItens " & vbCrLf &
              "                INNER JOIN NotasFiscais " & vbCrLf &
              "                   ON NotasFiscaisXItens.Empresa_Id      = NotasFiscais.Empresa_Id " & vbCrLf &
              "                  AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id" & vbCrLf &
              "                  AND NotasFiscaisXItens.Cliente_Id      = NotasFiscais.Cliente_Id " & vbCrLf &
              "                  AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscais.EndCliente_Id" & vbCrLf &
              "                  AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id " & vbCrLf &
              "                  AND NotasFiscaisXItens.Serie_Id        = NotasFiscais.Serie_Id " & vbCrLf &
              "                  AND NotasFiscaisXItens.Nota_Id         = NotasFiscais.Nota_Id " & vbCrLf &
              "                 LEFT JOIN SubOperacoes " & vbCrLf &
              "                   ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id " & vbCrLf &
              "                  AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
              "                Where SubOperacoes.Classe not in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "', '" & eClassesOperacoes.FISCAL.ToString & "')" & vbCrLf &
              "	                 AND NotasFiscais.Situacao   in (1,4)" & vbCrLf

        If txtPeriodoFinal.Text <> "" And ckConsiderarPeriodo.Checked Then
            Sql &= " and NotasFiscais.DataDaNota <= '" & txtPeriodoFinal.Text.ToSqlDate() & "'" & vbCrLf
        End If

        Sql &= "                GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Pedido, NotasFiscaisXItens.Produto_Id " & vbCrLf &
               "               ) AS NFI " & vbCrLf &
               "       ON NFI.Empresa_Id    = P.Empresa_Id " & vbCrLf &
               "      AND NFI.EndEmpresa_id = P.EndEmpresa_Id" & vbCrLf &
               "      AND NFI.Pedido        = P.Pedido_Id " & vbCrLf &
               "      AND NFI.Produto_Id    = PxI.Produto_Id " & vbCrLf

        '***********************  FATURADO NAO ENTREGUE  ***************************
        Sql &= "     LEFT JOIN (SELECT NotasFiscais.Pedido, " & vbCrLf &
               "                       NFI.Produto_Id, " & vbCrLf &
               "                       SUM(Case" & vbCrLf &
               "                             when SubOperacoes.Classe = '" & eClassesOperacoes.GLOBAL.ToString & "' " & vbCrLf &
               "                               then Case" & vbCrLf &
               "                                     when SubOperacoes.Devolucao = 'S'" & vbCrLf &
               "                                       then NFI.QuantidadeFiscal * - 1" & vbCrLf &
               "                                       else NFI.QuantidadeFiscal" & vbCrLf &
               "                                    end " & vbCrLf &
               "                               else 0" & vbCrLf &
               "                           end  " & vbCrLf &
               "                           )EntregueGlobal," & vbCrLf &
               "                       SUM(Case" & vbCrLf &
               "                             when SubOperacoes.Classe = '" & eClassesOperacoes.REMESSAS.ToString & "' " & vbCrLf &
               "                               then Case" & vbCrLf &
               "                                     when SubOperacoes.Devolucao = 'S'" & vbCrLf &
               "                                       then NFI.QuantidadeFiscal * - 1" & vbCrLf &
               "                                       else NFI.QuantidadeFiscal" & vbCrLf &
               "                                    end " & vbCrLf &
               "                               else 0" & vbCrLf &
               "                           end " & vbCrLf &
               "                           )EntregueRemessa" & vbCrLf &
               "                  FROM NotasFiscaisXItens NFI " & vbCrLf &
               "                 INNER JOIN NotasFiscais " & vbCrLf &
               "                    ON NFI.Empresa_Id = NotasFiscais.Empresa_Id " & vbCrLf &
               "                   AND NFI.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id " & vbCrLf &
               "                   AND NFI.Cliente_Id = NotasFiscais.Cliente_Id" & vbCrLf &
               "                   AND NFI.EndCliente_Id = NotasFiscais.EndCliente_Id " & vbCrLf &
               "                   AND NFI.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id" & vbCrLf &
               "                   AND NFI.Serie_Id = NotasFiscais.Serie_Id" & vbCrLf &
               "                   AND NFI.Nota_Id = NotasFiscais.Nota_Id" & vbCrLf &
               "                  LEFT JOIN SubOperacoes" & vbCrLf &
               "                    ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id" & vbCrLf &
               "                   AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
               "                 Where SubOperacoes.Classe  in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.REMESSAS.ToString & "')" & vbCrLf &
               "                   and NotasFiscais.situacao in (1,4)" & vbCrLf &
               "                 GROUP BY NotasFiscais.Pedido, NFI.Produto_Id" & vbCrLf &
               "               ) AS FNE" & vbCrLf &
               "        ON FNE.Pedido     = P.Pedido_Id" & vbCrLf &
               "       AND FNE.Produto_Id = PxI.Produto_Id" & vbCrLf &
               "      Left Join Produtos prd" & vbCrLf &
               "        on prd.Produto_id = PxI.Produto_id" & vbCrLf

        Sql &= " WHERE (P.Situacao =  1)" & vbCrLf &
               "   AND (SO.Operacao_Id <> 96)" & vbCrLf &
               "   AND so.Classe not in ('" & eClassesOperacoes.CONTAEORDEM.ToString & "', '" & eClassesOperacoes.FISCAL.ToString & "')" & vbCrLf

        'MOSTRAR APENAS PRODUTOS DE MATÉRIA PRIMA - FURLAN 11/12/2024
        If Left(Session("ssEmpresa").ToString, 8) = "24450490" Then
            Sql &= "   AND prd.Grupo in('10101','10102','20101','20102','20103','20104','20105','20106','30101','90203')" & vbCrLf
        End If

        If ddlMarca.SelectedIndex > 0 Then
            Sql &= "   AND prd.Marca = " & ddlMarca.SelectedValue & vbCrLf
        End If


        If Not String.IsNullOrWhiteSpace(DdlUnidadeConsultaTitulos.SelectedValue) Then
            Sql &= "AND P.UnidadeDeNegocio = '" & DdlUnidadeConsultaTitulos.SelectedValue & "'" & vbCrLf
        End If

        If DdlEmpresaConsultaTitulos.SelectedIndex > 0 Then
            Dim strEmpresa As String() = DdlEmpresaConsultaTitulos.SelectedValue.Split("-")

            If chkConsolidarEmpresa.Checked Then
                Sql &= " AND left(P.Empresa_Id,8) ='" & strEmpresa(0).Substring(0, 8) & "'"
                Parametros &= "Empresa Consolidada: " & DdlEmpresaConsultaTitulos.SelectedItem.Text & vbCrLf
            Else
                Sql &= " AND P.Empresa_Id    ='" & strEmpresa(0) & "' " & vbCrLf &
                       " AND P.EndEmpresa_Id = " & strEmpresa(1) & " " & vbCrLf
                Parametros = "Empresa:" & DdlEmpresaConsultaTitulos.SelectedItem.Text & vbCrLf
            End If
        End If

        If txtCodigoCliente.Value.Length > 0 Then
            Dim strCliente As String() = txtCodigoCliente.Value.Split(";")
            If chkConsolidarCliente.Checked Then
                Sql &= " AND left(P.Cliente, 8)    ='" & strCliente(0).Substring(0, 8) & "'" & vbCrLf

                Parametros &= "Cliente Consolidado: " & txtClientes.Text & vbCrLf
            Else
                Sql &= " AND P.Cliente    ='" & strCliente(0) & "'" & vbCrLf &
                       " AND P.EndCliente = " & strCliente(1) & vbCrLf
                Parametros &= "Cliente: " & txtClientes.Text & vbCrLf
            End If
        End If

        If lstClasses.GetSelectedValues().Count > 0 Then
            Sql &= " AND SO.Classe in ('" & String.Join("', '", getListClasses()) & "')"
            Parametros &= "Classes: " & String.Join(", ", getListClasses()) & vbCrLf
        End If

        Dim fin As String = String.Join("','", lstFinalidade.GetSelecteds)
        If fin.Length > 0 Then
            Sql &= "   AND isnull(P.Finalidade,0) " & IIf(rdComFinalidade.Checked, "", "not") & " in ('" & fin & "')"
        End If

        If cmbOperacao.SelectedIndex > 0 Then
            Sql &= "And ( P.operacao = '" & cmbOperacao.SelectedValue & "') " & vbCrLf
            Parametros &= "Operacao: " & cmbOperacao.SelectedItem.Text
        End If

        If cmbSubOperacao.SelectedIndex > 0 Then
            Sql &= "And ( P.Suboperacao = '" & cmbSubOperacao.SelectedValue.Split("-")(1) & "') " & vbCrLf
            Parametros &= "SubOperacao: " & cmbSubOperacao.SelectedItem.Text
        End If

        If rbTodas.Checked Then
            'Todas Entradas e Saídas
        ElseIf rbEntradas.Checked Then 'somente Entradas
            Sql &= "   AND (SO.EntradaSaida = 'E')" & vbCrLf
            Parametros &= "Pedidos de Entrada" & vbCrLf
        ElseIf rbSaidas.Checked Then 'somente Saídas
            Sql &= "   AND (SO.EntradaSaida = 'S')" & vbCrLf
            Parametros &= "Pedidos de Saida" & vbCrLf
        End If

        If ucSelecaoProduto.TemSelecionado Then
            Dim ret As ArrayList
            ret = ucSelecaoProduto.GetSqlEParametrosRelatorio("PxI.Grupo", "PxI.Produto_Id", "", True)
            Sql &= " And " & ret(0)
            Parametros &= ret(1) & vbCrLf
        ElseIf ddlMarca.SelectedIndex > 0 Then
            Sql &= " And PxI.Marca = " & ddlMarca.SelectedValue & vbCrLf
            Parametros &= "Marca Produto: " & ddlMarca.SelectedItem.Text & vbCrLf
        End If

        If cmbSafra.SelectedIndex > 0 Then
            Sql &= " AND ( P.Safra = '" & cmbSafra.SelectedValue & "' ) " & vbCrLf
            Parametros &= "Safra: " & cmbSafra.SelectedItem.Text & vbCrLf
        End If

        If cmbMoeda.SelectedIndex > 0 Then
            Sql &= " AND ( P.Moeda = '" & cmbMoeda.SelectedValue & "' ) " & vbCrLf
            Parametros &= "Moeda: " & cmbMoeda.SelectedItem.Text & vbCrLf
        End If

        If rdTrocaSim.Checked Then
            Sql &= "   AND isnull(p.troca,0) = 1" & vbCrLf
            Parametros &= "Pedidos de Troca" & vbCrLf
        ElseIf rdTrocaNao.Checked Then
            Sql &= "   AND isnull(p.troca,0) = 0" & vbCrLf
            Parametros &= "Pedidos que não são de Troca" & vbCrLf
        End If

        If chkFaturadoNaoEntregue.Checked Then
            Sql &= " And isnull(FNE.EntregueGlobal,0) - isnull(FNE.EntregueRemessa,0) <> 0"
            Parametros &= "Com Saldo de Faturado nao Entregue" & vbCrLf
        End If

        If txtPeriodoInicial.Text <> "" And txtPeriodoFinal.Text <> "" And ckConsiderarPeriodo.Checked Then
            Sql &= " AND (P.DataPedido between '" & txtPeriodoInicial.Text.ToSqlDate() & "' and '" & txtPeriodoFinal.Text.ToSqlDate() & "') " & vbCrLf
        End If

        If ddlRepresentante.SelectedIndex > 0 Then
            Dim rep As String() = ddlRepresentante.SelectedValue.Split("-")
            Sql &= " AND exists (Select 1 from Comissoes C where C.Empresa_Id = P.Empresa_id and C.EndEmpresa_Id = P.EndEmpresa_Id and C.Pedido_id = P.Pedido_Id and Representante_id ='" & rep(0) & "' and EndRepresentante_id = " & rep(1) & " )"
            Parametros &= "Representante: " & ddlRepresentante.SelectedItem.Text & vbCrLf
        End If

        'Sql &= "UPDATE #SaldoPedido Set QuantidadeSoma = EntregueNota, TotalOficial = EntregueNotaValor Where (Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' OR Classe = '" & eClassesOperacoes.AFIXAR.ToString & "'); " & vbCrLf
        Sql &= "UPDATE #SaldoPedido Set QuantidadeSoma = EntregueNota, TotalOficial = EntregueNotaValor Where (Classe = '" & eClassesOperacoes.AFIXAR.ToString & "'); " & vbCrLf

        Return Sql
    End Function

    Private Function getListClasses() As List(Of String)
        Try
            Dim lst As New List(Of String)

            For Each item As String In lstClasses.GetSelectedValues()
                lst.Add(item)
            Next
            Return lst
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Function getSql(ByVal Tipo As String, Optional ByRef Parametros As String = "") As String

        Dim Sql As String = ""

        If Tipo = "Base" Then

            Sql = SqlBase(Parametros)

            Sql &= "Select " & IIf(chkConsolidarEmpresa.Checked, "matriz.nome as EmpresaNome, '999' as EmpresaReduzido, matriz.cidade as EmpresaCidade,", "SP.EmpresaNome,SP.EmpresaReduzido,SP.EmpresaCidade,") & vbCrLf &
                   "       SP.Empresa_id," & vbCrLf &
                   "       SP.Pedido," & vbCrLf &
                   "       SP.PedidoEfetivo," & vbCrLf &
                   "       SP.EntradaSaida," & vbCrLf &
                   "       SP.Classificacao," & vbCrLf &
                   "       '(' + SP.Cifrao +') ' + SP.FreteCIFFOB AS Moeda," & vbCrLf &
                   "       SP.CondicaoPagamento," & vbCrLf &
                   "       SP.Cliente," & vbCrLf &
                   "       SP.NomeCliente," & vbCrLf &
                   "       SP.Representante," & vbCrLf &
                   "       SP.NomeRepresentante," & vbCrLf &
                   "       SP.Complemento," & vbCrLf &
                   "       SP.Cidade," & vbCrLf &
                   "       SP.Estado," & vbCrLf &
                   "       SP.DataPedido," & vbCrLf &
                   "       SP.DataEntrega," & vbCrLf &
                   "       SP.CodigoProduto," & vbCrLf &
                   "       SP.NomeProduto," & vbCrLf &
                   "       SP.Classe," & vbCrLf &
                   "       CONVERT(varchar, SP.Operacao + '-' + SP.SubOperacao) AS Operacao," & vbCrLf &
                   "       SP.DescOperacao," & vbCrLf &
                   "       SP.Embalagem," & vbCrLf &
                   "       SP.Quantidade," & vbCrLf &
                   "       SP.QuantidadeSoma," & vbCrLf &
                   "       SP.QuantidadeItemPedido," & vbCrLf &
                   "       SP.TotalOficial AS ValorOficial," & vbCrLf &
                   "       SP.ValorLiquido,       " & vbCrLf &
                   "       SP.ValorLiquidoMoeda,      " & vbCrLf &
                   "       SP.TotalMoeda AS ValorMoeda," & vbCrLf &
                   "       SP.FaturadoNaoEntregue," & vbCrLf &
                   "       SP.EntregueNota," & vbCrLf &
                   "       SP.EntregueNotaValor," & vbCrLf &
                   "       SP.DevolucaoNota," & vbCrLf &
                   "       SP.DevolucaoNotaValor," & vbCrLf &
                   "       Case When (SP.Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' AND SP.QuantidadeItemPedido = 0) Then case when ISNULL(SP.PesoDoItem, 0) > 0 then (SP.EntregueNota - SP.DevolucaoNota) * SP.PesoDoItem else (SP.EntregueNota - SP.DevolucaoNota) End else (SP.QuantidadeSoma - SP.EntregueNota + SP.DevolucaoNota) End AS SaldoNota, " & vbCrLf &
                   "       Case When (SP.Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' AND SP.QuantidadeItemPedido = 0) Then (SP.EntregueNotaValor - SP.DevolucaoNotaValor) Else (SP.totaloficial - SP.EntregueNotaValor + SP.DevolucaoNotaValor) End AS SaldoNotaValor, " & vbCrLf

            If rbSaca.Checked Then
                Parametros &= "SP.Unitario * 60" & vbCrLf

                Sql &= " Case" & vbCrLf &
                       "   When SP.Embalagem = 1 " & vbCrLf &
                       "     Then SP.UnitarioLiquido * 60 " & vbCrLf &
                       "     Else SP.UnitarioLiquido " & vbCrLf &
                       " End AS UnitarioLiquido, " & vbCrLf &
                       " Case" & vbCrLf &
                       "   When SP.Embalagem = 1 " & vbCrLf &
                       "     Then SP.UnitarioLiquidoMoeda * 60 " & vbCrLf &
                       "     Else SP.UnitarioLiquidoMoeda " & vbCrLf &
                       " End As UnitarioLiquidoMoeda, " & vbCrLf &
                       " Case" & vbCrLf &
                       "   When SP.Embalagem = 1 " & vbCrLf &
                       "     Then SP.UnitarioOficial * 60 " & vbCrLf &
                       "     Else SP.UnitarioOficial " & vbCrLf &
                       " End AS UnitarioOficial, " & vbCrLf &
                       " Case" & vbCrLf &
                       "   When SP.Embalagem = 1 " & vbCrLf &
                       "     Then SP.UnitarioMoeda * 60 " & vbCrLf &
                       "     Else SP.UnitarioMoeda " & vbCrLf &
                       " End As UnitarioMoeda, " & vbCrLf &
                       " Case" & vbCrLf &
                       "   When isnull(SP.QuantidadeSoma,0) > 0" & vbCrLf &
                       "     then Case when SP.Embalagem = 1 " & vbCrLf &
                       "            Then (SP.TotalOficial / SP.QuantidadeSoma) * 60 " & vbCrLf &
                       "            Else (SP.TotalOficial / SP.QuantidadeSoma) " & vbCrLf &
                       "          end" & vbCrLf &
                       "     else 0" & vbCrLf &
                       " End AS UnitarioOficialNota, " & vbCrLf &
                       " Case" & vbCrLf &
                       "   When isnull(SP.QuantidadeSoma,0) > 0" & vbCrLf &
                       "     then Case when SP.Embalagem = 1 " & vbCrLf &
                       "            Then (SP.TotalMoeda / SP.QuantidadeSoma) * 60 " & vbCrLf &
                       "            Else (SP.TotalMoeda / SP.QuantidadeSoma) " & vbCrLf &
                       "          end" & vbCrLf &
                       "      Else 0" & vbCrLf &
                       " End As UnitarioMoedaNota " & vbCrLf

            ElseIf rbKg.Checked Then
                Sql &= " SP.UnitarioOficial," & vbCrLf &
                       " SP.UnitarioMoeda," & vbCrLf &
                       " SP.UnitarioLiquido,        " & vbCrLf &
                       " SP.UnitarioLiquidoMoeda,   " & vbCrLf &
                       " case" & vbCrLf &
                       "    When isnull(SP.QuantidadeSoma,0) > 0" & vbCrLf &
                       "      then (SP.TotalOficial / SP.QuantidadeSoma)" & vbCrLf &
                       "      else 0" & vbCrLf &
                       " end UnitarioOficialNota," & vbCrLf &
                       " case" & vbCrLf &
                       "    when isnull(SP.QuantidadeSoma,0) > 0" & vbCrLf &
                       "      then (SP.TotalMoeda / SP.QuantidadeSoma)" & vbCrLf &
                       "      else 0" & vbCrLf &
                       " end UnitarioMoedaNota," & vbCrLf &
                       " SP.Observacoes" & vbCrLf
            End If

            Sql &= " , Case When (SP.Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' AND SP.QuantidadeItemPedido = 0) Then case when ISNULL(SP.PesoDoItem, 0) > 0 then (SP.EntregueNota - SP.DevolucaoNota) * SP.PesoDoItem else (SP.EntregueNota - SP.DevolucaoNota) End else (SP.QuantidadeSoma - SP.EntregueNota + SP.DevolucaoNota) END * SP.PesoDoItem AS SaldoKgNota, " & vbCrLf &
                   " Case When (SP.Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' AND SP.QuantidadeItemPedido = 0) Then (SP.EntregueNota - SP.DevolucaoNota) Else SP.QuantidadeItemPedido - (SP.EntregueNota * SP.PesoDoItem) End AS SaldoPedido, " & vbCrLf &
                   " SP.PedidoBloqueado" & vbCrLf &
                   " INTO #Saldo " & vbCrLf &
                   " FROM #SaldoPedido SP " & vbCrLf

            If chkConsolidarEmpresa.Checked Then
                Sql &= "  Inner join ClientesXEmpresas cxe" & vbCrLf &
                       "     on left(cxe.Empresa_Id,8) =  Left(SP.Empresa_id,8)" & vbCrLf &
                       "    and cxe.Matriz = 'S'" & vbCrLf &
                       "  Inner join Clientes Matriz" & vbCrLf &
                       "     on Matriz.Cliente_Id  = cxe.Empresa_Id" & vbCrLf &
                       "    and matriz.Endereco_Id = cxe.EndEmpresa_Id" & vbCrLf
            End If

            Sql &= "SELECT Empresa_id, EmpresaNome, EmpresaReduzido, Empresacidade, Pedido,PedidoEfetivo, EntradaSaida, classificacao, Moeda, CondicaoPagamento, Cliente, NomeCliente, Representante, NomeRepresentante, Complemento, Cidade, Estado, DataPedido, DataEntrega, Ge.Descricao, prd.NCM, CodigoProduto, CodigoProdutoTerceiro, NomeProduto, Classe," & vbCrLf &
                   "       Operacao, DescOperacao, Quantidade, QuantidadeItemPedido, UnitarioOficial, UnitarioMoeda, ValorOficial, ValorMoeda, FaturadoNaoEntregue, EntregueNota," & vbCrLf &
                   "       EntregueNotaValor, DevolucaoNota, DevolucaoNotaValor, SaldoNota, CONVERT(DECIMAL(18, 4), SaldoKgNota) AS SaldoNFKG,  SaldoNotaValor, CONVERT(DECIMAL(18, 4), SaldoPedido) AS SaldoPedido, " & vbCrLf &
                   "       Case When Classificacao = 'O' then UnitarioOficial     Else UnitarioMoeda     end as Unitario," & vbCrLf &
                   "       Case When Classificacao = 'O' then UnitarioOficialNota Else UnitarioMoedaNota end as UnitarioNF," & vbCrLf &
                   "       Case When Classificacao = 'O' then ValorOficial        Else ValorMoeda    end as Valor," & vbCrLf &
                   "       Case When Classificacao = 'O' then UnitarioLiquido     Else UnitarioLiquidoMoeda end as UnitarioLiquido," & vbCrLf &
                   "       Case When Classificacao = 'O' then ValorLiquido        Else ValorLiquidoMoeda    end as ValorLiquido," & vbCrLf &
                   "       Observacoes, PedidoBloqueado" & vbCrLf &
                   "  From #Saldo " & vbCrLf &
                   "  inner Join Produtos prd " & vbCrLf &
                   "  on prd.Produto_id = CodigoProduto " & vbCrLf &
                   "  inner Join GruposDeEstoques Ge " & vbCrLf &
                   "  on prd.grupo = Ge.Grupo_Id " & vbCrLf &
                   " Where 1 = 1 " & vbCrLf

            If chkSaldoQuantidade.Checked Then
                Parametros &= "Com Saldo de QuantidadeSoma" & vbCrLf
                Sql &= " And  SaldoNota > 0 " & vbCrLf
            End If

            If chkSaldoValor.Checked Then
                Parametros &= "Com Saldo de Valor" & vbCrLf
                Sql &= " And SaldoNotaValor > 0 " & vbCrLf
            End If

            If rdPorProduto.Checked Then
                Sql &= " order by Empresa_Id, NomeProduto,NomeCliente,DataPedido" & vbCrLf
            ElseIf rdPorPedido.Checked Then
                Sql &= " order by Empresa_Id, Pedido" & vbCrLf
            ElseIf rdPorCliente.Checked Then
                Sql &= " order by Empresa_Id, NomeCliente" & vbCrLf
            End If

            Return Sql  '"RelacaoDePedidos"

        ElseIf Tipo = "PorPedido" Then
            Sql &= "Select " & IIf(chkConsolidarEmpresa.Checked, "matriz.nome as EmpresaNome, '999' as EmpresaReduzido, matriz.cidade as EmpresaCidade,", "SP.EmpresaNome,SP.EmpresaReduzido,SP.EmpresaCidade,") & vbCrLf &
                   "       SP.NomeProduto, SP.Classe, CONVERT(varchar, SP.Operacao + '-' + SP.SubOperacao + ' ' + SP.DescOperacao) AS Operacao, " & vbCrLf &
                   "       case when SUM(SP.QuantidadeSoma) > 0 then (SUM(SP.TotalOficial) / SUM(SP.QuantidadeSoma)) else 0 end AS UnitReal, " & vbCrLf &
                   "       case when SUM(SP.QuantidadeSoma) > 0 then (SUM(SP.TotalMoeda) / SUM(SP.QuantidadeSoma)) else 0 end AS UnitMoeda, " & vbCrLf &
                   "       SUM(SP.QuantidadeSoma) AS Saldo, SUM(SP.TotalOficial) AS ValorReais, " & vbCrLf &
                   "       SUM(SP.TotalMoeda) AS ValorDolar " & vbCrLf &
                   "  FROM #SaldoPedido SP" & vbCrLf

            If chkConsolidarEmpresa.Checked Then
                Sql &= "  Inner join ClientesXEmpresas cxe" & vbCrLf &
                       "     on left(cxe.Empresa_Id,8) =  Left(SP.Empresa_id,8)" & vbCrLf &
                       "    and cxe.Matriz = 'S'" & vbCrLf &
                       "  Inner join Clientes Matriz" & vbCrLf &
                       "     on Matriz.Cliente_Id  = cxe.Empresa_Id" & vbCrLf &
                       "    and matriz.Endereco_Id = cxe.EndEmpresa_Id" & vbCrLf
            End If

            Sql &= " WHERE 1 = 1"

            If chkSaldoQuantidade.Checked Then
                Sql &= " And Case When (SP.Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' AND SP.QuantidadeItemPedido = 0) Then (SP.EntregueNota - SP.DevolucaoNota) Else (SP.QuantidadeSoma - SP.EntregueNota + SP.DevolucaoNota) End > 0 " & vbCrLf
                'Sql &= " And Case When (Classe = '"& eClassesOperacoes.DEPOSITOS.ToString &"' OR Classe = '"& eClassesOperacoes.AFIXAR.ToString &"') Then (EntregueNota - DevolucaoNota) Else (Quantidade - EntregueNota + DevolucaoNota) End > 0 " & vbCrLf
            End If

            If chkSaldoValor.Checked Then
                Sql &= " And Case When (SP.Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' AND SP.QuantidadeItemPedido = 0) Then (SP.EntregueNotaValor - SP.DevolucaoNotaValor) Else (SP.totaloficial - SP.EntregueNotaValor + SP.DevolucaoNotaValor) End > 0 " & vbCrLf
                'Sql &= " And Case When (Classe = '"& eClassesOperacoes.DEPOSITOS.ToString &"' OR Classe = '"& eClassesOperacoes.AFIXAR.ToString &"') Then (EntregueNotaValor - DevolucaoNotaValor) Else (totaloficial - EntregueNotaValor + DevolucaoNotaValor) End > 0 " & vbCrLf
            End If

            Sql &= " GROUP BY SP.NomeProduto, SP.Classe, CONVERT(varchar, SP.Operacao + '-' + SP.SubOperacao + ' ' + SP.DescOperacao)" & vbCrLf &
            IIf(chkConsolidarEmpresa.Checked, ",matriz.nome, matriz.cidade", ",SP.EmpresaNome, SP.EmpresaReduzido, SP.EmpresaCidade")

            Return Sql '"ResumoRelacaoPedidos"

        ElseIf Tipo = "PorCliente" Then
            Sql &= "Select " & IIf(chkConsolidarEmpresa.Checked, "matriz.nome as EmpresaNome, '999' as EmpresaReduzido, matriz.cidade as EmpresaCidade,", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade,") & vbCrLf &
                   "        Ge.Grupo_id," & vbCrLf &
                   "        Ge.Descricao as DescGrupo," & vbCrLf &
                   "        S.CodigoProduto," & vbCrLf &
                   "        S.NomeProduto," & vbCrLf &
                   "        S.EntradaSaida," & vbCrLf &
                   "        M.Descricao + ' - ' + S.Cifrao as Moeda," & vbCrLf &
                   "        case" & vbCrLf &
                   "           when S.Troca       = 1  then 'Troca'" & vbCrLf &
                   "           When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "           When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "           When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "           When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "           When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "        End as CondicaoPagamento,  " & vbCrLf &
                   "        sum(S.QuantidadeSoma) as Quantidade," & vbCrLf &
                   "        sum(case" & vbCrLf &
                   "              When S.Classificacao = 'O'" & vbCrLf &
                   "                then S.TotalOficial" & vbCrLf &
                   "                else S.TotalMoeda" & vbCrLf &
                   "            end) as Valor," & vbCrLf &
                   "        case" & vbCrLf &
                   "          When sum(S.QuantidadeSoma) = 0" & vbCrLf &
                   "            then 0" & vbCrLf &
                   "            Else (sum(case" & vbCrLf &
                   "                        When S.Classificacao = 'O'" & vbCrLf &
                   "                          then S.TotalOficial" & vbCrLf &
                   "                          else S.TotalMoeda" & vbCrLf &
                   "                      end) / (sum(S.QuantidadeSoma))) " & vbCrLf &
                   "        end as Unitario," & vbCrLf &
                   "        case                                       " & vbCrLf &
                   "          When sum(S.QuantidadeSoma) = 0               " & vbCrLf &
                   "            then 0                                 " & vbCrLf &
                   "            Else (sum(case                         " & vbCrLf &
                   "                        When S.Classificacao = 'O' " & vbCrLf &
                   "                          then S.valorLiquido      " & vbCrLf &
                   "                          else S.valorLiquidoMoeda " & vbCrLf &
                   "                      end) / (sum(S.QuantidadeSoma)))  " & vbCrLf &
                   "        end as UnitarioLiquido,                     " & vbCrLf &
                   "        sum(case                                   " & vbCrLf &
                   "              When S.Classificacao = 'O'           " & vbCrLf &
                   "                then S.valorliquido                " & vbCrLf &
                   "                Else S.valorliquidomoeda         " & vbCrLf &
                   "            end) as ValorLiquido                  " & vbCrLf &
                   "   into #Base" & vbCrLf &
                   "   from #SaldoPedido S" & vbCrLf &
                   "  inner Join Produtos prd" & vbCrLf &
                   "     on prd.Produto_id = S.CodigoProduto" & vbCrLf &
                   "  inner Join GruposDeEstoques Ge" & vbCrLf &
                   "     on prd.grupo = Ge.Grupo_Id" & vbCrLf &
                   "  Inner Join Moedas M" & vbCrLf &
                   "     on M.Moeda_id = S.Moeda" & vbCrLf

            If chkConsolidarEmpresa.Checked Then
                Sql &= "  Inner join ClientesXEmpresas cxe" & vbCrLf &
                       "     on left(cxe.Empresa_Id,8) =  Left(S.Empresa_id,8)" & vbCrLf &
                       "    and cxe.Matriz = 'S'" & vbCrLf &
                       "  Inner join Clientes Matriz" & vbCrLf &
                       "     on Matriz.Cliente_Id  = cxe.Empresa_Id" & vbCrLf &
                       "    and matriz.Endereco_Id = cxe.EndEmpresa_Id" & vbCrLf
            End If

            If chkSaldoQuantidade.Checked Then
                Sql &= "  where (Case " & vbCrLf &
                       "           When ((Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "') OR Classe = '" & eClassesOperacoes.AFIXAR.ToString & "')" & vbCrLf &
                       "             Then (EntregueNota - DevolucaoNota)" & vbCrLf &
                       "             Else (QuantidadeSoma - EntregueNota + DevolucaoNota)" & vbCrLf &
                       "         End) > 0" & vbCrLf
            End If

            If chkSaldoValor.Checked Then
                Sql &= IIf(chkSaldoQuantidade.Checked, " or ", " Where ") &
                       "    (Case " & vbCrLf &
                       "       When (Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "')" & vbCrLf &
                       "         Then (EntregueNotaValor - DevolucaoNotaValor)" & vbCrLf &
                       "         Else (totaloficial - EntregueNotaValor + DevolucaoNotaValor)" & vbCrLf &
                       "     End) > 0" & vbCrLf
            End If

            Sql &= " Group By ge.Grupo_id, Ge.Descricao, S.CodigoProduto, S.NomeProduto, S.EntradaSaida, M.Descricao + ' - ' + S.Cifrao," & vbCrLf &
                   "          case" & vbCrLf &
                   "             when S.Troca      = 1   then 'Troca'" & vbCrLf &
                   "             When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "             When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "             When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "             When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "             When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "          End, " & vbCrLf &
                   IIf(chkConsolidarEmpresa.Checked, "matriz.nome, matriz.cidade", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade") & vbCrLf

            Sql &= "Select * from (" & vbCrLf &
                   "Select  Ge.Grupo_id," & vbCrLf &
                   "        Ge.Descricao as DescGrupo," & vbCrLf &
                   IIf(chkConsolidarEmpresa.Checked, "matriz.nome as EmpresaNome, '999' as EmpresaReduzido, matriz.cidade as EmpresaCidade,", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade,") & vbCrLf &
                   "        Ge.Grupo_id as CodigoProduto," & vbCrLf &
                   "        Ge.Descricao as NomeProduto," & vbCrLf &
                   "        S.EntradaSaida," & vbCrLf &
                   "        M.Descricao + ' - ' + S.Cifrao as Moeda," & vbCrLf &
                   "        case" & vbCrLf &
                   "           when S.Troca       = 1  then 'Troca'" & vbCrLf &
                   "           When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "           When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "           When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "           When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "           When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "        End as CondicaoPagamento," & vbCrLf &
                   "        sum(S.QuantidadeSoma) as Quantidade," & vbCrLf &
                   "        sum(case" & vbCrLf &
                   "              When S.Classificacao = 'O'" & vbCrLf &
                   "                then S.TotalOficial" & vbCrLf &
                   "                else S.TotalMoeda" & vbCrLf &
                   "            end) as Valor," & vbCrLf &
                   "        case" & vbCrLf &
                   "          When sum(S.QuantidadeSoma) = 0" & vbCrLf &
                   "            then 0" & vbCrLf &
                   "            Else (sum(case" & vbCrLf &
                   "                        When S.Classificacao = 'O'" & vbCrLf &
                   "                          then S.TotalOficial" & vbCrLf &
                   "                          else S.TotalMoeda" & vbCrLf &
                   "                      end) / (sum(S.QuantidadeSoma)))" & vbCrLf &
                   "        end as Unitario," & vbCrLf &
                   "        case                                       " & vbCrLf &
                   "          When sum(S.QuantidadeSoma) = 0               " & vbCrLf &
                   "            then 0                                 " & vbCrLf &
                   "            Else (sum(case                         " & vbCrLf &
                   "                        When S.Classificacao = 'O' " & vbCrLf &
                   "                          then S.valorLiquido      " & vbCrLf &
                   "                          else S.valorLiquidoMoeda " & vbCrLf &
                   "                      end) / (sum(S.QuantidadeSoma)))  " & vbCrLf &
                   "        end as UnitarioLiquido,                    " & vbCrLf &
                   "        sum(case                                   " & vbCrLf &
                   "              When S.Classificacao = 'O'           " & vbCrLf &
                   "                then S.valorliquido                " & vbCrLf &
                   "                Else S.valorliquidomoeda         " & vbCrLf &
                   "            end) as ValorLiquido                  " & vbCrLf &
                   "   from #SaldoPedido S" & vbCrLf

            If chkConsolidarEmpresa.Checked Then
                Sql &= "  Inner join ClientesXEmpresas cxe" & vbCrLf &
                       "     on left(cxe.Empresa_Id,8) =  Left(S.Empresa_id,8)" & vbCrLf &
                       "    and cxe.Matriz = 'S'" & vbCrLf &
                       "  Inner join Clientes Matriz" & vbCrLf &
                       "     on Matriz.Cliente_Id  = cxe.Empresa_Id" & vbCrLf &
                       "    and Matriz.Endereco_Id = cxe.EndEmpresa_Id" & vbCrLf
            End If

            Sql &= "  inner Join Produtos prd" & vbCrLf &
                   "     on prd.Produto_id = S.CodigoProduto" & vbCrLf &
                   "  inner Join GruposDeEstoques Ge" & vbCrLf &
                   "     on left(prd.grupo,1) = Ge.Grupo_Id" & vbCrLf &
                   "  Inner Join Moedas M" & vbCrLf &
                   "     on M.Moeda_id = S.Moeda" & vbCrLf &
                   " Group By ge.Grupo_id, Ge.Descricao,  S.EntradaSaida, M.Descricao + ' - ' + S.Cifrao," & vbCrLf &
                   "          case" & vbCrLf &
                   "             when S.Troca      = 1   then 'Troca'" & vbCrLf &
                   "             When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "             When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "             When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "             When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "             When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "          End, " & vbCrLf &
                   IIf(chkConsolidarEmpresa.Checked, "matriz.nome, matriz.cidade", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade") & vbCrLf &
                   "  union all" & vbCrLf &
                   " select Ge.Grupo_id," & vbCrLf &
                   "        Ge.Descricao as DescGrupo," & vbCrLf &
                   IIf(chkConsolidarEmpresa.Checked, "matriz.nome as EmpresaNome, '999' as EmpresaReduzido, matriz.cidade as EmpresaCidade,", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade,") & vbCrLf &
                   "        Ge.Grupo_id as CodigoProduto," & vbCrLf &
                   "        Ge.Descricao as NomeProduto," & vbCrLf &
                   "        S.EntradaSaida," & vbCrLf &
                   "        M.Descricao + ' - ' + S.Cifrao as Moeda," & vbCrLf &
                   "        case" & vbCrLf &
                   "           when S.Troca       = 1  then 'Troca'" & vbCrLf &
                   "           When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "           When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "           When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "           When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "           When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "        End as CondicaoPagamento, " & vbCrLf &
                   "        sum(S.QuantidadeSoma) as Quantidade," & vbCrLf &
                   "        sum(case" & vbCrLf &
                   "              When S.Classificacao = 'O'" & vbCrLf &
                   "                then S.TotalOficial" & vbCrLf &
                   "                else S.TotalMoeda" & vbCrLf &
                   "            end) as Valor," & vbCrLf &
                   "        case" & vbCrLf &
                   "          When sum(S.QuantidadeSoma) = 0" & vbCrLf &
                   "            then 0" & vbCrLf &
                   "            Else (sum(case" & vbCrLf &
                   "                        When S.Classificacao = 'O'" & vbCrLf &
                   "                          then S.TotalOficial" & vbCrLf &
                   "                          else S.TotalMoeda" & vbCrLf &
                   "                      end) / (sum(S.QuantidadeSoma))) " & vbCrLf &
                   "        end as Unitario," & vbCrLf &
                   "        case                                       " & vbCrLf &
                   "          When sum(S.QuantidadeSoma) = 0               " & vbCrLf &
                   "            then 0                                 " & vbCrLf &
                   "            Else (sum(case                         " & vbCrLf &
                   "                        When S.Classificacao = 'O' " & vbCrLf &
                   "                          then S.valorLiquido      " & vbCrLf &
                   "                          else S.valorLiquidoMoeda " & vbCrLf &
                   "                      end) / (sum(S.QuantidadeSoma)))  " & vbCrLf &
                   "        end as UnitarioLiquido,                     " & vbCrLf &
                   "        sum(case                                   " & vbCrLf &
                   "              When S.Classificacao = 'O'           " & vbCrLf &
                   "                then S.valorliquido                " & vbCrLf &
                   "                Else S.valorliquidomoeda         " & vbCrLf &
                   "            end) as ValorLiquido                  " & vbCrLf &
                   "   from #SaldoPedido S" & vbCrLf

            If chkConsolidarEmpresa.Checked Then
                Sql &= "  Inner join ClientesXEmpresas cxe" & vbCrLf &
                       "     on left(cxe.Empresa_Id,8) =  Left(S.Empresa_id,8)" & vbCrLf &
                       "    and cxe.Matriz = 'S'" & vbCrLf &
                       "  Inner join Clientes Matriz" & vbCrLf &
                       "     on Matriz.Cliente_Id  = cxe.Empresa_Id" & vbCrLf &
                       "    and Matriz.Endereco_Id = cxe.EndEmpresa_Id" & vbCrLf
            End If

            Sql &= "  inner Join Produtos prd" & vbCrLf &
                   "     on prd.Produto_id = S.CodigoProduto" & vbCrLf &
                   "  inner Join GruposDeEstoques Ge" & vbCrLf &
                   "     on left(prd.grupo,2) = Ge.Grupo_Id" & vbCrLf &
                   "  Inner Join Moedas M" & vbCrLf &
                   "     on M.Moeda_id = S.Moeda" & vbCrLf &
                   " Group By ge.Grupo_id, Ge.Descricao,  S.EntradaSaida, M.Descricao + ' - ' + S.Cifrao," & vbCrLf &
                   "          case" & vbCrLf &
                   "             when S.Troca      = 1   then 'Troca'" & vbCrLf &
                   "             When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "             When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "             When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "             When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "             When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "          End, " & vbCrLf &
                   IIf(chkConsolidarEmpresa.Checked, "matriz.nome, matriz.cidade", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade") & vbCrLf &
                   "  union all" & vbCrLf &
                   " select Ge.Grupo_id," & vbCrLf &
                   "        Ge.Descricao as DescGrupo," & vbCrLf &
                   IIf(chkConsolidarEmpresa.Checked, "matriz.nome as EmpresaNome, '999' as EmpresaReduzido, matriz.cidade as EmpresaCidade,", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade,") & vbCrLf &
                   "        Ge.Grupo_id as CodigoProduto," & vbCrLf &
                   "        Ge.Descricao as NomeProduto," & vbCrLf &
                   "        S.EntradaSaida," & vbCrLf &
                   "        M.Descricao + ' - ' + S.Cifrao as Moeda," & vbCrLf &
                   "        case" & vbCrLf &
                   "           when S.Troca       = 1  then 'Troca'" & vbCrLf &
                   "           When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "           When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "           When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "           When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "           When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "        End as CondicaoPagamento," & vbCrLf &
                   "        sum(S.QuantidadeSoma) as Quantidade," & vbCrLf &
                   "        sum(case" & vbCrLf &
                   "              When S.Classificacao = 'O'" & vbCrLf &
                   "                then S.TotalOficial" & vbCrLf &
                   "                else S.TotalMoeda" & vbCrLf &
                   "            end) as Valor," & vbCrLf &
                   "        case" & vbCrLf &
                   "          When sum(S.QuantidadeSoma) = 0" & vbCrLf &
                   "            then 0" & vbCrLf &
                   "            Else (sum(case" & vbCrLf &
                   "                        When S.Classificacao = 'O'" & vbCrLf &
                   "                          then S.TotalOficial" & vbCrLf &
                   "                          else S.TotalMoeda" & vbCrLf &
                   "                      end) / (sum(S.QuantidadeSoma))) " & vbCrLf &
                   "        end as Unitario," & vbCrLf &
                   "        case                                       " & vbCrLf &
                   "          When sum(S.QuantidadeSoma) = 0               " & vbCrLf &
                   "            then 0                                 " & vbCrLf &
                   "            Else (sum(case                         " & vbCrLf &
                   "                        When S.Classificacao = 'O' " & vbCrLf &
                   "                          then S.valorLiquido      " & vbCrLf &
                   "                          else S.valorLiquidoMoeda " & vbCrLf &
                   "                      end) / (sum(S.QuantidadeSoma)))  " & vbCrLf &
                   "        end as UnitarioLiquido,                     " & vbCrLf &
                   "        sum(case                                   " & vbCrLf &
                   "              When S.Classificacao = 'O'           " & vbCrLf &
                   "                then S.valorliquido                " & vbCrLf &
                   "                Else S.valorliquidomoeda        " & vbCrLf &
                   "            end) as ValorLiquido                  " & vbCrLf &
                   "   from #SaldoPedido S" & vbCrLf

            If chkConsolidarEmpresa.Checked Then
                Sql &= "  Inner join ClientesXEmpresas cxe" & vbCrLf &
                       "     on left(cxe.Empresa_Id,8) =  Left(S.Empresa_id,8)" & vbCrLf &
                       "    and cxe.Matriz = 'S'" & vbCrLf &
                       "  Inner join Clientes Matriz" & vbCrLf &
                       "     on Matriz.Cliente_Id  = cxe.Empresa_Id" & vbCrLf &
                       "    and Matriz.Endereco_Id = cxe.EndEmpresa_Id" & vbCrLf
            End If

            Sql &= "  inner Join Produtos prd" & vbCrLf &
                   "     on prd.Produto_id = S.CodigoProduto" & vbCrLf &
                   "  inner Join GruposDeEstoques Ge" & vbCrLf &
                   "     on left(prd.grupo,3) = Ge.Grupo_Id" & vbCrLf &
                   "  Inner Join Moedas M" & vbCrLf &
                   "     on M.Moeda_id = S.Moeda" & vbCrLf &
                   " Group By ge.Grupo_id, Ge.Descricao,  S.EntradaSaida, M.Descricao + ' - ' + S.Cifrao," & vbCrLf &
                   "          case" & vbCrLf &
                   "             when S.Troca      = 1   then 'Troca'" & vbCrLf &
                   "             When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "             When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "             When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "             When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "             When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "          End, " & vbCrLf &
                   IIf(chkConsolidarEmpresa.Checked, "Matriz.nome, Matriz.cidade", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade") & vbCrLf &
                   "  union all" & vbCrLf &
                   " select Ge.Grupo_id," & vbCrLf &
                   "        Ge.Descricao as DescGrupo," & vbCrLf &
                   IIf(chkConsolidarEmpresa.Checked, "matriz.nome as EmpresaNome, '999' as EmpresaReduzido, matriz.cidade as EmpresaCidade,", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade,") & vbCrLf &
                   "        Ge.Grupo_id as CodigoProduto," & vbCrLf &
                   "        Ge.Descricao as NomeProduto," & vbCrLf &
                   "        S.EntradaSaida," & vbCrLf &
                   "        M.Descricao + ' - ' + S.Cifrao as Moeda," & vbCrLf &
                   "        case" & vbCrLf &
                   "           when S.Troca       = 1  then 'Troca'" & vbCrLf &
                   "           When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "           When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "           When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "           When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "           When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "        End as CondicaoPagamento," & vbCrLf &
                   "        sum(S.QuantidadeSoma) as Quantidade," & vbCrLf &
                   "        sum(case" & vbCrLf &
                   "              When S.Classificacao = 'O'" & vbCrLf &
                   "                then S.TotalOficial" & vbCrLf &
                   "                else S.TotalMoeda" & vbCrLf &
                   "            end) as Valor," & vbCrLf &
                   "        case" & vbCrLf &
                   "          When sum(S.QuantidadeSoma) = 0" & vbCrLf &
                   "            then 0" & vbCrLf &
                   "            Else (sum(case" & vbCrLf &
                   "                        When S.Classificacao = 'O'" & vbCrLf &
                   "                          then S.TotalOficial" & vbCrLf &
                   "                          else S.TotalMoeda" & vbCrLf &
                   "                      end) / (sum(S.QuantidadeSoma))) " & vbCrLf &
                   "        end as Unitario," & vbCrLf &
                   "        case                                       " & vbCrLf &
                   "          When sum(S.QuantidadeSoma) = 0               " & vbCrLf &
                   "            then 0                                 " & vbCrLf &
                   "            Else (sum(case                         " & vbCrLf &
                   "                        When S.Classificacao = 'O' " & vbCrLf &
                   "                          then S.valorLiquido      " & vbCrLf &
                   "                          else S.valorLiquidoMoeda " & vbCrLf &
                   "                      end) / (sum(S.QuantidadeSoma)))  " & vbCrLf &
                   "        end as UnitarioLiquido,                    " & vbCrLf &
                   "        sum(case                                   " & vbCrLf &
                   "              When S.Classificacao = 'O'           " & vbCrLf &
                   "                then S.valorliquido                " & vbCrLf &
                   "                Else S.valorliquidomoeda         " & vbCrLf &
                   "            end) as ValorLiquido                  " & vbCrLf &
                   "   from #SaldoPedido S" & vbCrLf

            If chkConsolidarEmpresa.Checked Then
                Sql &= "  Inner join ClientesXEmpresas cxe" & vbCrLf &
                       "     on left(cxe.Empresa_Id,8) =  Left(S.Empresa_id,8)" & vbCrLf &
                       "    and cxe.Matriz = 'S'" & vbCrLf &
                       "  Inner join Clientes Matriz" & vbCrLf &
                       "     on Matriz.Cliente_Id  = cxe.Empresa_Id" & vbCrLf &
                       "    and Matriz.Endereco_Id = cxe.EndEmpresa_Id" & vbCrLf
            End If

            Sql &= "  inner Join Produtos prd" & vbCrLf &
                   "     on prd.Produto_id = S.CodigoProduto" & vbCrLf &
                   "  inner Join GruposDeEstoques Ge" & vbCrLf &
                   "     on prd.grupo = Ge.Grupo_Id" & vbCrLf &
                   "  Inner Join Moedas M" & vbCrLf &
                   "     on M.Moeda_id = S.Moeda" & vbCrLf &
                   " Group By ge.Grupo_id, Ge.Descricao, S.EntradaSaida, M.Descricao + ' - ' + S.Cifrao," & vbCrLf &
                   "          case" & vbCrLf &
                   "             when S.Troca      = 1   then 'Troca'" & vbCrLf &
                   "             When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "             When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "             When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "             When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "             When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "          End, " & vbCrLf &
                   IIf(chkConsolidarEmpresa.Checked, "matriz.nome, matriz.cidade", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade") & vbCrLf &
                   "  union all" & vbCrLf &
                   " select Ge.Grupo_id," & vbCrLf &
                   "        Ge.Descricao as DescGrupo," & vbCrLf &
                   IIf(chkConsolidarEmpresa.Checked, "matriz.nome as EmpresaNome, '999' as EmpresaReduzido, matriz.cidade as EmpresaCidade,", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade,") & vbCrLf &
                   "        S.CodigoProduto," & vbCrLf &
                   "        S.NomeProduto," & vbCrLf &
                   "        S.EntradaSaida," & vbCrLf &
                   "        M.Descricao + ' - ' + S.Cifrao as Moeda," & vbCrLf &
                   "        case" & vbCrLf &
                   "           when S.Troca       = 1  then 'Troca'" & vbCrLf &
                   "           When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "           When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "           When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "           When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "           When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "        End as CondicaoPagamento," & vbCrLf &
                   "        sum(S.QuantidadeSoma) as Quantidade," & vbCrLf &
                   "        sum(case" & vbCrLf &
                   "              When S.Classificacao = 'O'" & vbCrLf &
                   "                then S.TotalOficial" & vbCrLf &
                   "                else S.TotalMoeda" & vbCrLf &
                   "            end) as Valor," & vbCrLf &
                   "        case" & vbCrLf &
                   "          When sum(S.QuantidadeSoma) = 0" & vbCrLf &
                   "            then 0" & vbCrLf &
                   "            Else (sum(case" & vbCrLf &
                   "                        When S.Classificacao = 'O'" & vbCrLf &
                   "                          then S.TotalOficial" & vbCrLf &
                   "                          else S.TotalMoeda" & vbCrLf &
                   "                      end) / (sum(S.QuantidadeSoma))) " & vbCrLf &
                   "        end as Unitario," & vbCrLf &
                   "        case                                       " & vbCrLf &
                   "          When sum(S.QuantidadeSoma) = 0               " & vbCrLf &
                   "            then 0                                 " & vbCrLf &
                   "            Else (sum(case                         " & vbCrLf &
                   "                        When S.Classificacao = 'O' " & vbCrLf &
                   "                          then S.valorLiquido      " & vbCrLf &
                   "                          else S.valorLiquidoMoeda " & vbCrLf &
                   "                      end) / (sum(S.QuantidadeSoma)))  " & vbCrLf &
                   "        end as UnitarioLiquido,                     " & vbCrLf &
                   "        sum(case                                   " & vbCrLf &
                   "              When S.Classificacao = 'O'           " & vbCrLf &
                   "                then S.valorliquido                " & vbCrLf &
                   "                Else S.valorliquidomoeda         " & vbCrLf &
                   "            end) as ValorLiquido                  " & vbCrLf &
                   "   from #SaldoPedido S" & vbCrLf

            If chkConsolidarEmpresa.Checked Then
                Sql &= "  Inner join ClientesXEmpresas cxe" & vbCrLf &
                       "     on left(cxe.Empresa_Id,8) =  Left(S.Empresa_id,8)" & vbCrLf &
                       "    and cxe.Matriz = 'S'" & vbCrLf &
                       "  Inner join Clientes Matriz" & vbCrLf &
                       "     on Matriz.Cliente_Id  = cxe.Empresa_Id" & vbCrLf &
                       "    and Matriz.Endereco_Id = cxe.EndEmpresa_Id" & vbCrLf
            End If

            Sql &= "  inner Join Produtos prd" & vbCrLf &
                   "     on prd.Produto_id = S.CodigoProduto" & vbCrLf &
                   "  inner Join GruposDeEstoques Ge" & vbCrLf &
                   "     on prd.grupo = Ge.Grupo_Id" & vbCrLf &
                   "  Inner Join Moedas M" & vbCrLf &
                   "     on M.Moeda_id = S.Moeda" & vbCrLf &
                   " Group By ge.Grupo_id, Ge.Descricao, S.CodigoProduto, S.NomeProduto, S.EntradaSaida, M.Descricao + ' - ' + S.Cifrao," & vbCrLf &
                   "          case" & vbCrLf &
                   "             when S.Troca      = 1   then 'Troca'" & vbCrLf &
                   "             When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "             When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "             When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "             When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "             When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "          End, " & vbCrLf &
                   IIf(chkConsolidarEmpresa.Checked, "matriz.nome, matriz.cidade", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade") & vbCrLf &
                   " ) sb" & vbCrLf &
                   " order by empresaReduzido, grupo_id, len(codigoProduto)" & vbCrLf &
                   "--*****************************" & vbCrLf &
                   "--Resumo de Moeda por Cliente" & vbCrLf &
                   "--*****************************" & vbCrLf &
                    "SELECT S.NomeCliente," & vbCrLf &
                    IIf(chkConsolidarEmpresa.Checked, "'999' as EmpresaReduzido,", "S.EmpresaReduzido,") & vbCrLf &
                    "       case" & vbCrLf &
                    "          when S.Troca      = 1   then 'Troca'" & vbCrLf &
                    "          When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                    "          When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                    "          When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                    "          When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                    "          When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                    "       End CondicaoPagamento," & vbCrLf &
                    "       M.Descricao + ' - ' + M.Cifrao as NomeMoeda," & vbCrLf &
                    "       sum(S.QuantidadeSoma) as Quantidade," & vbCrLf &
                    "       sum(case" & vbCrLf &
                    "             when S.Classificacao = 'O'" & vbCrLf &
                    "               then S.totaloficial" & vbCrLf &
                    "               else S.totalmoeda" & vbCrLf &
                    "           end) as Valor," & vbCrLf &
                    "        case" & vbCrLf &
                    "          when sum(S.QuantidadeSoma) = 0" & vbCrLf &
                    "            then 0" & vbCrLf &
                    "            else sum(case" & vbCrLf &
                    "					   when S.Classificacao = 'O'" & vbCrLf &
                    "					     then S.totaloficial" & vbCrLf &
                    "					     else S.totalmoeda" & vbCrLf &
                    "				     end) / sum(S.QuantidadeSoma)" & vbCrLf &
                    "        end as Unitario, EntradaSaida, " & vbCrLf &
                    "        case                                       " & vbCrLf &
                    "          When sum(S.QuantidadeSoma) = 0               " & vbCrLf &
                    "            then 0                                 " & vbCrLf &
                    "            Else (sum(case                         " & vbCrLf &
                    "                        When S.Classificacao = 'O' " & vbCrLf &
                    "                          then S.valorLiquido      " & vbCrLf &
                    "                          else S.valorLiquidoMoeda " & vbCrLf &
                    "                      end) / (sum(S.QuantidadeSoma)))  " & vbCrLf &
                    "        end as UnitarioLiquido,                     " & vbCrLf &
                    "        sum(case                                   " & vbCrLf &
                    "              When S.Classificacao = 'O'           " & vbCrLf &
                    "                then S.valorliquido                " & vbCrLf &
                    "                Else S.valorliquidomoeda         " & vbCrLf &
                    "            end) as ValorLiquido                  " & vbCrLf &
                    "  FROM #SaldoPedido S" & vbCrLf &
                    " Inner Join Moedas M" & vbCrLf &
                    "    on S.Moeda = M.Moeda_Id" & vbCrLf &
                    " Inner Join Pedidos P" & vbCrLf &
                    "    on S.Empresa_id    = P.Empresa_id" & vbCrLf &
                    "   and S.EndEmpresa_id = P.EndEmpresa_id" & vbCrLf &
                    "   and S.Pedido        = P.Pedido_id" & vbCrLf &
                    "  Left join Pagamentos Pg" & vbCrLf &
                    "    on Pg.Pagamento_Id = P.CondicaoPagamento" & vbCrLf &
                    " GROUP BY S.NomeCliente," & vbCrLf &
                    "       case" & vbCrLf &
                    "          when S.Troca      = 1   then 'Troca'" & vbCrLf &
                    "          When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                    "          When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                    "          When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                    "          When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                    "          When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                    "       End," & vbCrLf &
                    "       M.Descricao + ' - ' + M.Cifrao, EntradaSaida" & vbCrLf &
                    IIf(chkConsolidarEmpresa.Checked, "", ",S.EmpresaReduzido") & vbCrLf

            Return Sql 'Resumo de Moeda por Cliente

        ElseIf Tipo = "PorProduto" Then
            Sql &= "SELECT S.CodigoProduto, S.NomeProduto," & vbCrLf &
                   IIf(chkConsolidarEmpresa.Checked, "matriz.nome as EmpresaNome, '999' as EmpresaReduzido, matriz.cidade as EmpresaCidade,", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade,") & vbCrLf &
                   "       case" & vbCrLf &
                   "          when S.Troca      = 1   then 'Troca'" & vbCrLf &
                   "          When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "          When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "          When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "          When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "          When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "       End CondicaoPagamento," & vbCrLf &
                   "       M.Descricao + ' - ' + M.Cifrao as NomeMoeda," & vbCrLf &
                   "       sum(S.QuantidadeSoma) as Quantidade," & vbCrLf &
                   "       sum(case" & vbCrLf &
                   "             when S.Classificacao = 'O'" & vbCrLf &
                   "               then S.totaloficial" & vbCrLf &
                   "               else S.totalmoeda" & vbCrLf &
                   "           end) as Valor," & vbCrLf &
                   "        case " & vbCrLf &
                   "          when sum(S.QuantidadeSoma) = 0" & vbCrLf &
                   "            then 0" & vbCrLf &
                   "            else sum(case" & vbCrLf &
                   "					   when S.Classificacao = 'O'" & vbCrLf &
                   "					     then S.totaloficial" & vbCrLf &
                   "					     else S.totalmoeda" & vbCrLf &
                   "				     end) / sum(S.QuantidadeSoma)" & vbCrLf &
                   "        end as Unitario, " & vbCrLf &
                   "        case                                   " & vbCrLf &
                   "          when sum(S.QuantidadeSoma) = 0               " & vbCrLf &
                   "            then 0                                 " & vbCrLf &
                   "            else sum(case                          " & vbCrLf &
                   "					   when S.Classificacao = 'O'  " & vbCrLf &
                   "					     then S.valorLiquido       " & vbCrLf &
                   "					     else S.valorLiquidoMoeda  " & vbCrLf &
                   "				     end) / sum(S.QuantidadeSoma)      " & vbCrLf &
                   "        end as UnitarioLiquido,                    " & vbCrLf &
                   "        sum(case                                   " & vbCrLf &
                   "             when S.Classificacao = 'O'            " & vbCrLf &
                   "               then S.valorLiquido                 " & vbCrLf &
                   "               else S.valorLiquidoMoeda            " & vbCrLf &
                   "           end) as ValorLiquido,                   " & vbCrLf &
                   "        EntradaSaida" & vbCrLf &
                   "  FROM #SaldoPedido S" & vbCrLf

            If chkConsolidarEmpresa.Checked Then
                Sql &= "  Inner join ClientesXEmpresas cxe" & vbCrLf &
                       "     on left(cxe.Empresa_Id,8) =  Left(S.Empresa_id,8)" & vbCrLf &
                       "    and cxe.Matriz = 'S'" & vbCrLf &
                       "  Inner join Clientes Matriz" & vbCrLf &
                       "     on Matriz.Cliente_Id  = cxe.Empresa_Id" & vbCrLf &
                       "    and Matriz.Endereco_Id = cxe.EndEmpresa_Id" & vbCrLf
            End If

            Sql &= " Inner Join Moedas M" & vbCrLf &
                   "    on S.Moeda = M.Moeda_Id" & vbCrLf &
                   " Inner Join Pedidos P" & vbCrLf &
                   "    on S.Empresa_id    = P.Empresa_Id" & vbCrLf &
                   "   and S.EndEmpresa_id = P.EndEmpresa_Id" & vbCrLf &
                   "   and S.Pedido        = P.Pedido_id" & vbCrLf &
                   "  Left join Pagamentos Pg" & vbCrLf &
                   "    on Pg.Pagamento_Id = P.CondicaoPagamento" & vbCrLf &
                   " WHERE 1 = 1" & vbCrLf

            If chkSaldoQuantidade.Checked Then
                Sql &= " And Case When (Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' AND S.QuantidadeItemPedido = 0) Then (EntregueNota - DevolucaoNota) Else (QuantidadeSoma - EntregueNota + DevolucaoNota) End > 0 " & vbCrLf
                'Sql &= " And Case When (Classe = '"& eClassesOperacoes.DEPOSITOS.ToString &"' OR Classe = '"& eClassesOperacoes.AFIXAR.ToString &"') Then (EntregueNota - DevolucaoNota) Else (Quantidade - EntregueNota + DevolucaoNota) End > 0 " & vbCrLf
            End If

            If chkSaldoValor.Checked Then
                Sql &= " And Case When (Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' AND S.QuantidadeItemPedido = 0) Then (EntregueNotaValor - DevolucaoNotaValor) Else (totaloficial - EntregueNotaValor + DevolucaoNotaValor) End > 0 " & vbCrLf
                'Sql &= " And Case When (Classe = '"& eClassesOperacoes.DEPOSITOS.ToString &"' OR Classe = '"& eClassesOperacoes.AFIXAR.ToString &"') Then (EntregueNotaValor - DevolucaoNotaValor) Else (totaloficial - EntregueNotaValor + DevolucaoNotaValor) End > 0 " & vbCrLf
            End If

            Sql &= " GROUP BY S.CodigoProduto, S.NomeProduto," & vbCrLf &
                   "       case" & vbCrLf &
                   "          when S.Troca      = 1   then 'Troca'" & vbCrLf &
                   "          When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "          When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "          When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "          When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "          When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "       End," & vbCrLf &
                   "       M.Descricao + ' - ' + M.Cifrao, EntradaSaida," & vbCrLf &
                   IIf(chkConsolidarEmpresa.Checked, "matriz.nome, matriz.cidade", "S.EmpresaReduzido, S.EmpresaNome, S.EmpresaCidade") & vbCrLf


            Sql &= "SELECT Case" & vbCrLf &
                   "          when S.Troca      = 1   then 'Troca'" & vbCrLf &
                   "          When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "          When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "          When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "          When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "          When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "       End CondicaoPagamento," & vbCrLf &
                   "       M.Descricao + ' - ' + M.Cifrao as NomeMoeda," & vbCrLf &
                   "       sum(S.QuantidadeSoma) as Quantidade," & vbCrLf &
                   "       sum(case" & vbCrLf &
                   "             when S.Classificacao = 'O'" & vbCrLf &
                   "               then S.totaloficial" & vbCrLf &
                   "               else S.totalmoeda" & vbCrLf &
                   "           end) as Valor," & vbCrLf &
                   "        case " & vbCrLf &
                   "          when sum(S.QuantidadeSoma) = 0" & vbCrLf &
                   "            then 0" & vbCrLf &
                   "            else sum(case" & vbCrLf &
                   "					   when S.Classificacao = 'O'" & vbCrLf &
                   "					     then S.totaloficial" & vbCrLf &
                   "					     else S.totalmoeda" & vbCrLf &
                   "				     end) / sum(S.QuantidadeSoma)" & vbCrLf &
                   "        end as Unitario, " & vbCrLf &
                   "        sum(case                                  " & vbCrLf &
                   "             when S.Classificacao = 'O'           " & vbCrLf &
                   "               then S.valorLiquido                " & vbCrLf &
                   "        Else S.valorLiquidomoeda                  " & vbCrLf &
                   "           end) as ValorLiquido,                  " & vbCrLf &
                   "        case                                      " & vbCrLf &
                   "          when sum(S.QuantidadeSoma) = 0              " & vbCrLf &
                   "            then 0                                " & vbCrLf &
                   "            else sum(case                         " & vbCrLf &
                   "					   when S.Classificacao = 'O' " & vbCrLf &
                   "					     then S.ValorLiquido      " & vbCrLf &
                   "					     else S.ValorLiquidoMoeda " & vbCrLf &
                   "				     end) / sum(S.QuantidadeSoma)     " & vbCrLf &
                   "        end as UnitarioLiquido,                   " & vbCrLf &
                   "       EntradaSaida" & vbCrLf &
                   "  FROM #SaldoPedido S" & vbCrLf &
                   " Inner Join Moedas M" & vbCrLf &
                   "    on S.Moeda = M.Moeda_Id" & vbCrLf &
                   " Inner Join Pedidos P" & vbCrLf &
                   "    on S.Empresa_id    = P.Empresa_Id" & vbCrLf &
                   "   and S.EndEmpresa_id = P.EndEmpresa_id" & vbCrLf &
                   "   and S.Pedido        = P.Pedido_id" & vbCrLf &
                   "  Left join Pagamentos Pg" & vbCrLf &
                   "    on Pg.Pagamento_Id = P.CondicaoPagamento" & vbCrLf &
                   " WHERE 1 = 1" & vbCrLf

            If chkSaldoQuantidade.Checked Then
                Sql &= " And Case When (Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' AND AND S.QuantidadeItemPedido = 0) Then (EntregueNota - DevolucaoNota) Else (QuantidadeSoma - EntregueNota + DevolucaoNota) End > 0 " & vbCrLf
                'Sql &= " And Case When (Classe = '"& eClassesOperacoes.DEPOSITOS.ToString &"' OR Classe = '"& eClassesOperacoes.AFIXAR.ToString &"') Then (EntregueNota - DevolucaoNota) Else (Quantidade - EntregueNota + DevolucaoNota) End > 0 " & vbCrLf
            End If

            If chkSaldoValor.Checked Then
                Sql &= " And Case When (Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' AND AND S.QuantidadeItemPedido = 0) Then (EntregueNotaValor - DevolucaoNotaValor) Else (totaloficial - EntregueNotaValor + DevolucaoNotaValor) End > 0 " & vbCrLf
                'Sql &= " And Case When (Classe = '"& eClassesOperacoes.Depositos.ToString &"' OR Classe = '"& eClassesOperacoes.AFIXAR.ToString &"') Then (EntregueNotaValor - DevolucaoNotaValor) Else (totaloficial - EntregueNotaValor + DevolucaoNotaValor) End > 0 " & vbCrLf
            End If

            Sql &= " GROUP BY case" & vbCrLf &
                   "             when S.Troca      = 1   then 'Troca'" & vbCrLf &
                   "             When S.Avista is null and s.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' then 'A Fixar'" & vbCrLf &
                   "             When S.Avista is null   then 'Nao Informado'" & vbCrLf &
                   "             When S.Financeiro = 'N' then 'Sem Financeiro'" & vbCrLf &
                   "             When S.Avista     = 0   then 'A Prazo'" & vbCrLf &
                   "             When S.Avista     = 1   then 'A Vista'" & vbCrLf &
                   "          End," & vbCrLf &
                   "          M.Descricao + ' - ' + M.Cifrao, EntradaSaida" & vbCrLf

            Return Sql ' "Resumo Moeda Produto"
        End If

        Return Nothing
    End Function

    Protected Sub cmbSafra_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSafra.SelectedIndexChanged
        Try
            If cmbSafra.SelectedIndex > 0 Then
                ckConsiderarPeriodo.Enabled = True
                ckConsiderarPeriodo.Checked = False
            Else
                ckConsiderarPeriodo.Enabled = False
                ckConsiderarPeriodo.Checked = True
            End If
            txtPeriodoInicial.Parent.Visible = ckConsiderarPeriodo.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlMarca_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlMarca.SelectedIndexChanged
        Try
            If ddlMarca.SelectedIndex = 0 Then
                ucSelecaoProduto.WhereProduto = ""
            Else
                ucSelecaoProduto.WhereProduto = "marca = " & ddlMarca.SelectedValue
            End If

            ucSelecaoProduto.CarregarNivel(1)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio("PDF")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio("EXCEL")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcelDados_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcelDados.Click
        Try
            EmitirRelatorio("DADOS")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Tipo As String)
        Try
            If Funcoes.VerificaPermissao("RelacaoDePedidos", "RELATORIO") Then
                If String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
                    MsgBox(Me.Page, "Empresa é obrigatório.")
                    Exit Sub
                End If

                Dim Parametros As String = ""
                Dim sql As String = ""
                sql = getSql("Base", Parametros)
                Dim ds As DataSet = New DataSet()

                If rdPorProduto.Checked Then
                    Parametros &= "Relatório por Produto "
                    sql &= getSql("PorProduto")
                    ds = Banco.ConsultaDataSet(sql, "Consulta")
                    ds.Tables(0).TableName = "RelacaoDePedidos"
                    ds.Tables(1).TableName = "ResumoMoedaProduto"
                    ds.Tables(2).TableName = "ResumoMoedaRelatorio"
                ElseIf rdPorCliente.Checked Then
                    Parametros &= "Relatório por Cliente "
                    sql &= getSql("PorCliente")
                    ds = Banco.ConsultaDataSet(sql, "Consulta")
                    ds.Tables(0).TableName = "RelacaoDePedidos"
                    ds.Tables(1).TableName = "ResumoProduto"
                    ds.Tables(2).TableName = "ResumoMoedaCliente"
                ElseIf rdPorPedido.Checked Then
                    Parametros &= "Relatório por Pedido "
                    sql &= getSql("PorPedido")
                    ds = Banco.ConsultaDataSet(sql, "Consulta")
                    ds.Tables(0).TableName = "RelacaoDePedidos"
                    ds.Tables(1).TableName = "ResumoRelacaoPedidos"
                End If

                Dim CrName As String = ""
                If rdPorProduto.Checked Then
                    CrName = "Cr_RelacaoDePedidosRevendaProduto"
                Else
                    If rdPorCliente.Checked Then
                        If radPeso.Checked Then
                            CrName = "Cr_RelacaoDePedidosRevenda"
                        Else
                            CrName = "Cr_RelacaoDePedidosValorRevenda"
                        End If
                    Else
                        If radPeso.Checked Then
                            CrName = "Cr_RelacaoDePedidos"
                        Else
                            CrName = "Cr_RelacaoDePedidosValor"
                        End If
                    End If
                End If

                Dim pEmpresa As String() = DdlEmpresaConsultaTitulos.SelectedValue.Split("-")
                Dim Empresa As New [Lib].Negocio.Cliente(pEmpresa(0), pEmpresa(1))

                Dim texto As String = "Relação de Pedidos"
                If String.IsNullOrWhiteSpace(txtPeriodoInicial.Text) AndAlso String.IsNullOrWhiteSpace(txtPeriodoFinal.Text) AndAlso ckConsiderarPeriodo.Checked Then
                    texto = texto & " - Periodo de " & txtPeriodoInicial.Text & " a " & txtPeriodoFinal.Text
                End If

                Dim param As New Dictionary(Of String, Object)()
                param.Add("Parametros", Parametros)
                param.Add("OmitirValores", chkOmitirValores.Checked)
                param.Add("OmitirResumoProduto", chkOmitirResumoProduto.Checked)
                param.Add("TituloRelatorio", texto)
                If Tipo = "EXCEL" Then
                    EmitirExcel(ds)
                ElseIf Tipo = "DADOS" Then
                    EmitirExcelDados(ds)
                Else
                    Funcoes.BindReport(Me.Page, ds, CrName, IIf(Tipo = "PDF", eExportType.PDF, eExportType.ExcelCrystal), param)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirExcel(ByRef dsExcel As DataSet)
        Try
            Dim objEmpresa As New [Lib].Negocio.Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Dim dt As DataTable = dsExcel.Tables(0)
            dt.Columns.Remove("EmpresaNome")
            dt.Columns.Remove("EmpresaReduzido")
            dt.Columns.Remove("EmpresaCidade")
            dt.Columns.Remove("UnitarioMoeda")
            dt.Columns.Remove("ValorMoeda")
            dt.Columns.Remove("FaturadoNaoEntregue")
            dt.Columns.Remove("DevolucaoNota")
            dt.Columns.Remove("DevolucaoNotaValor")
            dt.Columns.Remove("Complemento")

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Relação de pedidos.")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "Relação de pedidos.")
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    Dim TipoRel As String = String.Empty

                    If rbEntradas.Checked Then
                        TipoRel = "ENTRADAS"
                    ElseIf rbSaidas.Checked Then
                        TipoRel = "SAÍDAS"
                    Else
                        TipoRel = "ENTRADAS/SAÍDAS"
                    End If

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtPeriodoInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtPeriodoFinal.Text)) & " - " & TipoRel)
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "DATA : " & String.Format(Now(), "dd/MM/yyyy"))
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In dt.Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A6:H" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In dt.Rows
                        columnIndex = 1
                        For Each col As DataColumn In dt.Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                        'formatando células numéricas
                        worksheet.Cells(String.Format("W{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000" ' quantidade
                        worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000" ' quantidade
                        worksheet.Cells(String.Format("Y{0}", rowIndex)).Style.Numberformat.Format = "#,########0.00000000_ ;[Red]-#,########0.00000000" ' unitário
                        worksheet.Cells(String.Format("Z{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00" ' valor
                        worksheet.Cells(String.Format("AA{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00" ' valor
                        worksheet.Cells(String.Format("AB{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00" ' valor
                        worksheet.Cells(String.Format("AC{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000" ' quantidade
                        worksheet.Cells(String.Format("AD{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000" ' quantidade
                        worksheet.Cells(String.Format("AE{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00" ' valor
                        worksheet.Cells(String.Format("AF{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000" ' quantidade 

                        '"#,##0.00_ ;[Red]-#,##0.00" ' quantidade ' unitário
                        worksheet.Cells(String.Format("AG{0}", rowIndex)).Style.Numberformat.Format = "#,########0.00000000_ ;[Red]-#,########0.00000000" ' unitário
                        worksheet.Cells(String.Format("AH{0}", rowIndex)).Style.Numberformat.Format = "#,########0.00000000_ ;[Red]-#,########0.00000000" ' unitário
                        worksheet.Cells(String.Format("AI{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00" ' valor
                        worksheet.Cells(String.Format("AJ{0}", rowIndex)).Style.Numberformat.Format = "#,########0.00000000_ ;[Red]-#,########0.00000000" ' unitário
                        worksheet.Cells(String.Format("AK{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00" ' valor

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("U{0}:AN{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("U{0}:AN{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("U{0}:AN{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("U{0}:AN{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("U{0}:AN{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(String.Format("U{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    'criando colunas de totalizadores "QuantidadeItemPedido"
                    worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("X{0}", rowIndex)).Formula = String.Format("=SUM(V6:V{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("X6:X{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"

                    'criando colunas de totalizadores "Valor"
                    worksheet.Cells(String.Format("AJ{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("AJ{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("AJ{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("AJ{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("AJ{0}", rowIndex)).Formula = String.Format("=SUM(AJ6:AJ{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("AJ6:AJ{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"

                    'criando colunas de totalizadores "ValorLiquido"
                    worksheet.Cells(String.Format("AL{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("AL{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("AL{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("AL{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("AL{0}", rowIndex)).Formula = String.Format("=SUM(AL6:AL{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("AL6:AL{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            'download do arquivo pelo browser
            'Download(fileName)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirExcelDados(ByRef dsExcel As DataSet)
        Try
            Dim objEmpresa As New [Lib].Negocio.Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Dim dt As DataTable = dsExcel.Tables(0)
            dt.Columns.Remove("EmpresaNome")
            dt.Columns.Remove("EmpresaReduzido")
            dt.Columns.Remove("EmpresaCidade")
            dt.Columns.Remove("UnitarioMoeda")
            dt.Columns.Remove("ValorMoeda")
            dt.Columns.Remove("FaturadoNaoEntregue")
            dt.Columns.Remove("DevolucaoNota")
            dt.Columns.Remove("DevolucaoNotaValor")
            dt.Columns.Remove("Complemento")

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Relação de pedidos.")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "Relação de pedidos.")
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    Dim TipoRel As String = String.Empty

                    If rbEntradas.Checked Then
                        TipoRel = "ENTRADAS"
                    ElseIf rbSaidas.Checked Then
                        TipoRel = "SAÍDAS"
                    Else
                        TipoRel = "ENTRADAS/SAÍDAS"
                    End If

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtPeriodoInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtPeriodoFinal.Text)) & " - " & TipoRel)
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "DATA : " & String.Format(Now(), "dd/MM/yyyy"))
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In dt.Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A6:H" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In dt.Rows
                        columnIndex = 1
                        For Each col As DataColumn In dt.Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                        'formatando células numéricas
                        worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000" ' quantidade
                        worksheet.Cells(String.Format("Y{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000" ' quantidade
                        worksheet.Cells(String.Format("Z{0}", rowIndex)).Style.Numberformat.Format = "#,########0.00000000_ ;[Red]-#,########0.00000000" ' unitário
                        worksheet.Cells(String.Format("AA{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00" ' valor
                        worksheet.Cells(String.Format("AB{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00" ' valor
                        worksheet.Cells(String.Format("AC{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00" ' valor
                        worksheet.Cells(String.Format("AD{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000" ' quantidade
                        worksheet.Cells(String.Format("AE{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000" ' quantidade
                        worksheet.Cells(String.Format("AF{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00" ' valor
                        worksheet.Cells(String.Format("AG{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000" ' quantidade 

                        '"#,##0.00_ ;[Red]-#,##0.00" ' quantidade ' unitário
                        worksheet.Cells(String.Format("AH{0}", rowIndex)).Style.Numberformat.Format = "#,########0.00000000_ ;[Red]-#,########0.00000000" ' unitário
                        worksheet.Cells(String.Format("AI{0}", rowIndex)).Style.Numberformat.Format = "#,########0.00000000_ ;[Red]-#,########0.00000000" ' unitário
                        worksheet.Cells(String.Format("AJ{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00" ' valor
                        worksheet.Cells(String.Format("AK{0}", rowIndex)).Style.Numberformat.Format = "#,########0.00000000_ ;[Red]-#,########0.00000000" ' unitário
                        worksheet.Cells(String.Format("AL{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00" ' valor

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("U{0}:AN{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("U{0}:AN{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("U{0}:AN{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("U{0}:AN{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("U{0}:AN{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(String.Format("U{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    'criando colunas de totalizadores "Quantidade"
                    worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("X{0}", rowIndex)).Formula = String.Format("=SUM(X6:X{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("X6:X{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"
                    ' Alinhamento à direita
                    worksheet.Cells(String.Format("X6:X{0}", rowIndex)).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right

                    'criando colunas de totalizadores "QuantidadeItemPedido"
                    worksheet.Cells(String.Format("Y{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("Y{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("Y{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("Y{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("Y{0}", rowIndex)).Formula = String.Format("=SUM(Y6:Y{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("Y6:Y{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"
                    ' Alinhamento à direita
                    worksheet.Cells(String.Format("Y6:Y{0}", rowIndex)).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right


                    'criando colunas de totalizadores "Valor"
                    worksheet.Cells(String.Format("AJ{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("AJ{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("AJ{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("AJ{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("AJ{0}", rowIndex)).Formula = String.Format("=SUM(AJ6:AJ{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("AJ6:AJ{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"
                    ' Alinhamento à direita
                    worksheet.Cells(String.Format("AJ6:AJ{0}", rowIndex)).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right


                    'criando colunas de totalizadores "ValorLiquido"
                    worksheet.Cells(String.Format("AL{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("AL{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("AL{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("AL{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("AL{0}", rowIndex)).Formula = String.Format("=SUM(AL6:AL{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("AL6:AL{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"
                    ' Alinhamento à direita
                    worksheet.Cells(String.Format("AL6:AL{0}", rowIndex)).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            'download do arquivo pelo browser
            'Download(fileName)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelacaoDePedidos")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ckConsiderarPeriodo_CheckedChanged(sender As Object, e As EventArgs) Handles ckConsiderarPeriodo.CheckedChanged
        txtPeriodoFinal.Parent.Visible = ckConsiderarPeriodo.Checked
    End Sub

End Class