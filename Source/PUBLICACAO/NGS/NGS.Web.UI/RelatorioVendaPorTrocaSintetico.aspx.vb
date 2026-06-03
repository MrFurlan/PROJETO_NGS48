Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioVendaPorTrocaSintetico
    Inherits BasePage

    Private objCliente As [Lib].Negocio.Cliente

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioVendaPorTrocaSintetico", "ACESSAR") Then
                    Limpar()

                    ddl.Carregar(cmbEmpresaOrigem, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
                    ddl.Carregar(cmbEmpresaDestino, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
                    ddl.Carregar(ddlSafraOrigem, CarregarDDL.Tabela.Safra, "", True)
                    ddl.Carregar(ddlSafraDestino, CarregarDDL.Tabela.Safra, "", True)

                    ddl.Carregar(ddlMoedaVenda, CarregarDDL.Tabela.Moeda, "", True)
                    ddl.Carregar(ddlMoedaCompra, CarregarDDL.Tabela.Moeda, "", True)

                    txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
                    txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub





    Private Function ValidarCampos() As Boolean
        If cmbEmpresaOrigem.SelectedIndex = 0 And cmbEmpresaDestino.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione empresa origem ou destino")
            Return False
        ElseIf txtDataInicial.Text.Length = 0 Or IsDate(txtDataInicial.Text) = False Then
            MsgBox(Me.Page, "Informe uma data inicial válida")
            txtDataInicial.Focus()
            Return False
        ElseIf txtDataFinal.Text.Length = 0 Or IsDate(txtDataFinal.Text) = False Then
            MsgBox(Me.Page, "Informe uma data final válida")
            txtDataFinal.Focus()
            Return False
        ElseIf CDate(txtDataInicial.Text) > CDate(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Data inicial não pode ser maior que a data final")
            txtDataInicial.Focus()
            Return False
        End If
        Return True
    End Function

    Protected Sub cmdBuscaClienteOrigem_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HttpContext.Current.Session("ssCampo") = "Livre"
            hdnControlePopup.Value = "Compra"
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTituloDIV("Consulta de Clientes")
            Popup.ConsultaDeClientes(Me.Page, "objVendaPorTrocaSintetico" & HID.Value.ToString)
            Popup.CenterDialog(Me.Page, "divConsultaCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaClienteDestino_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HttpContext.Current.Session("ssCampo") = "Livre"
            hdnControlePopup.Value = "Venda"
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTituloDIV("Consulta de Clientes")
            Popup.ConsultaDeClientes(Me.Page, "objVendaPorTrocaSintetico" & HID.Value.ToString)
            Popup.CenterDialog(Me.Page, "divConsultaCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        If Funcoes.VerificaPermissao("RelatorioVendaPorTrocaSintetico", "RELATORIO") Then
            Dim sqlProd As String = ""
            Dim DescricaoProduto As String = ""
            Dim dsProduto As New DataSet
            Dim ParametrosVenda As String = "<-- Venda -->" & vbCrLf
            Dim ParametrosCompra As String = "<-- Compra -->" & vbCrLf
            Dim PPeriodo As String = ""

            Dim ds As New DataSet
            Dim SQL As String
            Dim Cliente As String = ""
            Dim strEmpresaOrigem() As String = cmbEmpresaOrigem.SelectedValue.Split("-")
            Dim strEmpresaDestino() As String = cmbEmpresaDestino.SelectedValue.Split("-")
            Dim strClienteOrigem() As String = txtCodigoClienteOrigem.Value.Split(";")
            Dim strClienteDestino() As String = txtCodigoClienteDestino.Value.Split(";")


            'Moeda exibicao
            If rdOficial.Checked Then
                ParametrosVenda &= "Moeda de Exibição Oficial" & vbCrLf
                ParametrosCompra &= "Moeda de Exibição Oficial" & vbCrLf
            Else
                ParametrosVenda &= "Moeda de Exibição Estrangeira" & vbCrLf
                ParametrosCompra &= "Moeda de Exibição Estrangeira" & vbCrLf
            End If


            Dim Where As String = ""
            Where = " Select distinct P.Empresa_id, P.EndEmpresa_Id, P.Pedido_Id" & vbCrLf & _
                    "   into #Pedidos" & vbCrLf & _
                    "   From Pedidos P" & vbCrLf & _
                     "  INNER JOIN PedidoXItem AS PxI" & vbCrLf & _
                    " 	  ON P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf & _
                    "	 AND P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf & _
                    "	 AND P.Pedido_Id     = PxI.Pedido_Id" & vbCrLf & _
                    "  INNER JOIN Produtos AS Pd" & vbCrLf & _
                    "     ON PxI.Produto_Id  = Pd.Produto_Id" & vbCrLf & _
                    "  Inner Join Operacoes Op" & vbCrLf & _
                    "     ON Op.Operacao_Id     = P.Operacao" & vbCrLf & _
                    "  Where P.Situacao        =  1 " & vbCrLf & _
                    "    and op.Classe         = '" & eClassesOperacoes.VENDAS.ToString & "' " & vbCrLf & _
                    "    And isnull(p.troca,0) = 1" & vbCrLf

            If chkPeriodo.Checked Then
                PPeriodo = "Periodo de: " & txtDataInicial.Text.ToSqlDate() & " a " & txtDataFinal.Text.ToSqlDate()
                Where &= " and P.DataEntrega BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf
            End If

            If strEmpresaOrigem(0).Length > 0 Then
                ParametrosVenda &= "Empresa: " & cmbEmpresaOrigem.SelectedItem.Text & vbCrLf
                Where &= " AND P.Empresa_Id    = '" & strEmpresaOrigem(0) & "'" & vbCrLf
                Where &= " AND P.EndEmpresa_Id = " & strEmpresaOrigem(1) & " " & vbCrLf
            End If

            If strClienteOrigem(0).Length > 0 Then
                If chkConsolidaClienteVenda.Checked Then
                    ParametrosVenda &= "Cliente Consolidado"
                    Where &= " AND left(P.Cliente,8) = '" & strClienteOrigem(0).ToString.Substring(0, 8) & "'" & vbCrLf
                Else
                    Where &= " AND P.Cliente= '" & strClienteOrigem(0) & "'" & vbCrLf
                    Where &= " AND P.EndCliente= " & strClienteOrigem(1) & " " & vbCrLf
                End If
                ParametrosVenda &= "Cliente: " & txtClienteOrigem.Text & vbCrLf
            End If

            If ucSelecaoProdutoVenda.TemSelecionado Then
                Dim ret As New ArrayList
                ret = ucSelecaoProdutoVenda.GetSqlEParametrosRelatorio("Pd.Grupo", "PxI.Produto_Id", "", True)
                Where &= " and " & ret(0) & vbCrLf
                ParametrosVenda &= ret(1)
            End If

            If ddlSafraOrigem.SelectedIndex > 0 Then
                ParametrosVenda &= "Safra: " & ddlSafraOrigem.SelectedItem.Text & vbCrLf
                Where &= " AND P.Safra = '" & ddlSafraOrigem.SelectedValue & "'" & vbCrLf
            End If

            If ddlMoedaVenda.SelectedIndex > 0 Then
                ParametrosVenda &= "Moeda dos pedidos: " & ddlMoedaVenda.SelectedItem.Text & vbCrLf
                Where &= " and P.Moeda = " & ddlMoedaVenda.SelectedValue & vbCrLf
            End If

            If strEmpresaDestino(0).Length > 0 Or strClienteDestino(0).Length > 0 Or ucSelecaoProdutoVenda.TemSelecionado Or ddlSafraDestino.SelectedIndex > 0 Or ddlMoedaCompra.SelectedIndex > 0 Then
                Where &= " And P.Pedido_Id in (Select P.PedidoTroca" & vbCrLf & _
                       "						From Pedidos P" & vbCrLf & _
                       "					   INNER JOIN PedidoXItem AS PxI" & vbCrLf & _
                       " 					      ON P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf & _
                       "						 AND P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf & _
                       "						 AND P.Pedido_Id     = PxI.Pedido_Id" & vbCrLf & _
                       "					   INNER JOIN Produtos AS Pd" & vbCrLf & _
                       "						  ON PxI.Produto_Id  = Pd.Produto_Id" & vbCrLf & _
                       "                       Inner Join SubOperacoes SO" & vbCrLf & _
                       "                          ON SO.Operacao_Id     = P.Operacao" & vbCrLf & _
                       "                         And SO.SubOperacoes_Id = P.SubOperacao" & vbCrLf & _
                       "				       Where P.Situacao        =  1 " & vbCrLf & _
                       "                         And isnull(P.Troca,0) = 1" & vbCrLf

                If strEmpresaDestino(0).Length > 0 Then
                    ParametrosCompra &= "Empresa: " & cmbEmpresaDestino.SelectedItem.Text & vbCrLf
                    Where &= " AND P.Empresa_Id    ='" & strEmpresaDestino(0) & "'" & vbCrLf
                    Where &= " AND P.EndEmpresa_Id = " & strEmpresaDestino(1) & " " & vbCrLf
                End If

                If strClienteDestino(0).Length > 0 Then
                    If chkConsolidarClienteCompra.Checked Then
                        ParametrosCompra &= "Cliente Consolidado"
                        Where &= " AND left(P.Cliente,8)    ='" & strClienteDestino(0).ToString.Substring(0, 8) & "'" & vbCrLf
                    Else
                        Where &= " AND P.Cliente    ='" & strClienteDestino(0) & "'" & vbCrLf
                        Where &= " AND P.EndCliente = " & strClienteDestino(1) & " " & vbCrLf
                    End If
                    ParametrosCompra &= "Cliente: " & txtClienteDestino.Text & vbCrLf
                End If


                If ucSelecaoProdutoCompra.TemSelecionado Then
                    Dim ret As New ArrayList
                    ret = ucSelecaoProdutoCompra.GetSqlEParametrosRelatorio("Pd.Grupo", "PxI.Produto_Id", "", True)
                    Where &= " and " & ret(0) & vbCrLf
                    ParametrosCompra &= ret(1)
                End If

                If ddlSafraDestino.SelectedIndex > 0 Then
                    ParametrosCompra &= "Safra: " & ddlSafraDestino.SelectedItem.Text & vbCrLf
                    Where &= " AND P.Safra = '" & ddlSafraDestino.SelectedValue & "'" & vbCrLf
                End If

                If ddlMoedaCompra.SelectedIndex > 0 Then
                    ParametrosCompra &= "Moeda dos pedidos: " & ddlMoedaCompra.SelectedItem.Text & vbCrLf
                    Where &= " and P.Moeda = " & ddlMoedaCompra.SelectedValue & vbCrLf
                End If

                Where &= ")"
            End If
            ParametrosVenda &= PPeriodo


            ' --****************************  COMPRA  *******************************
            SQL = Where & ";" & vbCrLf & _
                  " SELECT PV.Pedido_id," & vbCrLf & _
                  "        PC.Cliente, " & vbCrLf & _
                  "        PC.EndCliente, " & vbCrLf & _
                  "        CC.Nome, " & vbCrLf & _
                  "        PIC.Produto_Id, " & vbCrLf & _
                  "        PrdC.Nome as NomeProduto,  " & vbCrLf & _
                  "        sum(PIC.Quantidade) AS Quantidade," & vbCrLf & _
                  "        sum(" & IIf(rdMoeda.Checked, "PEC.ValorMoeda", "PEC.ValorOficial") & ") AS ValorOficial, " & vbCrLf & _
                  "        sum(ISNULL(NFI.EntregueNota, 0)) AS EntregueNota, " & vbCrLf & _
                  "        sum(ISNULL(NFI.EntregueNotaValor, 0)" & IIf(rdMoeda.Checked, "/ case when M.classificacao = 'O' or PV.IndiceFixado = 0 then co.indice else PV.IndiceFixado end", "") & ")  AS EntregueNotaValor, " & vbCrLf & _
                  "        convert(Numeric(13,10),0)  AS Percentual  " & vbCrLf & _
                  "   into #Compra" & vbCrLf & _
                  "   FROM Pedidos AS PV   " & vbCrLf & _
                  "  INNER JOIN Pedidos AS PC   " & vbCrLf & _
                  "     ON PV.EmpresaTroca    = PC.Empresa_Id  " & vbCrLf & _
                  "    AND PV.EndEmpresaTroca = PC.EndEmpresa_Id " & vbCrLf & _
                  "    AND PV.PedidoTroca     = PC.Pedido_Id    " & vbCrLf & _
                  "  INNER JOIN (Select Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id," & vbCrLf & _
                  "					    SUM(CASE " & vbCrLf & _
                  "						      WHEN TipoDeLancamento = 'E'" & vbCrLf & _
                  "							    THEN Quantidade * - 1 " & vbCrLf & _
                  "							    ELSE Quantidade " & vbCrLf & _
                  "						    END) AS Quantidade" & vbCrLf & _
                  "                from PedidoXItemXLancamento" & vbCrLf & _
                  "               Group by Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id" & vbCrLf & _
                  "              ) AS PIC    " & vbCrLf & _
                  "     ON PC.Empresa_Id    = PIC.Empresa_Id    " & vbCrLf & _
                  "    AND PC.EndEmpresa_Id = PIC.EndEmpresa_Id " & vbCrLf & _
                  "    AND PC.Pedido_Id     = PIC.Pedido_Id    " & vbCrLf & _
                  "  INNER JOIN Clientes AS CC    " & vbCrLf & _
                  "     ON PC.Cliente    = CC.Cliente_Id    " & vbCrLf & _
                  "    AND PC.EndCliente = CC.Endereco_Id   " & vbCrLf & _
                  "  INNER JOIN Produtos AS PrdC  " & vbCrLf & _
                  "     ON PIC.Produto_Id  = PrdC.Produto_Id " & vbCrLf & _
                  "  INNER JOIN PedidosXEncargos AS PEC   " & vbCrLf & _
                  "     ON PIC.Empresa_Id     = PEC.Empresa_Id   " & vbCrLf & _
                  "    AND PIC.EndEmpresa_Id  = PEC.EndEmpresa_Id " & vbCrLf & _
                  "    AND PIC.Pedido_Id      = PEC.Pedido_Id    " & vbCrLf & _
                  "    AND PIC.Produto_Id     = PEC.Produto_Id   " & vbCrLf & _
                  "    AND PEC.Encargo_Id     ='" & IIf(rdLiquido.Checked, "LIQUIDO", "PRODUTO") & "'" & vbCrLf & _
                  "  LEFT JOIN (SELECT NotasFiscais.Pedido, " & vbCrLf & _
                  "                   SUM(case" & vbCrLf & _
                  "                          When SubOperacoes.Devolucao = 'N'" & vbCrLf & _
                  "                            then ISNULL(NotasFiscaisXItens.QuantidadeFisica, 0) " & vbCrLf & _
                  "                            else ISNULL(NotasFiscaisXItens.QuantidadeFisica, 0) * -1 " & vbCrLf & _
                  "                        End " & vbCrLf & _
                  "                        ) AS EntregueNota, " & vbCrLf & _
                  "                    SUM(case " & vbCrLf & _
                  "                          When SubOperacoes.Devolucao = 'N' " & vbCrLf & _
                  "                            then ISNULL(NotasFiscaisXEncargos.Valor, 0) " & vbCrLf & _
                  "                            else ISNULL(NotasFiscaisXEncargos.Valor, 0) * -1 " & vbCrLf & _
                  "                        End " & vbCrLf & _
                  "                        ) AS EntregueNotaValor " & vbCrLf & _
                  "              FROM NotasFiscaisXItens  " & vbCrLf & _
                  "             INNER JOIN NotasFiscais  " & vbCrLf & _
                  "                ON NotasFiscaisXItens.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf & _
                  "               AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf & _
                  "               AND NotasFiscaisXItens.Cliente_Id      = NotasFiscais.Cliente_Id " & vbCrLf & _
                  "               AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf & _
                  "               AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf & _
                  "               AND NotasFiscaisXItens.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf & _
                  "               AND NotasFiscaisXItens.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf & _
                  "             INNER JOIN NotasFiscaisXEncargos" & vbCrLf & _
                  "                ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id " & vbCrLf & _
                  "               AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id " & vbCrLf & _
                  "               AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id " & vbCrLf & _
                  "               AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id " & vbCrLf & _
                  "               AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id " & vbCrLf & _
                  "               AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id " & vbCrLf & _
                  "               AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id " & vbCrLf & _
                  "               AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id " & vbCrLf & _
                  "               AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id " & vbCrLf & _
                  "               AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf & _
                  "               AND NotasFiscaisXEncargos.Encargo_Id   = '" & IIf(rdLiquido.Checked, "LIQUIDO", "PRODUTO") & "'" & vbCrLf & _
                  "              LEFT JOIN SubOperacoes  " & vbCrLf & _
                  "                ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id  " & vbCrLf & _
                  "               AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf & _
                  "             Where SubOperacoes.Classe not in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.COMPLEMENTACOES.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "') " & vbCrLf & _
                  "               and Notasfiscais.situacao = 1 " & vbCrLf & _
                  "             GROUP BY NotasFiscais.Pedido" & vbCrLf & _
                  "          ) AS NFI  " & vbCrLf & _
                  "     ON NFI.Pedido     = PC.Pedido_Id " & vbCrLf & _
                  "  Inner Join Moedas M" & vbCrLf & _
                  "     on M.Moeda_id = PV.Moeda" & vbCrLf & _
                  "  Inner Join Cotacoes Co" & vbCrLf & _
                  "     on Co.indexador_id = 2" & vbCrLf & _
                  "    and Co.Data_id      ='" & Date.Now.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                  "  Inner Join (Select P1.EmpresaTroca as Empresa_id, P1.EndEmpresaTroca as EndEmpresa_id, P1.PedidoTroca as Pedido_Id" & vbCrLf & _
                  "                From Pedidos P1" & vbCrLf & _
                  "               inner join #Pedidos tP1" & vbCrLf & _
                  "                  on P1.Empresa_id    = tP1.Empresa_id" & vbCrLf & _
                  "                 and P1.EndEmpresa_id = tP1.EndEmpresa_id" & vbCrLf & _
                  "                 and P1.Pedido_id     = tP1.Pedido_id" & vbCrLf & _
                  "              ) tP2" & vbCrLf & _
                  "     on PC.Empresa_id    = tP2.Empresa_id" & vbCrLf & _
                  "    and PC.EndEmpresa_id = tP2.EndEmpresa_id" & vbCrLf & _
                  "    and PC.Pedido_id     = tP2.Pedido_id" & vbCrLf & _
                  "  Group By PV.Pedido_id, " & vbCrLf & _
                  "        PC.Cliente, " & vbCrLf & _
                  "        PC.EndCliente, " & vbCrLf & _
                  "        CC.Nome, " & vbCrLf & _
                  "        PIC.Produto_Id, " & vbCrLf & _
                  "        PrdC.Nome " & vbCrLf

            '----************************** % COMPRA % *******************************
            SQL &= " Update #Compra set " & vbCrLf & _
                   "   Percentual = case when sb.TotalOficial = 0" & vbCrLf & _
                   "                       then 0" & vbCrLf & _
                   "                       else (valorOficial*100)/ sb.TotalOficial" & vbCrLf & _
                   "                end" & vbCrLf & _
                   "        From ( Select Pedido_id, " & vbCrLf & _
                   "         Cliente, " & vbCrLf & _
                   "         EndCliente, " & vbCrLf & _
                   "         Sum(ValorOficial) as TotalOficial " & vbCrLf & _
                   "              from #Compra " & vbCrLf & _
                   "              Group By Pedido_id, " & vbCrLf & _
                   "              Cliente, " & vbCrLf & _
                   "             EndCliente " & vbCrLf & _
                   "          ) as sb " & vbCrLf & _
                   "   Inner Join #Compra " & vbCrLf & _
                   "      on #Compra.Cliente    = sb.Cliente " & vbCrLf & _
                   "     and #Compra.EndCliente = sb.EndCliente " & vbCrLf & _
                   "     and #Compra.Pedido_Id  = sb.Pedido_Id  " & vbCrLf


            '----*****************************  VENDA  *******************************
            SQL &= " SELECT PV.Pedido_Id, " & vbCrLf & _
                   "        PV.Cliente, " & vbCrLf & _
                   "        PV.EndCliente, " & vbCrLf & _
                   "        CV.Nome, " & vbCrLf & _
                   "        GE.Grupo_Id, " & vbCrLf & _
                   "        GE.Descricao, " & vbCrLf & _
                   "        sum(" & IIf(rdMoeda.Checked, "PEV.ValorMoeda", "PEV.ValorOficial") & ")   AS ValorOficial,  " & vbCrLf & _
                   "        sum(ISNULL(NFI.EntregueNotaValor, 0) " & IIf(rdMoeda.Checked, "/ case when M.classificacao = 'O' or PV.IndiceFixado = 0 then co.indice else PV.IndiceFixado end", "") & ")  AS EntregueNotaValor, " & vbCrLf & _
                   "        convert(Numeric(13,10),0)              AS Percentual,   " & vbCrLf & _
                   "        sum(PIv.Quantidade)                    AS Quantidade " & vbCrLf & _
                   "   INTO #Venda " & vbCrLf & _
                   "   FROM Pedidos AS PV   " & vbCrLf & _
                   "  INNER JOIN (Select PxI.Empresa_Id," & vbCrLf & _
                   "                     PxI.EndEmpresa_Id," & vbCrLf & _
                   "                     PxI.Pedido_Id," & vbCrLf & _
                   "                     PxI.Produto_id," & vbCrLf & _
                   "                     sum(case" & vbCrLf & _
                   "                           When PxI.TipoDeLancamento = 'E'" & vbCrLf & _
                   "                              then Quantidade * -1" & vbCrLf & _
                   "                              else Quantidade" & vbCrLf & _
                   "                           end) Quantidade" & vbCrLf & _
                   "                from PedidoXItemXLancamento PxI" & vbCrLf & _
                   "               Group by PxI.Empresa_Id, PxI.EndEmpresa_Id, PxI.Pedido_Id, PxI.Produto_id" & vbCrLf & _
                   "               ) PIV" & vbCrLf & _
                   "      ON PV.Empresa_Id    = PIV.Empresa_Id" & vbCrLf & _
                   "     AND PV.EndEmpresa_Id = PIV.EndEmpresa_Id" & vbCrLf & _
                   "     AND PV.Pedido_Id     = PIV.Pedido_Id" & vbCrLf & _
                   "   INNER JOIN Clientes AS CV    " & vbCrLf & _
                   "      ON PV.Cliente    = CV.Cliente_Id  " & vbCrLf & _
                   "     AND PV.EndCliente = CV.Endereco_Id  " & vbCrLf & _
                   "   INNER JOIN Produtos AS PrdV    " & vbCrLf & _
                   "      ON PIV.Produto_Id  = PrdV.Produto_Id  " & vbCrLf & _
                   "   Inner Join GruposDeEstoques GE " & vbCrLf & _
                   "      on left(PrdV.Grupo,2)  = left(GE.Grupo_Id,2)   " & vbCrLf & _
                   "     and len(GE.Grupo_Id)    = 2 " & vbCrLf & _
                   "   INNER JOIN PedidosXEncargos AS PEV   " & vbCrLf & _
                   "      ON PIV.Empresa_Id     = PEV.Empresa_Id  " & vbCrLf & _
                   "     AND PIV.EndEmpresa_Id  = PEV.EndEmpresa_Id    " & vbCrLf & _
                   "     AND PIV.Pedido_Id      = PEV.Pedido_Id    " & vbCrLf & _
                   "     AND PIV.Produto_Id     = PEV.Produto_Id    " & vbCrLf & _
                   "     AND PEV.Encargo_Id     = '" & IIf(rdLiquido.Checked, "LIQUIDO", "PRODUTO") & "'" & vbCrLf & _
                   "    LEFT JOIN (SELECT NotasFiscais.Pedido,  " & vbCrLf & _
                   "                      NotasFiscaisXItens.Produto_id, " & vbCrLf & _
                   "                      SUM(case " & vbCrLf & _
                   "                            When SubOperacoes.Devolucao = 'N' " & vbCrLf & _
                   "                              then ISNULL(NotasFiscaisXItens.QuantidadeFisica, 0) " & vbCrLf & _
                   "                              else ISNULL(NotasFiscaisXItens.QuantidadeFisica, 0) * - 1 " & vbCrLf & _
                   "                          end " & vbCrLf & _
                   "                          ) AS EntregueNota,  " & vbCrLf & _
                   "                      SUM(case " & vbCrLf & _
                   "                            When SubOperacoes.Devolucao = 'N' " & vbCrLf & _
                   "                              then ISNULL(NotasFiscaisXEncargos.Valor, 0) " & vbCrLf & _
                   "                              else ISNULL(NotasFiscaisXEncargos.Valor, 0) * -1 " & vbCrLf & _
                   "                          end " & vbCrLf & _
                   "                         ) AS EntregueNotaValor " & vbCrLf & _
                   "               FROM NotasFiscais   " & vbCrLf & _
                   "              INNER JOIN NotasFiscaisXItens  " & vbCrLf & _
                   "                 ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id  " & vbCrLf & _
                   "                AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id  " & vbCrLf & _
                   "                AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                   "                AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id  " & vbCrLf & _
                   "                AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id  " & vbCrLf & _
                   "                AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id  " & vbCrLf & _
                   "                AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id  " & vbCrLf & _
                   "              INNER JOIN NotasFiscaisXEncargos" & vbCrLf & _
                   "                 ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id " & vbCrLf & _
                   "                AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id " & vbCrLf & _
                   "                AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id " & vbCrLf & _
                   "                AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id " & vbCrLf & _
                   "                AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id " & vbCrLf & _
                   "                AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id " & vbCrLf & _
                   "                AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id " & vbCrLf & _
                   "                AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id " & vbCrLf & _
                   "                AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id " & vbCrLf & _
                   "                AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf & _
                   "                AND NotasFiscaisXEncargos.Encargo_Id   = '" & IIf(rdLiquido.Checked, "LIQUIDO", "PRODUTO") & "'" & vbCrLf & _
                   "               LEFT JOIN SubOperacoes  " & vbCrLf & _
                   "                 ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id  " & vbCrLf & _
                   "                AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf & _
                   "              Where SubOperacoes.Classe not in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.COMPLEMENTACOES.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "') " & vbCrLf & _
                   "                and Notasfiscais.situacao = 1 " & vbCrLf & _
                   "              GROUP BY NotasFiscais.Pedido, NotasFiscaisXItens.Produto_id " & vbCrLf & _
                   "            ) AS NFI  " & vbCrLf & _
                   "     ON NFI.Pedido     = PV.Pedido_Id " & vbCrLf & _
                   "    AND NFI.Produto_id = PIV.Produto_Id  " & vbCrLf & _
                   "  Inner Join Moedas M" & vbCrLf & _
                   "     on M.Moeda_id = PV.Moeda" & vbCrLf & _
                   "  Inner Join Cotacoes Co" & vbCrLf & _
                   "     on Co.indexador_id = 2" & vbCrLf & _
                   "    and Co.Data_id      ='" & Date.Now.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                   "  Inner Join #Pedidos P1" & vbCrLf & _
                   "     on PV.Empresa_id    = P1.Empresa_id" & vbCrLf & _
                   "    and PV.EndEmpresa_id = P1.EndEmpresa_id" & vbCrLf & _
                   "    and PV.Pedido_id     = P1.Pedido_id" & vbCrLf & _
                   "  Group By PV.Pedido_Id, " & vbCrLf & _
                   "         PV.Cliente, " & vbCrLf & _
                   "         PV.EndCliente, " & vbCrLf & _
                   "         CV.Nome, " & vbCrLf & _
                   "         GE.Grupo_Id, " & vbCrLf & _
                   "         GE.Descricao " & vbCrLf

            '----************************** % VENDA % *******************************
            SQL &= " Update #Venda set " & vbCrLf & _
                   "   Percentual = case when sb.TotalOficial = 0" & vbCrLf & _
                   "                       then 0" & vbCrLf & _
                   "                       else (valorOficial*100)/ sb.TotalOficial" & vbCrLf & _
                   "                end" & vbCrLf & _
                   "   from (Select Pedido_Id, " & vbCrLf & _
                   "         Cliente, " & vbCrLf & _
                   "         EndCliente, " & vbCrLf & _
                   "                Sum(ValorOficial) as TotalOficial " & vbCrLf & _
                   "           from #Venda " & vbCrLf & _
                   "          Group By Pedido_Id, " & vbCrLf & _
                   "         Cliente, " & vbCrLf & _
                   "         EndCliente " & vbCrLf & _
                   "         ) as sb " & vbCrLf & _
                   "   Inner Join #Venda " & vbCrLf & _
                   "      on #Venda.Cliente    = sb.Cliente " & vbCrLf & _
                   "     and #Venda.EndCliente = sb.EndCliente " & vbCrLf & _
                   "     and #Venda.Pedido_id  = sb.Pedido_Id  " & vbCrLf & _
                   " SELECT Row_Number() over(PARTITION BY V.Pedido_id order by V.Pedido_Id) sequencia, " & vbCrLf & _
                   "        V.Pedido_Id, " & vbCrLf & _
                   "        V.Cliente, " & vbCrLf & _
                   "        V.EndCliente, " & vbCrLf & _
                   "        V.Nome, " & vbCrLf & _
                   "        V.Grupo_Id, " & vbCrLf & _
                   "        V.Descricao, " & vbCrLf & _
                   "        convert(numeric(18,2),(V.ValorOficial      * C.Percentual)/100) as ValorOficialGrupo, " & vbCrLf & _
                   "        convert(numeric(18,2),(V.EntregueNotaValor * C.Percentual)/100) as ValorEntregue, " & vbCrLf & _
                   "        C.Produto_id, " & vbCrLf & _
                   "        C.NomeProduto, " & vbCrLf & _
                   "        convert(numeric(18,2),(C.Quantidade        * V.Percentual)/100) as Quantidade, " & vbCrLf & _
                   " 	    convert(numeric(18,2),(C.ValorOficial      * V.Percentual)/100) as ValorOficial, " & vbCrLf & _
                   " 	    convert(numeric(18,2),(C.EntregueNota      * V.Percentual)/100) as EntregueNota, " & vbCrLf & _
                   " 	    convert(numeric(18,2),(C.EntregueNotaValor * V.Percentual)/100) as EntregueNotaValor, " & vbCrLf & _
                   "        convert(numeric(18,2),(V.Quantidade        * V.Percentual)/100) as QuantidadeV " & vbCrLf & _
                   "   Into #PreSelect " & vbCrLf & _
                   "   FROM #Venda V   " & vbCrLf & _
                   "  Inner Join #Compra C " & vbCrLf & _
                   "     on V.Pedido_Id  = C.Pedido_Id " & vbCrLf
            '"    and V.Cliente    = C.Cliente " & vbCrLf & _
            '"    and V.EndCliente = C.EndCliente " & vbCrLf & _

            SQL &= " Update #PreSelect set " & vbCrLf & _
                   "    #PreSelect.ValorOficialGrupo = #PreSelect.ValorOficialGrupo + SQ.ValorOficialGrupo " & vbCrLf & _
                   "   ,#PreSelect.ValorEntregue     = #PreSelect.ValorEntregue     + SQ.ValorEntregueGrupo " & vbCrLf & _
                   "   ,#PreSelect.Quantidade        = #PreSelect.Quantidade        + SQ.Quantidade " & vbCrLf & _
                   "   ,#PreSelect.ValorOficial      = #PreSelect.ValorOficial      + SQ.ValorOficial  " & vbCrLf & _
                   "   ,#PreSelect.EntregueNota      = #PreSelect.EntregueNota      + SQ.EntregueNota " & vbCrLf & _
                   "   ,#PreSelect.EntregueNotaValor = #PreSelect.EntregueNotaValor + SQ.EntregueNotaValor  " & vbCrLf & _
                   "  From ( " & vbCrLf & _
                   "        Select PS.Pedido_Id, " & vbCrLf & _
                   "               VD.ValorOficial      - sum(PS.ValorOficialGrupo) as ValorOficialGrupo, " & vbCrLf & _
                   "               VD.ValorEntregue     - sum(PS.ValorEntregue)     as ValorEntregueGrupo, " & vbCrLf & _
                   "               CP.Quantidade        - sum(PS.Quantidade)        as Quantidade, " & vbCrLf & _
                   " 	            CP.ValorOficial      - sum(PS.ValorOficial)      as ValorOficial, " & vbCrLf & _
                   " 	            CP.EntregueNota      - sum(PS.EntregueNota)      as EntregueNota, " & vbCrLf & _
                   " 	            CP.EntregueNotaValor - sum(PS.EntregueNotaValor) as EntregueNotaValor " & vbCrLf & _
                   "          from #PreSelect PS " & vbCrLf & _
                   "         Inner Join ( select Pedido_id, " & vbCrLf & _
                   "                             sum(ValorOficial)      as ValorOficial, " & vbCrLf & _
                   "                             sum(EntregueNotaValor) as ValorEntregue " & vbCrLf & _
                   "                        from #Venda  " & vbCrLf & _
                   "                       group by Pedido_id " & vbCrLf & _
                   "                     ) VD  " & vbCrLf & _
                   "            on PS.Pedido_Id  = VD.Pedido_Id " & vbCrLf & _
                   "         Inner Join ( select Pedido_id, " & vbCrLf & _
                   "                             sum(Quantidade)        as Quantidade, " & vbCrLf & _
                   "                             sum(ValorOficial)      as ValorOficial, " & vbCrLf & _
                   "                             sum(EntregueNota)      as EntregueNota, " & vbCrLf & _
                   "                             sum(EntregueNotaValor) as EntregueNotaValor  " & vbCrLf & _
                   "                        from #Compra  " & vbCrLf & _
                   "                       group by Pedido_id " & vbCrLf & _
                   "                     ) CP  " & vbCrLf & _
                   "            on PS.Pedido_Id  = CP.Pedido_Id " & vbCrLf & _
                   "         Group By PS.Pedido_Id,  " & vbCrLf & _
                   " 				 VD.ValorOficial, " & vbCrLf & _
                   " 				 VD.ValorEntregue, " & vbCrLf & _
                   " 				 CP.Quantidade, " & vbCrLf & _
                   " 				 CP.ValorOficial, " & vbCrLf & _
                   " 				 CP.EntregueNota, " & vbCrLf & _
                   " 				 CP.EntregueNotaValor " & vbCrLf & _
                   "       ) SQ " & vbCrLf & _
                   "  Inner Join #PreSelect " & vbCrLf & _
                   "     on #PreSelect.Pedido_Id  = SQ.Pedido_Id " & vbCrLf & _
                   "  Where #PreSelect.Sequencia  = 1 " & vbCrLf

            If rbCliente.Checked Then
                SQL &= " Select Cliente, " & vbCrLf & _
                       "        EndCliente, " & vbCrLf & _
                       "        Nome, " & vbCrLf & _
                       "        Grupo_Id, " & vbCrLf & _
                       "        Descricao, " & vbCrLf & _
                       "        sum(ValorOficialGrupo) as ValorOficialGrupo, " & vbCrLf & _
                       "        sum(ValorEntregue) as ValorEntregue, " & vbCrLf & _
                       "        Produto_id, " & vbCrLf & _
                       "        NomeProduto, " & vbCrLf & _
                       "        sum(Quantidade) as Quantidade, " & vbCrLf & _
                       " 	    sum(ValorOficial) as ValorOficial, " & vbCrLf & _
                       " 	    sum(EntregueNota) as EntregueNota, " & vbCrLf & _
                       " 	    sum(EntregueNotaValor) as EntregueNotaValor  " & vbCrLf & _
                       "   From #PreSelect " & vbCrLf & _
                       "  group by Cliente, EndCliente, Nome, Grupo_Id, Descricao, Produto_id, NomeProduto " & vbCrLf

            ElseIf rbGrupo.Checked Then
                SQL &= " Select Grupo_Id, " & vbCrLf & _
                "         Descricao, " & vbCrLf & _
                "         sum(ValorOficialGrupo) as ValorOficialGrupo, " & vbCrLf & _
                "         sum(ValorEntregue) as ValorEntregue, " & vbCrLf & _
                "         Produto_id, " & vbCrLf & _
                "         NomeProduto, " & vbCrLf & _
                "         sum(Quantidade) as Quantidade, " & vbCrLf & _
                " 	      sum(ValorOficial) as ValorOficial, " & vbCrLf & _
                " 	      sum(EntregueNota) as EntregueNota, " & vbCrLf & _
                " 	      sum(EntregueNotaValor) as EntregueNotaValor " & vbCrLf & _
                "   From #PreSelect " & vbCrLf & _
                "  group by Grupo_Id, Descricao, Produto_id, NomeProduto " & vbCrLf

            ElseIf rbPedido.Checked Then

                SQL &= " Select P.Pedido_Id, P.Cliente, " & vbCrLf & _
                                   "        P.EndCliente, " & vbCrLf & _
                                   "        P.Nome, " & vbCrLf & _
                                   "        P.Grupo_Id, " & vbCrLf & _
                                   "        P.Descricao, " & vbCrLf & _
                                   "        sum(P.ValorOficialGrupo) as ValorOficialGrupo, " & vbCrLf & _
                                   "        sum(P.ValorEntregue) as ValorEntregue, " & vbCrLf & _
                                   "        sum(P.QuantidadeV) as QuantidadeV,  " & vbCrLf & _
                                   "        P.Produto_id, " & vbCrLf & _
                                   "        P.NomeProduto, " & vbCrLf & _
                                   "        sum(P.Quantidade) as Quantidade, " & vbCrLf & _
                                   " 	    sum(P.ValorOficial) as ValorOficial, " & vbCrLf & _
                                   " 	    sum(P.EntregueNota) as EntregueNota, " & vbCrLf & _
                                   " 	    sum(P.EntregueNotaValor) as EntregueNotaValor,  " & vbCrLf & _
                                   " 	    Pxp.PedidoDestino_Id as PedidoDestino  " & vbCrLf & _
                                   "   From #PreSelect P " & vbCrLf & _
                                   "   LEFT JOIN PedidoXPedido Pxp" & vbCrLf & _
                                   "     ON P.Pedido_Id = Pxp.PedidoOrigem_Id" & vbCrLf & _
                                   "  group by P.Cliente, P.EndCliente, P.Nome, P.Grupo_Id, P.Descricao, P.Produto_id, P.NomeProduto, P.Pedido_Id, Pxp.PedidoDestino_Id " & vbCrLf

            End If

            ds = Banco.ConsultaDataSet(SQL, "RelatorioVendasPorTrocaSintetico")

            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("Grupo", "Grupo Prod.")
            parameters.Add("Cliente", "Cliente")
            parameters.Add("Produto", "Produto")
            parameters.Add("ValorEntregue", "Vlr.Entregue")
            parameters.Add("Quantidade", "Quantidade")
            parameters.Add("ValorOficial", "Vlr.Oficial.")
            parameters.Add("EntregaNota", "Entr.Nota")
            parameters.Add("EntregaValorNota", "Ent. Vlr Nota")
            parameters.Add("Titulo", "Relatório de Vendas Por Troca Sintético " & IIf(rbCliente.Checked, "Cliente", "Grupo"))
            parameters.Add("Pagina", "Página")
            parameters.Add("Data", "Emissão")
            parameters.Add("ValorOficialGrupo", "Vlr.Ofic.Grupo")
            parameters.Add("ParametrosVenda", ParametrosVenda & IIf(rdBruto.Checked, "Encargo Produto/Bruto", "Encargo Liquido"))
            parameters.Add("ParametrosCompra", ParametrosCompra & IIf(rdBruto.Checked, "Encargo Produto/Bruto", "Encargo Liquido"))

            Funcoes.BindReport(Me.Page, ds, IIf(rbCliente.Checked, "Cr_VendasPorTrocaSinteticoCli", IIf(rbGrupo.Checked, "Cr_VendasPorTrocaSinteticoGru", "Cr_VendasPorTrocaSinteticoPed")), IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
        Else
            MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
        End If
    End Sub

    Protected Sub chkPeriodo_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkPeriodo.CheckedChanged
        Try
            pnlData.Visible = chkPeriodo.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objVendaPorTrocaSintetico" & HID.Value.ToString) Is Nothing Then
            objCliente = CType(Session("objVendaPorTrocaSintetico" & HID.Value.ToString), [Lib].Negocio.Cliente)
            Dim Cliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            Select Case hdnControlePopup.Value.Replace(",", "")
                Case "Compra"
                    txtClienteOrigem.Text = Cliente.Text
                    txtCodigoClienteOrigem.Value = objCliente.Codigo & ";" & objCliente.CodigoEndereco
                Case "Venda"
                    txtClienteDestino.Text = Cliente.Text
                    txtCodigoClienteDestino.Value = objCliente.Codigo & ";" & objCliente.CodigoEndereco
            End Select
        End If
    End Sub

    Public Sub Limpar()
        Session.Remove("objVendaPorTrocaSintetico" & HID.Value)
        txtClienteOrigem.Text = ""
        txtCodigoClienteOrigem.Value = ""
        txtClienteDestino.Text = ""
        txtCodigoClienteDestino.Value = ""
        ddlSafraOrigem.SelectedIndex = 0
        ddlSafraDestino.SelectedIndex = 0
        txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
        cmbEmpresaOrigem.SelectedIndex = 0
        cmbEmpresaDestino.SelectedIndex = 0
        HID.Value = Guid.NewGuid.ToString
        ucConsultaClientes.SetarHID(HID.Value)
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioVendaPorTrocaSintetico")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class