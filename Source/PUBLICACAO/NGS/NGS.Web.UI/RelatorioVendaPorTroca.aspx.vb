Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class RelatorioVendaPorTroca
    Inherits BasePage

    Private dateEmissao As New DateTime
    Private intPagina As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RelatorioVendaPorTroca", "ACESSAR") Then
                ddl.Carregar(cmbEmpresaOrigem, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
                ddl.Carregar(cmbEmpresaDestino, CarregarDDL.Tabela.ClientesXEmpresas, "", True)

                ddl.Carregar(ddlSafraOrigem, CarregarDDL.Tabela.Safra, "", True)
                ddl.Carregar(ddlSafraDestino, CarregarDDL.Tabela.Safra, "", True)

                ddl.Carregar(ddlMoedaVenda, CarregarDDL.Tabela.Moeda, "", True)
                ddl.Carregar(ddlMoedaCompra, CarregarDDL.Tabela.Moeda, "", True)

                txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
                txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
                txtPedido.Text = ""
                HID.Value = Guid.NewGuid().ToString
                ucConsultaClientes.SetarHID(HID.Value)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Session("objClienteOrigem" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteOrigem" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtClienteOrigem.Text = itemCliente.Text
            txtCodigoClienteOrigem.Value = itemCliente.Value
            Session.Remove("objClienteOrigem" & HID.Value)
        End If

        If Session("objClienteDestino" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteDestino" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtClienteDestino.Text = itemCliente.Text
            txtCodigoClienteDestino.Value = itemCliente.Value
            Session.Remove("objClienteDestino" & HID.Value)
        End If
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
            ucConsultaClientes.SetarHID(HID.Value)
            Popup.ConsultaDeClientes(Me.Page, "objClienteOrigem" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaClienteDestino_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HttpContext.Current.Session("ssCampo") = "Livre"
            ucConsultaClientes.SetarHID(HID.Value)
            Popup.ConsultaDeClientes(Me.Page, "objClienteDestino" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub



    Public Sub EmitirRelatorio(ByVal pdf As Boolean)
        If Funcoes.VerificaPermissao("RelatorioVendaPorTroca", "RELATORIO") Then
            Dim sqlProd As String = ""
            Dim DescricaoProduto As String = ""
            Dim dsProduto As New DataSet
            Dim ParametrosVenda As String = "<-- Venda --> "
            Dim ParametrosCompra As String = "<-- Compra --> "
            Dim PPeriodo As String = ""

            Dim ds As New DataSet
            Dim sql As String
            Dim Cliente As String = ""
            Dim strEmpresaOrigem() As String = cmbEmpresaOrigem.SelectedValue.Split("-")
            Dim strEmpresaDestino() As String = cmbEmpresaDestino.SelectedValue.Split("-")
            'Dim strClienteOrigem() As String = txtCodigoClienteOrigem.Value.Split(";")
            'Dim strClienteDestino() As String = txtCodigoClienteDestino.Value.Split(";")

            Dim strClienteOrigem() As String = txtCodigoClienteOrigem.Value.Split("-")
            Dim strClienteDestino() As String = txtCodigoClienteDestino.Value.Split("-")

            'Moeda exibicao
            If rdOficial.Checked Then
                ParametrosVenda &= "Moeda de Exibição Oficial" & vbCrLf
                ParametrosCompra &= "Moeda de Exibição Oficial" & vbCrLf
            Else
                ParametrosVenda &= "Moeda de Exibição Estrangeira" & vbCrLf
                ParametrosCompra &= "Moeda de Exibição Estrangeira" & vbCrLf
            End If

            sql = "  Select distinct P.Pedido_Id" & vbCrLf & _
                    "    into #Pedidos" & vbCrLf & _
                    "    From Pedidos P" & vbCrLf & _
                    "   INNER JOIN PedidoXItem AS PxI" & vbCrLf & _
                    " 	   ON P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf & _
                    "	  AND P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf & _
                    "	  AND P.Pedido_Id     = PxI.Pedido_Id" & vbCrLf & _
                    "   INNER JOIN Produtos AS Pd" & vbCrLf & _
                    "      ON PxI.Produto_Id  = Pd.Produto_Id" & vbCrLf & _
                    "   Inner Join Operacoes Op" & vbCrLf & _
                    "      ON Op.Operacao_Id    = P.Operacao" & vbCrLf & _
                    "   Where P.Situacao        =  1" & vbCrLf & _
                    "     And Op.Classe         = '" & eClassesOperacoes.VENDAS.ToString & "'" & vbCrLf & _
                    "     And isnull(P.Troca,0) = 1" & vbCrLf

            If txtPedido.Text.Trim <> "" Then
                ParametrosVenda &= "Pedido: " & txtPedido.Text & vbCrLf
                sql &= " AND P.Pedido_Id   = " & txtPedido.Text & vbCrLf
            ElseIf txtPedidoCompra.Text.Trim <> "" Then
                ParametrosCompra &= "Pedido: " & txtPedido.Text & vbCrLf
                sql &= " AND P.PedidoTroca =  " & txtPedido.Text & vbCrLf
            Else
                If ChkPeriodo.Checked Then
                    PPeriodo = "Periodo de: " & txtDataInicial.Text.ToSqlDate() & " a " & txtDataFinal.Text.ToSqlDate()
                    sql &= " and P.DataEntrega BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf
                End If

                If strEmpresaOrigem(0).Length > 0 Then
                    ParametrosVenda &= "Empresa: " & cmbEmpresaOrigem.SelectedItem.Text & vbCrLf
                    sql &= " AND P.Empresa_Id    = '" & strEmpresaOrigem(0) & "'" & vbCrLf
                    sql &= " AND P.EndEmpresa_Id = " & strEmpresaOrigem(1) & " " & vbCrLf
                End If

                If strClienteOrigem(0).Length > 0 Then
                    If chkConsClienteVenda.Checked Then
                        ParametrosVenda &= "Cliente Consolidado"
                        sql &= " AND left(P.Cliente,8) = '" & strClienteOrigem(0).ToString.Substring(0, 8) & "'" & vbCrLf
                    Else
                        sql &= " AND P.Cliente= '" & strClienteOrigem(0) & "'" & vbCrLf
                        sql &= " AND P.EndCliente= " & strClienteOrigem(1) & " " & vbCrLf
                    End If
                    ParametrosVenda &= "Cliente: " & txtClienteOrigem.Text & vbCrLf
                End If

                If ucSelecaoProdutoVenda.TemSelecionado Then
                    Dim ret As New ArrayList
                    ret = ucSelecaoProdutoVenda.GetSqlEParametrosRelatorio("Pd.Grupo", "PxI.Produto_Id", "", True)
                    sql &= " and " & ret(0) & vbCrLf
                    ParametrosVenda &= ret(1)
                End If

                If ddlSafraOrigem.SelectedIndex > 0 Then
                    ParametrosVenda &= "Safra: " & ddlSafraOrigem.SelectedItem.Text & vbCrLf
                    sql &= " AND P.Safra = '" & ddlSafraOrigem.SelectedValue & "'" & vbCrLf
                End If

                If ddlMoedaVenda.SelectedIndex > 0 Then
                    ParametrosVenda &= "Moeda dos pedidos: " & ddlMoedaVenda.SelectedItem.Text & vbCrLf
                    sql &= " and P.Moeda = " & ddlMoedaVenda.SelectedValue & vbCrLf
                End If

                If strEmpresaDestino(0).Length > 0 Or strClienteDestino(0).Length > 0 Or ucSelecaoProdutoCompra.TemSelecionado Or ddlSafraDestino.SelectedIndex > 0 Or ddlMoedaCompra.SelectedIndex > 0 Then
                    sql &= " And P.Pedido_Id in (Select P.PedidoTroca" & vbCrLf & _
                           "					   From Pedidos P" & vbCrLf & _
                           "				      INNER JOIN PedidoXItem AS PxI" & vbCrLf & _
                           " 				         ON P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf & _
                           "				  	    AND P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf & _
                           "				  	    AND P.Pedido_Id     = PxI.Pedido_Id" & vbCrLf & _
                           "				      INNER JOIN Produtos AS Pd" & vbCrLf & _
                           "				  	     ON PxI.Produto_Id  = Pd.Produto_Id" & vbCrLf & _
                           "                      Inner Join SubOperacoes SO" & vbCrLf & _
                           "                         ON SO.Operacao_Id     = P.Operacao" & vbCrLf & _
                           "                        And SO.SubOperacoes_Id = P.SubOperacao" & vbCrLf & _
                           "				      Where P.Situacao        = 1 " & vbCrLf & _
                           "                        And isnull(P.troca,0) = 1 "


                    If strEmpresaDestino(0).Length > 0 Then
                        ParametrosCompra &= "Empresa: " & cmbEmpresaDestino.SelectedItem.Text & vbCrLf
                        sql &= "                        AND P.Empresa_Id    ='" & strEmpresaDestino(0) & "'" & vbCrLf
                        sql &= "                        AND P.EndEmpresa_Id = " & strEmpresaDestino(1) & " " & vbCrLf
                    End If

                    If strClienteDestino(0).Length > 0 Then
                        If chkConsClienteCompra.Checked Then
                            ParametrosCompra &= "Cliente Consolidado"
                            sql &= "                        AND left(P.Cliente,8)    ='" & strClienteDestino(0).ToString.Substring(0, 8) & "'" & vbCrLf
                        Else
                            sql &= "                        AND P.Cliente    ='" & strClienteDestino(0) & "'" & vbCrLf
                            sql &= "                        AND P.EndCliente = " & strClienteDestino(1) & " " & vbCrLf
                        End If
                        ParametrosCompra &= "Cliente: " & txtClienteDestino.Text & vbCrLf
                    End If

                    If ucSelecaoProdutoCompra.TemSelecionado Then
                        Dim ret As New ArrayList
                        ret = ucSelecaoProdutoCompra.GetSqlEParametrosRelatorio("Pd.Grupo", "PxI.Produto_Id", "", True)
                        sql &= "                        And " & ret(0) & vbCrLf
                        ParametrosCompra &= ret(1)
                    End If

                    If ddlSafraDestino.SelectedIndex > 0 Then
                        ParametrosCompra &= "Safra: " & ddlSafraDestino.SelectedItem.Text & vbCrLf
                        sql &= "                        AND P.Safra = '" & ddlSafraDestino.SelectedValue & "'" & vbCrLf
                    End If

                    If ddlMoedaCompra.SelectedIndex > 0 Then
                        ParametrosCompra &= "Moeda dos pedidos: " & ddlMoedaCompra.SelectedItem.Text & vbCrLf
                        sql &= " and P.Moeda = " & ddlMoedaCompra.SelectedValue & vbCrLf
                    End If

                    sql &= "					)"
                End If
                ParametrosCompra &= PPeriodo
            End If

            If RdLiquido.Checked Then
                ParametrosVenda &= "Encargo Liquido" & vbCrLf
                ParametrosCompra &= "Encargo Liquido" & vbCrLf
            Else
                ParametrosVenda &= "Encargo Produto / Bruto" & vbCrLf
                ParametrosCompra &= "Encargo Produto / Bruto" & vbCrLf
            End If

            sql &= "SELECT NFE.Empresa_Id, NFE.EndEmpresa_Id,NFE.Cliente_Id, " & vbCrLf & _
                   "       NFE.EndCliente_Id, NFE.EntradaSaida_Id," & vbCrLf & _
                   "       NFE.Serie_Id, NFE.Nota_Id," & vbCrLf & _
                   "       NFE.Produto_Id, NFE.CFOP_Id," & vbCrLf & _
                   "       NFE.Sequencia_id, NFE.Encargo_id," & vbCrLf & _
                   "       NFE.Valor" & vbCrLf & _
                   "  INTO #Encargos" & vbCrLf & _
                   "  FROM NotasFiscaisXEncargos AS NFE " & vbCrLf & _
                   "  JOIN NotasFiscais AS NF " & vbCrLf & _
                   "    ON NFE.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                   "   AND NFE.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                   "   AND NFE.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                   "   AND NFE.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                   "   AND NFE.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                   "   AND NFE.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                   "   AND NFE.Nota_Id         = NF.Nota_Id" & vbCrLf & _
                   "  LEFT JOIN SubOperacoes   " & vbCrLf & _
                   "    ON NF.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf & _
                   "   AND NF.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
                   " WHERE NFE.Encargo_Id = '" & IIf(RdLiquido.Checked, "LIQUIDO", "PRODUTO") & "'" & vbCrLf & _
                   "   AND SubOperacoes.Classe not in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.COMPLEMENTACOES.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "')" & vbCrLf & _
                   "   AND NF.Pedido IN (SELECT P.Pedido_Id FROM #Pedidos P)" & vbCrLf

            If chkValorizacao.Checked And IsDate(txtDataRef.Text) Then
                sql &= "   AND NF.Movimento <='" & CDate(txtDataRef.Text).ToString("yyyy-MM-dd") & "'"
            End If


            '*** SQL Venda -----------------------------------------------------------------------------------------------------------------------------
            sql &= "; SELECT PV.Empresa_Id as EmpresaOrigem,  " & vbCrLf & _
                   " 	    PV.EndEmpresa_Id as EndEmpresaOrigem,   " & vbCrLf & _
                   " 	    PV.Pedido_Id as PedidoOrigem,    " & vbCrLf & _
                   "        PV.Cliente, " & vbCrLf & _
                   "        PV.EndCliente, " & vbCrLf & _
                   "        CV.Nome  + ' - ' +  " & vbCrLf & _
                   "        CV.Complemento AS Nome, " & vbCrLf & _
                   "        PrdV.Produto_Id, " & vbCrLf & _
                   "        PrdV.Nome as NomeProduto, " & vbCrLf & _
                   "        PV.DataPedido, " & vbCrLf & _
                   "        PV.DataEntrega, " & vbCrLf & _
                   "        SUM(CASE " & vbCrLf & _
                   "		        WHEN PIV.TipoDeLancamento = 'E'" & vbCrLf & _
                   "			     THEN PIV.Quantidade * - 1 " & vbCrLf & _
                   "			     ELSE PIV.Quantidade " & vbCrLf & _
                   "		       END) AS Quantidade," & vbCrLf & _
                   "        PEV.ValorOficial, " & vbCrLf & _
                   "        PEV.ValorMoeda, " & vbCrLf & _
                   "        PV.Moeda, " & vbCrLf & _
                   "        PV.indicefixado," & vbCrLf & _
                   "        ISNULL(NFI.EntregueNota, 0) AS EntregueNota, " & vbCrLf & _
                   "        ISNULL(NFI.EntregueNotaValor" & IIf(rdMoeda.Checked, "/ case when M.classificacao = 'O' or PV.IndiceFixado = 0 then co.indice else PV.IndiceFixado end", "") & ",0) AS EntregueNotaValor, " & vbCrLf & _
                   "        ISNULL(NFI.DevolucaoNota, 0) AS DevolucaoNota, " & vbCrLf & _
                   "        ISNULL(NFI.DevolucaoNotaValor" & IIf(rdMoeda.Checked, "/ case when M.classificacao = 'O' or PV.IndiceFixado = 0 then co.indice else PV.IndiceFixado end", "") & ",0) AS DevolucaoNotaValor " & vbCrLf & _
                   "   Into #Venda " & vbCrLf & _
                   "   FROM Pedidos AS PV  " & vbCrLf & _
                   "  INNER JOIN PedidoXItemxLancamento AS PIV   " & vbCrLf & _
                   "     ON PV.Empresa_Id    = PIV.Empresa_Id   " & vbCrLf & _
                   "    AND PV.EndEmpresa_Id = PIV.EndEmpresa_Id   " & vbCrLf & _
                   "    AND PV.Pedido_Id     = PIV.Pedido_Id   " & vbCrLf & _
                   "  INNER JOIN Clientes AS CV   " & vbCrLf & _
                   "     ON PV.Cliente    = CV.Cliente_Id   " & vbCrLf & _
                   "    AND PV.EndCliente = CV.Endereco_Id   " & vbCrLf & _
                   "  INNER JOIN Produtos AS PrdV   " & vbCrLf & _
                   "     ON PIV.Produto_Id  = PrdV.Produto_Id   " & vbCrLf & _
                   "  INNER JOIN PedidosXEncargos AS PEV  " & vbCrLf & _
                   "     ON PIV.Empresa_Id     = PEV.Empresa_Id   " & vbCrLf & _
                   "    AND PIV.EndEmpresa_Id  = PEV.EndEmpresa_Id   " & vbCrLf & _
                   "    AND PIV.Pedido_Id      = PEV.Pedido_Id   " & vbCrLf & _
                   "    AND PIV.Produto_Id     = PEV.Produto_Id   " & vbCrLf & _
                   "    AND PEV.Encargo_Id     ='" & IIf(RdLiquido.Checked, "LIQUIDO", "PRODUTO") & "'" & vbCrLf & _
                   "   LEFT JOIN (SELECT NotasFiscais.Pedido,    " & vbCrLf & _
                   "                     NotasFiscaisXItens.Produto_id,    " & vbCrLf & _
                   "                     SUM( case   " & vbCrLf & _
                   "                            When SubOperacoes.Devolucao = 'S'   " & vbCrLf & _
                   "                              then ISNULL(NotasFiscaisXItens.QuantidadeFisica, 0)   " & vbCrLf & _
                   "                              else 0   " & vbCrLf & _
                   "                           End    " & vbCrLf & _
                   "                         ) AS DevolucaoNota,    " & vbCrLf & _
                   "                     SUM( case   " & vbCrLf & _
                   "                            When SubOperacoes.Devolucao = 'S'   " & vbCrLf & _
                   "                              then ISNULL(Enc.Valor, 0)   " & vbCrLf & _
                   "                              else 0   " & vbCrLf & _
                   "                           End   " & vbCrLf & _
                   "                         ) AS DevolucaoNotaValor,   " & vbCrLf & _
                   "                     SUM( case   " & vbCrLf & _
                   "                            When SubOperacoes.Devolucao = 'N'   " & vbCrLf & _
                   "                              then ISNULL(NotasFiscaisXItens.QuantidadeFisica, 0)   " & vbCrLf & _
                   "                              else 0   " & vbCrLf & _
                   "                           End   " & vbCrLf & _
                   "                         ) AS EntregueNota,    " & vbCrLf & _
                   "                     SUM( case   " & vbCrLf & _
                   "                            When SubOperacoes.Devolucao = 'N' " & vbCrLf & _
                   "                              then ISNULL(Enc.Valor, 0) " & vbCrLf & _
                   "                              else 0 " & vbCrLf & _
                   "                           End " & vbCrLf & _
                   "                        ) AS EntregueNotaValor " & vbCrLf & _
                   "                FROM NotasFiscaisXItens  " & vbCrLf & _
                   "               INNER JOIN NotasFiscais  " & vbCrLf & _
                   "                  ON NotasFiscaisXItens.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Cliente_Id      = NotasFiscais.Cliente_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf & _
                   "               INNER JOIN #Encargos Enc" & vbCrLf & _
                   "                  ON NotasFiscaisXItens.Empresa_Id      = Enc.Empresa_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.EndEmpresa_Id   = Enc.EndEmpresa_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Cliente_Id      = Enc.Cliente_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.EndCliente_Id   = Enc.EndCliente_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.EntradaSaida_Id = Enc.EntradaSaida_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Serie_Id        = Enc.Serie_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Nota_Id         = Enc.Nota_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Produto_Id      = Enc.Produto_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.CFOP_Id         = Enc.CFOP_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Sequencia_Id    = Enc.Sequencia_id" & vbCrLf & _
                   "                 AND Enc.Encargo_Id   = '" & IIf(RdLiquido.Checked, "LIQUIDO", "PRODUTO") & "'" & vbCrLf & _
                   "                LEFT JOIN SubOperacoes  " & vbCrLf & _
                   "                  ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id  " & vbCrLf & _
                   "                 AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf & _
                   "               Where SubOperacoes.Classe not in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.COMPLEMENTACOES.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "') " & vbCrLf & _
                   "               GROUP BY NotasFiscais.Pedido, NotasFiscaisXItens.Produto_id  " & vbCrLf & _
                   "             ) AS NFI  " & vbCrLf & _
                   "      ON NFI.Pedido     = PV.Pedido_Id " & vbCrLf & _
                   "     AND NFI.Produto_id = PIV.Produto_Id  " & vbCrLf & _
                   "   Inner Join Moedas M" & vbCrLf & _
                   "      on M.Moeda_id = PV.Moeda" & vbCrLf & _
                   "   Inner Join Cotacoes Co" & vbCrLf & _
                   "      on Co.indexador_id = 2" & vbCrLf & _
                   "     and Co.Data_id      ='" & Date.Now.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                   "   WHERE PV.Pedido_Id in (Select Pedido_id From #Pedidos)" & vbCrLf & _
                   "   Group By PV.Empresa_Id , PV.EndEmpresa_Id ,PV.Pedido_Id , PV.Cliente,PV.EndCliente, CV.Nome + ' - ' + CV.Complemento, PrdV.Produto_Id, " & vbCrLf & _
                   "         PrdV.Nome, PV.DataPedido, PV.DataEntrega, PEV.ValorOficial, PEV.ValorMoeda, PV.Moeda, PV.indicefixado, NFI.EntregueNota,NFI.EntregueNotaValor" & IIf(rdMoeda.Checked, "/case when M.classificacao = 'O' or PV.IndiceFixado = 0 then co.indice else PV.IndiceFixado end", "") & ",NFI.DevolucaoNota,NFI.DevolucaoNotaValor" & IIf(rdMoeda.Checked, "/case when M.classificacao = 'O' or PV.IndiceFixado = 0 then co.indice else PV.IndiceFixado end", "") & ";" & vbCrLf

            '*** SQL Compra -----------------------------------------------------------------------------------------------------------------------------

            sql &= "SELECT PV.Empresa_Id as EmpresaOrigem,    " & vbCrLf & _
                   "       PV.EndEmpresa_Id as EndEmpresaOrigem,    " & vbCrLf & _
                   " 	   PV.Pedido_Id as PedidoOrigem,     " & vbCrLf & _
                   "       PC.Empresa_Id, " & vbCrLf & _
                   "       PC.EndEmpresa_Id, " & vbCrLf & _
                   "       PC.Pedido_Id, " & vbCrLf & _
                   "       PC.Cliente, " & vbCrLf & _
                   "       PC.EndCliente, " & vbCrLf & _
                   "       CC.Nome + ' - ' + " & vbCrLf & _
                   "       CC.Complemento AS Nome, " & vbCrLf & _
                   "       PrdC.Produto_Id, " & vbCrLf & _
                   "       PrdC.Nome as NomeProduto, " & vbCrLf & _
                   "       PC.DataPedido, " & vbCrLf & _
                   "       PC.DataEntrega, " & vbCrLf & _
                   "       SUM(CASE " & vbCrLf & _
                   "		     WHEN PIC.TipoDeLancamento = 'E'" & vbCrLf & _
                   "			   THEN PIC.Quantidade * - 1 " & vbCrLf & _
                   "			   ELSE PIC.Quantidade " & vbCrLf & _
                   "		   END) AS Quantidade," & vbCrLf & _
                   "       PEC.ValorOficial, " & vbCrLf & _
                   "       PEC.ValorMoeda, " & vbCrLf & _
                   "       PC.Moeda, " & vbCrLf & _
                   "       PC.indicefixado," & vbCrLf & _
                   "       ISNULL(NFI.EntregueNota, 0) AS EntregueNota, " & vbCrLf & _
                   "       ISNULL(NFI.EntregueNotaValor" & IIf(rdMoeda.Checked, "/ case when M.classificacao = 'O' or PV.IndiceFixado = 0 then co.indice else PV.IndiceFixado end", "") & ",0) AS EntregueNotaValor, " & vbCrLf & _
                   "       ISNULL(NFI.DevolucaoNota, 0) AS DevolucaoNota, " & vbCrLf & _
                   "       ISNULL(NFI.DevolucaoNotaValor" & IIf(rdMoeda.Checked, "/ case when M.classificacao = 'O' or PV.IndiceFixado = 0 then co.indice else PV.IndiceFixado end", "") & ",0) AS DevolucaoNotaValor " & vbCrLf & _
                   "  Into #Compra " & vbCrLf & _
                   "  FROM Pedidos AS PV   " & vbCrLf & _
                   " INNER JOIN Pedidos AS PC   " & vbCrLf & _
                   "    ON PV.EmpresaTroca    = PC.Empresa_Id    " & vbCrLf & _
                   "   AND PV.EndEmpresaTroca = PC.EndEmpresa_Id    " & vbCrLf & _
                   "   AND PV.PedidoTroca     = PC.Pedido_Id    " & vbCrLf & _
                   " INNER JOIN PedidoXItemxLancamento AS PIC    " & vbCrLf & _
                   "    ON PC.Empresa_Id    = PIC.Empresa_Id    " & vbCrLf & _
                   "   AND PC.EndEmpresa_Id = PIC.EndEmpresa_Id    " & vbCrLf & _
                   "   AND PC.Pedido_Id     = PIC.Pedido_Id    " & vbCrLf & _
                   " INNER JOIN Clientes AS CC    " & vbCrLf & _
                   "    ON PC.Cliente    = CC.Cliente_Id    " & vbCrLf & _
                   "   AND PC.EndCliente = CC.Endereco_Id    " & vbCrLf & _
                   " INNER JOIN Produtos AS PrdC    " & vbCrLf & _
                   "    ON PIC.Produto_Id  = PrdC.Produto_Id    " & vbCrLf & _
                   " INNER JOIN PedidosXEncargos AS PEC   " & vbCrLf & _
                   "    ON PIC.Empresa_Id     = PEC.Empresa_Id    " & vbCrLf & _
                   "   AND PIC.EndEmpresa_Id  = PEC.EndEmpresa_Id    " & vbCrLf & _
                   "   AND PIC.Pedido_Id      = PEC.Pedido_Id    " & vbCrLf & _
                   "   AND PIC.Produto_Id     = PEC.Produto_Id    " & vbCrLf & _
                   "   AND PEC.Encargo_Id     = '" & IIf(RdLiquido.Checked, "LIQUIDO", "PRODUTO") & "'" & vbCrLf & _
                   "  LEFT JOIN (SELECT NotasFiscais.Pedido,   " & vbCrLf & _
                   "                    NotasFiscaisXItens.Produto_id,   " & vbCrLf & _
                   "                    SUM(case  " & vbCrLf & _
                   "                          When SubOperacoes.Devolucao = 'S'  " & vbCrLf & _
                   "                            then ISNULL(NotasFiscaisXItens.QuantidadeFisica, 0)  " & vbCrLf & _
                   "                            else 0  " & vbCrLf & _
                   "                        End) AS DevolucaoNota,   " & vbCrLf & _
                   "                    SUM(case  " & vbCrLf & _
                   "                          When SubOperacoes.Devolucao = 'S'  " & vbCrLf & _
                   "                            then ISNULL(Enc.Valor, 0)  " & vbCrLf & _
                   "                            else 0  " & vbCrLf & _
                   "                        End) AS DevolucaoNotaValor,  " & vbCrLf & _
                   "                    SUM(case  " & vbCrLf & _
                   "                          When SubOperacoes.Devolucao = 'N'  " & vbCrLf & _
                   "                            then ISNULL(NotasFiscaisXItens.QuantidadeFisica, 0)  " & vbCrLf & _
                   "                            else 0  " & vbCrLf & _
                   "                        End) AS EntregueNota,   " & vbCrLf & _
                   "                    SUM(case  " & vbCrLf & _
                   "                          When SubOperacoes.Devolucao = 'N'  " & vbCrLf & _
                   "                            then ISNULL(Enc.Valor, 0)  " & vbCrLf & _
                   "                            else 0  " & vbCrLf & _
                   "                        End) AS EntregueNotaValor  " & vbCrLf & _
                   "                FROM NotasFiscaisXItens  " & vbCrLf & _
                   "               INNER JOIN NotasFiscais   " & vbCrLf & _
                   "                  ON NotasFiscaisXItens.Empresa_Id = NotasFiscais.Empresa_Id   " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id   " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Cliente_Id = NotasFiscais.Cliente_Id  " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.EndCliente_Id = NotasFiscais.EndCliente_Id   " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id   " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Serie_Id = NotasFiscais.Serie_Id   " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Nota_Id = NotasFiscais.Nota_Id   " & vbCrLf & _
                   "               INNER JOIN #Encargos Enc" & vbCrLf & _
                   "                  ON NotasFiscaisXItens.Empresa_Id      = Enc.Empresa_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.EndEmpresa_Id   = Enc.EndEmpresa_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Cliente_Id      = Enc.Cliente_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.EndCliente_Id   = Enc.EndCliente_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.EntradaSaida_Id = Enc.EntradaSaida_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Serie_Id        = Enc.Serie_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Nota_Id         = Enc.Nota_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Produto_Id      = Enc.Produto_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.CFOP_Id         = Enc.CFOP_Id " & vbCrLf & _
                   "                 AND NotasFiscaisXItens.Sequencia_Id    = Enc.Sequencia_id" & vbCrLf & _
                   "                 AND Enc.Encargo_Id   = '" & IIf(RdLiquido.Checked, "LIQUIDO", "PRODUTO") & "'" & vbCrLf & _
                   "                LEFT JOIN SubOperacoes   " & vbCrLf & _
                   "                  ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id  " & vbCrLf & _
                   "                 AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id  " & vbCrLf & _
                   "               Where SubOperacoes.Classe not in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.COMPLEMENTACOES.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "')  " & vbCrLf & _
                   "               GROUP BY NotasFiscais.Pedido, NotasFiscaisXItens.Produto_id   " & vbCrLf & _
                   "             ) AS NFI" & vbCrLf & _
                   "      ON NFI.Pedido     = PC.Pedido_Id" & vbCrLf & _
                   "     AND NFI.Produto_id = PIC.Produto_Id" & vbCrLf & _
                   "   Inner Join Moedas M" & vbCrLf & _
                   "      on M.Moeda_id = PV.Moeda" & vbCrLf & _
                   "   Inner Join Cotacoes Co" & vbCrLf & _
                   "      on Co.indexador_id = 2" & vbCrLf & _
                   "     and Co.Data_id      ='" & Date.Now.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                   "   WHERE PC.Pedido_Id in (Select PedidoTroca From Pedidos where Pedido_Id in (Select Pedido_id from #Pedidos))" & vbCrLf & _
                   "   Group By PV.Empresa_Id, PV.EndEmpresa_Id, PV.Pedido_Id,PC.Empresa_Id,PC.EndEmpresa_Id,PC.Pedido_Id," & vbCrLf & _
                   "         PC.Cliente, PC.EndCliente, CC.Nome + ' - ' + CC.Complemento, PrdC.Produto_Id,PrdC.Nome , PC.DataPedido, PC.DataEntrega, PEC.ValorOficial, PEC.ValorMoeda, PC.indicefixado, PC.Moeda,NFI.EntregueNota,NFI.EntregueNotaValor" & IIf(rdMoeda.Checked, "/case when M.classificacao = 'O' or PV.IndiceFixado = 0 then co.indice else PV.IndiceFixado end", "") & ",NFI.DevolucaoNota,NFI.DevolucaoNotaValor" & IIf(rdMoeda.Checked, "/case when M.classificacao = 'O' or PV.IndiceFixado = 0 then co.indice else PV.IndiceFixado end", "") & ";" & vbCrLf

            '***********************************************************************************************************************************************************************************
            '************************************************     Atualiza Valor de Dolar onde os Pedidos sao em REAIS     *********************************************************************
            '***********************************************************************************************************************************************************************************

            sql &= "Update V set" & vbCrLf & _
                   "   ValorMoeda = V.ValorOficial / case" & vbCrLf & _
                   "                                   when C.indicefixado = 0" & vbCrLf & _
                   "                                     then ct.Indice" & vbCrLf & _
                   "                                     else C.indicefixado" & vbCrLf & _
                   "                                 end" & vbCrLf & _
                   " from #Venda V" & vbCrLf & _
                   "inner Join #compra C" & vbCrLf & _
                   "   on V.EmpresaOrigem    = C.EmpresaOrigem" & vbCrLf & _
                   "  and V.EndEmpresaOrigem = C.EndEmpresaOrigem" & vbCrLf & _
                   "  and V.PedidoOrigem     = C.PedidoOrigem" & vbCrLf & _
                   "inner join cotacoes ct" & vbCrLf & _
                   "   on ct.indexador_Id = 2" & vbCrLf & _
                   "  and ct.data_id      = V.DataPedido" & vbCrLf & _
                   "Where V.ValorMoeda    = 0;" & vbCrLf


            sql &= "Update C set" & vbCrLf & _
                   "   ValorMoeda = C.ValorOficial / case" & vbCrLf & _
                   "                                   when V.indicefixado = 0" & vbCrLf & _
                   "                                     then ct.Indice" & vbCrLf & _
                   "                                     else V.indicefixado" & vbCrLf & _
                   "                                 end" & vbCrLf & _
                   " from #Venda V" & vbCrLf & _
                   "inner Join #compra C" & vbCrLf & _
                   "   on V.EmpresaOrigem    = C.EmpresaOrigem" & vbCrLf & _
                   "  and V.EndEmpresaOrigem = C.EndEmpresaOrigem" & vbCrLf & _
                   "  and V.PedidoOrigem     = C.PedidoOrigem" & vbCrLf & _
                   "inner join cotacoes ct" & vbCrLf & _
                   "   on ct.indexador_Id = 2" & vbCrLf & _
                   "  and ct.data_id      = C.DataPedido" & vbCrLf & _
                   "Where C.ValorMoeda    = 0;" & vbCrLf


            '*************************************************************************************************************************************
            '************************************************     RESUMO     *********************************************************************
            '*************************************************************************************************************************************
            sql &= "Select V.EmpresaOrigem, V.EndEmpresaOrigem, V.PedidoOrigem, V.Produto_id," & vbCrLf

            If rdMoeda.Checked Then
                sql &= "       round((V.ValorMoeda / (C.ValorMoeda / C.Quantidade)) / 60,4) as SacosTroca," & vbCrLf & _
                       "       ((((V.ValorMoeda / (C.ValorMoeda / C.Quantidade)) /60)/ v.Quantidade)) MediaTrocaSC," & vbCrLf & _
                       "       ((C.ValorMoeda / C.Quantidade) * 60) MediaTrocaUN," & vbCrLf

                If Not chkValorizacao.Checked Then
                    sql &= "       0 as Valorizacao"
                ElseIf rdContratado.Checked Then
                    sql &= "      (v.valormoeda) -  ((V.ValorMoeda / (C.ValorMoeda / C.Quantidade)) / 60) * (" & Str(txtValorRef.Text) & " * 60)  as Valorizacao" & vbCrLf
                Else
                    sql &= "      (V.EntregueNotaValor - V.DevolucaoNotaValor)" & vbCrLf & _
                           "      - " & vbCrLf & _
                           "      ((" & vbCrLf & _
                           "        (V.EntregueNotaValor - V.DevolucaoNotaValor)" & vbCrLf & _
                           "         /" & vbCrLf & _
                           "        (C.ValorMoeda / C.Quantidade)" & vbCrLf & _
                           "      ) / 60) * (" & Str(txtValorRef.Text) & " * 60)  as Valorizacao" & vbCrLf
                End If
            Else
                sql &= "       round((V.ValorOficial / (C.ValorOficial / C.Quantidade)) / 60,4) as SacosTroca," & vbCrLf & _
                       "       ((((V.ValorOficial / (C.ValorOficial / C.Quantidade)) /60)/ v.Quantidade)) MediaTrocaSC," & vbCrLf & _
                       "       ((C.ValorOficial / C.Quantidade) * 60) MediaTroc,"

                If Not chkValorizacao.Checked Then
                    sql &= "       0 as Valorizacao"
                ElseIf rdContratado.Checked Then
                    sql &= "      (v.valorOficial) -  ((V.ValorOficial / (C.ValorOficial / C.Quantidade)) / 60) * (" & Str(txtValorRef.Text) & " * 60)  as Valorizacao" & vbCrLf
                Else
                    sql &= "      (V.EntregueNotaValor - V.DevolucaoNotaValor)" & vbCrLf & _
                           "      - " & vbCrLf & _
                           "      ((" & vbCrLf & _
                           "        (V.EntregueNotaValor - V.DevolucaoNotaValor)" & vbCrLf & _
                           "         /" & vbCrLf & _
                           "        (C.ValorOficial / C.Quantidade)" & vbCrLf & _
                           "      ) / 60) * (" & Str(txtValorRef.Text) & " *60)  as Valorizacao" & vbCrLf
                End If
            End If

            sql &= "   from #venda V" & vbCrLf & _
                   "  inner Join Produtos Pv" & vbCrLf & _
                   "     on v.produto_id = Pv.Produto_id" & vbCrLf & _
                   "  inner Join #compra C" & vbCrLf & _
                   "     on V.EmpresaOrigem    = C.EmpresaOrigem" & vbCrLf & _
                   "    and V.EndEmpresaOrigem = C.EndEmpresaOrigem" & vbCrLf & _
                   "    and V.PedidoOrigem     = C.PedidoOrigem" & vbCrLf & _
                   "  inner Join Produtos Pc" & vbCrLf & _
                   "     on c.produto_id = Pc.Produto_id " & vbCrLf & _
                   "  where V.Quantidade > 0 and C.Quantidade > 0;" & vbCrLf


            sql &= "select GE.Grupo_Id, GE.Descricao, V.Produto_id, V.NomeProduto, PV.Unidade as UnidadeMedida, " & vbCrLf & _
                   "       sum(V.Quantidade) as Quantidade, " & vbCrLf

            If rdMoeda.Checked Then
                sql &= "       sum(V.ValorMoeda) as ValorMoeda," & vbCrLf & _
                       "       sum((V.ValorMoeda / (C.ValorMoeda / C.Quantidade)) / 60) as SacosTroca," & vbCrLf & _
                       "       sum((V.ValorMoeda / (C.ValorMoeda / C.Quantidade)) / 60) / sum(V.Quantidade) MediaTrocaSC," & vbCrLf & _
                       "       round((sum(C.ValorMoeda) / sum(C.Quantidade)) * 60,4) MediaTrocaUN," & vbCrLf
                If Not chkValorizacao.Checked Then
                    sql &= "       0 as Valorizacao"
                ElseIf rdContratado.Checked Then
                    sql &= "      sum((v.valormoeda) - ((V.ValorMoeda / (C.ValorMoeda / C.Quantidade)) / 60) * (" & Str(txtValorRef.Text) & " * 60))  as Valorizacao" & vbCrLf
                Else
                    sql &= " sum((V.EntregueNotaValor - V.DevolucaoNotaValor)" & vbCrLf & _
                           "      - " & vbCrLf & _
                           "      ((" & vbCrLf & _
                           "        (V.EntregueNotaValor - V.DevolucaoNotaValor)" & vbCrLf & _
                           "         /" & vbCrLf & _
                           "        (C.ValorMoeda / C.Quantidade)" & vbCrLf & _
                           "      ) / 60) * (" & Str(txtValorRef.Text) & " * 60))  as Valorizacao" & vbCrLf
                End If
            Else
                sql &= "       sum(V.ValorOficial) as ValorMoeda," & vbCrLf & _
                       "       sum((V.Valoroficial / (C.ValorOficial / C.Quantidade)) / 60) as SacosTroca," & vbCrLf & _
                       "       sum((V.Valoroficial / (C.Valoroficial / C.Quantidade)) / 60) / sum(V.Quantidade) MediaTrocaSC," & vbCrLf & _
                       "       round((sum(C.Valoroficial) / sum(C.Quantidade)) * 60,4) MediaTrocaUN," & vbCrLf

                If Not chkValorizacao.Checked Then
                    sql &= "       0 as Valorizacao"
                ElseIf rdContratado.Checked Then
                    sql &= "      sum((v.valorOficial) - ((V.ValorOficial / (C.ValorOficial / C.Quantidade)) / 60) * (" & Str(txtValorRef.Text) & " * 60))  as Valorizacao" & vbCrLf
                Else
                    sql &= "  sum((V.EntregueNotaValor - V.DevolucaoNotaValor)" & vbCrLf & _
                           "      - " & vbCrLf & _
                           "      ((" & vbCrLf & _
                           "        (V.EntregueNotaValor - V.DevolucaoNotaValor)" & vbCrLf & _
                           "         /" & vbCrLf & _
                           "        (C.ValorOficial / C.Quantidade)" & vbCrLf & _
                           "      ) / 60) * (" & Str(txtValorRef.Text) & " *60))  as Valorizacao" & vbCrLf
                End If
            End If


            sql &= "  from #venda V " & vbCrLf & _
                   " Inner Join Produtos PV " & vbCrLf & _
                   "    on PV.Produto_id = V.Produto_id " & vbCrLf & _
                   " Inner Join GruposDeEstoques GE " & vbCrLf & _
                   "    on Ge.Grupo_Id = left(PV.Grupo,2) " & vbCrLf & _
                   " inner Join #compra C " & vbCrLf & _
                   "    on V.EmpresaOrigem    = C.EmpresaOrigem " & vbCrLf & _
                   "   and V.EndEmpresaOrigem = C.EndEmpresaOrigem " & vbCrLf & _
                   "   and V.PedidoOrigem     = C.PedidoOrigem " & vbCrLf & _
                   "   And V.Quantidade > 0 " & vbCrLf & _
                   "  inner Join Produtos Pc" & vbCrLf & _
                   "     on c.produto_id = Pc.Produto_id " & vbCrLf & _
                   "  where V.Quantidade > 0 and C.Quantidade > 0" & vbCrLf & _
                   " Group by GE.Grupo_Id, GE.Descricao, V.Produto_id, V.NomeProduto, PV.Unidade" & vbCrLf & _
                   " order by GE.Grupo_Id,  V.NomeProduto;" & vbCrLf

            sql &= " Select * from #Venda;" & vbCrLf & _
                   " Select * from #Compra;"

            ds = Banco.ConsultaDataSet(sql, "Consulta")
            ds.Tables("Consulta").TableName = "ResumoPedido"
            ds.Tables("Consulta1").TableName = "ResumoRelatorio"
            ds.Tables("Consulta2").TableName = "RelatorioVendasPorTroca"
            ds.Tables("Consulta3").TableName = "Compra"

            AlimentaCrptRelatorios(ds, "Cr_RelatorioVendasPorTroca", ParametrosVenda, ParametrosCompra, pdf)
        Else
            MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
        End If
    End Sub

    Public Sub AlimentaCrptRelatorios(ByVal Ds As DataSet, ByVal Caminho As String, ByVal ParametrosV As String, ByVal ParametrosC As String, ByVal pdf As Boolean)
        Try
            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("ParametrosV", ParametrosV)
            parameters.Add("ParametrosC", ParametrosC)
            parameters.Add("Nome", HttpContext.Current.Session("ssNomeEmpresa"))
            parameters.Add("Cidade", HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa"))
            parameters.Add("Zerado", "True")
            parameters.Add("Titulo", "Relatório de Vendas Por Troca")
            parameters.Add("Pagina", "Página")
            parameters.Add("Data", "Emissão")

            parameters.Add("Pedido", "Pedido")
            parameters.Add("Produto", "Produto")
            parameters.Add("Cliente", "Cliente")
            parameters.Add("Moeda", "Moeda")
            parameters.Add("Quantidade", "Quantidade")
            parameters.Add("ValorOficial", IIf(rdOficial.Checked, "Vlr.Oficial", "Vlr.Moeda"))
            parameters.Add("TotalOficial", IIf(rdOficial.Checked, "Total Ofic.", "Total Moeda"))

            parameters.Add("EntregaNota", "Entr.Nota")
            parameters.Add("DevolucaoNota", "Dev. Nota")
            parameters.Add("EntregaValorNota", "Ent. Vlr Nota")
            parameters.Add("DevolucaoValorNota", "Dev. Vlr Nota")

            parameters.Add("ValorRef", IIf(IsNumeric(txtValorRef.Text) AndAlso CDec(txtValorRef.Text) > 0, txtValorRef.Text, "0"))

            parameters.Add("TipoDeResumo", IIf(rbResumoParcial.Checked, "Parcial", "Completo"))

            Funcoes.BindReport(Me.Page, Ds, Caminho, IIf(pdf, eExportType.PDF, eExportType.ExcelCrystalDados), parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaPedido.Click
        Try
            If cmbEmpresaOrigem.SelectedIndex = -1 Or txtCodigoClienteOrigem.Value.Length = 0 Then
                MsgBox(Me.Page, "É necessário informar a empresa e/ou cliente para buscar o número do pedido")
            Else
                Dim strJScript As String = ""
                Dim strCodigoEmpresaOrigem As String() = cmbEmpresaOrigem.SelectedValue.Split("-")
                Dim strCodClienteOrigem As String() = txtCodigoClienteOrigem.Value.Split(";")

                HttpContext.Current.Session("ssCampo") = "ApenasNumeroPedido"
                HttpContext.Current.Session("ssCnpjDaEmpresa") = strCodigoEmpresaOrigem(0)
                HttpContext.Current.Session("ssEndDaEmpresa") = strCodigoEmpresaOrigem(1)
                HttpContext.Current.Session("txtCnpjDoCliente") = strCodClienteOrigem(0)
                HttpContext.Current.Session("txtEndDoCliente") = strCodClienteOrigem(1)

                'HttpContext.Current.Session("CodigoProduto") = cmbProdutoOrigem.SelectedValue
                HttpContext.Current.Session("CodigoSafra") = ddlSafraOrigem.SelectedValue

                strJScript += "var x = (screen.height / 2) - 250; "
                strJScript += "var y = (screen.width / 2) - 400; "
                strJScript += "window.open(""ConsultaPedidos.aspx?pedido=" & txtPedido.ClientID & "&tipo=v"", """", ""resizable=no, menubar=no, scrollbars=yes, width=800, height=500, top="" + x + "", left="" + y + """");"
                ScriptManager.RegisterClientScriptBlock(Me, cmdBuscaPedido.GetType(), "Extratos", strJScript, True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaPedidoCompra_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If cmbEmpresaDestino.SelectedIndex = -1 Or txtCodigoClienteDestino.Value.Length = 0 Then
                MsgBox(Me.Page, "É necessário informar a empresa e/ou cliente para buscar o número do pedido")
            Else
                Dim strJScript As String = ""
                Dim strCodigoEmpresaDestino As String() = cmbEmpresaDestino.SelectedValue.Split("-")
                Dim strCodClienteDestino As String() = txtCodigoClienteDestino.Value.Split(";")

                HttpContext.Current.Session("ssCampo") = "ApenasNumeroPedido"
                HttpContext.Current.Session("ssCnpjDaEmpresa") = strCodigoEmpresaDestino(0)
                HttpContext.Current.Session("ssEndDaEmpresa") = strCodigoEmpresaDestino(1)
                HttpContext.Current.Session("txtCnpjDoCliente") = strCodClienteDestino(0)
                HttpContext.Current.Session("txtEndDoCliente") = strCodClienteDestino(1)

                'HttpContext.Current.Session("CodigoProduto") = cmbProdutoDestino.SelectedValue
                HttpContext.Current.Session("CodigoSafra") = ddlSafraDestino.SelectedValue

                strJScript += "var x = (screen.height / 2) - 250; "
                strJScript += "var y = (screen.width / 2) - 400; "
                strJScript += "window.open(""ConsultaPedidos.aspx?pedido=" & txtPedidoCompra.ClientID & "&tipo=v"", """", ""resizable=no, menubar=no, scrollbars=yes, width=800, height=500, top="" + x + "", left="" + y + """");"
                ScriptManager.RegisterClientScriptBlock(Me, cmdBuscaPedido.GetType(), "Extratos", strJScript, True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ChkPeriodo_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ChkPeriodo.CheckedChanged
        Try
            pnlData.Visible = ChkPeriodo.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(sender As Object, e As EventArgs) Handles lnkPdf.Click
        Try
            If ValidarCampos() Then
                EmitirRelatorio(True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try
            If ValidarCampos() Then
                EmitirRelatorio(False)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            txtClienteOrigem.Text = ""
            txtCodigoClienteOrigem.Value = ""
            txtClienteDestino.Text = ""
            txtCodigoClienteDestino.Value = ""
            ddlSafraOrigem.SelectedIndex = 0
            ddlSafraDestino.SelectedIndex = 0
            txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
            txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
            txtPedido.Text = ""
            cmbEmpresaOrigem.SelectedIndex = 0
            cmbEmpresaDestino.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioVendaPorTroca")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


End Class