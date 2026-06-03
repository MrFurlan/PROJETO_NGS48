Imports System.Data
Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis


Public Class RelatorioDeNotasSemCte
    Inherits BasePage

    Dim Sql As String
    Dim ds As DataSet
    Private objPesagem As [Lib].Negocio.Pesagem

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeNotasSemCte", "ACESSAR") Then
                    BuscaEmpresa()
                    Limpar()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub Limpar()
        'txtCliente.Text = ""
        'txtCodigoCliente.Value = ""
        'TxtTransportador.Text =String.empty
        'txtCodigoTransportador.Value = String.Empty

        SetStringEmptyinControls({txtCliente, TxtTransportador, txtPedido})
        SetStringEmptyinControls({txtCodigoCliente, txtCodigoTransportador})
        rbEntrada.Checked = True
        rbSaida.Checked = False
        txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
        Funcoes.VerificaEmpresa(ddlEmpresa)
        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub BuscarPedidos()
        HttpContext.Current.Session("ssCampo" & HID.Value) = "Pedidos"
        Dim strEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")
        Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
        Dim parameters As New Dictionary(Of String, Object)
        HttpContext.Current.Session("ssTipoRetorno") = "objPedidoAxP" & HID.Value
        parameters("empresa") = strEmpresa(0)
        parameters("enderecoEmpresa") = strEmpresa(1)
        parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
        parameters("endCliente_ID") = IIf(strCliente.Length > 1, strCliente(1), "")
        ucConsultaPedidos.Limpar()
        ucConsultaPedidos.BindGridView(parameters)
        Popup.ConsultaDePedidos(Me.Page, "objPedidoAxP" & HID.Value, "txtNome")
    End Sub

    'Private Sub CarregarObjetoSessao()
    '    objPesagem = CType(Session("objPesagem" & HID.Value), [Lib].Negocio.Pesagem)
    'End Sub

    'Private Sub SalvarObjetoSessao()
    '    Session("objPesagem" & HID.Value) = objPesagem
    'End Sub

    Private Function getParam() As String
        Dim param As String = "Parametros:"
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            param &= "Empresa: " & ddlEmpresa.SelectedItem.Text.Split(".")(0) & " - " & ddlEmpresa.SelectedItem.Text.Split(" ")(1) & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            param &= "Cliente: " & txtCliente.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            param &= "Pedido: " & txtPedido.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) AndAlso Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            param &= "Perído: " & txtDataInicial.Text & " à " & txtDataFinal.Text & " - "
        End If
        param &= IIf(rbEntrada.Checked, "Ent/Sai: Entrada.", "Ent/Sai: Saida.")

        Return param
    End Function

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Session("objClienteExtrato" & HID.Value) IsNot Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteExtrato" & HID.Value), [Lib].Negocio.Cliente))
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClienteExtrato" & HID.Value)
            ElseIf Session("objClienteTRAAxL" & HID.Value) IsNot Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteTRAAxL" & HID.Value), [Lib].Negocio.Cliente))
                TxtTransportador.Text = itemCliente.Text
                txtCodigoTransportador.Value = itemCliente.Value
                Session.Remove("objClienteTRAAxL" & HID.Value)
            ElseIf Session("objPedidoAxP" & HID.Value) IsNot Nothing Then
                'CarregarObjetoSessao()
                'If CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).CodigoEmpresa <> objPesagem.CodigoEmpresa Or CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).EnderecoEmpresa <> objPesagem.EnderecoEmpresa Then
                '    MsgBox(Me.Page, "Empresa do Pedido é diferente da Empresa Fornecedora.")
                'ElseIf CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).CodigoCliente <> objPesagem.CodigoCliente Or CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).EnderecoCliente <> objPesagem.EnderecoCliente Then
                '    MsgBox(Me.Page, "Fornecedor do Pedido é diferente do Fornecedor informado.")
                'ElseIf CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).Itens.Count <> 1 Then
                '    MsgBox(Me.Page, "Pedido com mais de um item não pode ser usado para agrupamento.")
                'Else
                '    objPesagem.CodigoPedido = CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).Codigo
                '    objPesagem.CodigoTabelaDeClassificacao = CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).Itens(0).Classificacao.Codigo
                '    objPesagem.CodigoProduto = CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).Itens(0).CodigoProduto
                '    SalvarObjetoSessao()
                '    txtPedido.Text = objPesagem.CodigoPedido
                '    txtCodigoPedido.Value = objPesagem.CodigoPedido
                'End If



                txtPedido.Text = CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).Codigo
                'txtCodigoPedido.Value = CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).CodigoEmpresa

                Session.Remove("objPedidoAxP" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidarCampos() As Boolean
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
            Return False
        ElseIf txtCliente.Text.Length = 0 Then
            MsgBox(Me.Page, "Cliente não foi selecionado.")
            Return False
        End If

        Return True
    End Function

    Protected Sub BtnConsultarTransportador_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnConsultarTransportador.Click
        Try
            Session("ssCampo" & HID.Value) = "LivreClasse"
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente(eTipoCliente.Transportadores)
            Popup.ConsultaDeClientes(Me.Page, "objClienteTRAAxL" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBuscaCliente.Click
        Try
            Dim strJScript As String = ""
            HttpContext.Current.Session("ssCampo") = "Livre"
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente("4,5") 'Cliente/Fornecedor
            Popup.ConsultaDeClientes(Me.Page, "objClienteExtrato" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.")
            ElseIf txtCodigoCliente.Value.ToString().Length() = 0 Then
                MsgBox(Me.Page, "Cliente não foi selecionado.")
            Else
                BuscarPedidos()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioDeNotasSemCte", "RELATORIO") Then

                Dim sql As String = "  SELECT nf.Pedido, nf.Nota_Id, nf.Serie_Id, nf.EntradaSaida_Id, nf.Empresa_Id as Empresa_Id," & vbCrLf & _
                                    "	      emp.Nome as NomeEmpresa, nf.Cliente_Id as Cliente_Id, nf.EndCliente_Id, cli.Nome as NomeCliente," & vbCrLf & _
                                    "	      nf.Movimento, nxi.QuantidadeFiscal as Quantidade, nxi.Valor," & vbCrLf & _
                                    "	      case" & vbCrLf & _
                                    "	      	when isnull(nxt.Proprietario,'') = ''" & vbCrLf & _
                                    "	      		then ''" & vbCrLf & _
                                    "	      		else nxt.Proprietario + ' - ' + trp.Nome" & vbCrLf & _
                                    "	      end as Transportador" & vbCrLf & _
                                    "    FROM NotasFiscais AS nf" & vbCrLf & _
                                    "   INNER JOIN NotasFiscaisXItens AS nxi" & vbCrLf & _
                                    "  	   ON nf.Empresa_Id = nxi.Empresa_Id" & vbCrLf & _
                                    "  	  and nf.EndEmpresa_Id = nxi.EndEmpresa_Id" & vbCrLf & _
                                    "  	  and nf.Cliente_Id = nxi.Cliente_Id" & vbCrLf & _
                                    "  	  and nf.EndCliente_Id = nxi.EndCliente_Id" & vbCrLf & _
                                    "  	  and nf.EntradaSaida_Id = nxi.EntradaSaida_Id" & vbCrLf & _
                                    "  	  and nf.Serie_Id = nxi.Serie_Id" & vbCrLf & _
                                    "  	  and nf.Nota_Id = nxi.Nota_Id" & vbCrLf & _
                                    "   INNER JOIN Clientes cli1" & vbCrLf & _
                                    "  	   ON cli1.Cliente_Id = nf.Cliente_Id" & vbCrLf & _
                                    "  	  and cli1.Endereco_Id = nf.EndCliente_Id" & vbCrLf & _
                                    "   INNER JOIN SubOperacoes sub" & vbCrLf & _
                                    "  	   ON nxi.Operacao = sub.Operacao_Id" & vbCrLf & _
                                    "  	  and nxi.SubOperacao = sub.SubOperacoes_Id" & vbCrLf & _
                                    "    LEFT JOIN NotasFiscaisXTransportadores nxt" & vbCrLf & _
                                    " 	   ON nf.Empresa_Id = nxt.Empresa_Id" & vbCrLf & _
                                    " 	  and nf.EndEmpresa_Id = nxt.EndEmpresa_Id" & vbCrLf & _
                                    " 	  and nf.Cliente_Id = nxt.Cliente_Id" & vbCrLf & _
                                    " 	  and nf.EndCliente_Id = nxt.EndCliente_Id" & vbCrLf & _
                                    " 	  and nf.EntradaSaida_Id = nxt.EntradaSaida_Id" & vbCrLf & _
                                    " 	  and nf.Serie_Id = nxt.Serie_Id" & vbCrLf & _
                                    " 	  and nf.Nota_Id = nxt.Nota_Id" & vbCrLf & _
                                    "   INNER JOIN Clientes trp" & vbCrLf & _
                                    " 	   ON trp.Cliente_Id = nxt.Proprietario" & vbCrLf & _
                                    " 	  and trp.Endereco_Id = nxt.EndProprietario "

                'Rodrigo - 02/08/2022 - Alterado para o OUTER APPLY pois quando tem uma nota relacionada e um frete relacionado acaba dando erro
                '"    LEFT JOIN NotasXNotas nxn" & vbCrLf & _
                '" 	   on nxn.OrigemEmpresa_Id = nf.Empresa_Id" & vbCrLf & _
                '" 	  and nxn.OrigemEndEmpresa_Id = nf.EndEmpresa_Id" & vbCrLf & _
                '" 	  and nxn.OrigemCliente_Id = nf.Cliente_Id" & vbCrLf & _
                '" 	  and nxn.OrigemEndCliente_Id = nf.EndCliente_Id" & vbCrLf & _
                '" 	  and nxn.OrigemEntradaSaida_Id = nf.EntradaSaida_Id" & vbCrLf & _
                '" 	  and nxn.OrigemSerie_Id = nf.Serie_Id" & vbCrLf & _
                '" 	  and nxn.OrigemNota_Id = nf.Nota_Id" & vbCrLf & _
                '" 	  and nxn.OrigemSerie_Id <> 'REC'" & vbCrLf & _
                '" 	  and nxn.Serie_Id <> 'REC'" & vbCrLf & _
                '"    LEFT JOIN NotasFiscais cte" & vbCrLf & _
                '" 	   ON nxn.Empresa_Id = cte.Empresa_Id" & vbCrLf & _
                '" 	  and nxn.EndEmpresa_Id = cte.EndEmpresa_Id" & vbCrLf & _
                '" 	  and nxn.Cliente_Id = cte.Cliente_Id" & vbCrLf & _
                '" 	  and nxn.EndCliente_Id = cte.EndCliente_Id" & vbCrLf & _
                '" 	  and nxn.EntradaSaida_Id = cte.EntradaSaida_Id" & vbCrLf & _
                '" 	  and nxn.Serie_Id = cte.Serie_Id" & vbCrLf & _
                '" 	  and nxn.Nota_Id = cte.Nota_Id" & vbCrLf & _
                '" 	  and cte.TipoDeDocumento in(2,57) " & vbCrLf & _
                '" 	  and cte.Serie_Id <> 'REC'" & vbCrLf & _
                '"    LEFT JOIN NotasFiscaisXItens ctexi" & vbCrLf & _
                '"      ON cte.Empresa_Id = ctexi.Empresa_Id" & vbCrLf & _
                '"     and cte.EndEmpresa_Id = ctexi.EndEmpresa_Id" & vbCrLf & _
                '"     and cte.Cliente_Id = ctexi.Cliente_Id" & vbCrLf & _
                '"     and cte.EndCliente_Id = ctexi.EndCliente_Id" & vbCrLf & _
                '"     and cte.EntradaSaida_Id = ctexi.EntradaSaida_Id" & vbCrLf & _
                '"     and cte.Serie_Id = ctexi.Serie_Id" & vbCrLf & _
                '"     and cte.Nota_Id = ctexi.Nota_Id" & vbCrLf & 

                sql &= " OUTER APPLY (  SELECT COUNT(cte.Nota_Id) as Quantidade "
                sql &= ""
                sql &= "                FROM NotasXNotas AS nxn "
                sql &= ""
                sql &= "                INNER JOIN NotasFiscais cte ON nxn.Empresa_Id = cte.Empresa_Id "
                sql &= "                AND nxn.EndEmpresa_Id = cte.EndEmpresa_Id "
                sql &= "                AND nxn.Cliente_Id = cte.Cliente_Id "
                sql &= "                AND nxn.EndCliente_Id = cte.EndCliente_Id "
                sql &= "                AND nxn.EntradaSaida_Id = cte.EntradaSaida_Id "
                sql &= "                AND nxn.Serie_Id = cte.Serie_Id "
                sql &= "				AND nxn.Nota_Id = cte.Nota_Id "
                sql &= "				AND cte.TipoDeDocumento in(2,57) "
                sql &= "                AND cte.Serie_Id <> 'REC' "
                sql &= ""
                sql &= "                WHERE nxn.OrigemEmpresa_Id = nf.Empresa_Id "
                sql &= "				AND nxn.OrigemEndEmpresa_Id = nf.EndEmpresa_Id "
                sql &= "				AND nxn.OrigemCliente_Id = nf.Cliente_Id "
                sql &= "				AND nxn.OrigemEndCliente_Id = nf.EndCliente_Id "
                sql &= "				AND nxn.OrigemEntradaSaida_Id = nf.EntradaSaida_Id "
                sql &= "				AND nxn.OrigemSerie_Id = nf.Serie_Id "
                sql &= "				AND nxn.OrigemNota_Id = nf.Nota_Id "
                sql &= "				AND nxn.OrigemSerie_Id <> 'REC' "
                sql &= ""
                sql &= " ) AS CTE "

                sql &= "     LEFT JOIN Clientes cli" & vbCrLf & _
                                    "      ON cli.Cliente_Id = nf.Cliente_Id" & vbCrLf & _
                                    "     and cli.Endereco_Id = nf.EndCliente_Id" & vbCrLf & _
                                    "    LEFT JOIN Clientes emp" & vbCrLf & _
                                    "      ON emp.Cliente_Id = nf.Empresa_Id" & vbCrLf & _
                                    "     and emp.Endereco_Id = nf.EndEmpresa_Id" & vbCrLf & _
                                    "   WHERE nf.Situacao in (1,4,7)" & vbCrLf & _
                                    "     AND nf.TipoDeDocumento = 1" & vbCrLf
                If rbEntrada.Checked Then
                    sql &= "     AND nf.EntradaSaida_Id = 'E'" & vbCrLf & _
                           "     AND nf.CIFFOB = 'FOB'" & vbCrLf
                ElseIf rbSaida.Checked Then
                    sql &= "     AND nf.EntradaSaida_Id = 'S'" & vbCrLf & _
                           "     AND nf.CIFFOB = 'CIF'" & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                    sql &= "     AND nf.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                           "     AND nf.EndEmpresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(1) & "'" & vbCrLf
                End If

                If IsDate(txtDataInicial.Text) AndAlso IsDate(txtDataInicial.Text) Then
                    sql &= "     AND nf.Movimento BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                    sql &= "     AND nf.Cliente_Id = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf
                End If
                If Not String.IsNullOrWhiteSpace(txtCodigoTransportador.Value) Then
                    sql &= "     AND NxT.Proprietario = '" & txtCodigoTransportador.Value.Split("-")(0) & "'" & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
                    sql &= "     AND NF.Pedido = '" & txtPedido.Text & "'" & vbCrLf
                End If

                'sql &= "     AND isnull(cte.Nota_Id,0) = 0"
                sql &= " AND CTE.Quantidade = 0 "
                sql &= "     ORDER BY nf.Nota_Id, nf.Serie_Id, nf.Movimento" & vbCrLf

                ds = Banco.ConsultaDataSet(sql, "RelatorioDeNotasSemCte")

                Dim param As New Dictionary(Of String, Object)
                param.Add("ConsultaParam", getParam())

                Funcoes.BindReport(Me.Page, ds, "Cr_RelatorioDeNotasSemCte", eExportType.PDF, param)

            Else
                MsgBox(Me.Page, "Usuário sem permissão para visualizar o relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class