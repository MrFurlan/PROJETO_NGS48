Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class NotaFiscalXPesagem
    Inherits BasePage

    Private Sql As String
    Private mensagemErro As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        'If Not IsPostBack And IsConnect Then
        '    If Not Funcoes.VerificaPermissao("NotaFiscalXPesagem", "ACESSAR") Then
        MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
        Exit Sub
        '    End If
        '    CarregarEmpresas()
        '    LimparCampos()
        'End If
    End Sub

    Private Sub CarregarEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub LimparCampos()
        Session.Remove("objNFxPesagem")
        Session.Remove("objNFxRomaneio")
        ddlEmpresa.SelectedIndex = 0
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""
        txtPedido.Text = ""
        txtNotaFiscal.Text = ""
        txtPesagem.Text = ""
        btnPesagem.Enabled = False
        btnPedido.Enabled = True
        imgConfirmar.Visible = False
        txtDataInicial.Text = Now().ToString("dd/MM/yyyy")
        txtDataFinal.Text = Now().ToString("dd/MM/yyyy")
        gridNF.DataSource = Nothing
        gridNF.DataBind()
        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
    End Sub

    Private Sub CarregarNotasFiscais()
        Dim Emp() As String = ddlEmpresa.SelectedValue.ToString.Split("-")

        Sql = "SELECT NotasFiscais.EntradaSaida_Id, NotasFiscais.Nota_Id, NotasFiscais.Serie_Id, " & vbCrLf &
              "       NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, Clientes.Nome, " & vbCrLf &
              "       NotasFiscais.Operacao, NotasFiscais.SubOperacao, NotasFiscais.Movimento AS Data, " & vbCrLf &
              "       NotasFiscaisXItens.QuantidadeFiscal AS PesoFiscal, NotasFiscaisXItens.Valor, nfxr.Romaneio_Id " & vbCrLf &
              " FROM  NotasFiscais " & vbCrLf &
              " 		INNER JOIN  NotasFiscaisXItens " & vbCrLf &
              "				    ON  NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf &
              "				    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
              "				    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf &
              "				    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf &
              "				    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf &
              "				    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf &
              "				    AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf &
              " 		INNER JOIN  NotasFiscaisXRomaneios nfxr" & vbCrLf &
              "				    ON  NotasFiscaisXItens.Empresa_Id      = nfxr.Empresa_Id " & vbCrLf &
              "				    AND NotasFiscaisXItens.EndEmpresa_Id   = nfxr.EndEmpresa_Id " & vbCrLf &
              "				    AND NotasFiscaisXItens.Cliente_Id      = nfxr.Cliente_Id " & vbCrLf &
              "				    AND NotasFiscaisXItens.EndCliente_Id   = nfxr.EndCliente_Id " & vbCrLf &
              "				    AND NotasFiscaisXItens.EntradaSaida_Id = nfxr.EntradaSaida_Id " & vbCrLf &
              "				    AND NotasFiscaisXItens.Serie_Id        = nfxr.Serie_Id " & vbCrLf &
              "				    AND NotasFiscaisXItens.Nota_Id         = nfxr.Nota_Id " & vbCrLf &
              "		    INNER JOIN  Romaneios " & vbCrLf &
              "				    ON  nfxr.Empresa_Id    = Romaneios.Empresa_Id " & vbCrLf &
              "				    AND nfxr.EndEmpresa_Id = Romaneios.EndEmpresa_Id " & vbCrLf &
              "				    AND nfxr.Romaneio_Id   = Romaneios.Romaneio_Id " & vbCrLf &
              "	 	    INNER JOIN Produtos " & vbCrLf &
              "				    ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id " & vbCrLf &
              "		    INNER JOIN  Clientes " & vbCrLf &
              "				    ON  NotasFiscais.Cliente_Id     = Clientes.Cliente_Id " & vbCrLf &
              "				    AND NotasFiscais.EndCliente_Id  = Clientes.Endereco_Id " & vbCrLf &
              "		    INNER JOIN  SubOperacoes " & vbCrLf &
              "				    ON  NotasFiscais.Operacao    = SubOperacoes.Operacao_Id " & vbCrLf &
              "				    AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf &
              "WHERE (NotasFiscais.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
              "  AND (SubOperacoes.QuantidadeFisico = 'S') " & vbCrLf &
              "  AND (Produtos.Agrupar = 'N') " & vbCrLf &
              "  AND (NotasFiscais.TipoDeDocumento = 1) " & vbCrLf &
              "  AND (NotasFiscais.Empresa_Id      = '" & Emp(0) & "')" & vbCrLf &
              "  AND (NotasFiscais.EndEmpresa_Id   = " & Emp(1) & ")" & vbCrLf

        If txtCodigoCliente.Value.ToString.Length > 0 Then
            Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
            Sql &= "  AND (NotasFiscais.Cliente_Id    = '" & strCliente(0) & "')" & vbCrLf &
                   "  AND (NotasFiscais.EndCliente_Id = " & strCliente(1) & ")" & vbCrLf
        End If

        Sql &= "  AND (NotasFiscais.EntradaSaida_Id = '" & IIf(rdSaida.Checked, "S", "E") & "')" & vbCrLf

        If txtPedido.Text.Length > 0 Then Sql &= "  AND (NotasFiscais.Pedido = " & txtPedido.Text & ")" & vbCrLf

        Sql &= "  AND (Romaneios.Processo = 'NOTA FISCAL')"

        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(Sql, "NotasFiscais")

        gridNF.DataSource = ds
        gridNF.DataBind()
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteNotaXPes" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteNotaXPes" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteNotaXPes" & HID.Value)
        ElseIf Not Session("objPedidoSelecionadoNotaXPes" & HID.Value) Is Nothing Then
            Dim Emp() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
            If CType(Session("objPedidoSelecionadoNotaXPes" & HID.Value), [Lib].Negocio.Pedido).CodigoEmpresa <> Emp(0) Or CType(Session("objPedidoSelecionadoNotaXPes" & HID.Value), [Lib].Negocio.Pedido).EnderecoEmpresa.ToString <> Emp(1) Then
                MsgBox(Me.Page, "Empresa do Pedido é diferente da Empresa selecionada.")
            ElseIf CType(Session("objPedidoSelecionadoNotaXPes" & HID.Value), [Lib].Negocio.Pedido).CodigoCliente <> strCliente(0) Or CType(Session("objPedidoSelecionadoNotaXPes" & HID.Value), [Lib].Negocio.Pedido).EnderecoCliente.ToString <> strCliente(1) Then
                MsgBox(Me.Page, "Cliente do Pedido é diferente do Cliente selecionado.")
            Else
                txtPedido.Text = CType(Session("objPedidoSelecionadoNotaXPes" & HID.Value), [Lib].Negocio.Pedido).Codigo
            End If
            Session.Remove("objPedidoSelecionadoNotaXPes" & HID.Value)
        ElseIf Not Session("RomaneioNFxP" & HID.Value) Is Nothing Then
            Dim ROM As [Lib].Negocio.Romaneio = Session("RomaneioNFxP" & HID.Value)
            Session.Remove("RomaneioNFxP" & HID.Value)
            Session.Remove("ProcurandoRomaneio" & Session.SessionID)

            If ROM.CodigoPedido = CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).CodigoPedido AndAlso
               ROM.EntradaSaida = CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).EntradaSaida.ToString.Substring(0, 1) AndAlso
               ROM.CodigoOperacao = CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).CodigoOperacao AndAlso
               ROM.CodigoSubOperacao = CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).CodigoSubOperacao Then
                Session("objNFxRomaneio" & HID.Value) = ROM
                txtPesagem.Text = ROM.Pesagens(0).CodigoPesagem
                imgConfirmar.Visible = True
            Else
                MsgBox(Me.Page, "Pedido do Romaneio " & ROM.CodigoPedido.ToString & " não pode ser diferente do Pedido da Nota Fiscal " & CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).CodigoPedido.ToString)
            End If
        End If
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        txtCodigoCliente.Value = ""
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteNotaXPes" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
        ElseIf txtCodigoCliente.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Cliente não foi selecionado.")
        Else
            HttpContext.Current.Session("ssCampo") = "Pedidos"

            Dim strJavaScript As String = "var x = (screen.height / 2) - 200; "
            strJavaScript &= "var y = (screen.width / 2) - 400; "
            strJavaScript &= "window.open(""ConsultaPedidos.aspx?url=PesoDeChegada&tipo=NotaXPes"

            Dim Emp() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            'strJavaScript &= "&emp=" & Emp(0)
            'strJavaScript &= "&ende=" & Emp(1)

            Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
            'If strCliente(0) <> "" Then strJavaScript &= "&cli=" & strCliente(0)
            'If strCliente.Length > 1 Then strJavaScript &= "&endc=" & strCliente(1)

            ucConsultaPedidos.SetarHID(HID.Value)
            Dim parameters As New Dictionary(Of String, Object)
            parameters("empresa") = Emp(0)
            parameters("enderecoEmpresa") = Emp(1)
            parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
            parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
            parameters("es") = IIf(rdEntrada.Checked, "E", "S")
            ucConsultaPedidos.BindGridView(parameters)
            Popup.ConsultaDePedidos(Me.Page, "objPedidoSelecionadoNotaXPes" & HID.Value)

            'ucConsultaPedidos.Carregar()

            'strJavaScript &= """, """", ""resizable=no, menubar=no, scrollbars=Yes, width=800, height=400, top="" + x + "", left="" + y + """");"
            'ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), "BuscarPedido", strJavaScript, True)
        End If
    End Sub

    Protected Sub lnkSelecionarNF_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim lnk As LinkButton = CType(sender, LinkButton)
        Dim row As GridViewRow = CType(lnk.NamingContainer, GridViewRow)
        Dim Emp() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
        Dim objNotaFiscal As New [Lib].Negocio.NotaFiscal()
        objNotaFiscal.CodigoEmpresa = Emp(0)
        objNotaFiscal.EnderecoEmpresa = Emp(1)
        objNotaFiscal.CodigoCliente = gridNF.Rows(row.RowIndex).Cells(4).Text()
        objNotaFiscal.EnderecoCliente = gridNF.Rows(row.RowIndex).Cells(5).Text()
        objNotaFiscal.EntradaSaida = IIf(gridNF.Rows(row.RowIndex).Cells(1).Text() = "S", eEntradaSaida.Saida, eEntradaSaida.Entrada)
        objNotaFiscal.Serie = gridNF.Rows(row.RowIndex).Cells(3).Text()
        objNotaFiscal.Codigo = gridNF.Rows(row.RowIndex).Cells(2).Text()
        objNotaFiscal = New [Lib].Negocio.NotaFiscal(objNotaFiscal)
        Session("objNFxPesagem" & HID.Value) = objNotaFiscal

        txtNotaFiscal.Text = objNotaFiscal.Codigo
        btnPesagem.Enabled = True
    End Sub

    Protected Sub btnPesagem_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).Itens(0).Produto.Agrupar = "N" And CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).SubOperacao.Classe <> eClassesOperacoes.GLOBAL Then
            If CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).SubOperacao.QuantidadeFisico Then
                Session("ProcurandoRomaneio" & Session.SessionID) = True
                Session("objNotaFiscal" & HID.Value) = CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal)
                Session("Unitario" & HID.Value.ToString) = CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).Itens(0).Unitario
                Session("Op" & HID.Value.ToString) = CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).CodigoOperacao
                Session("SOp" & HID.Value.ToString) = CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).CodigoSubOperacao
                ucConsultaRomaneios.SetarHID(HID.Value)
                Popup.ConsultaDeRomaneios(Me, "RomaneioNFxP" & HID.Value)
                ucConsultaRomaneios.LimparCampos()
                ucConsultaRomaneios.CargaRomaneios()

                'Dim xPop As String = 250
                'Dim yPop As String = 400
                'ScriptManager.RegisterStartupScript(Me, Me.GetType(), Guid.NewGuid().ToString(), "window.open('ConsultaRomaneios.aspx?Tipo=NFxP&Emp=" & CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).CodigoEmpresa & "&EndEmp=" & CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).EnderecoEmpresa & "&Pedido=" & CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).CodigoPedido & "&Op=" & CType(Session("objNFxPesagem" + HID.Value), [Lib].Negocio.NotaFiscal).CodigoOperacao & "&SOp=" & CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).CodigoSubOperacao & "&Unitario=" & Str(CType(Session("objNFxPesagem" & HID.Value), [Lib].Negocio.NotaFiscal).Itens(0).ValorUnitario) & "', '', 'resizable=no, menubar=no, scrollbars=No, width=800, height=500, top=" + xPop + ", left=" + yPop + "');", True)
            Else
                MsgBox(Me.Page, "Operação da Nota Fiscal não é para Quantidade Física")
            End If
        Else
            MsgBox(Me.Page, "Não é informado Pesagem para este Produto")
        End If
    End Sub

    Protected Sub imgConfirmar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If Funcoes.VerificaPermissao("NotaFiscalXPesagem", "GRAVAR") Then
            Dim sqls As New ArrayList

            Sql = " Delete NotasFiscaisXRomaneios" & vbCrLf &
                  "  Where Empresa_Id      ='" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).CodigoEmpresa & "'" & vbCrLf &
                  "    and EndEmpresa_Id   = " & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).EnderecoEmpresa & vbCrLf &
                  "    and Cliente_Id      ='" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).CodigoCliente & "'" & vbCrLf &
                  "    and EndCliente_Id   = " & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).EnderecoCliente & vbCrLf &
                  "    and EntradaSaida_Id ='" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                  "    and Serie_Id        ='" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).Serie & "'" & vbCrLf &
                  "    and Nota_Id         = " & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).Codigo & vbCrLf &
                  "    and Romaneio_Id     = " & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).CodigoRomaneio
            sqls.Add(Sql)

            Sql = " Delete Romaneios" & vbCrLf &
                  "  Where Empresa_Id      ='" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).CodigoEmpresa & "'" & vbCrLf &
                  "    and EndEmpresa_Id   = " & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).EnderecoEmpresa & vbCrLf &
                  "    and Romaneio_Id     = " & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).CodigoRomaneio
            sqls.Add(Sql)

            Sql = " Insert into NotasFiscaisXRomaneios(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Romaneio_Id)" & vbCrLf &
                  " Values('" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).CodigoEmpresa & "'" & vbCrLf &
                          "," & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).EnderecoEmpresa & vbCrLf &
                          ",'" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).CodigoCliente & "'" & vbCrLf &
                          "," & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).EnderecoCliente & vbCrLf &
                          ",'" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                          ",'" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).Serie & "'" & vbCrLf &
                          "," & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).Codigo & vbCrLf &
                          "," & CType(Session("objNFxRomaneio"), [Lib].Negocio.Romaneio).Codigo & ")"
            sqls.Add(Sql)

            Sql = " Update NotasFiscaisxItens set " & vbCrLf &
                  "       Deposito         ='" & CType(Session("objNFxRomaneio"), [Lib].Negocio.Romaneio).CodigoDeposito & "'" & vbCrLf &
                  "      ,EndDeposito      = " & CType(Session("objNFxRomaneio"), [Lib].Negocio.Romaneio).EnderecoDeposito & vbCrLf &
                  "      ,QuantidadeFisica = " & CType(Session("objNFxRomaneio"), [Lib].Negocio.Romaneio).PesoLiquido & vbCrLf &
                  "  Where Empresa_Id      ='" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).CodigoEmpresa & "'" & vbCrLf &
                  "    and EndEmpresa_Id   = " & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).EnderecoEmpresa & vbCrLf &
                  "    and Cliente_Id      ='" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).CodigoCliente & "'" & vbCrLf &
                  "    and EndCliente_Id   = " & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).EnderecoCliente & vbCrLf &
                  "    and EntradaSaida_Id ='" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                  "    and Serie_Id        ='" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).Serie & "'" & vbCrLf &
                  "    and Nota_Id         = " & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).Codigo
            sqls.Add(Sql)

            Sql = " Update NotasFiscais set " & vbCrLf &
                  "       Deposito         ='" & CType(Session("objNFxRomaneio"), [Lib].Negocio.Romaneio).CodigoDeposito & "'" & vbCrLf &
                  "      ,EndDeposito      = " & CType(Session("objNFxRomaneio"), [Lib].Negocio.Romaneio).EnderecoDeposito & vbCrLf &
                  "  Where Empresa_Id      ='" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).CodigoEmpresa & "'" & vbCrLf &
                  "    and EndEmpresa_Id   = " & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).EnderecoEmpresa & vbCrLf &
                  "    and Cliente_Id      ='" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).CodigoCliente & "'" & vbCrLf &
                  "    and EndCliente_Id   = " & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).EnderecoCliente & vbCrLf &
                  "    and EntradaSaida_Id ='" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                  "    and Serie_Id        ='" & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).Serie & "'" & vbCrLf &
                  "    and Nota_Id         = " & CType(Session("objNFxPesagem"), [Lib].Negocio.NotaFiscal).Codigo
            sqls.Add(Sql)

            If Banco.GravaBanco(sqls) Then
                Session.Remove("objNFxPesagem")
                Session.Remove("objNFxRomaneio")
                txtNotaFiscal.Text = ""
                txtPesagem.Text = ""
                btnPesagem.Enabled = False
                imgConfirmar.Visible = False

                CarregarNotasFiscais()
            Else
                MsgBox(Me.Page, "Erro de Gravação: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte.")
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        If Funcoes.VerificaPermissao("NotaFiscalXPesagem", "LEITURA") Then
            If ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.")
            Else
                CarregarNotasFiscais()
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar Registro(s).")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        LimparCampos()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "NotaFiscalXPesagem")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class