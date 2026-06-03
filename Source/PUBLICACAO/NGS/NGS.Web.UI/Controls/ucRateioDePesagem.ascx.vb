Imports NGS.Lib.Negocio

Public Class ucRateioDePesagem
    Inherits BaseUserControl
    Private objLaudo As [Lib].Negocio.Pesagem

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    End Sub

#Region "Sessão"
    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Private Sub SessaoSalvaLaudo()
        Session("ssLaudo" & HID.Value.ToString) = objLaudo
    End Sub

    Private Sub SessaoRecuperaLaudo()
        objLaudo = CType(Session("ssLaudo" & HID.Value.ToString), [Lib].Negocio.Pesagem)
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Session("objClienteRxP" & HID.Value) IsNot Nothing Then
            Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objClienteRxP" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(pCliente)
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteRxP" & HID.Value)
            btnPedido.Enabled = True
            CarregarPopupPedido(True)
        ElseIf Session("objPedidoRxP" & HID.Value) IsNot Nothing Then
            SessaoRecuperaLaudo()

            If Not CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Itens(0).Produto.CodigoGrupo = objLaudo.Produto.CodigoGrupo Then
                MsgBox(Me.Page, "Produto com Grupo diferente do Laudo não pode ser usado")
                Limpar()
                Exit Sub
            End If

            'For intRegistro As Integer = gridRateio.Rows.Count - 1 To 0 Step -1
            '    If gridRateio.Rows(intRegistro).Cells(3).Text = CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Codigo Then
            '        MsgBox(Me.Page, "Pedido já foi usado no Rateio")
            '        Limpar()
            '        Exit Sub
            '    End If
            'Next intRegistro

            txtPedido.Text = CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Codigo
            Dim subOperacaoDoPedido As String = CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).CodigoOperacao & "-" & CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).CodigoSubOperacao
            Dim subOperacoes As New [Lib].Negocio.ListSubOperacao(" So.Operacao_Id = " & CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).CodigoOperacao & " AND So.EntradaSaida = '" & objLaudo.EntradaSaida & "'")
            ddlSubOperacao.Items.Insert(0, "")
            For Each objSubOperacao As SubOperacao In subOperacoes
                ddlSubOperacao.Items.Add(New ListItem(objSubOperacao.Operacao.Codigo.ToString("00") & "-" & objSubOperacao.Codigo.ToString("00") & " " & objSubOperacao.Descricao, _
                                                      objSubOperacao.Operacao.Codigo & "-" & objSubOperacao.Codigo))
            Next
            Dim intIndice As Integer = ddlSubOperacao.Items.IndexOf(ddlSubOperacao.Items.FindByValue(subOperacaoDoPedido))
            If Not intIndice = -1 Then
                ddlSubOperacao.SelectedIndex = intIndice
            End If

            If ddlSubOperacao.SelectedIndex > 0 Then CarregarSaldoPedido(objLaudo.CodigoEmpresa, objLaudo.EnderecoEmpresa, txtPedido.Text)

            Funcoes.InserirLinhaEmBranco(ddlDeposito)

            Dim dep As String = String.Empty
            Dim enddep As Integer = 0
            For intDeposito As Integer = CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Depositos.Count - 1 To 0 Step -1
                If CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Depositos(intDeposito).Tipo = "DE" Then
                    ddlDeposito.Items.Add(New ListItem(Funcoes.AlinharEsquerda(CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Depositos(intDeposito).Deposito.Nome, 28, ".") & " - " & _
                                                       Funcoes.AlinharEsquerda(CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Depositos(intDeposito).Deposito.Cidade, 20, ".") & " " & _
                                                       CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Depositos(intDeposito).Deposito.CodigoEstado & " " & _
                                                       Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Depositos(intDeposito).Codigo), 18, ".") & "-" & _
                                                       CStr(CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Depositos(intDeposito).CodigoEndereco) & "-" & _
                                                       CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Depositos(intDeposito).Deposito.Reduzido, _
                                                       CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Depositos(intDeposito).Codigo & "-" & _
                                                       CStr(CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Depositos(intDeposito).CodigoEndereco)))
                    If dep.Length = 0 Then
                        dep = CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Depositos(intDeposito).Codigo
                        enddep = CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Depositos(intDeposito).CodigoEndereco
                    End If
                End If
            Next intDeposito

            intIndice = ddlDeposito.Items.IndexOf(ddlDeposito.Items.FindByValue(dep & "-" & enddep))
            If Not intIndice = -1 Then
                ddlDeposito.SelectedIndex = intIndice
            End If
        End If
    End Sub

#End Region

    Public Sub CarregarPopupPedido(ByVal AnoCorrente As Boolean)
        SessaoRecuperaLaudo()
        Dim parameters As New Dictionary(Of String, Object)
        Session("ssCampo" & HID.Value) = "Laudo"
        Dim objCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
        parameters("CodigoEmpresa") = objLaudo.CodigoEmpresa
        parameters("CodigoEndEmpresa") = objLaudo.EnderecoEmpresa
        parameters("CodigoCliente") = objCliente(0)
        parameters("CodigoEndCliente") = objCliente(1)
        parameters("AnoCorrente") = AnoCorrente
        parameters("Consolidado") = False
        Dim ucConsultaPedidos As ucConsultaPedidos = DirectCast(Me.NamingContainer.FindControl("ucConsultaPedidos"), ucConsultaPedidos)
        If ucConsultaPedidos IsNot Nothing Then
            ucConsultaPedidos.SetarHID(HID.Value)
            ucConsultaPedidos.MainUserControl = Me
            ucConsultaPedidos.BindGridView(parameters)
            Popup.ConsultaDePedidos(Me.Page, "objPedidoRxP" & HID.Value.ToString)
        End If
    End Sub

    Private Sub CarregarSaldoPedido(ByVal Empresa As String, ByVal EndEmpresa As String, ByVal Pedido As String)
        Dim sOp() As String = ddlSubOperacao.SelectedValue.ToString.Split("-")
        Dim suboperacao As New [Lib].Negocio.SubOperacao(sOp(0), sOp(1))

        objLaudo = New Pesagem()
        objLaudo.CodigoEmpresa = Empresa
        objLaudo.EnderecoEmpresa = EndEmpresa
        objLaudo.CodigoPedido = Pedido

        Dim dsSaldoPedido As DataSet = objLaudo.SaldoDePedidos(objLaudo, False)
        If suboperacao.Devolucao Then
            txtSaldoPedido.Text = dsSaldoPedido.Tables(0).Rows(0).Item("SaldoDevolucao")
        Else
            txtSaldoPedido.Text = dsSaldoPedido.Tables(0).Rows(0).Item("Saldo")
        End If

        If suboperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso (suboperacao.Classe = eClassesOperacoes.AFIXAR Or suboperacao.Classe = eClassesOperacoes.DEPOSITOS) Then
            txtQuantidade.ReadOnly = False
        Else
            If CInt(txtSaldoPedido.Text) > 0 Then txtQuantidade.ReadOnly = False
        End If
    End Sub

    Public Sub CarregarAutorizacaoDeRetirada()
        SessaoRecuperaLaudo()
        Dim parametros As New Dictionary(Of String, Object)
        parametros("emp") = objLaudo.CodigoEmpresa
        parametros("endemp") = objLaudo.EnderecoEmpresa
        Dim objCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
        parametros("cli") = objCliente(0)
        parametros("endcli") = objCliente(1)
        parametros("ped") = txtPedido.Text
        Dim sOp() As String = ddlSubOperacao.SelectedValue.ToString.Split("-")
        Dim suboperacao As New [Lib].Negocio.SubOperacao(sOp(0), sOp(1))
        parametros("classe") = suboperacao.Classe
        parametros("romaneio") = True
        Dim ucConsultaAutorizacaoDeRetirada As ucConsultaAutorizacaoDeRetirada = DirectCast(Me.NamingContainer.FindControl("ucConsultaAutorizacaoDeRetirada"), ucConsultaAutorizacaoDeRetirada)
        If ucConsultaAutorizacaoDeRetirada IsNot Nothing Then
            ucConsultaAutorizacaoDeRetirada.SetarHID(HID.Value)
            ucConsultaAutorizacaoDeRetirada.MainUserControl = Me
            ucConsultaAutorizacaoDeRetirada.BindGridView(parametros)
            Popup.ConsultaDeAutorizacaoDeRetirada(Me.Page, "objAutorizacaoRxP" & HID.Value.ToString)
        End If
    End Sub

    Public Sub CarregarAutorizacao(Par As Hashtable)
        SessaoRecuperaLaudo()
        Dim sOp() As String = ddlSubOperacao.SelectedValue.ToString.Split("-")
        Dim suboperacao As New [Lib].Negocio.SubOperacao(sOp(0), sOp(1))
        Dim Autorizacao As New [Lib].Negocio.AutorizacaoDeRetirada(objLaudo.CodigoEmpresa, objLaudo.EnderecoEmpresa, Par("Pedido"), Par("Autorizacao"), suboperacao.Classe)
        If (Autorizacao.SaldoFisico < Convert.ToDouble(txtQuantidade.Text)) Then
            MsgBox(Me.Page, "Autorização com saldo insuficiente - Saldo: " & Autorizacao.SaldoFisico)
        Else
            txtAutorizacao.Text = Session("objAutorizacaoRxP" & HID.Value.ToString)
            AdicionarRateio()
        End If
    End Sub

    Private Sub AdicionarRateio()
        SessaoRecuperaLaudo()

        If CInt(txtQuantidade.Text) < CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(0).Item("Liquido") Then
            Dim ParcialPercentual As Decimal = CInt(txtQuantidade.Text) / CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(0).Item("Liquido")
            Dim ParcialBruto As Integer = CInt(CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(0).Item("Bruto") * ParcialPercentual)
            Dim ParcialDesconto As Integer = CInt(CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(0).Item("Bruto") * ParcialPercentual) - CInt(txtQuantidade.Text)

            CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(0).Item("Bruto") = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(0).Item("Bruto") - ParcialBruto
            CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(0).Item("Desconto") = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(0).Item("Desconto") - ParcialDesconto
            CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(0).Item("Liquido") = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(0).Item("Liquido") - CInt(txtQuantidade.Text)
            CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(0).Item("Percentual") = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(0).Item("Percentual") - ParcialPercentual

            Dim sOp() As String = ddlSubOperacao.SelectedValue.ToString.Split("-")
            Dim dep() As String = ddlDeposito.SelectedValue.ToString.Split("-")
            Dim newRow As DataRow = CType(Session("dtRomaneios" & HID.Value), DataTable).NewRow
            newRow.Item("Codigo") = CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).CodigoCliente & "-" & CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).EnderecoCliente
            newRow.Item("Nome") = CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Cliente.Nome
            newRow.Item("End") = CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).EnderecoEmpresa
            newRow.Item("Pedido") = CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Codigo
            newRow.Item("CodigoDeposito") = dep(0) & "-" & dep(1)
            newRow.Item("Produto") = CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).Itens(0).CodigoProduto
            newRow.Item("OP") = sOp(0)
            newRow.Item("SO") = sOp(1)
            If String.IsNullOrWhiteSpace(txtAutorizacao.Text) Then
                newRow.Item("Autorizacao") = "0"
            Else
                newRow.Item("Autorizacao") = txtAutorizacao.Text
            End If
            newRow.Item("Bruto") = ParcialBruto
            newRow.Item("Desconto") = ParcialDesconto
            newRow.Item("Liquido") = txtQuantidade.Text
            newRow.Item("Percentual") = ParcialPercentual
            newRow.Item("Principal") = False

            CType(Session("dtRomaneios" & HID.Value), DataTable).Rows.Add(newRow)

            gridRateio.DataSource = CType(Session("dtRomaneios" & HID.Value), DataTable)
            gridRateio.DataBind()

            For intRegistro As Integer = gridRateio.Rows.Count - 1 To 0 Step -1
                If CType(gridRateio.Rows(intRegistro).FindControl("chkPrincipal"), CheckBox).Checked = True Then
                    CType(gridRateio.Rows(intRegistro).FindControl("imgExcluir"), ImageButton).Visible = False
                Else
                    CType(gridRateio.Rows(intRegistro).FindControl("imgExcluir"), ImageButton).Visible = True
                End If
            Next intRegistro

            tdBotaoNovo.Visible = True

            LimparCampos("I")
        Else
            MsgBox(Me.Page, "Quantidade não pode ser maior que o Laudo original")
        End If
    End Sub

    Public Sub carregarRomaneios()
        SessaoRecuperaLaudo()

        Dim dtRomaneios As New DataTable()
        dtRomaneios.Columns.Add("Codigo")
        dtRomaneios.Columns.Add("Nome")
        dtRomaneios.Columns.Add("End")
        dtRomaneios.Columns.Add("Pedido")
        dtRomaneios.Columns.Add("CodigoDeposito")
        dtRomaneios.Columns.Add("Produto")
        dtRomaneios.Columns.Add("Op")
        dtRomaneios.Columns.Add("SO")
        dtRomaneios.Columns.Add("Autorizacao")
        dtRomaneios.Columns.Add("Bruto")
        dtRomaneios.Columns.Add("Desconto")
        dtRomaneios.Columns.Add("Liquido")
        dtRomaneios.Columns.Add("Percentual")
        dtRomaneios.Columns.Add("Principal")

        Session("dtRomaneios" & HID.Value) = dtRomaneios

        For Each row As [Lib].Negocio.Romaneio In objLaudo.Romaneios
            Dim newRow As DataRow = CType(Session("dtRomaneios" & HID.Value), DataTable).NewRow
            newRow.Item("Codigo") = objLaudo.CodigoCliente & "-" & objLaudo.EnderecoCliente
            newRow.Item("Nome") = objLaudo.Cliente.Nome
            newRow.Item("End") = objLaudo.Cliente.CodigoEndereco
            newRow.Item("Pedido") = row.CodigoPedido
            newRow.Item("CodigoDeposito") = objLaudo.CodigoDeposito & "-" & objLaudo.EnderecoDeposito
            newRow.Item("Produto") = row.CodigoProduto
            newRow.Item("OP") = row.CodigoOperacao
            newRow.Item("SO") = row.CodigoSubOperacao
            newRow.Item("Autorizacao") = row.CodigoAutorizacao
            newRow.Item("Bruto") = row.PesoBruto
            newRow.Item("Desconto") = row.Desconto
            newRow.Item("Liquido") = row.PesoLiquido
            newRow.Item("Percentual") = 1
            newRow.Item("Principal") = True
            CType(Session("dtRomaneios" & HID.Value), DataTable).Rows.Add(newRow)
        Next

        gridRateio.DataSource = CType(Session("dtRomaneios" & HID.Value), DataTable)
        gridRateio.DataBind()
    End Sub

    Public Sub LimparCampos(ByVal tipo As String)
        If String.IsNullOrWhiteSpace(tipo) Then
            tdBotaoNovo.Visible = False
            Session.Remove("dtRomaneios" & HID.Value)
        End If
        Session.Remove("objClienteRxP" & HID.Value)
        Session.Remove("objPedidoRxP" & HID.Value)
        Session.Remove("objAutorizacaoRxP" & HID.Value)
        Session.Remove("_MainUserControl")
        txtCliente.Text = String.Empty
        btnCliente.Enabled = True
        txtPedido.Text = String.Empty
        btnPedido.Enabled = False
        txtSaldoPedido.Text = String.Empty
        txtAutorizacao.Text = String.Empty
        ddlDeposito.Items.Clear()
        ddlSubOperacao.Items.Clear()
        txtQuantidade.Text = String.Empty
        txtQuantidade.ReadOnly = True
    End Sub

    Protected Sub lnkGravar_Click(sender As Object, e As EventArgs) Handles lnkGravar.Click
        Dim Sqls As New ArrayList
        SessaoRecuperaLaudo()

        For Each xRomaneio As Romaneio In objLaudo.Romaneios
            xRomaneio.IUD = "D"
            xRomaneio.SalvarSql(Sqls, False)
        Next

        For intRegistro As Integer = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows.Count - 1 To 0 Step -1
            Dim SqlN As String = "exec sp_Numerador '" & objLaudo.CodigoEmpresa & "'," & objLaudo.EnderecoEmpresa & ",110"
            Dim dsN As New DataSet
            dsN = Banco.ConsultaDataSet(SqlN, "Numerador")

            Dim pRomaneio As New Romaneio()
            pRomaneio.Codigo = dsN.Tables(0).Rows(0).Item(0)

            pRomaneio.IUD = "I"
            pRomaneio.CodigoEmpresa = objLaudo.CodigoEmpresa
            pRomaneio.EnderecoEmpresa = objLaudo.EnderecoEmpresa

            pRomaneio.EntradaSaida = objLaudo.EntradaSaida
            pRomaneio.CodigoPedido = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intRegistro).Item("Pedido")

            Dim dep() As String = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intRegistro).Item("CodigoDeposito").ToString.Split("-")
            pRomaneio.CodigoDeposito = dep(0)
            pRomaneio.EnderecoDeposito = dep(1)

            If objLaudo.EntradaSaida = "E" Then
                pRomaneio.CodigoDestino = objLaudo.CodigoEmpresa
                pRomaneio.EnderecoDestino = objLaudo.EnderecoEmpresa
            Else
                Dim cli() As String = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intRegistro).Item("Codigo").ToString.Split("-")
                pRomaneio.CodigoDestino = cli(0)
                pRomaneio.EnderecoDestino = cli(1)
            End If

            pRomaneio.CodigoTransbordo = ""
            pRomaneio.EnderecoTransbordo = 0

            pRomaneio.CodigoProduto = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intRegistro).Item("Produto")
            pRomaneio.CodigoOperacao = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intRegistro).Item("OP")
            pRomaneio.CodigoSubOperacao = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intRegistro).Item("SO")
            pRomaneio.Movimento = objLaudo.Movimento

            pRomaneio.PesoBruto = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intRegistro).Item("Bruto")
            pRomaneio.Desconto = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intRegistro).Item("Desconto")
            pRomaneio.PesoLiquido = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intRegistro).Item("Liquido")

            pRomaneio.Observacoes = ""
            pRomaneio.Processo = "Rateio"

            pRomaneio.CodigoAutorizacao = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intRegistro).Item("Autorizacao")

            pRomaneio.DescontosAnalises = New ListRomaneioXDesconto(pRomaneio)

            For Each row As PesagemXAnalises In objLaudo.Analises
                Dim RxD As New RomaneioXDesconto(pRomaneio)
                RxD.CodigoAnalise = row.CodigoAnalise
                RxD.Desconto = row.Desconto * CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intRegistro).Item("Percentual")
                RxD.Percentual = row.Percentual
                RxD.Indice = row.Indice
                pRomaneio.DescontosAnalises.Add(RxD)
            Next

            pRomaneio.SalvarSql(Sqls, False)

            Dim rxp As New RomaneioXPesagem()
            rxp.IUD = "I"
            rxp.CodigoEmpresa = pRomaneio.CodigoEmpresa
            rxp.EndEmpresa = pRomaneio.EnderecoEmpresa
            rxp.CodigoRomaneio = pRomaneio.Codigo
            rxp.CodigoPesagem = objLaudo.Codigo
            rxp.Sequencia = 0

            rxp.SalvarSql(Sqls)
        Next intRegistro

        If Banco.GravaBanco(Sqls) Then
            MsgBox(Me.Page, "Gravado com Sucesso.", eTitulo.Sucess)
            LimparCampos("")
            'CType(Me.Page, RateioDePesagemNovo).Limpar()
            'CType(Me.Page, RateioDePesagemNovo).Consultar(objLaudo.CodigoEmpresa, objLaudo.EnderecoEmpresa, objLaudo.Codigo, "I")

            Popup.CloseDialog(Me.Page, "divRateioDePesagem")
        Else
            MsgBox(Me.Page, "Erro ao Gravar Registro(Rateio): " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte.")
        End If
    End Sub

    Protected Sub imgExcluir_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        SessaoRecuperaLaudo()

        Dim Bruto As Decimal = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(row.RowIndex).Item("Bruto")
        Dim Desconto As Decimal = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(row.RowIndex).Item("Desconto")
        Dim Liquido As Decimal = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(row.RowIndex).Item("Liquido")
        Dim Percentual As Decimal = CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(row.RowIndex).Item("Percentual")

        CType(Session("dtRomaneios" & HID.Value), DataTable).Rows.Remove(CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(row.RowIndex))

        For intPrincipal As Integer = gridRateio.Rows.Count - 1 To 0 Step -1
            If gridRateio.Rows(intPrincipal).Cells(3).Text = objLaudo.CodigoPedido Then
                CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intPrincipal).Item("Bruto") += Bruto
                CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intPrincipal).Item("Desconto") += Desconto
                CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intPrincipal).Item("Liquido") += Liquido
                CType(Session("dtRomaneios" & HID.Value), DataTable).Rows(intPrincipal).Item("Percentual") += Percentual
            End If
        Next intPrincipal

        gridRateio.DataSource = CType(Session("dtRomaneios" & HID.Value), DataTable)
        gridRateio.DataBind()

        For intRegistro As Integer = gridRateio.Rows.Count - 1 To 0 Step -1
            If gridRateio.Rows(intRegistro).Cells(3).Text = objLaudo.CodigoPedido Then
                CType(gridRateio.Rows(intRegistro).FindControl("imgExcluir"), ImageButton).Visible = False
            Else
                CType(gridRateio.Rows(intRegistro).FindControl("imgExcluir"), ImageButton).Visible = True
            End If
        Next intRegistro

        If gridRateio.Rows.Count = 1 Then tdBotaoNovo.Visible = False

        LimparCampos("I")
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        LimparCampos("I")
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        LimparCampos("")
        'CType(Me.Page, RateioDePesagemNovo).Limpar()
        Popup.CloseDialog(Me.Page, "divRateioDePesagem")
    End Sub

    Protected Sub btnCliente_Click(sender As Object, e As EventArgs) Handles btnCliente.Click
        Dim ucConsultaClientes As ucConsultaClientes = DirectCast(Me.NamingContainer.FindControl("ucConsultaClientes"), ucConsultaClientes)
        If ucConsultaClientes IsNot Nothing Then
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarHID(HID.Value)
            ucConsultaClientes.MainUserControl = Me
            Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeClientes(Me.Page, "objClienteRxP" & HID.Value, txtNome.ClientID, True, 500)
        End If
    End Sub

    Protected Sub btnPedido_Click(sender As Object, e As EventArgs) Handles btnPedido.Click
        CarregarPopupPedido(False)
    End Sub

    Protected Sub ddlSubOperacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlSubOperacao.SelectedIndexChanged
        If ddlSubOperacao.SelectedIndex > 0 Then
            SessaoRecuperaLaudo()
            CarregarSaldoPedido(objLaudo.CodigoEmpresa, objLaudo.EnderecoEmpresa, txtPedido.Text)
        End If
    End Sub

    Protected Sub cmdQuantidade_Click(sender As Object, e As EventArgs) Handles cmdQuantidade.Click
        If String.IsNullOrWhiteSpace(txtQuantidade.Text) OrElse txtQuantidade.Text.Equals("0") Then
            MsgBox(Me.Page, "Quantidade deve ser maior que 0(zero)")
        ElseIf ddlDeposito.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Depósito não foi selecionado")
        ElseIf ddlSubOperacao.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Sub-Operação não foi selecionada")
        Else
            Dim sOp() As String = ddlSubOperacao.SelectedValue.ToString.Split("-")
            Dim suboperacao As New [Lib].Negocio.SubOperacao(sOp(0), sOp(1))

            If suboperacao.EntradaSaida = eEntradaSaida.Entrada Then
                If CInt(txtQuantidade.Text) > CInt(txtSaldoPedido.Text) Then
                    If (suboperacao.Classe = eClassesOperacoes.AFIXAR Or suboperacao.Classe = eClassesOperacoes.DEPOSITOS) Then
                        AdicionarRateio()
                    Else
                        MsgBox(Me.Page, "Quantidade não pode ser maior que o Saldo do Pedido")
                    End If
                Else
                    AdicionarRateio()
                End If
            Else
                If CInt(txtQuantidade.Text) > CInt(txtSaldoPedido.Text) Then
                    MsgBox(Me.Page, "Quantidade não pode ser maior que o Saldo do Pedido")
                Else
                    CarregarAutorizacaoDeRetirada()
                End If
            End If
        End If
    End Sub

End Class