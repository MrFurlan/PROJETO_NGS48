Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RateioDePesagem
    Inherits BasePage

    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim Empresa() As String
    Dim Deposito() As String
    Dim Cliente() As String
    Dim Destinatario() As String
    Dim CodigoOperacao() As String
    Dim Bruto As Integer = 0
    Dim Desconto As Integer = 0
    Dim Liquido As Integer = 0
    Dim Percentual As Decimal = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RateioDePesagem", "ACESSAR") Then
                    Limpar()
                    SetarEmpresa(Session("ssEmpresa"), Session("ssEndEmpresa"))
                    SetarEmpresaAlt(Session("ssEmpresa"), Session("ssEndEmpresa"))
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub SalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub RecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub CargaAnalises()
        Sql = "Select top 5 Analise_Id as Codigo, Descricao, 0 as Percentual, 0 as Indice, 0 as Desconto from analises"
        GridDescontos.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
        GridDescontos.DataBind()
    End Sub

    Private Sub Limpar()
        btnEmpresa.Enabled = True
        txtLaudo.Enabled = True
        lnkConsultar.Parent.Visible = True
        lnkLimpar.Parent.Visible = True
        DdlCopias.Enabled = True
        lnkExcluir.Parent.Visible = False
        lnkRelatorio.Parent.Visible = False
        divAnalises.Visible = False
        divRateio.Visible = False
        divMenuRateio.Visible = False

        txtLaudo.Text = String.Empty
        txtMovimento.Text = String.Empty
        txtRomaneio.Text = String.Empty
        chkFisico.Checked = False

        txtCliente.Text = String.Empty
        txtDeposito.Text = String.Empty
        txtProduto.Text = String.Empty
        txtDescricaoProduto.Text = String.Empty

        txtOperacao.Text = String.Empty
        txtSubOperacao.Text = String.Empty
        txtDescricaoOperacao.Text = String.Empty

        txtClassificacao.Text = String.Empty
        txtDescricaoClassificacao.Text = String.Empty

        txtPedido.Text = String.Empty

        txtBruto.Text = String.Empty
        txtLiquido.Text = String.Empty
        txtQuantidade.Text = String.Empty

        txtLaudo.ReadOnly = False

        GridDescontos.DataSource = Nothing
        GridDescontos.DataBind()

        GridRateio.DataSource = Nothing
        GridRateio.DataBind()

        txtCodigoCliente.Value = String.Empty
        txtCodigoDeposito.Value = String.Empty
        txtCodigoOperacao.Value = String.Empty
        txtCodigoDestinatario.Value = String.Empty
        txtPedidoDestinatario.Text = String.Empty
        ddlDepositoDestinatario.Items.Clear()
        ddlDepositoDestinatario.Enabled = False
        txtPedidoSaldo.Value = String.Empty
        txtPedidoEntregue.Value = String.Empty
        txtCodigoProdutoPedido.Value = String.Empty
        txtCodigoAutorizacaoDestinatario.Value = 0
        txtNotaFiscalAlt.Text = String.Empty
        DdlCopias.Enabled = False
        ddlOperacao.Items.Clear()
        ddlOperacao.Enabled = False

        cmdCliente.Enabled = False
        lnkConfirmar.Parent.Visible = False
        txtQuantidade.Enabled = False
        lnkFinalizar.Parent.Visible = False
        lnkExcluirRateio.Parent.Visible = False

        Session.Remove("objEmpresaRXT" & HID.Value)
        Session.Remove("objClienteRT" & HID.Value)
        Session.Remove("objItensPedidoSelecionadosRxT" & HID.Value)
        Session.Remove("Autorizacao" & HID.Value)
        Session.Remove("dtRomaneios" & HID.Value)
        Session.Remove("dtRomaneiosAlt" & HID.Value)
        Session.Remove("NotaParaVinculoDeRateio" & HID.Value)

        Dim dtRomaneios As New DataTable()
        dtRomaneios.Columns.Add("Codigo")
        dtRomaneios.Columns.Add("Pedido")
        dtRomaneios.Columns.Add("Deposito")
        dtRomaneios.Columns.Add("Produto")
        dtRomaneios.Columns.Add("Op")
        dtRomaneios.Columns.Add("SO")
        dtRomaneios.Columns.Add("Autorizacao")
        dtRomaneios.Columns.Add("Bruto")
        dtRomaneios.Columns.Add("Desconto")
        dtRomaneios.Columns.Add("Liquido")

        HID.Value = Guid.NewGuid().ToString()
        ucConsultaEmpresas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
        ucConsultaPedidosXNotas.SetarHID(HID.Value)

        LiberaEmpresa()

        Session("dtRomaneios" & HID.Value) = dtRomaneios
        Session("dtRomaneiosAlt" & HID.Value) = dtRomaneios
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            btnEmpresa.Enabled = False
            btnEmpresaAlt.Enabled = False
        End If
    End Sub

    Private Sub Consultar_Laudo()
        Empresa = txtCodigoEmpresa.Value.ToString.Split("-")

        Dim objPesagem As New [Lib].Negocio.Pesagem(Empresa(0), Empresa(1), txtLaudo.Text)

        If objPesagem.Codigo = 0 Then
            MsgBox(Me.Page, "Laudo de pesagem não encontrado.")
            txtLaudo.Focus()
        ElseIf objPesagem.PrimeiraPesagem = 0 Or objPesagem.SegundaPesagem = 0 Then
            MsgBox(Me.Page, "Laudo não foi encerrado.")
            txtLaudo.Focus()
        ElseIf Not objPesagem.CodigoSituacao = 1 Then
            MsgBox(Me.Page, "Laudo de pesagem com situação diferente de normal não pode ser usado.")
            txtLaudo.Focus()
        Else
            btnEmpresa.Enabled = False
            txtPedido.Text = objPesagem.CodigoPedido
            txtOperacao.Text = objPesagem.CodigoOperacao
            txtSubOperacao.Text = objPesagem.CodigoSubOperacao
            txtDescricaoOperacao.Text = objPesagem.SubOperacao.Descricao
            txtCodigoOperacao.Value = objPesagem.CodigoOperacao & "-" & objPesagem.CodigoSubOperacao & "-" & objPesagem.EntradaSaida & "-" & objPesagem.SubOperacao.Devolucao & "-" & objPesagem.SubOperacao.Classe
            txtMovimento.Text = objPesagem.Movimento.ToString("dd/MM/yyyy")
            txtProduto.Text = objPesagem.CodigoProduto
            txtDescricaoProduto.Text = objPesagem.Produto.Descricao
            txtClassificacao.Text = objPesagem.CodigoTabelaDeClassificacao
            txtDescricaoClassificacao.Text = objPesagem.TabelaDeClassificacao.Descricao

            txtCodigoCliente.Value = objPesagem.CodigoCliente & "-" & objPesagem.EnderecoCliente
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objPesagem.Cliente)
            txtCliente.Text = itemCliente.Text

            txtCodigoDeposito.Value = objPesagem.CodigoDeposito & "-" & objPesagem.EnderecoDeposito
            Dim itemDeposito As ListItem = Funcoes.FormatarListItemCliente(objPesagem.Deposito)
            txtDeposito.Text = itemDeposito.Text

            txtBruto.Text = objPesagem.BrutoBalanca
            txtLiquido.Text = objPesagem.Liquido

            '------Analises------------------------------------
            GridDescontos.DataSource = objPesagem.Analises.ToArray()
            GridDescontos.DataBind()

            If GridDescontos.Rows.Count > 0 Then divAnalises.Visible = True
            '------Romaneio------------------------------------

            Dim listRomXPesagem As ListRomaneioXPesagem = objPesagem.Romaneios(0).Pesagens()

            If listRomXPesagem IsNot Nothing AndAlso listRomXPesagem.Count > 0 Then
                txtRomaneio.Text = listRomXPesagem(0).CodigoRomaneio

                For Each row_RXP As [Lib].Negocio.RomaneioXPesagem In listRomXPesagem
                    Dim rom As New Romaneio(objPesagem.CodigoEmpresa, objPesagem.EnderecoEmpresa, row_RXP.CodigoRomaneio)
                    If rom.Codigo > 0 Then
                        txtCodigoAutorizacao.Value = rom.CodigoAutorizacao
                    End If
                Next
            End If

            For Each row As [Lib].Negocio.Romaneio In objPesagem.Romaneios
                Dim newRow As DataRow = CType(Session("dtRomaneios" & HID.Value), DataTable).NewRow
                newRow.Item("Codigo") = Funcoes.FormatarCpfCnpj(objPesagem.CodigoDepositante) & "-" & objPesagem.EnderecoDepositante & " " & Trim(objPesagem.Depositante.Nome)
                newRow.Item("Pedido") = row.CodigoPedido
                newRow.Item("Deposito") = Funcoes.FormatarCpfCnpj(row.CodigoDeposito) & "-" & row.EnderecoDeposito & " " & Trim(row.Deposito.Nome)
                newRow.Item("Produto") = row.CodigoProduto
                newRow.Item("OP") = row.CodigoOperacao
                newRow.Item("SO") = row.CodigoSubOperacao
                newRow.Item("Autorizacao") = row.CodigoAutorizacao
                newRow.Item("Bruto") = row.PesoBruto
                newRow.Item("Desconto") = row.Desconto
                newRow.Item("Liquido") = row.PesoLiquido
                CType(Session("dtRomaneios" & HID.Value), DataTable).Rows.Add(newRow)
            Next

            GridRateio.DataSource = CType(Session("dtRomaneios" & HID.Value), DataTable) 'objPesagem.Romaneios
            GridRateio.DataBind()

            divRateio.Visible = True

            If objPesagem.TemNota Then
                divRateio.Visible = False
                cmdCliente.Enabled = False
                lnkConfirmar.Parent.Visible = False
                lnkFinalizar.Parent.Visible = False
                lnkExcluirRateio.Parent.Visible = False
                txtQuantidade.Enabled = False
                lnkRelatorio.Parent.Visible = True
                DdlCopias.Enabled = False

                MsgBox(Me.Page, "Laudo de pesagem vinculado com nota fiscal não pode ser feito rateio.")
            ElseIf objPesagem.TemRateio Then
                cmdCliente.Enabled = False
                lnkConfirmar.Parent.Visible = False
                lnkFinalizar.Parent.Visible = False
                lnkExcluirRateio.Parent.Visible = False
                txtQuantidade.Enabled = False
                lnkExcluir.Parent.Visible = True
                lnkRelatorio.Parent.Visible = True
                DdlCopias.Enabled = False

                'Limpar()

                MsgBox(Me.Page, "Laudo já possui Rateio.")
            Else
                cmdCliente.Enabled = True
                DdlCopias.Enabled = True
                Sql = "  DELETE AuxiliarXRateio" & vbCrLf & _
                    "     WHERE Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                    "       AND EndEmpresa_Id = " & Empresa(1) & vbCrLf
                SqlArray.Add(Sql)

                Cliente = txtCodigoCliente.Value.ToString.Split("-")
                Deposito = txtCodigoDeposito.Value.ToString.Split("-")

                Sql = "  INSERT INTO AuxiliarXRateio" & vbCrLf & _
                    "                (Empresa_Id, EndEmpresa_Id, Laudo_Id, Cliente_Id, EndCliente_Id, Pedido_Id, Operacao, SubOperacao, Percentual, Bruto, Desconto, Liquido, Produto, Autorizacao, Deposito, EndDeposito) " & vbCrLf & _
                    "           VALUES  ('" & Empresa(0) & "', " & Empresa(1) & ", " & txtLaudo.Text & ", '" & Cliente(0) & "', " & Cliente(1) & ", " & txtPedido.Text & ", " & txtOperacao.Text & ", " & vbCrLf & _
                    "                       " & txtSubOperacao.Text & ", 1.00, " & txtBruto.Text & ", " & CDec(txtBruto.Text) - CDec(txtLiquido.Text) & ", " & txtLiquido.Text & ", '" & txtProduto.Text & "', " & txtCodigoAutorizacao.Value & ", '" & Deposito(0) & "', " & Deposito(1) & ")" & vbCrLf

                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, "Erro ao Gravar Registro(Consultar_Laudo): " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte.")
                End If
            End If
        End If
    End Sub

    Protected Sub cmdCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteRXP" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objEmpresaPXR" & HID.Value.ToString) Is Nothing Then
            SetarEmpresa(CType(Session("objEmpresaPXR" & HID.Value), [Lib].Negocio.Cliente).Codigo, CType(Session("objEmpresaPXR" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco)
            Session.Remove("objEmpresaPXR" & HID.Value)
        ElseIf Not Session("objEmpresaPXRAlt" & HID.Value.ToString) Is Nothing Then
            SetarEmpresaAlt(CType(Session("objEmpresaPXRAlt" & HID.Value), [Lib].Negocio.Cliente).Codigo, CType(Session("objEmpresaPXRAlt" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco)
            Session.Remove("objEmpresaPXRAlt" & HID.Value)
        ElseIf Not Session("objClienteRXP" & HID.Value) Is Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteRXP" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtDestinatario.Text = itemCliente.Text
            txtCodigoDestinatario.Value = itemCliente.Value
            Session.Remove("objClienteRXP" & HID.Value)
            ListarPedidos()
        ElseIf Not Session("objClienteConsRXP" & HID.Value) Is Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteConsRXP" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtClienteSel.Text = itemCliente.Text
            txtCodigoClienteSel.Value = itemCliente.Value
            Session.Remove("objClienteConsRXP" & HID.Value)
            ListarPedidosSel()
        ElseIf Not Session("objPedidoRxP" & HID.Value) Is Nothing Then
            Empresa = txtCodigoEmpresa.Value.ToString.Split("-")
            If Empresa(0) = CType(Session("objPedidoRxP" & HID.Value), [Lib].Negocio.Pedido).CodigoEmpresa And Empresa(1) = CType(Session("objPedidoRxP" & HID.Value), [Lib].Negocio.Pedido).EnderecoEmpresa Then
                txtPedidoDestinatario.Text = CType(Session("objPedidoRxP" & HID.Value), [Lib].Negocio.Pedido).Codigo

                ddlDepositoDestinatario.Enabled = True
                ddlDepositoDestinatario.Items.Clear()
                Funcoes.InserirLinhaEmBranco(ddlDepositoDestinatario)
                Dim coddep As String = String.Empty
                Dim endcoddep As String = String.Empty

                For Each row In CType(Session("objPedidoRxP" & HID.Value), [Lib].Negocio.Pedido).Depositos
                    If row.Tipo = "DE" Then
                        ddlDepositoDestinatario.Items.Add(New ListItem(Funcoes.FormatarCpfCnpj(row.Codigo) & "-" & row.CodigoEndereco & " " & row.Deposito.Nome, row.Codigo & "-" & row.CodigoEndereco))

                        If row.Principal Then
                            coddep = row.Codigo
                            endcoddep = row.CodigoEndereco
                        End If
                    End If
                Next

                Dim intDep As Integer = ddlDepositoDestinatario.Items.IndexOf(ddlDepositoDestinatario.Items.FindByValue(coddep & "-" & endcoddep))
                If Not intDep = -1 Then
                    ddlDepositoDestinatario.SelectedIndex = intDep
                End If

                Dim oPe() As String = txtCodigoOperacao.Value.ToString.Split("-")
                Dim subOperacoes As New [Lib].Negocio.ListSubOperacao(" So.Operacao_Id = " & CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).CodigoOperacao & " AND So.EntradaSaida = '" & oPe(2) & "'")
                ddlOperacao.Enabled = True
                ddlOperacao.Items.Clear()
                ddlOperacao.Items.Insert(0, "")
                For Each objSubOperacao As SubOperacao In subOperacoes
                    ddlOperacao.Items.Add(New ListItem(objSubOperacao.Operacao.Codigo.ToString("00") & "-" & objSubOperacao.Codigo.ToString("00") & " " & objSubOperacao.Descricao, _
                                                          objSubOperacao.Operacao.Codigo & "-" & objSubOperacao.Codigo))
                Next

                Dim intIndice As Integer = ddlOperacao.Items.IndexOf(ddlOperacao.Items.FindByValue(CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).CodigoOperacao & "-" & CType(Session("objPedidoRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).CodigoSubOperacao))
                If Not intIndice = -1 Then
                    ddlOperacao.SelectedIndex = intIndice
                End If

                txtCodigoProdutoPedido.Value = CType(Session("objPedidoRxP" & HID.Value), [Lib].Negocio.Pedido).Itens(0).CodigoProduto
                divMenuRateio.Visible = True
                lnkConfirmar.Parent.Visible = True
                txtQuantidade.Enabled = True
                txtQuantidade.Focus()
            Else
                MsgBox(Me.Page, "Empresa do Pedido selecionado é diferente da Empresa atual.")
            End If
            Session.Remove("objPedidoRxP" & HID.Value)
        ElseIf Not Session("objPedidoSelRxP" & HID.Value) Is Nothing Then
            Dim PedidoSel As Pedido = Session("objPedidoSelRxP" & HID.Value)

            Empresa = txtCodigoEmpresaAlt.Value.ToString.Split("-")
            If Empresa(0) = PedidoSel.CodigoEmpresa And Empresa(1) = PedidoSel.EnderecoEmpresa Then
                txtPedidoSel.Text = PedidoSel.Codigo
                'Dim saldoPedidoSel As New ListSaldoPedidoXNota(Empresa(0), Empresa(1), CType(Session("objPedidoSelRxP" & HID.Value), [Lib].Negocio.Pedido).CodigoCarteira, CType(Session("objPedidoSelRxP" & HID.Value), [Lib].Negocio.Pedido).EnderecoCliente, "", "", "", CType(Session("objPedidoSelRxP" & HID.Value), [Lib].Negocio.Pedido).Codigo, ListSaldoPedidoXNota.Situacao.Todos, True, CType(Session("objPedidoSelRxP" & HID.Value), [Lib].Negocio.Pedido).CodigoOperacao, CType(Session("objPedidoSelRxP" & HID.Value), [Lib].Negocio.Pedido).CodigoSubOperacao)
                Dim Parametros As New Hashtable
                Parametros.Add("Empresa", Empresa(0))
                Parametros.Add("EndEmpresa", Empresa(1))
                Parametros.Add("Pedido", txtPedidoSel.Text)
                Dim SaldoPedido As New ListSaldoPedido2015(Parametros)
                txtPedidoSaldoSel.Value = SaldoPedido(0).SaldoQtdeDiretoFisica

                txtCodigoAutorizacaoAlt.Value = 0

                If CInt(gridRomaneios.Rows(gridRomaneios.SelectedIndex).Cells(11).Text) > txtPedidoSaldoSel.Value AndAlso SaldoPedido(0).Tipo = 1 Then 'Testará o saldo somente se o tipo do Saldo pedido for 1 (Direta/Global)'
                    MsgBox(Me.Page, "Saldo " & String.Format("{0:N0}", CDec(txtPedidoSaldoSel.Value)) & " do Pedido é insuficiente para novo rateio")
                Else
                    Dim oPe() As String = txtCodigoOperacaoAlt.Value.ToString.Split("-")
                    Dim objOpe = New [Lib].Negocio.SubOperacao(gridRomaneios.Rows(gridRomaneios.SelectedIndex).Cells(6).Text, gridRomaneios.Rows(gridRomaneios.SelectedIndex).Cells(7).Text)
                    Dim EntSai As String = IIf(objOpe.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                    Dim subOperacoes As New [Lib].Negocio.ListSubOperacao(" So.Operacao_Id = " & CType(Session("objPedidoSelRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).CodigoOperacao & " AND So.EntradaSaida = '" & EntSai & "'")

                    ddlDepositoDestinatarioSel.Enabled = True
                    ddlDepositoDestinatarioSel.Items.Clear()
                    Funcoes.InserirLinhaEmBranco(ddlDepositoDestinatarioSel)
                    Dim coddep As String = String.Empty
                    Dim endcoddep As String = String.Empty

                    For Each row In CType(Session("objPedidoSelRxP" & HID.Value), [Lib].Negocio.Pedido).Depositos
                        If row.Tipo = "DE" Then
                            ddlDepositoDestinatarioSel.Items.Add(New ListItem(Funcoes.FormatarCpfCnpj(row.Codigo) & "-" & row.CodigoEndereco & " " & row.Deposito.Nome, row.Codigo & "-" & row.CodigoEndereco))

                            If row.Principal Then
                                coddep = row.Codigo
                                endcoddep = row.CodigoEndereco
                            End If
                        End If
                    Next

                    Dim intDep As Integer = ddlDepositoDestinatarioSel.Items.IndexOf(ddlDepositoDestinatarioSel.Items.FindByValue(coddep & "-" & endcoddep))
                    If Not intDep = -1 Then
                        ddlDepositoDestinatarioSel.SelectedIndex = intDep
                    End If

                    ddlOperacaoSel.Enabled = True
                    ddlOperacaoSel.Items.Clear()
                    ddlOperacaoSel.Items.Insert(0, "")
                    For Each objSubOperacao As SubOperacao In subOperacoes
                        ddlOperacaoSel.Items.Add(New ListItem(objSubOperacao.Operacao.Codigo.ToString("00") & "-" & objSubOperacao.Codigo.ToString("00") & " " & objSubOperacao.Descricao, _
                                                              objSubOperacao.Operacao.Codigo & "-" & objSubOperacao.Codigo & "-" & EntSai))
                    Next

                    Dim intIndice As Integer = ddlOperacaoSel.Items.IndexOf(ddlOperacaoSel.Items.FindByValue(CType(Session("objPedidoSelRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).CodigoOperacao & "-" & CType(Session("objPedidoSelRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).CodigoSubOperacao & "-" & IIf(CType(Session("objPedidoSelRxP" & HID.Value.ToString), [Lib].Negocio.Pedido).SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))
                    If Not intIndice = -1 Then
                        ddlOperacaoSel.SelectedIndex = intIndice

                        If PedidoSel.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
                            Dim parameters As New Dictionary(Of String, Object)
                            parameters.Add("ped", txtPedidoSel.Text)
                            parameters.Add("emp", Empresa(0))
                            parameters.Add("endemp", Empresa(1))
                            parameters.Add("cli", "")
                            parameters.Add("endcli", "")
                            parameters.Add("romaneio", True)
                            parameters.Add("classe", CType(Session("objPedidoSelRxP" & HID.Value), [Lib].Negocio.Pedido).SubOperacao.Classe)
                            ucConsultaAutorizacaoDeRetirada.BindGridView(parameters)
                            Popup.ConsultaDeAutorizacaoDeRetirada(Me.Page, "objSelRateioDePesagem" & HID.Value)
                        Else
                            btnAltRomaneio.Enabled = True
                            btnBuscarNF.Enabled = True
                        End If
                    End If
                End If
                Session.Remove("objPedidoSelRxP" & HID.Value)
            Else
                MsgBox(Me.Page, "Empresa do pedido selecionado é diferente da empresa atual.")
            End If
            Session.Remove("objPedidoSelRxP" & HID.Value)
        End If
    End Sub

    Public Sub CarregarNotaFiscal()
        Try
            If Not Session("objNFConsultaRAT" & HID.Value) Is Nothing Then
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsultaRAT" & HID.Value), NotaFiscal))

                If objNotaFiscal.Codigo > 0 Then
                    If Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento.ToString("dd-MM-yyyy"), "NOTAS FISCAIS") Then
                        MsgBox(Me.Page, "Movimento de Notas Fiscais já Fechado para esta data...")
                    ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento.ToString("dd-MM-yyyy"), "PRODUCAO") Then
                        MsgBox(Me.Page, "Movimento de Produção já Fechado para esta data....")
                    ElseIf Session("NotaParaVinculoDeRateio" & HID.Value) = "N" AndAlso objNotaFiscal.CodigoRomaneio = 0 Then
                        MsgBox(Me.Page, "Nota Fiscal não possui Romaneio vinculado.")
                    ElseIf objNotaFiscal.CodigoRomaneio > 0 AndAlso Not objNotaFiscal.Romaneio.Processo = "NOTA FISCAL" Then
                        MsgBox(Me.Page, "Nota Fiscal não pode ser utilizada pois já está vinculada com uma Pesagem.")
                    Else

                        If Session("NotaParaVinculoDeRateio" & HID.Value) = "S" Then
                            txtNotaFiscalAlt.Text = objNotaFiscal.Codigo
                            btnGravarNF.Visible = True
                            btnBuscarNF.Enabled = False
                            btnAltRomaneio.Enabled = False
                        Else
                            lnkConsultarAR.Parent.Visible = False
                            lnkConsultarNF.Parent.Visible = False
                            txtNotaAlt.Enabled = False
                            txtLaudoAlt.Enabled = False
                            btnEmpresaAlt.Enabled = False
                            txtCodigoEmpresaAlt.Value = objNotaFiscal.CodigoEmpresa & "-" & objNotaFiscal.EnderecoEmpresa
                            txtMovimentoAlt.Text = objNotaFiscal.Movimento.ToString("dd/MM/yyyy")
                            txtCodigoClienteAlt.Value = objNotaFiscal.CodigoCliente & "-" & objNotaFiscal.EnderecoCliente
                            Dim itemClienteAlt As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Cliente)
                            txtClienteAlt.Text = itemClienteAlt.Text
                            txtPedidoAlt.Text = objNotaFiscal.CodigoPedido
                            txtCodigoOperacaoAlt.Value = objNotaFiscal.CodigoOperacao & "-" & objNotaFiscal.CodigoSubOperacao & "-" & objNotaFiscal.EntradaSaida & "-" & objNotaFiscal.SubOperacao.Devolucao & "-" & objNotaFiscal.SubOperacao.Classe
                            txtProdutoAlt.Value = objNotaFiscal.Itens(0).CodigoProduto

                            Dim listRomXPesagem As ListRomaneioXPesagem = objNotaFiscal.Romaneio.Pesagens()

                            txtCodigoAutorizacaoAlt.Value = objNotaFiscal.Romaneio.CodigoAutorizacao

                            Dim dtRomaneios As New DataTable()
                            dtRomaneios.Columns.Add("Romaneio")
                            dtRomaneios.Columns.Add("Codigo")
                            dtRomaneios.Columns.Add("Pedido")
                            dtRomaneios.Columns.Add("Deposito")
                            dtRomaneios.Columns.Add("Produto")
                            dtRomaneios.Columns.Add("Op")
                            dtRomaneios.Columns.Add("SO")
                            dtRomaneios.Columns.Add("Autorizacao")
                            dtRomaneios.Columns.Add("Bruto")
                            dtRomaneios.Columns.Add("Desconto")
                            dtRomaneios.Columns.Add("Liquido")

                            Session("dtRomaneiosAlt" & HID.Value) = dtRomaneios

                            Dim newRow As DataRow = CType(Session("dtRomaneiosAlt" & HID.Value), DataTable).NewRow
                            newRow.Item("Romaneio") = objNotaFiscal.Romaneio.Codigo
                            newRow.Item("Codigo") = Funcoes.FormatarCpfCnpj(objNotaFiscal.CodigoCliente) & "-" & objNotaFiscal.EnderecoCliente & " " & objNotaFiscal.Cliente.Nome
                            newRow.Item("Pedido") = objNotaFiscal.CodigoPedido
                            newRow.Item("Deposito") = Funcoes.FormatarCpfCnpj(objNotaFiscal.CodigoDeposito) & "-" & objNotaFiscal.EnderecoDeposito & " " & objNotaFiscal.Deposito.Nome
                            newRow.Item("Produto") = objNotaFiscal.Itens(0).CodigoProduto
                            newRow.Item("OP") = objNotaFiscal.CodigoOperacao
                            newRow.Item("SO") = objNotaFiscal.CodigoSubOperacao
                            newRow.Item("Autorizacao") = objNotaFiscal.CodigoAutorizacao
                            newRow.Item("Bruto") = objNotaFiscal.Romaneio.PesoBruto
                            newRow.Item("Desconto") = objNotaFiscal.Romaneio.Desconto
                            newRow.Item("Liquido") = objNotaFiscal.Romaneio.PesoLiquido
                            CType(Session("dtRomaneiosAlt" & HID.Value), DataTable).Rows.Add(newRow)

                            gridRomaneios.DataSource = CType(Session("dtRomaneiosAlt" & HID.Value), DataTable)
                            gridRomaneios.DataBind()

                            CType(gridRomaneios.Rows(0).FindControl("imgRemoverRomaneio"), ImageButton).Visible = True

                        End If

                        SalvaNotaFiscal()
                    End If
                Else
                    MsgBox(Me.Page, "Houve um problema ao selecionar a Nota Fiscal, refaça o processo!")
                End If
            ElseIf Not Session("objNFConsultaRATNOT" & HID.Value) Is Nothing Then
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsultaRATNOT" & HID.Value), NotaFiscal))

                If objNotaFiscal.Codigo > 0 Then
                    If Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento.ToString("dd-MM-yyyy"), "NOTAS FISCAIS") Then
                        MsgBox(Me.Page, "Movimento de Notas Fiscais já Fechado para esta data...")
                    ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento.ToString("dd-MM-yyyy"), "PRODUCAO") Then
                        MsgBox(Me.Page, "Movimento de Produção já Fechado para esta data....")
                    ElseIf objNotaFiscal.CodigoRomaneio = 0 Then
                        MsgBox(Me.Page, "Nota Fiscal não possui Romaneio vinculado.")
                    ElseIf objNotaFiscal.CodigoRomaneio > 0 AndAlso Not objNotaFiscal.Romaneio.Processo = "NOTA FISCAL" Then
                        MsgBox(Me.Page, "Nota Fiscal não pode ser utilizada pois já está vinculada com uma Pesagem.")
                    Else

                        'lnkConsultarAR.Parent.Visible = False
                        'lnkConsultarNF.Parent.Visible = False
                        'txtNotaAlt.Enabled = False
                        'txtLaudoAlt.Enabled = False
                        'btnEmpresaAlt.Enabled = False
                        'txtCodigoEmpresaAlt.Value = objNotaFiscal.CodigoEmpresa & "-" & objNotaFiscal.EnderecoEmpresa
                        'txtMovimentoAlt.Text = objNotaFiscal.Movimento.ToString("dd/MM/yyyy")
                        'txtCodigoClienteAlt.Value = objNotaFiscal.CodigoCliente & "-" & objNotaFiscal.EnderecoCliente
                        'Dim itemClienteAlt As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Cliente)
                        'txtClienteAlt.Text = itemClienteAlt.Text
                        'txtPedidoAlt.Text = objNotaFiscal.CodigoPedido
                        'txtCodigoOperacaoAlt.Value = objNotaFiscal.CodigoOperacao & "-" & objNotaFiscal.CodigoSubOperacao & "-" & objNotaFiscal.EntradaSaida & "-" & objNotaFiscal.SubOperacao.Devolucao & "-" & objNotaFiscal.SubOperacao.Classe
                        'txtProdutoAlt.Value = objNotaFiscal.Itens(0).CodigoProduto

                        'Dim listRomXPesagem As ListRomaneioXPesagem = objNotaFiscal.Romaneio.Pesagens()

                        'txtCodigoAutorizacaoAlt.Value = objNotaFiscal.Romaneio.CodigoAutorizacao

                        'Dim dtRomaneios As New DataTable()
                        'dtRomaneios.Columns.Add("Romaneio")
                        'dtRomaneios.Columns.Add("Codigo")
                        'dtRomaneios.Columns.Add("Pedido")
                        'dtRomaneios.Columns.Add("Deposito")
                        'dtRomaneios.Columns.Add("Produto")
                        'dtRomaneios.Columns.Add("Op")
                        'dtRomaneios.Columns.Add("SO")
                        'dtRomaneios.Columns.Add("Autorizacao")
                        'dtRomaneios.Columns.Add("Bruto")
                        'dtRomaneios.Columns.Add("Desconto")
                        'dtRomaneios.Columns.Add("Liquido")

                        'Session("dtRomaneiosAlt" & HID.Value) = dtRomaneios

                        'Dim newRow As DataRow = CType(Session("dtRomaneiosAlt" & HID.Value), DataTable).NewRow
                        'newRow.Item("Romaneio") = objNotaFiscal.Romaneio.Codigo
                        'newRow.Item("Codigo") = Funcoes.FormatarCpfCnpj(objNotaFiscal.CodigoCliente) & "-" & objNotaFiscal.EnderecoCliente & " " & objNotaFiscal.Cliente.Nome
                        'newRow.Item("Pedido") = objNotaFiscal.CodigoPedido
                        'newRow.Item("Deposito") = Funcoes.FormatarCpfCnpj(objNotaFiscal.CodigoDeposito) & "-" & objNotaFiscal.EnderecoDeposito & " " & objNotaFiscal.Deposito.Nome
                        'newRow.Item("Produto") = objNotaFiscal.Itens(0).CodigoProduto
                        'newRow.Item("OP") = objNotaFiscal.CodigoOperacao
                        'newRow.Item("SO") = objNotaFiscal.CodigoSubOperacao
                        'newRow.Item("Autorizacao") = objNotaFiscal.CodigoAutorizacao
                        'newRow.Item("Bruto") = objNotaFiscal.Romaneio.PesoBruto
                        'newRow.Item("Desconto") = objNotaFiscal.Romaneio.Desconto
                        'newRow.Item("Liquido") = objNotaFiscal.Romaneio.PesoLiquido
                        'CType(Session("dtRomaneiosAlt" & HID.Value), DataTable).Rows.Add(newRow)

                        'gridRomaneios.DataSource = CType(Session("dtRomaneiosAlt" & HID.Value), DataTable)
                        'gridRomaneios.DataBind()

                        'CType(gridRomaneios.Rows(0).FindControl("imgRemoverRomaneio"), ImageButton).Visible = True

                        SalvaNotaFiscal()
                    End If
                Else
                    MsgBox(Me.Page, "Houve um problema ao selecionar a Nota Fiscal, refaça o processo!")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub SetarEmpresaAlt(ByVal Empresa As String, ByVal EndEmpresa As String)
        Dim objEmpresa As New [Lib].Negocio.Cliente(Empresa, EndEmpresa)
        Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objEmpresa)

        txtEmpresaAlt.Text = itemEmpresa.Text
        txtCodigoEmpresaAlt.Value = itemEmpresa.Value
    End Sub

    Private Sub ListarPedidos()
        Dim empresa() As String = txtCodigoEmpresa.Value.Split("-")
        Destinatario = txtCodigoDestinatario.Value.ToString.Split("-")
        Dim objPedido = New [Lib].Negocio.Pedido(empresa(0), empresa(1), txtPedido.Text)
        Dim parameters As New Dictionary(Of String, Object)

        parameters("empresa") = empresa(0)
        parameters("enderecoEmpresa") = empresa(1)
        parameters("cliente") = Destinatario(0)
        parameters("enderecoCliente") = Destinatario(1)
        parameters("produto") = txtProduto.Text
        parameters("es") = IIf(objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
        parameters("situacao") = 1
        Session("ssCampo" & HID.Value) = "Pedidos"
        Popup.ConsultaDePedidos(Me.Page, "objPedidoRxP" & HID.Value.ToString, "txtPedido")
        ucConsultaPedidos.BindGridView(parameters)
    End Sub

    Private Sub ListarPedidosSel()
        Dim empresa() As String = txtCodigoEmpresaAlt.Value.Split("-")
        Destinatario = txtCodigoClienteSel.Value.ToString.Split("-")
        Dim Operacao() As String = txtCodigoOperacaoAlt.Value.ToString.Split("-")
        Dim parameters As New Dictionary(Of String, Object)

        RecuperaNotaFiscal()

        parameters("empresa") = empresa(0)
        parameters("enderecoEmpresa") = empresa(1)
        parameters("cliente") = Destinatario(0)
        parameters("enderecoCliente") = Destinatario(1)
        parameters("produto") = txtProdutoAlt.Value
        parameters("es") = Operacao(2)
        parameters("situacao") = 1

        If Not objNotaFiscal Is Nothing AndAlso Not objNotaFiscal.Pedido Is Nothing AndAlso Not objNotaFiscal.Pedido.PedidoEfetivo Is Nothing Then
            parameters("pedidoefetivo") = objNotaFiscal.Pedido.PedidoEfetivo
        End If

        Session("ssCampo" & HID.Value) = "Pedidos"
        Popup.ConsultaDePedidos(Me.Page, "objPedidoSelRxP" & HID.Value.ToString, "txtPedido")

        ucConsultaPedidos.BindGridView(parameters)
    End Sub

    Private Sub GravaParcial()
        Empresa = txtCodigoEmpresa.Value.ToString.Split("-")
        Destinatario = txtCodigoDestinatario.Value.ToString.Split("-")
        CodigoOperacao = ddlOperacao.SelectedValue.ToString.Split("-")
        Deposito = ddlDepositoDestinatario.SelectedValue.ToString.Split("-")
        Dim ParcialPercentual As Decimal = CInt(txtQuantidade.Text) / Liquido
        Dim ParcialBruto As Integer = CInt(Bruto * ParcialPercentual)
        Dim ParcialDesconto As Integer = CInt(Bruto * ParcialPercentual) - CInt(txtQuantidade.Text)

        Sql = "  INSERT INTO AuxiliarXRateio" & vbCrLf & _
              " (Empresa_Id" & vbCrLf & _
              ", EndEmpresa_Id" & vbCrLf & _
              ", Laudo_Id" & vbCrLf & _
              ", Cliente_Id " & vbCrLf & _
              ", EndCliente_Id" & vbCrLf & _
              ", Pedido_Id" & vbCrLf & _
              ", Operacao" & vbCrLf & _
              ", SubOperacao" & vbCrLf & _
              ", Percentual" & vbCrLf & _
              ", Bruto" & vbCrLf & _
              ", Desconto" & vbCrLf & _
              ", Liquido" & vbCrLf & _
              ", Produto" & vbCrLf & _
              ", Autorizacao" & vbCrLf & _
              ", Deposito" & vbCrLf & _
              ", EndDeposito)" & vbCrLf & _
              "  VALUES(" & vbCrLf & _
              "'" & Empresa(0) & "'" & vbCrLf & _
              ", " & Empresa(1) & vbCrLf & _
              ", " & GridRateio.Rows.Count - 1 & vbCrLf

        Sql &= ", '" & Destinatario(0) & "'" & vbCrLf & _
               ", " & Destinatario(1) & vbCrLf & _
               ", " & txtPedidoDestinatario.Text & vbCrLf & _
               ", " & CodigoOperacao(0) & vbCrLf & _
               ", " & CodigoOperacao(1) & vbCrLf & _
               ", " & Replace(ParcialPercentual, ",", ".") & vbCrLf & _
               ", " & ParcialBruto & vbCrLf & _
               ", " & ParcialDesconto & vbCrLf & _
               ", " & txtQuantidade.Text & vbCrLf & _
               ", '" & txtCodigoProdutoPedido.Value & "'" & vbCrLf & _
               ", " & txtCodigoAutorizacaoDestinatario.Value & vbCrLf & _
               ", '" & Deposito(0) & "'" & vbCrLf & _
               ", " & Deposito(1) & vbCrLf & _
               ")" & vbCrLf

        SqlArray.Add(Sql)

        Sql = " Update AuxiliarXRateio" & vbCrLf & _
              "    Set Percentual = " & Replace((Percentual - ParcialPercentual), ",", ".") & vbCrLf & _
              " ,Bruto = " & Bruto - ParcialBruto & vbCrLf & _
              " ,Desconto = " & Desconto - ParcialDesconto & vbCrLf & _
              " ,Liquido = " & Liquido - CInt(txtQuantidade.Text) & vbCrLf & _
              " WHERE Empresa_Id = '" & Empresa(0) & "' AND EndEmpresa_Id = " & Empresa(1) & "And Laudo_Id = " & txtLaudo.Text & vbCrLf

        SqlArray.Add(Sql)

        If Banco.GravaBanco(SqlArray) = False Then
            MsgBox(Me.Page, "Erro ao Gravar Registro(GravaParcial): " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte.")
        Else
            Sql = " SELECT AxR.Cliente_Id AS Cliente, AxR.EndCliente_Id AS Endereco, Cli.Nome, AxR.Pedido_Id AS Pedido, AxR.Operacao, " & vbCrLf & _
                  "        AxR.SubOperacao, AxR.Bruto, AxR.Desconto, AxR.Liquido, AxR.Produto, AxR.Autorizacao, AxR.Deposito, AxR.EndDeposito, Dep.Nome as NomeDeposito " & vbCrLf & _
                  "   FROM AuxiliarXRateio AxR " & vbCrLf & _
                  "   INNER JOIN Clientes Cli" & vbCrLf & _
                  "     ON AxR.Cliente_Id    = Cli.Cliente_Id " & vbCrLf & _
                  "    AND AxR.EndCliente_Id = Cli.Endereco_Id " & vbCrLf & _
                  "   LEFT JOIN Clientes Dep" & vbCrLf & _
                  "     ON AxR.Deposito    = Dep.Cliente_Id " & vbCrLf & _
                  "    AND AxR.EndDeposito = Dep.Endereco_Id " & vbCrLf & _
                  "  WHERE Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                  "    AND EndEmpresa_Id = " & Empresa(1) & vbCrLf
            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

            CType(Session("dtRomaneios" & HID.Value), DataTable).Clear()

            For Each row As DataRow In ds.Tables(0).Rows
                Dim newRow As DataRow = CType(Session("dtRomaneios" & HID.Value), DataTable).NewRow
                newRow.Item("Codigo") = Funcoes.FormatarCpfCnpj(row("Cliente")) & "-" & row("Endereco") & " " & row("Nome")
                newRow.Item("Pedido") = row("Pedido")
                newRow.Item("Deposito") = Funcoes.FormatarCpfCnpj(row("Deposito")) & "-" & row("EndDeposito") & " " & row("NomeDeposito")
                newRow.Item("Produto") = row("Produto")
                newRow.Item("OP") = row("Operacao")
                newRow.Item("SO") = row("SubOperacao")
                newRow.Item("Autorizacao") = row("Autorizacao")
                newRow.Item("Bruto") = row("Bruto")
                newRow.Item("Desconto") = row("Desconto")
                newRow.Item("Liquido") = row("Liquido")
                CType(Session("dtRomaneios" & HID.Value), DataTable).Rows.Add(newRow)
            Next

            GridRateio.DataSource = CType(Session("dtRomaneios" & HID.Value), DataTable)
            GridRateio.DataBind()

            txtDestinatario.Text = String.Empty
            txtCodigoDestinatario.Value = String.Empty
            txtPedidoDestinatario.Text = String.Empty
            ddlDepositoDestinatario.Items.Clear()
            txtCodigoAutorizacaoDestinatario.Value = 0
            ddlOperacao.Items.Clear()
            txtQuantidade.Text = String.Empty
        End If
    End Sub

    Public Sub CarregarAutorizacao(par As Hashtable)
        If Not Session("objRateioDePesagem" & HID.Value) Is Nothing Then
            txtCodigoAutorizacaoDestinatario.Value = Session("objRateioDePesagem" & HID.Value)
            Empresa = txtCodigoEmpresa.Value.ToString.Split("-")
            Dim pAutorizacao = New [Lib].Negocio.AutorizacaoDeRetirada(Empresa(0), Empresa(1), par("Pedido"), par("Autorizacao"), True) 'txtCodigoAutorizacaoDestinatario.Value
            Session.Remove("objRateioDePesagem" & HID.Value)

            If CInt(txtQuantidade.Text) > pAutorizacao.SaldoFisico Then
                MsgBox(Me.Page, "Saldo " & pAutorizacao.SaldoFisico.ToString("N0") & " da Autorização " & pAutorizacao.Autorizacao & " do Pedido " & pAutorizacao.CodigoPedido & " selecionando é insuficiente para Rateio.")
            Else
                If VerificaSaldo() Then
                    lnkFinalizar.Parent.Visible = True
                    lnkExcluirRateio.Parent.Visible = True
                    GravaParcial()
                End If
            End If
        ElseIf Not Session("objSelRateioDePesagem" & HID.Value) Is Nothing Then
            txtCodigoAutorizacaoAlt.Value = Session("objSelRateioDePesagem" & HID.Value)
            btnAltRomaneio.Enabled = True
            btnBuscarNF.Enabled = True
            Session.Remove("objSelRateioDePesagem" & HID.Value)
        Else
            MsgBox(Me.Page, "Autorização de retirada não foi encontrada, pedido não pode ser utilizado.")
        End If
    End Sub


    Function VerificaSaldo(Optional ByVal Empresa() As String = Nothing, Optional ByVal Pedido As String = "") As Boolean
        Dim objPedido As [Lib].Negocio.Pedido = Nothing
        If Not Empresa Is Nothing Then
            objPedido = New [Lib].Negocio.Pedido(Empresa(0), Empresa(1), Pedido)
        End If

        If txtDestinatario.Text = "" Or txtQuantidade.Text = "" Then
            MsgBox(Me.Page, "Destinatário e/ou quantidade nao informada")
            Return False
        Else
            Sql = " SELECT Percentual, Bruto, Desconto, Liquido" & vbCrLf & _
                  "   FROM AuxiliarXRateio " & vbCrLf & _
                  "  WHERE Laudo_Id = " & txtLaudo.Text & vbCrLf
            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
                Percentual = Dr("Percentual")
                Bruto = Dr("Bruto")
                Desconto = Dr("Desconto")
                Liquido = Dr("Liquido")
            Next

            If Not objPedido Is Nothing AndAlso objPedido.Itens(0).QuantidadePedidoFaturamento = 0 AndAlso (objPedido.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS Or objPedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR) AndAlso Not objPedido.SubOperacao.Devolucao Then
                Return True
            ElseIf CInt(txtQuantidade.Text) > Liquido Then
                MsgBox(Me.Page, "Saldo insuficiente para novo rateio")
                Return False
            ElseIf CInt(txtQuantidade.Text) > txtPedidoSaldo.Value AndAlso Not CInt(txtPedidoDestinatario.Text) = CInt(txtPedido.Text) Then
                MsgBox(Me.Page, "Saldo " & String.Format("{0:N0}", CDec(txtPedidoSaldo.Value)) & " do Pedido é insuficiente para novo rateio")
                Return False
            Else
                Return True
            End If
        End If
    End Function

    Function VerificaSaldoParaExcluir(ByVal strCliente As String, ByVal strPedido As Integer) As Boolean
        Dim ds As New DataSet

        Sql = " SELECT Bruto, Desconto, Liquido" & vbCrLf & _
              "   FROM AuxiliarXRateio " & vbCrLf & _
              "  WHERE Cliente_Id = '" & strCliente & "'" & vbCrLf & _
              "    AND Pedido_Id = " & strPedido & vbCrLf
        ds = Banco.ConsultaDataSet(Sql, "Consulta")

        For Each Dr As DataRow In ds.Tables(0).Rows
            Bruto = Dr("Bruto")
            Desconto = Dr("Desconto")
            Liquido = Dr("Liquido")
        Next

        Return True
    End Function

    Function GravarRateio() As Boolean
        Dim dsAux As New DataSet
        Dim dsAnalises As New DataSet
        Dim ds As New DataSet
        Dim Romaneio As Integer = 0
        Empresa = txtCodigoEmpresa.Value.Split("-")

        Dim subOperacao() As String = txtCodigoOperacao.Value.ToString.Split("-")

        Sql = " SELECT Empresa_Id, EndEmpresa_Id, Laudo_Id, Cliente_Id, EndCliente_Id, Pedido_Id, Operacao, " & vbCrLf & _
              "        SubOperacao, Bruto, Desconto, Percentual, Liquido, Produto, Autorizacao, Deposito, EndDeposito " & vbCrLf & _
              "   FROM AuxiliarXRateio" & vbCrLf & _
              "  WHERE Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
              "    AND EndEmpresa_Id = " & Empresa(1) & vbCrLf

        dsAux = Banco.ConsultaDataSet(Sql, "ConsultaAuxiliar")

        If dsAux Is Nothing OrElse dsAux.Tables(0).Rows.Count = 0 Then
            MsgBox(Me.Page, "Tabela auxiliar do rateio não encontrada ou com problema, rateio não será feito. entre em contato com o suporte.")
            Return False
        Else
            For Each Dr As DataRow In dsAux.Tables(0).Rows
                If Not IsDBNull(Dr("Liquido")) AndAlso CInt(Dr("Liquido")) > 0 Then
                    Sql = "exec sp_Numerador '" & Empresa(0) & "'," & Empresa(1) & ",110"
                    ds = Banco.ConsultaDataSet(Sql, "Consulta")

                    If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                        MsgBox(Me.Page, "Numerador do Romaneio não encontrado")
                        Return False
                    Else
                        Romaneio = ds.Tables(0).Rows(0).Item(0)
                    End If

                    Sql = " INSERT INTO Romaneios" & vbCrLf & _
                        "           (Empresa_Id, EndEmpresa_Id, Romaneio_Id, EntradaSaida, Pedido, Deposito, EndDeposito, Destino, EndDestino, Transbordo, EndTransbordo, " & vbCrLf & _
                        "            Produto, Operacao, SubOperacao, Movimento, PesoBruto, Desconto, PesoLiquido, Processo, Autorizacao) " & vbCrLf & _
                        "  VALUES ('" & Empresa(0) & "', " & Empresa(1) & ", " & Romaneio & ", '" & subOperacao(2) & "'" & ", " & Dr("Pedido_Id") & ", '" & Dr("Deposito") & "', " & vbCrLf & _
                        "           " & Dr("EndDeposito") & ", '', 0, '', 0, '" & Dr("Produto") & "', " & Dr("Operacao") & ", " & Dr("SubOperacao") & ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "', " & vbCrLf & _
                        "           " & Dr("Bruto") & ", " & Dr("Desconto") & ", " & Dr("Liquido") & ", 'Rateio', " & Dr("Autorizacao") & ")"

                    SqlArray.Add(Sql)

                    Sql = " INSERT INTO RomaneiosXPesagens" & vbCrLf & _
                          "           (Empresa_Id, EndEmpresa_Id, Romaneio_Id, Pesagem_Id, Sequencia_Id)" & vbCrLf & _
                          "   VALUES  ('" & Empresa(0) & "', " & Empresa(1) & ", " & Romaneio & ", " & txtLaudo.Text & ", 0)" & vbCrLf

                    SqlArray.Add(Sql)


                    '------Analises------------------------------------

                    Sql = " SELECT Analise_Id, Percentual, Indice, Desconto" & vbCrLf & _
                          "   FROM PesagemXAnalises" & vbCrLf & _
                          "  WHERE (Empresa_Id = '" & Empresa(0) & "') AND (EndEmpresa_Id = " & Empresa(1) & ") And Pesagem_Id =" & txtLaudo.Text & " And (Sequencia_Id = 0)" & vbCrLf
                    dsAnalises = Banco.ConsultaDataSet(Sql, "ConsultaDescontos")

                    If dsAnalises Is Nothing OrElse dsAnalises.Tables(0).Rows.Count = 0 Then
                        'MsgBox(Me.Page, "Análises para o Romaneio não foram encontradas.")
                        'Return False
                    Else
                        For Each Dra As DataRow In dsAnalises.Tables(0).Rows

                            Desconto = Dra("Desconto") * Dr("Percentual")

                            Sql = "INSERT INTO RomaneiosXDescontos (Empresa_Id, EndEmpresa_Id, Romaneio_Id, Analise_Id, Desconto, Percentual, Indice)" & vbCrLf & _
                                  "VALUES ('" & Empresa(0) & "'," & Empresa(1) & "," & Romaneio & "," & Dra("Analise_Id") & "," & Str(Desconto) & "," & Str(Dra("Percentual")) & "," & Str(Dra("Indice")) & ")"
                            SqlArray.Add(Sql)
                        Next
                    End If
                End If
            Next

            If GridRateio.Rows.Count > 0 Then
                Sql = "DELETE RomaneiosXPesagens " & vbCrLf & _
                      " WHERE Empresa_Id    ='" & Empresa(0) & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & Empresa(1) & vbCrLf & _
                      "   AND Romaneio_Id   = " & txtRomaneio.Text & vbCrLf & _
                      "   AND Pesagem_Id    = " & txtLaudo.Text & vbCrLf & _
                      "   AND Sequencia_Id  = 0" & vbCrLf
                SqlArray.Add(Sql)

                Sql = "DELETE RomaneiosXDescontos " & vbCrLf & _
                      " WHERE Empresa_Id    ='" & Empresa(0) & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & Empresa(1) & vbCrLf & _
                      "   AND Romaneio_Id   = " & txtRomaneio.Text & vbCrLf
                SqlArray.Add(Sql)

                Sql = "DELETE Romaneios " & vbCrLf & _
                      " WHERE Empresa_Id    ='" & Empresa(0) & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & Empresa(1) & vbCrLf & _
                      "   AND Romaneio_Id   = " & txtRomaneio.Text & vbCrLf
                SqlArray.Add(Sql)

                Sql = " DELETE AuxiliarXRateio "
                SqlArray.Add(Sql)
            End If

            If Not Banco.GravaBanco(SqlArray) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Return False
            Else
                MsgBox(Me.Page, "Rateio realizado com Sucesso.", eTitulo.Sucess)
                Return True
            End If
        End If
    End Function

    Private Sub ExcluirRateio()
        Dim SqlArray As New ArrayList
        Empresa = txtCodigoEmpresa.Value.Split("-")

        Dim Laudo As New [Lib].Negocio.Pesagem(Empresa(0), Empresa(1), txtLaudo.Text)

        For Each xRomaneio As Romaneio In Laudo.Romaneios
            If xRomaneio.CodigoPedido = Laudo.CodigoPedido Then Laudo.CodigoAutorizacao = xRomaneio.CodigoAutorizacao
        Next

        Laudo.DesfazerRateio(SqlArray)

        If Banco.GravaBanco(SqlArray) Then
            Limpar()
            MsgBox(Me.Page, "Rateio desfeito com Sucesso.", eTitulo.Sucess)
        Else
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        End If
    End Sub

    Private Function getDataSet(ByVal romaneio As String)
        Dim sql As String = " SELECT R.Romaneio_Id as Romaneio, P.Pesagem_Id AS Laudo, P.Produto, P.Placa, P.EntradaSaida, " & vbCrLf & _
                            "		case" & vbCrLf & _
                            "			when P.EntradaSaida = 'S'" & vbCrLf & _
                            "				then P.PrimeiraPesagem + R.PesoBruto" & vbCrLf & _
                            "				else P.PrimeiraPesagem" & vbCrLf & _
                            "			end as PrimeiraPesagem," & vbCrLf & _
                            "		case" & vbCrLf & _
                            "			when P.EntradaSaida = 'E'" & vbCrLf & _
                            "				then P.PrimeiraPesagem - R.PesoBruto" & vbCrLf & _
                            "				else P.SegundaPesagem" & vbCrLf & _
                            "			end as SegundaPesagem," & vbCrLf & _
                            "        R.PesoBruto AS BrutoBalanca, " & vbCrLf & _
                            "        R.Desconto AS Descontos, R.PesoLiquido AS Liquido, P.EntradaPatio, " & vbCrLf & _
                            "        P.EntradaBalanca, P.SaidaBalanca, P.Movimento, P.NumeroDaNota AS NotaFiscal, " & vbCrLf & _
                            "        P.SerieDaNota AS SerieNota, P.PesoFiscal, P.Observacoes, Produtos.Nome AS NomeProduto, " & vbCrLf & _
                            "        ClientePedido.Cliente_Id AS CodigoCliente, ClientePedido.Endereco_Id AS EndCliente, ClientePedido.Nome AS NomeCliente, " & vbCrLf & _
                            "        ClientePedido.Reduzido AS ReduzidoCliente, ClientePedido.Endereco AS EnderecoCliente, ClientePedido.Cidade AS CidadeCliente, " & vbCrLf & _
                            "        ClientePedido.Estado AS EstadoCliente, Transportes.Cliente_Id AS CodigoTransportador, Transportes.Endereco_Id AS EndTransportador, " & vbCrLf & _
                            "        Transportes.Nome AS NomeTransportador, Transportes.Reduzido AS ReduzidoTransportador, Transportes.Endereco AS EnderecoTransportador, " & vbCrLf & _
                            "        Transportes.Cidade AS CidadeTransportador, Transportes.Estado AS EstadoTransportador, Depositos.Cliente_Id AS CodigoDeposito, " & vbCrLf & _
                            "        Depositos.Endereco_Id AS EndDeposito, Depositos.Nome AS NomeDeposito, Depositos.Reduzido AS ReduzidoDeposito, Depositos.Endereco AS EnderecoDeposito, " & vbCrLf & _
                            "        Depositos.Cidade AS CidadeDeposito, Depositos.Estado AS EstadoDeposito, Depositos.Inscricao AS InscricaoDeposito, Placas.Placa01, " & vbCrLf & _
                            "        Placas.Placa02, Placas.Placa03, Placas.CidadePlaca, Placas.EstadoPlaca, isnull(Motorista.Nome,'') AS NomeMotorista, " & vbCrLf & _
                            "        isnull(Motorista.Cidade,'') AS CidadeMotorista, isnull(Motorista.Estado,'') AS EstadoMotorista, Placas.Habilitacao, CASE WHEN Placas.CpfMotorista IS NULL " & vbCrLf & _
                            "        THEN '' WHEN Placas.CpfMotorista = '' THEN '' ELSE SUBSTRING(Placas.CpfMotorista, 1, 3) + '.' + SUBSTRING(Placas.CpfMotorista, 4, 3) " & vbCrLf & _
                            "        + '.' + SUBSTRING(Placas.CpfMotorista, 7, 3) + '-' + SUBSTRING(Placas.CpfMotorista, 10, 2) END AS CpfMotorista, EstadoPlaca.Descricao AS NomeEstadoPlaca, " & vbCrLf & _
                            "        isnull(EstadoPlacaMotorista.Descricao,'') AS NomeEstadoMotorista, cast(R.Pedido as varchar) as Pedido, Depositos.Numero AS NumeroDeposito, Depositos.Complemento AS ComplementoDeposito, " & vbCrLf & _
                            "        Depositos.Bairro AS BairroDeposito, Clientes.Numero AS NumeroCliente, Clientes.Complemento AS ComplementoCliente, Clientes.Bairro AS BairroCliente, " & vbCrLf & _
                            "        Clientes.Inscricao AS InscricaoCliente, P.Empresa_Id as CodigoEmpresa, P.EndEmpresa_id as EndEmpresa, Empresa.Nome as NomeEmpresa, " & vbCrLf & _
                            "        Empresa.Reduzido as ReduzidoEmpresa, Empresa.Endereco as EnderecoEmpresa, Empresa.Cidade as CidadeEmpresa, Empresa.Estado as EstadoEmpresa, " & vbCrLf & _
                            "        Empresa.Inscricao as InscricaoEmpresa, Empresa.Numero as NumeroEmpresa, Empresa.Complemento as ComplementoEmpresa, Empresa.Bairro as BairroEmpresa " & vbCrLf & _
                            "   FROM Romaneios R" & vbCrLf & _
                            "  INNER JOIN RomaneiosXPesagens RxP" & vbCrLf & _
                            "     ON R.Empresa_Id = RxP.Empresa_Id " & vbCrLf & _
                            "    AND R.EndEmpresa_Id = RxP.EndEmpresa_Id " & vbCrLf & _
                            "    AND R.Romaneio_Id = RxP.Romaneio_Id " & vbCrLf & _
                            "  INNER JOIN Pesagem P" & vbCrLf & _
                            "     ON P.Empresa_Id = RxP.Empresa_Id" & vbCrLf & _
                            "    AND P.EndEmpresa_Id = RxP.EndEmpresa_Id " & vbCrLf & _
                            "    AND P.Pesagem_Id = RxP.Pesagem_Id " & vbCrLf & _
                            "    AND P.Sequencia_Id = RxP.Sequencia_Id " & vbCrLf & _
                            "  INNER JOIN Pedidos Ped " & vbCrLf & _
                            "     ON R.Empresa_Id = Ped.Empresa_Id " & vbCrLf & _
                            "    AND R.EndEmpresa_Id = Ped.EndEmpresa_Id " & vbCrLf & _
                            "    AND R.Pedido = Ped.Pedido_Id " & vbCrLf & _
                            "  INNER JOIN Clientes AS ClientePedido " & vbCrLf & _
                            "     ON Ped.Cliente = ClientePedido.Cliente_Id " & vbCrLf & _
                            "    AND Ped.EndCliente = ClientePedido.Endereco_Id " & vbCrLf & _
                            "   LEFT OUTER JOIN Produtos " & vbCrLf & _
                            "     ON P.Produto = Produtos.Produto_Id " & vbCrLf & _
                            "  INNER JOIN Placas " & vbCrLf & _
                            "     ON P.Placa = Placas.Placa_Id " & vbCrLf & _
                            "   LEFT OUTER JOIN Clientes " & vbCrLf & _
                            "     ON P.Cliente = Clientes.Cliente_Id " & vbCrLf & _
                            "    AND P.EndCliente = Clientes.Endereco_Id " & vbCrLf & _
                            "   LEFT OUTER JOIN Clientes AS Transportes " & vbCrLf & _
                            "     ON P.Transportador = Transportes.Cliente_Id " & vbCrLf & _
                            "    AND P.EndTransportador = Transportes.Endereco_Id " & vbCrLf & _
                            "   LEFT OUTER JOIN Estados AS EstadoPlaca " & vbCrLf & _
                            "     ON Placas.EstadoPlaca = EstadoPlaca.Estado_Id " & vbCrLf & _
                            "   LEFT OUTER JOIN Clientes AS Depositos " & vbCrLf & _
                            "     ON P.Deposito = Depositos.Cliente_Id " & vbCrLf & _
                            "    AND P.EndDeposito = Depositos.Endereco_Id " & vbCrLf & _
                            "   LEFT OUTER JOIN Clientes as Empresa " & vbCrLf & _
                            "     ON P.Empresa_Id = Empresa.Cliente_Id " & vbCrLf & _
                            "    AND P.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf & _
                            "   LEFT OUTER JOIN Clientes as Motorista " & vbCrLf & _
                            "     ON Placas.CpfMotorista = Motorista.Cliente_Id " & vbCrLf & _
                            "    AND Placas.EndCpfMotorista = Motorista.Endereco_Id" & vbCrLf & _
                            "   LEFT OUTER JOIN Estados AS EstadoPlacaMotorista " & vbCrLf & _
                            "     ON Motorista.Estado = EstadoPlacaMotorista.Estado_Id " & vbCrLf & _
                            "  WHERE P.Empresa_Id    = '" & txtCodigoEmpresa.Value.ToString.Split("-")(0) & "' " & vbCrLf & _
                            "    AND P.EndEmpresa_Id = " & txtCodigoEmpresa.Value.ToString.Split("-")(1) & vbCrLf & _
                            "    AND R.Romaneio_Id in(" & romaneio & ") " & vbCrLf & _
                            "    AND P.Sequencia_Id  = 0" & vbCrLf & _
                            "   /* DESCONTOS */" & vbCrLf & _
                            " SELECT RxD.Analise_Id AS Analise, " & vbCrLf & _
                            "        CASE " & vbCrLf & _
                            "             WHEN RxD.Analise_Id = 6 " & vbCrLf & _
                            "                  THEN An.Descricao + ' ' +" & vbCrLf & _
                            "                       (SELECT Observacoes" & vbCrLf & _
                            "                          FROM Classificacoes" & vbCrLf & _
                            "                         WHERE (Analise_Id = 6) " & vbCrLf & _
                            "                           AND (Sequencia_Id = RxD.Indice) " & vbCrLf & _
                            "                           AND (Produto_Id = '101010001')) " & vbCrLf & _
                            "                  ELSE An.Descricao " & vbCrLf & _
                            "         END AS Descricao, " & vbCrLf & _
                            "         RxD.Percentual, RxD.Indice, RxD.Desconto" & vbCrLf & _
                            "    FROM RomaneiosXDescontos RxD " & vbCrLf & _
                            "   INNER JOIN Analises An " & vbCrLf & _
                            "      ON RxD.Analise_Id = An.Analise_Id" & vbCrLf & _
                            "   WHERE RxD.Empresa_Id = '" & txtCodigoEmpresa.Value.ToString.Split("-")(0) & "'" & vbCrLf & _
                            "     AND RxD.EndEmpresa_Id = " & txtCodigoEmpresa.Value.ToString.Split("-")(1) & vbCrLf & _
                            "     AND RxD.Romaneio_Id in (" & romaneio & ")" & vbCrLf & _
                            "   ORDER BY Analise" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "Laudo")
    End Function

    Private Sub Imprimir(ByVal romaneio As String, Optional ByVal reemissao As Boolean = False)
        Dim ds As DataSet = getDataSet(romaneio)

        If ds IsNot Nothing AndAlso ds.Tables.Count > 1 Then
            ds.Tables(1).TableName = "Analises"
        End If

        Dim param As New Dictionary(Of String, Object)
        param.Add("Reemissao", IIf(reemissao, "Reemissão", ""))

        Funcoes.BindReport(Me.Page, ds, "Cr_LaudoDePesagem", eExportType.PDF, param)
    End Sub

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEmpresa.Click
        Try
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me, "objEmpresaPXR" & HID.Value.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub SetarEmpresa(ByVal Empresa As String, ByVal EndEmpresa As String)
        Dim objEmpresa As New [Lib].Negocio.Cliente(Empresa, EndEmpresa)
        Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objEmpresa)

        txtEmpresa.Text = itemEmpresa.Text
        txtCodigoEmpresa.Value = itemEmpresa.Value
    End Sub

    Private Function getDataSetRomaneio()
        Dim sql As String = "SELECT Empresa_Id, EndEmpresa_Id, Romaneio_Id, Pesagem_Id, Sequencia_Id " & vbCrLf & _
                            "  FROM RomaneiosXPesagens " & vbCrLf & _
                            " WHERE Pesagem_Id         = " & txtLaudo.Text & vbCrLf & _
                            "   AND Empresa_id    = " & txtCodigoEmpresa.Value.ToString.Split("-")(0) & vbCrLf & _
                            "   AND EndEmpresa_Id = " & txtCodigoEmpresa.Value.ToString.Split("-")(1) & vbCrLf

        Return Banco.ConsultaDataSet(sql, "Consulta")
    End Function

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If String.IsNullOrWhiteSpace(txtEmpresa.Text.Trim) Then
                MsgBox(Me.Page, "Informe a empresa.")
            ElseIf String.IsNullOrWhiteSpace(txtLaudo.Text.Trim) Then
                MsgBox(Me.Page, "Informe o Laudo.")
            Else

                Dim ds As DataSet = getDataSetRomaneio()

                If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                    If ds.Tables(0).Rows.Count = 1 Then
                        MsgBox(Me.Page, "Laudo não rateado.")
                    Else
                        For Each row As DataRow In ds.Tables(0).Rows
                            Imprimir(row("Romaneio_Id"), True)
                        Next
                        Limpar()
                    End If
                Else
                    MsgBox(Me.Page, "Nenhum registro encontrado.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnEmpresaAlt_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnEmpresaAlt.Click
        Try
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me, "objEmpresaPXRAlt" & HID.Value.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaSelecionados() As Boolean
        If String.IsNullOrWhiteSpace(txtEmpresaAlt.Text) Then
            MsgBox(Me.Page, "Informe a empresa para consulta.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtLaudoAlt.Text) Then
            MsgBox(Me.Page, "Informe o laudo para consulta.")
            Return False
        End If
        Return True
    End Function

    Protected Sub btnSelCliente_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSelCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteConsRXP" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LimparRomaneios()

        lnkConsultarAR.Parent.Visible = True
        lnkConsultarNF.Parent.Visible = True

        txtLaudoAlt.Enabled = True
        txtNotaAlt.Enabled = True

        btnBuscarNF.Enabled = False
        btnGravarNF.Visible = False

        txtLaudoAlt.Text = String.Empty
        txtNotaAlt.Text = String.Empty
        txtClienteAlt.Text = String.Empty
        txtMovimentoAlt.Text = String.Empty
        txtCodigoClienteAlt.Value = String.Empty
        txtPedidoAlt.Text = String.Empty
        txtCodigoAutorizacaoAlt.Value = String.Empty
        txtCodigoOperacaoAlt.Value = String.Empty
        gridRomaneios.DataSource = Nothing
        gridRomaneios.DataBind()
        txtClienteSel.Text = String.Empty
        txtCodigoClienteSel.Value = String.Empty
        txtPedidoSel.Text = String.Empty
        txtProdutoAlt.Value = String.Empty
        txtNotaFiscalAlt.Text = String.Empty

        txtCodigoAutorizacaoAltSel.Value = String.Empty
        txtPedidoSaldoSel.Value = String.Empty

        ddlDepositoDestinatarioSel.Items.Clear()
        ddlDepositoDestinatarioSel.Enabled = False

        ddlOperacaoSel.Items.Clear()
        ddlOperacaoSel.Enabled = False

        btnEmpresaAlt.Enabled = True
        btnSelCliente.Enabled = False
        btnAltRomaneio.Enabled = False

        LiberaEmpresa()

        'Dim dtRomaneios As New DataTable()
        'dtRomaneios.Columns.Add("Romaneio")
        'dtRomaneios.Columns.Add("Codigo")
        'dtRomaneios.Columns.Add("Pedido")
        'dtRomaneios.Columns.Add("Deposito")
        'dtRomaneios.Columns.Add("Produto")
        'dtRomaneios.Columns.Add("Op")
        'dtRomaneios.Columns.Add("SO")
        'dtRomaneios.Columns.Add("Autorizacao")
        'dtRomaneios.Columns.Add("Bruto")
        'dtRomaneios.Columns.Add("Desconto")
        'dtRomaneios.Columns.Add("Liquido")

        Session("NotaParaVinculoDeRateio" & HID.Value) = "N"

        'Session("dtRomaneiosAlt" & HID.Value) = dtRomaneios
    End Sub

    Protected Sub gridRomaneios_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles gridRomaneios.SelectedIndexChanged
        Try

            If Not String.IsNullOrWhiteSpace(txtNotaAlt.Text) AndAlso CInt(txtNotaAlt.Text) > 0 Then Exit Sub

            txtClienteSel.Text = ""
            txtCodigoClienteSel.Value = ""
            txtPedidoSel.Text = ""

            btnSelCliente.Enabled = True

            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteConsRXP" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub btnBuscarNF_Click(sender As Object, e As EventArgs) Handles btnBuscarNF.Click
        Try
            Empresa = txtCodigoEmpresaAlt.Value.ToString.Split("-")

            objNotaFiscal = New [Lib].Negocio.NotaFiscal()

            objNotaFiscal.CodigoEmpresa = Empresa(0)
            objNotaFiscal.CodigoSituacao = 1
            objNotaFiscal.CodigoTipoDeDocumento = 1
            objNotaFiscal.CodigoPedido = txtPedidoSel.Text
            objNotaFiscal.DataNota = Now.AddYears(-2)
            objNotaFiscal.Movimento = Now()

            Session("NotaParaVinculoDeRateio" & HID.Value) = "S"

            Session("objNotaFiscal" & HID.Value) = objNotaFiscal
            Session("ssCampo" & HID.Value) = "RateioDePesagem"

            ucConsultaPedidosXNotas.BindGridView()
            Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaNOTRAT" & HID.Value)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub btnGravarNF_Click(sender As Object, e As EventArgs) Handles btnGravarNF.Click
        Try
            If ddlOperacaoSel.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Operação não foi selecionada.")
            ElseIf ddlDepositoDestinatarioSel.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Depósito não foi selecionado.")
            ElseIf String.IsNullOrWhiteSpace(txtNotaFiscalAlt.Text) Then
                MsgBox(Me.Page, "Nota Fiscal não foi informada.")
            ElseIf CInt(txtNotaFiscalAlt.Text) = 0 Then
                MsgBox(Me.Page, "Nota Fiscal não foi informada.")
            ElseIf String.IsNullOrWhiteSpace(gridRomaneios.Rows(gridRomaneios.SelectedIndex).Cells(1).Text) Then
                MsgBox(Me.Page, "Romaneio não foi selecionado.")
            ElseIf CInt(gridRomaneios.Rows(gridRomaneios.SelectedIndex).Cells(1).Text) = 0 Then
                MsgBox(Me.Page, "Romaneio não foi selecionado.")
            Else

                RecuperaNotaFiscal()

                Dim pRomaneio As Integer = gridRomaneios.Rows(gridRomaneios.SelectedIndex).Cells(1).Text

                Dim sql As String
                Dim Sqls As New ArrayList

                Dim obs As String = objNotaFiscal.ObservacoesControleInterno
                If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                    obs = obs & ". Vinculado Romaneio pelo Rateio de Pesagem em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " - Usuário " & HttpContext.Current.Session("ssNomeUsuario")
                Else
                    obs = "Vinculado Romaneio pelo Rateio de Pesagem em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " - Usuário " & HttpContext.Current.Session("ssNomeUsuario")
                End If

                sql = " Update NotasFiscais set " & vbCrLf & _
                      "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                      "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf & _
                      "      ObservacoesControleInterno  = '" & obs & "'" & vbCrLf & _
                      "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                      "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                      "	   and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                      "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                      "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                      "    and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                      "    and Nota_Id         = " & objNotaFiscal.Codigo & ";"
                Sqls.Add(sql)

                sql = "UPDATE NotasFiscaisXItens SET" & vbCrLf & _
                             "    QuantidadeFisica = " & CType(Session("dtRomaneiosAlt" & HID.Value), DataTable).Rows(gridRomaneios.SelectedIndex).Item("Liquido") & vbCrLf & _
                             " WHERE Empresa_id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                             "   AND EndEmpresa_id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                             "   AND Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                             "   AND EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                             "   AND EntradaSaida_Id ='" & Left(objNotaFiscal.EntradaSaida.ToString, 1) & "'" & vbCrLf & _
                             "   AND Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                             "   AND Nota_Id         = " & objNotaFiscal.Codigo & ";"
                Sqls.Add(sql)


                sql = " Insert into NotasFiscaisXRomaneios(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Romaneio_Id)" & vbCrLf & _
                      " Values('" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                      "," & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                      ",'" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                      "," & objNotaFiscal.EnderecoCliente & vbCrLf & _
                      ",'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                      ",'" & objNotaFiscal.Serie & "'" & vbCrLf & _
                      "," & objNotaFiscal.Codigo & vbCrLf & _
                      "," & pRomaneio & ");"
                Sqls.Add(sql)


                sql = "Update Romaneios set EntradaSaida =  '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                      "  ,Pedido       = " & objNotaFiscal.CodigoPedido & vbCrLf & _
                      "  ,Deposito     ='" & objNotaFiscal.CodigoDeposito & "'" & vbCrLf & _
                      "  ,EndDeposito  = " & objNotaFiscal.EnderecoDeposito & vbCrLf & _
                      "  ,Destino      ='" & objNotaFiscal.CodigoDestino & "'" & vbCrLf & _
                      "  ,EndDestino   = " & objNotaFiscal.EnderecoDestino & vbCrLf & _
                      "  ,Operacao     = " & objNotaFiscal.CodigoOperacao & vbCrLf & _
                      "  ,SubOperacao  = " & objNotaFiscal.CodigoSubOperacao & vbCrLf & _
                      "  ,Autorizacao  = " & objNotaFiscal.CodigoAutorizacao & vbCrLf & _
                      " Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                      "   and Romaneio_Id   = " & pRomaneio & vbCrLf
                Sqls.Add(sql)

                If Banco.GravaBanco(Sqls) Then
                    For Each row As DataRow In CType(Session("dtRomaneiosAlt" & HID.Value), DataTable).Rows
                        If row("Romaneio") = pRomaneio Then
                            row("Pedido") = txtPedidoSel.Text
                            row("OP") = objNotaFiscal.CodigoOperacao
                            row("SO") = objNotaFiscal.CodigoSubOperacao
                            row("Autorizacao") = objNotaFiscal.CodigoAutorizacao
                        End If
                    Next

                    gridRomaneios.DataSource = CType(Session("dtRomaneiosAlt" & HID.Value), DataTable)
                    gridRomaneios.DataBind()

                    MsgBox(Me.Page, "Rateio Vinculado com Sucesso na Nota Fiscal.", eTitulo.Sucess)

                    Limpar()
                    LimparRomaneios()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub



    Protected Sub btnAltRomaneio_Click(sender As Object, e As EventArgs) Handles btnAltRomaneio.Click
        Try
            If ddlOperacaoSel.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Operação não foi selecionada.")
                Exit Sub
            ElseIf ddlDepositoDestinatarioSel.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Depósito não foi selecionado.")
                Exit Sub
            End If

            Dim sql As String
            Dim Sqls As New ArrayList

            Empresa = txtCodigoEmpresaAlt.Value.ToString.Split("-")
            Dim pPedido = New [Lib].Negocio.Pedido(Empresa(0), Empresa(1), txtPedidoSel.Text)
            Dim subOperacao() As String = ddlOperacaoSel.SelectedValue.ToString.Split("-")
            Deposito = ddlDepositoDestinatarioSel.SelectedValue.ToString.Split("-")
            Dim pDestino As String = ""
            Dim pEndDestino As Integer = 0
            Dim pRomaneio As Integer = gridRomaneios.Rows(gridRomaneios.SelectedIndex).Cells(1).Text

            For Each pXd As [Lib].Negocio.PedidoXDeposito In pPedido.Depositos
                If pXd.Tipo = "OD" AndAlso pDestino.Length = 0 Then
                    pDestino = pXd.Codigo
                    pEndDestino = pXd.CodigoEndereco
                End If
            Next

            sql = "Update Romaneios set EntradaSaida =  '" & subOperacao(2) & "'" & vbCrLf & _
                  "  ,Pedido       = " & txtPedidoSel.Text & vbCrLf & _
                  "  ,Deposito     ='" & Deposito(0) & "'" & vbCrLf & _
                  "  ,EndDeposito  = " & Deposito(1) & vbCrLf & _
                  "  ,Destino      ='" & pDestino & "'" & vbCrLf & _
                  "  ,EndDestino   = " & pEndDestino & vbCrLf & _
                  "  ,Operacao     = " & subOperacao(0) & vbCrLf & _
                  "  ,SubOperacao  = " & subOperacao(1) & vbCrLf & _
                  "  ,Autorizacao  = " & txtCodigoAutorizacaoAlt.Value & vbCrLf & _
                  " Where Empresa_Id    ='" & Empresa(0) & "'" & vbCrLf & _
                  "   and EndEmpresa_Id = " & Empresa(1) & vbCrLf & _
                  "   and Romaneio_Id   = " & pRomaneio & vbCrLf
            Sqls.Add(sql)

            If Banco.GravaBanco(Sqls) Then
                For Each row As DataRow In CType(Session("dtRomaneiosAlt" & HID.Value), DataTable).Rows
                    If row("Romaneio") = pRomaneio Then
                        row("Pedido") = txtPedidoSel.Text
                        row("OP") = subOperacao(0)
                        row("SO") = subOperacao(1)
                        row("Autorizacao") = txtCodigoAutorizacaoAlt.Value
                    End If
                Next

                gridRomaneios.DataSource = CType(Session("dtRomaneiosAlt" & HID.Value), DataTable)
                gridRomaneios.DataBind()

                MsgBox(Me.Page, "Rateio realizado com Sucesso.", eTitulo.Sucess)

                Limpar()
                LimparRomaneios()
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlOperacaoSel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlOperacaoSel.SelectedIndexChanged
        Try
            If ddlOperacaoSel.SelectedIndex > 0 Then
                btnAltRomaneio.Enabled = True
                btnBuscarNF.Enabled = True
            Else
                btnAltRomaneio.Enabled = False
                btnBuscarNF.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'RATEIO
    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("RateioDePesagem", "ACESSAR") Then
                If txtCodigoEmpresa.Value.ToString.Length = 0 Then
                    MsgBox(Me.Page, "Empresa não foi informada.")
                ElseIf txtLaudo.Text.Length = 0 Then
                    MsgBox(Me.Page, "Número da pesagem não foi informado.")
                ElseIf Funcoes.VerificaPermissao("RateioDePesagem", "LEITURA") Then
                    Consultar_Laudo()
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("RateioDePesagem", "EXCLUIR") Then
                ExcluirRateio()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
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

    'ALTERAR RATEIO
    Protected Sub lnkConsultarAR_Click(sender As Object, e As EventArgs) Handles lnkConsultarAR.Click
        Try
            If Funcoes.VerificaPermissao("RateioDePesagem", "ACESSAR") Then
                If ValidaSelecionados() Then
                    Empresa = txtCodigoEmpresaAlt.Value.ToString.Split("-")

                    Try
                        Dim objPesagem As New [Lib].Negocio.Pesagem(Empresa(0), Empresa(1), txtLaudoAlt.Text)

                        If objPesagem.Codigo = 0 Then
                            MsgBox(Me.Page, "Laudo de pesagem não encontrado.")
                            txtLaudo.Focus()
                            Exit Sub
                        ElseIf objPesagem.TemNota AndAlso Not objPesagem.TemRomaneioSemNota Then
                            MsgBox(Me.Page, "Romaneios do laudo de pesagem ja possui nota fiscal.")
                            txtLaudo.Focus()
                            Exit Sub
                        End If

                        lnkConsultarAR.Parent.Visible = False
                        lnkConsultarNF.Parent.Visible = False
                        txtNotaAlt.Text = String.Empty
                        txtNotaAlt.Enabled = False
                        txtLaudoAlt.Enabled = False

                        btnEmpresaAlt.Enabled = False
                        txtCodigoEmpresaAlt.Value = objPesagem.CodigoEmpresa & "-" & objPesagem.EnderecoEmpresa
                        txtMovimentoAlt.Text = objPesagem.Movimento.ToString("dd/MM/yyyy")
                        txtCodigoClienteAlt.Value = objPesagem.CodigoCliente & "-" & objPesagem.EnderecoCliente
                        Dim itemClienteAlt As ListItem = Funcoes.FormatarListItemCliente(objPesagem.Cliente)
                        txtClienteAlt.Text = itemClienteAlt.Text
                        txtPedidoAlt.Text = objPesagem.CodigoPedido
                        txtCodigoOperacaoAlt.Value = objPesagem.CodigoOperacao & "-" & objPesagem.CodigoSubOperacao & "-" & objPesagem.EntradaSaida & "-" & objPesagem.SubOperacao.Devolucao & "-" & objPesagem.SubOperacao.Classe
                        txtProdutoAlt.Value = objPesagem.CodigoProduto

                        Dim listRomXPesagem As ListRomaneioXPesagem = objPesagem.Romaneios(0).Pesagens()

                        If listRomXPesagem IsNot Nothing AndAlso listRomXPesagem.Count > 0 Then
                            For Each row_RXP As [Lib].Negocio.RomaneioXPesagem In listRomXPesagem
                                Dim rom As New Romaneio(objPesagem.CodigoEmpresa, objPesagem.EnderecoEmpresa, row_RXP.CodigoRomaneio)
                                If rom.Codigo > 0 Then
                                    txtCodigoAutorizacaoAlt.Value = rom.CodigoAutorizacao
                                End If
                            Next
                        End If

                        Dim dtRomaneios As New DataTable()
                        dtRomaneios.Columns.Add("Romaneio")
                        dtRomaneios.Columns.Add("Codigo")
                        dtRomaneios.Columns.Add("Pedido")
                        dtRomaneios.Columns.Add("Deposito")
                        dtRomaneios.Columns.Add("Produto")
                        dtRomaneios.Columns.Add("Op")
                        dtRomaneios.Columns.Add("SO")
                        dtRomaneios.Columns.Add("Autorizacao")
                        dtRomaneios.Columns.Add("Bruto")
                        dtRomaneios.Columns.Add("Desconto")
                        dtRomaneios.Columns.Add("Liquido")

                        Session("dtRomaneiosAlt" & HID.Value) = dtRomaneios

                        For Each row As [Lib].Negocio.Romaneio In objPesagem.Romaneios
                            If Not row.TemNotaFiscal Then
                                Dim newRow As DataRow = CType(Session("dtRomaneiosAlt" & HID.Value), DataTable).NewRow
                                newRow.Item("Romaneio") = row.Codigo
                                newRow.Item("Codigo") = Funcoes.FormatarCpfCnpj(row.Pedido.CodigoCliente) & "-" & row.Pedido.EnderecoCliente & " " & row.Pedido.Cliente.Nome
                                newRow.Item("Pedido") = row.CodigoPedido
                                newRow.Item("Deposito") = Funcoes.FormatarCpfCnpj(row.CodigoDeposito) & "-" & row.EnderecoDeposito & " " & row.Deposito.Nome
                                newRow.Item("Produto") = row.CodigoProduto
                                newRow.Item("OP") = row.CodigoOperacao
                                newRow.Item("SO") = row.CodigoSubOperacao
                                newRow.Item("Autorizacao") = row.CodigoAutorizacao
                                newRow.Item("Bruto") = row.PesoBruto
                                newRow.Item("Desconto") = row.Desconto
                                newRow.Item("Liquido") = row.PesoLiquido
                                CType(Session("dtRomaneiosAlt" & HID.Value), DataTable).Rows.Add(newRow)
                            End If
                        Next

                        gridRomaneios.DataSource = CType(Session("dtRomaneiosAlt" & HID.Value), DataTable)
                        gridRomaneios.DataBind()

                        CType(gridRomaneios.Rows(0).FindControl("imgRemoverRomaneio"), ImageButton).Enabled = False

                    Catch ex As Exception
                        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                    End Try
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultarNF_Click(sender As Object, e As EventArgs) Handles lnkConsultarNF.Click
        Try
            If Funcoes.VerificaPermissao("RateioDePesagem", "ACESSAR") Then

                If String.IsNullOrWhiteSpace(txtEmpresaAlt.Text) Then
                    MsgBox(Me.Page, "Informe a empresa para consulta.")
                ElseIf String.IsNullOrWhiteSpace(txtNotaAlt.Text) Then
                    MsgBox(Me.Page, "Informe a nota fiscal para consulta.")
                ElseIf CInt(txtNotaAlt.Text) = 0 Then
                    MsgBox(Me.Page, "Informe a nota fiscal para consulta.")
                Else
                    Empresa = txtCodigoEmpresaAlt.Value.ToString.Split("-")

                    objNotaFiscal = New [Lib].Negocio.NotaFiscal()

                    objNotaFiscal.CodigoEmpresa = Empresa(0)
                    objNotaFiscal.EnderecoEmpresa = Empresa(1)
                    objNotaFiscal.Codigo = txtNotaAlt.Text
                    objNotaFiscal.CodigoSituacao = 1
                    objNotaFiscal.CodigoTipoDeDocumento = 1

                    Session("objNotaFiscal" & HID.Value) = objNotaFiscal
                    Session("ssCampo" & HID.Value) = "RateioDePesagem"

                    ucConsultaPedidosXNotas.BindGridView()
                    Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaRAT" & HID.Value)

                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparAR_Click(sender As Object, e As EventArgs) Handles lnkLimparAR.Click
        Try
            Limpar()
            LimparRomaneios()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RateioDePesagem")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConfirmar_Click(sender As Object, e As EventArgs) Handles lnkConfirmar.Click
        Try
            If String.IsNullOrWhiteSpace(txtQuantidade.Text) OrElse CDec(txtQuantidade.Text) = 0 Then
                MsgBox(Me.Page, "Informe uma quantidade para fazer o rateio.")
                txtQuantidade.Focus()
            ElseIf ddlOperacao.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Operação não foi selecionada.")
            ElseIf ddlDepositoDestinatario.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Depósito não foi selecionado.")
            Else
                Dim Operacao() As String = ddlOperacao.SelectedValue.ToString.Split("-")
                Dim sOP = New [Lib].Negocio.SubOperacao(Operacao(0), Operacao(1))
                Empresa = txtCodigoEmpresa.Value.ToString.Split("-")
                txtCodigoAutorizacaoDestinatario.Value = 0

                Dim objLaudo = New [Lib].Negocio.Pesagem()
                objLaudo.CodigoEmpresa = Empresa(0)
                objLaudo.EnderecoEmpresa = Empresa(1)
                objLaudo.CodigoPedido = txtPedidoDestinatario.Text

                Dim dsSaldoPedido As DataSet = objLaudo.SaldoDePedidos(objLaudo, False)

                txtPedidoEntregue.Value = dsSaldoPedido.Tables(0).Rows(0).Item("Entregue")
                If sOP.Devolucao Then
                    txtPedidoSaldo.Value = dsSaldoPedido.Tables(0).Rows(0).Item("SaldoDevolucao")
                Else
                    txtPedidoSaldo.Value = dsSaldoPedido.Tables(0).Rows(0).Item("Saldo")
                End If

                If sOP.EntradaSaida = eEntradaSaida.Saida Then
                    Dim parameters As New Dictionary(Of String, Object)
                    parameters.Add("ped", txtPedidoDestinatario.Text)
                    parameters.Add("emp", Empresa(0))
                    parameters.Add("endemp", Empresa(1))
                    parameters.Add("cli", "")
                    parameters.Add("endcli", "")
                    parameters.Add("romaneio", True)
                    parameters.Add("classe", sOP.Classe)
                    ucConsultaAutorizacaoDeRetirada.BindGridView(parameters)
                    Popup.ConsultaDeAutorizacaoDeRetirada(Me.Page, "objRateioDePesagem" & HID.Value)
                Else
                    If VerificaSaldo(Empresa, txtPedidoDestinatario.Text) Then
                        GravaParcial()
                        lnkFinalizar.Parent.Visible = True
                        lnkExcluirRateio.Parent.Visible = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgRemoverRomaneio_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            RecuperaNotaFiscal()

            Dim sqls As New ArrayList

            Sql = "DELETE FROM notasfiscaisXromaneios" & vbCrLf & _
               "where Empresa_Id  = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
               "and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
               "and Cliente_Id    = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
               "and EndCliente_Id = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
               "and Romaneio_Id   = " & objNotaFiscal.CodigoRomaneio & ";"
            sqls.Add(Sql)

            Sql = "UPDATE NotasFiscaisXItens SET" & vbCrLf &
                         "    QuantidadeFisica   = 0" & vbCrLf &
                         " WHERE Empresa_id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                         "   AND EndEmpresa_id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                         "   AND Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                         "   AND EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                         "   AND EntradaSaida_Id ='" & Left(objNotaFiscal.EntradaSaida.ToString, 1) & "'" & vbCrLf &
                         "   AND Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                         "   AND Nota_Id         = " & objNotaFiscal.Codigo & ";"
            sqls.Add(Sql)

            Dim obs As String = objNotaFiscal.ObservacoesControleInterno
            If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                obs = obs & ". Desvinculado Romaneio pelo Rateio de Pesagem em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " - Usuário " & HttpContext.Current.Session("ssNomeUsuario")
            Else
                obs = "Desvinculado Romaneio pelo Rateio de Pesagem em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " - Usuário " & HttpContext.Current.Session("ssNomeUsuario")
            End If

            Sql = " Update NotasFiscais set " & vbCrLf & _
                  "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                  "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf & _
                  "      ObservacoesControleInterno  = '" & obs & "'" & vbCrLf & _
                  "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                  "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                  "	   and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                  "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                  "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                  "    and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                  "    and Nota_Id         = " & objNotaFiscal.Codigo & ";"
            sqls.Add(Sql)

            Sql = " Delete From Romaneiosxdescontos" & vbCrLf & _
                  " Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                  "   and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                  "   and Romaneio_Id   = " & objNotaFiscal.Romaneio.Codigo & ";"
            sqls.Add(Sql)

            Sql = " Delete From Romaneios " & vbCrLf & _
                  "  Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                  "    and EndEmpresa_id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                  "    and Romaneio_Id   = " & objNotaFiscal.Romaneio.Codigo & ";"
            sqls.Add(Sql)

            If Banco.GravaBanco(sqls) Then
                MsgBox(Me.Page, "Romaneio da Nota Fiscal desvinculado com Sucesso.")
                Limpar()
                LimparRomaneios()
            Else
                MsgBox(Me.Page, "Erro ao Desvincular o Romaneio da Nota Fiscal.")
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnExcluirItemRateio_Click(sender As Object, e As EventArgs)
        Try
            If GridRateio.Rows.Count = 1 Then
                MsgBox(Me.Page, "Lista não pode ser excluída.")
            Else
                Dim btn As ImageButton = CType(sender, ImageButton)
                Dim rowbtn As GridViewRow = CType(btn.NamingContainer, GridViewRow)

                VerificaSaldoParaExcluir(rowbtn.Cells(1).Text(), rowbtn.Cells(4).Text())

                Empresa = txtCodigoEmpresa.Value.Split("-")

                Sql = "  DELETE AuxiliarXRateio " & vbCrLf & _
                      "   WHERE Laudo_Id = 0 And Cliente_Id = '" & rowbtn.Cells(1).Text() & "'" & vbCrLf & _
                      "     And Pedido_Id = " & rowbtn.Cells(4).Text() & vbCrLf
                SqlArray.Add(Sql)

                Sql = "  Update AuxiliarXRateio" & vbCrLf & _
                      "  Set Bruto = Bruto + " & Bruto & vbCrLf & _
                      ", Desconto = Desconto +" & Desconto & vbCrLf & _
                      ", Liquido = Liquido + " & Liquido & vbCrLf & _
                      "  WHERE Empresa_Id = '" & Empresa(0) & "' AND EndEmpresa_Id = " & Empresa(1) & "And Laudo_Id = " & txtLaudo.Text & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) Then
                    Sql = " SELECT AxR.Cliente_Id AS Cliente, AxR.EndCliente_Id AS Endereco, Cli.Nome, AxR.Pedido_Id AS Pedido, AxR.Operacao, " & vbCrLf & _
                          "        AxR.SubOperacao, AxR.Bruto, AxR.Desconto, AxR.Liquido, AxR.Produto, AxR.Autorizacao " & vbCrLf & _
                          "   FROM AuxiliarXRateio AxR " & vbCrLf & _
                          "   LEFT OUTER JOIN Clientes Cli " & vbCrLf & _
                          "     ON AxR.Cliente_Id    = Cli.Cliente_Id " & vbCrLf & _
                          "    AND AxR.EndCliente_Id = Cli.Endereco_Id " & vbCrLf & _
                          "  WHERE Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                          "    AND EndEmpresa_Id = " & Empresa(1) & vbCrLf

                    Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

                    CType(Session("dtRomaneios" & HID.Value), DataTable).Clear()

                    For Each row As DataRow In ds.Tables(0).Rows

                        Dim newRow As DataRow = CType(Session("dtRomaneios" & HID.Value), DataTable).NewRow

                        newRow.Item("Codigo") = row("Cliente")
                        newRow.Item("Nome") = row("Nome")
                        newRow.Item("End") = row("Endereco")
                        newRow.Item("Pedido") = row("Pedido")
                        newRow.Item("Produto") = row("Produto")
                        newRow.Item("OP") = row("Operacao")
                        newRow.Item("SO") = row("SubOperacao")
                        newRow.Item("Autorizacao") = row("Autorizacao")
                        newRow.Item("Bruto") = row("Bruto")
                        newRow.Item("Desconto") = row("Desconto")
                        newRow.Item("Liquido") = row("Liquido")
                        CType(Session("dtRomaneios" & HID.Value), DataTable).Rows.Add(newRow)
                    Next

                    GridRateio.DataSource = CType(Session("dtRomaneios" & HID.Value), DataTable)
                    GridRateio.DataBind()

                    txtDestinatario.Text = ""
                    txtQuantidade.Text = ""
                End If
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkFinalizar_Click(sender As Object, e As EventArgs) Handles lnkFinalizar.Click
        Try
            If Funcoes.VerificaPermissao("RateioDePesagem", "GRAVAR") Then
                If GridRateio.Rows.Count > 0 Then
                    If GravarRateio() Then
                        Dim ds As DataSet = getDataSetRomaneio()
                        For Each row As DataRow In ds.Tables(0).Rows
                            Imprimir(row("Romaneio_Id"), True)
                        Next
                        Limpar()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar rateio.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExtratoPedido_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgExtratoPedido.Click
        If txtPedido.Text.Length = 0 Then
            MsgBox(Me.Page, "Pedido deve ser selecionado")
        Else
            Dim empresa() As String = txtCodigoEmpresa.Value.Split("-")
            Dim cliente() As String = txtCodigoCliente.Value.ToString.Split("-")
            Extrato.Emitir(Me.Page, FinanceiroNovo, empresa(0), empresa(1), "T", txtPedido.Text)
        End If
    End Sub

    Protected Sub imgExtratoPedidoDestinatario_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgExtratoPedidoDestinatario.Click
        If txtPedidoDestinatario.Text.Length = 0 Then
            MsgBox(Me.Page, "Pedido do Destinatário deve ser selecionado")
        Else
            Dim empresa() As String = txtCodigoEmpresa.Value.Split("-")
            Extrato.Emitir(Me.Page, FinanceiroNovo, empresa(0), empresa(1), "T", txtPedido.Text)
        End If
    End Sub
End Class