Imports System.IO
Imports System.Data
Imports System.Net.Configuration
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Threading
Imports System.Xml
Imports System.Globalization

Public Class NotaFiscalXItens
    Inherits BasePage

#Region "Variáveis"

    Private Sql As String
    Private erroMsg As String = String.Empty
    Private MsgAlerta As String = String.Empty
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal
    Private objPlaca As [Lib].Negocio.Placa
    Private objListaDeCombinacoesLoteEmbalagem As [Lib].Negocio.ListSaldoProdutoEstoqueLoteEmbalagem
    Private objListaDeNotasFiscaisReferenciais As [Lib].Negocio.ListNotaFiscalReferencial
    Private objListaDeNotasFiscaisReferenciaisAnteriores As [Lib].Negocio.ListNotaFiscalReferencial
    Private chaveXMLautomatico As String
    Private i As Integer
    Dim logs As FuncoesLogs
    Dim bloco1 As String
    Dim bloco2 As String
    Dim bloco3 As String

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If IsConnect AndAlso Not IsPostBack Then

                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Expedicao.aspx", eTitulo.Info)
                    Exit Sub
                End If

                If Funcoes.VerificaPermissao("NotaFiscalXItens", "ACESSAR") Then
                    TabVencimentosOld.Visible = Not FinanceiroNovo
                    TabVencimentos.Visible = FinanceiroNovo
                    txtDataDeEmissao.Text = Today.ToString("dd/MM/yyyy")
                    txtDataDeEntrada.Text = Today.ToString("dd/MM/yyyy")
                    txtHoraDaSaida.Text = Now.ToString("HH:mm:ss")
                    ddl.Carregar(ddlTipoDeDocumento, CarregarDDL.Tabela.TipoDeDocumento, "", False)
                    ddl.Carregar(ddlUF, CarregarDDL.Tabela.EstadosUF, " Estado_Id not in('DR','SX','EX')", True)
                    ddl.Carregar(ddlPaisDestino, CarregarDDL.Tabela.Pais, "", True)
                    ddlTipoConhec.SelectedValue = 10
                    ddl.Carregar(cmbFinalidade, CarregarDDL.Tabela.Finalidade, "")
                    primeiraVez.Value = True

                    If Not Session("chaveXMLautomacao" & HID.Value) Is Nothing Then
                        chaveXMLautomatico = Session("chaveXMLautomacao" & HID.Value).ToString()
                        Session.Remove("chaveXMLautomacao" & HID.Value)
                    End If

                    LimparCampos("", "", False, Session("ssEmpresa"), Session("ssEndEmpresa"))
                    txtReajuste.Visible = False
                    btnReajusta.Visible = False

                    If Not Left(Session("ssEmpresa"), 8) = "24450490" Then chkEspelho.Checked = True

                    ddl.Carregar(ddlViaDeTransporteDI, CarregarDDL.Tabela.ViaDeTransportesSefaz, "", True)

                    If Not LancarSaldoInicial Then carregarNotaXMLautomatico(sender, e)

                    If Not Directory.Exists("C:\NGS\Log\NotasFiscais") Then Directory.CreateDirectory("C:\NGS\Log\NotasFiscais")

                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
            'SessaoVerificaCarregaObjetos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnTransferencia_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnTransferencia.Click
        Try
            If Funcoes.VerificaPermissao("TRANSFERENCIA", "ACESSAR") Then
                SessaoRecuperaNotaFiscal()
                ucTransferencias.InicializarUC(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                Popup.ConsultaTransferencias(Me.Page, "objTransferencias" & HID.Value)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar Transferência!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnRecontabilizar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnRecontabilizar.Click
        Try
            If Funcoes.VerificaPermissao("RECONTABILIZAR", "ALTERAR") Then

                SessaoRecuperaNotaFiscal()

                If objNotaFiscal.Itens.Count > 0 Then
                    Dim sqls As New ArrayList
                    objNotaFiscal.Razao.ExcluiContabilizacaoNotaSQL(sqls)
                    objNotaFiscal.Razao.ContabilizarNotaSql(sqls)

                    If Banco.GravaBanco(sqls) Then
                        Dim objNotaAntes = New NotaFiscal(objNotaFiscal)
                        objNotaFiscal = New NotaFiscal(objNotaAntes)
                        objNotaFiscal.NotaFiscalOriginal = New NotaFiscal(objNotaFiscal)
                        SessaoSalvaNotaFiscal()

                        AtualizaFormularioComAClasse()

                        MsgBox(Me.Page, "Nota Recontabilizada com Sucesso.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, "Erro Durante a Recontabilizacão da nota")
                    End If
                Else
                    MsgBox(Me.Page, "Consute a Nota Fiscal para recontabilização")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para recontabilizar Nota Fiscal")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub imgExtratoPedido_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.CodigoCliente.Length = 0 Then
                MsgBox(Me.Page, "Cliente não foi selecionado")
                Exit Sub
            End If

            If objNotaFiscal.Pedido IsNot Nothing Then Extrato.Emitir(Me.Page, FinanceiroNovo, objNotaFiscal.Pedido.CodigoEmpresa, objNotaFiscal.Pedido.EnderecoEmpresa, "T",
                                                                      objNotaFiscal.Pedido.Codigo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgRomaneio_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgRomaneio.Click
        Try
            SessaoRecuperaNotaFiscal()

            Funcoes.BindRomaneio(Me.Page, objNotaFiscal.CodigoEmpresa, objNotaFiscal.CodigoRomaneio, objNotaFiscal.Itens(0).CodigoProduto)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnEspelho_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ImprimirEspelho()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub txtObservacoesInternas_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNotaFiscal()
            objNotaFiscal.ObservacoesControleInterno = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacoesInternas.Text)
            SessaoSalvaNotaFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub btnEncargos_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNotaFiscal()

            If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.IUD = "U" Then BtnRecontabilizar.Enabled = False 'Desabilitado caso entre aqui porque estava contabilizando a nota no ajustar encargos em nota sem ALTERAÇÃO/NOSSA EMISSÃO.

            Dim btn As Button = CType(sender, Button)
            Dim row As GridViewRow = CType(btn.NamingContainer, GridViewRow)
            ucNFEncargo.Limpar()
            ucNFEncargo.Inicializar(row.RowIndex)
            Popup.NFEncargo(Me.Page, "objNFEncargo" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Observações do Produto
    Protected Sub imgObsProduto_Click(sender As Object, e As ImageClickEventArgs)
        Try
            SessaoRecuperaNotaFiscal()
            If Not objNotaFiscal Is Nothing Then
                Dim Imgproduto As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(Imgproduto.NamingContainer, GridViewRow)
                ucNFObsProduto.Limpar()
                ucNFObsProduto.CarregarObs(row.RowIndex)
                Popup.NFObsProduto(Me.Page, "ObjNFObsProduto" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub DesabilitaLinhaProduto(ByVal idLinha As Integer)
        Try
            Dim i As Integer = 0
            While i < gridItens.Rows.Count

                If i = idLinha Then
                    CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled = False
                    CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled = False
                End If

                i += 1
            End While
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgSelecionar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            SessaoRecuperaNotaFiscal()

            Dim chkProduto As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(chkProduto.NamingContainer, GridViewRow)

            If objNotaFiscal.SubOperacao.Devolucao Then
                If objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                    If CDec(CType(gridItens.Rows(row.RowIndex).FindControl("txtQuantidadeItem"), TextBox).Text) <> objNotaFiscal.Itens(row.RowIndex).QuantidadeFiscal Then
                        ItemNotaOK(row.RowIndex)
                        Exit Sub
                    End If
                Else
                    If CDec(CType(gridItens.Rows(row.RowIndex).FindControl("txtTotalItem"), TextBox).Text) <> objNotaFiscal.Itens(row.RowIndex).ValorTotal Then
                        ItemNotaOK(row.RowIndex)
                        Exit Sub
                    End If
                End If

                Session("ssCampo" & HID.Value) = "LivreClasse"
                ucNotaDeDevolucaoXNota.SetarIndice(row.RowIndex)
                ucNotaDeDevolucaoXNota.BindGridView()
                Popup.ConsultaDeNotaDeDevolucaoXNota(Me.Page, "objNotaDeDevolucaoXNota" & HID.Value)
            ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES And objNotaFiscal.SubOperacao.Memorando And objNotaFiscal.EntradaSaida = eEntradaSaida.Saida Then
                If Not objNotaFiscal.NotasReferenciais Is Nothing AndAlso objNotaFiscal.NotasReferenciais.Count > 0 Then
                    MsgBox(Me.Page, "Já existem notas referenciais. Para prosseguir, estas devem ser excluídas!")
                    TabContainer1.ActiveTab = TabExportacao
                Else
                    Dim Empresa() As String = {objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa}

                    Dim Parametros As New Dictionary(Of String, Object)()
                    Parametros.Add("TipoDeDocumento", eTipoDeDocumento.Nota)
                    Parametros.Add("EntradaSaida", eEntradaSaida.Entrada)
                    Parametros.Add("Empresa", Empresa)
                    Parametros.Add("Produto", objNotaFiscal.Itens(row.RowIndex).Produto.Codigo)
                    Parametros.Add("QuantidadeFiscalNF", objNotaFiscal.TotalQuantidadeFiscal)
                    Parametros.Add("TipoReferencial", eTipoReferencial.EXP)

                    ucNotaFiscalReferencial.ConsultarNotas(Parametros)

                    Popup.ConsultaNotaFiscalReferencial(Me.Page, "objNotaFiscalReferencial" & HID.Value)
                End If
            Else
                MsgBox(Me.Page, "Esta opção somente pode ser usada em notas de devolução ou de Exportação!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub BtnVencimentos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnVencimentos.Click
        Try
            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.Codigo = 0 Then
                MsgBox(Me.Page, "Número da Nota não foi informado")
            ElseIf objNotaFiscal.Serie.Length = 0 Then
                MsgBox(Me.Page, "Série da Nota não foi informada")
            ElseIf Not objNotaFiscal.SubOperacao.Devolucao Then
                If objNotaFiscal.Pedido.MomentoFinanceiro <> 3 Or objNotaFiscal.IUD = "U" Or Not objNotaFiscal.SubOperacao.Financeiro Then
                    If objNotaFiscal.Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR AndAlso objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.COMPLEMENTACOES Then
                        'LIBERA FINANCEIRO PARA COMPLEMENTAÇÃO DO AFIXAR - FURLAN - 18/09/2024
                    Else
                        MsgBox(Me.Page, "Momento Financeiro pelo Pedido não pode ser alterado na Nota Fiscal")
                        Exit Sub
                    End If
                End If

                If objNotaFiscal.Cliente.ApenasAVista AndAlso Not objNotaFiscal.Pedido.CondicaoPagamento.AVista Then
                    MsgBox(Me.Page, "Para Cliente com Condição de Pagamento apenas a Vista, o Pedido deve sestar na mesma condição.")
                    Exit Sub
                End If

                If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.ValorLiquido) = 0 Then
                    MsgBox(Me.Page, "Valor Liquido da Nota Fiscal não pode ser Zero.")
                    Exit Sub
                End If

                Dim parameters = New Dictionary(Of String, Object)
                parameters.Add("tipo", "FRT")
                parameters.Add("Origem", "NF")
                parameters.Add("Indice", objNotaFiscal.IndiceNota)
                ucVencimentos.Limpar()
                ucVencimentos.BindGridView(parameters)
                Popup.ConsultaDeVencimentos(Me.Page, "objVencimentosNxI" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao gerar Finaceiro da Nota Fiscal")
            lnkNovo.Parent.Visible = False
        End Try
    End Sub

    Protected Sub ddlTipoDeDocumento_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNotaFiscal()
            objNotaFiscal.CodigoTipoDeDocumento = ddlTipoDeDocumento.SelectedValue
            SessaoSalvaNotaFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub BtnExportacaoSalvar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnExportacaoSalvar.Click
        Try
            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.EXPORTACOES Then
                LimparDadosdaExportacao()
                MsgBox(Me.Page, "Os Dados de Exportacoes só sao Permitidos em operacoes cuja classe seja EXPORTACAO.")
                Exit Sub
            End If

            If txtNrDespacho.Text.Length > 0 AndAlso Not txtNrDespacho.Text.Length = 11 Then
                MsgBox(Me.Page, "Caso informe o Número do Despacho, o mesmo deve ter 11 dígitos.")
                Exit Sub
            End If

            If ddlPaisDestino.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Selecione o Pais de Destino da Mercadoria.")
                Exit Sub
            End If

            Dim exp As New [Lib].Negocio.NotaFiscalXExportacao(objNotaFiscal)

            If txtNrDespacho.Text.Length > 0 Then
                exp.NrDespachoExp = txtNrDespacho.Text
                exp.DataDespachoExp = CDate(txtDataDespacho.Text)
            Else
                exp.NrDespachoExp = ""
                exp.DataDespachoExp = Nothing
            End If

            exp.CodigoPaisDestino = ddlPaisDestino.SelectedValue
            exp.Navio = txtNavio.Text

            If String.IsNullOrWhiteSpace(txtDataAverba.Text) Then
                exp.DataAverbacao = Nothing
            Else
                exp.DataAverbacao = CDate(txtDataAverba.Text)
            End If

            exp.FaturaExportacao = Trim(txtFaturaExportacao.Text)

            'EXPORTACAO e IMPORTACAO DRAWBAK
            If (objNotaFiscal.CodigoOperacao = 35 Or objNotaFiscal.CodigoOperacao = 36) And objNotaFiscal.CodigoSubOperacao = 3 Then
                If TxtNumAtoConcessorio.Text.Length = 0 Then
                    MsgBox(Me.Page, "Informe Numero do Ato Concessório(Drawback).")
                Else
                    exp.NumAtoConcessorio = TxtNumAtoConcessorio.Text

                    If TxtDtaRegAtoConcessorio.Text.Length = 0 Then
                        MsgBox(Me.Page, "Informe a Data de registro Do Ato Concessório.")
                    Else
                        exp.DtaRegAtoConcessorio = CDate(TxtDtaRegAtoConcessorio.Text)
                    End If

                    If TxtDtaValidAtoConcessorio.Text.Length = 0 Then
                        MsgBox(Me.Page, "Informe a Data de Validade Do Ato Concessório.")
                    Else
                        exp.DtaValidAtoConcessorio = CDate(TxtDtaValidAtoConcessorio.Text)
                    End If
                End If
            End If

            If objNotaFiscal.IUD <> "I" Then
                Dim sqls As New ArrayList
                exp.IUD = "D"
                exp.SalvarSql(sqls)
                exp.IUD = "I"
                exp.SalvarSql(sqls)

                If Banco.GravaBanco(sqls) Then
                    objNotaFiscal.DadosDaExportacao = exp
                Else
                    MsgBox(Me.Page, "Erro ao Salvar dados da exportacao")
                End If

            Else
                objNotaFiscal.DadosDaExportacao = exp
            End If

            SessaoSalvaNotaFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgAdicionarRE_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If txtRe.Text.Trim.Length = 0 Then
                MsgBox(Me.Page, "Informe o Número do RE")
                Exit Sub
            ElseIf Funcoes.OnlyNumbers(txtRe.Text.Trim).Length > 12 Then
                MsgBox(Me.Page, "O Registro de Exportação não deve ultrapassar 12 números.")
                Exit Sub
            ElseIf Not IsDate(txtDataRe.Text) Then
                MsgBox(Me.Page, "Informe uma Data Válida")
                Exit Sub
            ElseIf ddlUF.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Selecione o Estado do Produtor / Fabricante")
                Exit Sub
            End If

            SessaoRecuperaNotaFiscal()
            Dim re As New NotaFiscalXRE(objNotaFiscal)
            re.IUD = "I"
            re.RegistroDeExportacao = txtRe.Text
            re.DataRegistroDeExportacao = CDate(txtDataRe.Text)
            re.UfProdutor = ddlUF.SelectedValue

            If objNotaFiscal.IUD = "I" Then
                objNotaFiscal.DadosDaExportacaoRE.Add(re)
                gridRE.DataSource = objNotaFiscal.DadosDaExportacaoRE
                gridRE.DataBind()
            Else
                If re.Salvar Then
                    re.IUD = ""
                    objNotaFiscal.DadosDaExportacaoRE.Add(re)
                    gridRE.DataSource = objNotaFiscal.DadosDaExportacaoRE
                    gridRE.DataBind()
                Else
                    MsgBox(Me.Page, "Erro Ao Incluir Registro de Exportação")
                    Exit Sub
                End If
            End If

            txtRe.Text = ""
            txtDataRe.Text = ""
            ddlUF.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnLimparExportacao_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnLimparExportacao.Click
        Try
            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.IUD = "I" Then
                objNotaFiscal.DadosDaExportacao = Nothing
                objNotaFiscal.DadosDaExportacaoRE = Nothing
            Else
                If Not objNotaFiscal.DadosDaExportacao Is Nothing Then
                    objNotaFiscal.DadosDaExportacao.IUD = "D"
                    objNotaFiscal.DadosDaExportacao.Salvar()
                    objNotaFiscal.DadosDaExportacao = Nothing
                End If

                For Each re As NotaFiscalXRE In objNotaFiscal.DadosDaExportacaoRE
                    re.IUD = "D"
                    re.Salvar()
                Next
                objNotaFiscal.DadosDaExportacaoRE = Nothing
            End If

            SessaoSalvaNotaFiscal()
            LimparDadosdaExportacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgDeletarRe_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim cancRe As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(cancRe.NamingContainer, GridViewRow)

            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.IUD <> "I" Then
                objNotaFiscal.DadosDaExportacaoRE(row.RowIndex).IUD = "D"
                If objNotaFiscal.DadosDaExportacaoRE(row.RowIndex).Salvar() Then objNotaFiscal.DadosDaExportacaoRE.RemoveAt(row.RowIndex)
            Else
                objNotaFiscal.DadosDaExportacaoRE.RemoveAt(row.RowIndex)
            End If
            SessaoSalvaNotaFiscal()
            gridRE.DataSource = objNotaFiscal.DadosDaExportacaoRE.ToArray
            gridRE.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgConfirmarDI_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            SessaoRecuperaNotaFiscal()

            If objNotaFiscal.DadosDaImportacao Is Nothing Then
                objNotaFiscal.DadosDaImportacao = New [Lib].Negocio.NotaFiscalXImportacao(objNotaFiscal)
                SessaoSalvaNotaFiscal()
            End If

            If gridItens.SelectedIndex = -1 Then
                MsgBox(Me.Page, "Produto não foi selecionado.")
            ElseIf objNotaFiscal.Itens(gridItens.SelectedIndex).CFOP > 3000 AndAlso objNotaFiscal.Itens(gridItens.SelectedIndex).CFOP < 4000 Then
                If validar_DI() Then

                    objNotaFiscal.DadosDaImportacao.NumeroDeclaracaoImportacao = txtDI.Text
                    If txtDataDI.Text.Length > 0 Then objNotaFiscal.DadosDaImportacao.DataDeclaracaoImportacao = CDate(txtDataDI.Text)
                    objNotaFiscal.DadosDaImportacao.RegistroDeImportacao = TxtRegImportacao.Text
                    If TxtDataRegImp.Text.Length > 0 Then objNotaFiscal.DadosDaImportacao.DataRegistroDeImportacao = CDate(TxtDataRegImp.Text)

                    objNotaFiscal.DadosDaImportacao.LocalEmbarqueImportacao = Trim(txtEmbarqueDI.Text).ToUpper()
                    If txtDataEmbarqueDI.Text.Length > 0 Then objNotaFiscal.DadosDaImportacao.DataEmbarqueImportacao = CDate(txtDataEmbarqueDI.Text)
                    objNotaFiscal.DadosDaImportacao.EstadoEmbarqueImportacao = txtCodigoUFEmbarqueDI.Value

                    objNotaFiscal.DadosDaImportacao.LocalDesembarqueImportacao = Trim(txtDesembarqueDI.Text).ToUpper()
                    If txtDataDesembarqueDI.Text.Length > 0 Then objNotaFiscal.DadosDaImportacao.DataDesembarqueImportacao = CDate(txtDataDesembarqueDI.Text)
                    objNotaFiscal.DadosDaImportacao.EstadoDesembarqueImportacao = txtCodigoUFDesembarqueDI.Value

                    Dim strCliImp() As String = txtCodigolblFabricanteDI.Value.ToString.Split("-")
                    objNotaFiscal.DadosDaImportacao.CodigoFabricante = strCliImp(0)
                    objNotaFiscal.DadosDaImportacao.EndFabricante = strCliImp(1)

                    '*****************************************
                    objNotaFiscal.DadosDaImportacao.NumAtoConcessorio = TxtNumAtoConcessorioImp.Text
                    If TxtDtaRegAtoConcessorioImp.Text.Length > 0 Then objNotaFiscal.DadosDaImportacao.DtaRegAtoConcessorio = CDate(TxtDtaRegAtoConcessorioImp.Text)
                    If TxtDtaValidAtoConcessorioImp.Text.Length > 0 Then objNotaFiscal.DadosDaImportacao.DtaValidAtoConcessorio = CDate(TxtDtaValidAtoConcessorioImp.Text)

                    objNotaFiscal.DadosDaImportacao.NrFatura = TxtNrFaturaImp.Text

                    objNotaFiscal.DadosDaImportacao.ViaDeTransporte = ddlViaDeTransporteDI.SelectedValue
                    objNotaFiscal.DadosDaImportacao.TipoDeImportacao = ddlTipoDeImportacaoDI.SelectedValue

                    objNotaFiscal.DadosDaImportacao.ValorVAFRMM = txtValorVAFRMMDI.Text

                    SessaoSalvaNotaFiscal()
                    imgConfirmarDI.Enabled = False
                    TabContainer1.ActiveTabIndex = 0
                Else
                    MsgBox(Me.Page, erroMsg)
                End If
            Else
                MsgBox(Me.Page, "Declaração de Importação só é obrigatória para Operações com CFOP na casa de 3000.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ImgAddConhec_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If txtNrConhecimento.Text.Trim.Length = 0 Then
                MsgBox(Me.Page, "Informe o Numero do Conhec. De Embarque")
                Exit Sub
            End If

            If Not IsDate(txtDataNrConhecim.Text) Then
                MsgBox(Me.Page, "Informe uma Data Para Conhecimento Valida")
                Exit Sub
            End If

            If ddlTipoConhec.SelectedIndex = 0 Or ddlTipoConhec.Text.Trim = "" Then
                MsgBox(Me.Page, "Selecione o Tipo do Conhecimento")
                Exit Sub
            End If

            SessaoRecuperaNotaFiscal()
            Dim CE As New NotaFiscalXCE(objNotaFiscal)
            CE.IUD = "I"
            CE.ConhecimentoDeEmbarque = txtNrConhecimento.Text
            CE.DataConhecimento = CDate(txtDataNrConhecim.Text)
            CE.TipoConhecimento = ddlTipoConhec.SelectedValue

            If objNotaFiscal.IUD = "I" Then
                objNotaFiscal.DadosDaExportacaoCE.Add(CE)
                GridConhecimento.DataSource = objNotaFiscal.DadosDaExportacaoCE
                GridConhecimento.DataBind()
            Else
                If CE.Salvar Then
                    CE.IUD = ""
                    objNotaFiscal.DadosDaExportacaoCE.Add(CE)
                    GridConhecimento.DataSource = objNotaFiscal.DadosDaExportacaoCE
                    GridConhecimento.DataBind()
                Else
                    MsgBox(Me.Page, "Erro Ao Incluir Conhecimento De Embarque")
                    Exit Sub
                End If
            End If

            txtNrConhecimento.Text = ""
            txtDataNrConhecim.Text = ""
            ddlTipoConhec.SelectedValue = 10
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ImgExcluirConhec_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim cancCE As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(cancCE.NamingContainer, GridViewRow)

            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.IUD <> "I" Then
                objNotaFiscal.DadosDaExportacaoCE(row.RowIndex).IUD = "D"
                If objNotaFiscal.DadosDaExportacaoCE(row.RowIndex).Salvar() Then objNotaFiscal.DadosDaExportacaoCE.RemoveAt(row.RowIndex)
            Else
                objNotaFiscal.DadosDaExportacaoCE.RemoveAt(row.RowIndex)
            End If
            SessaoSalvaNotaFiscal()
            GridConhecimento.DataSource = objNotaFiscal.DadosDaExportacaoCE.ToArray
            GridConhecimento.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub bntBuscarNFReferencial_Click(ByVal sender As Object, ByVal e As EventArgs) Handles bntBuscarNFReferencial.Click
        Try
            ucNFReferencialSaida.BindGridView()
            Popup.ConsultaNFReferencialSaida(Me, "objNFReferencialSaida" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub bntBuscarNFProdutor_Click(ByVal sender As Object, ByVal e As EventArgs) Handles bntBuscarNFProdutor.Click
        Try
            ucNFProdutor.BindGridView()
            Popup.ConsultaNFProdutor(Me, "objNFDeProdutor" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub TxtQtdeEmbalagem_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim txtQtdeEmb As TextBox = CType(sender, TextBox)
            Dim row As GridViewRow = CType(txtQtdeEmb.NamingContainer, GridViewRow)

            SessaoRecuperaLoteEmbalagem()
            objListaDeCombinacoesLoteEmbalagem(row.RowIndex).QtdeEmbalagem = txtQtdeEmb.Text

            If Not objListaDeCombinacoesLoteEmbalagem(row.RowIndex).PesoVariavel Then
                objListaDeCombinacoesLoteEmbalagem(row.RowIndex).QtdeEmbXCapacidade = objListaDeCombinacoesLoteEmbalagem(row.RowIndex).QtdeEmbalagem * objListaDeCombinacoesLoteEmbalagem(row.RowIndex).CapacidadeEmbalagem
                CType(gridEmbalagem.Rows(row.RowIndex).FindControl("txtQtdeDeProduto"), TextBox).Text = objListaDeCombinacoesLoteEmbalagem(row.RowIndex).QtdeEmbXCapacidade.ToString("N4")
            Else
                objListaDeCombinacoesLoteEmbalagem(row.RowIndex).QtdeEmbXCapacidade = objListaDeCombinacoesLoteEmbalagem(row.RowIndex).QtdeEmbalagem * objListaDeCombinacoesLoteEmbalagem(row.RowIndex).PesoSaco
                CType(gridEmbalagem.Rows(row.RowIndex).FindControl("txtQtdeDeProduto"), TextBox).Text = objListaDeCombinacoesLoteEmbalagem(row.RowIndex).QtdeEmbXCapacidade.ToString("N4")
            End If

            SessaoSalvaLoteEmbalagem()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlDeposito_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlDeposito.SelectedIndex > 0 Then
                SessaoRecuperaNotaFiscal()
                Dim dep() As String = ddlDeposito.SelectedValue.ToString.Split("-")
                objNotaFiscal.CodigoDeposito = dep(0)
                objNotaFiscal.EnderecoDeposito = dep(1)
                objNotaFiscal.Deposito = New [Lib].Negocio.Cliente(objNotaFiscal.CodigoDeposito, objNotaFiscal.EnderecoDeposito)
                SessaoSalvaNotaFiscal()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlOrigemDestino_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlOrigemDestino.SelectedIndexChanged
        Try
            If ddlOrigemDestino.SelectedIndex > 0 Then
                SessaoRecuperaNotaFiscal()
                Dim dep() As String = ddlOrigemDestino.SelectedValue.ToString.Split("-")
                objNotaFiscal.CodigoDestino = dep(0)
                objNotaFiscal.EnderecoDestino = dep(1)
                objNotaFiscal.Destino = New [Lib].Negocio.Cliente(objNotaFiscal.CodigoDestino, objNotaFiscal.EnderecoDestino)
                SessaoSalvaNotaFiscal()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEmbarque_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNotaFiscal()
            If ddlEmbarque.SelectedIndex > 0 Then
                Dim dep() As String = ddlEmbarque.SelectedValue.ToString.Split("-")
                objNotaFiscal.CodigoLocalEmbarque = dep(0)
                objNotaFiscal.EndLocalEmbarque = dep(1)
            Else
                objNotaFiscal.CodigoLocalEmbarque = String.Empty
            End If
            SessaoSalvaNotaFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlTransbordo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlTransbordo.SelectedIndex > 0 Then
                SessaoRecuperaNotaFiscal()
                Dim dep() As String = ddlTransbordo.SelectedValue.ToString.Split("-")
                objNotaFiscal.CodigoTransbordo = dep(0)
                objNotaFiscal.EnderecoTransbordo = dep(1)
                objNotaFiscal.Transbordo = New [Lib].Negocio.Cliente(objNotaFiscal.CodigoTransbordo, objNotaFiscal.EnderecoTransbordo)
                SessaoSalvaNotaFiscal()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEntrega_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlEntrega.SelectedIndexChanged
        Try
            If ddlEntrega.SelectedIndex > 0 Then
                SessaoRecuperaNotaFiscal()
                Dim dep() As String = ddlEntrega.SelectedValue.ToString.Split("-")
                objNotaFiscal.CodigoEntrega = dep(0)
                objNotaFiscal.EnderecoEntrega = dep(1)
                objNotaFiscal.Entrega = New [Lib].Negocio.Cliente(objNotaFiscal.CodigoEntrega, objNotaFiscal.EnderecoEntrega)
                SessaoSalvaNotaFiscal()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgDelNFReferencial_Click(sender As Object, e As ImageClickEventArgs)
        Try
            Dim cancRe As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(cancRe.NamingContainer, GridViewRow)

            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.IUD <> "I" Then
                objNotaFiscal.NotasReferenciais(row.RowIndex).IUD = "D"
                If objNotaFiscal.NotasReferenciais(row.RowIndex).Salvar() Then objNotaFiscal.NotasReferenciais.RemoveAt(row.RowIndex)
            Else
                objNotaFiscal.NotasReferenciais.RemoveAt(row.RowIndex)
            End If
            SessaoSalvaNotaFiscal()

            CompoeObsDasNotasFiscaisReferenciais()

            grdNotasReferenciais.DataSource = objNotaFiscal.NotasReferenciais
            grdNotasReferenciais.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnSalvarNotasReferenciais_Click(sender As Object, e As EventArgs) Handles btnSalvarNotasReferenciais.Click
        Try
            SessaoRecuperaNotaFiscal()
            Try
                If objNotaFiscal.IUD <> "I" Then
                    If objNotaFiscal.NotasReferenciais.Count > 0 Then
                        For Each objNfRef In objNotaFiscal.NotasReferenciais
                            If Not objNfRef.Salvar() Then
                                Throw New Exception("Falha!!!")
                            End If
                        Next
                    End If
                Else
                    MsgBox(Me.Page, "Inclua a Nota Fiscal Principal antes de salvar as Notas Referenciais!")
                    Exit Sub
                End If
                MsgBox(Me.Page, "Gravação das Notas Referenciais concluída com sucesso!!!")
            Catch ex As Exception
                MsgBox(Me.Page, "Falha na gravação das Notas Referenciais!!!")
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grdNotasReferenciais_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles grdNotasReferenciais.RowDataBound
        Try
            SessaoRecuperaNotaFiscal()
            If e.Row.RowType = DataControlRowType.Footer Then
                e.Row.Cells(1).HorizontalAlign = HorizontalAlign.Right
                e.Row.Cells(1).Text = "Totais:"

                e.Row.Cells(2).HorizontalAlign = HorizontalAlign.Right
                e.Row.Cells(2).Text = objNotaFiscal.NotasReferenciais.QuantidadeTotal.ToString("#,0.00")

                e.Row.Cells(3).HorizontalAlign = HorizontalAlign.Right
                e.Row.Cells(3).Text = objNotaFiscal.NotasReferenciais.ValorTotal.ToString("c")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Métodos"

    Private Function ConstroiDataSetNota() As DataSet

        SessaoRecuperaNotaFiscal()

        Dim dsNota As New DataSet
        Dim dtNota As New DataTable("NotaFiscal")
        Dim dtitem As New DataTable("ItemDaNota")
        Dim dtVencimentos As New DataTable("Vencimentos")

        'Nota
        dtNota.Columns.Add("NaturezaOperacaoNota", GetType(String))
        dtNota.Columns.Add("NomeEmpresa", GetType(String))
        dtNota.Columns.Add("EnderecoEmpresa", GetType(String))
        dtNota.Columns.Add("NumeroEmpresa", GetType(String))
        dtNota.Columns.Add("ComplementoEmpresa", GetType(String))
        dtNota.Columns.Add("BairroEmpresa", GetType(String))
        dtNota.Columns.Add("EstadoEmpresa", GetType(String))
        dtNota.Columns.Add("CEPEmpresa", GetType(String))
        dtNota.Columns.Add("IEEmpresa", GetType(String))
        dtNota.Columns.Add("CNPJEmpresa", GetType(String))
        dtNota.Columns.Add("NumeroNota", GetType(String))
        dtNota.Columns.Add("EntradaSaidaNota", GetType(String))
        dtNota.Columns.Add("NomeCliente", GetType(String))
        dtNota.Columns.Add("CNPJCliente", GetType(String))
        dtNota.Columns.Add("DataEmissaoNota", GetType(DateTime))
        dtNota.Columns.Add("SerieNota", GetType(String))
        dtNota.Columns.Add("EnderecoCliente", GetType(String))
        dtNota.Columns.Add("ComplementoCliente", GetType(String))
        dtNota.Columns.Add("BairroCliente", GetType(String))
        dtNota.Columns.Add("CEPCliente", GetType(String))
        dtNota.Columns.Add("DataSaidaNota", GetType(DateTime))
        dtNota.Columns.Add("MunicipioCliente", GetType(String))
        dtNota.Columns.Add("TelefoneCliente", GetType(String))
        dtNota.Columns.Add("EstadoCliente", GetType(String))
        dtNota.Columns.Add("IECliente", GetType(String))
        dtNota.Columns.Add("HoraSaidaNota", GetType(String))
        dtNota.Columns.Add("SituacaoNota", GetType(String))
        dtNota.Columns.Add("TipoDocumento", GetType(String))
        dtNota.Columns.Add("BaseCalculoICMSNota", GetType(Decimal))
        dtNota.Columns.Add("ValorICMSNota", GetType(Decimal))
        dtNota.Columns.Add("BaseICMSNota", GetType(Decimal))
        dtNota.Columns.Add("BaseICMSSTNota", GetType(Decimal))
        dtNota.Columns.Add("ValorICMSSTNota", GetType(Decimal))
        dtNota.Columns.Add("TotalProdutosNota", GetType(Decimal))
        dtNota.Columns.Add("ValorFreteNota", GetType(Decimal))
        dtNota.Columns.Add("ValorSeguroNota", GetType(Decimal))
        dtNota.Columns.Add("DescontoNota", GetType(Decimal))
        dtNota.Columns.Add("OutrasDespesas", GetType(Decimal))
        dtNota.Columns.Add("ValorIPINota", GetType(Decimal))
        dtNota.Columns.Add("ValorTotalNota", GetType(Decimal))
        dtNota.Columns.Add("NomeTransportador", GetType(String))
        dtNota.Columns.Add("FretePorConta", GetType(String))
        dtNota.Columns.Add("PlacasTransportador", GetType(String))
        dtNota.Columns.Add("EstadoPlacaTransportador", GetType(String))
        dtNota.Columns.Add("EnderecoTransportador", GetType(String))
        dtNota.Columns.Add("CidadeTransportador", GetType(String))
        dtNota.Columns.Add("EstadoTransportador", GetType(String))
        dtNota.Columns.Add("CNPJTransportador", GetType(String))
        dtNota.Columns.Add("IETransportador", GetType(String))
        dtNota.Columns.Add("QuantidadeTransportador", GetType(Decimal))
        dtNota.Columns.Add("EspecieTransportador", GetType(String))
        dtNota.Columns.Add("MarcaTransportador", GetType(String))
        dtNota.Columns.Add("NumeracaodoTransportador", GetType(String))
        dtNota.Columns.Add("PesoRomaneioTransportador", GetType(Decimal))
        dtNota.Columns.Add("PesoBrutoTransportador", GetType(Decimal))
        dtNota.Columns.Add("PesoLiquidoTransportador", GetType(Decimal))
        dtNota.Columns.Add("DadosAdicionais", GetType(String))
        dtNota.Columns.Add("ObservacoesFiscais", GetType(String))
        dtNota.Columns.Add("ObservacoesInternas", GetType(String))
        dtNota.Columns.Add("Logotipo", GetType(System.Byte()))

        'Item Nota
        dtitem.Columns.Add("CodigoProduto", GetType(String))
        dtitem.Columns.Add("NomeProduto", GetType(String))
        dtitem.Columns.Add("LoteProduto", GetType(String))
        dtitem.Columns.Add("NCMProduto", GetType(String))
        dtitem.Columns.Add("CSTProduto", GetType(String))
        dtitem.Columns.Add("CFOPProduto", GetType(String))
        dtitem.Columns.Add("UnidadeProduto", GetType(String))
        dtitem.Columns.Add("QuantidadeProduto", GetType(Decimal))
        dtitem.Columns.Add("UnitarioProduto", GetType(Decimal))
        dtitem.Columns.Add("TotalProduto", GetType(Decimal))
        dtitem.Columns.Add("BaseCalculoICMSProduto", GetType(Decimal))
        dtitem.Columns.Add("ValorICMSProduto", GetType(Decimal))
        dtitem.Columns.Add("IPIProduto", GetType(Decimal))
        dtitem.Columns.Add("AliquotaICMSProduto", GetType(Decimal))
        dtitem.Columns.Add("AliquotaIPIProduto", GetType(Decimal))

        'Vencimentos Nota
        dtVencimentos.Columns.Add("DataProrrogacao", GetType(String))
        dtVencimentos.Columns.Add("Codigo", GetType(Integer))
        dtVencimentos.Columns.Add("ValorDoDocumento", GetType(Decimal))

        Dim newNota As DataRow = dtNota.NewRow()

        newNota.Item("NaturezaOperacaoNota") = txtNaturezaDaOperacao.Text
        newNota.Item("NomeEmpresa") = txtNomeDaEmpresa.Text
        newNota.Item("EnderecoEmpresa") = objNotaFiscal.Empresa.Endereco
        newNota.Item("NumeroEmpresa") = objNotaFiscal.Empresa.Numero
        newNota.Item("ComplementoEmpresa") = objNotaFiscal.Empresa.Complemento
        newNota.Item("BairroEmpresa") = objNotaFiscal.Empresa.Bairro
        newNota.Item("EstadoEmpresa") = objNotaFiscal.Empresa.Estado.Codigo
        newNota.Item("CEPEmpresa") = objNotaFiscal.Empresa.CEP
        newNota.Item("NomeEmpresa") = txtNomeDaEmpresa.Text
        newNota.Item("IEEmpresa") = txtInscricaoDaEmpresa.Text
        newNota.Item("CNPJEmpresa") = txtCnpjDaEmpresa.Text.Replace(".", "").Replace("/", "").Replace("-", "")
        newNota.Item("NumeroNota") = txtNumero.Text
        newNota.Item("EntradaSaidaNota") = txtES.Text
        newNota.Item("NomeCliente") = txtNomeDoCliente.Text
        newNota.Item("CNPJCliente") = txtCnpjDoCliente.Text.Replace(".", "").Replace("/", "").Replace("-", "")
        newNota.Item("DataEmissaoNota") = txtDataDeEmissao.Text
        newNota.Item("SerieNota") = txtSerie.Text
        newNota.Item("EnderecoCliente") = txtEnderecoDoCliente.Text
        newNota.Item("ComplementoCliente") = txtComplementoDoCliente.Text
        newNota.Item("BairroCliente") = txtBairroDoCliente.Text
        newNota.Item("CEPCliente") = txtCepDoCliente.Text
        newNota.Item("DataSaidaNota") = txtDataDeEntrada.Text
        newNota.Item("MunicipioCliente") = txtCidadeDoCliente.Text
        newNota.Item("TelefoneCliente") = txtTelefoneDoCliente.Text
        newNota.Item("EstadoCliente") = txtEstadoDoCliente.Text
        newNota.Item("IECliente") = txtInscricaoDoCliente.Text
        newNota.Item("HoraSaidaNota") = txtHoraDaSaida.Text
        newNota.Item("SituacaoNota") = txtSituacao.Text
        newNota.Item("TipoDocumento") = ddlTipoDeDocumento.SelectedItem
        newNota.Item("BaseCalculoICMSNota") = txtBaseIcmsNota.Text
        newNota.Item("ValorICMSNota") = txtValorIcmsNota.Text
        newNota.Item("BaseICMSNota") = txtBaseIcmsNota.Text
        newNota.Item("BaseICMSSTNota") = txtValorIcmsSTNota.Text
        newNota.Item("ValorICMSSTNota") = txtValorBaseIcmsSTNota.Text
        newNota.Item("TotalProdutosNota") = txtValorTotalDosProdutos.Text
        newNota.Item("ValorFreteNota") = txtValorFrete.Text
        newNota.Item("ValorSeguroNota") = txtSeguro.Text
        newNota.Item("DescontoNota") = txtDesconto.Text
        newNota.Item("OutrasDespesas") = txtOutras.Text
        newNota.Item("ValorIPINota") = txtValorIPINota.Text
        newNota.Item("ValorTotalNota") = txtValorTotalDaNota.Text
        newNota.Item("NomeTransportador") = txtNomeDoTransportador.Text
        newNota.Item("FretePorConta") = ddlFrete.SelectedItem
        newNota.Item("PlacasTransportador") = txtPlacas.Text
        newNota.Item("EstadoPlacaTransportador") = txtEstadoDaPlaca.Text
        newNota.Item("EnderecoTransportador") = txtEnderecoDoTransportador.Text
        newNota.Item("CidadeTransportador") = txtCidadeDoTransportador.Text
        newNota.Item("EstadoTransportador") = txtEstadoDoTransportador.Text
        newNota.Item("CNPJTransportador") = txtCnpjDoTransportador.Text.Replace(".", "").Replace("/", "").Replace("-", "")
        newNota.Item("IETransportador") = txtInscricaoDoTransportador.Text
        newNota.Item("QuantidadeTransportador") = txtVolumes.Text
        newNota.Item("EspecieTransportador") = txtEspecie.Text
        newNota.Item("MarcaTransportador") = txtMarca.Text
        newNota.Item("NumeracaodoTransportador") = txtNumeracao.Text

        If txtPesoRomaneio.Text.Length > 0 Then
            newNota.Item("PesoRomaneioTransportador") = txtPesoRomaneio.Text
        Else
            newNota.Item("PesoRomaneioTransportador") = 0
        End If

        If txtPesoBruto.Text.Length > 0 Then
            newNota.Item("PesoBrutoTransportador") = txtPesoBruto.Text
        Else
            newNota.Item("PesoBrutoTransportador") = 0
        End If

        If txtPesoLiquido.Text.Length > 0 Then
            newNota.Item("PesoLiquidoTransportador") = txtPesoLiquido.Text
        Else
            newNota.Item("PesoLiquidoTransportador") = 0
        End If

        newNota.Item("DadosAdicionais") = objNotaFiscal.ObservacoesDeEmbarque
        newNota.Item("ObservacoesFiscais") = txtObservacoesFiscais.Text
        newNota.Item("ObservacoesInternas") = txtObservacoesFiscais.Text

        Dim strCaminhoImagem As String = HttpContext.Current.Server.MapPath("~/Images/" & objNotaFiscal.Empresa.Imagem)
        newNota.Item("Logotipo") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)

        dtNota.Rows.Add(newNota)

        For Each item In objNotaFiscal.Itens

            Dim newItem As DataRow = dtitem.NewRow()

            newItem.Item("CodigoProduto") = item.CodigoProduto

            Dim produtoNota As Produto = New Produto(item.CodigoProduto)

            newItem.Item("NomeProduto") = produtoNota.Nome

            Dim nLotes As String = String.Empty

            For Each l In item.Lotes

                If nLotes.Length > 0 Then
                    nLotes += "," & l.Lote
                Else
                    nLotes += l.Lote
                End If

            Next

            newItem.Item("LoteProduto") = nLotes
            newItem.Item("NCMProduto") = item.Produto.NCM
            newItem.Item("CSTProduto") = item.OperacaoEstado.CodigoSTICMS
            newItem.Item("CFOPProduto") = item.OperacaoEstado.CodigoFiscal
            newItem.Item("UnidadeProduto") = item.Produto.Unidade
            newItem.Item("QuantidadeProduto") = item.QuantidadeFiscal
            newItem.Item("UnitarioProduto") = item.Unitario
            newItem.Item("TotalProduto") = item.ValorTotal
            newItem.Item("BaseCalculoICMSProduto") = (From e In item.Encargos Where e.Encargo.Codigo = "ICMS" OrElse e.Encargo.Codigo = "ICMS A REC." Select e.Base).FirstOrDefault
            newItem.Item("ValorICMSProduto") = (From e In item.Encargos Where e.Encargo.Codigo = "ICMS" OrElse e.Encargo.Codigo = "ICMS A REC." Select e.Valor).FirstOrDefault
            newItem.Item("IPIProduto") = (From e In item.Encargos Where e.Encargo.Codigo = "IPI" Select e.Valor).FirstOrDefault
            newItem.Item("AliquotaICMSProduto") = (From e In item.Encargos Where e.Encargo.Codigo = "ICMS" OrElse e.Encargo.Codigo = "ICMS A REC." Select e.Percentual).FirstOrDefault
            newItem.Item("AliquotaIPIProduto") = (From e In item.Encargos Where e.Encargo.Codigo = "IPI" Select e.Percentual).FirstOrDefault

            dtitem.Rows.Add(newItem)

        Next

        bloco1 = ""
        bloco2 = ""
        bloco3 = ""
        Dim contVencimento As Integer = 1

        For Each vencimento In objNotaFiscal.VencimentosNota

            Dim newVencimento As DataRow = dtVencimentos.NewRow()

            newVencimento.Item("DataProrrogacao") = vencimento.Prorrogacao
            newVencimento.Item("Codigo") = contVencimento
            newVencimento.Item("ValorDoDocumento") = vencimento.ValorDoDocumento

            dtVencimentos.Rows.Add(newVencimento)

            If contVencimento <= 5 Then
                bloco1 &= vencimento.Prorrogacao & "            " & Format(contVencimento, "000") & "            " & Format(vencimento.ValorDoDocumento, "#,##0.00") & Environment.NewLine
            ElseIf contVencimento > 5 And contVencimento <= 10 Then
                bloco2 &= vencimento.Prorrogacao & "            " & Format(contVencimento, "000") & "            " & Format(vencimento.ValorDoDocumento, "#,##0.00") & Environment.NewLine
            Else
                bloco3 &= vencimento.Prorrogacao & "            " & Format(contVencimento, "000") & "            " & Format(vencimento.ValorDoDocumento, "#,##0.00") & Environment.NewLine
            End If

            contVencimento = contVencimento + 1

        Next

        dsNota.Tables.Add(dtNota)
        dsNota.Tables.Add(dtitem)
        dsNota.Tables.Add(dtVencimentos)

        Return dsNota

    End Function

    Private Sub carregarNotaXMLautomatico(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not chaveXMLautomatico Is Nothing Then
            chk_nfe.Checked = True
            txtChaveNFe.Text = chaveXMLautomatico
            lnkVerificarChaveNFE_Click(sender, e)
        End If
    End Sub

    Private Sub ImprimirEspelho()
        Try
            SessaoRecuperaNotaFiscal()
            Dim espelho As New [Lib].Negocio.NotaFiscalEspelho
            espelho.ExibirEspelho(Me.Page, objNotaFiscal)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub LimparGrid()
        gridItens.DataBind()
        'DgEncargos.DataBind()
        gridVencimentosPedido.DataBind()
        gridVencimentosNota.DataBind()
    End Sub

    Private Sub IniciarAlteracaoNotaFiscal()
        SessaoRecuperaNotaFiscal()

        If Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, txtDataDeEntrada.Text, "NOTAS FISCAIS") Then
            MsgBox(Me.Page, "Movimento de Notas Fiscais já Fechado para esta data...")
        ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, txtDataDeEntrada.Text, "CONTABIL") Then
            MsgBox(Me.Page, "Movimento Contábil já Fechado para esta data...")
        ElseIf objNotaFiscal.NotaFiscalOriginal.NotasOrigemDestino.Where(Function(s) s.TipoDeDocumento.Codigo = eTipoDeDocumento.CT_E).Any Then  'VinculadaCTE(objNotaFiscal.NotaFiscalOriginal.TipoDeDocumento) Then
            BtnCancelarSefaz.Enabled = False
            lnkExcluir.Parent.Visible = False
            Dim CodigoCte As Integer = objNotaFiscal.NotaFiscalOriginal.NotasOrigemDestino.Where(Function(s) s.TipoDeDocumento.Codigo = eTipoDeDocumento.CT_E).FirstOrDefault.Codigo
            MsgBox(Me.Page, "Não é possível alterar a nota fiscal " & objNotaFiscal.Codigo & ", pois está vinculada ao conhecimento de transporte " & CodigoCte & "!")
        ElseIf objNotaFiscal.NotaFiscalOriginal.NotasOrigemDestino.Where(Function(s) s.TipoDeDocumento.Codigo = eTipoDeDocumento.CTRC).Any Then  'VinculadaCTE(objNotaFiscal.NotaFiscalOriginal.TipoDeDocumento) Then
            BtnCancelarSefaz.Enabled = False
            lnkExcluir.Parent.Visible = False
            Dim CodigoCte As Integer = objNotaFiscal.NotaFiscalOriginal.NotasOrigemDestino.Where(Function(s) s.TipoDeDocumento.Codigo = eTipoDeDocumento.CTRC).FirstOrDefault.Codigo
            MsgBox(Me.Page, "Não é possível alterar a nota fiscal " & objNotaFiscal.Codigo & ", pois está vinculada ao conhecimento de transporte " & CodigoCte & "!")
        Else
            For Each row As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
                If row.Produto.ControlarNumeroDoLote AndAlso row.SubOperacao.ControlarNumeroDoLote OrElse (row.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso row.SubOperacao.ControlarNumeroDoLote AndAlso (row.Produto.CodigoGrupo = "10101" Or row.Produto.CodigoGrupo = "10102" Or row.Produto.CodigoGrupo = "30101" Or row.Produto.CodigoGrupo = "30102")) Then
                    If row.QuantidadeFiscal <> row.Lotes.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.Quantidade) Then
                        MsgBox(Me.Page, "Quantidade do Lote informada diferente da quantidade do Produto, verifique o item " & row.CodigoProduto & "-" & row.Produto.Nome, eTitulo.Info)
                        Exit Sub
                    End If
                End If
            Next

            Dim original As [Lib].Negocio.NotaFiscal = objNotaFiscal.NotaFiscalOriginal

            If objNotaFiscal.CodigoAutorizacao > 0 Then
                If objNotaFiscal.Autorizacao.SaldoFiscal + IIf(objNotaFiscal.Itens(0).PesoQuantidade = "P", original.Itens(0).PesoFiscal < objNotaFiscal.Itens(0).PesoFiscal, original.Itens(0).QuantidadeFiscal < objNotaFiscal.Itens(0).QuantidadeFiscal) Then
                    MsgBox(Me.Page, "Saldo Fiscal da Autorizacao Insuficiente para geracao da Nota, Saldo:" & objNotaFiscal.Autorizacao.SaldoFiscal + original.Itens(0).PesoFiscal & " Nota Fiscal: " & objNotaFiscal.Itens(0).PesoFiscal, eTitulo.Info)
                    Exit Sub
                End If
                If objNotaFiscal.Autorizacao.SaldoFisico + original.Itens(0).PesoLiquido < objNotaFiscal.Itens(0).PesoLiquido Then
                    MsgBox(Me.Page, "Saldo Fisico da Autorizacao Insuficiente para geracao da Nota, Saldo:" & objNotaFiscal.Autorizacao.SaldoFisico + original.Itens(0).PesoLiquido & " Nota Fiscal: " & objNotaFiscal.Itens(0).PesoLiquido, eTitulo.Info)
                    Exit Sub
                End If
            End If

            If Not objNotaFiscal.SubOperacao.Devolucao AndAlso
                     objNotaFiscal.SubOperacao.Financeiro AndAlso
                     objNotaFiscal.Pedido.MomentoFinanceiro = 3 Then

                If objNotaFiscal.VencimentosNota Is Nothing Then
                    MsgBox(Me.Page, "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!", eTitulo.Info)
                    Exit Sub
                ElseIf objNotaFiscal.VencimentosNota.Count = 0 Then
                    MsgBox(Me.Page, "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!", eTitulo.Info)
                    Exit Sub
                ElseIf Not objNotaFiscal.VencimentosNota.Sum(Function(s) s.ValorDoDocumento) = objNotaFiscal.TotalNota Then
                    MsgBox(Me.Page, "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!", eTitulo.Info)
                    Exit Sub
                End If
            End If

            If String.IsNullOrWhiteSpace(objNotaFiscal.ObservacoesControleInterno) Then
                objNotaFiscal.ObservacoesControleInterno = "ALTERADO NO PROCESSO NotaFiscalXItens EM " & Now().ToString("yyyy-MM-dd HH:mm:ss") & " PELO IP " & lblAcessoUsuario.Text
            Else
                objNotaFiscal.ObservacoesControleInterno = objNotaFiscal.ObservacoesControleInterno & ". ALTERADO NO PROCESSO NotaFiscalXItens EM " & Now().ToString("yyyy-MM-dd HH:mm:ss") & " PELO IP " & lblAcessoUsuario.Text
            End If

            If objNotaFiscal.NossaEmissao Then
                If objNotaFiscal.Salvar Then
                    LimparCampos("U", objNotaFiscal.Codigo, False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                Else
                    MsgBox(Me.Page, "Erro ao Alterar a Nota: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage" & HID.Value).ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            Else

                Dim Sqls As New ArrayList

                original.IUD = "D"
                original.CarregandoNota = True
                original.SalvarSql(Sqls)

                'Arquivo
                If ucFile.Parent.Visible Then
                    ucFile.Salvar(objNotaFiscal.Arquivos)
                End If

                objNotaFiscal.IUD = "I"

                objNotaFiscal.CarregandoNota = True
                If String.IsNullOrWhiteSpace(txtSerie.Text) Then
                    MsgBox(Me.Page, "Série Nota Fiscal não foi informada.")
                    Exit Sub
                End If
                objNotaFiscal.Serie = Trim(txtSerie.Text)
                If txtSerie.Text.Length > 0 Then objNotaFiscal.SerieNotaProdutor = Trim(txtSerie.Text)
                txtSerie.Text = objNotaFiscal.Serie
                objNotaFiscal.UsuarioInclusao = original.UsuarioInclusao
                objNotaFiscal.DataInclusao = original.DataInclusao
                objNotaFiscal.UsuarioAlteracao = Session("ssNomeUsuario")
                objNotaFiscal.DataAlteracao = Date.Now

                If Not objNotaFiscal.VencimentosNota Is Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                    For Each dr As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosNota
                        dr.IUD = "I"
                    Next
                    objNotaFiscal.VencimentosNota.NF = objNotaFiscal
                    'Atualiza a provisão do pedido.
                    For Each dr As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosPedido.Where(Function(t) t.CodigoProvisao = eProvisao.Provisao)
                        dr.IUD = IIf(dr.IUD = "I", "I", "U")
                        If dr.IUD = "U" Then dr.SalvarSql(Sqls, False)
                    Next

                End If

                objNotaFiscal.SalvarSql(Sqls)

                If Banco.GravaBanco(Sqls) Then
                    If chkEspelho.Checked Then
                        Dim espelho As New NotaFiscalEspelho
                        espelho.ExibirEspelho(Me.Page, objNotaFiscal)
                    End If
                    LimparCampos("U", objNotaFiscal.Codigo, False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                Else
                    MsgBox(Me.Page, "Erro ao Alterar a Nota: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage" & HID.Value).ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            End If
        End If

    End Sub

    Public Sub RemoverArquivoBD()
        SessaoRecuperaNotaFiscal()

        If Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento, "NOTAS FISCAIS") Then
            MsgBox(Me.Page, "Movimento da Nota Fiscal já fechado para esta data", eTitulo.Info)
        Else
            If objNotaFiscal.IUD = "U" Then
                ucFile.Salvar(objNotaFiscal.Arquivos)

                Dim Sqls As New ArrayList

                For Each arq In objNotaFiscal.Arquivos
                    If arq.IUD = "D" Then
                        arq.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                        arq.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                        arq.CodigoCliente = objNotaFiscal.CodigoCliente
                        arq.EnderecoCliente = objNotaFiscal.EnderecoCliente
                        arq.CodigoNota = objNotaFiscal.Codigo
                        arq.Serie = objNotaFiscal.Serie
                        arq.CodigoPedido = objNotaFiscal.CodigoPedido
                        arq.SalvarSql(Sqls)
                    End If
                Next

                If Sqls.Count > 0 Then
                    If Banco.GravaBanco(Sqls) Then
                        SessaoSalvaNotaFiscal()
                        MsgBox(Me.Page, "Arquivo removido com sucesso. Caso tenha consultado a nota apenas para remover o arquivo, não precisa clicar em alterar pois o mesmo já foi removido.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub AdicionarArquivoBD()
        SessaoRecuperaNotaFiscal()

        If Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento, "NOTAS FISCAIS") Then
            MsgBox(Me.Page, "Movimento da Nota Fiscal já fechado para esta data", eTitulo.Info)
        Else
            If objNotaFiscal.IUD = "U" Then
                ucFile.Salvar(objNotaFiscal.Arquivos)

                Dim Sqls As New ArrayList

                For Each arq In objNotaFiscal.Arquivos
                    If arq.IUD = "I" Then
                        arq.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                        arq.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                        arq.CodigoCliente = objNotaFiscal.CodigoCliente
                        arq.EnderecoCliente = objNotaFiscal.EnderecoCliente
                        arq.CodigoNota = objNotaFiscal.Codigo
                        arq.Serie = objNotaFiscal.Serie
                        arq.CodigoPedido = objNotaFiscal.CodigoPedido
                        arq.SalvarSql(Sqls)
                        arq.IUD = "U"
                    End If
                Next

                If Sqls.Count > 0 Then
                    If Banco.GravaBanco(Sqls) Then
                        SessaoSalvaNotaFiscal()
                        MsgBox(Me.Page, "Arquivo adicionado com sucesso. Caso tenha consultado a nota apenas para adicionar o arquivo, não precisa clicar em alterar pois o mesmo já foi adicionado.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub IniciarExclusaoNotaFiscal()
        SessaoRecuperaNotaFiscal()

        If Funcoes.VerificaAcesso(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, txtDataDeEntrada.Text, "NOTAS FISCAIS") = False Then
            MsgBox(Me.Page, "Movimento de Notas Fiscais já Fechado para esta data...")
        ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, txtDataDeEntrada.Text, "CONTABIL") Then
            MsgBox(Me.Page, "Movimento Contábil já Fechado para esta data...")
        ElseIf objNotaFiscal.NotaFiscalOriginal.NotasOrigemDestino.Where(Function(s) s.TipoDeDocumento.Codigo = eTipoDeDocumento.CTRC).Any Then
            BtnCancelarSefaz.Enabled = False
            lnkExcluir.Parent.Visible = False
            Dim CodigoCte As Integer = objNotaFiscal.NotaFiscalOriginal.NotasOrigemDestino.Where(Function(s) s.TipoDeDocumento.Codigo = eTipoDeDocumento.CTRC).FirstOrDefault.Codigo
            MsgBox(Me.Page, "Não é possível excluir a nota fiscal " & objNotaFiscal.Codigo & ", pois está vinculada ao conhecimento de transporte " & CodigoCte & "!")
        ElseIf objNotaFiscal.NotaFiscalOriginal.NotasOrigemDestino.Where(Function(s) s.TipoDeDocumento.Codigo = eTipoDeDocumento.CT_E).Any Then
            BtnCancelarSefaz.Enabled = False
            lnkExcluir.Parent.Visible = False
            Dim CodigoCte As Integer = objNotaFiscal.NotaFiscalOriginal.NotasOrigemDestino.Where(Function(s) s.TipoDeDocumento.Codigo = eTipoDeDocumento.CT_E).FirstOrDefault.Codigo
            MsgBox(Me.Page, "Não é possível excluir a nota fiscal " & objNotaFiscal.Codigo & ", pois está vinculada ao conhecimento de transporte " & CodigoCte & "!")
        ElseIf objNotaFiscal.VencimentosNota.Any(Function(s) s.CodigoSituacao = eSituacao.RemessaBancaria) Then
            BtnCancelarSefaz.Enabled = False
            lnkExcluir.Parent.Visible = False
            MsgBox(Me.Page, "Não é possível excluir a nota fiscal " & objNotaFiscal.Codigo & ", o financeiro está em Cobrança Bancária.")
        ElseIf objNotaFiscal.VencimentosNota.Any(Function(s) s.CodigoSituacao = eSituacao.EndossoTitulo) Then
            BtnCancelarSefaz.Enabled = False
            lnkExcluir.Parent.Visible = False
            MsgBox(Me.Page, "Não é possível excluir a nota fiscal " & objNotaFiscal.Codigo & ", o financeiro está com Endosso.")
        ElseIf objNotaFiscal.Itens(0).CodigoFixacao > 0 AndAlso Not objNotaFiscal.Pedido.Vencimentos.Any(Function(s) s.PedidoFixacao = objNotaFiscal.Itens(0).CodigoFixacao And s.Provisao = eProvisao.Provisao) Then
            BtnCancelarSefaz.Enabled = False
            lnkExcluir.Parent.Visible = False
            MsgBox(Me.Page, "Não é possível excluir a nota fiscal " & objNotaFiscal.Codigo & ", o financeiro da Fixação já foi Liberado.")
        Else
            lnkAtualizar.Parent.Visible = False
            objNotaFiscal.IUD = "D"
            If objNotaFiscal.NossaEmissao And objNotaFiscal.Eletronica Then
                If objNotaFiscal.ChaveNFE.Length = 0 Then
                    MsgBox(Me.Page, "Esta nota não possui a chave de NF-e da SEFAZ!")
                    Exit Sub
                ElseIf objNotaFiscal.ProtocoloNota.Length = 0 Then
                    MsgBox(Me.Page, "Esta nota não possui o número de protocolo da SEFAZ!")
                    Exit Sub
                End If

                Dim fm As New FilesManager()
                If fm.IsConnect() Then
                    lnkExcluir.Parent.Visible = False
                    TabContainer1.ActiveTabIndex = 3
                    txtObservacaoCancelamento.Focus()
                Else
                    MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
                End If
                'Furlan VinculoDeNotaFiscal.aspx qdo estiver funcionando volta ao normal aqui - FURLAN 19/05/2022 DESCOMENTEI ABAIXO, ACOMPANHAR PARA VER SE SURGE PROBLEMA
            ElseIf objNotaFiscal.TemNotaTroca AndAlso Not objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.CONTAEORDEM Then
                If objNotaFiscal.Codigo = objNotaFiscal.NotaTrocaOrigem.Codigo _
                And objNotaFiscal.Serie = objNotaFiscal.NotaTrocaOrigem.Serie _
                And objNotaFiscal.CodigoCliente = objNotaFiscal.NotaTrocaOrigem.CodigoCliente Then
                    MsgBox(Me.Page, "Nota não pode ser excluída, possui vínculo com " & IIf(objNotaFiscal.NotaTrocaDestino.EntradaSaida.ToString().Substring(0, 1) = "E", "ENTRADA ", "SAÍDA ") & " " & objNotaFiscal.NotaTrocaDestino.Codigo.ToString() & "-" & objNotaFiscal.NotaTrocaDestino.Serie & " do dia " & objNotaFiscal.NotaTrocaDestino.Movimento.ToString("dd-MM-yyyy"))
                Else
                    MsgBox(Me.Page, "Nota não pode ser excluída, possui vínculo com " & IIf(objNotaFiscal.NotaTrocaOrigem.EntradaSaida.ToString().Substring(0, 1) = "E", "ENTRADA ", "SAÍDA ") & " " & objNotaFiscal.NotaTrocaOrigem.Codigo.ToString() & "-" & objNotaFiscal.NotaTrocaOrigem.Serie & " do dia " & objNotaFiscal.NotaTrocaOrigem.Movimento.ToString("dd-MM-yyyy"))
                End If
            ElseIf objNotaFiscal.TemNotaTroca _
               AndAlso Not objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.CONTAEORDEM _
               AndAlso objNotaFiscal.Codigo = objNotaFiscal.NotaTrocaOrigem.Codigo _
               AndAlso objNotaFiscal.Serie = objNotaFiscal.NotaTrocaOrigem.Serie _
               AndAlso objNotaFiscal.CodigoCliente = objNotaFiscal.NotaTrocaOrigem.CodigoCliente Then
                MsgBox(Me.Page, "Nota não pode ser excluída, possui vínculo com " & IIf(objNotaFiscal.NotaTrocaDestino.EntradaSaida.ToString().Substring(0, 1) = "E", "ENTRADA ", "SAÍDA ") & " " & objNotaFiscal.NotaTrocaDestino.Codigo.ToString() & "-" & objNotaFiscal.NotaTrocaDestino.Serie & " do dia " & objNotaFiscal.NotaTrocaDestino.Movimento.ToString("dd-MM-yyyy"))
            Else
                'Retorna o valor aos titulos que foram alterados pela emissão da NF de Devolução
                If objNotaFiscal.SubOperacao.Devolucao AndAlso (objNotaFiscal.VencimentosPedido.Any() OrElse objNotaFiscal.VencimentosNota.Any()) Then
                    Dim ReajFinanceiro = New ReajusteFinanceiro()
                    ReajFinanceiro.CancelamentoNotaDeDevolucao(objNotaFiscal)
                End If

                If objNotaFiscal.Salvar Then
                    'colocar no salvar da nota
                    'If objNotaFiscal.Pedido.MomentoFinanceiro = 3 Then
                    '    Dim objPedidoNotaFiscal As New [Lib].Negocio.Pedido(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.CodigoPedido)
                    '    Dim arrSQL As New ArrayList()


                    '    If (objPedidoNotaFiscal IsNot Nothing AndAlso objPedidoNotaFiscal.Vencimentos IsNot Nothing) Then
                    '        Dim objVencimentos As New Hashtable
                    '        Dim i As Integer
                    '        For i = 0 To objPedidoNotaFiscal.Vencimentos.Count - 1
                    '            objVencimentos.Add(i, objPedidoNotaFiscal.Vencimentos(i).Codigo)
                    '        Next
                    '        Session("objPedVencimentos" & HID.Value) = objVencimentos
                    '    End If

                    '    If Not objNotaFiscal.SubOperacao.Devolucao Then
                    '        objPedidoNotaFiscal.Vencimentos.CriarParcelamento(True, CType(Session("objPedVencimentos" & HID.Value), Hashtable))
                    '        objPedidoNotaFiscal.Vencimentos.ModificarHistorico(eTabelas.Pedido, New String() {objPedidoNotaFiscal.Codigo.ToString()})
                    '        arrSQL.AddRange(objPedidoNotaFiscal.Vencimentos.GetSQL(eModoAlteracao.Exclusao))
                    '        arrSQL.AddRange(objPedidoNotaFiscal.Vencimentos.GetSQL(eModoAlteracao.Inclusao, UsuarioServidor.NomeServidor.ToUpper(), objPedidoNotaFiscal.MomentoFinanceiro))

                    '        If Banco.GravaBanco(arrSQL) Then
                    '            LimparCampos("E", objNotaFiscal.Codigo, False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                    '        Else
                    '            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage" & HID.Value))
                    '        End If
                    '    Else
                    '        LimparCampos("E", objNotaFiscal.Codigo, False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                    '    End If
                    'Else
                    '    LimparCampos("E", objNotaFiscal.Codigo, False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                    'End If
                    LimparCampos("E", objNotaFiscal.Codigo, False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage" & HID.Value))
                End If
            End If
        End If
    End Sub

    Public Sub AtualizaNossaEmissaoEletronica()
        SessaoRecuperaNotaFiscal()
        chk_NossaEmissao.Enabled = True
        'chk_nfe.Enabled = True
        txtChaveNFe.ReadOnly = True
        lnkVerificarChaveNFE.Visible = False

        If objNotaFiscal.EntradaSaida = eEntradaSaida.Saida Then

            If LancarSaldoInicial Then
                chk_NossaEmissao.Checked = False
                chk_NossaEmissao.Enabled = False

                objNotaFiscal.NossaEmissao = False

                txtChaveNFe.ReadOnly = False
            ElseIf objNotaFiscal.Empresa.Empresa.NossaEmissao Then
                chk_NossaEmissao.Checked = True
                chk_NossaEmissao.Enabled = False

                objNotaFiscal.NossaEmissao = True

                txtChaveNFe.ReadOnly = False
            Else
                chk_NossaEmissao.Checked = False
                chk_NossaEmissao.Enabled = True

                objNotaFiscal.NossaEmissao = False
            End If

            If LancarSaldoInicial Then
                chk_nfe.Checked = False
                chk_nfe.Enabled = False
                objNotaFiscal.Eletronica = False
            ElseIf objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
                chk_nfe.Checked = True
                chk_nfe.Enabled = False
                objNotaFiscal.Eletronica = True
            ElseIf objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
                chk_nfe.Checked = True
                chk_nfe.Enabled = True
                objNotaFiscal.Eletronica = True
                lnkVerificarChaveNFE.Visible = True
                txtChaveNFe.ReadOnly = False
                chk_NossaEmissao.Enabled = False
            Else
                chk_nfe.Checked = False
                chk_nfe.Enabled = False
                objNotaFiscal.Eletronica = False
            End If

            objNotaFiscal.ChaveNFE = ""
            txtChaveNFe.Text = ""
        ElseIf chk_NossaEmissao.Checked Then
            objNotaFiscal.NossaEmissao = True
            txtNumero.Enabled = False
            txtSerie.Enabled = False

            If objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
                chk_nfe.Checked = True
                chk_nfe.Enabled = False
                objNotaFiscal.Eletronica = True
            Else
                chk_nfe.Checked = False
                chk_nfe.Enabled = False
                objNotaFiscal.Eletronica = False
            End If

            objNotaFiscal.ChaveNFE = ""
            txtChaveNFe.Text = ""
        ElseIf chk_NossaEmissao.Checked = False And chk_nfe.Checked Then
            txtNumero.Enabled = True
            txtSerie.Enabled = True
            objNotaFiscal.NossaEmissao = False
            objNotaFiscal.Eletronica = True
            objNotaFiscal.ChaveNFE = txtChaveNFe.Text
            txtChaveNFe.ReadOnly = False
            lnkVerificarChaveNFE.Visible = True

            'Importação do xml de notas de entrada
            chk_nfe.Enabled = SessaoDsXML() Is Nothing
            chk_NossaEmissao.Enabled = SessaoDsXML() Is Nothing

        ElseIf chk_NossaEmissao.Checked = False And chk_nfe.Checked = False Then
            txtNumero.Enabled = True
            txtSerie.Enabled = True
            objNotaFiscal.NossaEmissao = False
            objNotaFiscal.Eletronica = False
            objNotaFiscal.ChaveNFE = ""
            txtChaveNFe.Text = ""
        End If

        'chkEspelho.Checked = objNotaFiscal.NossaEmissao

        If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
            If Not VerificarModo(objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa, objNotaFiscal.NossaEmissao, True) Then
                lnkNovo.Parent.Visible = False
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
                btnInutilizar.Enabled = False
                lnkConsultar.Parent.Visible = False
                lnkEspelho.Parent.Visible = False
            End If
        End If

        If LancarSaldoInicial Then
            Dim numeroNF As String = Left((New Random).Next, 6)
            objNotaFiscal.Codigo = CInt(numeroNF)

            objNotaFiscal.Serie = "F"

            Dim dataNota As String = "31/12/" & Today.AddYears(-1).ToString("yyyy")
            objNotaFiscal.DataNota = CDate(dataNota)
            objNotaFiscal.Movimento = objNotaFiscal.DataNota

            txtDataDeEmissao.Text = objNotaFiscal.DataNota.ToString("dd/MM/yyyy")
            txtDataDeEntrada.Text = objNotaFiscal.Movimento.ToString("dd/MM/yyyy")
        End If

        txtNumero.Text = objNotaFiscal.Codigo
        txtSerie.Text = objNotaFiscal.Serie

        SessaoSalvaNotaFiscal()
    End Sub

    Public Sub AtualizaFormulario()
        If FinanceiroNovo And objNotaFiscal.SubOperacao.Financeiro Then
            ucFinanceiro.CarregarResumo()
            ucFinanceiro.AtualizarValorNotaOuFixacaoOuTroca()
        End If

        If objNotaFiscal.CriarRomaneio OrElse objNotaFiscal.CodigoRomaneio > 0 Then
            If objNotaFiscal.EntradaSaida = eEntradaSaida.Saida Then
                objNotaFiscal.PesoBruto = objNotaFiscal.Romaneio.PesoBruto
                objNotaFiscal.PesoLiquido = objNotaFiscal.Romaneio.PesoLiquido
            End If
            txtPesoBruto.Text = objNotaFiscal.PesoBruto.ToString("N4")
            txtPesoLiquido.Text = objNotaFiscal.PesoLiquido.ToString("N4")
        Else
            txtPesoBruto.Text = objNotaFiscal.PesoBruto.ToString("N4")
            txtPesoLiquido.Text = objNotaFiscal.PesoLiquido.ToString("N4")
        End If

        'Acrescentei para resolver o problema da Quantidade
        objNotaFiscal.Quantidade = objNotaFiscal.PesoLiquido
        txtVolumes.Text = objNotaFiscal.Quantidade.ToString("N4")

        txtEspecie.Text = objNotaFiscal.Especie
        txtMarca.Text = objNotaFiscal.Marca

        If objNotaFiscal.Numero = "1" Then objNotaFiscal.Numero = ""

        txtNumeracao.Text = objNotaFiscal.Numero
        txtRomaneio.Text = objNotaFiscal.CodigoRomaneio

        txtSeguro.Text = objNotaFiscal.ValorSeguro.ToString("N2")
        txtOutras.Text = objNotaFiscal.ValorAduaneira.ToString("N2")
        txtValorFrete.Text = objNotaFiscal.ValorFrete.ToString("N2")
        txtDesconto.Text = objNotaFiscal.ValorDesconto.ToString("N2")
        txtValorTotalDosProdutos.Text = objNotaFiscal.TotalProduto.ToString("N2")
        txtValorTotalDaNota.Text = objNotaFiscal.TotalNota.ToString("N2")
        txtBaseIcmsNota.Text = objNotaFiscal.ValorBaseIcms.ToString("N2")
        txtValorIcmsNota.Text = objNotaFiscal.ValorIcms.ToString("N2")
        txtValorIPINota.Text = objNotaFiscal.ValorIPI.ToString("N2")
        txtValorBaseIcmsSTNota.Text = objNotaFiscal.ValorBaseIcmsST.ToString("N2")
        txtValorIcmsSTNota.Text = objNotaFiscal.ValorIcmsST.ToString("N2")
    End Sub

    Public Sub AtualizaComEncargos()
        AtualizaFormularioComAClasse()

        SessaoRecuperaNotaFiscal()

        ''Furlan testando
        'If Not objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.IUD = "U" Then
        '    Dim y As Integer = 0
        '    While y < gridItens.Rows.Count
        '        ItemNotaOK(y)

        '        y += 1
        '    End While
        'End If

        If Not objNotaFiscal.NossaEmissao AndAlso
            objNotaFiscal.IUD = "U" AndAlso
            Not objNotaFiscal.NotaFiscalOriginal Is Nothing AndAlso
            Not objNotaFiscal.TotalNota = objNotaFiscal.NotaFiscalOriginal.TotalNota Then
            'Financeiro
            If Not objNotaFiscal.VencimentosPedido Is Nothing AndAlso objNotaFiscal.VencimentosPedido.Count > 0 Then
                'gridVencimentosPedido.DataSource = objNotaFiscal.VencimentosPedido.ToArray
                'Ajuste para Selecinar o Título e carregar no Financeiro - Furlan 30/07/2025
                gridVencimentosPedido.DataSource = From tit In objNotaFiscal.VencimentosPedido
                                                   Select New With {.Codigo = tit.Codigo, .ReceberPagar = tit.ReceberPagar, .DescricaoMoeda = tit.DescricaoMoeda, .DescricaoProvisao = tit.DescricaoProvisao, .Prorrogacao = tit.Prorrogacao, .ValorDoDocumento = tit.ValorDoDocumento, .MoedaValorDoDocumento = tit.MoedaValorDoDocumento, .CodigoCifrado = IIf(tit.ReceberPagar = "P", Funcoes.Cifrar("ContasAPagar-" & tit.Codigo), Funcoes.Cifrar("ContasAReceber-" & tit.Codigo))}
                gridVencimentosPedido.DataBind()

                If objNotaFiscal.VencimentosNota Is Nothing OrElse objNotaFiscal.VencimentosNota.Count = 0 Then
                    gridVencimentosNota.DataSource = Nothing
                    gridVencimentosNota.DataBind()
                End If
            End If

            If objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                If Not objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada AndAlso objNotaFiscal.TotalNota <> objNotaFiscal.VencimentosNota.Sum(Function(x) x.ValorDoDocumento) Then
                    objNotaFiscal.VencimentosNota.ReajFinanceiro = New ReajusteFinanceiro(objNotaFiscal, False)
                    objNotaFiscal.VencimentosNota.ReajFinanceiro.ReajustaNotaDeEntrada()
                    'gridVencimentosNota.DataSource = objNotaFiscal.VencimentosNota.ToArray
                    'Ajuste para Selecinar o Título e carregar no Financeiro - Furlan 30/07/2025
                    gridVencimentosNota.DataSource = From tit In objNotaFiscal.VencimentosNota
                                                     Select New With {.Codigo = tit.Codigo, .ReceberPagar = tit.ReceberPagar, .DescricaoMoeda = tit.DescricaoMoeda, .DescricaoProvisao = tit.DescricaoProvisao, .Prorrogacao = tit.Prorrogacao, .ValorDoDocumento = tit.ValorDoDocumento, .MoedaValorDoDocumento = tit.MoedaValorDoDocumento, .CodigoCifrado = IIf(tit.ReceberPagar = "P", Funcoes.Cifrar("ContasAPagar-" & tit.Codigo), Funcoes.Cifrar("ContasAReceber-" & tit.Codigo))}
                    gridVencimentosNota.DataBind()

                    'gridVencimentosPedido.DataSource = objNotaFiscal.VencimentosPedido.ToArray
                    'Ajuste para Selecinar o Título e carregar no Financeiro - Furlan 30/07/2025
                    gridVencimentosPedido.DataSource = From tit In objNotaFiscal.VencimentosPedido
                                                       Select New With {.Codigo = tit.Codigo, .ReceberPagar = tit.ReceberPagar, .DescricaoMoeda = tit.DescricaoMoeda, .DescricaoProvisao = tit.DescricaoProvisao, .Prorrogacao = tit.Prorrogacao, .ValorDoDocumento = tit.ValorDoDocumento, .MoedaValorDoDocumento = tit.MoedaValorDoDocumento, .CodigoCifrado = IIf(tit.ReceberPagar = "P", Funcoes.Cifrar("ContasAPagar-" & tit.Codigo), Funcoes.Cifrar("ContasAReceber-" & tit.Codigo))}
                    gridVencimentosPedido.DataBind()
                End If
            Else
                objNotaFiscal.VencimentosNota.Clear()
                objNotaFiscal.VencimentosPedido = Nothing
                gridVencimentosNota.DataSource = Nothing
                gridVencimentosNota.DataBind()
            End If

            SessaoSalvaNotaFiscal()
        End If
    End Sub

    Public Sub AtualizaFormularioComAClasse()
        SessaoRecuperaNotaFiscal()

        Try
            If objNotaFiscal Is Nothing Then Exit Sub

            ConsutaNotasReferenciais(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.CodigoCliente, objNotaFiscal.EnderecoCliente, objNotaFiscal.EntradaSaida.ToString.Substring(0, 1), objNotaFiscal.Serie, objNotaFiscal.Codigo)

            If objNotaFiscal.IUD = "I" Then
                lnkNovo.Parent.Visible = True
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False

                'Importação do xml de notas de entrada
                If SessaoDsXML IsNot Nothing Then
                    chk_nfe.Checked = objNotaFiscal.Eletronica
                    'chk_nfe.Enabled = False
                    txtChaveNFe.ReadOnly = True
                    txtNumero.Enabled = True
                    txtSerie.Enabled = True
                End If
            Else
                Session("Carregando" & HID.Value) = True
                lnkNovo.Parent.Visible = False
                btnInutilizar.Enabled = False

                'If objNotaFiscal.CodigoSituacao = eSituacao.Normal Then
                '    lnkExcluir.Parent.Visible = True
                '    lnkAtualizar.Parent.Visible = True
                'End If

                chk_nfe.Checked = objNotaFiscal.Eletronica
                chk_NossaEmissao.Checked = objNotaFiscal.NossaEmissao
                chk_nfe.Enabled = False
                chk_NossaEmissao.Enabled = False
                imgUsuario.Visible = True
                lblUsuario.Visible = True
                If objNotaFiscal.UsuarioAlteracao.Length > 0 Then
                    lblUsuario.Text = objNotaFiscal.UsuarioInclusao & "/" & objNotaFiscal.UsuarioAlteracao
                Else
                    lblUsuario.Text = objNotaFiscal.UsuarioInclusao
                End If
                btnEmpresa.Enabled = False
                btnCliente.Enabled = False
                btnPedido.Enabled = False

                If objNotaFiscal.NossaEmissao Then DesabilitarBotoes()
            End If

            txtNomeDaEmpresa.Text = objNotaFiscal.Empresa.Nome
            txtInscricaoDaEmpresa.Text = objNotaFiscal.Empresa.InscricaoEstadual
            txtCnpjDaEmpresa.Text = objNotaFiscal.Empresa.CodigoFormatado

            txtNomeDoCliente.Text = objNotaFiscal.Cliente.Nome
            txtInscricaoDoCliente.Text = objNotaFiscal.Cliente.InscricaoEstadual
            txtCnpjDoCliente.Text = objNotaFiscal.Cliente.CodigoFormatado
            txtEnderecoDoCliente.Text = objNotaFiscal.Cliente.Endereco
            txtComplementoDoCliente.Text = objNotaFiscal.Cliente.Complemento
            txtBairroDoCliente.Text = objNotaFiscal.Cliente.Bairro
            txtCepDoCliente.Text = objNotaFiscal.Cliente.CEP
            txtCidadeDoCliente.Text = objNotaFiscal.Cliente.Cidade
            txtTelefoneDoCliente.Text = objNotaFiscal.Cliente.Telefone
            txtEstadoDoCliente.Text = objNotaFiscal.Cliente.CodigoEstado
            txtInscricaoDoCliente.Text = objNotaFiscal.Cliente.InscricaoEstadual

            If objNotaFiscal.DataHoraNFE IsNot Nothing AndAlso objNotaFiscal.DataHoraNFE.HasValue Then
                txtHoraDaSaida.Text = CDate(objNotaFiscal.DataHoraNFE).TimeOfDay.ToString()
            End If

            ddlTipoDeDocumento.SelectedValue = objNotaFiscal.CodigoTipoDeDocumento

            txtRomaneio.Text = objNotaFiscal.CodigoRomaneio

            If objNotaFiscal.IUD = "U" AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.CodigoRomaneio = 0 Then
                btnClassificacao.Enabled = True
            ElseIf objNotaFiscal.CodigoRomaneio > 0 Then
                txtPesoRomaneio.Text = objNotaFiscal.Romaneio.PesoLiquido
                If objNotaFiscal.IUD = "U" AndAlso (objNotaFiscal.Romaneio.Pesagens Is Nothing OrElse objNotaFiscal.Romaneio.Pesagens.Count = 0) Then
                    btnRomaneio.Visible = False
                    imgRomaneio.Visible = True
                End If
            End If

            txtDataDeEmissao.Text = objNotaFiscal.DataNota.ToString("dd/MM/yyyy")
            txtDataDeEntrada.Text = objNotaFiscal.Movimento.ToString("dd/MM/yyyy")

            If objNotaFiscal.CodigoPedido > 0 AndAlso objNotaFiscal.Pedido.Itens.Count > 0 AndAlso objNotaFiscal.Pedido.Itens(0).Fixacoes.Count > 0 AndAlso objNotaFiscal.Pedido.Itens(0).Fixacoes(0).IndiceFixado > 0 Then
                txtPedido.Text = objNotaFiscal.CodigoPedido & " - Indice Fixado " & objNotaFiscal.Pedido.Itens(0).Fixacoes(0).IndiceFixado
            ElseIf objNotaFiscal.CodigoPedido > 0 AndAlso objNotaFiscal.Pedido.IndiceFixado > 0 Then
                txtPedido.Text = objNotaFiscal.CodigoPedido & " - Indice Fixado " & objNotaFiscal.Pedido.IndiceFixado
            Else
                txtPedido.Text = objNotaFiscal.CodigoPedido
            End If

            If objNotaFiscal.CodigoPedido > 0 Then imgExtratoPedido.Visible = True

            BuscarSubOperacoes(objNotaFiscal.CodigoOperacao, objNotaFiscal.CodigoSubOperacao)
            SelecionarIndiceCombo(cmbFinalidade, objNotaFiscal.CodigoFinalidade)

            txtAutorizacao.Text = objNotaFiscal.CodigoAutorizacao
            txtCessaoDeCredito.Text = objNotaFiscal.CodigoProcuracao

            If objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                If objNotaFiscal.CodigoFinalidade = 13 Or objNotaFiscal.CodigoFinalidade = 23 Then
                    ChkTroca.Visible = False
                    BtnNotas.Visible = True
                Else
                    ChkTroca.Visible = objNotaFiscal.Itens(0).Produto.Agrupar = "N"
                    BtnNotas.Visible = False
                End If
            Else
                ChkTroca.Visible = False
                BtnNotas.Visible = objNotaFiscal.Itens(0).Produto.Agrupar = "N"
            End If
            ChkTroca.Checked = objNotaFiscal.Troca

            CarregarDepositosPedido()

            Dim dtItens As New DataTable("Itens")

            dtItens.Columns.Add("Produto", Type.GetType("System.String"))
            dtItens.Columns.Add("NomeProduto", Type.GetType("System.String"))
            dtItens.Columns.Add("Sequencia", Type.GetType("System.String"))
            dtItens.Columns.Add("BaseCalculo", Type.GetType("System.String"))
            dtItens.Columns.Add("Lote", Type.GetType("System.String"))
            dtItens.Columns.Add("Classificacao", Type.GetType("System.String"))
            dtItens.Columns.Add("Embalagem", Type.GetType("System.String"))
            dtItens.Columns.Add("Saldo", Type.GetType("System.Decimal"))
            dtItens.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
            dtItens.Columns.Add("QuantidadeFisica", Type.GetType("System.Decimal"))
            dtItens.Columns.Add("Unitario", Type.GetType("System.Decimal"))
            dtItens.Columns.Add("Total", Type.GetType("System.Decimal"))

            For Each row As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
                Dim drItem As DataRow = dtItens.NewRow()
                If String.IsNullOrWhiteSpace(objNotaFiscal.Especie) Then objNotaFiscal.Especie = row.Produto.Embalagem.Descricao
                row.CodigoDeposito = objNotaFiscal.CodigoDeposito
                row.EnderecoDeposito = objNotaFiscal.EnderecoDeposito
                drItem("Produto") = row.CodigoProduto
                drItem("Sequencia") = row.Sequencia
                drItem("Lote") = row.Lote
                drItem("Classificacao") = row.Classificacao
                drItem("Embalagem") = row.CodigoEmbalagemIndea + "-" + row.CodigoTipoDeEmbalagem + "-" + row.CapacidadeEmbalagem.ToString
                drItem("BaseCalculo") = 1
                drItem("Unitario") = "0"

                drItem("Saldo") = IIf(row.QuantidadeFiscal = 0, row.SaldoValorOficial.ToString("N4"), row.QuantidadeFiscal.ToString("N4"))
                'If Not objNotaFiscal.SubOperacao.Devolucao Then
                drItem("Total") = row.ValorTotal.ToString("N2")
                drItem("Unitario") = row.Unitario.ToString("N10")
                'End If

                If (row.Produto.ControlarLote Or row.Produto.ControlarEmbalagem) And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL Then
                    drItem("NomeProduto") = "$" & row.Produto.Nome
                    If objNotaFiscal.IUD <> "U" Then
                        row.QuantidadeFiscal = IIf(row.Lote.Length > 0 Or row.CodigoEmbalagemIndea.Length > 0, row.QuantidadeFiscal, 0)
                    End If
                Else

                    If objNotaFiscal.Empresa.Empresa.UsarRegistroMinAgr Then
                        drItem("NomeProduto") = row.Produto.Nome & "-" & row.Produto.Descricao & "(" & row.Produto.RegistroMinisterioAgricultura & ")"
                    ElseIf objNotaFiscal.Empresa.Empresa.UsarDescricaoProduto Then
                        drItem("NomeProduto") = row.Produto.Nome & "-" & row.Produto.Descricao
                    Else
                        drItem("NomeProduto") = row.Produto.Nome
                    End If
                End If

                drItem("Quantidade") = row.QuantidadeFiscal.ToString("N4")
                drItem("QuantidadeFisica") = row.QuantidadeFisica.ToString("N4")
                dtItens.Rows.Add(drItem)
            Next

            If objNotaFiscal.Cliente.CodigoEstado = "EX" Then
                If objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada AndAlso Not objNotaFiscal.SubOperacao.Devolucao _
                    OrElse objNotaFiscal.EntradaSaida = eEntradaSaida.Saida AndAlso objNotaFiscal.SubOperacao.Devolucao Then
                    'Importação
                    TabImportacao.Visible = (objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada AndAlso objNotaFiscal.Cliente.CodigoEstado = "EX")
                Else
                    'Exportação
                    TabExportacao.Visible = (objNotaFiscal.EntradaSaida = eEntradaSaida.Saida AndAlso objNotaFiscal.Cliente.CodigoEstado = "EX")
                End If
            End If

            txtPesoBruto.Text = objNotaFiscal.PesoBruto.ToString("N4")
            txtPesoLiquido.Text = objNotaFiscal.PesoLiquido.ToString("N4")

            txtVolumes.Text = objNotaFiscal.Quantidade.ToString("N4")

            txtMarca.Text = objNotaFiscal.Marca

            If objNotaFiscal.Numero = "1" Then objNotaFiscal.Numero = ""
            txtNumeracao.Text = objNotaFiscal.Numero

            txtEspecie.Text = objNotaFiscal.Especie

            txtNaturezaDaOperacao.Text = objNotaFiscal.NaturezaDaOperacao

            gridItens.DataSource = dtItens
            gridItens.DataBind()

            Dim i As Integer = 0
            While i < gridItens.Rows.Count
                If Mid(gridItens.Rows(i).Cells(2).Text, 1, 1) = "$" And Not objNotaFiscal.Itens(i).SubOperacao.Devolucao And objNotaFiscal.Itens(i).SubOperacao.Classe <> eClassesOperacoes.GLOBAL Then
                    CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled = False
                End If

                If objNotaFiscal.TemNotaTroca Then
                    If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.CONTAEORDEM Then
                        CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled = False
                        If objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada Then
                            CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled = True
                        Else
                            CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled = False
                        End If
                    ElseIf objNotaFiscal.NotaDeTroca.Itens(0).QuantidadeFisica <= objNotaFiscal.Itens(0).SaldoPedidoFisico Then
                        CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled = True
                    Else
                        CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled = False
                    End If
                End If

                If objNotaFiscal.SubOperacao.Devolucao Then CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled = False

                gridItens.Rows(i).Cells(2).Text = gridItens.Rows(i).Cells(2).Text.Replace("$", "")

                If objNotaFiscal.IUD = "U" And objNotaFiscal.NossaEmissao Then CType(gridItens.Rows(i).FindControl("btnItem"), Button).Enabled = False

                If SessaoDsXML IsNot Nothing OrElse objNotaFiscal.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE ICMS") OrElse objNotaFiscal.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE IPI") Then
                    CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled = False
                    CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled = False
                    CType(gridItens.Rows(i).FindControl("txtTotalItem"), TextBox).Enabled = False
                End If

                'Travar unitário cfe. solicitação e-mail dia 18/11/2021 Jonathan - Furlan 19/11/2021
                'Or Left(objNotaFiscal.Empresa.Codigo, 8) = "44979506" _

                If objNotaFiscal.NossaEmissao AndAlso (Left(objNotaFiscal.Empresa.Codigo, 8) = "05366261" _
                                                       Or Left(objNotaFiscal.Empresa.Codigo, 8) = "38198213" _
                                                       Or Left(objNotaFiscal.Empresa.Codigo, 8) = "62747840" _
                                                       Or Left(objNotaFiscal.Empresa.Codigo, 8) = "62780383" _
                                                       Or Left(objNotaFiscal.Empresa.Codigo, 8) = "63358210" _
                                                       Or Left(objNotaFiscal.Empresa.Codigo, 8) = "40938762" _
                                                       Or Left(objNotaFiscal.Empresa.Codigo, 8) = "44005444" _
                                                       Or Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" _
                                                       Or Left(objNotaFiscal.Empresa.Codigo, 8) = "48984539") Then
                    CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled = False
                End If

                i += 1
            End While

            'Kitio
            If FinanceiroNovo And objNotaFiscal.SubOperacao.Financeiro Then
                If objNotaFiscal.IUD = "I" Then
                    ucFinanceiro.CarregarResumo()
                    ucFinanceiro.AtualizarValorNotaOuFixacaoOuTroca()
                Else
                    ucFinanceiro.CarregarControle()
                    'ucFinanceiro.CarregarResumo()
                End If
            End If

            txtSeguro.Text = objNotaFiscal.ValorSeguro.ToString("N2")
            txtOutras.Text = objNotaFiscal.ValorAduaneira.ToString("N2")
            txtValorFrete.Text = objNotaFiscal.ValorFrete.ToString("N2")
            txtDesconto.Text = objNotaFiscal.ValorDesconto.ToString("N2")
            txtValorTotalDosProdutos.Text = objNotaFiscal.TotalProduto.ToString("N2")
            txtValorTotalDaNota.Text = objNotaFiscal.TotalNota.ToString("N2")
            txtBaseIcmsNota.Text = objNotaFiscal.ValorBaseIcms.ToString("N2")
            txtValorIcmsNota.Text = objNotaFiscal.ValorIcms.ToString("N2")
            txtValorIPINota.Text = objNotaFiscal.ValorIPI.ToString("N2")
            txtValorBaseIcmsSTNota.Text = objNotaFiscal.ValorBaseIcmsST.ToString("N2")
            txtValorIcmsSTNota.Text = objNotaFiscal.ValorIcmsST.ToString("N2")
            txtES.Text = objNotaFiscal.EntradaSaida.ToString.Substring(0, 1)
            txtNumero.Text = objNotaFiscal.Codigo
            txtSerie.Text = objNotaFiscal.Serie
            txtChaveNFe.Text = objNotaFiscal.ChaveNFE

            txtSituacao.Text = objNotaFiscal.Situacao.Descricao

            If objNotaFiscal.CodigoTransportador.Length > 0 Then
                txtNomeDoTransportador.Text = objNotaFiscal.Transportador.Nome
                txtCidadeDoTransportador.Text = objNotaFiscal.Transportador.Cidade
                txtEstadoDoTransportador.Text = objNotaFiscal.Transportador.CodigoEstado
                txtCnpjDoTransportador.Text = objNotaFiscal.Transportador.Codigo
                txtInscricaoDoTransportador.Text = objNotaFiscal.Transportador.InscricaoEstadual
                txtEnderecoDoTransportador.Text = objNotaFiscal.Transportador.Endereco & ", " & objNotaFiscal.Transportador.Numero
            End If

            If objNotaFiscal.PlacaTransportador.Length > 0 Then
                txtPlacas.Text = objNotaFiscal.PlacaDetalhes.Placa01 & " " & objNotaFiscal.PlacaDetalhes.Placa02 & " " & objNotaFiscal.PlacaDetalhes.Placa03
                txtEstadoDaPlaca.Text = objNotaFiscal.PlacaDetalhes.EstadoPlaca01 & " " & objNotaFiscal.PlacaDetalhes.EstadoPlaca02 & " " & objNotaFiscal.PlacaDetalhes.EstadoPlaca03
            Else
                txtPlacas.Text = ""
                txtCodigoPlaca.Value = ""
                txtEstadoDaPlaca.Text = ""
            End If

            txtObservacoesFiscais.Text = objNotaFiscal.Observacoes

            'txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque
            If String.IsNullOrWhiteSpace(observacaoAutorizacao.Value) Then
                txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
            Else
                txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
            End If

            txtObservacoesInternas.Text = objNotaFiscal.ObservacoesControleInterno

            '***********************************************************
            '***************  Contabilização  *****************************
            '***********************************************************
            objNotaFiscal.LancamentosContabeis.CalcularSaldo()
            gridContabilizacao.DataSource = objNotaFiscal.LancamentosContabeis
            gridContabilizacao.DataBind()
            '--------------------------------------------------------------'

            '***********************************************************
            '***************  Vencimentos  *****************************
            '***********************************************************

            If Not objNotaFiscal.VencimentosPedido Is Nothing Then
                'gridVencimentosPedido.DataSource = objNotaFiscal.VencimentosPedido.ToArray
                'Ajuste para Selecinar o Título e carregar no Financeiro - Furlan 30/07/2025
                gridVencimentosPedido.DataSource = From tit In objNotaFiscal.VencimentosPedido
                                                   Select New With {.Codigo = tit.Codigo, .ReceberPagar = tit.ReceberPagar, .DescricaoMoeda = tit.DescricaoMoeda, .DescricaoProvisao = tit.DescricaoProvisao, .Prorrogacao = tit.Prorrogacao, .ValorDoDocumento = tit.ValorDoDocumento, .MoedaValorDoDocumento = tit.MoedaValorDoDocumento, .CodigoCifrado = IIf(tit.ReceberPagar = "P", Funcoes.Cifrar("ContasAPagar-" & tit.Codigo), Funcoes.Cifrar("ContasAReceber-" & tit.Codigo))}
            Else
                gridVencimentosPedido.DataSource = Nothing
            End If
            gridVencimentosPedido.DataBind()

            If Not objNotaFiscal.VencimentosNota Is Nothing Then
                'gridVencimentosNota.DataSource = objNotaFiscal.VencimentosNota.ToArray
                'Ajuste para Selecinar o Título e carregar no Financeiro - Furlan 30/07/2025
                gridVencimentosNota.DataSource = From tit In objNotaFiscal.VencimentosNota
                                                 Select New With {.Codigo = tit.Codigo, .ReceberPagar = tit.ReceberPagar, .DescricaoMoeda = tit.DescricaoMoeda, .DescricaoProvisao = tit.DescricaoProvisao, .Prorrogacao = tit.Prorrogacao, .ValorDoDocumento = tit.ValorDoDocumento, .MoedaValorDoDocumento = tit.MoedaValorDoDocumento, .CodigoCifrado = IIf(tit.ReceberPagar = "P", Funcoes.Cifrar("ContasAPagar-" & tit.Codigo), Funcoes.Cifrar("ContasAReceber-" & tit.Codigo))}
            Else
                gridVencimentosNota.DataSource = Nothing
            End If
            gridVencimentosNota.DataBind()

            '*******************************************************************
            '***************  Dados da Exportacao  *****************************
            '*******************************************************************
            If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES Then
                pnlExportacao.Enabled = True
            Else
                pnlExportacao.Enabled = False
            End If

            If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR OrElse
                objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.COMPLEMENTACOES Then
                btnProcuracao.Enabled = False
            End If

            GridConhecimento.DataSource = objNotaFiscal.DadosDaExportacaoCE.ToArray
            GridConhecimento.DataBind()

            gridRE.DataSource = objNotaFiscal.DadosDaExportacaoRE.ToArray
            gridRE.DataBind()

            If Not objNotaFiscal.DadosDaExportacao Is Nothing Then
                If objNotaFiscal.DadosDaExportacao.NrDespachoExp.Length > 0 Then
                    txtNrDespacho.Text = objNotaFiscal.DadosDaExportacao.NrDespachoExp
                    txtDataDespacho.Text = objNotaFiscal.DadosDaExportacao.DataDespachoExp
                End If

                If objNotaFiscal.DadosDaExportacao.FaturaExportacao.Length > 0 Then
                    txtFaturaExportacao.Text = objNotaFiscal.DadosDaExportacao.FaturaExportacao
                End If

                If objNotaFiscal.DadosDaExportacao.CodigoPaisDestino = 0 Then
                    ddlPaisDestino.SelectedIndex = 0
                Else
                    ddlPaisDestino.SelectedValue = objNotaFiscal.DadosDaExportacao.CodigoPaisDestino
                End If

                txtNavio.Text = objNotaFiscal.DadosDaExportacao.Navio

                If objNotaFiscal.DadosDaExportacao.DataAverbacao IsNot Nothing Then
                    txtDataAverba.Text = CDate(objNotaFiscal.DadosDaExportacao.DataAverbacao).ToString("dd/MM/yyyy")
                Else
                    txtDataAverba.Text = ""
                End If

                If objNotaFiscal.DadosDaExportacao.NumAtoConcessorio.Length > 0 Then
                    TxtNumAtoConcessorio.Text = objNotaFiscal.DadosDaExportacao.NumAtoConcessorio
                    TxtDtaRegAtoConcessorio.Text = objNotaFiscal.DadosDaExportacao.DtaRegAtoConcessorio
                    TxtDtaValidAtoConcessorio.Text = objNotaFiscal.DadosDaExportacao.DtaValidAtoConcessorio
                End If

                objListaDeNotasFiscaisReferenciais = New ListNotaFiscalReferencial(objNotaFiscal, eTipoReferencial.EXP)
                If objListaDeNotasFiscaisReferenciais.Count > 0 Then
                    objNotaFiscal.NotasReferenciais = objListaDeNotasFiscaisReferenciais
                    grdNotasReferenciais.DataSource = objNotaFiscal.NotasReferenciais
                    grdNotasReferenciais.DataBind()
                Else
                    grdNotasReferenciais.DataSource = Nothing
                    grdNotasReferenciais.DataBind()
                End If
            End If

            If Not objNotaFiscal.DadosDaImportacao Is Nothing Then
                If objNotaFiscal.DadosDaImportacao.NumeroDeclaracaoImportacao.Length > 0 Then
                    txtDI.Text = objNotaFiscal.DadosDaImportacao.NumeroDeclaracaoImportacao
                    txtDataDI.Text = objNotaFiscal.DadosDaImportacao.DataDeclaracaoImportacao
                End If

                If objNotaFiscal.DadosDaImportacao.RegistroDeImportacao.Length > 0 Then
                    TxtRegImportacao.Text = objNotaFiscal.DadosDaImportacao.RegistroDeImportacao
                    TxtDataRegImp.Text = objNotaFiscal.DadosDaImportacao.DataRegistroDeImportacao
                End If

                If objNotaFiscal.DadosDaImportacao.LocalEmbarqueImportacao.Length > 0 Then
                    txtEmbarqueDI.Text = objNotaFiscal.DadosDaImportacao.LocalEmbarqueImportacao
                    txtDataEmbarqueDI.Text = objNotaFiscal.DadosDaImportacao.DataEmbarqueImportacao
                    txtCodigoUFEmbarqueDI.Value = objNotaFiscal.DadosDaImportacao.EstadoEmbarqueImportacao
                    lblUFEmbarqueDI.Text = objNotaFiscal.DadosDaImportacao.EstadoEmbarqueImportacao
                End If

                If objNotaFiscal.DadosDaImportacao.LocalDesembarqueImportacao.Length > 0 Then
                    txtDesembarqueDI.Text = objNotaFiscal.DadosDaImportacao.LocalDesembarqueImportacao
                    txtDataDesembarqueDI.Text = objNotaFiscal.DadosDaImportacao.DataDesembarqueImportacao
                    txtCodigoUFDesembarqueDI.Value = objNotaFiscal.DadosDaImportacao.EstadoDesembarqueImportacao
                    lblUFDesembarqueDI.Text = objNotaFiscal.DadosDaImportacao.EstadoDesembarqueImportacao
                End If

                If objNotaFiscal.DadosDaImportacao.CodigoFabricante.Length > 0 Then
                    lblFabricanteDI.Text = objNotaFiscal.DadosDaImportacao.Fabricante.Nome & " - " & objNotaFiscal.DadosDaImportacao.Fabricante.Cidade & "/" & objNotaFiscal.DadosDaImportacao.Fabricante.CodigoEstado
                    txtCodigolblFabricanteDI.Value = objNotaFiscal.DadosDaImportacao.Fabricante.Codigo & "-" & objNotaFiscal.DadosDaImportacao.Fabricante.CodigoEndereco
                End If

                If objNotaFiscal.DadosDaImportacao.NumAtoConcessorio.Length > 0 Then
                    TxtNumAtoConcessorioImp.Text = objNotaFiscal.DadosDaImportacao.NumAtoConcessorio
                    TxtDtaRegAtoConcessorioImp.Text = objNotaFiscal.DadosDaImportacao.DtaRegAtoConcessorio
                    TxtDtaValidAtoConcessorioImp.Text = objNotaFiscal.DadosDaImportacao.DtaValidAtoConcessorio
                End If

                If objNotaFiscal.DadosDaImportacao.NrFatura.Length > 0 Then
                    TxtNrFaturaImp.Text = objNotaFiscal.DadosDaImportacao.NrFatura
                End If

                If objNotaFiscal.DadosDaImportacao.ViaDeTransporte > 0 Then
                    ddlViaDeTransporteDI.SelectedValue = objNotaFiscal.DadosDaImportacao.ViaDeTransporte
                End If

                If objNotaFiscal.DadosDaImportacao.TipoDeImportacao > 0 Then
                    ddlTipoDeImportacaoDI.SelectedValue = objNotaFiscal.DadosDaImportacao.TipoDeImportacao
                End If

                txtValorVAFRMMDI.Text = objNotaFiscal.DadosDaImportacao.ValorVAFRMM
            End If

            SessaoSalvaNotaFiscal()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message.ToString(), eTitulo.Erro)
        End Try

    End Sub

    Private Sub DesabilitarBotoes()
        btnRomaneio.Enabled = False
        btnClassificacao.Enabled = False
        BtnRetirada.Enabled = False
        BtnNotas.Enabled = False
        cmbFinalidade.Enabled = False
        btnTransportador.Enabled = False
        btnPlaca.Enabled = False
        btnObservacoesFiscais.Enabled = False
        BtnEmbalagem.Enabled = False
        imgConfirmarDI.Enabled = False
    End Sub

    Private Sub HabilitarBotoes()
        btnCliente.Enabled = True
        btnPedido.Enabled = True
        btnRomaneio.Enabled = True
        btnClassificacao.Enabled = True
        BtnRetirada.Enabled = True
        BtnNotas.Enabled = True
        btnDeposito.Enabled = True
        BtnTransbordo.Enabled = True
        BtnOrigemDestino.Enabled = True
        BtnEntrega.Enabled = True
        cmbFinalidade.Enabled = True
        btnTransportador.Enabled = True
        btnPlaca.Enabled = True
        btnObservacoesFiscais.Enabled = True
        BtnEmbalagem.Enabled = True
        imgConfirmarDI.Enabled = True
    End Sub

    Private Function TemRomaneio(ByVal Pedido As String) As Boolean
        SessaoRecuperaNotaFiscal()
        If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL Then
            Session("MsgControle" & HID.Value) = "Não existe romaneios para Notas com Operação Global"
            Return False
        End If

        Dim ds As New DataSet
        Dim operacao() As String = cmbSubOperacao.SelectedValue.Split("-")
        Sql = "SELECT top(1) 'TEM' as Tem" & vbCrLf &
              "  FROM Romaneios R" & vbCrLf &
              " INNER JOIN Produtos Pr" & vbCrLf &
              "    ON R.Produto = Pr.Produto_Id" & vbCrLf &
              " INNER JOIN RomaneiosXPesagens RxP" & vbCrLf &
              "    ON R.Empresa_Id    = RxP.Empresa_Id" & vbCrLf &
              "   AND R.EndEmpresa_Id = RxP.EndEmpresa_Id" & vbCrLf &
              "   AND R.Romaneio_Id   = RxP.Romaneio_Id" & vbCrLf &
              " INNER JOIN Pesagem P" & vbCrLf &
              "    ON RxP.Empresa_Id    = P.Empresa_Id" & vbCrLf &
              "   AND RxP.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf &
              "   AND RxP.Pesagem_Id    = P.Pesagem_Id" & vbCrLf &
              "   AND RxP.Sequencia_Id  = P.Sequencia_Id" & vbCrLf &
              "  FULL OUTER JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf &
              "    ON R.Empresa_Id    = nfxr.Empresa_Id" & vbCrLf &
              "   AND R.EndEmpresa_Id = nfxr.EndEmpresa_Id" & vbCrLf &
              "   AND R.Romaneio_Id   = nfxr.Romaneio_Id" & vbCrLf &
              "  LEFT JOIN NotasFiscais NF" & vbCrLf &
              "    ON nfxr.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
              "   AND nfxr.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
              "   AND nfxr.Cliente_Id      = NF.Cliente_id" & vbCrLf &
              "   AND nfxr.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf &
              "   AND nfxr.Nota_Id         = NF.Nota_Id" & vbCrLf &
              "   AND nfxr.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf &
              "   AND nfxr.Serie_Id        = NF.Serie_Id " & vbCrLf &
              " WHERE (R.Empresa_id    ='" & objNotaFiscal.CodigoEmpresa & "')" & vbCrLf &
              "   AND (R.EndEmpresa_id = " & objNotaFiscal.EnderecoEmpresa & ")" & vbCrLf &
              "   AND (R.Pedido        = " & Pedido & ")" & vbCrLf &
              "   AND (R.Movimento     <= '" & objNotaFiscal.Movimento.ToString("yyy-MM-dd") & "')" & vbCrLf &
              "   AND (nfxr.Nota_Id IS NULL OR NF.situacao NOT IN (1,4,7))  " & vbCrLf &
              "   AND (P.Situacao        = 1) " & vbCrLf &
              "   AND R.Operacao       = " & operacao(0) & vbCrLf &
              "   AND R.SubOperacao    = " & operacao(1) & vbCrLf


        ds = Banco.ConsultaDataSet(Sql, "Romaneios")

        If ds Is Nothing Then
            Return False
        ElseIf ds.Tables Is Nothing Then
            Return False
        ElseIf ds.Tables.Count = 0 Then
            Return False
        ElseIf ds.Tables(0).Rows.Count = 0 Then
            Return False
        Else
            Return True
        End If

    End Function

    Public Sub ListarClientes()
        ucConsultaClientes.Limpar()
        Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
        Popup.ConsultaDeClientes(Me, "objClienteNXI" & HID.Value, txtNome.ClientID, True, 500)
    End Sub

    Private Sub ListarRomaneios(ByVal Pedido As String)
        SessaoRecuperaNotaFiscal()
        Dim operacao() As String = cmbSubOperacao.SelectedValue.Split("-")
        Session("ProcurandoRomaneio" & HID.Value) = True
        Session("Op" & HID.Value) = operacao(0)
        Session("SOp" & HID.Value) = operacao(1)
        Session("Unitario" & HID.Value) = Str(objNotaFiscal.Itens(0).Unitario)
        Popup.ConsultaDeRomaneios(Me, "objRomaneioNXI" & HID.Value, "", True, 1000)
        ucConsultaRomaneios.LimparCampos()
        ucConsultaRomaneios.CargaRomaneios()
    End Sub

    Public Sub ConsultaPedidos(Optional ByVal carregarDados As Boolean = False)
        SessaoRecuperaNotaFiscal()
        If String.IsNullOrWhiteSpace(objNotaFiscal.CodigoCliente) Then
            MsgBox(Me.Page, "Cliente não foi selecionado")
            Exit Sub
        End If
        ucPedidoxSaldo.LimparCampos()
        If (carregarDados) Then
            ucPedidoxSaldo.CarregarDadosNotaFiscalXItem()
        End If
        Popup.ConsultaDePedidoXSaldoXFinanceiro(Me, "objItensPedidoSelecionadosNXI" & HID.Value, "txtPedido", True, 500)
    End Sub

    Private Sub BuscarSubOperacoes(ByVal Operacao As String, ByVal SubOperacao As String)
        cmbSubOperacao.Items.Clear()

        Dim objSubOperacoes As New [Lib].Negocio.ListSubOperacao()

        If objSubOperacoes.Selecionar(Operacao) Then
            For Each objSubOperacao As [Lib].Negocio.SubOperacao In objSubOperacoes
                cmbSubOperacao.Items.Add(New ListItem(objSubOperacao.Operacao.Codigo.ToString("00") & "-" & objSubOperacao.Codigo.ToString("00") & " " & objSubOperacao.Descricao,
                                                      Operacao & "-" & objSubOperacao.Codigo))
            Next

            cmbSubOperacao.Items.Insert(0, "")
            cmbSubOperacao.SelectedIndex = 0
            cmbSubOperacao.SelectedValue = Operacao & "-" & SubOperacao
        Else : MostrarTelaErro(objSubOperacoes.Erro)
        End If
    End Sub

    Public Function BuscarRomaneio() As Boolean
        SessaoRecuperaNotaFiscal()
        If objNotaFiscal.CodigoPedido = 0 Then
            Session("MsgControle" & HID.Value) = "Para consultar o(s) Romaneio(s) é necessário informar o Pedido"
            Return False
        ElseIf objNotaFiscal.CodigoRomaneio > 0 Then
            Session("MsgControle" & HID.Value) = "Romaneio já foi selecionado"
            Return False
        Else
            If objNotaFiscal.Itens(0).Produto.Agrupar = "N" And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL Then
                If TemRomaneio(objNotaFiscal.CodigoPedido) = True AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico Then
                    ListarRomaneios(objNotaFiscal.CodigoPedido)
                Else
                    If objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                        If objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.Itens(0).Produto.ControlarPesagem Then
                            BuscaClassificacao(False)
                        Else
                            BuscaProcuracao()
                        End If
                    Else
                        BuscaNotaTroca()
                    End If
                End If
            Else
                Session("MsgControle" & HID.Value) = "Não é informado a classificação para este produto!"
                Return False
            End If
        End If
        Return True
    End Function

    Public Function BuscaClassificacao(Optional ByVal delay As Boolean = False) As Boolean
        SessaoRecuperaNotaFiscal()
        If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL Then
            Session("MsgControle" & HID.Value) = "Não é possivel haver classificação de produto para notas globais."
            Return False
        End If
        If objNotaFiscal.CodigoRomaneio > 0 AndAlso Not objNotaFiscal.Romaneio.Processo = "NOTA FISCAL" Then
            Session("MsgControle" & HID.Value) = "A nota já possui um romaneio / classificação."
            Return False
        End If

        Dim validaRomaneio As Boolean = False

        If objNotaFiscal.Itens.Count > 1 Then
            For Each item In objNotaFiscal.Itens
                If item.Produto.Agrupar = "S" Then
                    validaRomaneio = True
                    Exit For
                End If
            Next
        End If

        If validaRomaneio Then Return False

        Session("ProcurandoClassificacao" & HID.Value) = True

        Popup.ConsultaDeNotaFiscalXClassificacao(Me, "objClassificacaoNXI" & HID.Value, "txtPrimeiraPesagem", delay, 1000)

        ucNotaFiscalXClassificacao.LimparCampos()
        ucNotaFiscalXClassificacao.BindGridView()
        ucNotaFiscalXClassificacao.CarregarPercentualPadraoClassificacao(objNotaFiscal.Quantidade, objNotaFiscal.Pedido.Itens(0).Produto.Codigo, False)

        Return True
    End Function

    Public Function BuscaProcuracao() As Boolean
        SessaoRecuperaNotaFiscal()
        If (objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.SubOperacao IsNot Nothing AndAlso objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso objNotaFiscal.SubOperacao.Devolucao AndAlso objNotaFiscal.Itens IsNot Nothing AndAlso objNotaFiscal.Itens.Count > 0 AndAlso objNotaFiscal.Itens(0).Produto.Agrupar = "N" AndAlso objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL) Then
            Session("ProcurandoProcuracao" & HID.Value) = True
            Dim parameters As New Dictionary(Of String, Object)
            parameters("tipo") = "NxI"
            parameters("pedc") = objNotaFiscal.CodigoPedido
            ucConsultaProcuracao.BindGridView(parameters)
            Popup.ConsultaDeProcuracao(Me, "objProcuracaoNxI" & HID.Value)
            Return True
        Else
            Session("MsgControle" & HID.Value) = "A Cessão de Crédito só pode ser vinculadas em notas de saída por devolução de produtos a granel!"
            If objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso
                (objNotaFiscal.SubOperacao.QuantidadeFisico OrElse objNotaFiscal.SubOperacao.QuantidadeFiscal) Then BuscaAutorizacao()
            Return False
        End If
    End Function

    Private Sub CarregarLoteEmbalagem()
        SessaoRecuperaNotaFiscal()
        If objNotaFiscal.IUD = "U" Then Exit Sub
        If (objNotaFiscal.Itens(gridItens.SelectedIndex).Produto.ControlarLote Or objNotaFiscal.Itens(gridItens.SelectedIndex).Produto.ControlarEmbalagem) And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL Then
            objListaDeCombinacoesLoteEmbalagem = New [Lib].Negocio.ListSaldoProdutoEstoqueLoteEmbalagem(objNotaFiscal.Itens(gridItens.SelectedIndex).CodigoProduto)


            If objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada Then
                objListaDeCombinacoesLoteEmbalagem.CarregarPossiveisEntradasLoteClassificacao(objNotaFiscal.Pedido.CodigoSafra)
                gridEmbalagem.DataSource = objListaDeCombinacoesLoteEmbalagem
            Else
                objListaDeCombinacoesLoteEmbalagem.CarregaSaldoDisponivelSaidaLoteClassificao(objNotaFiscal.SubOperacao, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                gridEmbalagem.DataSource = objListaDeCombinacoesLoteEmbalagem
                If objListaDeCombinacoesLoteEmbalagem.Count = 0 Then
                    objNotaFiscal.Itens.Remove(objNotaFiscal.Itens(gridItens.SelectedIndex))
                    objNotaFiscal.AtualizaTotais()
                    SessaoSalvaLoteEmbalagem()
                    SessaoSalvaNotaFiscal()
                    AtualizaFormularioComAClasse()
                    TabContainer1.ActiveTabIndex = 0
                    MsgBox(Me.Page, "Este Produto nao contem saldo, para emissao.")
                End If
            End If

            'Ocultar colunas Classificação e PesoSaco e mostrar a coluna Validade quando o lote for de Defensivos
            Dim contLoteDefensivo As Integer = 0, contLoteSemente As Integer = 0
            For Each objItemCombinacoes As [Lib].Negocio.SaldoProdutoEstoqueLoteEmbalagem In objListaDeCombinacoesLoteEmbalagem
                If objItemCombinacoes.TipoDeLote = 2 Then
                    contLoteDefensivo = contLoteDefensivo + 1
                ElseIf objItemCombinacoes.TipoDeLote = 1 Then
                    contLoteSemente = contLoteSemente + 1
                End If
            Next
            If contLoteDefensivo > 0 Then
                gridEmbalagem.Columns(6).Visible = True
                gridEmbalagem.Columns(7).Visible = False
                gridEmbalagem.Columns(8).Visible = False
            ElseIf contLoteSemente > 0 Then
                gridEmbalagem.Columns(6).Visible = True
                gridEmbalagem.Columns(7).Visible = True
                gridEmbalagem.Columns(8).Visible = True
            Else
                gridEmbalagem.Columns(6).Visible = False
                gridEmbalagem.Columns(7).Visible = False
                gridEmbalagem.Columns(8).Visible = False
            End If
            gridEmbalagem.DataBind()
        End If

        SessaoSalvaLoteEmbalagem()
        SessaoSalvaNotaFiscal()
    End Sub


    'Public Function ValidarNotaFiscal() As Boolean

    '    'Verificar se foi feito a recusa da Nota Fiscal
    '    If objNotaFiscal.TemRecusa Then
    '        erroMsg = "Nota Fiscal não pode ser lançada pois a mesma foi lançada como recusada."
    '        Return False
    '    End If

    '    'Verifica se o operação está com CLASSE cadastrada.
    '    If String.IsNullOrWhiteSpace(objNotaFiscal.Operacao.CodigoClasse) Then
    '        erroMsg = "Operação: " & objNotaFiscal.CodigoOperacao & " - " & objNotaFiscal.Operacao.Descricao & ". está sem Classe cadastrada!"
    '        Return False
    '    End If

    '    If Not Funcoes.VerificaAcesso(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, txtDataDeEntrada.Text, "NOTAS FISCAIS") Then
    '        erroMsg = "Movimento da Nota Fiscal já Fechado para esta data."
    '        Return False
    '    End If

    '    If Not Funcoes.VerificaAcesso(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, txtDataDeEntrada.Text, "CONTABIL") Then
    '        erroMsg = "Movimento Contábil já Fechado para esta data."
    '        Return False
    '    End If

    '    If objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada AndAlso objNotaFiscal.Empresa.Empresa.ObrigaChaveNf AndAlso Not objNotaFiscal.ChaveNFE.Length = 44 Then
    '        erroMsg = "Obrigatório a informação da Chave da Nota Eletrônica do Fornecedor."
    '        Return False
    '    End If

    '    'Solicitação via e-mail em 04/09/2023 pela Nutri - Furlan - 14/09/2023
    '    If objNotaFiscal.CodigoEmpresa = "05366261000224" AndAlso objNotaFiscal.EntradaSaida = eEntradaSaida.Saida AndAlso objNotaFiscal.CodigoLocalEmbarque.Length = 0 Then
    '        erroMsg = "Local de Embarque não foi informado, o mesmo é obrigatório para o Estado do MS."
    '        Return False
    '    End If

    '    'Verifica  a obrigação da Informação do Navio - DESABILITADO ATÉ ESTABILIZAR A IMPLANTAÇÃO
    '    'If objNotaFiscal.Empresa.Empresa.ObrigaNavio AndAlso txtNaviosXInvoice.Text.Length = 0 Then
    '    '    erroMsg = "Invoice não foi selecionada."
    '    '    Return False
    '    'End If

    '    'Verifica se a suboperação tem a Situação PIS/Cofins cadastrada.

    '    'Verifica se a soma das notas de relacionadas na devolucao batem com a nota de Devolucao.
    '    If objNotaFiscal.SubOperacao.Devolucao AndAlso Not objNotaFiscal.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE ICMS") _
    '        AndAlso ([Enum].Parse(GetType(eClassesOperacoes), objNotaFiscal.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.DEPOSITOS Or
    '                 [Enum].Parse(GetType(eClassesOperacoes), objNotaFiscal.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.COMPRAS Or
    '                 [Enum].Parse(GetType(eClassesOperacoes), objNotaFiscal.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.COMPRASAORDEM Or
    '                 [Enum].Parse(GetType(eClassesOperacoes), objNotaFiscal.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.EXPORTACOES Or
    '                 [Enum].Parse(GetType(eClassesOperacoes), objNotaFiscal.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.VENDAS Or
    '                 [Enum].Parse(GetType(eClassesOperacoes), objNotaFiscal.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.VENDASAORDEM) Then

    '        For Each item In objNotaFiscal.Itens

    '            If item.CodigoEmbalagem = 1 AndAlso item.QuantidadeFiscal > 0 Then
    '                'ainda vou testar - Furlan - 23/08/2022
    '                Dim vString As String = CStr(item.QuantidadeFiscal)
    '                Dim vDecimal As Decimal = CDec(vString.Split(",")(1))

    '                If vDecimal > 0 Then
    '                    erroMsg = "Produto a GRANEL não pode ter casa decimal." & vbCrLf
    '                    Return False
    '                End If

    '                If item.QuantidadeFisica > 0 Then
    '                    vString = CStr(item.QuantidadeFisica)
    '                    vDecimal = CDec(vString.Split(",")(1))

    '                    If vDecimal > 0 Then
    '                        erroMsg = "Produto a GRANEL não pode ter casa decimal." & vbCrLf
    '                        Return False
    '                    End If
    '                End If
    '            End If

    '            If objNotaFiscal.SubOperacao.QuantidadeFiscal AndAlso item.NotasDevolucao.Count = 0 Then
    '                erroMsg = "Não foi selecionado nenhuma nota fiscal do item Devolvido: " & item.CodigoProduto & " - " & item.Produto.Nome & " - Sequência: " & item.Sequencia
    '                Return False
    '            End If

    '            If objNotaFiscal.SubOperacao.QuantidadeFiscal AndAlso item.NotasDevolucao.SomaQtde <> item.QuantidadeFiscal Then
    '                erroMsg = "A Soma das Quantidades das notas de devolvidas nao batem com a quantidade do item Devolvido: " & item.CodigoProduto & " - " & item.Produto.Nome & " - Sequência: " & item.Sequencia
    '                Return False
    '            End If

    '            If Math.Round(item.NotasDevolucao.SomaVlr, 2, MidpointRounding.AwayFromZero) <> Math.Round(item.ValorTotal, 2, MidpointRounding.AwayFromZero) Then
    '                erroMsg = "A Soma dos Valores das notas devolvidas não batem com o Valor do item Devolvido: " & item.CodigoProduto & " - " & item.Produto.Nome & " - Sequência: " & item.Sequencia
    '                Return False
    '            End If

    '            Dim qDev As Decimal = 0
    '            Dim vDev As Decimal = 0

    '            For Each dev In item.NotasDevolucao
    '                qDev += dev.QuantidadeDevolucao
    '                vDev += dev.ValorDevolucao
    '            Next

    '            If qDev <> item.QuantidadeFiscal Then
    '                erroMsg = "A Soma das Quantidades das notas de devolvidas nao batem com a quantidade do item Devolvido: " & item.CodigoProduto & " - " & item.Produto.Nome & " - Sequência: " & item.Sequencia
    '                Return False
    '            End If

    '            If vDev <> item.ValorTotal Then
    '                erroMsg = "A Soma dos Valores das notas devolvidas não batem com o Valor do item Devolvido: " & item.CodigoProduto & " - " & item.Produto.Nome & " - Sequência: " & item.Sequencia
    '                Return False
    '            End If

    '            If objNotaFiscal.Empresa.CodigoEstado = "PR" AndAlso objNotaFiscal.EntradaSaida = eEntradaSaida.Saida AndAlso objNotaFiscal.ValorIcms = 0 Then
    '                If String.IsNullOrEmpty(item.OperacaoEstado.CodigoBeneficio) Then
    '                    erroMsg = "Código do Benefício não foi informado, verifique na Operação X Encargos o código: " & item.OperacaoEstado.Codigo
    '                    Return False
    '                End If
    '            End If

    '            If objNotaFiscal.Empresa.CodigoEstado = "PR" AndAlso objNotaFiscal.EntradaSaida = eEntradaSaida.Saida AndAlso objNotaFiscal.ValorIcms > 0 Then
    '                If Not String.IsNullOrEmpty(item.OperacaoEstado.CodigoBeneficio) Then
    '                    erroMsg = "Código do Benefício não pode ser informado, verifique na Operação X Encargos o código: " & item.OperacaoEstado.Codigo
    '                    Return False
    '                End If
    '            End If
    '        Next
    '    End If

    '    'Repete a Validacao da chave da nota caso seja eletronica S e nao seja nossa emissao
    '    If objNotaFiscal.Eletronica AndAlso objNotaFiscal.ChaveNFE.Length > 0 AndAlso Not objNotaFiscal.NossaEmissao Then

    '        Dim valida As Boolean = True

    '        If (Mid(objNotaFiscal.ChaveNFE, 7, 14) = "02935843000105" Or
    '            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "01409655000180" Or
    '            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "78393592000146" Or
    '            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "82951310000156" Or
    '            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "87958674000181" Or
    '            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "76416890000189" Or
    '            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "03507415000578") Then
    '            valida = False
    '        End If

    '        objNotaFiscal.ChaveNFE = objNotaFiscal.ChaveNFE.Replace(".", "")

    '        If valida AndAlso objNotaFiscal.CodigoCliente.Length = 11 AndAlso Not Mid(objNotaFiscal.ChaveNFE, 10, 11) = objNotaFiscal.CodigoCliente Then
    '            erroMsg = "Cliente na Chave Eletrônica diferente do informado na Nota Fiscal."
    '            Return False
    '        ElseIf valida AndAlso objNotaFiscal.CodigoCliente.Length = 14 AndAlso Not Mid(objNotaFiscal.ChaveNFE, 7, 14) = objNotaFiscal.CodigoCliente Then
    '            MsgBox(Me.Page, "Cliente na Chave Eletrônica diferente do informado na Nota Fiscal.")
    '            Return False
    '        ElseIf Not CInt(Left(objNotaFiscal.ChaveNFE, 2)) = objNotaFiscal.Cliente.Municipio.EstadoIbge Then
    '            erroMsg = "Estado do Cliente na Chave Eletrônica diferente do informado na Nota Fiscal."
    '            Return False
    '        ElseIf Not CInt(Mid(objNotaFiscal.ChaveNFE, 26, 9)) = CInt(txtNumero.Text) Then
    '            erroMsg = "Número da Nota na Chave Eletrônica diferente do informado na Nota Fiscal."
    '            Return False
    '        ElseIf Not CInt(Mid(objNotaFiscal.ChaveNFE, 23, 3)) = CInt(txtSerie.Text) Then
    '            erroMsg = "Série da Nota na Chave Eletrônica diferente do informado na Nota Fiscal."
    '            Return False
    '        ElseIf Not Mid(objNotaFiscal.ChaveNFE, 3, 2) = String.Format("{0:yy}", CDate(txtDataDeEmissao.Text)) Then
    '            erroMsg = "Ano da Nota na Chave Eletrônica diferente do informado na Nota Fiscal."
    '            Return False
    '            'ElseIf Not Mid(objNotaFiscal.ChaveNFE, 5, 2) = String.Format("{0:MM}", CDate(txtDataDeEmissao.Text)) Then
    '            '    erroMsg = "Mês da Nota na Chave Eletrônica diferente do informado na Nota Fiscal."
    '            '    Return False
    '        End If

    '    End If

    '    'Verifica obrigação da Nota de Produtor (NUTRI) - 27/08/2020 - FURLAN
    '    If objNotaFiscal.Empresa.Empresa.ObrigaNfProdutor AndAlso objNotaFiscal.ObrigaNFProdutor Then

    '        If objNotaFiscal.NotaTrocaOrigem Is Nothing Then
    '            erroMsg = "Nota Fiscal do Produtor não foi selecionada, faça a busca novamente."
    '            Return False
    '        ElseIf objNotaFiscal.NotaTrocaOrigem.Itens.Count = 0 Then
    '            erroMsg = "Nota Fiscal do Produtor não foi selecionada, faça a busca novamente."
    '            Return False
    '        ElseIf objNotaFiscal.NotaTrocaOrigem.Itens(0).QuantidadeFiscal = 0 AndAlso objNotaFiscal.NotaTrocaOrigem.Itens(0).ValorTotal = 0 Then
    '            erroMsg = "Nota Fiscal do Produtor não foi selecionada, faça a busca novamente."
    '            Return False
    '        ElseIf String.IsNullOrEmpty(objNotaFiscal.Cliente.InscricaoEstadual) OrElse objNotaFiscal.Cliente.InscricaoEstadual.Contains("ISENTO") Then
    '            erroMsg = "Inscrição Estadual do Produtor não foi informada, verifique no Cadastro de Clientes."
    '            Return False
    '        End If
    '    End If

    '    'furlan - Não deixar gravar encargo caso não tenha conta débito ou crédito - 18/07/2014
    '    '*********************************************************************************************************************************************************
    '    If objNotaFiscal.SubOperacao.Contabil AndAlso Not Left(objNotaFiscal.CodigoEmpresa, 8) = "04854422" AndAlso Not Left(objNotaFiscal.CodigoEmpresa, 8) = "03189063" AndAlso Not Left(objNotaFiscal.CodigoEmpresa, 8) = "24450490" Then
    '        For Each item In objNotaFiscal.Itens
    '            If String.IsNullOrWhiteSpace(item.SubOperacao.CodigoGrupoContas) Then
    '                erroMsg = "Operação " & objNotaFiscal.CodigoOperacao & "-" & objNotaFiscal.CodigoSubOperacao & " não possui conta contábil, verifique."
    '                Return False
    '            End If

    '            For Each enc In item.Encargos
    '                If Not enc.Codigo = "LIQUIDO" AndAlso enc.Valor > 0 Then
    '                    If String.IsNullOrWhiteSpace(enc.OperacaoEncargo.CodigoDebitaConta) AndAlso String.IsNullOrWhiteSpace(enc.OperacaoEncargo.CodigoCreditaConta) AndAlso enc.OperacaoEncargo.Sinal <> "=" Then
    '                        erroMsg = "Produto " & item.CodigoProduto & " ENCARGO " & enc.Codigo & " não possui conta DÉBITO/CRÉDITO, verifique."
    '                        Return False
    '                    End If
    '                End If
    '            Next
    '        Next
    '    End If
    '    'ATÉ AQUI

    '    If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES AndAlso objNotaFiscal.Empresa.Empresa.RegistroDeExportacao Then
    '        If objNotaFiscal.DadosDaExportacao Is Nothing Then
    '            erroMsg = "A Exportacao Exige um Registro de Exportacao, para Emissao da Nota."
    '            Return False
    '        End If
    '    End If

    '    If objNotaFiscal.FinanceiroNovo AndAlso objNotaFiscal.SubOperacao.Financeiro AndAlso Not objNotaFiscal.Pedido.SubOperacao.Classe.Equals(eClassesOperacoes.AFIXAR) _
    '       AndAlso Not (Not objNotaFiscal.Pedido.FinanceiroNovo AndAlso objNotaFiscal.Pedido.MomentoFinanceiro = eTipoFaturamento.Pedido) Then
    '        If (objNotaFiscal.Titulos.Count + objNotaFiscal.Titulos.AdiantamentosAbertos.Count = 0) Then
    '            erroMsg = "A suboperação possui financeiro, porém não há programação financeira para a nota fiscal!"
    '            Return False
    '        End If

    '        If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
    '            If (objNotaFiscal.Titulos.AdiantamentosAbertos.Sum(Function(T) T.VlrBaixa) + objNotaFiscal.Titulos.Sum(Function(T) T.Valores.EncargoValorDocumento.Valor)) <> objNotaFiscal.TotalNota Then
    '                erroMsg = "A suboperação possui financeiro, porém o valor financeiro programado não corresponde com o valor da nota fiscal!"
    '                Return False
    '            End If
    '        Else
    '            If (objNotaFiscal.Titulos.AdiantamentosAbertos.Sum(Function(T) T.VlrBaixa) + objNotaFiscal.Titulos.Sum(Function(T) T.Valores.EncargoValorDocumento.Valor)) <> Math.Round(objNotaFiscal.TotalNota / objNotaFiscal.IndiceNota, 2, MidpointRounding.AwayFromZero) Then
    '                erroMsg = "A suboperação possui financeiro, porém o valor financeiro programado não corresponde com o valor da nota fiscal!"
    '                Return False
    '            End If
    '        End If
    '    End If

    '    If Not objNotaFiscal.SubOperacao.Devolucao AndAlso
    '       objNotaFiscal.SubOperacao.Financeiro AndAlso
    '       objNotaFiscal.Pedido.MomentoFinanceiro = 3 Then

    '        If objNotaFiscal.VencimentosNota Is Nothing Then
    '            erroMsg = "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!"
    '            Return False
    '        ElseIf objNotaFiscal.VencimentosNota.Count = 0 Then
    '            erroMsg = "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!"
    '            Return False
    '        ElseIf Not objNotaFiscal.VencimentosNota.Sum(Function(s) s.ValorDoDocumento) = objNotaFiscal.TotalNota Then
    '            erroMsg = "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!"
    '            Return False
    '        End If
    '    End If

    '    If Not objNotaFiscal.VencimentosNota Is Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
    '        If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
    '            If Not objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoProvisao = 2).Sum(Function(s) s.ValorDoDocumento) = objNotaFiscal.TotalNota Then
    '                erroMsg = "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!"
    '                Return False
    '            End If
    '        Else
    '            If Math.Abs(Math.Round(objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoProvisao = 2).Sum(Function(s) s.MoedaValorDoDocumento), 2, MidpointRounding.AwayFromZero) - Math.Round(objNotaFiscal.TotalNotaMoeda, 2, MidpointRounding.AwayFromZero)) > 1 Then
    '                erroMsg = "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!"
    '                Return False
    '            End If
    '        End If

    '        Dim dataParcela = Today()
    '        Dim primeiro = True

    '        For Each parcela In objNotaFiscal.VencimentosNota
    '            If primeiro AndAlso objNotaFiscal.VencimentosNota.Count > 1 AndAlso parcela.Prorrogacao >= dataParcela Then
    '                dataParcela = parcela.Prorrogacao
    '                primeiro = False
    '            End If

    '            If objNotaFiscal.Empresa.Empresa.ControlaDataMovimentoNFG AndAlso parcela.Prorrogacao < dataParcela Then
    '                erroMsg = "Verifique o(s) vencimento(s) no financeiro programado pois a ordem não está correspondente!"
    '                Return False
    '            ElseIf parcela.ValorDoDocumento = 0 Then
    '                erroMsg = "Verifique o(s) vencimento(s), valor programado no título " & parcela.Codigo & " não pode ser 0(ZERO)."
    '                Return False
    '            End If

    '            dataParcela = parcela.Prorrogacao
    '        Next
    '    End If

    '    Dim saldoEmProvisao As Decimal
    '    For Each tit As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosPedido
    '        If tit.CodigoProvisao = 3 Then
    '            saldoEmProvisao += tit.ValorLiquido
    '        End If
    '    Next

    '    If objNotaFiscal.EntradaSaida = eEntradaSaida.Saida And objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso Not objNotaFiscal.Romaneio Is Nothing Then
    '        Dim qtde As Decimal
    '        For Each item As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
    '            qtde += item.QuantidadeFiscal
    '        Next

    '        If objNotaFiscal.Itens.Count = 1 Then
    '            If objNotaFiscal.Itens(0).Produto.Unidade = "KG" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "KGS" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "TON" Then
    '                If objNotaFiscal.Romaneio.PesoLiquido <> qtde Then
    '                    erroMsg = "Peso Fisico e Fiscal não pode ser Diferente."
    '                    Return False
    '                End If
    '            End If
    '        End If
    '    End If

    '    If objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada And objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso Not objNotaFiscal.Romaneio Is Nothing Then
    '        Dim qtde As Decimal
    '        For Each item As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
    '            qtde += item.QuantidadeFisica
    '        Next
    '        If objNotaFiscal.Romaneio.PesoLiquido <> qtde Then
    '            erroMsg = "Peso Físico da Nota não pode ser diferente do Romaneio, limpe e rafaça o processo."
    '            Return False
    '        End If
    '    End If

    '    If String.IsNullOrWhiteSpace(objNotaFiscal.CodigoDeposito) Then
    '        erroMsg = "Selecione o Deposito"
    '        Return False
    '    End If

    '    If String.IsNullOrWhiteSpace(objNotaFiscal.CodigoDestino) Then
    '        erroMsg = "Selecione o Deposito de origem/destino"
    '        Return False
    '    End If

    '    If Not [Enum].Parse(GetType(eClassesOperacoes), objNotaFiscal.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.IMPORTACOES AndAlso objNotaFiscal.Cliente.CodigoEstado = "EX" AndAlso objNotaFiscal.CodigoLocalEmbarque.ToString.Length = 0 Then
    '        erroMsg = "Local de Embarque não foi selecionado"
    '        Return False
    '    End If

    '    For Each row As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
    '        If row.Produto.ControlarNumeroDoLote AndAlso row.SubOperacao.ControlarNumeroDoLote OrElse (row.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso row.SubOperacao.ControlarNumeroDoLote AndAlso (row.Produto.CodigoGrupo = "10101" Or row.Produto.CodigoGrupo = "10102" Or row.Produto.CodigoGrupo = "30101" Or row.Produto.CodigoGrupo = "30102")) Then

    '            If objNotaFiscal.SubOperacao.EstoqueFisico Then
    '                If objNotaFiscal.Itens(0).Produto.Unidade = "KG" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "KGS" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "TON" Then
    '                    'NÃO FAZ NADA - LIBERADO PARA BAXI - FURLAN - 08/04/2024
    '                ElseIf row.QuantidadeFisica <> row.Lotes.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.Quantidade) Then
    '                    erroMsg = "Quantidade do Lote informada diferente da quantidade do Produto, verifique o item " & row.CodigoProduto & "-" & row.Produto.Nome
    '                    Return False
    '                End If
    '            Else
    '                If row.QuantidadeFiscal <> row.Lotes.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.Quantidade) Then
    '                    erroMsg = "Quantidade do Lote informada diferente da quantidade do Produto, verifique o item " & row.CodigoProduto & "-" & row.Produto.Nome
    '                    Return False
    '                End If
    '            End If

    '            For Each nLote In row.Lotes
    '                If nLote.Validade < objNotaFiscal.Movimento Then
    '                    erroMsg = "Validade informada no Lote não pode ser menor que a data da Nota Fiscal, verifique o item " & row.CodigoProduto & "-" & row.Produto.Nome & " lote " & nLote.Lote
    '                    Return False
    '                End If
    '            Next
    '        End If

    '        If row.Produto.ControlarNumeroDoLote AndAlso row.SubOperacao.ControlarNumeroDoLote OrElse (row.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso row.SubOperacao.Devolucao AndAlso row.SubOperacao.ControlarNumeroDoLote AndAlso (row.Produto.CodigoGrupo = "10101" Or row.Produto.CodigoGrupo = "10102" Or row.Produto.CodigoGrupo = "30101" Or row.Produto.CodigoGrupo = "30102")) Then

    '            If objNotaFiscal.SubOperacao.EstoqueFisico Then
    '                If objNotaFiscal.Itens(0).Produto.Unidade = "KG" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "KGS" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "TON" Then
    '                    'NÃO FAZ NADA - LIBERADO PARA BAXI - FURLAN - 08/04/2024
    '                ElseIf row.QuantidadeFisica <> row.Lotes.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.Quantidade) Then
    '                    erroMsg = "Quantidade do Lote informada diferente da quantidade do Produto, verifique o item " & row.CodigoProduto & "-" & row.Produto.Nome
    '                    Return False
    '                End If
    '            Else
    '                If row.QuantidadeFiscal <> row.Lotes.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.Quantidade) Then
    '                    erroMsg = "Quantidade do Lote informada diferente da quantidade do Produto, verifique o item " & row.CodigoProduto & "-" & row.Produto.Nome
    '                    Return False
    '                End If
    '            End If

    '            For Each nLote In row.Lotes
    '                If nLote.Validade < objNotaFiscal.Movimento Then
    '                    erroMsg = "Validade informada no Lote não pode ser menor que a data da Nota Fiscal, verifique o item " & row.CodigoProduto & "-" & row.Produto.Nome & " lote " & nLote.Lote
    '                    Return False
    '                End If
    '            Next
    '        End If

    '        If row.QuantidadeFiscal > row.SaldoPedidoFiscal And (row.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS Or row.SubOperacao.Classe = eClassesOperacoes.AFIXAR) And row.SubOperacao.Devolucao = True Then
    '            erroMsg = "A Quantidade informada não pode ser maior que o saldo do Pedido para Devolução"
    '            Return False
    '        End If

    '        If row.Produto.ControlarLote And row.Classificacao = "" And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL Then
    '            erroMsg = row.Produto.Nome & ", Este Produto é Controlado por Lote, informe o lote ao qual o produto pertence."
    '            Return False
    '        End If

    '        If row.Produto.ControlarEmbalagem And row.CodigoEmbalagem = 0 And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL And (objNotaFiscal.SubOperacao.QuantidadeFiscal Or objNotaFiscal.SubOperacao.QuantidadeFisico) Then
    '            erroMsg = row.Produto.Nome & ", Este Produto é Controlado por Embalagem, informe a Embalagem o qual o produto é Comercializado."
    '            Return False
    '        End If

    '        If row.Produto.ControlarPecas AndAlso row.SubOperacao.ControlarPecas AndAlso row.NumeroPecas = 0 Then
    '            erroMsg = row.Produto.Nome & "Não foi informado a quantidade de Peças/Meios do Produto."
    '            Return False
    '        End If

    '        'USADO PARA LIBERAR OPERAÇÕES DE IMPORTAÇÃO DA RT GRÃOS - FURLAN - 26/03/2024
    '        Dim validaIMP As Boolean = True
    '        If Left(objNotaFiscal.CodigoEmpresa, 8) = "24450490" AndAlso objNotaFiscal.SubOperacao.EstoqueFisico = False Then
    '            validaIMP = False
    '        End If

    '        If validaIMP And row.CFOP > 3000 AndAlso row.CFOP < 4000 And TabImportacao.Visible Then
    '            If objNotaFiscal.DadosDaImportacao Is Nothing Then
    '                erroMsg = "Declaração de Importação não foi Informada."
    '                Return False
    '            ElseIf Trim(objNotaFiscal.DadosDaImportacao.NumeroDeclaracaoImportacao).Length < 11 Then
    '                erroMsg = "Declaração de Importação do Produto " & row.CodigoProduto & " deve ter 11 dígitos."
    '                Return False
    '            ElseIf Trim(objNotaFiscal.DadosDaImportacao.LocalDesembarqueImportacao).Length = 0 Then
    '                erroMsg = "Local do Desembarque da Declaração de Importação do Produto " & row.CodigoProduto & " deve ser informado."
    '                Return False
    '            ElseIf objNotaFiscal.DadosDaImportacao.EstadoDesembarqueImportacao.Length = 0 Then
    '                erroMsg = "Estado do Desembarque da Declaração de Importação do Produto " & row.CodigoProduto & " deve ser informado."
    '                Return False
    '            ElseIf Trim(objNotaFiscal.DadosDaImportacao.CodigoFabricante).Length = 0 Then
    '                erroMsg = "Fabricante que está na Declaração de Importação do Produto " & row.CodigoProduto & " deve ser informado."
    '                Return False
    '            End If
    '        End If
    '    Next

    '    If Not objNotaFiscal.SubOperacao.Devolucao AndAlso
    '        objNotaFiscal.EntradaSaida = eEntradaSaida.Saida AndAlso
    '        objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.QuantidadeFiscal) = 0 AndAlso
    '        (objNotaFiscal.NotasReferenciais Is Nothing OrElse objNotaFiscal.NotasReferenciais.Count = 0) Then

    '        'TRATANDO PARA NOTA DE TRANSFERÊNCIA DE ICMS - FURLAN - 28/11/2023
    '        If Left(objNotaFiscal.CodigoEmpresa, 8) = "24450490" AndAlso objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.IMPORTACOES Then
    '            'LIBERA PARA INCLUIR POIS É PARA GERAR VARIAÇÃO CAMBIAL NEGATIVA - FURLAN - 16/07/2024
    '        ElseIf Not objNotaFiscal.SubOperacao.Descricao.Contains("TRANSFERENCIA DE CREDITO DE ICMS") Then
    '            erroMsg = "Nota fiscal Referencial não foi selecionada."
    '            Return False
    '        End If
    '    End If

    '    If objNotaFiscal.CodigoFinalidade = 28 AndAlso (objNotaFiscal.NotasReferenciais Is Nothing OrElse objNotaFiscal.NotasReferenciais.Count = 0) Then
    '        erroMsg = "Nota fiscal Referencial não foi selecionada."
    '        Return False
    '    End If

    '    If objNotaFiscal.Itens(0).CodigoFixacao = 0 AndAlso
    '       objNotaFiscal.Itens(0).Produto.ControlarPesagem AndAlso
    '       objNotaFiscal.CodigoRomaneio = 0 AndAlso
    '       objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso
    '       objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL Then
    '        erroMsg = "É Obrigatorio a seleção de um Romaneio ou fazer a classificaçao do produto, Lembrando que a classificação pela nota gera um Romaneio Fisico = Não."
    '        Return False
    '    End If

    '    If Not Session("TotalProcuracao" & HID.Value) Is Nothing Then
    '        If objNotaFiscal.ProcuracaoSaldoPedido > 0 And objNotaFiscal.ProcuracaoSaldoPedido < objNotaFiscal.Itens(0).SaldoValorOficial Then
    '            erroMsg = "Verifique o saldo referente as procurações. Não é permitido valores negativos."
    '            Return False
    '        End If
    '    End If

    '    If objNotaFiscal.CodigoRomaneio > 0 And
    '               objNotaFiscal.CriarRomaneio = False And
    '               (objNotaFiscal.CodigoFinalidade = 20 OrElse objNotaFiscal.CodigoFinalidade = 22) And
    '               objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) = "S" Then
    '        erroMsg = "Devolução fisica não pode ser usada com a finalidade de compra código 20 ou 22."
    '        Return False
    '    End If

    '    If objNotaFiscal.CodigoRomaneio > 0 And
    '               objNotaFiscal.CriarRomaneio = True And
    '               objNotaFiscal.SubOperacao.Devolucao = True And
    '               (objNotaFiscal.CodigoFinalidade <> 20 AndAlso objNotaFiscal.CodigoFinalidade <> 22) And
    '               objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) = "S" And
    '               (objNotaFiscal.CodigoAutorizacao = 0) And Not objNotaFiscal.TemNotaTroca Then
    '        erroMsg = "Devolução fisico = não, só pode ser feita com a finalidade código 20 ou 22."
    '        Return False
    '    End If

    '    If objNotaFiscal.NossaEmissao AndAlso txtCodigoTransportador.Value.ToString.Length > 0 AndAlso txtCodigoPlaca.Value.ToString.Length = 0 Then
    '        erroMsg = "Placa não foi selecionada."
    '        Return False
    '    End If

    '    If objNotaFiscal.NossaEmissao AndAlso Not objNotaFiscal.CodigoTransportador Is Nothing AndAlso objNotaFiscal.CodigoTransportador.Length > 0 Then
    '        If objNotaFiscal.PlacaDetalhes Is Nothing Then
    '            erroMsg = "Placa não foi selecionada."
    '            Return False
    '        ElseIf objNotaFiscal.PlacaDetalhes.Placa01.Length = 0 Then
    '            erroMsg = "Placa não foi selecionada."
    '            Return False
    '        End If
    '    End If

    '    If AutorizacaoENecessaria(objNotaFiscal) AndAlso Not objNotaFiscal.CodigoAutorizacao > 0 Then
    '        erroMsg = "Você deve selecionar uma autorizacao de retirada, caso não exista crie uma."
    '        Return False
    '    End If


    '    If objNotaFiscal.CodigoAutorizacao > 0 Then
    '        If objNotaFiscal.SubOperacao.QuantidadeFiscal AndAlso objNotaFiscal.Autorizacao.SaldoFiscal < IIf(objNotaFiscal.Itens(0).PesoQuantidade = "P", objNotaFiscal.Itens(0).PesoFiscal, objNotaFiscal.Itens(0).QuantidadeFiscal) Then
    '            erroMsg = "Saldo Fiscal Insuficiente para geração da Nota, Saldo:" & objNotaFiscal.Autorizacao.SaldoFiscal & " Nota Fiscal: " & objNotaFiscal.Itens(0).PesoFiscal
    '            Return False
    '        End If
    '        If objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.Autorizacao.SaldoFisico < objNotaFiscal.Itens(0).PesoLiquido Then
    '            erroMsg = "Saldo Físico Insuficiente para geração da Nota, Saldo:" & objNotaFiscal.Autorizacao.SaldoFisico & " Nota Fiscal: " & objNotaFiscal.Itens(0).PesoLiquido
    '            Return False
    '        End If
    '    End If

    '    If CDate(txtDataDeEmissao.Text) > Now.Date Then
    '        erroMsg = "A Data de Emissão não pode ser maior que a Data Atual."
    '        Return False
    '    End If


    '    If CDate(txtDataDeEntrada.Text) < CDate(txtDataDeEmissao.Text) Then
    '        erroMsg = "A Data de Emissão não pode ser posterior a Data do Movimento."
    '        Return False
    '    End If

    '    If String.IsNullOrWhiteSpace(txtNumero.Text) Then
    '        erroMsg = "Número da nota não foi informado."
    '        Return False
    '    End If

    '    If CInt(txtNumero.Text) = 0 Then
    '        MsgBox(Me.Page, "Número da Nota Fiscal não pode ser 0.")
    '        Return False
    '    End If

    '    If String.IsNullOrWhiteSpace(objNotaFiscal.Serie) Then
    '        erroMsg = "Serie não foi informada."
    '        Return False
    '    End If

    '    If cmbFinalidade.SelectedIndex = 0 Then
    '        erroMsg = "Finalidade não foi selecionada."
    '        Return False
    '    End If

    '    If chk_nfe.Checked AndAlso Not chk_NossaEmissao.Checked AndAlso txtES.Text = "E" AndAlso txtChaveNFe.Text.Replace(".", "").Length <> 44 Then
    '        erroMsg = "Informe a chave da nota eletrônica, deve ter 44 dígitos."
    '        Return False
    '    End If

    '    If gridItens.Rows.Count = 0 Then
    '        erroMsg = "Verifique itens da Nota Fiscal."
    '        Return False
    '    End If

    '    'If DgEncargos.Rows.Count = 0 Then
    '    '    erroMsg = "Verifique encargos da Nota Fiscal."
    '    '    Return False
    '    'End If

    '    If Not objNotaFiscal.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE ICMS") AndAlso
    '        Not objNotaFiscal.SubOperacao.Descricao.Contains("TRANSFERENCIA DE CREDITO DE ICMS") AndAlso
    '        Not objNotaFiscal.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE PRECO A FIXAR") AndAlso
    '        Not objNotaFiscal.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE IPI") Then
    '        If objNotaFiscal.TotalNota <= 0 Then
    '            erroMsg = "Valor da Nota Fiscal não pode ser zero."
    '            Return False
    '        End If

    '        'ADICIONADO EM 02-04-2013 REMOVER EM 2014
    '        For Each row In objNotaFiscal.Itens
    '            If row.ValorTotal = 0 Then
    '                erroMsg = row.Produto.Nome & " - Valor da Nota Fiscal não pode ser zero. ENTRE EM CONTATO COM A TI."
    '                Return False
    '                If row.Encargos.EncProduto.Valor = 0 Then
    '                    erroMsg = row.Produto.Nome & " - Encargo - Valor do Encargo da Nota Fiscal não pode ser zero. ENTRE EM CONTATO COM A TI."
    '                    Return False
    '                End If
    '            End If
    '        Next
    '    End If

    '    For Each prd As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
    '        If prd.Encargos.EncProduto.ValorPeso = eValorPeso.Valor Then
    '            If Math.Round((prd.ValorTotal), 2, MidpointRounding.AwayFromZero) <> prd.Encargos.EncProduto.Base Then
    '                erroMsg = "O Valor do produto " & prd.CodigoProduto & " - " & prd.Produto.Nome & " não confere com a Base De Calculo dos encargos. Cliquem em OK para atualizar os Totais do Produto e se o problema persisitr informe o TI."
    '                Return False
    '            End If
    '        Else
    '            If prd.QuantidadeFiscal <> prd.Encargos.EncProduto.Base Then
    '                erroMsg = "A Quantidade do produto " & prd.CodigoProduto & " - " & prd.Produto.Nome & " não confere com a Base De Calculo dos encargos, Cliquem em OK para atualizar os Totais do Produto e se o problema persisitr informe o TI."
    '                Return False
    '            End If
    '        End If

    '        For Each enc In prd.Encargos
    '            If (enc.Codigo = "ICMS" OrElse enc.Codigo = "ICMS A REC.") AndAlso enc.SituacaoTributaria = 0 AndAlso enc.Valor = 0 Then
    '                erroMsg = "Valor do Icms Tributado Integralmente não foi encontrado no Produto " & prd.CodigoProduto & "-" & prd.Produto.Nome & "."
    '                Return False
    '            End If
    '        Next

    '        If prd.Encargos.EncProduto.SituacaoTributariaIPI = 0 Or prd.Encargos.EncProduto.SituacaoTributariaIPI = 50 Then
    '            If Not prd.Encargos.Where(Function(s) s.Codigo.Contains("IPI")).Count > 0 Then
    '                erroMsg = "Produto " & prd.CodigoProduto & "-" & prd.Produto.Nome & " com SituacaoTributariaIPI = " & prd.Encargos.EncProduto.SituacaoTributariaIPI & " não existe o Encargo IPI parametrizado. Verifique parametrização na OperaçãoXEncargos."
    '                Return False
    '            ElseIf prd.Encargos.Where(Function(s) s.Codigo.Contains("IPI")).Sum(Function(s) s.Valor = 0) Then
    '                erroMsg = "Produto " & prd.CodigoProduto & "-" & prd.Produto.Nome & " com SituacaoTributariaIPI = " & prd.Encargos.EncProduto.SituacaoTributariaIPI & " não pode ser Zero(0)."
    '                Return False
    '            End If
    '        End If
    '    Next

    '    'If FinanceiroNovo AndAlso objNotaFiscal.Pedido.CondicaoPagamento IsNot Nothing AndAlso objNotaFiscal.Pedido.CondicaoPagamento.Antecipado AndAlso objNotaFiscal.TotalNota > objNotaFiscal.Titulos.AdiantamentosAbertos.ValorTotalInformadoParaBaixa Then
    '    '    erroMsg = "Para Pagamento/Recebimento Antecipado é necessário ter o valor: " & objNotaFiscal.TotalNota.ToString("N2") & " em Adiantamento"
    '    '    Return False
    '    'End If

    '    'If FinanceiroNovo AndAlso Not objNotaFiscal.SubOperacao.Devolucao AndAlso
    '    '    objNotaFiscal.SubOperacao.Financeiro AndAlso
    '    '    objNotaFiscal.TotalNota > objNotaFiscal.Pedido.Titulos.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao).Sum(Function(s) s.Valores.EncargoValorLiquido.ValorOficial) + objNotaFiscal.Titulos.AdiantamentosAbertos.ValorTotalInformadoParaBaixa AndAlso
    '    '    Not objNotaFiscal.Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR Then
    '    '    erroMsg = "Valor Total da Nota Fiscal não é compatível com o financeiro do pedido"
    '    '    Return False
    '    'End If

    '    'Obrigatória informação do número do processo de drawback para CFOP: 
    '    ' 7127: Venda de produção do estabelecimento sob o regime de drawback 
    '    ' 7211: Devolução de compras p/ industrialização sob o regime de drawback
    '    If (objNotaFiscal.CFOP.Codigo = 7127 OrElse objNotaFiscal.CFOP.Codigo = 7211) AndAlso (objNotaFiscal.DadosDaExportacao Is Nothing OrElse String.IsNullOrWhiteSpace(objNotaFiscal.DadosDaExportacao.NumAtoConcessorio)) Then
    '        erroMsg = "Para exportação em regime de Drawback é obrigatório o preenchimento do campo: Num. Ato Concessório (Drawbak), na aba exportação"
    '        Return False
    '    End If

    '    'Para NF de importação não preencher o campo Local de Embarque/Coleta.
    '    If [Enum].Parse(GetType(eClassesOperacoes), objNotaFiscal.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.IMPORTACOES AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.CodigoLocalEmbarque) Then
    '        erroMsg = "O campo Local de Embarque/Coleta não deve ser preenchido em Nota Fiscal de importação"
    '        Return False
    '    End If


    '    'Para NF de exportação, obrigatório JABER/ZÉLIO - 22/10/2018
    '    If Left(objNotaFiscal.CodigoEmpresa, 8) = "03189063" _
    '        AndAlso objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES _
    '        AndAlso (objNotaFiscal.DadosDaExportacao Is Nothing OrElse objNotaFiscal.DadosDaExportacao.FaturaExportacao.Length = 0) _
    '        AndAlso objNotaFiscal.EntradaSaida = eEntradaSaida.Saida Then

    '        erroMsg = "Número da Fatura de Exportação é obrigatório"
    '        Return False
    '    End If

    '    'Validação Xml de Notas de Entrada'
    '    'If SessaoDsXML IsNot Nothing Then
    '    'erroMsg = DocumentoEletronico.ValidaNFEntradaComXML(objNotaFiscal, SessaoDsXML)
    '    '    If Not String.IsNullOrWhiteSpace(erroMsg) Then
    '    '        Return False
    '    '    End If
    '    'End If

    '    'Validar Totals dos Produtos e Total da Nota Fiscal
    '    If SessaoDsXML IsNot Nothing Then

    '        If objNotaFiscal.TemSegCodBarra AndAlso Not txtSegCodBarra.Text.Length = 36 Then
    '            erroMsg = "O Segundo Código de Barras não está preenchido corretamente."
    '            Return False
    '        End If

    '        Dim XmlValorTotalProduto As Decimal = CDec(CType(Session("dsXML" & HID.Value), DataSet).Tables("ICMSTot").Rows(0)("vProd").ToString().Replace(".", ","))
    '        Dim XmlValorTotalNota As Decimal = CDec(CType(Session("dsXML" & HID.Value), DataSet).Tables("ICMSTot").Rows(0)("vNF").ToString().Replace(".", ","))

    '        Dim temDesconto As Boolean = False

    '        If CDec(CType(Session("dsXML" & HID.Value), DataSet).Tables("ICMSTot").Rows(0)("vDesc").ToString().Replace(".", ",")) > 0 Then temDesconto = True

    '        If Not objNotaFiscal.TotalProduto = XmlValorTotalProduto Then
    '            erroMsg = "O valor total do(s) Produto(s) no XML " & XmlValorTotalProduto & " está diferente do valor da NF " & objNotaFiscal.TotalProduto.ToString()
    '            Return False
    '        End If

    '        If Not objNotaFiscal.TotalNota = XmlValorTotalNota Then
    '            Dim totalGeral As Decimal = objNotaFiscal.TotalNota

    '            For Each prd In objNotaFiscal.Itens
    '                For Each enc In prd.Encargos

    '                    If Not enc.Codigo = "PRODUTO" AndAlso enc.Sinal = "+" Then
    '                        totalGeral -= enc.Valor
    '                    ElseIf enc.Sinal = "-" Then
    '                        totalGeral += enc.Valor
    '                    End If

    '                    'If enc.Codigo = "FUNRURAL" OrElse enc.Codigo = "SENAR" OrElse enc.Codigo = "FETHAB" OrElse enc.Codigo = "IAGRO" Then
    '                    '    totalGeral += enc.Valor
    '                    'End If
    '                Next
    '            Next

    '            If Not temDesconto AndAlso Not totalGeral = XmlValorTotalNota Then
    '                erroMsg = "O valor total da NF no XML " & XmlValorTotalNota & " está diferente do valor da NF " & objNotaFiscal.TotalNota.ToString()
    '                Return False
    '            End If
    '        End If
    '    End If

    '    erroMsg = ""
    '    Return True
    'End Function

    Public Sub BuscarPlaca()
        If txtCodigoTransportador.Value.Length > 0 Then
            ucConsultaPlacas.Limpar()
            Popup.ConsultaDePlacas(Me.Page, "objPlaca" & HID.Value)
            Popup.CenterDialog(Me.Page, "divConsultaPlacas")
        Else
            MsgBox(Me.Page, "Transportador não foi selecionado")
        End If
    End Sub

    Public Function BuscaNotaTroca() As Boolean
        SessaoRecuperaNotaFiscal()

        'Sai da funcao se a nota for de entrada e a finalidade nao for
        '13 - Compra em Deposito
        'ou
        '23 - Compra por transferencia de Titularidade
        'If objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) = "E" And objNotaFiscal.CodigoFinalidade <> 13 And objNotaFiscal.CodigoFinalidade <> 23 Then Exit Function

        If objNotaFiscal.Itens.Count = 0 Then
            Session("MsgControle" & HID.Value) = "A Nota deve conter um Produto para a troca de nota ser realizada"
            Return False
        End If

        'Vamos ter que tratar esse tipo de troca(Pedido com mais de 1 item) - Furlan - 19/02/2024
        'If objNotaFiscal.Itens(0).Produto.Agrupar.Equals("S") Then
        '    Session("MsgControle" & HID.Value) = "Não é permitido fazer a troca de nota para este produto"
        '    Return False
        'End If

        If objNotaFiscal.Pedido.SubOperacao.PrecoFixo = False And objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) = "E" And objNotaFiscal.CodigoFinalidade = 13 Then
            Session("MsgControle" & HID.Value) = "A Devolução para Compra só é permitida em operações de preço Fixo"
            Return False
        End If

        Session("ProcurandoTrocaDeNota" & HID.Value) = True

        ucConsultaNotaTroca.InicializarFormulario()
        Popup.ConsultaDeNotaTroca(Me, "objTrocaDeNotaNXI" & HID.Value)

        Return True
    End Function

    Private Sub BuscaVendaAOrdem()
        Session.Remove("SemVendaAOrdem" & HID.Value)
        Session("ProcurandoVendaAOrdem" & HID.Value) = True
        ucConsultaNotaVendaAOrdem.BindGridView()
        Popup.ConsultaDeNotaVendaAOrdem(Me, "objVendaAOrdemNXI" & HID.Value)
    End Sub

    Public Function BuscaAutorizacao() As Boolean
        SessaoRecuperaNotaFiscal()

        If Left(objNotaFiscal.CodigoEmpresa, 8) = "24450490" Then
            Return False
        End If

        'If objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.SubOperacao IsNot Nothing AndAlso _
        '         objNotaFiscal.SubOperacao.EntradaSaida.ToString.Substring(0, 1) = "S" AndAlso _
        '         objNotaFiscal.Itens IsNot Nothing AndAlso objNotaFiscal.Itens.Count > 0 AndAlso _
        '         objNotaFiscal.Itens(0).Produto.Agrupar = "N" AndAlso _
        '         (objNotaFiscal.CodigoFinalidade <> 20 AndAlso objNotaFiscal.CodigoFinalidade <> 22) Then
        '    ucConsultaAutorizacaoDeRetirada.BindGridView()
        '    Popup.ConsultaDeAutorizacaoDeRetirada(Me, "objAutorizacaoNXI" & HID.Value, True, 500)
        '    Return True
        'End If

        If objNotaFiscal.AutorizacaoENecessaria(objNotaFiscal) Then
            Popup.ConsultaDeAutorizacaoDeRetirada(Me, "objAutorizacaoNXI" & HID.Value, True, 500)
            ucConsultaAutorizacaoDeRetirada.BindGridView()
            Return True
        End If

        objNotaFiscal.CodigoAutorizacao = 0
        objNotaFiscal.Autorizacao = Nothing
        SessaoSalvaNotaFiscal()

        If objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.SubOperacao IsNot Nothing AndAlso objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
            Session("MsgControle" & HID.Value) = "A Autorização é usada somente em Operações de Saída"
            Return False
        End If

        If objNotaFiscal.Itens IsNot Nothing AndAlso objNotaFiscal.Itens.Count > 0 AndAlso objNotaFiscal.Itens(0).Produto.Agrupar = "S" Then
            Session("MsgControle" & HID.Value) = "A Autorização não pode ser usada com este produto"
            Return False
        End If

        If (objNotaFiscal.CodigoFinalidade = 14 OrElse objNotaFiscal.CodigoFinalidade = 20 OrElse objNotaFiscal.CodigoFinalidade = 22) Then
            Session("MsgControle" & HID.Value) = "A Autorização não é Requerida com as Finalidades 14, 20 ou 22"
            Return False
        End If
        Return True
    End Function

    Private Function AutorizacaoENecessaria(ByRef objNF As NotaFiscal) As Boolean

        If objNF IsNot Nothing AndAlso objNF.SubOperacao IsNot Nothing AndAlso
                 objNF.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso
                 (objNF.SubOperacao.QuantidadeFisico OrElse objNF.SubOperacao.QuantidadeFiscal) AndAlso
                 objNF.SubOperacao.Classe <> eClassesOperacoes.GLOBAL AndAlso
                 objNF.Itens IsNot Nothing AndAlso objNF.Itens.Count > 0 AndAlso
                 (objNF.Itens(0).Produto.Agrupar = "N" Or objNF.Itens(0).Produto.AutorizacaoDeRetirada) AndAlso
                 (objNF.CodigoFinalidade <> 14 AndAlso objNF.CodigoFinalidade <> 20 AndAlso objNF.CodigoFinalidade <> 22) Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Sub LimparCampos(ByVal evento As String, ByVal Numero As String, ByVal NossaEmissao As Boolean, ByVal pEmpresa As String, ByVal pEndEmpresa As Integer)
        If Not objNotaFiscal Is Nothing Then
            If String.IsNullOrWhiteSpace(objNotaFiscal.IUD) Then
                SessaoRecuperaNotaFiscal()
                logs = New FuncoesLogs(1, Session("ssCnpjDaEmpresa" & HID.Value))
                logs.RegistrarLog(objNotaFiscal.IUD, objNotaFiscal)
            End If
        End If

        Session.Remove("objHorarioServidor" & HID.Value)
        Session.Remove("objPedidoSelecionado" & HID.Value)
        Session.Remove("objItensPedidoSelecionadosNXI" & HID.Value)
        Session.Remove("objClienteNXI" & HID.Value)
        Session.Remove("objDepositoNXI" & HID.Value)
        Session.Remove("objEmbarqueNXI" & HID.Value)
        Session.Remove("objEntregaNXI" & HID.Value)
        Session.Remove("objTransportadorNXI" & HID.Value)
        Session.Remove("objOrigemDestinoNXI" & HID.Value)
        Session.Remove("objTransbordoNXI" & HID.Value)
        Session.Remove("ProcurandoRomaneio" & HID.Value)
        Session.Remove("ProcurandoTrocaDeNota" & HID.Value)
        Session.Remove("ProcurandoClassificacao" & HID.Value)
        Session.Remove("objAutorizacaoNXI" & HID.Value)
        Session.Remove("SemTrocaDeNota" & HID.Value)
        Session.Remove("SemAutorizacao" & HID.Value)
        Session.Remove("SemProcuracao" & HID.Value)
        Session.Remove("MsgControle" & HID.Value)
        Session.Remove("ssCampo" & HID.Value)
        Session.Remove("ProcurandoRomaneio" & HID.Value)
        Session.Remove("ProcurandoClassificacao" & HID.Value)
        Session.Remove("ProcurandoProcuracao" & HID.Value)
        Session.Remove("TotalProcuracao" & HID.Value)
        Session.Remove("ProcurandoTrocaDeNota" & HID.Value)
        Session.Remove("ssMessage" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)
        Session.Remove("objNotaFiscalOriginal" & HID.Value)
        Session.Remove("objEmpresaNXI" & HID.Value)
        Session.Remove("objPlaca" & HID.Value)
        Session.Remove("objNFConsultaNXI" & HID.Value)
        Session.Remove("objTrocaDeNotaNXI" & HID.Value)
        Session.Remove("objRomaneioNXI" & HID.Value)
        Session.Remove("objProcuracaoNxI" & HID.Value)
        Session.Remove("TotalProcuracao" & HID.Value)
        Session.Remove("ProcurandoProcuracao" & HID.Value)
        Session.Remove("SemProcuracao" & HID.Value)
        Session.Remove("objClassificacaoNXI" & HID.Value)
        Session.Remove("ProcurandoClassificacao" & HID.Value)
        Session.Remove("objAutorizacaoNXI" & HID.Value)
        Session.Remove("SemAutorizacao" & HID.Value)
        Session.Remove("Observacoes" & HID.Value)
        Session.Remove("SemVendaAOrdem" & HID.Value)
        Session.Remove("ProcurandoVendaAOrdem" & HID.Value)
        Session.Remove("objVendaAOrdemNXI" & HID.Value)
        Session.Remove("objUFEmbarqueDI" & HID.Value)
        Session.Remove("objUFDesembarqueDI" & HID.Value)
        Session.Remove("objFabricanteDI" & HID.Value)
        Session.Remove("objEmailNFe" & HID.Value)
        Session.Remove("ObjNFObsProduto" & HID.Value)
        Session.Remove("objNFDeProdutor" & HID.Value)
        Session.Remove("objNFReferencialSaida" & HID.Value)
        Session.Remove("objEstoqueMinimo" & HID.Value)
        Session.Remove("objLoteFornecedor" & HID.Value)
        Session.Remove("chaveXMLautomacao" & HID.Value)
        Session.Remove("objConsultarNaviosXInvoice" & HID.Value)
        Session.Remove("objNavioXInvoice" & HID.Value)
        Session.Remove("objTransferencias" & HID.Value)
        Session.Remove("dsXml" & HID.Value)
        Session.Remove("ssNomeArquivoPedido" & HID.Value)

        SessaoDsXML = Nothing

        chkImportarProdutoUnico.Checked = False
        chkAGruparNCM.Checked = False

        HID.Value = Guid.NewGuid().ToString
        ucConsultaEmpresas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        ucPedidoxSaldo.SetarHID(HID.Value)
        ucConsultaNotaTroca.SetarHID(HID.Value)
        ucConsultaRomaneios.SetarHID(HID.Value)
        ucConsultaObservacoes.SetarHID(HID.Value)
        ucConsultaPlacas.SetarHID(HID.Value)
        ucConsultaEstados.SetarHID(HID.Value)
        ucConsultaCodMunicipios.SetarHID(HID.Value)
        ucNotaFiscalXClassificacao.SetarHID(HID.Value)
        ucConsultaNotaVendaAOrdem.SetarHID(HID.Value)
        ucConsultaAutorizacaoDeRetirada.SetarHID(HID.Value)
        ucConsultaEncargosPlanoDeContas.SetarHID(HID.Value)
        ucConsultaDadosBancarios.SetarHID(HID.Value)
        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucNotaFiscalReferencial.SetarHID(HID.Value)
        ucNotaDeDevolucaoXNota.SetarHID(HID.Value)
        ucConsultaProcuracao.SetarHID(HID.Value)
        ucInutilizacao.SetarHID(HID.Value)
        ucVencimentos.SetarHID(HID.Value)
        ucMonitorDeNotas.SetarHID(HID.Value)
        ucEmailNFe.SetarHID(HID.Value)
        ucFile.Clear()
        ucNFObsProduto.SetarHID(HID.Value)
        ucNFEncargo.SetarHID(HID.Value)
        ucConsultaProduto.SetarHID(HID.Value)
        ucNFProdutor.SetarHID(HID.Value)
        ucNFReferencialSaida.SetarHID(HID.Value)
        ucConsultarNaviosXInvoice.SetarHID(HID.Value)
        ucTransferencias.SetarHID(HID.Value)

        txtNaviosXInvoice.Text = String.Empty

        Dim dtLoteFornecedor As New DataTable("ItemLoteFornecedor")
        dtLoteFornecedor.Columns.Add("Produto", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Lote", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Fabricado", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Validade", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
        dtLoteFornecedor.Columns.Add("Consumo", Type.GetType("System.Decimal"))
        Session("objLoteFornecedor" & HID.Value) = dtLoteFornecedor

        Dim dtEstoqueMinimo As New DataTable("ItemEstoqueMinimo")
        dtEstoqueMinimo.Columns.Add("Produto", Type.GetType("System.String"))
        dtEstoqueMinimo.Columns.Add("Nome", Type.GetType("System.String"))
        dtEstoqueMinimo.Columns.Add("EstoqueMinimo", Type.GetType("System.String"))
        dtEstoqueMinimo.Columns.Add("Faturando", Type.GetType("System.String"))
        dtEstoqueMinimo.Columns.Add("Saldo", Type.GetType("System.String"))
        Session("objEstoqueMinimo" & HID.Value) = dtEstoqueMinimo

        TabImportacao.Visible = False
        TabExportacao.Visible = False

        bntBuscarNFReferencial.Visible = False
        bntBuscarNFProdutor.Visible = False
        TabNotasDeProdutor.Visible = False

        Session("Carregando" & HID.Value) = False
        LimparDadosdaExportacao()

        imgConfirmarDI.Enabled = False
        imgUsuario.Visible = False
        lblUsuario.Visible = False
        btnDeposito.Enabled = True
        imgExtratoPedido.Visible = False
        imgRomaneio.Visible = False
        btnRomaneio.Visible = True

        HabilitarBotoes()

        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        lnkLimpar.Parent.Visible = True
        lnkEspelho.Parent.Visible = False
        lnkEnviarSEFAZ.Parent.Visible = False
        lnkEnviarEmail.Parent.Visible = False
        lnkPdf.Parent.Visible = False
        lnkImpressora.Parent.Visible = False

        'lkn_ModeloNota.Parent.Visible = True
        'lnkVisualizar.Parent.Parent.Parent.Visible = False
        IdVisualizarNFe.Visible = False
        IdModelNota.Visible = False
        IdVisualizar.Visible = False

        BtnRecontabilizar.Enabled = False
        btnInutilizar.Enabled = True
        BtnVencimentos.Enabled = True
        btnProcuracao.Enabled = True

        txtNomeDaEmpresa.Text = ""
        txtInscricaoDaEmpresa.Text = ""
        txtCnpjDaEmpresa.Text = ""

        lblCliente.Text = "DESTINATÁRIO/REMETENTE"
        txtNomeDoCliente.Text = ""
        txtInscricaoDoCliente.Text = ""
        txtCnpjDoCliente.Text = ""
        txtEnderecoDoCliente.Text = ""
        txtComplementoDoCliente.Text = ""
        txtBairroDoCliente.Text = ""
        txtCidadeDoCliente.Text = ""
        txtTelefoneDoCliente.Text = ""
        txtEstadoDoCliente.Text = ""
        txtInscricaoDoCliente.Text = ""
        txtCepDoCliente.Text = ""
        txtCodigoCliente.Value = ""
        txtChaveNFe.Text = ""
        txtSegCodBarra.Text = ""
        observacaoValores.Value = ""
        observacaoAutorizacao.Value = ""

        txtDataDeEmissao.Enabled = True
        txtDataDeEntrada.Enabled = True

        txtDataDeEmissao.Text = Now.ToString("dd/MM/yyyy")
        txtDataDeEntrada.Text = Now.ToString("dd/MM/yyyy")
        txtHoraDaSaida.Text = Now.ToString("HH:mm:ss")
        txtPedido.Text = ""

        cmbSubOperacao.Items.Clear()
        cmbFinalidade.SelectedIndex = 0
        cmbFinalidade.Enabled = True
        ddlTipoDeDocumento.SelectedValue = eTipoDeDocumento.Nota

        txtRomaneio.Text = ""
        txtPesoRomaneio.Text = ""
        txtCessaoDeCredito.Text = ""
        txtAutorizacao.Text = ""

        gridItens.DataSource = Nothing
        gridItens.DataBind()

        gridVencimentosNota.DataSource = Nothing
        gridVencimentosNota.DataBind()
        gridVencimentosPedido.DataSource = Nothing
        gridVencimentosPedido.DataBind()
        ucVencimentos.Limpar()

        gridContabilizacao.DataSource = Nothing
        gridContabilizacao.DataBind()

        gridEmbalagem.DataSource = Nothing
        gridEmbalagem.DataBind()

        gridNotasDeProdutor.DataSource = Nothing
        gridNotasDeProdutor.DataBind()

        ddlDeposito.Enabled = True
        ddlOrigemDestino.Enabled = True
        ddlEmbarque.Enabled = True
        ddlTransbordo.Enabled = True
        ddlEntrega.Enabled = True
        btnDeposito.Enabled = True
        BtnOrigemDestino.Enabled = True
        btnEmbarque.Enabled = True
        BtnTransbordo.Enabled = True

        ddlDeposito.Items.Clear()
        ddlOrigemDestino.Items.Clear()
        ddlEmbarque.Items.Clear()
        ddlTransbordo.Items.Clear()
        ddlEntrega.Items.Clear()

        txtValorTotalDosProdutos.Text = ""
        txtValorTotalDaNota.Text = ""
        txtBaseIcmsNota.Text = ""
        txtValorIcmsNota.Text = ""
        txtValorIPINota.Text = ""
        txtValorBaseIcmsSTNota.Text = ""
        txtValorIcmsSTNota.Text = ""
        txtValorFrete.Text = ""
        txtDesconto.Text = ""
        txtOutras.Text = ""
        txtChaveNFe.ReadOnly = True
        txtSegCodBarra.ReadOnly = True
        lnkVerificarChaveNFE.Visible = False
        txtTextoSegCodBarra.Visible = False
        txtSegCodBarra.Visible = False
        txtNumero.Enabled = True
        txtSerie.Enabled = True
        txtES.Text = ""
        txtNumero.Text = ""
        txtSerie.Text = ""
        chk_NossaEmissao.Enabled = True
        chk_NossaEmissao.Checked = False
        chk_nfe.Enabled = True
        chk_nfe.Checked = False
        txtSituacao.Text = ""
        ChkTroca.Checked = False

        txtNomeDoTransportador.Text = ""
        txtEnderecoDoTransportador.Text = ""
        txtCidadeDoTransportador.Text = ""
        txtEstadoDoTransportador.Text = ""
        txtCnpjDoTransportador.Text = ""
        txtInscricaoDoTransportador.Text = ""
        txtCodigoTransportador.Value = ""

        txtPlacas.Text = ""
        txtEstadoDaPlaca.Text = ""
        txtCodigoPlaca.Value = ""

        txtVolumes.Text = ""
        txtEspecie.Text = ""
        txtMarca.Text = ""
        txtNumeracao.Text = ""

        txtPesoBruto.Text = ""
        txtPesoLiquido.Text = ""
        txtObservacoesDeEmbarque.Text = ""

        txtObservacoesFiscais.Text = ""
        txtObservacoesFiscais.Enabled = False

        txtObservacaoCancelamento.Text = ""
        txtObservacoesInternas.Text = ""
        txtNaturezaDaOperacao.Text = ""

        txtDI.Text = ""
        txtDataDI.Text = ""
        TxtRegImportacao.Text = ""
        TxtDataRegImp.Text = ""
        TxtNumAtoConcessorioImp.Text = ""
        TxtDtaRegAtoConcessorioImp.Text = ""
        TxtDtaValidAtoConcessorioImp.Text = ""
        txtEmbarqueDI.Text = ""
        lblUFEmbarqueDI.Text = ""
        txtDataEmbarqueDI.Text = ""
        txtDesembarqueDI.Text = ""
        lblUFDesembarqueDI.Text = ""
        txtDataDesembarqueDI.Text = ""
        lblFabricanteDI.Text = ""
        TxtNrFaturaImp.Text = ""
        txtCodigoUFEmbarqueDI.Value = ""
        txtCodigoUFDesembarqueDI.Value = ""
        txtCodigolblFabricanteDI.Value = ""
        ddlViaDeTransporteDI.ClearSelection()
        ddlTipoDeImportacaoDI.ClearSelection()

        objNotaFiscal = New [Lib].Negocio.NotaFiscal
        objNotaFiscal.IUD = "I"

        objNotaFiscal.Usuario = Session("ssNomeUsuario")
        objNotaFiscal.UsuarioInclusao = Session("ssNomeUsuario")
        objNotaFiscal.DataInclusao = Format(Date.Now, "yyyy-MM-dd HH:mm:ss")
        ddlUsuarios.Items.Clear()
        ddlUsuarios.Items.Add("Inc.- " & objNotaFiscal.UsuarioInclusao)

        CarregarDDLFrete(False)

        SessaoSalvaNotaFiscal()

        If primeiraVez.Value = True Then
            primeiraVez.Value = False
            SetarEmpresa(pEmpresa, pEndEmpresa, True)
        Else
            SetarEmpresa(pEmpresa, pEndEmpresa, False)
        End If

        txtNumViasDeImpressao.Text = objNotaFiscal.Empresa.Empresa.ViasNFE
        txtChaveNFe.ReadOnly = Not chk_NossaEmissao.Checked
        grdNotasReferenciadas.DataSource = Nothing
        grdNotasReferenciadas.DataBind()

        If FinanceiroNovo Then
            ucFinanceiro.SetarHID(HID.Value)
            ucFinanceiro.SetarOrigem = "NF"
            ucFinanceiro.CarregarControle()
        End If

        TabContainer1.ActiveTabIndex = 0

        Dim Ip As String = Context.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
        If Ip = String.Empty Then
            Ip = Context.Request.ServerVariables("REMOTE_ADDR")
            If Ip = "::1" Then Ip = "127.0.0.1"
        End If

        lblAcessoUsuario.Text = Ip

        If evento = "E" Then
            MsgBox(Me.Page, "Nota fiscal " & Numero & " removida com Sucesso.", eTitulo.Sucess)
        ElseIf evento = "I" Then
            MsgBox(Me.Page, "Nota fiscal " & Numero & " incluída com Sucesso.", eTitulo.Sucess)
        ElseIf evento = "U" Then
            MsgBox(Me.Page, "Nota fiscal " & Numero & " alterada com Sucesso.", eTitulo.Sucess)
        ElseIf evento = "C" Then
            MsgBox(Me.Page, "Nota fiscal " & Numero & " cancelada com Sucesso.", eTitulo.Sucess)
        End If
    End Sub

    Protected Sub chkImportarXML_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim chk As CheckBox = CType(sender, CheckBox)

        If chk.ID = "chkImportarProdutoUnico" And chk.Checked Then
            chkAGruparNCM.Checked = False
        ElseIf chk.ID = "chkAGruparNCM" And chk.Checked Then
            chkImportarProdutoUnico.Checked = False
        End If

    End Sub

    Public Sub LimparDadosdaExportacao()
        txtNrDespacho.Text = ""
        txtDataDespacho.Text = ""
        txtNrConhecimento.Text = ""
        txtDataNrConhecim.Text = ""
        txtRe.Text = ""
        txtDataRe.Text = ""
        ddlUF.SelectedIndex = 0
        ddlPaisDestino.SelectedIndex = 0
        txtNavio.Text = ""
        ddlTipoConhec.SelectedValue = 10
        txtDataAverba.Text = ""
        TxtNumAtoConcessorio.Text = ""
        TxtDtaRegAtoConcessorio.Text = ""
        TxtDtaValidAtoConcessorio.Text = ""

        GridConhecimento.DataSource = Nothing
        GridConhecimento.DataBind()
        gridRE.DataSource = Nothing
        gridRE.DataBind()

        grdNotasReferenciais.DataSource = Nothing
        grdNotasReferenciais.DataBind()

    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            btnEmpresa.Enabled = False
        End If

        SessaoRecuperaNotaFiscal()

        If objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica AndAlso objNotaFiscal.Empresa.Empresa.NossaEmissao Then
            Dim fileEmpresa As String = "C:\NGS\xmlCopy\" & objNotaFiscal.Empresa.Codigo & ".xml"

            If Not File.Exists(fileEmpresa) Then
                MsgBox(Me.Page, "Arquivo de Configuracão da Empresa não foi encontrado.", "~/Expedicao.aspx")
                Exit Sub
            End If

            Dim xmlConf As New XmlDocument
            xmlConf.Load(fileEmpresa)

            Session("objHorarioServidor" & HID.Value) = xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("horarioServidor").InnerText
        End If
    End Sub

    Public Sub AjustarEncargos(ByVal pValorAjusteEncargo As Decimal)
        SessaoRecuperaNotaFiscal()
        For Each item In objNotaFiscal.Itens
            For Each enc In item.Encargos
                If enc.Codigo = "FACS" Then
                    enc.Valor -= pValorAjusteEncargo
                End If
            Next
        Next
        SessaoSalvaNotaFiscal()
        objNotaFiscal.AtualizaTotais()
        'AtualizaFormulario()
    End Sub
#End Region

#Region "Úteis"
    Private Sub SelecionarIndiceCombo(ByVal Combo As DropDownList, ByVal Valor As String)
        If Valor = "0" Then
            Combo.SelectedIndex = 0
        Else
            Combo.SelectedIndex = Combo.Items.IndexOf(Combo.Items.FindByValue(Valor))
        End If
    End Sub

    Private Sub MostrarTelaErro(ByVal Erro As Exception)
        Session("ssMessage" & HID.Value) = Erro.Message & vbCrLf & vbCrLf & Erro.StackTrace
        MsgBox(Me.Page, Session("ssMessage" & HID.Value))
    End Sub

    Public Sub MensagemControle()
        If Not Session("MsgControle" & HID.Value) Is Nothing Then
            MsgBox(Me.Page, Session("MsgControle" & HID.Value).ToString)
            Session.Remove("MsgControle" & HID.Value)
        End If
    End Sub
#End Region

#Region "Sessão"
    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub SessaoSalvaListaDeNotasFiscaisReferenciaisAnteriores()
        Session("objListaDeNotasFiscaisReferenciaisAnteriores" & HID.Value) = objListaDeNotasFiscaisReferenciaisAnteriores
    End Sub

    Private Sub SessaoRecuperaListaDeNotasFiscaisReferenciaisAnteriores()
        objListaDeNotasFiscaisReferenciaisAnteriores = Session("ListaDeNotasFiscaisReferenciaisAnteriores" & HID.Value)
    End Sub

    Private Sub SessaoSalvaLoteEmbalagem()
        Session("objListaDeCombinacoesLoteEmbalagem" + HID.Value) = objListaDeCombinacoesLoteEmbalagem
    End Sub

    Private Sub SessaoRecuperaLoteEmbalagem()
        objListaDeCombinacoesLoteEmbalagem = CType(Session("objListaDeCombinacoesLoteEmbalagem" + HID.Value), [Lib].Negocio.ListSaldoProdutoEstoqueLoteEmbalagem)
    End Sub

    'Private Sub SessaoVerificaCarregaObjetos()
    '    CarregarNotasDevolucao()
    'End Sub

    Public Property SessaoDsXML() As DataSet
        Get
            Return CType(Session("dsXML" & HID.Value), DataSet)
        End Get
        Set(ByVal value As DataSet)
            If value Is Nothing Then
                Session.Remove("dsXml" & HID.Value)
            Else
                Session("dsXml" & HID.Value) = value
            End If
        End Set
    End Property

    Public Sub CarregarVencimentos()
        If Not Session("Vencimentos" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()
            'gridVencimentosNota.DataSource = objNotaFiscal.VencimentosNota
            'Ajuste para Selecinar o Título e carregar no Financeiro - Furlan 30/07/2025
            gridVencimentosNota.DataSource = From tit In objNotaFiscal.VencimentosNota
                                             Select New With {.Codigo = tit.Codigo, .ReceberPagar = tit.ReceberPagar, .DescricaoMoeda = tit.DescricaoMoeda, .DescricaoProvisao = tit.DescricaoProvisao, .Prorrogacao = tit.Prorrogacao, .ValorDoDocumento = tit.ValorDoDocumento, .MoedaValorDoDocumento = tit.MoedaValorDoDocumento, .CodigoCifrado = IIf(tit.ReceberPagar = "P", Funcoes.Cifrar("ContasAPagar-" & tit.Codigo), Funcoes.Cifrar("ContasAReceber-" & tit.Codigo))}
            gridVencimentosNota.DataBind()

            'gridVencimentosPedido.DataSource = objNotaFiscal.VencimentosPedido
            'Ajuste para Selecinar o Título e carregar no Financeiro - Furlan 30/07/2025
            gridVencimentosPedido.DataSource = From tit In objNotaFiscal.VencimentosPedido
                                               Select New With {.Codigo = tit.Codigo, .ReceberPagar = tit.ReceberPagar, .DescricaoMoeda = tit.DescricaoMoeda, .DescricaoProvisao = tit.DescricaoProvisao, .Prorrogacao = tit.Prorrogacao, .ValorDoDocumento = tit.ValorDoDocumento, .MoedaValorDoDocumento = tit.MoedaValorDoDocumento, .CodigoCifrado = IIf(tit.ReceberPagar = "P", Funcoes.Cifrar("ContasAPagar-" & tit.Codigo), Funcoes.Cifrar("ContasAReceber-" & tit.Codigo))}
            gridVencimentosPedido.DataBind()

            Session.Remove("Vencimentos" & HID.Value)
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Session("objEmpresaNXI" & HID.Value) IsNot Nothing Then
                SetarEmpresa(CType(Session("objEmpresaNXI" & HID.Value), [Lib].Negocio.Cliente).Codigo, CType(Session("objEmpresaNXI" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco, True)
                Session.Remove("objEmpresaNXI" & HID.Value)
                ListarClientes()
            ElseIf Session("objClienteNXI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objClienteNXI" & HID.Value), [Lib].Negocio.Cliente)
                objNotaFiscal.Cliente = pCliente
                'If objNotaFiscal.Cliente.FisicaJuridica = "J" Then
                '    'chk_nfe.Checked = True
                '    chk_nfe_CheckedChanged(chk_nfe, New EventArgs)
                'End If
                If Not ValidarCertidaoClienteNota() Then
                    Session.Remove("objClienteNXI" & HID.Value)
                    Exit Sub
                End If
                With objNotaFiscal.Cliente
                    objNotaFiscal.CodigoCliente = .Codigo
                    objNotaFiscal.EnderecoCliente = .CodigoEndereco
                    SetarCliente(objNotaFiscal.Cliente)
                End With

                SessaoSalvaNotaFiscal()

                If Not LancarSaldoInicial Then EnviarNFePendentes()

                Session.Remove("objClienteNXI" & HID.Value)

                ConsultaPedidos(True)

            ElseIf Session("objTransportadorNXI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                objNotaFiscal.Transportador = CType(Session("objTransportadorNXI" & HID.Value), [Lib].Negocio.Cliente)
                Dim itemTransportador As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Transportador)
                With objNotaFiscal.Transportador
                    objNotaFiscal.CodigoTransportador = .Codigo
                    objNotaFiscal.EnderecoTransportador = .CodigoEndereco
                    txtNomeDoTransportador.Text = .Nome
                    txtCodigoTransportador.Value = itemTransportador.Value
                    txtCnpjDoTransportador.Text = .CodigoFormatado
                    txtEnderecoDoTransportador.Text = .Endereco & ", " & .Numero.ToString()
                    txtCidadeDoTransportador.Text = .Cidade
                    txtEstadoDoTransportador.Text = .CodigoEstado
                    txtInscricaoDoTransportador.Text = .InscricaoEstadual
                End With
                Session.Remove("objTransportadorNXI" & HID.Value)
                SessaoSalvaNotaFiscal()
                BuscarPlaca()
            ElseIf Session("objDepositoNXI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                If (objNotaFiscal.CodigoEmpresa = objNotaFiscal.CodigoCliente And
                  objNotaFiscal.EnderecoEmpresa = objNotaFiscal.EnderecoCliente) AndAlso
                  objNotaFiscal.CodigoCliente <> CType(Session("objClienteNxID" & HID.Value), [Lib].Negocio.Cliente).Codigo Then
                    Session.Remove("objClienteNxID" & HID.Value)
                    MsgBox(Me.Page, "Depósito não pode ser diferente do Código do Cliente")
                    Exit Sub
                End If
                objNotaFiscal.Deposito = CType(Session("objClienteNxID" & HID.Value), [Lib].Negocio.Cliente)
                objNotaFiscal.CodigoDeposito = objNotaFiscal.Deposito.Codigo
                objNotaFiscal.EnderecoDeposito = objNotaFiscal.Deposito.CodigoEndereco
                Dim itemDeposito As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Deposito)
                With objNotaFiscal.Deposito
                    objNotaFiscal.CodigoDeposito = .Codigo
                    objNotaFiscal.EnderecoDeposito = .CodigoEndereco
                    Dim intIndice As Integer = ddlDeposito.Items.IndexOf(ddlDeposito.Items.FindByValue(.Codigo & "-" & CStr(.CodigoEndereco)))
                    If intIndice = -1 Then
                        Funcoes.AdicionarClienteAoDDL(ddlDeposito, objNotaFiscal.Deposito)
                    End If
                End With
                ddlDeposito.SelectedValue = objNotaFiscal.CodigoDeposito & "-" & objNotaFiscal.EnderecoDeposito
                For Each row As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
                    row.CodigoDeposito = objNotaFiscal.CodigoDeposito
                    row.EnderecoDeposito = objNotaFiscal.EnderecoDeposito
                Next
                If objNotaFiscal.CodigoRomaneio > 0 Then
                    objNotaFiscal.Romaneio.CodigoDeposito = objNotaFiscal.CodigoDeposito
                    objNotaFiscal.Romaneio.EnderecoDeposito = objNotaFiscal.EnderecoDeposito
                End If
                txtCodigoTransportador.Value = ""
                Session.Remove("objDepositoNXI" & HID.Value)
                SessaoSalvaNotaFiscal()
            ElseIf Session("objEmbarqueNXI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                objNotaFiscal.CodigoLocalEmbarque = CType(Session("objEmbarqueNXI" & HID.Value), [Lib].Negocio.Cliente).Codigo
                objNotaFiscal.EndLocalEmbarque = CType(Session("objEmbarqueNXI" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
                Dim intIndice As Integer = ddlEmbarque.Items.IndexOf(ddlEmbarque.Items.FindByValue(objNotaFiscal.CodigoLocalEmbarque & "-" & CStr(objNotaFiscal.EndLocalEmbarque)))
                If intIndice = -1 Then
                    Funcoes.AdicionarClienteAoDDL(ddlEmbarque, objNotaFiscal.LocalEmbarque)
                End If

                objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|Local de Embarque: " & CType(Session("objEmbarqueNXI" & HID.Value), [Lib].Negocio.Cliente).CodigoFormatado & " " & CType(Session("objEmbarqueNXI" & HID.Value), [Lib].Negocio.Cliente).Nome & " - " & CType(Session("objEmbarqueNXI" & HID.Value), [Lib].Negocio.Cliente).Cidade & "/" & CType(Session("objEmbarqueNXI" & HID.Value), [Lib].Negocio.Cliente).CodigoEstado
                If String.IsNullOrWhiteSpace(observacaoAutorizacao.Value) Then
                    txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
                Else
                    txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
                End If

                ddlEmbarque.SelectedValue = objNotaFiscal.CodigoLocalEmbarque & "-" & objNotaFiscal.EndLocalEmbarque
                txtCodigoTransportador.Value = ""
                Session.Remove("objEmbarqueNXI" & HID.Value)
                SessaoSalvaNotaFiscal()
            ElseIf Session("objTransbordoNXI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                objNotaFiscal.Transbordo = CType(Session("objTransbordoNXI" & HID.Value), [Lib].Negocio.Cliente)
                objNotaFiscal.CodigoTransbordo = objNotaFiscal.Transbordo.Codigo
                objNotaFiscal.EnderecoTransbordo = objNotaFiscal.Transbordo.CodigoEndereco
                Dim intIndice As Integer = ddlTransbordo.Items.IndexOf(ddlTransbordo.Items.FindByValue(objNotaFiscal.CodigoTransbordo & "-" & CStr(objNotaFiscal.EnderecoTransbordo)))
                If intIndice = -1 Then
                    Funcoes.AdicionarClienteAoDDL(ddlTransbordo, objNotaFiscal.Transbordo)
                End If
                ddlTransbordo.SelectedValue = objNotaFiscal.CodigoTransbordo & "-" & objNotaFiscal.EnderecoTransbordo
                txtCodigoTransportador.Value = ""
                Session.Remove("objTransbordoNXI" & HID.Value)
                SessaoSalvaNotaFiscal()
            ElseIf Session("objOrigemDestinoNXI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                If (objNotaFiscal.CodigoEmpresa = objNotaFiscal.CodigoCliente And
                                      objNotaFiscal.EnderecoEmpresa = objNotaFiscal.EnderecoCliente) AndAlso
                                      objNotaFiscal.CodigoCliente = CType(Session("objOrigemDestinoNXI" & HID.Value), [Lib].Negocio.Cliente).Codigo Then
                    MsgBox(Me.Page, "Origem/Destino não pode ser e mesmo que o Código do Cliente")
                    Session.Remove("objOrigemDestinoNXI" & HID.Value)
                    Exit Sub
                End If
                objNotaFiscal.Destino = CType(Session("objOrigemDestinoNXI" & HID.Value), [Lib].Negocio.Cliente)
                objNotaFiscal.CodigoDestino = objNotaFiscal.Destino.Codigo
                objNotaFiscal.EnderecoDestino = objNotaFiscal.Destino.CodigoEndereco
                Dim intIndice As Integer = ddlEmbarque.Items.IndexOf(ddlEmbarque.Items.FindByValue(objNotaFiscal.CodigoDestino & "-" & CStr(objNotaFiscal.EnderecoDestino)))
                If intIndice = -1 Then
                    Funcoes.AdicionarClienteAoDDL(ddlOrigemDestino, objNotaFiscal.Destino)
                End If
                ddlOrigemDestino.SelectedValue = objNotaFiscal.CodigoDestino & "-" & objNotaFiscal.EnderecoDestino
                If objNotaFiscal.CodigoRomaneio > 0 Then
                    objNotaFiscal.Romaneio.CodigoDestino = objNotaFiscal.CodigoDestino
                    objNotaFiscal.Romaneio.EnderecoDestino = objNotaFiscal.EnderecoDestino
                End If
                txtCodigoTransportador.Value = ""
                Session.Remove("objOrigemDestinoNXI" & HID.Value)
                SessaoSalvaNotaFiscal()
            ElseIf Session("objEntregaNXI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                objNotaFiscal.Entrega = CType(Session("objEntregaNXI" & HID.Value), [Lib].Negocio.Cliente)
                objNotaFiscal.CodigoEntrega = objNotaFiscal.Entrega.Codigo
                objNotaFiscal.EnderecoEntrega = objNotaFiscal.Entrega.CodigoEndereco
                Dim intIndice As Integer = ddlEntrega.Items.IndexOf(ddlEntrega.Items.FindByValue(objNotaFiscal.CodigoEntrega & "-" & CStr(objNotaFiscal.EnderecoEntrega)))
                If intIndice = -1 Then
                    Funcoes.AdicionarClienteAoDDL(ddlEntrega, objNotaFiscal.Entrega)
                End If
                ddlEntrega.SelectedValue = objNotaFiscal.CodigoEntrega & "-" & objNotaFiscal.EnderecoEntrega
                Session.Remove("objEntregaNXI" & HID.Value)
                SessaoSalvaNotaFiscal()
            ElseIf Session("objFabricanteDI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                objNotaFiscal.DadosDaImportacao.CodigoFabricante = CType(Session("objFabricanteDI" & HID.Value), [Lib].Negocio.Cliente).Codigo
                objNotaFiscal.DadosDaImportacao.EndFabricante = CType(Session("objFabricanteDI" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
                lblFabricanteDI.Text = CType(Session("objFabricanteDI" & HID.Value), [Lib].Negocio.Cliente).Nome & " - " & RTrim(CType(Session("objFabricanteDI" & HID.Value), [Lib].Negocio.Cliente).Cidade) & "/" & CType(Session("objFabricanteDI" & HID.Value), [Lib].Negocio.Cliente).CodigoEstado
                txtCodigolblFabricanteDI.Value = CType(Session("objFabricanteDI" & HID.Value), [Lib].Negocio.Cliente).Codigo & "-" & CType(Session("objFabricanteDI" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
                Session.Remove("objFabricanteDI" & HID.Value)
                SessaoSalvaNotaFiscal()
            ElseIf Session("objPlaca" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                objPlaca = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa)
                objNotaFiscal.PlacaTransportador = LTrim(RTrim(objPlaca.Placa01.ToUpper))
                txtCodigoPlaca.Value = objNotaFiscal.PlacaTransportador
                txtPlacas.Text = objPlaca.Placa01.ToUpper
                txtEstadoDaPlaca.Text = objPlaca.EstadoPlaca01
                If objPlaca.Placa02.Length > 0 Then
                    txtPlacas.Text = txtPlacas.Text & " " & objPlaca.Placa02.ToUpper
                End If
                If objPlaca.Placa03.Length > 0 Then
                    txtPlacas.Text = txtPlacas.Text & " " & objPlaca.Placa03.ToUpper
                End If
                If objPlaca.Placa04.Length > 0 Then
                    txtPlacas.Text = txtPlacas.Text & " " & objPlaca.Placa04.ToUpper
                End If

                If Not objPlaca.Motorista Is Nothing AndAlso objPlaca.Motorista.Codigo.Length > 0 AndAlso Not objNotaFiscal.ObservacoesDeEmbarque.Contains("Motorista:") Then
                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|Motorista: " & objPlaca.Motorista.Nome & " - CPF " & objPlaca.Motorista.CodigoFormatado & " - CNH " & objPlaca.Habilitacao & " - Placa " & objPlaca.Placa01 & " " & objPlaca.Placa02 & " " & objPlaca.Placa03 & " " & objPlaca.Placa04
                    'txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque
                    If String.IsNullOrWhiteSpace(observacaoAutorizacao.Value) Then
                        txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
                    Else
                        txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
                    End If
                End If

                Session.Remove("objPlaca" & HID.Value)

                SessaoSalvaNotaFiscal()
            ElseIf Session("objUFDesembarqueDI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                objNotaFiscal.DadosDaImportacao.EstadoDesembarqueImportacao = CType(Session("objUFDesembarqueDI" & HID.Value), [Lib].Negocio.Estado).Codigo
                lblUFDesembarqueDI.Text = CType(Session("objUFDesembarqueDI" & HID.Value), [Lib].Negocio.Estado).Codigo & " - " & CType(Session("objUFDesembarqueDI" & HID.Value), [Lib].Negocio.Estado).Descricao
                txtCodigoUFDesembarqueDI.Value = CType(Session("objUFDesembarqueDI" & HID.Value), [Lib].Negocio.Estado).Codigo
                Session.Remove("objUFDesembarqueDI" & HID.Value)
                SessaoSalvaNotaFiscal()
            ElseIf Session("objUFEmbarqueDI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                objNotaFiscal.DadosDaImportacao.EstadoEmbarqueImportacao = CType(Session("objUFEmbarqueDI" & HID.Value), [Lib].Negocio.Estado).Codigo
                lblUFEmbarqueDI.Text = CType(Session("objUFEmbarqueDI" & HID.Value), [Lib].Negocio.Estado).Codigo & " - " & CType(Session("objUFEmbarqueDI" & HID.Value), [Lib].Negocio.Estado).Descricao
                txtCodigoUFEmbarqueDI.Value = CType(Session("objUFEmbarqueDI" & HID.Value), [Lib].Negocio.Estado).Codigo
                Session.Remove("objUFEmbarqueDI" & HID.Value)
                SessaoSalvaNotaFiscal()
            ElseIf Session("objItensPedidoSelecionadosNXI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                If Not ValidarCertidaoClienteNota() Then
                    Session.Remove("objItensPedidoSelecionadosNXI")
                    'Importacao do xml de notas fiscais de entrada
                    If SessaoDsXML IsNot Nothing Then
                        SessaoDsXML() = Nothing
                        ucFile.Clear()
                    End If

                    Exit Sub
                End If

                objNotaFiscal.DataDevolucao = objNotaFiscal.Movimento

                If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
                    objNotaFiscal.Especie = objNotaFiscal.Pedido.DescricaoEmbalagem
                End If

                'OBRIGATÓRIO O ENVIO DA DANFE E XML QUANDO FOR NOSSA EMISSAO - 14/09/2016 - FURLAN
                'VERIFICA SE EXISTE O MUNICIPIO IBGE - FURLAN - 03/10/2016 - FURLAN
                If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
                    If String.IsNullOrWhiteSpace(objNotaFiscal.Cliente.EmailNFE) Then
                        MsgBox(Me.Page, "Email para Recebimento da NFE é obrigatório quando for NOSSA EMISSÃO, favor preencher no Cadastro do Cliente " & objNotaFiscal.Cliente.CodigoFormatado & "-" & objNotaFiscal.Cliente.Nome & " o campo EMAIL NFE.")
                        Exit Sub
                    ElseIf objNotaFiscal.Cliente.Municipio.CodigoIbge = 0 Then
                        MsgBox(Me.Page, "Código do Município da tabela do IBGE não foi encontrado, favor verificar no Cadastro do Cliente " & objNotaFiscal.Cliente.CodigoFormatado & "-" & objNotaFiscal.Cliente.Nome & " o campo CIDADE e o código do Município informado.")
                        Exit Sub
                    End If
                    trFile.Visible = False
                End If

                'PAGAMENTO/RECEBIMENTO ANTECIPADO O DEVE TER TITULO EM ADIANTAMENTO
                If FinanceiroNovo And objNotaFiscal.SubOperacao.Financeiro Then
                    If objNotaFiscal.IUD = "I" Then
                        ucFinanceiro.CarregarResumo()
                        ucFinanceiro.AtualizarValorNotaOuFixacaoOuTroca()
                    End If

                    If objNotaFiscal.Pedido.CondicaoPagamento IsNot Nothing AndAlso objNotaFiscal.Pedido.CondicaoPagamento.Antecipado AndAlso objNotaFiscal.Titulos.AdiantamentosAbertos.Count = 0 Then
                        Session.Remove("objItensPedidoSelecionadosNXI")
                        MsgBox(Me.Page, "Pagamento/Recebimento antecipado deve ter titulo em adiantamento")
                        Exit Sub
                    End If
                End If

                If objNotaFiscal.SubOperacao.ProprietarioDaMercadoria Then
                    Dim numdep As Integer
                    numdep = objNotaFiscal.Pedido.Depositos.Where(Function(s) s.Tipo = "PM").Count
                    If numdep = 0 Then
                        MsgBox(Me.Page, "Não foi informado o Proprietário da Mercadoria no Pedido.")
                        Exit Sub
                    End If

                    objNotaFiscal.CodigoProprietarioDaMercadoria = objNotaFiscal.Pedido.Depositos.Where(Function(s) s.Tipo = "PM")(0).Codigo
                    objNotaFiscal.EnderecoProprietarioDaMercadoria = objNotaFiscal.Pedido.Depositos.Where(Function(s) s.Tipo = "PM")(0).CodigoEndereco
                End If

                SetarCliente(objNotaFiscal.Cliente)
                txtDataDeEmissao.Enabled = False
                txtDataDeEntrada.Enabled = False
                chk_NossaEmissao.Checked = objNotaFiscal.NossaEmissao
                CarregarDepositosPedido()

                '2013.08.27 - Acrescentadas verificações para permitir diferença de 1 centavo entre os valores dos pedidos para troca'
                If objNotaFiscal.Pedido.Troca Then
                    If [Enum].Parse(GetType(eClassesOperacoes), objNotaFiscal.Pedido.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.COMPRAS Then

                        If Not objNotaFiscal.SubOperacao.Devolucao AndAlso ((objNotaFiscal.Pedido.CodigoMoeda = eTiposMoeda.Oficial AndAlso Math.Abs(objNotaFiscal.Pedido.PedidoTroca.Itens.LiquidoOficial - objNotaFiscal.Pedido.Itens.LiquidoOficial) > 0.01) _
                            OrElse (objNotaFiscal.Pedido.CodigoMoeda = eTiposMoeda.MoedaEstrangeira AndAlso Math.Abs(objNotaFiscal.Pedido.PedidoTroca.Itens.LiquidoOficial - objNotaFiscal.Pedido.Itens.LiquidoOficial) > 0.1)) Then
                            Session.Remove("objItensPedidoSelecionadosNXI" & HID.Value)
                            MsgBox(Me.Page, "Os pedidos de troca estão com valores diferentes!")
                            Exit Sub
                        End If

                    ElseIf objNotaFiscal.Pedido.PedidoTroca Is Nothing Then
                        Session.Remove("objItensPedidoSelecionadosNXI" & HID.Value)
                        MsgBox(Me.Page, "Pedido de Troca não está vinculado, favor verificar!")
                        Exit Sub
                    Else
                        If Not objNotaFiscal.SubOperacao.Devolucao _
                            AndAlso ((objNotaFiscal.Pedido.CodigoMoeda = eTiposMoeda.Oficial AndAlso Math.Abs(objNotaFiscal.Pedido.PedidoTroca.Itens.LiquidoOficial - objNotaFiscal.Pedido.Itens.LiquidoOficial) > 0.01) _
                            OrElse (objNotaFiscal.Pedido.CodigoMoeda = eTiposMoeda.MoedaEstrangeira AndAlso Math.Abs(objNotaFiscal.Pedido.PedidoTroca.Itens.LiquidoOficial - objNotaFiscal.Pedido.Itens.LiquidoOficial) > 0.1)) Then

                            Session.Remove("objItensPedidoSelecionadosNXI" & HID.Value)
                            MsgBox(Me.Page, "Os pedidos de troca estão com valores diferentes!")
                            Exit Sub
                        End If
                    End If
                End If

                If objNotaFiscal.Pedido.Itens(0).Fixacoes.Count > 0 AndAlso objNotaFiscal.Pedido.Itens(0).Fixacoes(0).IndiceFixado > 0 Then
                    txtPedido.Text = objNotaFiscal.Pedido.Codigo & " - Indice Fixado " & objNotaFiscal.Pedido.Itens(0).Fixacoes(0).IndiceFixado
                ElseIf objNotaFiscal.Pedido.IndiceFixado > 0 Then
                    txtPedido.Text = objNotaFiscal.Pedido.Codigo & " - Indice Fixado " & objNotaFiscal.Pedido.IndiceFixado
                Else
                    txtPedido.Text = objNotaFiscal.Pedido.Codigo
                End If

                If objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                    lblCliente.Text = "REMETENTE"
                Else
                    lblCliente.Text = "DESTINATÁRIO"
                End If

                objNotaFiscal.CodigoPedido = objNotaFiscal.Pedido.Codigo

                If Not objNotaFiscal.Pedido.Finalidade Is Nothing Then
                    SelecionarIndiceCombo(cmbFinalidade, objNotaFiscal.Pedido.CodigoFinalidade)
                    objNotaFiscal.CodigoFinalidade = objNotaFiscal.Pedido.CodigoFinalidade
                Else
                    SelecionarIndiceCombo(cmbFinalidade, 1)
                    objNotaFiscal.CodigoFinalidade = 1
                End If

                objNotaFiscal.CodigoSituacao = 1

                erroMsg = String.Empty
                MsgAlerta = String.Empty

                If CarregarItensComSaldo(CType(obj, [Lib].Negocio.SaldoPedido2015), erroMsg, MsgAlerta) Then
                    SessaoRecuperaNotaFiscal()
                Else
                    If erroMsg.Length > 0 Then
                        If String.IsNullOrWhiteSpace(MsgAlerta) Then
                            MsgBox(Me.Page, erroMsg, eTitulo.Erro, False)
                        Else
                            MsgBox(Me.Page, MsgAlerta & erroMsg, eTitulo.Erro, False)
                        End If
                    End If
                    LimparCampos("", "", False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                    Exit Sub
                End If

                If Not String.IsNullOrWhiteSpace(MsgAlerta) Then
                    MsgBox(Me.Page, MsgAlerta, eTitulo.Info, False)
                End If

                ''Enviar E-mail caso tenha chego no estoque mínimo - Furlan - 28/10/2020
                'If CType(Session("objEstoqueMinimo" & HID.Value), DataTable).Rows.Count > 0 Then enviarEstoqueMinimo()

                'Nenhum
                If objNotaFiscal.SubOperacao.Devolucao AndAlso Not objNotaFiscal.Pedido.FreteCIFFOB = eTiposFrete.NEN Then 'Quando é devolução Inverte'
                    objNotaFiscal.CIFFOB = IIf(objNotaFiscal.Pedido.FreteCIFFOB = eTiposFrete.CIF, eTiposFrete.FOB, eTiposFrete.CIF)
                Else 'senão coloca o que estiver no pedido'
                    objNotaFiscal.CIFFOB = objNotaFiscal.Pedido.FreteCIFFOB
                End If

                CarregarDDLFrete(False)

                If objNotaFiscal.Pedido.FreteCIFFOB = eTiposFrete.NEN OrElse Not objNotaFiscal.SubOperacao.QuantidadeFisico Then

                    If Left(objNotaFiscal.CodigoEmpresa, 8) = "05272759" OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.OUTRAS OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.TRANSFERENCIAS Then
                        'SE FOR A 'QUÍMICA LIBERA POIS NÃO USA ESTOQUE FÍSICO FURLAN - 29/09/2022
                    Else
                        ddlFrete.SelectedValue = eTiposFrete.NEN.ToString
                        ddlFrete.Enabled = False
                    End If
                End If

                'INICIA OBSERVACAO NOTA FISCAL AQUI - FURLAN - 25/07-2014
                If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
                    'lnkVisualizar.Parent.Parent.Parent.Visible = True
                    'lkn_ModeloNota.Parent.Visible = True
                    IdVisualizarNFe.Visible = True
                    IdModelNota.Visible = True
                    IdVisualizar.Visible = True

                    If objNotaFiscal.Empresa.Empresa.ObservacaoSefazNFE.Length > 0 Then
                        Dim obsSefaz() As String = objNotaFiscal.Empresa.Empresa.ObservacaoSefazNFE.Split(";")
                        If objNotaFiscal.Empresa.CodigoEstado = "MT" Then
                            If obsSefaz.Length > 1 AndAlso Left(objNotaFiscal.CodigoEmpresa, 8) = "04854422" AndAlso objNotaFiscal.Itens.Count > 0 AndAlso objNotaFiscal.Itens(0).Produto.ControlarEmbalagem Then
                                objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(1)
                            Else
                                objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(0)
                            End If
                        ElseIf objNotaFiscal.Empresa.CodigoEstado = "MS" Then
                            objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(0)
                        ElseIf objNotaFiscal.Empresa.CodigoEstado = "PR" Then
                            If objNotaFiscal.Empresa.Cidade = "UMUARAMA" Or objNotaFiscal.Empresa.Cidade = "CASCAVEL" Then
                                objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(0)
                            Else
                                If obsSefaz.Length > 1 AndAlso (objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10101" OrElse
                                    objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10201" OrElse
                                    objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10202" OrElse
                                    objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10203") Then

                                    objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(1)
                                Else
                                    objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(0)
                                End If
                            End If
                        End If
                    End If

                    'REMOVIDO POR NÃO LEMBRAR O OBJETIVO, VINCULO DE NOTA DE PRODUTOR OBRIGATORIO APENAS PARA ENTRADA - FURLAN - 06/12/2023
                    'OrElse (objNotaFiscal.SubOperacao.Devolucao And objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.VENDAS)
                    If objNotaFiscal.Empresa.Empresa.ObrigaNfProdutor _
                            AndAlso (objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada And (objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.COMPRAS OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.COMPRASGERAIS OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REAJUSTES)) _
                            AndAlso objNotaFiscal.Cliente.CodigoCategoria < 4 Then

                        objNotaFiscal.ObrigaNFProdutor = True

                        TabNotasDeProdutor.Visible = True
                        bntBuscarNFProdutor.Visible = True
                    End If

                End If

                If Not objNotaFiscal.SubOperacao.Devolucao AndAlso objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso objNotaFiscal.IUD = "I" Then
                    bntBuscarNFReferencial.Visible = True
                ElseIf (objNotaFiscal.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE ICMS") Or objNotaFiscal.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE IPI")) AndAlso objNotaFiscal.IUD = "I" Then
                    bntBuscarNFReferencial.Visible = True
                ElseIf objNotaFiscal.ObrigaNFProdutor AndAlso objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REAJUSTES AndAlso objNotaFiscal.IUD = "I" Then
                    bntBuscarNFReferencial.Visible = True
                ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.IMPORTACOES AndAlso objNotaFiscal.IUD = "I" Then
                    bntBuscarNFReferencial.Visible = True
                Else
                    bntBuscarNFReferencial.Visible = False
                End If

                If objNotaFiscal.Itens(0).OperacaoEstado.CodigoBeneficio.Length > 0 Then
                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|" & objNotaFiscal.Itens(0).OperacaoEstado.BeneficioICMS.Descricao
                End If

                If objNotaFiscal.ObservacoesDeEmbarque.Length > 0 Then
                    If Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" Then
                        objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|PEDIDO: " & objNotaFiscal.Pedido.PedidoEfetivo & ". "
                    Else
                        objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|PEDIDO: " & objNotaFiscal.CodigoPedido & ". "
                    End If
                Else
                    If Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" Then
                        objNotaFiscal.ObservacoesDeEmbarque = "|PEDIDO: " & objNotaFiscal.Pedido.PedidoEfetivo & ". "
                    Else
                        objNotaFiscal.ObservacoesDeEmbarque = "|PEDIDO: " & objNotaFiscal.CodigoPedido & ". "
                    End If
                End If

                'POR HORA MANUAL POIS AINDA NÃO TEM O NR DO PROCESSO, DEPOIS VAMOS VER PARA COLOCARMOS EM UMA TABELA PARA PEGAR AUTOMÁTICO - FURLAN - 05-07-2024
                If (Left(objNotaFiscal.CodigoEmpresa, 8) = "44979506" OrElse Left(objNotaFiscal.CodigoEmpresa, 8) = "62780383") AndAlso objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10101" Then
                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "PRODUTO PROVENIENTE DE ARMAZEM PARTICIPANTE CNPJ: " & objNotaFiscal.Empresa.CodigoFormatado & "."
                End If

                'CARREGAR INVOICE DO PEDIDO
                If objNotaFiscal.Pedido.InvoiceNavio > 0 Then
                    Dim objPedidoNxI = New NavioXInvoice(objNotaFiscal.Pedido.InvoiceNavio)
                    objNotaFiscal.InvoiceNavio = objNotaFiscal.Pedido.InvoiceNavio
                    txtNaviosXInvoice.Text = objPedidoNxI.Navio_Id & " - " & objPedidoNxI.Descricao
                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "NOME DO NAVIO: " & objPedidoNxI.Descricao & ". ORIGEM: " & objPedidoNxI.Pais.Descricao & "."
                End If

                If objNotaFiscal.Pedido.PedidoEfetivo.Length > 0 Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & " - CONFIRMACAO DE NEGOCIO: " & objNotaFiscal.Pedido.PedidoEfetivo
                If objNotaFiscal.Pedido.Contrato.Length > 0 Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & " - CONTRATO: " & objNotaFiscal.Pedido.Contrato

                Dim Certidao As New [Lib].Negocio.CertidaoNegativa(objNotaFiscal.CodigoEmpresa, False)
                If Certidao.CodigoCliente.Length > 0 Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "| Certidao Negativa de Debitos nr " & Certidao.Numero & " / Cod. de Autenticidade " & Certidao.CodigoAutenticidade

                If objNotaFiscal.Operacao.CodigoClasse = "VENDAS" AndAlso objNotaFiscal.Pedido.Troca Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "| Pedido com Vencimento em " & objNotaFiscal.Pedido.DataVencimentoPedido

                'TERMINAR COM PARAMETRO DA EMPRESA REF. ICMS
                'If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.IMPORTACOES Then
                '    For Each item In objNotaFiscal.Itens
                '        For Each enc In item.Encargos
                '            If enc.Codigo = "CUSTO ICMS" Then
                '                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "AFRMM").Sum(Function(t) t.Valor)
                '            ElseIf enc.Codigo = "FRETES" Then
                '                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "FRETES").Sum(Function(t) t.Valor)
                '            ElseIf enc.Codigo = "ICMS" Then
                '                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "ICMS").Sum(Function(t) t.Valor)
                '            ElseIf enc.Codigo = "CUSTO ICMS" Then
                '                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "CUSTO ICMS").Sum(Function(t) t.Valor)
                '            End If
                '        Next
                '    Next
                'Else
                'End If
                Dim teveEncrgo As Boolean = False

                For Each item In objNotaFiscal.Itens

                    If objNotaFiscal.CodigoOperacao = objNotaFiscal.Pedido.CodigoOperacao AndAlso objNotaFiscal.CodigoSubOperacao = objNotaFiscal.Pedido.CodigoSubOperacao AndAlso item.QuantidadeFiscal = objNotaFiscal.Pedido.Itens.FirstOrDefault(Function(s) s.CodigoProduto = item.CodigoProduto).QuantidadePedidoFaturamento Then
                        For Each enc In item.Encargos
                            If enc.Codigo = "IPI" Then
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "IPI").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            ElseIf enc.Codigo = "FRETES" Then
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "FRETES").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            ElseIf enc.Codigo = "DESP.ADUANEIRAS" Then
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "DESP.ADUANEIRAS").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            ElseIf enc.Codigo = "DESCONTOS" Then
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "DESCONTOS").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            ElseIf enc.Codigo = "SEGURO" Then
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "SEGURO").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            ElseIf enc.Codigo = "ICMS" Then
                                enc.Base = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "ICMS" OrElse s.CodigoEncargo = "ICMS A REC.").Sum(Function(t) t.BaseOficial)
                                enc.Percentual = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "ICMS" OrElse s.CodigoEncargo = "ICMS A REC.").Sum(Function(t) t.Percentual)
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "ICMS" OrElse s.CodigoEncargo = "ICMS A REC.").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            ElseIf enc.Codigo = "CUSTO ICMS" Then
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "CUSTO ICMS").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            End If
                        Next
                    End If
                Next

                If teveEncrgo Then objNotaFiscal.AtualizaTotais()

                'Verificação para evitar erro de índice quando existir mensagem de Erro no cálculo do valor unitário do produto
                observacaoValores.Value = String.Empty

                If Not objNotaFiscal.SubOperacao.Devolucao AndAlso objNotaFiscal.Itens.Count > 0 Then
                    For Each enc In objNotaFiscal.Itens(0).Encargos
                        If enc.Valor > 0 Then
                            If enc.Codigo.Contains("FUNRURAL") Then
                                observacaoValores.Value = observacaoValores.Value & "|FUNRURAL - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " R$ por Unidade", " % ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".") & " R$"
                            ElseIf enc.Codigo.Contains("FABOV") Then
                                observacaoValores.Value = observacaoValores.Value & "|FABOV - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " R$ por Unidade ", " % ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".") & " R$"
                            ElseIf enc.Codigo.Contains("SENAR") Then
                                observacaoValores.Value = observacaoValores.Value & "|SENAR - " & enc.Percentual.ToString("N2").Replace(",", ".") & " % - R$ " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                            ElseIf enc.Codigo.Contains("FETHAB") Then
                                observacaoValores.Value = observacaoValores.Value & "|FETHAB - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " R$ por Unidade", " % ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".") & " R$"
                            ElseIf enc.Codigo.Contains("IAGRO") Then
                                observacaoValores.Value = observacaoValores.Value & "|IAGRO - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " R$ por Unidade", " % ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".") & " R$"
                            ElseIf enc.Codigo.Contains("FACS") Then
                                observacaoValores.Value = observacaoValores.Value & "|FACS - " & enc.Percentual.ToString.Replace(",", ".") & " % - R$ " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                            ElseIf enc.Codigo.Contains("FUNDEMS") Then
                                observacaoValores.Value = observacaoValores.Value & "|FUNDEMS - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " R$ por Unidade", " % ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".") & " R$"
                            ElseIf enc.Codigo.Contains("FUNDERSUL") Then
                                observacaoValores.Value = observacaoValores.Value & "|FUNDERSUL - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " R$ por Unidade", " % ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".") & " R$"
                            End If

                        End If
                    Next
                End If

                If Not String.IsNullOrWhiteSpace(objNotaFiscal.PlacaTransportador) AndAlso Not objNotaFiscal.ObservacoesDeEmbarque.Contains("Motorista:") Then
                    Dim objPlaca As New [Lib].Negocio.Placa(objNotaFiscal.PlacaTransportador)
                    If Not objPlaca.Motorista Is Nothing Then
                        objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|Motorista: " & objPlaca.Motorista.Nome & " - CPF " & objPlaca.Motorista.CodigoFormatado & " - CNH " & objPlaca.Habilitacao & " - Placa " & objPlaca.Placa01 & " " & objPlaca.Placa02 & " " & objPlaca.Placa03 & " " & objPlaca.Placa04
                    Else
                        objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & " - Placa " & objPlaca.Placa01 & " " & objPlaca.Placa02 & " " & objPlaca.Placa03 & " " & objPlaca.Placa04
                    End If
                End If


                'Código do Resgistro do Ministéri da Agriculcura Tabela Produtos
                If objNotaFiscal.Empresa.Empresa.UsarRegistroMinAgr Then

                    Dim nrRegistro As String = ""
                    Dim separa As String = ""
                    Dim temRegistro As Boolean = False

                    For Each item In objNotaFiscal.Itens

                        If item.Produto.RegistroMinisterioAgricultura.Length > 0 Then
                            nrRegistro = nrRegistro & item.Produto.RegistroMinisterioAgricultura & separa
                            separa = ","
                            temRegistro = True
                        End If
                    Next

                    If temRegistro Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "REGISTRO DO PRODUTO: " & nrRegistro

                End If

                If Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" AndAlso objNotaFiscal.Eletronica AndAlso objNotaFiscal.NossaEmissao AndAlso Left(objNotaFiscal.Itens(0).Produto.CodigoGrupo, 1) = "2" Then
                    Dim origemDaMercadoria As String = String.Empty

                    If objNotaFiscal.Pedido.InvoiceNavio > 0 Then
                        origemDaMercadoria = "IMPORTADO"
                    ElseIf objNotaFiscal.Itens(0).Encargos.EncProduto.SituacaoTributaria > 99 Then
                        origemDaMercadoria = "IMPORTADO"
                    Else
                        origemDaMercadoria = "NACIONAL"
                    End If

                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|PRODUTO " & origemDaMercadoria & ". NATUREZA FISICA DO PRODUTO: " & objNotaFiscal.Itens(0).Produto.EstadoFisico.Descricao & ". "
                    Dim dValidade = objNotaFiscal.Movimento.AddYears(1)
                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|MODO DE APLICACAO PRINCIPAL: VIA SOLO. ARMAZENAR EM LUGAR SECO, COBERTO E AREJADO. DATA DE FABRICACAO: " & objNotaFiscal.Movimento.ToString("dd/MM/yyyy") & " DATA DE VALIDADE: " & dValidade.ToString("dd/MM/yyyy") & ". "

                End If

                'Importação do xml de notas de entrada
                If SessaoDsXML IsNot Nothing Then
                    chk_nfe.Checked = True
                    If Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                        txtChaveNFe.Text = objNotaFiscal.ChaveNFE
                        txtChaveNFe.Enabled = False
                    End If
                End If

                SessaoSalvaNotaFiscal()

                AtualizaNossaEmissaoEletronica()

                AtualizaFormularioComAClasse()

                SessaoRecuperaNotaFiscal()

                Dim i As Integer = 0
                While i < gridItens.Rows.Count

                    If SessaoDsXML IsNot Nothing Then
                        CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled = False
                        CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled = False
                        CType(gridItens.Rows(i).FindControl("txtTotalItem"), TextBox).Enabled = False
                    ElseIf Not objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                        CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled = False
                        CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled = False
                        CType(gridItens.Rows(i).FindControl("txtTotalItem"), TextBox).Enabled = True
                    ElseIf Mid(gridItens.Rows(i).Cells(2).Text, 1, 1) = "$" Or objNotaFiscal.TemNotaTroca Then
                        If objNotaFiscal.NotaDeTroca.Itens(0).QuantidadeFisica <= objNotaFiscal.Itens(0).SaldoPedidoFisico Then
                            CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled = True
                        Else
                            CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled = False
                        End If
                        gridItens.Rows(i).Cells(2).Text = gridItens.Rows(i).Cells(2).Text.Replace("$", "")
                    End If
                    If objNotaFiscal.SubOperacao.Devolucao Then CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled = False

                    'Travar unitário cfe. solicitação e-mail dia 18/11/2021 Jonathan - Furlan 19/11/2021
                    '                                                           Or Left(objNotaFiscal.Empresa.Codigo, 8) = "44979506" _

                    If objNotaFiscal.NossaEmissao AndAlso (Left(objNotaFiscal.Empresa.Codigo, 8) = "05366261" _
                                                           Or Left(objNotaFiscal.Empresa.Codigo, 8) = "38198213" _
                                                           Or Left(objNotaFiscal.Empresa.Codigo, 8) = "62747840" _
                                                           Or Left(objNotaFiscal.Empresa.Codigo, 8) = "62780383" _
                                                           Or Left(objNotaFiscal.Empresa.Codigo, 8) = "63358210" _
                                                           Or Left(objNotaFiscal.Empresa.Codigo, 8) = "40938762" _
                                                           Or Left(objNotaFiscal.Empresa.Codigo, 8) = "44005444" _
                                                           Or Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" _
                                                           Or Left(objNotaFiscal.Empresa.Codigo, 8) = "48984539") Then
                        CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled = False
                    End If

                    i += 1
                End While

                lnkNovo.Parent.Visible = True
                lnkConsultar.Parent.Visible = False

                If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.CONTAEORDEM Then
                    If objNotaFiscal.SubOperacao.Devolucao Then
                        If objNotaFiscal.Itens(0).Produto.Agrupar = "N" AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico Then
                            BuscaClassificacao()
                        Else
                            objNotaFiscal.Itens(0).PesoFiscal = 0
                            objNotaFiscal.Itens(0).QuantidadeFisica = 0
                            objNotaFiscal.Itens(0).QuantidadeFiscal = 0
                            objNotaFiscal.Itens(0).Unitario = 0
                            objNotaFiscal.Itens(0).ValorTotal = 0

                            SessaoSalvaNotaFiscal()

                            AtualizaFormularioComAClasse()
                        End If
                    Else
                        BuscaVendaAOrdem()
                    End If
                ElseIf Not objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso
                    (objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.COMPRASAORDEM Or
                     objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.VENDASAORDEM) Then
                    BuscaNotaTroca()
                ElseIf objNotaFiscal.Itens.Count > 0 AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.Itens(0).Produto.ControlarPesagem Then
                    If TemRomaneio(objNotaFiscal.CodigoPedido) = True Then
                        ListarRomaneios(objNotaFiscal.CodigoPedido)
                    Else
                        If Not BuscaNotaTroca() Then
                            BuscaClassificacao()
                        End If
                    End If
                End If

                If objNotaFiscal.SubOperacao.Devolucao AndAlso Not objNotaFiscal.NotaOrigemImportacaoXML Is Nothing AndAlso objNotaFiscal.NotaOrigemImportacaoXML.Count() > 0 Then

                    Dim iItemNota As Integer = 0

                    For Each ObjItemNF In objNotaFiscal.Itens

                        For Each notaDevolucao In ObjItemNF.NotasDevolucao

                            notaDevolucao.IUD = ""
                            notaDevolucao.QuantidadeDevolucao = 0
                            notaDevolucao.ValorDevolucao = 0

                            ' casa com a NF referenciada no XML
                            For Each notaRefXML As [Lib].Negocio.NotaFiscal In objNotaFiscal.NotaOrigemImportacaoXML _
                                .Where(Function(x) x.CodigoEmpresa = notaDevolucao.Nota.CodigoEmpresa _
                                             And x.CodigoCliente = notaDevolucao.Nota.CodigoCliente _
                                             And x.Codigo = notaDevolucao.NumeroNota)

                                ' 1) marcar intenção de gravação/edição
                                notaDevolucao.IUD = "U"      ' ou "I" se for primeira vez; "U" sinaliza que já existe ajuste

                                ' 2) preencher DEVOLUÇÃO (o que você quer manter na grid)
                                notaDevolucao.QuantidadeDevolucao = ObjItemNF.QuantidadeFiscal
                                notaDevolucao.ValorDevolucao = ObjItemNF.ValorTotal
                                notaDevolucao.ValorLiquidoDevolucao = ObjItemNF.ValorLiquido

                                ' 3) opcional: informação de nota para exibição/cálculo
                                notaDevolucao.UnitarioNota = ObjItemNF.Unitario

                                ' 4) NÃO é obrigatório, mas você pode espelhar “já devolvido”
                                '    se quiser mostrar acumulados históricos no cabeçalho
                                notaDevolucao.QuantidadeDevolvido = ObjItemNF.QuantidadeFiscal
                                notaDevolucao.ValorDevolvido = ObjItemNF.ValorTotal

                                ' 5) garantir CHAVE COMPLETA para a reabertura bater 1:1
                                notaDevolucao.EntradaSaida = notaDevolucao.EntradaSaida   ' já preenchido na criação
                                notaDevolucao.Serie = notaDevolucao.Serie          ' idem
                                notaDevolucao.NumeroNota = notaDevolucao.NumeroNota     ' idem
                                notaDevolucao.Sequencia = ObjItemNF.Sequencia
                                notaDevolucao.CodigoCFOP = ObjItemNF.CFOP

                                ' 6) marcar como origem (ajuste partiu desta devolução)
                                notaDevolucao.NotaOrigem = True

                            Next
                        Next

                        ' consolida por item p/ refletir na UI
                        Dim retornoList = From p In ObjItemNF.NotasDevolucao
                                          Group By p.ItemNota.CodigoProduto
                                           Into Qtde = Sum(p.QuantidadeDevolucao), Valor = Sum(p.ValorDevolucao)
                                          Where Qtde > 0 Or Valor > 0
                                          Select New RetornoNotasDevolucao With {
                                              .indexItem = iItemNota,
                                              .Quantidade = Qtde,
                                              .Valor = Valor
                                          }

                        If retornoList.Any() Then
                            CarregarNotasDevolucao(retornoList.ToList())
                        End If

                        iItemNota += 1
                    Next

                    ' persiste no estado da sessão p/ ser reencontrado ao reabrir
                    SessaoSalvaNotaFiscal()
                End If

            ElseIf Session("objNotaFiscalReferencial" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()

                'Dim msgObs As String = String.Empty
                'Dim msgObsNotas As String = String.Empty

                For Each objNotaFiscalReferencial In CType(obj, [Lib].Negocio.ListNotaFiscalReferencial)
                    objNotaFiscalReferencial.Parent = objNotaFiscal.Itens(0)
                    If Not objNotaFiscal.NotasReferenciais Is Nothing Then
                        objNotaFiscal.NotasReferenciais.Add(objNotaFiscalReferencial)
                    Else
                        objNotaFiscal.NotasReferenciais = New ListNotaFiscalReferencial()
                        objNotaFiscal.NotasReferenciais.Add(objNotaFiscalReferencial)
                    End If
                Next

                'For Each nfRef In objNotaFiscal.NotasReferenciais
                '    msgObsNotas &= CType(New Cliente(nfRef.Cliente_Id, nfRef.EndCliente_Id), [Lib].Negocio.Cliente).Nome & " CNPJ:" & Funcoes.FormatarCpfCnpj(nfRef.Cliente_Id) & " NF " & nfRef.Nota_Id & "-" & nfRef.Serie_Id & ", "
                'Next
                'msgObsNotas = msgObsNotas.Substring(0, msgObsNotas.Length - 2) & "|/exp|"

                'If objNotaFiscal.NotasReferenciais.Count > 1 Then
                '    msgObs = "|exp|MERC. ADQUIRIDA DAS EMPRESAS "
                'Else
                '    msgObs = "|MERC ADQUIRIDA DA EMPRESA "
                'End If



                'If Not objNotaFiscal.ObservacoesDeEmbarque.Contains(msgObs & msgObsNotas) Then
                '    If objNotaFiscal.ObservacoesDeEmbarque.Length > 0 Then
                '        objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & msgObs & msgObsNotas
                '    Else
                '        objNotaFiscal.ObservacoesDeEmbarque = msgObs & msgObsNotas
                '    End If
                'End If

                CompoeObsDasNotasFiscaisReferenciais()

                'If Not String.IsNullOrWhiteSpace(objNotaFiscal.ObservacoesDeEmbarque) Then
                '    txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque
                'End If
                If String.IsNullOrWhiteSpace(observacaoAutorizacao.Value) Then
                    txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
                Else
                    txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
                End If

                grdNotasReferenciais.DataSource = objNotaFiscal.NotasReferenciais
                grdNotasReferenciais.DataBind()

                Session.Remove("objNotaFiscalReferencial" & HID.Value)
                SessaoSalvaNotaFiscal()

            ElseIf Session("objNFReferencialSaida" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()

                Dim obsReferencial As String = String.Empty
                Dim obsVirgula As String = String.Empty

                If CType(obj, [Lib].Negocio.ListNotaFiscalReferencial).Count = 1 Then
                    obsReferencial = "|COMPLEMENTO REF. NOTA FISCAL "
                Else
                    obsReferencial = "|COMPLEMENTO REF. NOTAS FISCAIS "
                End If


                Dim dtNFReferencial As New DataTable("Item")
                dtNFReferencial.Columns.Add("Nota_id", Type.GetType("System.String"))
                dtNFReferencial.Columns.Add("TipoDeDocumento", Type.GetType("System.String"))
                dtNFReferencial.Columns.Add("Empresa_Id", Type.GetType("System.String"))
                dtNFReferencial.Columns.Add("EndEmpresa_Id", Type.GetType("System.String"))
                dtNFReferencial.Columns.Add("Cliente_Id", Type.GetType("System.String"))
                dtNFReferencial.Columns.Add("EndCliente_Id", Type.GetType("System.String"))
                dtNFReferencial.Columns.Add("ClienteNome", Type.GetType("System.String"))
                dtNFReferencial.Columns.Add("EntradaSaida_Id", Type.GetType("System.String"))
                dtNFReferencial.Columns.Add("Serie_Id", Type.GetType("System.String"))
                dtNFReferencial.Columns.Add("Produto_Id", Type.GetType("System.String"))
                dtNFReferencial.Columns.Add("CFOP_Id", Type.GetType("System.String"))
                dtNFReferencial.Columns.Add("Sequencia_Id", Type.GetType("System.String"))
                dtNFReferencial.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
                dtNFReferencial.Columns.Add("Valor", Type.GetType("System.Decimal"))
                dtNFReferencial.Columns.Add("TipoReferencial_Id", Type.GetType("System.String"))


                For Each objNotaFiscalReferencial In CType(obj, [Lib].Negocio.ListNotaFiscalReferencial)
                    objNotaFiscalReferencial.Parent = objNotaFiscal.Itens(0)

                    objNotaFiscalReferencial.EmpresaReferencial_Id = objNotaFiscal.CodigoEmpresa
                    objNotaFiscalReferencial.EndEmpresaReferencial_Id = objNotaFiscal.EnderecoEmpresa
                    objNotaFiscalReferencial.ClienteReferencial_Id = objNotaFiscal.CodigoCliente
                    objNotaFiscalReferencial.EndClienteReferencial_Id = objNotaFiscal.EnderecoCliente
                    'objNotaFiscalReferencial.EntradaSaidaReferencial_Id = eEntradaSaida.Saida
                    objNotaFiscalReferencial.EntradaSaidaReferencial_Id = objNotaFiscal.EntradaSaida
                    objNotaFiscalReferencial.SerieReferencial_Id = objNotaFiscal.Serie
                    objNotaFiscalReferencial.NotaReferencial_Id = objNotaFiscal.Codigo
                    objNotaFiscalReferencial.CFOPReferencial_Id = objNotaFiscal.Itens(0).CFOP
                    objNotaFiscalReferencial.ProdutoReferencial_Id = objNotaFiscal.Itens(0).CodigoProduto
                    objNotaFiscalReferencial.SequenciaReferencial_Id = objNotaFiscal.Itens(0).Sequencia

                    If Not objNotaFiscal.NotasReferenciais Is Nothing Then
                        objNotaFiscal.NotasReferenciais.Add(objNotaFiscalReferencial)
                    Else
                        objNotaFiscal.NotasReferenciais = New ListNotaFiscalReferencial()
                        objNotaFiscal.NotasReferenciais.Add(objNotaFiscalReferencial)
                    End If

                    obsReferencial += obsVirgula + objNotaFiscalReferencial.Nota_Id

                    obsVirgula = ", "

                    Dim drItem As DataRow = dtNFReferencial.NewRow()

                    drItem("Nota_id") = objNotaFiscalReferencial.Nota_Id
                    drItem("TipoDeDocumento") = eTipoDeDocumento.Nota
                    drItem("Empresa_Id") = objNotaFiscalReferencial.Empresa_Id
                    drItem("EndEmpresa_Id") = objNotaFiscalReferencial.EndEmpresa_Id
                    drItem("Cliente_Id") = objNotaFiscalReferencial.Cliente_Id
                    drItem("EndCliente_Id") = objNotaFiscalReferencial.EndCliente_Id
                    drItem("ClienteNome") = objNotaFiscal.Cliente.Nome
                    drItem("EntradaSaida_Id") = IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                    drItem("Serie_Id") = objNotaFiscalReferencial.Serie_Id
                    drItem("Produto_Id") = objNotaFiscalReferencial.Produto_Id
                    drItem("CFOP_Id") = objNotaFiscalReferencial.CFOP_Id
                    drItem("Sequencia_Id") = objNotaFiscalReferencial.Sequencia_Id
                    drItem("Quantidade") = objNotaFiscalReferencial.Quantidade
                    drItem("Valor") = objNotaFiscalReferencial.Valor
                    drItem("TipoReferencial_Id") = "NFC"

                    dtNFReferencial.Rows.Add(drItem)
                Next

                objNotaFiscal.ObservacoesDeEmbarque += obsReferencial

                'If Not String.IsNullOrWhiteSpace(objNotaFiscal.ObservacoesDeEmbarque) Then
                '    txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque
                'End If
                If String.IsNullOrWhiteSpace(observacaoAutorizacao.Value) Then
                    txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
                Else
                    txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
                End If

                grdNotasReferenciadas.DataSource = dtNFReferencial
                grdNotasReferenciadas.DataBind()

                objNotaFiscal.ModoComplemento = True

                SessaoSalvaNotaFiscal()

                Session.Remove("objNFReferencialSaida" & HID.Value)

                bntBuscarNFReferencial.Visible = False

            ElseIf Session("objNFDeProdutor" & HID.Value) IsNot Nothing Then

                SessaoRecuperaNotaFiscal()

                If objNotaFiscal.NotaTrocaOrigem Is Nothing OrElse objNotaFiscal.NotaTrocaOrigem.Itens.Count = 0 AndAlso objNotaFiscal.NotaTrocaOrigem.Itens(0).ValorTotal = 0 Then
                    MsgBox(Me.Page, "Nota Fiscal do Produtor não foi selecionada, faça a busca novamente!", eTitulo.Info)
                Else

                    Dim dtNFProdutor As New DataTable("Item")
                    dtNFProdutor.Columns.Add("Data", Type.GetType("System.String"))
                    dtNFProdutor.Columns.Add("Operacao", Type.GetType("System.String"))
                    dtNFProdutor.Columns.Add("Serie", Type.GetType("System.String"))
                    dtNFProdutor.Columns.Add("Nota", Type.GetType("System.String"))
                    dtNFProdutor.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
                    dtNFProdutor.Columns.Add("Valor", Type.GetType("System.Decimal"))

                    Dim drItem As DataRow = dtNFProdutor.NewRow()

                    drItem("Data") = objNotaFiscal.NotaTrocaOrigem.Movimento.ToString("dd/MM/yyyy")
                    drItem("Operacao") = objNotaFiscal.NotaTrocaOrigem.CodigoOperacao & "-" & objNotaFiscal.NotaTrocaOrigem.CodigoSubOperacao & " - " & objNotaFiscal.NotaTrocaOrigem.SubOperacao.Descricao
                    drItem("Serie") = objNotaFiscal.NotaTrocaOrigem.Serie
                    drItem("Nota") = objNotaFiscal.NotaTrocaOrigem.Codigo
                    If objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                        drItem("Quantidade") = objNotaFiscal.NotaTrocaOrigem.Itens(0).QuantidadeFiscal
                    Else
                        drItem("Quantidade") = 0
                    End If
                    drItem("Valor") = objNotaFiscal.NotaTrocaOrigem.Itens(0).ValorTotal
                    dtNFProdutor.Rows.Add(drItem)

                    gridNotasDeProdutor.DataSource = dtNFProdutor
                    gridNotasDeProdutor.DataBind()

                    If objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                        objNotaFiscal.Quantidade = objNotaFiscal.NotaTrocaOrigem.Itens(0).QuantidadeFiscal

                        objNotaFiscal.Itens(0).QuantidadeFiscal = objNotaFiscal.NotaTrocaOrigem.Itens(0).QuantidadeFiscal
                        objNotaFiscal.Itens(0).Unitario = objNotaFiscal.NotaTrocaOrigem.Itens(0).Unitario
                        objNotaFiscal.Itens(0).ValorTotal = objNotaFiscal.NotaTrocaOrigem.Itens(0).ValorTotal
                    Else
                        objNotaFiscal.Itens(0).Unitario = 0
                        objNotaFiscal.Itens(0).QuantidadeFiscal = 0
                        objNotaFiscal.Quantidade = 0
                    End If


                    CType(gridItens.Rows(0).FindControl("txtQuantidadeItem"), TextBox).Text = objNotaFiscal.Itens(0).QuantidadeFiscal.ToString("N4")
                    CType(gridItens.Rows(0).FindControl("txtUnitarioItem"), TextBox).Text = objNotaFiscal.Itens(0).Unitario.ToString("N10")
                    CType(gridItens.Rows(0).FindControl("txtTotalItem"), TextBox).Text = objNotaFiscal.Itens(0).ValorTotal.ToString("N2")

                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & ". NOTA DO PRODUTOR NR. " & objNotaFiscal.NotaTrocaOrigem.Codigo

                    SessaoSalvaNotaFiscal()

                    ItemNotaOK(0)

                    CType(gridItens.Rows(0).FindControl("txtQuantidadeItem"), TextBox).Enabled = False
                    CType(gridItens.Rows(0).FindControl("txtUnitarioItem"), TextBox).Enabled = False
                    CType(gridItens.Rows(0).FindControl("txtTotalItem"), TextBox).Enabled = False

                    bntBuscarNFProdutor.Visible = False
                End If

                Session.Remove("objNFDeProdutor" & HID.Value)


            ElseIf Session("objEmailNFe" & HID.Value) IsNot Nothing Then

                Try
                    SessaoRecuperaNotaFiscal()
                    objNotaFiscal = New [Lib].Negocio.NotaFiscal(objNotaFiscal)
                    Dim lst As [Lib].Negocio.ListCliente = CType(Session("objEmailNFe" & HID.Value), [Lib].Negocio.ListCliente)
                    Dim fm As New FilesManager()
                    If fm.IsConnect() Then
                        Dim Sqls As New ArrayList
                        Dim objFil As New [Lib].Negocio.Fil()
                        objFil.IUD = "I"
                        objFil.NomeArquivo = String.Format("email{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa)
                        objFil.Texto = getTextoEmail(objNotaFiscal, lst, Session("strAssunto" & HID.Value), Session("strMensagem" & HID.Value), CBool(Session("strCompactado" & HID.Value)))
                        objFil.SalvarSql(Sqls)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        End If

                        Dim resp As [Lib].Negocio.Resp = Nothing
                        Dim fileName As String = String.Format("resp-email{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa)

                        Dim tempoLimite As DateTime
                        tempoLimite = Now.AddSeconds(90)

                        While resp Is Nothing AndAlso Now < tempoLimite
                            resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                            System.Threading.Thread.Sleep(3000)
                        End While

                        If resp IsNot Nothing Then
                            Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                            Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                            Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                            Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                            If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "4014" Then
                                MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                                Exit Sub
                            Else
                                MsgBox(Me.Page, strMsg)
                            End If

                            Sqls.Clear()
                            Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-email{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                            Sqls.Add(Sql)

                            If Not Banco.GravaBanco(Sqls) Then
                                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                            End If
                        Else
                            MsgBox(Me.Page, "Falha no retorno da Sefaz, tente o reenvio do Email novamente.")
                        End If
                    Else
                        MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
                    End If

                    Session.Remove("objEmailNFe" & HID.Value)

                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
            ElseIf Session("objNavioXInvoice" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                Dim objNavioXInvoice = CType(Session("objNavioXInvoice" & HID.Value), [Lib].Negocio.NavioXInvoice)
                objNotaFiscal.InvoiceNavio = objNavioXInvoice.Codigo
                txtNaviosXInvoice.Text = objNavioXInvoice.Navio_Id & " - " & objNavioXInvoice.Descricao

                objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "NOME DO NAVIO: " & objNavioXInvoice.Descricao & ". ORIGEM: " & objNavioXInvoice.Pais.Descricao & "."

                SessaoSalvaNotaFiscal()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CompoeObsDasNotasFiscaisReferenciais()
        SessaoRecuperaNotaFiscal()

        Dim msgObsNotasReferenciais As String = String.Empty
        Dim temp As String = String.Empty

        Dim inicio As Integer = objNotaFiscal.ObservacoesDeEmbarque.IndexOf("|exp|")
        If inicio > 0 Then
            Dim final As Integer = objNotaFiscal.ObservacoesDeEmbarque.IndexOf("|/exp|")
            objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque.Remove(inicio, final - inicio + 6)
        End If

        Dim cnpj As String = String.Empty

        Dim ListaDeNotasReferenciaisOrdenadasPeloCnpj = objNotaFiscal.NotasReferenciais 'As New ListNotaFiscalReferencial()
        ListaDeNotasReferenciaisOrdenadasPeloCnpj.Sort(Function(nfr1 As NotaFiscalReferencial, nfr2 As NotaFiscalReferencial) nfr1.Cliente_Id.CompareTo(nfr2.Cliente_Id))

        For Each nfRef In ListaDeNotasReferenciaisOrdenadasPeloCnpj
            If Not cnpj.Equals(nfRef.Cliente_Id) Then
                cnpj = nfRef.Cliente_Id
                msgObsNotasReferenciais &= CType(New Cliente(nfRef.Cliente_Id, nfRef.EndCliente_Id), [Lib].Negocio.Cliente).Nome & " CNPJ:" & Funcoes.FormatarCpfCnpj(nfRef.Cliente_Id) & " NF " & nfRef.Nota_Id & "-" & nfRef.Serie_Id & ", "
            Else
                msgObsNotasReferenciais &= nfRef.Nota_Id & "-" & nfRef.Serie_Id & ", "
            End If
        Next

        If Not String.IsNullOrWhiteSpace(msgObsNotasReferenciais) Then
            msgObsNotasReferenciais = msgObsNotasReferenciais.Substring(0, msgObsNotasReferenciais.Length - 2) & "|/exp|"
            If objNotaFiscal.NotasReferenciais.Count > 1 Then
                msgObsNotasReferenciais = " |exp| MERC. ADQUIRIDA DAS EMPRESAS " & msgObsNotasReferenciais
            Else
                msgObsNotasReferenciais = " |exp| MERC ADQUIRIDA DA EMPRESA " & msgObsNotasReferenciais
            End If
        End If

        If Not objNotaFiscal.ObservacoesDeEmbarque.Contains(msgObsNotasReferenciais) Then
            If objNotaFiscal.ObservacoesDeEmbarque.Length > 0 Then
                objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & msgObsNotasReferenciais
            Else
                objNotaFiscal.ObservacoesDeEmbarque = msgObsNotasReferenciais
            End If
        End If
        'txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque
        If String.IsNullOrWhiteSpace(observacaoAutorizacao.Value) Then
            txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
        Else
            txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
        End If

        SessaoSalvaNotaFiscal()
    End Sub

    Public Function ValidarCertidaoClienteNota() As Boolean
        If objNotaFiscal.Empresa.Empresa.CertidaoNegativa And objNotaFiscal.Empresa.CodigoEstado = objNotaFiscal.Cliente.CodigoEstado Then
            Dim encontrou As Boolean = False
            For Each row As [Lib].Negocio.ClientexTipo In objNotaFiscal.Cliente.Tipos
                If row.CodigoTipo = 4 Or row.CodigoTipo = 5 Then
                    encontrou = True
                    Exit For
                End If
            Next

            Dim strMsg As String = ""
            Dim cn As [Lib].Negocio.CertidaoNegativa
            Dim DiasParaOVencimento As Integer = 0

            'Cliente
            If encontrou Then
                cn = New [Lib].Negocio.CertidaoNegativa(objNotaFiscal.Cliente.Codigo, True)
                DiasParaOVencimento = DateDiff("d", FormatDateTime(Date.Now, DateFormat.ShortDate), FormatDateTime(cn.DataValidade, DateFormat.ShortDate))
                If cn.CodigoCliente.Length = 0 OrElse DiasParaOVencimento < 0 Then
                    MsgBox(Me.Page, "Cliente sem certidão negativa! ")
                    Return False
                ElseIf DiasParaOVencimento > 0 AndAlso DiasParaOVencimento <= 5 Then
                    strMsg = "Falta(m) " & DiasParaOVencimento & " dia(s) para o Vencimento da Certidão Negativa do Cliente" & vbCrLf
                End If
            End If

            'Empresa
            cn = New [Lib].Negocio.CertidaoNegativa(objNotaFiscal.CodigoEmpresa, False)
            DiasParaOVencimento = DateDiff("d", FormatDateTime(Date.Now, DateFormat.ShortDate), FormatDateTime(cn.DataValidade, DateFormat.ShortDate))
            If cn.CodigoCliente.Length = 0 OrElse DiasParaOVencimento < 0 Then
                strMsg &= IIf(strMsg.Length > 0, ", ", "") & "EMPRESA Sem Certidao Negativa!"
                MsgBox(Me.Page, strMsg)
                Return False
            ElseIf DiasParaOVencimento > 0 AndAlso DiasParaOVencimento <= 2 Then
                strMsg &= IIf(strMsg.Length > 0, ", ", "") & "Falta(m) " & DiasParaOVencimento & " dia(s) para o Vencimento da Certidão Negativa da EMPRESA!" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(strMsg) Then
                MsgBox(Me.Page, strMsg)
            End If

        End If
        Return True
    End Function

    Public Sub SetarCliente(ByVal Cli As Cliente)
        Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Cliente)
        txtNomeDoCliente.Text = Cli.Nome
        txtCodigoCliente.Value = itemCliente.Value
        txtCnpjDoCliente.Text = Cli.CodigoFormatado
        txtEnderecoDoCliente.Text = Cli.Endereco & ", " & Cli.Numero.ToString()
        txtComplementoDoCliente.Text = Cli.Complemento
        txtBairroDoCliente.Text = Cli.Bairro
        txtCepDoCliente.Text = Cli.CEP
        txtCidadeDoCliente.Text = Cli.Cidade
        txtTelefoneDoCliente.Text = Cli.Telefone
        txtEstadoDoCliente.Text = Cli.CodigoEstado
        txtInscricaoDoCliente.Text = Cli.InscricaoEstadual
    End Sub

    Public Sub CarregarConsulta()

        If Not Session("objNFConsultaNXI" & HID.Value) Is Nothing Then

            objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsultaNXI" & HID.Value), NotaFiscal))
            objNotaFiscal.NotaFiscalOriginal = New NotaFiscal(objNotaFiscal)
            objNotaFiscal.IUD = "U"

            If objNotaFiscal.Empresa.Empresa.ControlaDataMovimentoNFG Then txtDataDeEntrada.Enabled = False

            ddlUsuarios.Items.Clear()
            If (objNotaFiscal.CodigoSituacao = 2 Or objNotaFiscal.CodigoSituacao = 9 Or objNotaFiscal.CodigoSituacao = 10) AndAlso (objNotaFiscal.UsuarioInclusao <> objNotaFiscal.UsuarioAlteracao) Then
                ddlUsuarios.Items.Add("Can.- " & objNotaFiscal.UsuarioAlteracao & " " & objNotaFiscal.DataAlteracao.ToString("dd/MM/yyyy"))
                ddlUsuarios.Items.Add("Inc.- " & objNotaFiscal.UsuarioInclusao & " " & objNotaFiscal.DataInclusao.ToString("dd/MM/yyyy"))
            ElseIf objNotaFiscal.UsuarioAlteracao IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.UsuarioAlteracao) AndAlso objNotaFiscal.UsuarioInclusao <> objNotaFiscal.UsuarioAlteracao Then
                ddlUsuarios.Items.Add("Alt.- " & objNotaFiscal.UsuarioAlteracao & " " & objNotaFiscal.DataAlteracao.ToString("dd/MM/yyyy"))
                ddlUsuarios.Items.Add("Inc.- " & objNotaFiscal.UsuarioInclusao & " " & objNotaFiscal.DataInclusao.ToString("dd/MM/yyyy"))
            Else
                ddlUsuarios.Items.Add("Inc.- " & objNotaFiscal.UsuarioInclusao & " " & objNotaFiscal.DataInclusao.ToString("dd/MM/yyyy"))
            End If

            If (objNotaFiscal.CFOP IsNot Nothing AndAlso objNotaFiscal.CFOP.Codigo <> 5353 AndAlso objNotaFiscal.CFOP.Codigo <> 6353 AndAlso objNotaFiscal.Pedido IsNot Nothing) Then

                Dim Parametros As New Hashtable
                Parametros.Add("TipoApuracao", 2)
                Parametros.Add("Empresa", objNotaFiscal.CodigoEmpresa)
                Parametros.Add("EndEmpresa", objNotaFiscal.EnderecoEmpresa)

                If objNotaFiscal.SubOperacao.Devolucao AndAlso Not objNotaFiscal.CodigoEmpresa = objNotaFiscal.CodigoCliente Then
                    Parametros.Add("Cliente", objNotaFiscal.CodigoCliente)
                    Parametros.Add("EndCliente", objNotaFiscal.EnderecoCliente)
                End If

                Parametros.Add("Pedido", objNotaFiscal.CodigoPedido)
                Parametros.Add("Devolucao", IIf(objNotaFiscal.SubOperacao.Devolucao, 1, 0))
                Parametros.Add("SoValor", IIf(Not objNotaFiscal.SubOperacao.QuantidadeFiscal And Not objNotaFiscal.SubOperacao.QuantidadeFisico, 1, 0))
                Parametros.Add("Classe", objNotaFiscal.Operacao.CodigoClasse)

                'Enviada a Nota  para que não seja utilizada na composição dos saldos.
                If objNotaFiscal.IUD = "U" Then
                    Parametros.Add("ExcetoNota", objNotaFiscal.Codigo)
                    Parametros.Add("ExcetoSerie", objNotaFiscal.Serie)
                    Parametros.Add("ExcetoCliente", objNotaFiscal.CodigoCliente)
                    Parametros.Add("ExcetoEndCliente", objNotaFiscal.EnderecoCliente)
                    Parametros.Add("ExcetoEntradaSaida", IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S"))
                End If

                Dim ListaItensPedido As New ListSaldoPedido2015(Parametros)

                'Dim ObjItemNF As [Lib].Negocio.NotaFiscalXItem

                For Each ObjItemNF In objNotaFiscal.Itens
                    Dim row As SaldoPedido2015 = ListaItensPedido.Where(Function(s) s.CodigoProduto = ObjItemNF.CodigoProduto).FirstOrDefault
                    If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL And objNotaFiscal.SubOperacao.Devolucao Then
                        ObjItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalGlobal - row.QtdeEntregueFiscalRemessa
                        ObjItemNF.SaldoPedidoFisico = row.QtdeEntregueFiscalGlobal - row.QtdeEntregueFiscalRemessa
                        ObjItemNF.SaldoValorOficial = row.VlrNotaOficialGlobalBruto - row.VlrNotaOficialRemessaBruto
                        ObjItemNF.SaldoValorMoeda = row.VlrNotaMoedaGlobalBruto - row.VlrNotaMoedaRemessaBruto
                        'GLOBAL NAO É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL And Not objNotaFiscal.SubOperacao.Devolucao Then
                        ObjItemNF.SaldoPedidoFiscal = row.SaldoQtdeGlobalFiscal
                        ObjItemNF.SaldoPedidoFisico = row.SaldoQtdeGlobalFiscal
                        ObjItemNF.SaldoValorOficial = row.SaldoValorOficialGlobalDireto
                        ObjItemNF.SaldoValorMoeda = row.SaldoValorMoedaGlobalDireto
                        'REMESSA É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REMESSAS And objNotaFiscal.SubOperacao.Devolucao Then
                        ObjItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalRemessa
                        ObjItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoRemessa
                        ObjItemNF.SaldoValorOficial = row.VlrNotaOficialRemessaBruto
                        ObjItemNF.SaldoValorMoeda = row.VlrNotaMoedaRemessaBruto
                        'REMESSA NAO É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REMESSAS And Not objNotaFiscal.SubOperacao.Devolucao Then
                        ObjItemNF.SaldoPedidoFiscal = row.SaldoQtdeRemessaFiscal
                        ObjItemNF.SaldoPedidoFisico = row.SaldoQtdeRemessaFisica
                        ObjItemNF.SaldoValorOficial = row.SaldoValorOficialRemessa
                        ObjItemNF.SaldoValorMoeda = row.SaldoValorMoedaRemessa
                        'AFIXAR É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR And objNotaFiscal.SubOperacao.Devolucao Then
                        ObjItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalAFixar - row.QtdeFixacao
                        ObjItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoAFixar - row.QtdeFixacao
                        'Criado um novo campo na StoredProcedure spSaldoPedido para trazer o valor
                        'total da fixação, mas correspondente ao unitário da NF. 
                        'O campo VlrfixacaoNF será utilizado quando for menor  do que o campo VlrfixacaoOficial senão este último será, para que o valor da devolução fique correto.
                        ObjItemNF.SaldoValorOficial = row.VlrNotaOficialAFixarBruto - IIf(row.VlrFixacaoNF < row.VlrFixacaoOficial, row.VlrFixacaoNF, row.VlrFixacaoOficial)
                        ObjItemNF.SaldoValorMoeda = row.VlrNotaMoedaAFixarBruto - row.VlrFixacaoMoeda

                        'Devolução utilizada ou para Reajuste de unitário ou para devolução de valor.
                    ElseIf Not row Is Nothing AndAlso row.Tipo = 1 AndAlso objNotaFiscal.SubOperacao.Devolucao AndAlso Not objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                        ObjItemNF.SaldoPedidoFiscal = Math.Abs(row.SaldoQtdeDiretoFiscal)
                        ObjItemNF.SaldoPedidoFisico = Math.Abs(row.SaldoQtdeDiretoFisica)
                        ObjItemNF.SaldoValorOficial = Math.Abs(row.SaldoValorOficialGlobalDireto)
                        ObjItemNF.SaldoValorMoeda = Math.Abs(row.SaldoValorMoedaGlobalDireto)

                    ElseIf Not row Is Nothing AndAlso row.Tipo = 1 AndAlso objNotaFiscal.SubOperacao.Devolucao Then
                        ObjItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalDireta
                        ObjItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoDireta
                        ObjItemNF.SaldoValorOficial = row.VlrNotaOficialDiretaBruto
                        ObjItemNF.SaldoValorMoeda = row.VlrNotaMoedaDiretaBruto
                    ElseIf Not row Is Nothing AndAlso row.Tipo = 2 AndAlso objNotaFiscal.SubOperacao.Devolucao Then
                        ObjItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalDeposito
                        ObjItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoDeposito
                        ObjItemNF.SaldoValorOficial = row.VlrNotaOficialDepositoBruto
                        ObjItemNF.SaldoValorMoeda = row.VlrNotaMoedaDepositoBruto
                    Else
                        'Verificar ainda pode nao estar contemplando tudo
                        If Not row Is Nothing AndAlso (row.QtdeProgramada = 0 And row.QtdeProgramadaComercializacao > 0) Then
                            ObjItemNF.SaldoPedidoFiscal = row.SaldoQtdeComercializacao
                            ObjItemNF.SaldoPedidoFisico = row.SaldoQtdeComercializacao
                            ObjItemNF.SaldoValorOficial = row.SaldoValorOficialGlobalDireto
                            ObjItemNF.SaldoValorMoeda = row.SaldoValorMoedaGlobalDireto
                        ElseIf Not row Is Nothing Then
                            ObjItemNF.SaldoPedidoFiscal = row.SaldoQtdeDiretoFiscal
                            ObjItemNF.SaldoPedidoFisico = row.SaldoQtdeDiretoFisica
                            ObjItemNF.SaldoValorOficial = row.SaldoValorOficialGlobalDireto
                            ObjItemNF.SaldoValorMoeda = row.SaldoValorMoedaGlobalDireto
                        End If
                    End If
                Next

            End If

            'CARREGAR INVOICE DA NFG
            If objNotaFiscal.InvoiceNavio > 0 Then
                Dim objNxI = New NavioXInvoice(objNotaFiscal.InvoiceNavio)
                txtNaviosXInvoice.Text = "(Invoice " & objNxI.Codigo & ") " & objNxI.Navio_Id & " - " & objNxI.Descricao
            End If

            SessaoSalvaNotaFiscal()

            AtualizaFormularioComAClasse()

            CarregarDDLFrete(False)

            SessaoRecuperaNotaFiscal()

            If (Left(objNotaFiscal.CodigoEmpresa, 8) = "05366261" _
                Or Left(objNotaFiscal.CodigoEmpresa, 8) = "03189063" _
                Or Left(objNotaFiscal.CodigoEmpresa, 8) = "62747840" _
                Or Left(objNotaFiscal.CodigoEmpresa, 8) = "62780383" _
                Or Left(objNotaFiscal.CodigoEmpresa, 8) = "63358210" _
                Or Left(objNotaFiscal.Empresa.Codigo, 8) = "38198213" _
                Or Left(objNotaFiscal.Empresa.Codigo, 8) = "40938762" _
                Or Left(objNotaFiscal.Empresa.Codigo, 8) = "44005444" _
                Or Left(objNotaFiscal.Empresa.Codigo, 8) = "44979506" _
                Or Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" _
                Or Left(objNotaFiscal.Empresa.Codigo, 8) = "48984539") _
                      AndAlso objNotaFiscal.NossaEmissao _
                      AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica _
                      AndAlso objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada _
                      AndAlso (objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.COMPRAS OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.COMPRASGERAIS OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REAJUSTES) _
                      AndAlso objNotaFiscal.Cliente.CodigoCategoria < 4 Then

                If grdNotasReferenciadas.Rows.Count > 0 Then

                    Dim dtNFProdutor As New DataTable("Item")
                    dtNFProdutor.Columns.Add("Data", Type.GetType("System.String"))
                    dtNFProdutor.Columns.Add("Operacao", Type.GetType("System.String"))
                    dtNFProdutor.Columns.Add("Serie", Type.GetType("System.String"))
                    dtNFProdutor.Columns.Add("Nota", Type.GetType("System.String"))
                    dtNFProdutor.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
                    dtNFProdutor.Columns.Add("Valor", Type.GetType("System.Decimal"))

                    Dim k As Integer = 0
                    While k < grdNotasReferenciadas.Rows.Count
                        If Trim(grdNotasReferenciadas.Rows(k).Cells(4).Text()) = objNotaFiscal.CodigoCliente Then

                            Dim nfp As New NotaFiscal
                            nfp.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                            nfp.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                            nfp.CodigoCliente = objNotaFiscal.CodigoCliente
                            nfp.EnderecoCliente = objNotaFiscal.EnderecoCliente
                            nfp.EntradaSaida = objNotaFiscal.EntradaSaida
                            nfp.Serie = grdNotasReferenciadas.Rows(k).Cells(8).Text()
                            nfp.Codigo = grdNotasReferenciadas.Rows(k).Cells(0).Text()

                            nfp = New NotaFiscal(nfp)

                            If nfp.TipoDeDocumento.Codigo = eTipoDeDocumento.NotaDeProdutor OrElse nfp.CodigoFinalidade = 30 Then

                                objNotaFiscal.NotaTrocaOrigem = New NotaFiscal(nfp)

                                objNotaFiscal.ObrigaNFProdutor = True

                                TabNotasDeProdutor.Visible = True

                                Dim drItem As DataRow = dtNFProdutor.NewRow()

                                drItem("Data") = objNotaFiscal.NotaTrocaOrigem.Movimento.ToString("dd/MM/yyyy")
                                drItem("Operacao") = objNotaFiscal.NotaTrocaOrigem.CodigoOperacao & "-" & objNotaFiscal.NotaTrocaOrigem.CodigoSubOperacao & " - " & objNotaFiscal.NotaTrocaOrigem.SubOperacao.Descricao
                                drItem("Serie") = objNotaFiscal.NotaTrocaOrigem.Serie
                                drItem("Nota") = objNotaFiscal.NotaTrocaOrigem.Codigo
                                drItem("Quantidade") = objNotaFiscal.NotaTrocaOrigem.Itens(0).QuantidadeFiscal
                                drItem("Valor") = objNotaFiscal.NotaTrocaOrigem.Itens(0).ValorTotal
                                dtNFProdutor.Rows.Add(drItem)
                            End If
                        End If

                        k += 1
                    End While

                    If dtNFProdutor.Rows.Count > 0 Then
                        gridNotasDeProdutor.DataSource = dtNFProdutor
                        gridNotasDeProdutor.DataBind()

                        If grdNotasReferenciadas.Rows.Count = 1 Then
                            grdNotasReferenciadas.DataSource = Nothing
                            grdNotasReferenciadas.DataBind()
                        End If
                    End If
                End If
            End If

            If Not objNotaFiscal.SubOperacao.Devolucao AndAlso
                objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso
                objNotaFiscal.CodigoSituacao = 4 Then

                If objNotaFiscal.NotasReferenciais Is Nothing OrElse objNotaFiscal.NotasReferenciais.Count = 0 Then
                    objNotaFiscal.NotasReferenciais = New ListNotaFiscalReferencial(objNotaFiscal, eTipoReferencial.NFC)
                End If

                For Each objNotaFiscalReferencial In objNotaFiscal.NotasReferenciais

                    If objNotaFiscalReferencial.EmpresaReferencial_Id = objNotaFiscalReferencial.Empresa_Id AndAlso
                        objNotaFiscalReferencial.ClienteReferencial_Id = objNotaFiscalReferencial.Cliente_Id AndAlso
                        objNotaFiscalReferencial.EndClienteReferencial_Id = objNotaFiscalReferencial.EndCliente_Id Then
                        objNotaFiscal.ModoComplemento = True
                    End If
                Next
            End If

            SessaoSalvaNotaFiscal()

            SessaoRecuperaNotaFiscal()

            If objNotaFiscal.SubOperacao.Financeiro AndAlso objNotaFiscal.CodigoPedido > 0 AndAlso objNotaFiscal.Pedido.FinanceiroNovo Then
                TabVencimentosOld.Visible = Not FinanceiroNovo
                TabVencimentos.Visible = FinanceiroNovo
                If FinanceiroNovo Then ucFinanceiro.CarregarResumo()
            Else
                TabVencimentosOld.Visible = True
                BtnVencimentos.Enabled = False
                TabVencimentos.Visible = False
            End If

            Session.Remove("objNFConsultaNXI" & HID.Value)

            If ucFile.Parent.Visible Then
                ucFile.Bind(objNotaFiscal.Arquivos)
                SessaoSalvaNotaFiscal()
            End If

            'LIBERADO PARA GERAR PDF DE NOTAS ELETRONICAS RECEBIDAS - FURLAN 04/08/2025
            'lnkPdf.Parent.Visible = Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ProtocoloNota) AndAlso objNotaFiscal.Eletronica
            'lnkImpressora.Parent.Visible = Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ProtocoloNota) AndAlso objNotaFiscal.NossaEmissao
            lnkPdf.Parent.Visible = Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) AndAlso objNotaFiscal.Eletronica
            lnkImpressora.Parent.Visible = Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) AndAlso objNotaFiscal.Eletronica

            lnkVerificarChaveNFE.Visible = Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) AndAlso Not objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Arquivos.Count = 0

            lnkEnviarEmail.Parent.Visible = Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) AndAlso
                                            Not String.IsNullOrWhiteSpace(objNotaFiscal.ProtocoloNota) AndAlso
                                            objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica
            lnkEspelho.Parent.Visible = objNotaFiscal IsNot Nothing
            lnkConsultar.Parent.Visible = False

            lnkExcluir.Parent.Visible = True

            If objNotaFiscal.NossaEmissao Then lnkExcluir.Parent.Visible = objNotaFiscal.CodigoSituacao = CInt(eSituacao.Normal)

            lnkAtualizar.Parent.Visible = Not String.IsNullOrWhiteSpace(objNotaFiscal.IUD) AndAlso objNotaFiscal.IUD = "U" AndAlso objNotaFiscal.CodigoSituacao = CInt(eSituacao.Normal)

            If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Eletronica Then lnkAtualizar.Parent.Visible = False

            BtnRecontabilizar.Enabled = objNotaFiscal.SubOperacao.Contabil

            If Not objNotaFiscal.VencimentosNota Is Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                If objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoProvisao = eProvisao.Baixa).ToList.Count > 0 Then
                    lnkExcluir.Parent.Visible = False
                    lnkAtualizar.Parent.Visible = False
                    MsgBox(Me.Page, "Nota fiscal de " & objNotaFiscal.EntradaSaida.ToString() & " com vencimento baixado não pode ser alterada nem excluída!")
                ElseIf objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoSituacao = eSituacao.RemessaBancaria).ToList.Count > 0 Then
                    lnkExcluir.Parent.Visible = False
                    lnkAtualizar.Parent.Visible = False
                    MsgBox(Me.Page, "Não é possível alterar a nota fiscal " & objNotaFiscal.Codigo & ", o financeiro está em Cobrança Bancária.")
                ElseIf objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoSituacao = eSituacao.EndossoTitulo).ToList.Count > 0 Then
                    lnkExcluir.Parent.Visible = False
                    lnkAtualizar.Parent.Visible = False
                    MsgBox(Me.Page, "Não é possível alterar a nota fiscal " & objNotaFiscal.Codigo & ", o financeiro está com Endosso.")
                ElseIf objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.RegistroMestre > 0).ToList.Count > 0 Then
                    lnkExcluir.Parent.Visible = False
                    lnkAtualizar.Parent.Visible = False
                    MsgBox(Me.Page, "Nota fiscal de " & objNotaFiscal.EntradaSaida.ToString() & " com vencimento agrupado não pode ser alterada nem excluída!")
                End If
            End If

            Dim temDevolucao As Boolean = False

            For Each ni In objNotaFiscal.Itens
                If Not ni.NotasDevolucao Is Nothing AndAlso ni.NotasDevolucao.Count > 0 Then temDevolucao = True
            Next

            If temDevolucao Then
                MsgBox(Me.Page, "Não é possível alterar nota fiscal com NOTAS DE DEVOLUÇÃO.")

                If objNotaFiscal.SubOperacao.Devolucao Then
                    lnkExcluir.Parent.Visible = True
                Else
                    lnkExcluir.Parent.Visible = False
                End If

                lnkAtualizar.Parent.Visible = False
            End If

            lnkNovo.Parent.Visible = String.IsNullOrWhiteSpace(objNotaFiscal.IUD) OrElse objNotaFiscal.IUD = "I"

            lnkEnviarSEFAZ.Parent.Visible = Not String.IsNullOrWhiteSpace(objNotaFiscal.IUD) AndAlso objNotaFiscal.IUD <> "I" AndAlso objNotaFiscal.NossaEmissao _
                AndAlso String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) AndAlso String.IsNullOrWhiteSpace(objNotaFiscal.ProtocoloNota) _
                AndAlso (String.IsNullOrWhiteSpace(objNotaFiscal.Retorno) OrElse (Not String.IsNullOrWhiteSpace(objNotaFiscal.Retorno) AndAlso objNotaFiscal.Retorno <> "100" AndAlso objNotaFiscal.Retorno <> "110" AndAlso objNotaFiscal.Retorno <> "302"))

            IdVisualizarNFe.Visible = False
            IdModelNota.Visible = False
            IdVisualizar.Visible = False

            If Session("ssNomeUsuario") = "FURLAN" OrElse Session("ssNomeUsuario") = "KAYNÃ" Then
                'If objNotaFiscal.NossaEmissao Then lnkVisualizar.Parent.Parent.Parent.Visible = True
                If objNotaFiscal.NossaEmissao Then
                    IdVisualizarNFe.Visible = True
                    IdVisualizar.Visible = True
                End If
            Else
                'lnkVisualizar.Parent.Parent.Parent.Visible = IsEnabled(objNotaFiscal) AndAlso objNotaFiscal.NossaEmissao
                IdVisualizarNFe.Visible = IsEnabled(objNotaFiscal) AndAlso objNotaFiscal.NossaEmissao
                IdVisualizar.Visible = IsEnabled(objNotaFiscal) AndAlso objNotaFiscal.NossaEmissao
            End If

            BtnRecontabilizar.Enabled = objNotaFiscal IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.IUD) AndAlso objNotaFiscal.IUD = "U"
            If Not objNotaFiscal.IUD = "U" AndAlso objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
                If Not VerificarModo(objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa, objNotaFiscal.NossaEmissao, False) Then
                    lnkAtualizar.Parent.Visible = False
                    lnkExcluir.Parent.Visible = False
                    btnInutilizar.Enabled = False
                    lnkConsultar.Parent.Visible = False
                End If
            End If

            ddlFrete.Enabled = Not objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.IUD = "U" AndAlso objNotaFiscal.CodigoSituacao = CInt(eSituacao.Normal)

            'Caso tenha MDFE vinculado, Nota não pode ser cancelada antes do cancelamento do mesmo - Furlan - 27/09/2016
            For n As Integer = 0 To grdNotasReferenciadas.Rows.Count - 1
                Dim tipodoc() As String = grdNotasReferenciadas.Rows(n).Cells(1).Text.Split("-")
                If Trim(tipodoc(0)) = "12" Then
                    lnkExcluir.Parent.Visible = False
                    Exit For
                End If
            Next

            If Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento, "NOTAS FISCAIS") Then
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
                BtnRecontabilizar.Enabled = False
                MsgBox(Me.Page, "Movimento Fiscal já fechado para esta data!")
            ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento, "CONTABIL") Then
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
                BtnRecontabilizar.Enabled = False
                MsgBox(Me.Page, "Movimento Contábil já fechado para esta data!")
            ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento, "PRODUCAO") Then
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
                BtnRecontabilizar.Enabled = False
                MsgBox(Me.Page, "Movimento do Estoque já fechado para esta data!")
            End If
        End If
    End Sub

    Private Sub ConsutaNotasReferenciais(ByVal empresa As String, ByVal endempresa As Integer, ByVal cliente As String, ByVal endcliente As Integer, ByVal entradasaida As String, ByVal serie As String, ByVal nota As Integer)
        Dim sql As String
        sql = " Select nr.Nota_Id, nr.Empresa_Id, nr.EndEmpresa_Id, nr.Cliente_Id, nr.EndCliente_Id, nr.EntradaSaida_Id, nr.Serie_Id," & vbCrLf &
              "        nr.Produto_Id, nr.CFOP_Id, nr.Sequencia_Id, nr.Quantidade, nr.Valor, nr.TipoReferencial_Id," & vbCrLf &
              "        CAST(TpDoc.Codigo_Id AS VARCHAR) + '-' + TpDoc.Descricao AS TipoDeDocumento, Cli.Nome AS ClienteNome" & vbCrLf &
              "   from notafiscalreferencial nr" & vbCrLf &
              "  Inner Join	NotasFiscais nf" & vbCrLf &
              "     on nf.Empresa_Id      = nr.EmpresaReferencial_Id" & vbCrLf &
              "    And nf.EndEmpresa_Id   = nr.EndEmpresaReferencial_Id" & vbCrLf &
              "    and nf.Cliente_Id      = nr.ClienteReferencial_Id" & vbCrLf &
              "    and nf.EndCliente_Id   = nr.EndClienteReferencial_Id" & vbCrLf &
              "    and nf.EntradaSaida_Id = nr.EntradaSaidaReferencial_Id" & vbCrLf &
              "    and nf.Serie_Id        = nr.SerieReferencial_Id" & vbCrLf &
              "    and nf.Nota_Id         = nr.NotaReferencial_Id" & vbCrLf &
              "  INNER JOIN TipoDeDocumento TpDoc" & vbCrLf &
              "     ON NF.TipoDeDocumento = TpDoc.Codigo_Id " & vbCrLf &
              "  INNER JOIN Clientes Cli" & vbCrLf &
              "     ON NF.Cliente_Id = Cli.Cliente_Id " & vbCrLf &
              "    AND NF.EndCliente_Id = Cli.Endereco_Id " & vbCrLf &
              "   Left Join Finalidades f" & vbCrLf &
              "     On f.Finalidade_Id = nf.Finalidade" & vbCrLf &
              "  Where nr.EmpresaReferencial_Id    ='" & empresa & "'" & vbCrLf &
              "    And nr.EndEmpresaReferencial_Id = " & endempresa & vbCrLf &
              "    And nr.ClienteReferencial_id    ='" & cliente & "'" & vbCrLf &
              "    And nr.NotaReferencial_Id       = " & nota & vbCrLf &
              "  Union All " & vbCrLf &
              " Select nxn.Nota_Id, nxn.Empresa_Id, nxn.EndEmpresa_Id, nxn.Cliente_Id, nxn.EndCliente_Id, nxn.EntradaSaida_Id," & vbCrLf &
              "        nxn.Serie_Id, nfxi.Produto_Id, nfxi.CFOP_Id, '' as Sequencia_Id, nfxi.QuantidadeFiscal as Quantidade, nfxi.Valor, " & vbCrLf &
              "        'Destino / ' + cast(f.Finalidade_Id as varchar) + '-' + f.Descricao as TipoReferencial_Id, " & vbCrLf &
              "        CAST(TpDoc.Codigo_Id AS VARCHAR) + '-' + TpDoc.Descricao AS TipoDeDocumento, Cli.Nome AS ClienteNome" & vbCrLf &
              "   from notasxnotas nxn" & vbCrLf &
              "  inner join Notasfiscais NF" & vbCrLf &
              "     on Nf.Empresa_id      = nxn.empresa_id" & vbCrLf &
              "    and nf.endempresa_id   = nxn.endempresa_id" & vbCrLf &
              "    and nf.cliente_id      = nxn.cliente_id" & vbCrLf &
              "    and nf.endcliente_id   = nxn.endcliente_id" & vbCrLf &
              "    and nf.EntradaSaida_Id = nxn.EntradaSaida_Id" & vbCrLf &
              "    and nf.Serie_Id        = nxn.Serie_Id" & vbCrLf &
              "    and nf.nota_id         = nxn.nota_id" & vbCrLf &
              "  Inner Join NotasFiscaisXItens nfxi" & vbCrLf &
              "     on nfxi.Empresa_Id = nf.Empresa_Id" & vbCrLf &
              "    and nfxi.endempresa_id   = nf.endempresa_id" & vbCrLf &
              "    and nfxi.cliente_id      = nf.cliente_id" & vbCrLf &
              "    and nfxi.endcliente_id   = nf.endcliente_id" & vbCrLf &
              "    and nfxi.EntradaSaida_Id = nf.EntradaSaida_Id" & vbCrLf &
              "    and nfxi.Serie_Id        = nf.Serie_Id" & vbCrLf &
              "    and nfxi.nota_id         = nf.Nota_id" & vbCrLf &
              "  INNER JOIN TipoDeDocumento TpDoc" & vbCrLf &
              "     ON NF.TipoDeDocumento = TpDoc.Codigo_Id " & vbCrLf &
              "  INNER JOIN Clientes Cli" & vbCrLf &
              "     ON NF.Cliente_Id = Cli.Cliente_Id " & vbCrLf &
              "    AND NF.EndCliente_Id = Cli.Endereco_Id " & vbCrLf &
              "   LEFT JOIN Finalidades f" & vbCrLf &
              "     ON f.Finalidade_Id = nf.Finalidade" & vbCrLf &
              "  where nxn.OrigemEmpresa_Id      = '" & empresa & "'" & vbCrLf &
              "    and nxn.OrigemEndEmpresa_Id   = " & endempresa & vbCrLf &
              "    and nxn.OrigemCliente_Id      = '" & cliente & "'" & vbCrLf &
              "    and nxn.OrigemEndCliente_Id   = " & endcliente & vbCrLf &
              "    and nxn.OrigemEntradaSaida_Id = '" & entradasaida & "'" & vbCrLf &
              "    and nxn.OrigemSerie_Id        = '" & serie & "'" & vbCrLf &
              "    and " & nota & " in (nxn.Origemnota_id)" & vbCrLf &
              "  Union All" & vbCrLf &
              " Select nxn.OrigemNota_Id, nxn.OrigemEmpresa_Id, nxn.OrigemEndEmpresa_Id, nxn.OrigemCliente_Id, nxn.OrigemEndCliente_Id, nxn.OrigemEntradaSaida_Id," & vbCrLf &
              "        nxn.OrigemSerie_Id, nfxi.Produto_Id, nfxi.CFOP_Id, '' as Sequencia_Id, nfxi.QuantidadeFiscal as Quantidade, nfxi.Valor, " & vbCrLf &
              "        'Origem / ' + cast(f.Finalidade_Id as varchar) + '-' + f.Descricao as TipoReferencial_Id," & vbCrLf &
              "        CAST(TpDoc.Codigo_Id AS VARCHAR) + '-' + TpDoc.Descricao AS TipoDeDocumento, Cli.Nome AS ClienteNome" & vbCrLf &
              "   from notasxnotas nxn" & vbCrLf &
              "  inner join Notasfiscais NF" & vbCrLf &
              "     on Nf.Empresa_id      = nxn.Origemempresa_id" & vbCrLf &
              "    and nf.endempresa_id   = nxn.Origemendempresa_id" & vbCrLf &
              "    and nf.cliente_id      = nxn.Origemcliente_id" & vbCrLf &
              "    and nf.endcliente_id   = nxn.Origemendcliente_id" & vbCrLf &
              "    and nf.EntradaSaida_Id = nxn.OrigemEntradaSaida_Id" & vbCrLf &
              "    and nf.Serie_Id        = nxn.OrigemSerie_Id" & vbCrLf &
              "    and nf.nota_id         = nxn.OrigemNota_id" & vbCrLf &
              "  Inner Join NotasFiscaisXItens nfxi" & vbCrLf &
              "     on nfxi.Empresa_Id = nf.Empresa_Id" & vbCrLf &
              "    and nfxi.endempresa_id   = nf.endempresa_id" & vbCrLf &
              "    and nfxi.cliente_id      = nf.cliente_id" & vbCrLf &
              "    and nfxi.endcliente_id   = nf.endcliente_id" & vbCrLf &
              "    and nfxi.EntradaSaida_Id = nf.EntradaSaida_Id" & vbCrLf &
              "    and nfxi.Serie_Id        = nf.Serie_Id" & vbCrLf &
              "    and nfxi.nota_id         = nf.Nota_id" & vbCrLf &
              "  INNER JOIN Clientes Cli" & vbCrLf &
              "     ON NF.Cliente_Id = Cli.Cliente_Id " & vbCrLf &
              "    AND NF.EndCliente_Id = Cli.Endereco_Id " & vbCrLf &
              "   Left Join Finalidades f" & vbCrLf &
              "     On f.Finalidade_Id = nf.Finalidade" & vbCrLf &
              "  INNER JOIN TipoDeDocumento TpDoc" & vbCrLf &
              "     ON NF.TipoDeDocumento = TpDoc.Codigo_Id " & vbCrLf &
              "  where nxn.Empresa_id      = '" & empresa & "'" & vbCrLf &
              "    and nxn.Endempresa_Id   = " & endempresa & vbCrLf &
              "    and nxn.Cliente_id      = '" & cliente & "'" & vbCrLf &
              "    and nxn.EndCliente_Id   = " & endcliente & vbCrLf &
              "    and nxn.EntradaSaida_Id = '" & entradasaida & "'" & vbCrLf &
              "    and nxn.Serie_Id        = '" & serie & "'" & vbCrLf &
              "    and " & nota & " in (nxn.Nota_Id)" & vbCrLf &
              "  UNION ALL" & vbCrLf &
              " SELECT NN.Nota_Id, NN.Empresa_Id, NN.EndEmpresa_Id, NN.Cliente_Id, NN.EndCliente_Id, NN.EntradaSaida_Id," & vbCrLf &
              "        NN.Serie_Id, NFxI.Produto_Id, NFxI.CFOP_Id, '' as Sequencia_Id, NFxI.QuantidadeFiscal as Quantidade, NFxI.Valor, " & vbCrLf &
              "        'Origem / ' + CAST(f.Finalidade_Id AS VARCHAR) + '-' + f.Descricao as TipoReferencial_Id," & vbCrLf &
              "        CAST(TpDoc.Codigo_Id AS VARCHAR) + '-' + TpDoc.Descricao AS TipoDeDocumento, Cli.Nome AS ClienteNome" & vbCrLf &
              "   FROM NotasxNotas NxN" & vbCrLf &
              "   LEFT JOIN NotasXNotas NN" & vbCrLf &
              "     ON NN.OrigemEmpresa_id      = NxN.empresa_id" & vbCrLf &
              "    AND NN.Origemendempresa_id   = NxN.endempresa_id" & vbCrLf &
              "    AND NN.Origemcliente_id      = NxN.cliente_id" & vbCrLf &
              "    AND NN.Origemendcliente_id   = NxN.endcliente_id" & vbCrLf &
              "    AND NN.OrigemEntradaSaida_Id = NxN.EntradaSaida_Id" & vbCrLf &
              "    AND NN.OrigemSerie_Id        = NxN.Serie_Id" & vbCrLf &
              "    AND NN.Origemnota_id         = NxN.nota_id" & vbCrLf &
              "  INNER JOIN Notasfiscais NF" & vbCrLf &
              "     ON NF.Empresa_id      = NN.empresa_id" & vbCrLf &
              "    AND NF.endempresa_id   = NN.endempresa_id" & vbCrLf &
              "    AND NF.cliente_id      = NN.cliente_id" & vbCrLf &
              "    AND NF.endcliente_id   = NN.endcliente_id" & vbCrLf &
              "    AND NF.EntradaSaida_Id = NN.EntradaSaida_Id" & vbCrLf &
              "    AND NF.Serie_Id        = NN.Serie_Id" & vbCrLf &
              "    AND NF.nota_id         = NN.Nota_id" & vbCrLf &
              "  INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf &
              "     ON NFxI.Empresa_Id = nf.Empresa_Id" & vbCrLf &
              "    AND NFxI.endempresa_id   = nf.endempresa_id" & vbCrLf &
              "    AND NFxI.cliente_id      = nf.cliente_id" & vbCrLf &
              "    AND NFxI.endcliente_id   = nf.endcliente_id" & vbCrLf &
              "    AND NFxI.EntradaSaida_Id = nf.EntradaSaida_Id" & vbCrLf &
              "    AND NFxI.Serie_Id        = nf.Serie_Id" & vbCrLf &
              "    AND NFxI.nota_id         = nf.Nota_id" & vbCrLf &
              "  INNER JOIN Clientes Cli" & vbCrLf &
              "     ON NF.Cliente_Id = Cli.Cliente_Id " & vbCrLf &
              "    AND NF.EndCliente_Id = Cli.Endereco_Id " & vbCrLf &
              "   LEFT JOIN Finalidades f" & vbCrLf &
              "     ON f.Finalidade_Id = nf.Finalidade" & vbCrLf &
              "  INNER JOIN TipoDeDocumento TpDoc" & vbCrLf &
              "     ON NF.TipoDeDocumento = TpDoc.Codigo_Id " & vbCrLf &
              "  WHERE NxN.OrigemEmpresa_id    = '" & empresa & "'" & vbCrLf &
              "    AND NxN.OrigemEndempresa_Id =  " & endempresa & vbCrLf &
              "    AND NxN.OrigemCliente_id    ='" & cliente & "'" & vbCrLf &
              "    AND " & nota & " IN (NxN.OrigemNota_Id)" & vbCrLf &
              "  Union All " & vbCrLf &
              " Select nd.Nota_Id, nd.EmpresaDevolucao_Id, nd.EndEmpresaDevolucao_Id, nd.ClienteDevolucao_Id, nd.EndClienteDevolucao_Id, nd.EntradaSaida_Id, nd.Serie_Id," & vbCrLf &
              "        nd.Produto_Id, nd.CFOP_Id, nd.Sequencia_Id, nd.Quantidade, nd.Valor, CAST(f.Finalidade_Id AS VARCHAR) + '-' + f.Descricao as TipoReferencial_Id," & vbCrLf &
              "        CAST(TpDoc.Codigo_Id AS VARCHAR) + '-' + TpDoc.Descricao AS TipoDeDocumento, Cli.Nome AS ClienteNome" & vbCrLf &
              "   from NotaFiscalDevolucaoXNotaFiscal nd" & vbCrLf &
              "  Inner Join	NotasFiscais nf" & vbCrLf &
              "     on nf.Empresa_Id      = nd.EmpresaDevolucao_Id" & vbCrLf &
              "    And nf.EndEmpresa_Id   = nd.EndEmpresaDevolucao_Id" & vbCrLf &
              "    and nf.Cliente_Id      = nd.ClienteDevolucao_Id" & vbCrLf &
              "    and nf.EndCliente_Id   = nd.EndClienteDevolucao_Id" & vbCrLf &
              "    and nf.EntradaSaida_Id = nd.EntradaSaidaDevolucao_Id" & vbCrLf &
              "    and nf.Serie_Id        = nd.SerieDevolucao_Id" & vbCrLf &
              "    and nf.Nota_Id         = nd.NotaDevolucao_Id" & vbCrLf &
              "  INNER JOIN TipoDeDocumento TpDoc" & vbCrLf &
              "     ON NF.TipoDeDocumento = TpDoc.Codigo_Id " & vbCrLf &
              "  INNER JOIN Clientes Cli" & vbCrLf &
              "     ON NF.Cliente_Id = Cli.Cliente_Id " & vbCrLf &
              "    AND NF.EndCliente_Id = Cli.Endereco_Id " & vbCrLf &
              "   Left Join Finalidades f" & vbCrLf &
              "     On f.Finalidade_Id = nf.Finalidade" & vbCrLf &
              "  Where nd.EmpresaDevolucao_Id    = '" & empresa & "'" & vbCrLf &
              "    And nd.EndEmpresaDevolucao_Id = " & endempresa & vbCrLf &
              "    And nd.ClienteDevolucao_id    = '" & cliente & "'" & vbCrLf &
              "    And nd.NotaDevolucao_Id       = " & nota & vbCrLf &
              " Union All" & vbCrLf &
              " Select nd.NotaDevolucao_Id, nd.EmpresaDevolucao_Id, nd.EndEmpresaDevolucao_Id, nd.ClienteDevolucao_Id, nd.EndClienteDevolucao_Id, nd.EntradaSaidaDevolucao_Id, nd.SerieDevolucao_Id," & vbCrLf &
              "        nd.Produto_Id, nd.CFOPDevolucao_Id, nd.Sequencia_Id, nd.Quantidade, nd.Valor, CAST(f.Finalidade_Id AS VARCHAR) + '-' + f.Descricao as TipoReferencial_Id," & vbCrLf &
              "        CAST(TpDoc.Codigo_Id AS VARCHAR) + '-' + TpDoc.Descricao AS TipoDeDocumento, Cli.Nome AS ClienteNome" & vbCrLf &
              "   from NotaFiscalDevolucaoXNotaFiscal nd" & vbCrLf &
              "  Inner Join	NotasFiscais nf" & vbCrLf &
              "     on nf.Empresa_Id      = nd.EmpresaDevolucao_Id" & vbCrLf &
              "    And nf.EndEmpresa_Id   = nd.EndEmpresaDevolucao_Id" & vbCrLf &
              "    and nf.Cliente_Id      = nd.ClienteDevolucao_Id" & vbCrLf &
              "    and nf.EndCliente_Id   = nd.EndClienteDevolucao_Id" & vbCrLf &
              "    and nf.EntradaSaida_Id = nd.EntradaSaida_Id" & vbCrLf &
              "    and nf.Serie_Id        = nd.Serie_Id" & vbCrLf &
              "    and nf.Nota_Id         = nd.Nota_Id" & vbCrLf &
              "  INNER JOIN TipoDeDocumento TpDoc" & vbCrLf &
              "     ON NF.TipoDeDocumento = TpDoc.Codigo_Id " & vbCrLf &
              "  INNER JOIN Clientes Cli" & vbCrLf &
              "     ON NF.Cliente_Id = Cli.Cliente_Id " & vbCrLf &
              "    AND NF.EndCliente_Id = Cli.Endereco_Id " & vbCrLf &
              "   Left Join Finalidades f" & vbCrLf &
              "     On f.Finalidade_Id = nf.Finalidade" & vbCrLf &
              "  Where nd.EmpresaDevolucao_Id    = '" & empresa & "'" & vbCrLf &
              "    And nd.EndEmpresaDevolucao_Id = " & endempresa & vbCrLf &
              "    And nd.ClienteDevolucao_id    = '" & cliente & "'" & vbCrLf &
              "    And nd.Nota_Id       = " & nota & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NotasReferenciais")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            grdNotasReferenciadas.DataSource = ds
            grdNotasReferenciadas.DataBind()
        End If
    End Sub

    Public Function CarregarRomaneio() As Boolean
        If Not Session("objRomaneioNXI" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()
            Dim ROM As [Lib].Negocio.Romaneio = Session("objRomaneioNXI" & HID.Value)
            Session.Remove("objRomaneioNXI" & HID.Value)
            Session.Remove("ProcurandoRomaneio" & HID.Value)

            If ROM.CodigoPedido = objNotaFiscal.CodigoPedido AndAlso
               ROM.EntradaSaida = objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) AndAlso
               ROM.CodigoOperacao = objNotaFiscal.CodigoOperacao AndAlso
               ROM.CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao Then
                If objNotaFiscal.Itens(0).Produto.ControlarRomaneio And ROM.PesoLiquido > objNotaFiscal.Itens(0).SaldoPedidoFisico And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR Then
                    MsgBox(Me.Page, "O Pedido nao tem saldo suficiente para o Romaneio Selecionado.")
                    LimparGrid()
                ElseIf objNotaFiscal.Itens(0).Produto.ControlarRomaneio And ROM.PesoLiquido > objNotaFiscal.Itens(0).SaldoPedidoFisico And objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida And objNotaFiscal.SubOperacao.Devolucao = True Then
                    MsgBox(Me.Page, "O Pedido nao tem saldo suficiente para o Romaneio Selecionado.")
                    LimparGrid()
                ElseIf objNotaFiscal.Itens(0).Produto.ControlarRomaneio And ROM.PesoLiquido > objNotaFiscal.Itens(0).SaldoPedidoFisico And objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada And objNotaFiscal.SubOperacao.Devolucao = True Then
                    MsgBox(Me.Page, "O Pedido nao tem saldo suficiente para o Romaneio Selecionado.")
                    LimparGrid()
                Else
                    btnDeposito.Enabled = False
                    objNotaFiscal.CriarRomaneio = False
                    objNotaFiscal.Romaneio = ROM
                    objNotaFiscal.Deposito = objNotaFiscal.Romaneio.Deposito
                    Dim itemDeposito As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Deposito)

                    With objNotaFiscal.Deposito
                        objNotaFiscal.CodigoDeposito = .Codigo
                        objNotaFiscal.EnderecoDeposito = .CodigoEndereco
                    End With

                    Dim intIndice As Integer = ddlDeposito.Items.IndexOf(ddlDeposito.Items.FindByValue(objNotaFiscal.CodigoDeposito & "-" & CStr(objNotaFiscal.EnderecoDeposito)))
                    If intIndice = -1 Then
                        Funcoes.AdicionarClienteAoDDL(ddlDeposito, objNotaFiscal.Deposito)
                    End If

                    ddlDeposito.SelectedValue = objNotaFiscal.CodigoDeposito & "-" & objNotaFiscal.EnderecoDeposito
                    ddlDeposito.Enabled = False

                    If objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada Then
                        If ROM.NF IsNot Nothing AndAlso ROM.NF.Numero > 0 AndAlso (objNotaFiscal.NotaProdutor <> ROM.NF.Numero OrElse objNotaFiscal.SerieNotaProdutor <> ROM.NF.Serie) Then
                            MsgBox(Me.Page, "O Numero da Nota Informada no Romaneio é diferente do Numero Informado Pelo Usuario para Emissao da Nota, Romaneio " & ROM.NF.Numero & "-" & ROM.NF.Serie & " Nota " & objNotaFiscal.NotaProdutor & "-" & objNotaFiscal.SerieNotaProdutor)
                        End If

                        objNotaFiscal.CodigoDestino = objNotaFiscal.Romaneio.Pesagens(0).Pesagem.CodigoDepositante
                        objNotaFiscal.EnderecoDestino = objNotaFiscal.Romaneio.Pesagens(0).Pesagem.EnderecoDepositante

                        Dim Parametros As New Hashtable
                        Parametros.Clear()
                        Parametros.Add("Empresa", objNotaFiscal.CodigoEmpresa)
                        Parametros.Add("EndEmpresa", objNotaFiscal.EnderecoEmpresa)
                        Parametros.Add("Pedido", objNotaFiscal.CodigoPedido)
                        Parametros.Add("TipoDeposito", "OD")
                        Parametros.Add("Deposito", objNotaFiscal.CodigoDestino)
                        Parametros.Add("EndDeposito", objNotaFiscal.EnderecoDestino)
                        ddl.Carregar(ddlOrigemDestino, CarregarDDL.Tabela.DepositosPedido, "", True, Parametros)
                    End If

                    If SessaoDsXML IsNot Nothing Then
                        If objNotaFiscal.Itens.Count = 1 Then
                            If objNotaFiscal.Itens(0).Produto.Unidade = "KG" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "KGS" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "TON" Then
                                objNotaFiscal.Itens(0).QuantidadeFisica = ROM.PesoLiquido
                                objNotaFiscal.Itens(0).PesoBruto = ROM.PesoBruto
                                objNotaFiscal.Itens(0).PesoLiquido = ROM.PesoLiquido
                                'objNotaFiscal.Itens(0).PesoFiscal = ROM.PesoLiquido
                            Else

                                Dim pesoDoItem = objNotaFiscal.Pedido.Itens.Where(Function(p) p.CodigoProduto = objNotaFiscal.Itens(0).CodigoProduto).FirstOrDefault().UnidadeComercializacaoFatorDeConversao

                                If pesoDoItem > 0 Then
                                    objNotaFiscal.Itens(0).PesoFiscal = (objNotaFiscal.Itens(0).SaldoPedidoFiscal * pesoDoItem)
                                Else
                                    objNotaFiscal.Itens(0).PesoFiscal = ROM.PesoLiquido
                                End If

                            End If

                        Else
                            For Each itemNT In objNotaFiscal.Itens
                                itemNT.QuantidadeFisica = itemNT.QuantidadeFiscal

                                If itemNT.Produto.Unidade = "KG" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "KGS" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "TON" Then
                                    itemNT.PesoFiscal = itemNT.SaldoPedidoFiscal
                                Else

                                    Dim pesoDoItem = objNotaFiscal.Pedido.Itens.Where(Function(p) p.CodigoProduto = itemNT.CodigoProduto).FirstOrDefault().UnidadeComercializacaoFatorDeConversao

                                    If pesoDoItem > 0 Then
                                        itemNT.PesoFiscal = (itemNT.SaldoPedidoFiscal * pesoDoItem)
                                    Else
                                        itemNT.PesoFiscal = itemNT.SaldoPedidoFiscal
                                    End If

                                End If
                            Next
                        End If
                    Else
                        Dim pesoFormatado As String = ROM.PesoLiquido.ToString("F4")


                        If objNotaFiscal.Itens(0).Produto.ControlarRomaneio Then
                            'objNotaFiscal.Itens(0).QuantidadeFiscal = ROM.PesoLiquido
                            'objNotaFiscal.Itens(0).QuantidadeFisica = ROM.PesoLiquido
                            objNotaFiscal.Itens(0).QuantidadeFiscal = pesoFormatado
                            objNotaFiscal.Itens(0).QuantidadeFisica = pesoFormatado
                            objNotaFiscal.Itens(0).PesoBruto = ROM.PesoBruto
                            objNotaFiscal.Itens(0).PesoLiquido = ROM.PesoLiquido
                            objNotaFiscal.Itens(0).PesoFiscal = ROM.PesoLiquido
                        Else
                            If objNotaFiscal.Itens.Count = 1 Then
                                If objNotaFiscal.Itens(0).Produto.Unidade = "KG" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "KGS" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "TON" Then
                                    'objNotaFiscal.Itens(0).PesoFiscal = objNotaFiscal.Itens(0).SaldoPedidoFiscal
                                    'objNotaFiscal.Itens(0).QuantidadeFiscal = ROM.PesoLiquido
                                    'objNotaFiscal.Itens(0).QuantidadeFisica = ROM.PesoLiquido
                                    objNotaFiscal.Itens(0).QuantidadeFiscal = pesoFormatado
                                    objNotaFiscal.Itens(0).QuantidadeFisica = pesoFormatado
                                    objNotaFiscal.Itens(0).PesoFiscal = ROM.PesoLiquido
                                Else

                                    Dim pesoDoItem = objNotaFiscal.Pedido.Itens.Where(Function(p) p.CodigoProduto = objNotaFiscal.Itens(0).CodigoProduto).FirstOrDefault().UnidadeComercializacaoFatorDeConversao

                                    If pesoDoItem > 0 Then
                                        objNotaFiscal.Itens(0).PesoFiscal = (objNotaFiscal.Itens(0).SaldoPedidoFiscal * pesoDoItem)
                                    Else
                                        'objNotaFiscal.Itens(0).PesoFiscal = objNotaFiscal.Itens(0).SaldoPedidoFiscal
                                        objNotaFiscal.Itens(0).PesoFiscal = ROM.PesoLiquido
                                    End If

                                End If
                            Else
                                For Each itemNT In objNotaFiscal.Itens
                                    itemNT.QuantidadeFiscal = itemNT.SaldoPedidoFiscal
                                    itemNT.QuantidadeFisica = itemNT.QuantidadeFiscal

                                    If itemNT.Produto.Unidade = "KG" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "KGS" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "TON" Then
                                        itemNT.PesoFiscal = itemNT.SaldoPedidoFiscal
                                    Else

                                        Dim pesoDoItem = objNotaFiscal.Pedido.Itens.Where(Function(p) p.CodigoProduto = itemNT.CodigoProduto).FirstOrDefault().UnidadeComercializacaoFatorDeConversao

                                        If pesoDoItem > 0 Then
                                            itemNT.PesoFiscal = (itemNT.SaldoPedidoFiscal * pesoDoItem)
                                        Else
                                            itemNT.PesoFiscal = itemNT.SaldoPedidoFiscal
                                        End If

                                    End If
                                Next
                            End If
                        End If
                    End If


                    'objNotaFiscal.Itens(0).SaldoValorOficial = Math.Max(objNotaFiscal.Itens(0).SaldoPedidoFisico, objNotaFiscal.Itens(0).SaldoPedidoFiscal)
                    '#FimBaseDeCalculo
                    'objNotaFiscal.Itens(0).ValorTotal = (objNotaFiscal.Itens(0).QuantidadeFiscal * objNotaFiscal.Itens(0).Unitario) / objNotaFiscal.Itens(0).Produto.BaseCalculo

                    If objNotaFiscal.SubOperacao.Devolucao Then
                        objNotaFiscal.Itens(0).ValorTotal = 0
                        objNotaFiscal.Itens(0).Unitario = 0
                    Else
                        objNotaFiscal.Itens(0).ValorTotal = Math.Round((objNotaFiscal.Itens(0).QuantidadeFiscal * objNotaFiscal.Itens(0).Unitario), 2, MidpointRounding.AwayFromZero)
                    End If

                    If objNotaFiscal.Itens.Count = 1 Then
                        If objNotaFiscal.Itens(0).Produto.Unidade = "KG" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "KGS" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "TON" Then
                            objNotaFiscal.Itens(0).Volumes = ROM.PesoLiquido
                            objNotaFiscal.Quantidade = ROM.PesoLiquido
                            objNotaFiscal.PesoBruto = ROM.PesoLiquido
                            objNotaFiscal.PesoLiquido = ROM.PesoLiquido
                        Else
                            objNotaFiscal.PesoBruto = ROM.PesoLiquido
                        End If
                    Else
                        objNotaFiscal.PesoBruto = ROM.PesoLiquido
                    End If

                    txtPesoRomaneio.Text = ROM.PesoLiquido

                    '-----------------------------------------

                    objNotaFiscal.CodigoRomaneio = ROM.Codigo

                    txtCodigoTransportador.Value = ROM.CodigoTransportador
                    objNotaFiscal.CodigoTransportador = ROM.CodigoTransportador
                    objNotaFiscal.EnderecoTransportador = ROM.EnderecoTransportador
                    objNotaFiscal.CodigoMotorista = ROM.CodigoMotorista
                    objNotaFiscal.EnderecoMotorista = ROM.EnderecoMotorista
                    objNotaFiscal.PlacaTransportador = ROM.Placa
                    txtCodigoPlaca.Value = objNotaFiscal.PlacaTransportador
                    objNotaFiscal.CodigoAutorizacao = ROM.CodigoAutorizacao

                    If ROM.CodigoAutorizacao > 0 AndAlso Not objNotaFiscal.ObservacoesDeEmbarque.Contains(". " & observacaoAutorizacao.Value & ". ") Then
                        observacaoAutorizacao.Value = objNotaFiscal.Autorizacao.Observacao
                        objNotaFiscal.ObservacoesDeEmbarque &= ". " & observacaoAutorizacao.Value & ". "
                    End If

                    If Not String.IsNullOrWhiteSpace(objNotaFiscal.Romaneio.Placa) AndAlso Not objNotaFiscal.ObservacoesDeEmbarque.Contains("Motorista:") Then
                        Dim objPlaca As New [Lib].Negocio.Placa(objNotaFiscal.Romaneio.Placa)
                        objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|Motorista: " & objPlaca.Motorista.Nome & " - CPF " & objPlaca.Motorista.CodigoFormatado & " - CNH " & objPlaca.Habilitacao & " - Placa " & objPlaca.Placa01 & " " & objPlaca.Placa02 & " " & objPlaca.Placa03 & " " & objPlaca.Placa04
                    End If

                    observacaoValores.Value = String.Empty

                    For Each enc In objNotaFiscal.Itens(0).Encargos
                        If enc.Codigo.Contains("FUNRURAL") Then
                            observacaoValores.Value = observacaoValores.Value & "|FUNRURAL - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                        ElseIf enc.Codigo.Contains("FABOV") Then
                            observacaoValores.Value = observacaoValores.Value & "|FABOV - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                        ElseIf enc.Codigo.Contains("SENAR") Then
                            observacaoValores.Value = observacaoValores.Value & "|SENAR - " & enc.Percentual.ToString("N2").Replace(",", ".") & " % - R$ = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                        ElseIf enc.Codigo.Contains("FETHAB") Then
                            observacaoValores.Value = observacaoValores.Value & "|FETHAB - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                        ElseIf enc.Codigo.Contains("IAGRO") Then
                            observacaoValores.Value = observacaoValores.Value & "|IAGRO - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                        ElseIf enc.Codigo.Contains("FUNDEMS") Then
                            observacaoValores.Value = observacaoValores.Value & "|FUNDEMS - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                        ElseIf enc.Codigo.Contains("FUNDERSUL") Then
                            observacaoValores.Value = observacaoValores.Value & "|FUNDERSUL - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                        ElseIf enc.Codigo.Contains("FACS") Then
                            observacaoValores.Value = observacaoValores.Value & "|FACS - " & enc.Percentual.ToString.Replace(",", ".") & " % - R$ = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                        End If
                    Next

                    SessaoSalvaNotaFiscal()
                    AtualizaFormularioComAClasse()
                    BuscaProcuracao()
                End If
            Else
                MsgBox(Me.Page, "Pedido do Romaneio " & ROM.CodigoPedido.ToString & " não pode ser diferente do Pedido da Nota Fiscal " & objNotaFiscal.CodigoPedido.ToString)
                LimparCampos("", "", False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
            End If

            Return True
        ElseIf Not Session("ProcurandoRomaneio" & HID.Value) Is Nothing Then
            Session.Remove("ProcurandoRomaneio" & HID.Value)

            SessaoRecuperaNotaFiscal()
            If Not BuscaNotaTroca() AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.Itens(0).Produto.ControlarPesagem Then
                BuscaClassificacao()
            End If

            Return True
        End If
        Return False
    End Function

    Public Function CarregarProcuracao() As Boolean
        If Not Session("objProcuracaoNxI" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()
            objNotaFiscal.ProcuracaoSaldoPedido = Session("TotalProcuracao" & HID.Value)
            objNotaFiscal.CodigoProcuracao = Session("objProcuracaoNxI" & HID.Value)
            cmbFinalidade.Enabled = True

            If objNotaFiscal.Procuracao.Saldo < objNotaFiscal.Itens(0).QuantidadeFiscal Then
                MsgBox(Me.Page, "O Peso da nota fiscal não pode ser maior que o saldo da Cessão de Crédito!")
                objNotaFiscal.CodigoProcuracao = 0
                objNotaFiscal.Procuracao = Nothing
                Dim parameters As New Dictionary(Of String, Object)
                parameters.Add("tipo", "NxI")
                parameters.Add("pedc", objNotaFiscal.CodigoPedido)
                Popup.ConsultaDeProcuracao(Me.Page, "objProcuracaoNxI" & HID.Value)
                ucConsultaProcuracao.BindGridView(parameters)
            Else
                txtCessaoDeCredito.Text = objNotaFiscal.CodigoProcuracao
                objNotaFiscal.CodigoFinalidade = 14
                cmbFinalidade.SelectedValue = 14
                cmbFinalidade.Enabled = False
            End If
            Session.Remove("objProcuracaoNxI" & HID.Value)
            Session.Remove("ProcurandoProcuracao" & HID.Value)
            SessaoSalvaNotaFiscal()
            If Not objNotaFiscal.CodigoAutorizacao > 0 AndAlso
                objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso
                (objNotaFiscal.SubOperacao.QuantidadeFisico OrElse objNotaFiscal.SubOperacao.QuantidadeFiscal) Then BuscaAutorizacao()
            Return True
        ElseIf Not Session("SemProcuracao" & HID.Value) Is Nothing Or Not Session("ProcurandoProcuracao" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()
            Session.Remove("SemProcuracao" & HID.Value)
            Session.Remove("ProcurandoProcuracao" & HID.Value)
            If Not objNotaFiscal.CodigoAutorizacao > 0 AndAlso
                objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso
                (objNotaFiscal.SubOperacao.QuantidadeFisico OrElse objNotaFiscal.SubOperacao.QuantidadeFiscal) Then BuscaAutorizacao()
            Return True
        End If
        Return False
    End Function

    Public Function CarregarClassificacao() As Boolean
        If Not Session("objClassificacaoNXI" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()

            If objNotaFiscal.Romaneio.PesoLiquido > objNotaFiscal.Itens(0).SaldoPedidoFisico AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS AndAlso objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR AndAlso Not objNotaFiscal.SubOperacao.Devolucao = True Then
                MsgBox(Me.Page, "O Pedido nao tem saldo suficiente para o Romaneio Selecionado.")
                LimparGrid()
                objNotaFiscal.CodigoRomaneio = 0
                SessaoSalvaNotaFiscal()
            ElseIf objNotaFiscal.Romaneio.PesoLiquido > objNotaFiscal.Itens(0).SaldoPedidoFisico AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso objNotaFiscal.SubOperacao.Devolucao = True Then
                MsgBox(Me.Page, "O Pedido nao tem saldo suficiente para o Romaneio Selecionado.")
                LimparGrid()
                objNotaFiscal.CodigoRomaneio = 0
                SessaoSalvaNotaFiscal()
            Else
                'If Session("Carregando" & HID.Value) = False Then
                '    objNotaFiscal.CriarRomaneio = True
                'Else
                '    objNotaFiscal.CriarRomaneio = False
                'End If
                'Se está incluindo a nota e vir aqui é porque tem que criar o romaneio - furlan - 11/08/2015
                If objNotaFiscal.IUD = "I" Then objNotaFiscal.CriarRomaneio = True

                txtPesoRomaneio.Text = objNotaFiscal.Romaneio.PesoLiquido

                If Not objNotaFiscal.Itens.Count > 1 Then
                    Dim pesoFormatado As String = objNotaFiscal.Romaneio.PesoBruto.ToString("F4")

                    'SE NÃO É IMPORTAÇÃO DE XML, DEVE ASSUMIR O QUE VEM DA CLASSIFICAÇÃO DO ROMANEIO - FURLAN - 28/05/2025
                    If SessaoDsXML Is Nothing Then
                        objNotaFiscal.Itens(0).QuantidadeFiscal = pesoFormatado
                    End If

                    If objNotaFiscal.SubOperacao.QuantidadeFisico Then
                        pesoFormatado = objNotaFiscal.Romaneio.PesoLiquido.ToString("F4")

                        objNotaFiscal.Itens(0).QuantidadeFisica = pesoFormatado
                    Else
                        objNotaFiscal.Itens(0).QuantidadeFisica = 0
                    End If

                    objNotaFiscal.Itens(0).PesoBruto = objNotaFiscal.Romaneio.PesoBruto
                    objNotaFiscal.Itens(0).PesoLiquido = objNotaFiscal.Romaneio.PesoLiquido
                    objNotaFiscal.Itens(0).PesoFiscal = objNotaFiscal.Romaneio.PesoBruto

                    'COMENTEI PORQUE INDEPENDETE SE É DEVOLUÇÃO OU NÃO, E NÃO É IMPORTAÇÃO DE XML, DEVE ASSUMIR O QUE VEM DA CLASSIFICAÇÃO DO ROMANEIO - FURLAN - 28/05/2025
                    'If objNotaFiscal.SubOperacao.Devolucao And SessaoDsXML Is Nothing Then
                    '    objNotaFiscal.Itens(0).ValorTotal = 0
                    '    objNotaFiscal.Itens(0).Unitario = 0
                    'Else
                    If SessaoDsXML Is Nothing Then
                        objNotaFiscal.Itens(0).ValorTotal = Math.Round((objNotaFiscal.Itens(0).QuantidadeFiscal * objNotaFiscal.Itens(0).Unitario), 2, MidpointRounding.AwayFromZero)
                    End If
                    'End If



                End If

                objNotaFiscal.CodigoRomaneio = objNotaFiscal.Romaneio.Codigo

                For Each item In objNotaFiscal.Itens
                    For Each enc As NotaFiscalXItemXEncargo In item.Encargos
                        If enc.Codigo.Contains("DESONERADO") Then
                            objNotaFiscal.Observacoes = objNotaFiscal.Observacoes & " | VALOR DE ICMS DISPENSADO DE R$ " & enc.Valor.ToString & " POR MOTIVO, " & enc.ObservacaoFiscal
                        End If
                    Next
                Next

                objNotaFiscal.PesoBruto = objNotaFiscal.Romaneio.PesoBruto
                objNotaFiscal.PesoLiquido = objNotaFiscal.Romaneio.PesoLiquido
                objNotaFiscal.Quantidade = objNotaFiscal.Romaneio.PesoLiquido

                observacaoValores.Value = String.Empty

                For Each enc In objNotaFiscal.Itens(0).Encargos
                    If enc.Codigo.Contains("FUNRURAL") Then
                        observacaoValores.Value = observacaoValores.Value & "|FUNRURAL - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("FABOV") Then
                        observacaoValores.Value = observacaoValores.Value & "|FABOV - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("SENAR") Then
                        observacaoValores.Value = observacaoValores.Value & "|SENAR - " & enc.Percentual.ToString("N2").Replace(",", ".") & " % - R$ = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("FETHAB") Then
                        observacaoValores.Value = observacaoValores.Value & "|FETHAB - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("IAGRO") Then
                        observacaoValores.Value = observacaoValores.Value & "|IAGRO - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("FUNDEMS") Then
                        observacaoValores.Value = observacaoValores.Value & "|FUNDEMS - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("FUNDERSUL") Then
                        observacaoValores.Value = observacaoValores.Value & "|FUNDERSUL - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("FACS") Then
                        observacaoValores.Value = observacaoValores.Value & "|FACS - " & enc.Percentual.ToString.Replace(",", ".") & " % - R$ = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    End If
                Next

                'Metodo alimenta informações importantes do IBS e CBS
                objNotaFiscal.AtualizarObservacoes()
                txtObservacoesFiscais.Text = objNotaFiscal.Observacoes
                'Metodo alimenta informações importantes do IBS e CBS

                SessaoSalvaNotaFiscal()
                AtualizaFormularioComAClasse()

                Session.Remove("objClassificacaoNXI" & HID.Value)
                Session.Remove("ProcurandoClassificacao" & HID.Value)
                BuscaProcuracao()
            End If
            Return True
        ElseIf Not Session("ProcurandoClassificacao" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()
            Session.Remove("ProcurandoClassificacao" & HID.Value)
            BuscaProcuracao()
            Return True
        End If
        Return False
    End Function

    Public Sub CarregarNotaClassificacao()

        SessaoRecuperaNotaFiscal()

        Dim objNotaAntes = New NotaFiscal(objNotaFiscal)
        objNotaFiscal = New NotaFiscal(objNotaAntes)
        objNotaFiscal.NotaFiscalOriginal = New NotaFiscal(objNotaFiscal)
        SessaoSalvaNotaFiscal()

        AtualizaFormularioComAClasse()

        Popup.CloseDialog(Me.Page, "divConsultaNotaFiscalXClassificacao")

    End Sub

    Public Sub CarregarAutorizacao(Par As Hashtable)

        Try
            If Not Session("objAutorizacaoNXI" & HID.Value) Is Nothing Then
                SessaoRecuperaNotaFiscal()

                Dim objAutorizacao As New [Lib].Negocio.AutorizacaoDeRetirada(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, Par("Pedido"), Par("Autorizacao"), objNotaFiscal.SubOperacao.Classe)
                If objAutorizacao.CodigoPedido = objNotaFiscal.CodigoPedido Then
                    objNotaFiscal.CodigoAutorizacao = Session("objAutorizacaoNXI" & HID.Value)
                    txtAutorizacao.Text = objNotaFiscal.CodigoAutorizacao

                    If objNotaFiscal.SubOperacao.QuantidadeFiscal AndAlso objNotaFiscal.Autorizacao.SaldoFiscal < IIf(objNotaFiscal.Itens(0).PesoQuantidade = "P", objNotaFiscal.Itens(0).PesoFiscal, objNotaFiscal.Itens(0).QuantidadeFiscal) Then

                        If objNotaFiscal.Autorizacao.SaldoFiscal > 0 AndAlso objNotaFiscal.Autorizacao.SaldoFiscal <= objNotaFiscal.Itens(0).QuantidadeFiscal Then
                            objNotaFiscal.Itens(0).QuantidadeFiscal = objNotaFiscal.Autorizacao.SaldoFiscal
                            objNotaFiscal.Itens(0).PesoFiscal = objNotaFiscal.Autorizacao.SaldoFiscal
                            MsgBox(Me.Page, "Saldo fiscal da autorização insuficiente para geração da nota, Saldo: " & objNotaFiscal.Autorizacao.SaldoFiscal & " Nota Fiscal: " & objNotaFiscal.Itens(0).PesoFiscal & ", vamos considerar o saldo da Autorização.")
                        Else
                            MsgBox(Me.Page, "Saldo fiscal da autorização insuficiente para geração da nota, Saldo: " & objNotaFiscal.Autorizacao.SaldoFiscal & " Nota Fiscal: " & objNotaFiscal.Itens(0).PesoFiscal)
                            LimparCampos("", "", False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                            Exit Sub
                        End If
                    End If
                    If objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.Autorizacao.SaldoFisico < objNotaFiscal.Itens(0).PesoLiquido Then

                        If objNotaFiscal.Autorizacao.SaldoFisico > 0 AndAlso objNotaFiscal.Autorizacao.SaldoFisico <= objNotaFiscal.Itens(0).PesoLiquido Then
                            'objNotaFiscal.Itens(0).QuantidadeFisica = objNotaFiscal.Autorizacao.SaldoFisico
                            objNotaFiscal.Itens(0).PesoLiquido = objNotaFiscal.Autorizacao.SaldoFisico
                            MsgBox(Me.Page, "Saldo físico da autorização insuficiente para geração da nota, Saldo: " & objNotaFiscal.Autorizacao.SaldoFisico & " Nota Fiscal: " & objNotaFiscal.Itens(0).PesoLiquido & ", vamos considerar o saldo da Autorização.")
                        Else
                            MsgBox(Me.Page, "Saldo físico da autorização insuficiente para geração da nota, Saldo: " & objNotaFiscal.Autorizacao.SaldoFisico & " Nota Fiscal: " & objNotaFiscal.Itens(0).PesoLiquido)
                            LimparCampos("", "", False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                            Exit Sub
                        End If
                    End If

                    observacaoAutorizacao.Value = objNotaFiscal.Autorizacao.Observacao

                    'If String.IsNullOrWhiteSpace(observacaoAutorizacao.Value) Then
                    '    If Not objNotaFiscal.ObservacoesDeEmbarque.Contains(". " & observacaoValores.Value) Then
                    '        objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
                    '    End If
                    '    txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque
                    '    'txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
                    'Else

                    '    If Not objNotaFiscal.ObservacoesDeEmbarque.Contains(". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value) Then
                    '        objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
                    '    End If
                    '    txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque
                    '    'txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
                    'End If
                    If String.IsNullOrWhiteSpace(observacaoAutorizacao.Value) Then
                        txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
                    Else
                        txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
                    End If

                    SessaoSalvaNotaFiscal()

                    AtualizaFormularioComAClasse()
                Else
                    MsgBox(Me.Page, "Pedido da autorização " & objAutorizacao.CodigoPedido.ToString & " não pode ser diferente do pedido da nota fiscal " & objNotaFiscal.CodigoPedido.ToString)
                    BuscaAutorizacao()
                End If
                Session.Remove("objAutorizacaoNXI" & HID.Value)
            ElseIf Not Session("SemAutorizacao" & HID.Value) Is Nothing Then
                SessaoRecuperaNotaFiscal()
                objNotaFiscal.CodigoAutorizacao = Session("SemAutorizacao" & HID.Value)
                Session.Remove("SemAutorizacao" & HID.Value)
                txtAutorizacao.Text = objNotaFiscal.CodigoAutorizacao
                SessaoSalvaNotaFiscal()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Original
    'Public Sub CarregarNotasDevolucao(Ret As RetornoNotasDevolucao)
    '    If Not Session("ssDevolucaoXNota" & HID.Value) Is Nothing Then
    '        SessaoRecuperaNotaFiscal()
    '        Dim index As Integer = Session("ssDevolucaoXNota" & HID.Value)
    '        If objNotaFiscal.SubOperacao.Devolucao _
    '         And objNotaFiscal.Itens(index).QuantidadeFiscal <> objNotaFiscal.Itens(index).NotasDevolucao.SomaQtde _
    '           Or Math.Round(objNotaFiscal.Itens(index).ValorTotal, 2) <> Math.Round(objNotaFiscal.Itens(index).NotasDevolucao.SomaVlr, 2) Then
    '            objNotaFiscal.Itens(index).NotasDevolucao = Nothing
    '        End If
    '        Session.Remove("ssDevolucaoXNota" & HID.Value)
    '        SessaoSalvaNotaFiscal()
    '    End If
    'End Sub

    'Original Acima
    'Public Sub CarregarNotasDevolucao(ListRet As List(Of RetornoNotasDevolucao))
    '    SessaoRecuperaNotaFiscal()
    '    Dim itemOrigem As [Lib].Negocio.NotaFiscalXItem = objNotaFiscal.Itens(ListRet(0).indexItem)
    '    For Each ret In ListRet
    '        Dim objItemNF As New [Lib].Negocio.NotaFiscalXItem(objNotaFiscal)

    '        With objItemNF
    '            .CodigoProduto = itemOrigem.CodigoProduto

    '            .Lote = itemOrigem.Lote
    '            .Classificacao = itemOrigem.Classificacao
    '            .CodigoEmbalagem = itemOrigem.CodigoEmbalagem
    '            .CodigoEmbalagemIndea = itemOrigem.CodigoEmbalagemIndea
    '            .CodigoTipoDeEmbalagem = itemOrigem.CodigoTipoDeEmbalagem
    '            .CapacidadeEmbalagem = itemOrigem.CapacidadeEmbalagem
    '            .QuantidadeDeEmbalagem = itemOrigem.QuantidadeDeEmbalagem


    '            If .Produto.ControlarEmbalagem Then
    '                .ObservacoesDoProduto = itemOrigem.QuantidadeDeEmbalagem & " " & .Embalagem.Descricao & " " & .TipoDeEmbalagem.Descricao & " " & IIf(itemOrigem.EmbalagemProduto.PesoVariavel, "", itemOrigem.CapacidadeEmbalagem & " " & .Produto.Unidade) & " / " & .Produto.DescricaoTecnica
    '            End If

    '            If .Produto.ControlarLote Then
    '                Dim L As New [Lib].Negocio.Lote(itemOrigem.Lote, .CodigoProduto)
    '                If L.Tipo = 2 Then
    '                    .ObservacoesDoProduto &= IIf(.ObservacoesDoProduto.Length = 0, "", ", ") & "Lote: " & .Lote & " Validade " & L.DataValidade.ToString("dd-MM-yyyy")
    '                Else
    '                    .ObservacoesDoProduto &= IIf(.ObservacoesDoProduto.Length = 0, "", ", ") & "Lote: " & .Lote & " GER " & L.Germinacao & IIf(L.Pureza = 0, "", " PUR " & L.Pureza) & " Peneira/Classif. " & .Classificacao & " Renasem " & L.Renasem & " Validade " & L.DataValidade.ToString("dd-MM-yyyy")
    '                End If

    '            End If

    '            .Sequencia = i
    '            i += 1

    '            .CodigoOperacao = objNotaFiscal.CodigoOperacao
    '            .CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao

    '            .CodigoPedido = itemOrigem.CodigoPedido
    '            .PesoQuantidade = itemOrigem.PesoQuantidade

    '            .PesoFiscal = ret.Quantidade
    '            If objNotaFiscal.SubOperacao.QuantidadeFisico Then
    '                .QuantidadeFisica = ret.Quantidade
    '            End If
    '            .QuantidadeFiscal = ret.Quantidade

    '            .SaldoValorOficial = itemOrigem.SaldoValorOficial
    '            .SaldoValorMoeda = itemOrigem.SaldoValorMoeda
    '            .SaldoPedidoFiscal = itemOrigem.SaldoPedidoFiscal
    '            .SaldoPedidoFisico = itemOrigem.SaldoPedidoFisico

    '            .UnitarioMoeda = itemOrigem.UnitarioMoeda
    '            .Unitario = ret.Unitario
    '            .ValorTotal = ret.Valor

    '            objItemNF.CarregandoEncargos = True
    '            .Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objItemNF, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
    '            objItemNF.CarregandoEncargos = False

    '            objItemNF.NotasDevolucao = New ListNotaFiscalDevolucaoXNotaFiscal(objItemNF)
    '            For Each x In itemOrigem.NotasDevolucao.Where(Function(s) s.UnitarioNota = ret.Unitario And (s.ValorDevolucao > 0 Or s.QuantidadeDevolucao > 0))
    '                objItemNF.NotasDevolucao.Add(x)
    '            Next
    '            objNotaFiscal.Itens.Add(objItemNF)
    '        End With
    '    Next
    '    objNotaFiscal.Itens.Remove(itemOrigem)

    '    '**********************************************************************************************
    '    '*******************  Cria a estrutura da Tabela que alimenta o Grid  *************************
    '    '**********************************************************************************************
    '    Dim dtItens As New DataTable("Itens")
    '    dtItens.Columns.Add("Produto", Type.GetType("System.String"))
    '    dtItens.Columns.Add("NomeProduto", Type.GetType("System.String"))
    '    dtItens.Columns.Add("BaseCalculo", Type.GetType("System.String"))
    '    dtItens.Columns.Add("Lote", Type.GetType("System.String"))
    '    dtItens.Columns.Add("Classificacao", Type.GetType("System.String"))
    '    dtItens.Columns.Add("Embalagem", Type.GetType("System.String"))
    '    dtItens.Columns.Add("Saldo", Type.GetType("System.Decimal"))
    '    dtItens.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
    '    dtItens.Columns.Add("QuantidadeFisica", Type.GetType("System.Decimal"))
    '    dtItens.Columns.Add("Unitario", Type.GetType("System.Decimal"))
    '    dtItens.Columns.Add("Total", Type.GetType("System.Decimal"))
    '    '******************************************************************************************************
    '    '************************** Tabela que Alimenta o Grid Da Tela  ***************************************
    '    '******************************************************************************************************
    '    For Each objItemNF As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
    '        Dim drItem As DataRow = dtItens.NewRow()
    '        objNotaFiscal.Especie = objItemNF.Produto.Embalagem.Descricao
    '        drItem("Produto") = objItemNF.CodigoProduto
    '        If objItemNF.Produto.ControlarLote Or objItemNF.Produto.ControlarEmbalagem Then
    '            drItem("NomeProduto") = "$" & objItemNF.Produto.Nome
    '        Else
    '            drItem("NomeProduto") = objItemNF.Produto.Nome
    '        End If
    '        drItem("Lote") = objItemNF.Lote
    '        drItem("Classificacao") = objItemNF.Classificacao
    '        drItem("Embalagem") = objItemNF.CodigoEmbalagemIndea + "-" + objItemNF.CodigoTipoDeEmbalagem + "-" + objItemNF.CapacidadeEmbalagem.ToString

    '        '#FimBaseDeCalculo
    '        'drItem("BaseCalculo") = objItemNF.Produto.BaseCalculo
    '        drItem("BaseCalculo") = 1

    '        If Not objItemNF.SubOperacao.QuantidadeFiscal Then
    '            drItem("Saldo") = objItemNF.SaldoValorOficial
    '        Else
    '            drItem("Saldo") = objItemNF.SaldoPedidoFiscal
    '        End If

    '        drItem("Quantidade") = objItemNF.QuantidadeFiscal
    '        drItem("QuantidadeFisica") = objItemNF.QuantidadeFisica.ToString("N4")
    '        drItem("Unitario") = objItemNF.Unitario
    '        drItem("Total") = objItemNF.ValorTotal
    '        dtItens.Rows.Add(drItem)
    '    Next

    '    objNotaFiscal.AtualizaTotais()
    '    objNotaFiscal.CarregandoItens = False

    '    'Importação do xml de notas de entrada
    '    If SessaoDsXML IsNot Nothing Then
    '        Dim ObsTemp As String = objNotaFiscal.Observacoes
    '        objNotaFiscal.AtualizarObservacoes()
    '        objNotaFiscal.Observacoes &= ObsTemp
    '    Else
    '        objNotaFiscal.AtualizarObservacoes()
    '    End If
    '    txtObservacoesFiscais.Text = objNotaFiscal.Observacoes
    '    txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque

    '    AtualizaFormulario()
    '    SessaoSalvaNotaFiscal()

    '    gridItens.DataSource = dtItens
    '    gridItens.DataBind()

    '    i = 0
    '    While i < gridItens.Rows.Count
    '        If Mid(gridItens.Rows(i).Cells(2).Text, 1, 1) = "$" Or objNotaFiscal.TemNotaTroca Then
    '            Dim Campo As TextBox = CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox)
    '            Campo.Enabled = False
    '            Campo = CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox)
    '            Campo.Enabled = False
    '            gridItens.Rows(i).Cells(2).Text = gridItens.Rows(i).Cells(2).Text.Replace("$", "")
    '        End If
    '        i += 1
    '    End While
    'End Sub

    Public Sub CarregarNotasDevolucao(ListRet As List(Of RetornoNotasDevolucao))
        SessaoRecuperaNotaFiscal()

        If ListRet IsNot Nothing AndAlso ListRet.Count = 0 Then
            MsgBox(Me.Page, "Não foi informado itens para devolução!", eTitulo.Info)
            Exit Sub
        End If

        Dim itemOrigem As [Lib].Negocio.NotaFiscalXItem = objNotaFiscal.Itens(ListRet(0).indexItem)
        Dim pesoNota As Decimal = 0

        If SessaoDsXML IsNot Nothing Then
            pesoNota = itemOrigem.QuantidadeFiscal
        Else
            For Each ret In ListRet

                itemOrigem.Unitario = ret.Unitario
                itemOrigem.ValorTotal = ret.Valor
                pesoNota += itemOrigem.QuantidadeFiscal

                'itemOrigem.Encargos.AtualizaLiquido()
                'CType(gridItens.Rows(ListRet(0).indexItem).FindControl("txtUnitarioItem"), TextBox).Text = ret.Unitario.ToString("N10")
                'CType(gridItens.Rows(ListRet(0).indexItem).FindControl("txtTotalItem"), TextBox).Text = ret.Valor.ToString("N2")
                'CType(gridItens.Rows(ListRet(0).indexItem).FindControl("txtUnitarioItem"), TextBox).Enabled = False
                'CType(gridItens.Rows(ListRet(0).indexItem).FindControl("txtQuantidadeItem"), TextBox).Enabled = False
            Next
        End If

        objNotaFiscal.DataDevolucao = objNotaFiscal.Movimento

        For Each dev In itemOrigem.NotasDevolucao
            If dev.ValorDevolucao > 0 Then
                If dev.DataDaNota < objNotaFiscal.DataDevolucao Then
                    objNotaFiscal.DataDevolucao = dev.DataDaNota
                End If
            End If
        Next

        itemOrigem.CarregandoEncargos = True
        itemOrigem.Encargos = Nothing
        itemOrigem.Encargos = New ListNotaFiscalXItemXEncargo(itemOrigem, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
        itemOrigem.CarregandoEncargos = False

        itemOrigem.Encargos.AtualizaLiquido()

        CType(gridItens.Rows(ListRet(0).indexItem).FindControl("txtUnitarioItem"), TextBox).Text = itemOrigem.Unitario.ToString("N10")
        CType(gridItens.Rows(ListRet(0).indexItem).FindControl("txtTotalItem"), TextBox).Text = itemOrigem.ValorTotal.ToString("N2")
        CType(gridItens.Rows(ListRet(0).indexItem).FindControl("txtUnitarioItem"), TextBox).Enabled = False
        CType(gridItens.Rows(ListRet(0).indexItem).FindControl("txtQuantidadeItem"), TextBox).Enabled = False


        'INDEXADOR DO PEDIDO FOR FIXO DEVE PEGAR O UNITÁRIO EM DÓLAR DO PEDIDO - FURLAN - 18/06/2025
        If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
            If objNotaFiscal.Pedido.CodigoIndexador = 99 OrElse objNotaFiscal.Pedido.IndexadorFixo Then
                itemOrigem.ValorTotalMoeda = Math.Round(itemOrigem.ValorTotal / objNotaFiscal.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
            Else
                Dim indiceDolar As Decimal = New Cotacao(objNotaFiscal.Pedido.CodigoIndexador, objNotaFiscal.DataNota).Indice 'PTAX
                itemOrigem.ValorTotalMoeda = Math.Round(itemOrigem.ValorTotal / indiceDolar, 2, MidpointRounding.AwayFromZero)
            End If
            If itemOrigem.QuantidadeFiscal > 0 Then
                itemOrigem.UnitarioMoeda = Math.Round(itemOrigem.ValorTotalMoeda / itemOrigem.QuantidadeFiscal, 10, MidpointRounding.AwayFromZero)
            End If
        End If

        objNotaFiscal.AtualizaTotais()

        If Not objNotaFiscal.Romaneio Is Nothing AndAlso objNotaFiscal.Romaneio.PesoLiquido > 0 Then
            objNotaFiscal.Itens(0).QuantidadeFiscal = pesoNota
            CType(gridItens.Rows(ListRet(0).indexItem).FindControl("txtQuantidadeItem"), TextBox).Text = pesoNota.ToString("N4")
        End If

        objNotaFiscal.PesoBruto = pesoNota
        objNotaFiscal.PesoLiquido = pesoNota
        objNotaFiscal.Quantidade = pesoNota

        observacaoValores.Value = String.Empty

        For Each enc In objNotaFiscal.Itens(0).Encargos
            If enc.Codigo.Contains("FUNRURAL") Then
                observacaoValores.Value = observacaoValores.Value & "|FUNRURAL - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
            ElseIf enc.Codigo.Contains("FABOV") Then
                observacaoValores.Value = observacaoValores.Value & "|FABOV - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
            ElseIf enc.Codigo.Contains("SENAR") Then
                observacaoValores.Value = observacaoValores.Value & "|SENAR - " & enc.Percentual.ToString("N2").Replace(",", ".") & " % - R$ = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
            ElseIf enc.Codigo.Contains("FETHAB") Then
                observacaoValores.Value = observacaoValores.Value & "|FETHAB - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
            ElseIf enc.Codigo.Contains("IAGRO") Then
                observacaoValores.Value = observacaoValores.Value & "|IAGRO - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
            ElseIf enc.Codigo.Contains("FUNDEMS") Then
                observacaoValores.Value = observacaoValores.Value & "|FUNDEMS - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
            ElseIf enc.Codigo.Contains("FUNDERSUL") Then
                observacaoValores.Value = observacaoValores.Value & "|FUNDERSUL - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
            ElseIf enc.Codigo.Contains("FACS") Then
                observacaoValores.Value = observacaoValores.Value & "|FACS - " & enc.Percentual.ToString.Replace(",", ".") & " % - R$ = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
            End If
        Next

        If String.IsNullOrWhiteSpace(observacaoAutorizacao.Value) Then
            txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
        Else
            txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
        End If

        AtualizaFormulario()
        SessaoSalvaNotaFiscal()
    End Sub

    Public Sub CarregarDepositosPedido()
        'Parametros Carga Deposito Pedido
        'Empresa
        'EndEmpresa
        'Pedido
        'TipoDeposito - DE - Deposito; OD - OrigemDestino; TR - Transbordo;  LE - Local de embarque
        'Deposito     - Esse Deposito sera adicionado mesmo q nao estaja cadastrado no pedido
        'EndDeposito
        Dim Parametros As New Hashtable
        Parametros.Add("Empresa", objNotaFiscal.CodigoEmpresa)
        Parametros.Add("EndEmpresa", objNotaFiscal.EnderecoEmpresa)
        Parametros.Add("Pedido", objNotaFiscal.CodigoPedido)
        Parametros.Add("TipoDeposito", "DE")
        Parametros.Add("Deposito", objNotaFiscal.CodigoDeposito)
        Parametros.Add("EndDeposito", objNotaFiscal.EnderecoDeposito)
        ddl.Carregar(ddlDeposito, CarregarDDL.Tabela.DepositosPedido, "", True, Parametros)

        Parametros.Clear()
        Parametros.Add("Empresa", objNotaFiscal.CodigoEmpresa)
        Parametros.Add("EndEmpresa", objNotaFiscal.EnderecoEmpresa)
        Parametros.Add("Pedido", objNotaFiscal.CodigoPedido)
        Parametros.Add("TipoDeposito", "OD")
        Parametros.Add("Deposito", objNotaFiscal.CodigoDestino)
        Parametros.Add("EndDeposito", objNotaFiscal.EnderecoDestino)
        ddl.Carregar(ddlOrigemDestino, CarregarDDL.Tabela.DepositosPedido, "", True, Parametros)

        Parametros.Clear()
        Parametros.Add("Empresa", objNotaFiscal.CodigoEmpresa)
        Parametros.Add("EndEmpresa", objNotaFiscal.EnderecoEmpresa)
        Parametros.Add("Pedido", objNotaFiscal.CodigoPedido)
        Parametros.Add("TipoDeposito", "TR")
        Parametros.Add("Deposito", objNotaFiscal.CodigoTransbordo)
        Parametros.Add("EndDeposito", objNotaFiscal.EnderecoTransbordo)
        ddl.Carregar(ddlTransbordo, CarregarDDL.Tabela.DepositosPedido, "", True, Parametros)

        Parametros.Clear()
        Parametros.Add("Empresa", objNotaFiscal.CodigoEmpresa)
        Parametros.Add("EndEmpresa", objNotaFiscal.EnderecoEmpresa)
        Parametros.Add("Pedido", objNotaFiscal.CodigoPedido)
        Parametros.Add("TipoDeposito", "LE")
        Parametros.Add("Deposito", objNotaFiscal.CodigoLocalEmbarque)
        Parametros.Add("EndDeposito", objNotaFiscal.EndLocalEmbarque)
        ddl.Carregar(ddlEmbarque, CarregarDDL.Tabela.DepositosPedido, "", True, Parametros)

        'btnDeposito.Enabled = ddlDeposito.Items.Count = 1
        'BtnOrigemDestino.Enabled = ddlOrigemDestino.Items.Count = 1
        'btnEmbarque.Enabled = ddlEmbarque.Items.Count = 1
        'BtnTransbordo.Enabled = ddlTransbordo.Items.Count = 1

        If objNotaFiscal.CodigoEntrega.Length > 0 Then
            Dim intIndice As Integer = ddlEntrega.Items.IndexOf(ddlEntrega.Items.FindByValue(objNotaFiscal.CodigoEntrega & "-" & CStr(objNotaFiscal.EnderecoEntrega)))
            If intIndice = -1 Then
                Funcoes.AdicionarClienteAoDDL(ddlEntrega, objNotaFiscal.Entrega)
            End If
            ddlEntrega.SelectedValue = objNotaFiscal.CodigoEntrega & "-" & objNotaFiscal.EnderecoEntrega
        End If

        If objNotaFiscal.CodigoDeposito.Length > 0 Then
            ddlDeposito.SelectedValue = objNotaFiscal.CodigoDeposito & "-" & objNotaFiscal.EnderecoDeposito
        ElseIf objNotaFiscal.Empresa.Empresa.SugereDeposito And ddlDeposito.Items.Count > 1 Then
            ddlDeposito.SelectedIndex = 1
            Dim dep() As String = ddlDeposito.SelectedValue.ToString.Split("-")
            objNotaFiscal.CodigoDeposito = dep(0)
            objNotaFiscal.EnderecoDeposito = dep(1)
        End If

        If objNotaFiscal.CodigoDestino.Length > 0 Then
            ddlOrigemDestino.SelectedValue = objNotaFiscal.CodigoDestino & "-" & objNotaFiscal.EnderecoDestino
        ElseIf objNotaFiscal.Empresa.Empresa.SugereDeposito And ddlOrigemDestino.Items.Count > 1 Then
            ddlOrigemDestino.SelectedIndex = 1
            Dim depOri() As String = ddlOrigemDestino.SelectedValue.ToString.Split("-")
            objNotaFiscal.CodigoDestino = depOri(0)
            objNotaFiscal.EnderecoDestino = depOri(1)
        End If

        If objNotaFiscal.CodigoLocalEmbarque.Length > 0 Then
            ddlEmbarque.SelectedValue = objNotaFiscal.CodigoLocalEmbarque & "-" & objNotaFiscal.EndLocalEmbarque
        End If

        If objNotaFiscal.CodigoTransbordo.Length > 0 Then
            ddlTransbordo.SelectedValue = objNotaFiscal.CodigoTransbordo & "-" & objNotaFiscal.EnderecoTransbordo
        End If
    End Sub

    Public Sub CarregarFixacao()
        If Not Session("objFixacao" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()
            txtDataDeEmissao.Enabled = False
            txtDataDeEntrada.Enabled = False
            btnProcuracao.Enabled = False
            chk_NossaEmissao.Checked = objNotaFiscal.NossaEmissao

            Dim chave As String() = Session("objFixacao" & HID.Value).ToString.Split(";")
            Dim Ped As String = chave(0)
            Dim Fix As String = chave(1)
            Dim Prod As String = chave(2)

            objNotaFiscal.CodigoPedido = Ped

            '***********************************************************************************
            '** Como a consulta pode ter sido consolidada Valida novamente cliente da Fixacao **
            '** e carrega novamente na nota    *************************************************
            '***********************************************************************************
            If Not ValidarCertidaoClienteNota() Then
                Exit Sub
            End If
            objNotaFiscal.CodigoCliente = objNotaFiscal.Pedido.CodigoCliente
            objNotaFiscal.EnderecoCliente = objNotaFiscal.Pedido.EnderecoCliente
            SetarCliente(objNotaFiscal.Cliente)
            '***********************************************************************************
            '***********************************************************************************

            If objNotaFiscal.Pedido.Itens(0).Fixacoes.Count > 0 AndAlso objNotaFiscal.Pedido.Itens(0).Fixacoes(0).IndiceFixado > 0 Then
                txtPedido.Text = objNotaFiscal.CodigoPedido & " - Indice Fixado " & objNotaFiscal.Pedido.Itens(0).Fixacoes(0).IndiceFixado
            Else
                txtPedido.Text = objNotaFiscal.CodigoPedido
            End If

            objNotaFiscal.CodigoDeposito = objNotaFiscal.CodigoEmpresa
            objNotaFiscal.EnderecoDeposito = objNotaFiscal.EnderecoEmpresa
            objNotaFiscal.CodigoDestino = objNotaFiscal.CodigoCliente
            objNotaFiscal.EnderecoDestino = objNotaFiscal.EnderecoCliente

            If Not objNotaFiscal.Pedido.Finalidade Is Nothing Then
                SelecionarIndiceCombo(cmbFinalidade, objNotaFiscal.Pedido.Finalidade.Codigo)
                objNotaFiscal.CodigoFinalidade = objNotaFiscal.Pedido.Finalidade.Codigo
            Else
                SelecionarIndiceCombo(cmbFinalidade, 1)
                objNotaFiscal.CodigoFinalidade = 1
            End If
            objNotaFiscal.CodigoSituacao = 1

            Dim PxIxF As [Lib].Negocio.Fixacao = Nothing
            For Each row As [Lib].Negocio.Fixacao In objNotaFiscal.Pedido.Itens(0).Fixacoes
                If row.Codigo = Fix Then
                    PxIxF = row
                    Exit For
                End If
            Next

            Dim nFix As Decimal = 0
            For Each row In objNotaFiscal.Pedido.Itens(0).Fixacoes
                If row.Codigo = Fix Then
                    For f As Integer = 0 To row.NotasFixacao.Count - 1
                        nFix += row.NotasFixacao(f).ValorFixacao
                    Next
                End If
            Next

            If objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso objNotaFiscal.SubOperacao.Devolucao AndAlso Not nFix > PxIxF.TotalOficial Then
                LimparGrid()
                MsgBox(Me.Page, "Valor da Fixação " & PxIxF.TotalOficial & " está maior que o valor de Notas à Fixar " & nFix)
                Exit Sub
            ElseIf objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso objNotaFiscal.SubOperacao.Devolucao AndAlso Not nFix < PxIxF.TotalOficial Then
                LimparGrid()
                MsgBox(Me.Page, "Valor das Notas à Fixar " & nFix & " está maior que o valor da Fixação " & PxIxF.TotalOficial)
                Exit Sub
            End If

            If Not objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
                objNotaFiscal.CodigoOperacao = PxIxF.CodigoOperacao
                objNotaFiscal.CodigoSubOperacao = PxIxF.CodigoSubOperacao
            End If

            objNotaFiscal.CarregandoItens = True
            objNotaFiscal.Itens.Clear()

            '**********************************************************************************************
            '************************* Carrega Item da Nota Fiscal ****************************************
            '**********************************************************************************************
            Dim objItemNF As New [Lib].Negocio.NotaFiscalXItem(objNotaFiscal)

            Dim Obs As String = ""

            With objItemNF
                .CodigoProduto = PxIxF.ItemPedido.CodigoProduto
                .CodigoPedido = objNotaFiscal.CodigoPedido
                .CodigoFixacao = Fix
                .PesoQuantidade = objItemNF.Produto.PesoQuantidade
                .Sequencia = 0
                .SaldoPedidoFiscal = 0
                .SaldoPedidoFisico = 0
                .PesoFiscal = 0
                .QuantidadeFisica = 0
                .QuantidadeFiscal = 0
                .SaldoValorOficial = 0
                .SaldoValorMoeda = 0
                .QuantidadeDeEmbalagem = 0
                .UnitarioMoeda = 0
                .Unitario = 0
                .ValorTotal = PxIxF.TotalOficial
                .CodigoOperacao = objNotaFiscal.CodigoOperacao
                .CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao

                '***********************************************************************************************
                '************************  Carregar encargos da fixacao no item da nota ************************
                '***********************************************************************************************
                Dim EncAFIXAR As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
                Dim EncPRODUTO As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
                Dim EncFABOV As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
                Dim EncFACS As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
                Dim EncFETHAB As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
                Dim EncIAGRO As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
                Dim EncFUNRURAL As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
                Dim EncSENAR As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
                Dim EncFUNDERSUL As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
                Dim EncFUNDEMS As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
                Dim VlrFunrural As Decimal = Decimal.Zero
                Dim VlrSenar As Decimal = Decimal.Zero

                'Carrega os encargos para o Item da Nota
                objItemNF.CarregandoEncargos = True
                .Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objItemNF, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})

                If .Encargos.Count = 0 Then
                    LimparGrid()
                    MsgBox(Me.Page, "O Produto " & .Produto.Nome & ", não tem encargos cadastrados na Operação:" & .CodigoOperacao & "-" & .CodigoSubOperacao)
                    Exit Sub
                End If

                'Atualiza os valores dos encargos com os valores da fixacao
                For Each NFxE As [Lib].Negocio.NotaFiscalXItemXEncargo In objItemNF.Encargos
                    For Each PIFxEnc As [Lib].Negocio.FixacaoXEncargo In PxIxF.Encargos
                        If NFxE.Codigo = PIFxEnc.CodigoEncargo Then
                            If NFxE.Codigo = "PRODUTO" Then EncPRODUTO = NFxE
                            If NFxE.Codigo = "FACS" Then EncFACS = NFxE
                            If NFxE.Codigo = "FETHAB" Then EncFETHAB = NFxE
                            If NFxE.Codigo = "IAGRO" Then EncIAGRO = NFxE
                            If NFxE.Codigo = "FETHAB GADO" Then EncFETHAB = NFxE
                            If NFxE.Codigo = "FUNRURAL" Then EncFUNRURAL = NFxE
                            If NFxE.Codigo = "FUNRURAL JUDICIAL" Then EncFUNRURAL = NFxE
                            If NFxE.Codigo = "FABOV" Then EncFABOV = NFxE
                            If NFxE.Codigo = "SENAR" Then EncSENAR = NFxE
                            If NFxE.Codigo = "AFIXAR" Then EncAFIXAR = NFxE
                            If NFxE.Codigo = "FUNDEMS" Then EncFUNDEMS = NFxE
                            If NFxE.Codigo = "FUNDERSUL" Then EncFUNDERSUL = NFxE

                            With NFxE
                                .Percentual = PIFxEnc.Percentual
                                .PercentualExibicao = PIFxEnc.Percentual
                                If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                                    .Base = PIFxEnc.BaseOficial
                                    .Valor = PIFxEnc.ValorOficial
                                Else
                                    .Base = PIFxEnc.BaseMoeda
                                    .Valor = PIFxEnc.ValorMoeda
                                End If
                            End With
                        End If
                    Next
                Next

                'retira o encargo FACS e FETAB da Nota
                If Not EncFACS Is Nothing Then objItemNF.Encargos.Remove(EncFACS)
                If Not EncFETHAB Is Nothing Then objItemNF.Encargos.Remove(EncFETHAB)
                If Not EncIAGRO Is Nothing Then objItemNF.Encargos.Remove(EncIAGRO)
                'Guarda o Valor original do Funrural
                If Not EncFUNRURAL Is Nothing Then VlrFunrural = EncFUNRURAL.Valor
                If Not EncSENAR Is Nothing Then VlrSenar = EncSENAR.Valor

                'Altera a base de calculo do encargos diminuindo a base atual pelo valor do encargo A Fixar
                Dim NovaBase As Decimal
                If objNotaFiscal.SubOperacao.Devolucao Then
                    NovaBase = nFix - EncPRODUTO.Base
                Else
                    NovaBase = EncPRODUTO.Base - EncAFIXAR.Valor
                End If
                'Dim NovaBase As Decimal = EncAFIXAR.Valor - nFix

                'Remove o Encargo a Fixar pq já fez pela Fixação a Transferência
                If Not EncAFIXAR Is Nothing Then objItemNF.Encargos.Remove(EncAFIXAR)

                objItemNF.ValorTotal = NovaBase
                'Muda o sinal do encargo a Fixar Para nao interfirir no liquido da nota
                If Not EncAFIXAR Is Nothing Then EncAFIXAR.Sinal = "="
                For Each row As [Lib].Negocio.NotaFiscalXItemXEncargo In objItemNF.Encargos
                    If Not Left(objNotaFiscal.CodigoEmpresa, 8) = "04854422" _
                       And row.Codigo <> "PRODUTO" _
                       And row.Codigo <> "LIQUIDO" _
                       And row.Codigo <> "AFIXAR" Then
                        row.Base = NovaBase
                        row.Valor = Math.Round(row.Base * (row.Percentual / 100), 2, MidpointRounding.AwayFromZero)
                    ElseIf row.Codigo <> "LIQUIDO" And row.Codigo <> "AFIXAR" Then
                        row.Base = NovaBase
                        row.Valor = Math.Round(row.Base * (row.Percentual / 100), 2, MidpointRounding.AwayFromZero)
                    End If
                Next
                objItemNF.Encargos.AtualizaLiquido()

                'Monta a Descricao da nota
                If objNotaFiscal.SubOperacao.Devolucao Then
                    Obs = " DEVOLUCAO PESO FIXADO REF. "
                Else
                    Obs = " PESO FIXADO DE " & EncAFIXAR.Base & " KGS REF."
                End If
                For Each NFxFIX As [Lib].Negocio.PedidoFixacaoXNotaFiscal In PxIxF.NotasFixacao
                    Obs &= " NF" & NFxFIX.NotaFiscalXItem.NotaFiscal.Codigo & " DE " & NFxFIX.QtdeFixacao & " KGS,"
                    'Obs &= " NF" & NFxFIX.NumeroNota & " DE " & NFxFIX.QtdeFixacao & " KGS,"
                Next
                'NAO PRECISA MOSTRAR ENCARGOS ANTECIPADOS - FURLAN - 29/03/2012
                'Obs &= vbCrLf & "ENCARGOS RECOLHIDOS ANTECIPADAMENTE:" & vbCrLf
                'If Not EncFACS Is Nothing Then Obs &= "FACS R$ " & EncFACS.Valor & vbCrLf
                'If Not EncFETHAB Is Nothing Then Obs &= "FETHAB R$ " & EncFETHAB.Valor & vbCrLf
                'If Not EncFUNRURAL Is Nothing Then Obs &= "FUNRURAL R$ " & VlrFunrural - EncFUNRURAL.Valor & vbCrLf

                .PesoFiscal = 0
                .QuantidadeFisica = 0
                .QuantidadeFiscal = 0

                objItemNF.CarregandoEncargos = False
                objNotaFiscal.Itens.Add(objItemNF)
            End With

            'INICIA OBSERVACAO NOTA FISCAL AQUI - FURLAN - 25/07-2014
            If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
                'lnkVisualizar.Parent.Parent.Parent.Visible = True
                IdVisualizarNFe.Visible = True
                IdModelNota.Visible = True
                IdVisualizar.Visible = True

                If objNotaFiscal.Empresa.Empresa.ObservacaoSefazNFE.Length > 0 Then
                    Dim obsSefaz() As String = objNotaFiscal.Empresa.Empresa.ObservacaoSefazNFE.Split(";")
                    If objNotaFiscal.Empresa.CodigoEstado = "MT" Then
                        If Left(objNotaFiscal.CodigoEmpresa, 8) = "04854422" AndAlso objNotaFiscal.Itens(0).Produto.ControlarEmbalagem Then
                            objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(1)
                        Else
                            objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(0)
                        End If
                    ElseIf objNotaFiscal.Empresa.CodigoEstado = "PR" Then
                        If objNotaFiscal.Empresa.Cidade = "UMUARAMA" Then
                            objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(0)
                        Else
                            If objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10101" OrElse
                                objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10201" OrElse
                                objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10202" OrElse
                                objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10203" Then
                                objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(1)
                            Else
                                objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(0)
                            End If
                        End If
                    End If
                End If
            End If

            If objNotaFiscal.ObservacoesDeEmbarque.Length > 0 Then
                objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "| USUARIO EMISSOR: " & objNotaFiscal.UsuarioInclusao & " - PEDIDO: " & objNotaFiscal.CodigoPedido & ". "
            Else
                objNotaFiscal.ObservacoesDeEmbarque = "USUARIO EMISSOR: " & objNotaFiscal.UsuarioInclusao & " - PEDIDO: " & objNotaFiscal.CodigoPedido & ". "
            End If

            If objNotaFiscal.Pedido.PedidoEfetivo.Length > 0 Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & " - CONFIRMACAO DE NEGOCIO: " & objNotaFiscal.Pedido.PedidoEfetivo
            If objNotaFiscal.Pedido.Contrato.Length > 0 Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & " - CONTRATO: " & objNotaFiscal.Pedido.Contrato

            Dim Certidao As New [Lib].Negocio.CertidaoNegativa(objNotaFiscal.CodigoEmpresa, False)
            If Certidao.CodigoCliente.Length > 0 Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "| Certidao Negativa de Debitos nr " & Certidao.Numero & " / Cod. de Autenticidade " & Certidao.CodigoAutenticidade

            objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & Obs

            observacaoValores.Value = String.Empty

            For Each enc In objNotaFiscal.Itens(0).Encargos
                If enc.Valor > 0 Then
                    If enc.Codigo.Contains("FUNRURAL") Then
                        observacaoValores.Value = observacaoValores.Value & "|FUNRURAL - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("FABOV") Then
                        observacaoValores.Value = observacaoValores.Value & "|FABOV - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("SENAR") Then
                        observacaoValores.Value = observacaoValores.Value & "|SENAR - " & enc.Percentual.ToString.Replace(",", ".") & " % - R$ " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("FETHAB") Then
                        observacaoValores.Value = observacaoValores.Value & "|FETHAB - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("IAGRO") Then
                        observacaoValores.Value = observacaoValores.Value & "|IAGRO - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("FACS") Then
                        observacaoValores.Value = observacaoValores.Value & "|FACS - " & enc.Percentual.ToString.Replace(",", ".") & " % - R$ " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    End If
                End If
            Next

            'Passei para cima - furlan - 04/09/2020
            'objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & observacaoValores.Value

            'objNotaFiscal.ObservacoesDoProduto = Obs
            'FINALIZA OBSERVACAO NOTA FISCAL AQUI - FURLAN - 25/07-2014

            objNotaFiscal.AtualizaTotais()
            objNotaFiscal.CarregandoItens = False

            'objNotaFiscal.AtualizarObservacoes()

            'Importação do xml de notas de entrada
            If SessaoDsXML IsNot Nothing Then
                Dim ObsTemp As String = objNotaFiscal.Observacoes
                objNotaFiscal.AtualizarObservacoes()
                objNotaFiscal.Observacoes &= ObsTemp
            Else
                objNotaFiscal.AtualizarObservacoes()
            End If

            If FinanceiroNovo Then
                objNotaFiscal.Titulos = objNotaFiscal.Pedido.Titulos
                objNotaFiscal.Titulos.NF = objNotaFiscal
            End If

            SessaoSalvaNotaFiscal()
            AtualizaFormularioComAClasse()

            Dim i As Integer = 0
            While i < gridItens.Rows.Count
                CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled = False
                CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled = False
                CType(gridItens.Rows(i).FindControl("txtTotalItem"), TextBox).Enabled = False
                i += 1
            End While

            SessaoSalvaNotaFiscal()
            Session.Remove("objFixacao" & HID.Value)
            AtualizaNossaEmissaoEletronica()
        End If
    End Sub

    Public Sub CarregarVendaAOrdem()
        If Not Session("objVendaAOrdemNXI" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()
            Session.Remove("SemVendaAOrdem" & HID.Value)
            Session.Remove("ProcurandoVendaAOrdem" & HID.Value)
            objNotaFiscal.NotaTrocaOrigem = Session("objVendaAOrdemNXI" & HID.Value)

            'SE FOR NUBA DEVE CONSIDERAR O QUE ESTIVER NO XML - FURLAN - 04-09-2025
            'SE FOR ORIX DEVE CONSIDERAR O QUE ESTIVER NO XML - FURLAN - 18-09-2025
            If (Left(objNotaFiscal.Empresa.Codigo, 8) = "53267147" OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "62747840" OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "62780383" OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "63358210") AndAlso SessaoDsXML IsNot Nothing Then
                'NÃO FAZ NADA
            Else
                objNotaFiscal.Itens.Clear()

                For Each item In objNotaFiscal.NotaTrocaOrigem.Itens
                    Dim objItemNF As New [Lib].Negocio.NotaFiscalXItem(objNotaFiscal)
                    With objItemNF
                        .CodigoProduto = item.CodigoProduto

                        .Lote = item.Lote
                        .Classificacao = item.Classificacao
                        .CodigoEmbalagem = item.CodigoEmbalagem
                        .CodigoEmbalagemIndea = item.CodigoEmbalagemIndea
                        .CodigoTipoDeEmbalagem = item.CodigoTipoDeEmbalagem
                        .CapacidadeEmbalagem = item.CapacidadeEmbalagem
                        .QuantidadeDeEmbalagem = item.QuantidadeDeEmbalagem
                        .ObservacoesDoProduto = item.ObservacoesDoProduto
                        .Sequencia = item.Sequencia

                        .CodigoOperacao = objNotaFiscal.CodigoOperacao
                        .CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao

                        .CodigoPedido = objNotaFiscal.CodigoPedido
                        .PesoQuantidade = item.Produto.PesoQuantidade

                        .PesoFiscal = item.QuantidadeFiscal
                        .PesoBruto = item.QuantidadeFiscal
                        .PesoLiquido = item.QuantidadeFiscal

                        .QuantidadeFiscal = item.QuantidadeFiscal

                        .SaldoValorOficial = item.ValorTotal
                        .SaldoValorMoeda = item.ValorTotalMoeda
                        .SaldoPedidoFiscal = item.QuantidadeFiscal
                        .SaldoPedidoFisico = objItemNF.Pedido.SaldoItensPedido.Where(Function(s) s.CodigoProduto = .CodigoProduto).FirstOrDefault.SaldoQtdeDiretoFisica

                        'AJUSTE PARA NUTRI - FURLAN - 07-05-2021
                        If Left(objNotaFiscal.Empresa.Codigo, 8) = "05366261" _
                            OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "38198213" _
                            OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "62747840" _
                            OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "62780383" _
                            OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "63358210" _
                            OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "40938762" _
                            OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "44005444" _
                            OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "48984539" Then
                            If objItemNF.Pedido.SaldoItensPedido.Where(Function(s) s.CodigoProduto = .CodigoProduto).FirstOrDefault.UnidadeComercializacao = "TON" Then
                                .UnitarioMoeda = objItemNF.Pedido.SaldoItensPedido.Where(Function(s) s.CodigoProduto = .CodigoProduto).FirstOrDefault.UnitarioMoeda
                                .Unitario = objItemNF.Pedido.SaldoItensPedido.Where(Function(s) s.CodigoProduto = .CodigoProduto).FirstOrDefault.UnitarioOficial
                            Else
                                .UnitarioMoeda = objItemNF.Pedido.SaldoItensPedido.Where(Function(s) s.CodigoProduto = .CodigoProduto).FirstOrDefault.UnitarioComercializacaoMoeda
                                .Unitario = objItemNF.Pedido.SaldoItensPedido.Where(Function(s) s.CodigoProduto = .CodigoProduto).FirstOrDefault.UnitarioComercializacaoOficial
                            End If
                        Else
                            .UnitarioMoeda = item.UnitarioMoeda
                            .Unitario = item.Unitario
                        End If

                        If item.ValorLiquido = item.ValorTotal Then
                            .ValorTotal = .QuantidadeFiscal * .Unitario
                        Else
                            .ValorTotal = item.ValorLiquido
                            .Unitario = item.ValorLiquido / item.QuantidadeFiscal
                        End If

                        objItemNF.CarregandoEncargos = True
                        .Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objItemNF, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                        objItemNF.CarregandoEncargos = False

                        objNotaFiscal.Itens.Add(objItemNF)
                    End With
                Next
            End If

            objNotaFiscal.AtualizaTotais()

            'SE FOR NUBA DEVE CONSIDERAR O QUE ESTIVER NO XML - FURLAN - 04-09-2025
            'SE FOR ORIX DEVE CONSIDERAR O QUE ESTIVER NO XML - FURLAN - 18-09-2025
            If Left(objNotaFiscal.Empresa.Codigo, 8) = "53267147" OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "62747840" OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "62780383" OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "63358210" Then
                'NÃO FAZ NADA
            Else
                Dim tot As Decimal = 0
                For Each item In objNotaFiscal.Itens
                    tot += item.QuantidadeFiscal
                Next

                objNotaFiscal.Quantidade = tot
            End If

            '***************   Deposito   ***********************************************
            objNotaFiscal.Deposito = objNotaFiscal.NotaDeTroca.Deposito
            Dim itemDeposito As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Deposito)

            With objNotaFiscal.Deposito
                objNotaFiscal.CodigoDeposito = .Codigo
                objNotaFiscal.EnderecoDeposito = .CodigoEndereco

                'ddlDeposito.SelectedValue = objNotaFiscal.Deposito.Codigo & "-" & CStr(objNotaFiscal.Deposito.CodigoEndereco)
                SelecionarIndiceCombo(ddlDeposito, objNotaFiscal.Deposito.Codigo & "-" & CStr(objNotaFiscal.Deposito.CodigoEndereco))
            End With
            '*****************************************************************************

            objNotaFiscal.CodigoTransportador = objNotaFiscal.NotaDeTroca.CodigoTransportador
            objNotaFiscal.EnderecoTransportador = objNotaFiscal.NotaDeTroca.EnderecoTransportador

            If objNotaFiscal.NotaDeTroca.CodigoTransportador.Length > 0 Then
                objNotaFiscal.CodigoTransportador = objNotaFiscal.NotaDeTroca.CodigoTransportador
                objNotaFiscal.EnderecoTransportador = objNotaFiscal.NotaDeTroca.EnderecoTransportador
            End If

            If objNotaFiscal.NotaDeTroca.PlacaTransportador.Length > 0 Then
                objNotaFiscal.PlacaTransportador = objNotaFiscal.NotaDeTroca.PlacaTransportador
                objNotaFiscal.PlacaDetalhes = objNotaFiscal.NotaDeTroca.PlacaDetalhes
            End If

            If Not objNotaFiscal.PlacaDetalhes Is Nothing AndAlso Not objNotaFiscal.PlacaDetalhes.Motorista Is Nothing AndAlso objNotaFiscal.PlacaDetalhes.Motorista.Codigo.Length > 0 AndAlso Not objNotaFiscal.ObservacoesDeEmbarque.Contains("Motorista:") Then
                objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|Motorista: " & objNotaFiscal.PlacaDetalhes.Motorista.Nome & " - CPF " & objNotaFiscal.PlacaDetalhes.Motorista.CodigoFormatado & " - CNH " & objNotaFiscal.PlacaDetalhes.Habilitacao & " - Placa " & objNotaFiscal.PlacaDetalhes.Placa01 & " " & objNotaFiscal.PlacaDetalhes.Placa02 & " " & objNotaFiscal.PlacaDetalhes.Placa03 & " " & objNotaFiscal.PlacaDetalhes.Placa04
                'txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque
                If String.IsNullOrWhiteSpace(observacaoAutorizacao.Value) Then
                    txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
                Else
                    txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
                End If
            ElseIf Funcoes.VerificaPermissao("ContaEOrdemSemPlaca", "ACESSAR") Then
                objNotaFiscal.CIFFOB = objNotaFiscal.Pedido.FreteCIFFOB
            Else
                MsgBox(Me.Page, "Placa não foi informada na nota de Origem, ajuste e tente novamente.")
                Exit Sub
            End If

            Session.Remove("objVendaAOrdemNXI" & HID.Value)

            SessaoSalvaNotaFiscal()

            If objNotaFiscal.Itens.Count > 0 AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.Itens(0).Produto.ControlarPesagem Then
                If TemRomaneio(objNotaFiscal.CodigoPedido) = True Then
                    ListarRomaneios(objNotaFiscal.CodigoPedido)
                Else
                    BuscaClassificacao(True)
                End If
            Else
                AtualizaFormularioComAClasse()
            End If
        ElseIf Not Session("SemVendaAOrdem" & HID.Value) Is Nothing Or Not Session("ProcurandoVendaAOrdem" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.EntradaSaida = eEntradaSaida.Saida Then
                If Not Session("SemVendaAOrdem" & HID.Value) Is Nothing Then
                    MsgBox(Me.Page, "Nota Fiscal de Venda A Ordem não foi encontrada.")
                Else
                    MsgBox(Me.Page, "Nota Fiscal de Venda A Ordem não foi selecionanda.")
                End If
                LimparCampos("", "", False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
            End If
        End If
    End Sub

    Public Function CarregarNotaTroca() As Boolean
        If Not Session("objTrocaDeNotaNXI" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()
            objNotaFiscal.NotaTrocaOrigem = Session("objTrocaDeNotaNXI" & HID.Value)

            If objNotaFiscal.NotaTrocaOrigem.CodigoProcuracao > 0 Then
                objNotaFiscal.CodigoProcuracao = objNotaFiscal.NotaTrocaOrigem.CodigoProcuracao
                objNotaFiscal.CodigoFinalidade = 14
                cmbFinalidade.SelectedValue = 14
                cmbFinalidade.Enabled = False
            End If

            Session.Remove("objTrocaDeNotaNXI" & HID.Value)
            Session.Remove("ProcurandoTrocaDeNota" & HID.Value)

            'If CInt(objNotaFiscal.Itens(0).CodigoProduto) <> CInt(objNotaFiscal.NotaTroca.Itens(0).CodigoProduto) Then
            '    MsgBox(Me.Page, "O Produto da Troca de nota deve ser o mesmo da Nota Atual")
            '    objNotaFiscal.NotaTroca = Nothing
            '    LimparGrid()
            '    Exit Sub
            'End If

            Dim SemSaldoFazerParcial As Boolean = False
            If objNotaFiscal.Itens(0).SaldoPedidoFiscal < objNotaFiscal.NotaDeTroca.Itens(0).QuantidadeFiscal And
               objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS And
               objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR Then
                MsgBox(Me.Page, "A Nota Atual Não possui Saldo suficiente para a Troca de Nota. Sera realizada uma troca parcial. Verifique as Diferenças entre fisico e fiscal.")
                SemSaldoFazerParcial = True
            ElseIf objNotaFiscal.Itens(0).SaldoPedidoFiscal < objNotaFiscal.NotaDeTroca.Itens(0).QuantidadeFiscal And
                   objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida And
                   objNotaFiscal.SubOperacao.Devolucao Then
                MsgBox(Me.Page, "A Nota Atual Não possui Saldo suficiente para a Troca de Nota. Sera realizada uma troca parcial. Verifique as Diferenças entre fisico e fiscal.")
                SemSaldoFazerParcial = True
            Else
                'Dim Saldo As Double = Math.Max(objNotaFiscal.NotaTroca.Itens(0).QuantidadeFisica, objNotaFiscal.NotaTroca.Itens(0).QuantidadeFiscal)
                objNotaFiscal.Itens(0).PesoFiscal = objNotaFiscal.NotaDeTroca.Itens(0).PesoFiscal
                objNotaFiscal.Itens(0).PesoBruto = objNotaFiscal.NotaDeTroca.Itens(0).PesoBruto
                objNotaFiscal.Itens(0).PesoLiquido = objNotaFiscal.NotaDeTroca.Itens(0).PesoLiquido
                objNotaFiscal.Itens(0).QuantidadeFiscal = objNotaFiscal.NotaDeTroca.Itens(0).QuantidadeFiscal
                objNotaFiscal.Itens(0).QuantidadeFisica = objNotaFiscal.NotaDeTroca.Itens(0).QuantidadeFisica
                'objNotaFiscal.Itens(0).SaldoValorOficial = objNotaFiscal.NotaDeTroca.Itens(0).SaldoValorOficial
            End If

            '***************   Deposito   ***********************************************
            objNotaFiscal.Deposito = objNotaFiscal.NotaDeTroca.Deposito

            With objNotaFiscal.Deposito
                objNotaFiscal.CodigoDeposito = .Codigo
                objNotaFiscal.EnderecoDeposito = .CodigoEndereco
            End With
            '*****************************************************************************

            '***************   Local de Embarque   ***********************************************
            objNotaFiscal.CodigoLocalEmbarque = objNotaFiscal.NotaDeTroca.CodigoCliente
            objNotaFiscal.EndLocalEmbarque = objNotaFiscal.NotaDeTroca.EnderecoCliente

            objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|Local de Embarque: " & objNotaFiscal.NotaDeTroca.Cliente.CodigoFormatado & " " & objNotaFiscal.NotaDeTroca.Cliente.Nome & " - " & objNotaFiscal.NotaDeTroca.Cliente.Cidade & "/" & objNotaFiscal.NotaDeTroca.Cliente.CodigoEstado

            '*****************************************************************************

            '#FimBaseDeCalculo
            'objNotaFiscal.Itens(0).ValorTotal = (objNotaFiscal.Itens(0).SaldoValorOficial * objNotaFiscal.Itens(0).Unitario) / objNotaFiscal.Itens(0).Produto.BaseCalculo
            'objNotaFiscal.Itens(0).ValorTotal = (objNotaFiscal.Itens(0).SaldoValorOficial * objNotaFiscal.Itens(0).Unitario)
            objNotaFiscal.Itens(0).ValorTotal = Math.Round((objNotaFiscal.Itens(0).QuantidadeFiscal * objNotaFiscal.Itens(0).Unitario), 2, MidpointRounding.AwayFromZero)

            objNotaFiscal.CodigoTransportador = objNotaFiscal.NotaDeTroca.CodigoTransportador
            objNotaFiscal.EnderecoTransportador = objNotaFiscal.NotaDeTroca.EnderecoTransportador

            objNotaFiscal.CodigoRomaneio = objNotaFiscal.NotaDeTroca.CodigoRomaneio
            Dim n As New [Lib].Negocio.Numerador(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, 110)
            objNotaFiscal.Romaneio.Codigo = n.Sequencia + 1
            objNotaFiscal.CodigoRomaneio = n.Sequencia + 1
            objNotaFiscal.CriarRomaneio = True

            If SemSaldoFazerParcial Then
                objNotaFiscal.Romaneio.PesoBruto = objNotaFiscal.Itens(0).PesoBruto
                objNotaFiscal.Romaneio.PesoLiquido = objNotaFiscal.Itens(0).PesoLiquido
            End If

            If objNotaFiscal.NotaDeTroca.CodigoTransportador.Length > 0 Then
                objNotaFiscal.CodigoTransportador = objNotaFiscal.NotaDeTroca.CodigoTransportador
                objNotaFiscal.EnderecoTransportador = objNotaFiscal.NotaDeTroca.EnderecoTransportador
            End If

            If objNotaFiscal.NotaDeTroca.PlacaTransportador.Length > 0 Then
                objNotaFiscal.PlacaTransportador = objNotaFiscal.NotaDeTroca.PlacaTransportador

                objPlaca = New [Lib].Negocio.Placa(objNotaFiscal.PlacaTransportador)

                txtCodigoPlaca.Value = objNotaFiscal.PlacaTransportador
                txtPlacas.Text = objPlaca.Placa01.ToUpper
                txtEstadoDaPlaca.Text = objPlaca.EstadoPlaca01
                If objPlaca.Placa02.Length > 0 Then
                    txtPlacas.Text = txtPlacas.Text & " " & objPlaca.Placa02.ToUpper
                End If
                If objPlaca.Placa03.Length > 0 Then
                    txtPlacas.Text = txtPlacas.Text & " " & objPlaca.Placa03.ToUpper
                End If
                If objPlaca.Placa04.Length > 0 Then
                    txtPlacas.Text = txtPlacas.Text & " " & objPlaca.Placa04.ToUpper
                End If

                If Not objPlaca.Motorista Is Nothing AndAlso objPlaca.Motorista.Codigo.Length > 0 AndAlso Not objNotaFiscal.ObservacoesDeEmbarque.Contains("Motorista:") Then
                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|Motorista: " & objPlaca.Motorista.Nome & " - CPF " & objPlaca.Motorista.CodigoFormatado & " - CNH " & objPlaca.Habilitacao & " - Placa " & objPlaca.Placa01 & " " & objPlaca.Placa02 & " " & objPlaca.Placa03 & " " & objPlaca.Placa04
                    'txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque
                    If String.IsNullOrWhiteSpace(observacaoAutorizacao.Value) Then
                        txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
                    Else
                        txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
                    End If
                End If
            End If

            SessaoSalvaNotaFiscal()

            AtualizaFormularioComAClasse()

            If objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso
                (objNotaFiscal.SubOperacao.QuantidadeFisico OrElse objNotaFiscal.SubOperacao.QuantidadeFiscal) Then BuscaAutorizacao()
            Return True
        ElseIf Not Session("SemTrocaDeNota" & HID.Value) Is Nothing Or Not Session("ProcurandoTrocaDeNota" & HID.Value) Is Nothing Then
            Session.Remove("SemTrocaDeNota" & HID.Value)
            Session.Remove("ProcurandoTrocaDeNota" & HID.Value)

            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.Itens(0).Produto.ControlarPesagem Then
                BuscaClassificacao(True)
            Else
                BuscaProcuracao()
            End If
            Return True
        End If
        Return False
    End Function

    Private Function CarregarItensComSaldo(ByVal ListaProdutos As [Lib].Negocio.SaldoPedido2015, ByRef erroMsg As String, ByRef MsgAlerta As String) As Boolean
        Dim Seq As Integer = 0
        Dim ind As Integer = 1
        Dim ValorUnitarioOficial As Decimal
        Dim ValorUnitarioMoeda As Decimal

        objNotaFiscal.CarregandoItens = True

        If SessaoDsXML Is Nothing Then
            objNotaFiscal.Itens.Clear()
        End If

        Dim ListaDeProdutosSelecionados = ListaProdutos.Itens.Where(Function(s) s.Selecionado).ToList

        '**************************************************************************************************
        '************************* 1 - Carrega Item da Nota Fiscal ****************************************
        '**************************************************************************************************
        For Each row As [Lib].Negocio.SaldoPedido2015 In ListaDeProdutosSelecionados
            If ind > 0 Then
                If row.CodigoProduto = ListaDeProdutosSelecionados(ind - 1).CodigoProduto Then
                    Seq += 1
                Else
                    Seq = 0
                End If
            End If
            ind += 1

            '**************************************************************************************************
            '************************* 2 - Carrega Valor unitario Oficial / Moeda    **************************
            '**************************************************************************************************
            'Qual o Indice
            objNotaFiscal.VerificarIndice() 'obsoleto?????

            Try

                If row.XmlvUnCom > 0 Then
                    ValorUnitarioOficial = row.XmlvUnCom
                    ValorUnitarioMoeda = row.UnitarioMoeda
                Else

                    'Nao Tem Qtde na nota
                    If Not objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                        ValorUnitarioOficial = 0
                        ValorUnitarioMoeda = 0
                        'DEPOSITO NAO É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS And Not objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = row.UnitarioOficial
                        ValorUnitarioMoeda = row.UnitarioMoeda
                        'DEPOSITO É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS And objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round(row.VlrNotaOficialDepositoBruto / row.QtdeEntregueFiscalDeposito, 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round(row.VlrNotaMoedaDepositoBruto / row.QtdeEntregueFiscalDeposito, 10, MidpointRounding.AwayFromZero)
                        'GLOBAL NAO É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL And Not objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round(row.SaldoValorOficialGlobalDireto / row.SaldoQtdeGlobalFiscal, 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round(row.SaldoValorMoedaGlobalDireto / row.SaldoQtdeGlobalFiscal, 10, MidpointRounding.AwayFromZero)
                        'GLOBAL É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL And objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round(row.VlrNotaOficialGlobalBruto / row.QtdeEntregueFiscalGlobal, 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round(row.VlrNotaMoedaGlobalBruto / row.QtdeEntregueFiscalGlobal, 10, MidpointRounding.AwayFromZero)
                        'REMESSA NAO É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REMESSAS And Not objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round(row.SaldoValorOficialRemessa / row.SaldoQtdeRemessaFiscal, 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round(row.SaldoValorMoedaRemessa / row.SaldoQtdeRemessaFiscal, 10, MidpointRounding.AwayFromZero)
                        'REMESSA É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REMESSAS And objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round(row.VlrNotaOficialRemessaBruto / row.QtdeEntregueFiscalRemessa, 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round(row.VlrNotaMoedaRemessaBruto / row.QtdeEntregueFiscalRemessa, 10, MidpointRounding.AwayFromZero)
                        'AFIXAR NAO É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR And Not objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = row.UnitarioOficial
                        ValorUnitarioMoeda = row.UnitarioMoeda
                        'AFIXAR É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR And objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round((row.VlrNotaOficialAFixarBruto - row.VlrFixacaoOficial) / (row.QtdeEntregueFiscalAFixar - row.QtdeFixacao), 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round((row.VlrNotaMoedaAFixarBruto - row.VlrFixacaoMoeda) / (row.QtdeEntregueFiscalAFixar - row.QtdeFixacao), 10, MidpointRounding.AwayFromZero)
                    ElseIf objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round(row.VlrNotaOficialDiretaBruto / row.QtdeEntregueFiscalDireta, 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round(row.VlrNotaMoedaDiretaBruto / row.QtdeEntregueFiscalDireta, 10, MidpointRounding.AwayFromZero)
                    Else

                        If row.SaldoValorOficialGlobalDireto > 0 AndAlso row.SaldoQtdeDiretoFiscal > 0 Then
                            ValorUnitarioOficial = Math.Round(row.SaldoValorOficialGlobalDireto / row.SaldoQtdeDiretoFiscal, 10, MidpointRounding.AwayFromZero)
                        End If

                        If row.SaldoValorMoedaGlobalDireto > 0 AndAlso row.SaldoQtdeDiretoFiscal > 0 Then

                            'DE INDEXADOR DO PEDIDO FOR FIXO DEVE PEGAR O UNITÁRIO EM DÓLAR DO PEDIDO - FURLAN - 28/01/2024
                            If objNotaFiscal.Pedido.CodigoIndexador = 99 OrElse objNotaFiscal.Pedido.IndexadorFixo Then
                                ValorUnitarioMoeda = row.UnitarioMoeda
                            Else
                                ValorUnitarioMoeda = Math.Round(row.SaldoValorMoedaGlobalDireto / row.SaldoQtdeDiretoFiscal, 10, MidpointRounding.AwayFromZero)
                            End If
                        End If

                    End If
                End If
            Catch ex As Exception
                erroMsg &= "Erro ao calcular o unitario do produto: " & row.CodigoProduto & " - " & row.NomeProduto & " \n "
            End Try

            '********************
            '**** 2 - FIM *******
            '********************

            Dim objItemNF As New [Lib].Negocio.NotaFiscalXItem(objNotaFiscal)

            objItemNF.CodigoProduto = row.CodigoProduto
            objItemNF.Sequencia = Seq

            If SessaoDsXML IsNot Nothing Then

                For Each itemNF In objNotaFiscal.Itens
                    If itemNF.CodigoProduto = objItemNF.CodigoProduto AndAlso itemNF.Sequencia = objItemNF.Sequencia Then
                        objItemNF = itemNF
                    ElseIf itemNF.CodigoProduto = objItemNF.CodigoProduto AndAlso objNotaFiscal.Itens.Where(Function(x) x.CodigoProduto = itemNF.CodigoProduto).Count = 1 Then
                        objItemNF = itemNF
                        objItemNF.Sequencia = Seq
                    End If
                Next

            End If

            'Levar informarção do InfaDProd do Produto para NFE caso tenha - Furlan - 21-12-2022
            If objItemNF.Produto.InfaDProd.Length > 0 Then objItemNF.ObservacoesDoProduto = objItemNF.Produto.InfaDProd

            objItemNF.CodigoPedido = ListaProdutos.CodigoPedido

            objItemNF.CodigoOperacao = objNotaFiscal.CodigoOperacao
            objItemNF.CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao

            objItemNF.PesoQuantidade = objItemNF.Produto.PesoQuantidade
            objItemNF.Lote = row.Lote
            objItemNF.Classificacao = row.Classificacao

            'objItemNF.CodigoEmbalagem = row.CodigoEmbalagem
            'objItemNF.CodigoEmbalagemIndea = row.EmbalagemIndea
            objItemNF.CodigoEmbalagem = objItemNF.Produto.CodigoEmbalagem
            objItemNF.CodigoEmbalagemIndea = objItemNF.Produto.Embalagem.EmbalagemIndea

            objItemNF.CodigoTipoDeEmbalagem = row.TipoDeEmbalagem
            objItemNF.CapacidadeEmbalagem = row.CapacidadeEmbalagem

            '**********************************************************************************************************************************
            '***************************************** 3 - SALDOS  ****************************************************************************
            '**********************************************************************************************************************************
            'Verifica o saldo do pedido de acordo com a Operacao da Nota

            'GLOBAL É DEVOLUCAO
            If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL And objNotaFiscal.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalGlobal - row.QtdeEntregueFiscalRemessa
                objItemNF.SaldoPedidoFisico = row.QtdeEntregueFiscalGlobal - row.QtdeEntregueFiscalRemessa
                objItemNF.SaldoValorOficial = row.VlrNotaOficialGlobalBruto - row.VlrNotaOficialRemessaBruto
                objItemNF.SaldoValorMoeda = row.VlrNotaMoedaGlobalBruto - row.VlrNotaMoedaRemessaBruto
                'GLOBAL NAO É DEVOLUCAO
            ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL And Not objNotaFiscal.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = row.SaldoQtdeGlobalFiscal
                objItemNF.SaldoPedidoFisico = row.SaldoQtdeGlobalFiscal
                objItemNF.SaldoValorOficial = row.SaldoValorOficialGlobalDireto
                objItemNF.SaldoValorMoeda = row.SaldoValorMoedaGlobalDireto
                'REMESSA É DEVOLUCAO
            ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REMESSAS And objNotaFiscal.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalRemessa
                objItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoRemessa
                objItemNF.SaldoValorOficial = row.VlrNotaOficialRemessaBruto
                objItemNF.SaldoValorMoeda = row.VlrNotaMoedaRemessaBruto
                'REMESSA NAO É DEVOLUCAO
            ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REMESSAS And Not objNotaFiscal.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = row.SaldoQtdeRemessaFiscal
                objItemNF.SaldoPedidoFisico = row.SaldoQtdeRemessaFisica
                objItemNF.SaldoValorOficial = row.SaldoValorOficialRemessa
                objItemNF.SaldoValorMoeda = row.SaldoValorMoedaRemessa
                'AFIXAR É DEVOLUCAO
            ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR And objNotaFiscal.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalAFixar - row.QtdeFixacao
                objItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoAFixar - row.QtdeFixacao
                'Criado um novo campo na StoredProcedure spSaldoPedido para trazer o valor
                'total da fixação, mas correspondente ao unitário da NF. 
                'O campo VlrfixacaoNF será utilizado quando for menor  do que o campo VlrfixacaoOficial senão este último será, para que o valor da devolução fique correto.
                objItemNF.SaldoValorOficial = row.VlrNotaOficialAFixarBruto - IIf(row.VlrFixacaoNF < row.VlrFixacaoOficial, row.VlrFixacaoNF, row.VlrFixacaoOficial)
                objItemNF.SaldoValorMoeda = row.VlrNotaMoedaAFixarBruto - row.VlrFixacaoMoeda

                'Devolução utilizada ou para Reajuste de unitário ou para devolução de valor.
            ElseIf objNotaFiscal.SubOperacao.Devolucao And row.Tipo = 1 AndAlso Not objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                objItemNF.SaldoPedidoFiscal = Math.Abs(row.SaldoQtdeDiretoFiscal)
                objItemNF.SaldoPedidoFisico = Math.Abs(row.SaldoQtdeDiretoFisica)
                objItemNF.SaldoValorOficial = Math.Abs(row.VlrNotaOficialDiretaBruto)
                objItemNF.SaldoValorMoeda = Math.Abs(row.VlrNotaMoedaDiretaBruto)
            ElseIf objNotaFiscal.SubOperacao.Devolucao And row.Tipo = 1 Then
                objItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalDireta
                objItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoDireta
                objItemNF.SaldoValorOficial = row.VlrNotaOficialDiretaBruto
                objItemNF.SaldoValorMoeda = row.VlrNotaMoedaDiretaBruto
            ElseIf objNotaFiscal.SubOperacao.Devolucao And row.Tipo = 3 Then
                objItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalDeposito
                objItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoDeposito
                objItemNF.SaldoValorOficial = row.VlrNotaOficialDepositoBruto
                objItemNF.SaldoValorMoeda = row.VlrNotaMoedaDepositoBruto
            Else
                'Verificar ainda pode nao estar contemplando tudo
                If row.QtdeProgramada = 0 And row.QtdeProgramadaComercializacao > 0 Then
                    objItemNF.SaldoPedidoFiscal = row.SaldoQtdeComercializacao
                    objItemNF.SaldoPedidoFisico = row.SaldoQtdeComercializacao
                    objItemNF.SaldoValorOficial = row.SaldoValorOficialGlobalDireto
                    objItemNF.SaldoValorMoeda = row.SaldoValorMoedaGlobalDireto
                Else
                    objItemNF.SaldoPedidoFiscal = row.SaldoQtdeDiretoFiscal
                    objItemNF.SaldoPedidoFisico = row.SaldoQtdeDiretoFisica
                    objItemNF.SaldoValorOficial = row.SaldoValorOficialGlobalDireto
                    objItemNF.SaldoValorMoeda = row.SaldoValorMoedaGlobalDireto
                End If
            End If

            '********************************************************************************************************************************************
            '***************************************** 3.1 - SALDOS vs ESTOQUE **************************************************************************
            '********************************************************************************************************************************************
            'Verifica o saldo do pedido de acordo com a Operacao da Nota não é maior que o saldo do produto em Estoque
            'Se o saldo Quantidade do produto em Estoque "Sem Embalagem" for menor que o saldo do pedido o saldo do pedido assume a Quantidade em Estoque
            'Acrescentei a classe CONTAEORDEM pois o estoque só deve ser checado para a VENDA OU COMPRA A ORDEM - FURLAN - 22/03/16 

            If objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL _
            And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.CONTAEORDEM _
            And objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida _
            And objItemNF.Produto.ControlarEstoque Then

                Dim SaldoProdutoEstoque As New [Lib].Negocio.ListSaldoProdutoEstoqueLoteEmbalagem(objItemNF.CodigoProduto)
                SaldoProdutoEstoque.CarregarResumoSaldoEmEstoque(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.SubOperacao)

                If objItemNF.SubOperacao.EstoqueFisico Or objItemNF.SubOperacao.EstoqueFiscal Then

                    If Not objItemNF.SubOperacao.Devolucao AndAlso objItemNF.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso objItemNF.SubOperacao.EstoqueFiscal AndAlso row.Produto.EstoqueMinimo > 0 Then

                        If (SaldoProdutoEstoque.SaldoFiscal - objItemNF.SaldoPedidoFiscal) < row.Produto.EstoqueMinimo Then

                            Dim ItemEstoqueMinimo As DataRow = CType(Session("objEstoqueMinimo" & HID.Value), DataTable).NewRow()

                            ItemEstoqueMinimo("Produto") = row.Produto.Codigo

                            If objNotaFiscal.Empresa.Empresa.UsarDescricaoProduto Then
                                ItemEstoqueMinimo("Nome") = row.Produto.Nome & "-" & row.Produto.Descricao
                            Else
                                ItemEstoqueMinimo("Nome") = row.Produto.Nome
                            End If

                            ItemEstoqueMinimo("EstoqueMinimo") = row.Produto.EstoqueMinimo.ToString("N4")
                            ItemEstoqueMinimo("Faturando") = objItemNF.SaldoPedidoFiscal
                            ItemEstoqueMinimo("Saldo") = SaldoProdutoEstoque.SaldoFiscal - objItemNF.SaldoPedidoFiscal
                            CType(Session("objEstoqueMinimo" & HID.Value), DataTable).Rows.Add(ItemEstoqueMinimo)
                        End If
                    End If

                    If SaldoProdutoEstoque.Count = 0 Then
                        objItemNF.SaldoValorOficial = 0
                        objItemNF.SaldoValorMoeda = 0
                        objItemNF.SaldoPedidoFiscal = 0
                        objItemNF.SaldoPedidoFisico = 0
                        erroMsg &= "Produto " & row.CodigoProduto & " - " & row.NomeProduto & " sem estoque. \n"
                    Else
                        Dim faltaSaldo As Boolean = False

                        If objItemNF.SubOperacao.EstoqueFisico AndAlso objItemNF.SaldoPedidoFisico > SaldoProdutoEstoque.SaldoFisico Then
                            MsgAlerta &= "O Produto " & row.CodigoProduto & " - " & row.NomeProduto & " tem em Estoque FISICO -> " & SaldoProdutoEstoque(0).SaldoFisico.ToString() & " que é menor que o saldo do Pedido -> " & objItemNF.SaldoPedidoFisico & ". Será liberado apenas o que tem em Estoque. \n"

                            faltaSaldo = True
                        End If

                        If objItemNF.SubOperacao.EstoqueFiscal AndAlso objItemNF.SaldoPedidoFiscal > SaldoProdutoEstoque.SaldoFiscal Then
                            MsgAlerta &= "O Produto " & row.CodigoProduto & " - " & row.NomeProduto & " tem em Estoque FISCAL -> " & SaldoProdutoEstoque(0).SaldoFiscal.ToString() & " que é menor que o saldo do Pedido -> " & objItemNF.SaldoPedidoFiscal & ". Será liberado apenas o que tem em Estoque. \n"

                            faltaSaldo = True
                        End If

                        If faltaSaldo Then
                            objItemNF.SaldoPedidoFiscal = SaldoProdutoEstoque.SaldoFiscal
                            objItemNF.SaldoPedidoFisico = SaldoProdutoEstoque.SaldoFisico
                            If ValorUnitarioOficial > 0 Then
                                objItemNF.SaldoValorOficial = Math.Round(SaldoProdutoEstoque.SaldoFiscal * ValorUnitarioOficial, 2, MidpointRounding.AwayFromZero)
                                objItemNF.SaldoValorMoeda = Math.Round(SaldoProdutoEstoque.SaldoFiscal * ValorUnitarioMoeda, 2, MidpointRounding.AwayFromZero)
                            End If
                        End If
                    End If
                End If
            End If

            'Na Devolucao Apura o estoque do produto de acordo com o lote/classificacao que foi recebido
            'Acrescentei a classe CONTAEORDEM pois o estoque só deve ser checado para a VENDA OU COMPRA A ORDEM - FURLAN - 22/03/16 
            If Not objNotaFiscal.SubOperacao.Devolucao _
         And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL _
         And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.CONTAEORDEM _
         And objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida _
         And objItemNF.Produto.ControlarEstoque _
         And (objItemNF.Produto.ControlarEmbalagem Or objItemNF.Produto.ControlarLote) _
         And objNotaFiscal.SubOperacao.EstoqueFiscal _
        Then
                Dim SaldoProdutoEstoque As New [Lib].Negocio.ListSaldoProdutoEstoqueLoteEmbalagem(objItemNF.CodigoProduto)
                SaldoProdutoEstoque.CarregaSaldoDisponivelSaidaLoteClassificao(objNotaFiscal.SubOperacao, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objItemNF.Lote, objItemNF.Classificacao, objItemNF.CodigoEmbalagem, objItemNF.CodigoTipoDeEmbalagem, objItemNF.CapacidadeEmbalagem)

                If SaldoProdutoEstoque.Count = 0 Then
                    objItemNF.SaldoPedidoFiscal = 0
                    objItemNF.SaldoPedidoFisico = 0
                    objItemNF.SaldoValorOficial = 0
                    objItemNF.SaldoValorMoeda = 0
                    erroMsg &= "Produto não pode ser liberado, saldo em Estoque = 0. \n" & vbCrLf
                Else
                    Dim faltaSaldo As Boolean = False

                    If objItemNF.SubOperacao.EstoqueFiscal AndAlso objItemNF.SaldoPedidoFiscal > SaldoProdutoEstoque(0).SaldoFiscal Then
                        MsgAlerta &= "O Produto " & row.CodigoProduto & " - " & row.NomeProduto & " tem em Estoque FISCAL -> " & SaldoProdutoEstoque(0).SaldoFiscal.ToString() & " que é menor que o saldo do Pedido -> " & SaldoProdutoEstoque(0).SaldoFiscal & ". Será liberado apenas o que tem em Estoque. \n"

                        faltaSaldo = True
                    End If

                    If objItemNF.SubOperacao.EstoqueFisico AndAlso objItemNF.SaldoPedidoFisico > SaldoProdutoEstoque(0).SaldoFisico Then
                        MsgAlerta &= "O Produto " & row.CodigoProduto & " - " & row.NomeProduto & " tem em Estoque FISICO -> " & SaldoProdutoEstoque(0).SaldoFiscal.ToString() & " que é menor que o saldo do Pedido -> " & SaldoProdutoEstoque(0).SaldoFiscal & ". Será liberado apenas o que tem em Estoque. \n"

                        faltaSaldo = True
                    End If

                    If faltaSaldo Then
                        objItemNF.SaldoPedidoFiscal = SaldoProdutoEstoque(0).SaldoFiscal
                        objItemNF.SaldoPedidoFisico = SaldoProdutoEstoque(0).SaldoFisico
                        objItemNF.SaldoValorOficial = Math.Round(SaldoProdutoEstoque(0).SaldoFiscal * ValorUnitarioOficial, 2, MidpointRounding.AwayFromZero)
                        objItemNF.SaldoValorMoeda = Math.Round(SaldoProdutoEstoque(0).SaldoFiscal * ValorUnitarioMoeda, 2, MidpointRounding.AwayFromZero)
                    End If
                End If
            End If
            '********************
            '**** 3 - FIM *******
            '********************

            If SessaoDsXML Is Nothing Then

                If row.XmlqCom > 0 Then
                    objItemNF.PesoFiscal = row.XmlqCom

                    If objItemNF.SubOperacao.QuantidadeFisico Then
                        objItemNF.QuantidadeFisica = row.XmlqCom
                    Else
                        objItemNF.QuantidadeFisica = 0
                    End If

                    objItemNF.QuantidadeFiscal = row.XmlqCom

                    If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS _
                        OrElse objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR Then
                        row.QtdeProgramada = objItemNF.QuantidadeFiscal
                    End If
                Else
                    If Not objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                        objItemNF.PesoFiscal = 0
                        objItemNF.QuantidadeFisica = 0
                        objItemNF.QuantidadeFiscal = 0
                        objNotaFiscal.Quantidade = 0
                    Else
                        objItemNF.PesoFiscal = objItemNF.SaldoPedidoFiscal
                        objItemNF.QuantidadeFiscal = objItemNF.SaldoPedidoFiscal
                        objItemNF.QuantidadeFisica = IIf(objItemNF.SubOperacao.QuantidadeFisico, objItemNF.SaldoPedidoFiscal, 0)
                        objNotaFiscal.Quantidade += objItemNF.QuantidadeFiscal
                    End If

                End If

            Else

                objItemNF.PesoFiscal = objItemNF.QuantidadeFiscal

                If objItemNF.SubOperacao.QuantidadeFisico Then
                    objItemNF.QuantidadeFisica = objItemNF.QuantidadeFiscal
                Else
                    objItemNF.QuantidadeFisica = 0
                End If

            End If

            'Esta em duplicidade abaixo comentado em 24-04-2015 passou de 2 meses dessa data pode apagar
            'If objItemNF.CapacidadeEmbalagem > 0 Then
            '    objItemNF.QuantidadeDeEmbalagem = objItemNF.QuantidadeFiscal / objItemNF.CapacidadeEmbalagem
            'Else
            '    objItemNF.QuantidadeDeEmbalagem = 0
            'End If

            If SessaoDsXML Is Nothing Then

                If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    objItemNF.UnitarioMoeda = IIf(ValorUnitarioMoeda < 0, 0, ValorUnitarioMoeda)
                    objItemNF.Unitario = IIf(ValorUnitarioOficial < 0, 0, ValorUnitarioOficial)
                Else
                    If objNotaFiscal.Pedido.CodigoIndexador = 99 OrElse objNotaFiscal.Pedido.IndexadorFixo Then
                        objItemNF.Unitario = Math.Round(objNotaFiscal.Pedido.IndiceFixado * ValorUnitarioMoeda, 10, MidpointRounding.AwayFromZero)
                    Else
                        Dim indiceDolar As Decimal = New Cotacao(objNotaFiscal.Pedido.CodigoIndexador, objNotaFiscal.DataNota).Indice 'PTAX

                        objItemNF.Unitario = Math.Round(indiceDolar * ValorUnitarioMoeda, 10, MidpointRounding.AwayFromZero)
                    End If

                    objItemNF.UnitarioMoeda = IIf(ValorUnitarioMoeda < 0, 0, ValorUnitarioMoeda)

                End If

                If row.XmlvProd > 0 Then
                    objItemNF.ValorTotal = row.XmlvProd
                    objItemNF.ValorTotalMoeda = 0
                Else
                    'Quando a nota é somente valor, ela nao tem cotacao e o valor é somente em reais
                    If Not objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                        'Pendente alimentar o valor
                        objItemNF.ValorTotal = objItemNF.SaldoValorOficial
                        objItemNF.ValorTotalMoeda = objItemNF.SaldoValorMoeda
                    Else
                        objItemNF.ValorTotal = Math.Round(objItemNF.QuantidadeFiscal * objItemNF.Unitario, 2, MidpointRounding.AwayFromZero)
                        objItemNF.ValorTotalMoeda = Math.Round(objItemNF.QuantidadeFiscal * objItemNF.UnitarioMoeda, 2, MidpointRounding.AwayFromZero)
                    End If
                End If

            Else

                If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then

                    'objItemNF.Unitario = IIf(ValorUnitarioOficial < 0, 0, ValorUnitarioOficial)
                    objItemNF.Unitario = objItemNF.ValorTotal / objItemNF.QuantidadeFiscal

                    If objNotaFiscal.Pedido.CodigoIndexador = 99 OrElse objNotaFiscal.Pedido.IndexadorFixo Then

                        'objItemNF.Unitario = objNotaFiscal.Pedido.IndiceFixado * ValorUnitarioMoeda
                        If objItemNF.Unitario > 0 Then
                            objItemNF.UnitarioMoeda = objItemNF.Unitario / objNotaFiscal.Pedido.IndiceFixado
                            objItemNF.ValorTotalMoeda = objItemNF.ValorTotal / objNotaFiscal.Pedido.IndiceFixado
                        End If

                    Else

                        Dim indiceDolar As Decimal = New Cotacao(objNotaFiscal.Pedido.CodigoIndexador, objNotaFiscal.DataNota).Indice 'PTAX
                        'objItemNF.Unitario = indiceDolar * ValorUnitarioMoeda

                        If objItemNF.Unitario > 0 Then
                            objItemNF.UnitarioMoeda = objItemNF.Unitario / indiceDolar
                            objItemNF.ValorTotalMoeda = objItemNF.ValorTotal / indiceDolar
                        End If

                    End If


                    'objItemNF.UnitarioMoeda = IIf(ValorUnitarioMoeda < 0, 0, ValorUnitarioMoeda)

                End If

            End If

            objItemNF.CodigoOperacao = objNotaFiscal.CodigoOperacao
            objItemNF.CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao
            objItemNF.Encargos = Nothing

            '**********************************************************************************************************************************
            '*********************************************** OBSERVACOES DO PRODUTO ***********************************************************
            '**********************************************************************************************************************************
            If objItemNF.EmbalagemProduto IsNot Nothing AndAlso objItemNF.EmbalagemProduto.PesoVariavel Then
                objItemNF.QuantidadeDeEmbalagem = objItemNF.QuantidadeFiscal / objItemNF.ProdutoLoteClassificacao.PesoSaco
                objItemNF.ObservacoesDoProduto &= objItemNF.QuantidadeDeEmbalagem & " " & objItemNF.Embalagem.Descricao & " " & objItemNF.TipoDeEmbalagem.Descricao & " " & IIf(objItemNF.EmbalagemProduto.PesoVariavel, objItemNF.ProdutoLoteClassificacao.PesoSaco, objItemNF.CapacidadeEmbalagem & " " & objItemNF.Produto.Unidade) & " / " & objItemNF.Produto.DescricaoTecnica
            ElseIf objItemNF.CapacidadeEmbalagem > 0 Then
                objItemNF.QuantidadeDeEmbalagem = objItemNF.QuantidadeFiscal / objItemNF.CapacidadeEmbalagem
                objItemNF.ObservacoesDoProduto &= objItemNF.QuantidadeDeEmbalagem & " " & objItemNF.Embalagem.Descricao & " " & objItemNF.TipoDeEmbalagem.Descricao & " " & objItemNF.CapacidadeEmbalagem & " " & objItemNF.Produto.Unidade & " / " & objItemNF.Produto.DescricaoTecnica
            End If

            If objItemNF.Lote.Length > 0 Then
                Dim L As New [Lib].Negocio.Lote(objItemNF.Lote, objItemNF.CodigoProduto)
                If L.Tipo = 2 Then
                    objItemNF.ObservacoesDoProduto &= IIf(objItemNF.ObservacoesDoProduto.Length = 0, "", ", ") & "Lote: " & objItemNF.Lote & " Validade " & L.DataValidade.ToString("dd-MM-yyyy")
                Else
                    objItemNF.ObservacoesDoProduto &= IIf(objItemNF.ObservacoesDoProduto.Length = 0, "", ", ") & "Lote: " & objItemNF.Lote & " GER " & L.Germinacao & IIf(L.Pureza = 0, "", " PUR " & L.Pureza) & " Peneira/Classif. " & objItemNF.Classificacao & " Renasem " & L.Renasem & " Validade " & L.DataValidade.ToString("dd-MM-yyyy")
                End If
            End If

            If Not String.IsNullOrWhiteSpace(row.XmlinfAdProd) Then
                objItemNF.ObservacoesDoProduto &= row.XmlinfAdProd
            End If
            '**********************************************************************************************************************************
            '**********************************************************************************************************************************

            If objNotaFiscal.SubOperacao.QuantidadeFiscal And objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira And objItemNF.SaldoValorMoeda <= 0 And objItemNF.SubOperacao.Classe <> eClassesOperacoes.AFIXAR Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & " esta com saldo 0 ou negativo não pode ser usado para emissão de Nota Fiscal. Saldo: " & objItemNF.SaldoValorMoeda & " Pedido em moeda Extrangeira. \n"
            ElseIf IIf(objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, objItemNF.SaldoValorOficial, objItemNF.SaldoValorMoeda) <= 0 _
                And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS _
                And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR _
                And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.FRETES _
                And objNotaFiscal.Operacao.CodigoClasse <> eClassesOperacoes.COMPRAS.ToString _
                And Not objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE ICMS") _
                And Not objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE IPI") _
                And Not objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTO DE PESO") _
                And objNotaFiscal.SubOperacao.QuantidadeFiscal _
                And Not objNotaFiscal.SubOperacao.NotaDebito _
                And Not objNotaFiscal.SubOperacao.NotaCredito _
                And Not objNotaFiscal.SubOperacao.Devolucao Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & " esta com Saldo de Valor 0 ou negativo e não pode ser usado para emissão de Nota Fiscal. Saldo: " & objItemNF.SaldoValorOficial & " \n"
            ElseIf row.QtdeProgramadaComercializacao < 0 Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & " esta com saldo negativo na Quantidade Programada para Comercialização e não pode ser usado para emissão de Nota Fiscal. Saldo: " & row.QtdeProgramadaComercializacao & " \n"
            ElseIf objItemNF.Encargos Is Nothing OrElse objItemNF.Encargos.Count = 0 Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & ", não tem encargos cadastrados na Operação:" & objItemNF.CodigoOperacao & "-" & objItemNF.CodigoSubOperacao & " \n"
            ElseIf objItemNF.CFOP = 0 Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & ", Operação:" & objItemNF.CodigoOperacao & "-" & objItemNF.CodigoSubOperacao & " está sem CFOP informada \n"
            ElseIf objItemNF.CFOP < 5000 And objNotaFiscal.EntradaSaida = eEntradaSaida.Saida Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & ", Operação " & objItemNF.CodigoOperacao & "-" & objItemNF.CodigoSubOperacao & " com CFOP " & objItemNF.CFOP & " não pode ser usada na saída. \n"
            ElseIf objItemNF.CFOP > 5000 And objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & ", Operação " & objItemNF.CodigoOperacao & "-" & objItemNF.CodigoSubOperacao & " com CFOP " & objItemNF.CFOP & " não pode ser usada na entrada. \n"
            ElseIf Not SessaoDsXML Is Nothing _
                And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS _
                And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR _
                And Not objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE ICMS") _
                And Not objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE IPI") _
                And Not objNotaFiscal.SubOperacao.Devolucao And row.QtdeProgramada < objItemNF.QuantidadeFiscal Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & " esta com saldo da quantidade Programada menor que a quantidade da Nota Fiscal. Saldo Programado: " & row.QtdeProgramada & " - Quantidade da nota: " & objItemNF.QuantidadeFiscal & " \n"
            ElseIf objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial AndAlso objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS _
                And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR _
                And objItemNF.SaldoValorOficial < objItemNF.ValorTotal _
                And Not objNotaFiscal.SubOperacao.Devolucao = True Then
                erroMsg &= "O pedido: " & objNotaFiscal.Pedido.Codigo & " com o produto: " & objItemNF.Produto.Nome & " esta com saldo de valor fiscal do pedido menor que o valor da nota Fiscal. Saldo Fiscal: " & objItemNF.SaldoValorOficial & " - Valor da nota: " & objItemNF.ValorTotal & " \n"
            ElseIf objItemNF.Encargos.EncProduto.SituacaoTributariaPISCOFINS = 0 Then
                erroMsg &= "Situação Tributária PISCOFINS não cadastrada para o Produto " & objItemNF.CodigoProduto & "-" & objItemNF.Produto.Nome & ", Operação " & objNotaFiscal.CodigoOperacao & "-" & objNotaFiscal.CodigoSubOperacao & " " & objNotaFiscal.SubOperacao.Descricao
            ElseIf (Not objNotaFiscal.SubOperacao.QuantidadeFiscal AndAlso (Math.Abs(objItemNF.SaldoValorOficial) > 0 Or objNotaFiscal.SubOperacao.Devolucao)) _
            OrElse objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE ICMS") _
            OrElse objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE IPI") _
            OrElse objNotaFiscal.SubOperacao.Descricao.ToString.Contains("TRANSFERENCIA DE CREDITO DE ICMS") _
            OrElse (objNotaFiscal.SubOperacao.Devolucao = True AndAlso objNotaFiscal.SubOperacao.QuantidadeFiscal AndAlso objItemNF.SaldoValorOficial > 0 AndAlso objItemNF.ValorTotal > 0) _
            OrElse (objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso objNotaFiscal.SubOperacao.Devolucao = True AndAlso objItemNF.SaldoValorOficial > 0) _
            OrElse (row.QtdeProgramada = 0 AndAlso row.QtdeProgramadaComercializacao > 0 AndAlso ((objNotaFiscal.SubOperacao.Devolucao AndAlso row.QtdeComercializacaoEntregue > 0) OrElse (objNotaFiscal.SubOperacao.Devolucao = False AndAlso row.SaldoQtdeComercializacao))) _
            OrElse (objItemNF.QuantidadeFiscal > 0 OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.OUTRAS OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.FRETES) _
            OrElse objNotaFiscal.SubOperacao.NotaDebito _
            OrElse objNotaFiscal.SubOperacao.NotaCredito _
            AndAlso objItemNF.Encargos.Count > 0 Then

                If SessaoDsXML Is Nothing Then
                    objNotaFiscal.Itens.Add(objItemNF)
                Else
                    For Each itemNF In objNotaFiscal.Itens
                        If itemNF.CodigoProduto = objItemNF.CodigoProduto AndAlso itemNF.Sequencia = objItemNF.Sequencia Then
                            itemNF = objItemNF
                        End If
                    Next
                End If

            ElseIf (Not objNotaFiscal.SubOperacao.NotaDebito OrElse Not objNotaFiscal.SubOperacao.NotaCredito) AndAlso objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.OUTRAS Then
                erroMsg &= "O item não pôde ser adicionado, verifique as configurações de Operações, Operações X Encargos, Saldo no Extrato de Pedido e outros."
            End If
        Next

        Dim itensAgrupadoPorProduto = From item In objNotaFiscal.Itens
                                      Group item By item.CodigoProduto Into Grupo = Group
                                      Select New With {
                                          .CodigoProduto = CodigoProduto,
                                          .SaldoPedidoFiscal = Grupo.Sum(Function(x) x.SaldoPedidoFiscal),
                                          .SaldoPedidoFisico = Grupo.Sum(Function(x) x.SaldoPedidoFisico),
                                          .SaldoValorMoeda = Grupo.Sum(Function(x) x.SaldoValorMoeda),
                                          .SaldoValorOficial = Grupo.Sum(Function(x) x.SaldoValorOficial),
                                          .Itens = Grupo.ToList()
                                      }

        For Each itemNF In objNotaFiscal.Itens

            itemNF.CodigoOperacao = objNotaFiscal.CodigoOperacao
            itemNF.CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao

            itemNF.PesoFiscal = itemNF.QuantidadeFiscal

            If itemNF.SubOperacao.QuantidadeFisico Then

                itemNF.QuantidadeFisica = itemNF.QuantidadeFiscal

                If objNotaFiscal.CriarRomaneio = False Then

                    Dim pesoDoItem = itemNF.Produto.UnidadesDeComercializacao.Where(Function(p) p.CodigoUnidade = itemNF.Produto.Unidade).FirstOrDefault().FatorConversao
                    If pesoDoItem > 0 Then
                        'objNotaFiscal.Itens(0).PesoFiscal = (objNotaFiscal.Itens(0).SaldoPedidoFiscal * pesoDoItem)
                        itemNF.PesoFiscal = itemNF.QuantidadeFiscal * pesoDoItem
                        itemNF.PesoLiquido = itemNF.QuantidadeFiscal * pesoDoItem
                        itemNF.PesoBruto = itemNF.QuantidadeFiscal * pesoDoItem
                    End If

                End If

            End If

            For Each itemSaldo In itensAgrupadoPorProduto.Where(Function(x) x.CodigoProduto = itemNF.CodigoProduto)

                itemNF.SaldoPedidoFiscal = itemSaldo.SaldoPedidoFiscal
                itemNF.SaldoPedidoFisico = itemSaldo.SaldoPedidoFisico
                itemNF.SaldoValorMoeda = itemSaldo.SaldoValorMoeda
                itemNF.SaldoValorOficial = itemSaldo.SaldoValorOficial


                itemSaldo.SaldoPedidoFiscal -= itemNF.PesoFiscal
                itemSaldo.SaldoPedidoFisico -= itemNF.PesoFiscal
                itemSaldo.SaldoValorMoeda -= itemNF.ValorTotal
                itemSaldo.SaldoValorOficial -= itemNF.ValorTotalMoeda

            Next

            If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then

                If SessaoDsXML Is Nothing Then

                    If objNotaFiscal.Pedido.CodigoIndexador = 99 OrElse objNotaFiscal.Pedido.IndexadorFixo Then
                        itemNF.Unitario = Math.Round(objNotaFiscal.Pedido.IndiceFixado * ValorUnitarioMoeda, 10, MidpointRounding.AwayFromZero)
                    Else
                        Dim indiceDolar As Decimal = New Cotacao(objNotaFiscal.Pedido.CodigoIndexador, objNotaFiscal.DataNota).Indice 'PTAX

                        itemNF.Unitario = Math.Round(indiceDolar * ValorUnitarioMoeda, 10, MidpointRounding.AwayFromZero)
                    End If

                    itemNF.UnitarioMoeda = IIf(ValorUnitarioMoeda < 0, 0, ValorUnitarioMoeda)

                Else

                    itemNF.Unitario = Math.Round(itemNF.ValorTotal / itemNF.QuantidadeFiscal, 10, MidpointRounding.AwayFromZero)

                    If objNotaFiscal.Pedido.CodigoIndexador = 99 OrElse objNotaFiscal.Pedido.IndexadorFixo Then

                        itemNF.UnitarioMoeda = Math.Round(itemNF.Unitario / objNotaFiscal.Pedido.IndiceFixado, 10, MidpointRounding.AwayFromZero)
                        itemNF.ValorTotalMoeda = Math.Round(itemNF.ValorTotal / objNotaFiscal.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)

                    Else

                        Dim indiceDolar As Decimal = New Cotacao(objNotaFiscal.Pedido.CodigoIndexador, objNotaFiscal.DataNota).Indice 'PTAX
                        itemNF.UnitarioMoeda = Math.Round(itemNF.Unitario / indiceDolar, 10, MidpointRounding.AwayFromZero)
                        itemNF.ValorTotalMoeda = Math.Round(itemNF.ValorTotal / indiceDolar, 2, MidpointRounding.AwayFromZero)

                    End If

                End If

            End If

            Try

                If itemNF.Encargos Is Nothing Then

                    'Poder instanciar os encargos dos itens que faltam

                End If

            Catch ex As Exception

                If SessaoDsXML Is Nothing Then

                    Throw New Exception(ex.Message)

                Else

                    If String.IsNullOrWhiteSpace(itemNF.CodigoProduto) OrElse ListaDeProdutosSelecionados.Count <> objNotaFiscal.Itens.Count() Then
                        Throw New Exception("Revise o pedido, a quantidade de itens do pedido está difente da quantidade de itens importados!")
                    End If

                End If

            End Try

        Next

        If objNotaFiscal.Itens.Count = 0 Then
            LimparGrid()
            Return False
        ElseIf Not String.IsNullOrWhiteSpace(erroMsg) Then
            LimparGrid()
            Return False
        End If

        objNotaFiscal.CarregandoItens = False
        objNotaFiscal.AtualizaTotais()
        'objNotaFiscal.CarregandoItens = False

        'Importação do xml de notas de entrada
        If SessaoDsXML IsNot Nothing Then
            Dim ObsTemp As String = objNotaFiscal.Observacoes
            objNotaFiscal.AtualizarObservacoes()
            objNotaFiscal.Observacoes &= ObsTemp
        Else
            objNotaFiscal.AtualizarObservacoes()
        End If

        'Se não tiver controle de lotes limpamos os lotes que são importados no XML
        If SessaoDsXML IsNot Nothing Then
            For Each itemNota In objNotaFiscal.Itens
                If itemNota.Produto.ControlarLote = False AndAlso Not itemNota.Lotes Is Nothing AndAlso itemNota.Lotes.Count > 0 Then
                    itemNota.Lotes = Nothing
                End If
            Next
        End If

        SessaoSalvaNotaFiscal()

        'kitio
        If objNotaFiscal.SubOperacao.Financeiro Then 'AndAlso objNotaFiscal.Pedido.MomentoFinanceiro <> 2
            TabVencimentosOld.Visible = Not FinanceiroNovo
            TabVencimentos.Visible = FinanceiroNovo
            If FinanceiroNovo Then
                ucFinanceiro.CarregarResumo()
                ucFinanceiro.AtualizarValorNotaOuFixacaoOuTroca()
            End If
        Else
            TabVencimentosOld.Visible = False
            TabVencimentos.Visible = False
        End If

        '**********************
        '**** 1 - FIM *********
        '**********************
        Session.Remove("objItensPedidoSelecionadosNXI" & HID.Value)

        Return True
    End Function

    Private Sub SetarEmpresa(ByVal pEmpresa As String, ByVal pEndEmpresa As Integer, ByVal verificamodo As Boolean)
        Dim objEmpresa As New [Lib].Negocio.Cliente(pEmpresa, pEndEmpresa)

        SessaoRecuperaNotaFiscal()

        objNotaFiscal.Empresa = CType(objEmpresa, [Lib].Negocio.Cliente)

        With objNotaFiscal.Empresa
            objNotaFiscal.CodigoEmpresa = .Codigo
            objNotaFiscal.EnderecoEmpresa = .CodigoEndereco

            txtNomeDaEmpresa.Text = .Nome
            txtInscricaoDaEmpresa.Text = .InscricaoEstadual
            txtCnpjDaEmpresa.Text = .CodigoFormatado
            txtMarca.Text = .Empresa.Marca
            objNotaFiscal.Marca = .Empresa.Marca
        End With

        If objNotaFiscal.Empresa.Empresa.ObrigaNavio Then
            divNaviosXInvoice.Visible = True
        Else
            divNaviosXInvoice.Visible = False
        End If

        SessaoSalvaNotaFiscal()

        'Verifica se a empresa está habilitada para gravar arquivo
        Dim Empresa As New [Lib].Negocio.Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
        trFile.Visible = Not String.IsNullOrWhiteSpace(Empresa.Empresa.PathDownloadNFe)

        'If objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica AndAlso objNotaFiscal.Empresa.Empresa.NossaEmissao AndAlso verificamodo Then
        '    If VerificarModo(objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa, True, False) Then
        '        If btnModo.Text.ToUpper().Trim() = "MODO NORMAL" Then VerificarContingencia()
        '    End If
        'End If

        LiberaEmpresa()

    End Sub

#End Region

#Region "Grid"
    Protected Sub ItemNotaOK(ByVal intIndice As Integer)
        SessaoRecuperaNotaFiscal()

        'O'PERAÇÃO DE IMPORTAÇÃO DO NAVIO VALE O QUE ESTÁ NO PEDIDO, NÃO PODE SER ATUALIZADA - FURLAN - 14/10/2024
        If Left(objNotaFiscal.CodigoEmpresa, 8) = "24450490" AndAlso objNotaFiscal.CodigoOperacao = 35 AndAlso (objNotaFiscal.CodigoSubOperacao = 2 OrElse objNotaFiscal.CodigoSubOperacao = 11 OrElse objNotaFiscal.CodigoSubOperacao = 12) Then
            Exit Sub
        End If

        Dim objItemNF As NotaFiscalXItem = objNotaFiscal.Itens(intIndice)

        Dim CapEmbalagem As Decimal = objItemNF.CapacidadeEmbalagem

        Dim Quantidade As Decimal = CType(gridItens.Rows(intIndice).FindControl("txtQuantidadeItem"), TextBox).Text
        Dim Unitario As Decimal = CType(gridItens.Rows(intIndice).FindControl("txtUnitarioItem"), TextBox).Text

        Dim Valor As Decimal = 0

        If objItemNF.Produto.CodigoEmbalagem = 1 AndAlso Quantidade > 0 Then
            'ainda vou testar - Furlan - 23/08/2022
            Dim vString As String = CType(gridItens.Rows(intIndice).FindControl("txtQuantidadeItem"), TextBox).Text
            Dim vDecimal As Decimal = CDec(vString.Split(",")(1))

            If vDecimal > 0 Then
                MsgBox(Me.Page, "Produto a GRANEL não pode ter casa decimal.", eTitulo.Info)

                CType(gridItens.Rows(0).FindControl("txtQuantidadeItem"), TextBox).Text = Quantidade

                Exit Sub
            End If
        End If

        'Caso Tenha Devolução o valor deve ser sempre o que está na tela, pois já vem calculado de acordo com a quantidade selecionada para devolução - Furlan - 09/08/2016
        If objNotaFiscal.SubOperacao.Devolucao AndAlso Not objItemNF.NotasDevolucao Is Nothing AndAlso objItemNF.NotasDevolucao.Count > 0 Then
            Valor = CType(gridItens.Rows(intIndice).FindControl("txtTotalItem"), TextBox).Text
        ElseIf objNotaFiscal.SubOperacao.NotaCredito OrElse objNotaFiscal.SubOperacao.NotaDebito Then
            Valor = CType(gridItens.Rows(intIndice).FindControl("txtTotalItem"), TextBox).Text
        Else
            Valor = IIf(Quantidade = 0, CType(gridItens.Rows(intIndice).FindControl("txtTotalItem"), TextBox).Text, Math.Round((Quantidade * Unitario), 2, MidpointRounding.AwayFromZero))
        End If

        Dim ValorMoeda As Decimal
        If objItemNF.UnitarioMoeda = 0 Then
            ValorMoeda = 0
        Else
            ValorMoeda = IIf(Quantidade = 0, CType(gridItens.Rows(intIndice).FindControl("txtTotalItem"), TextBox).Text, Math.Round((Quantidade * objItemNF.UnitarioMoeda), 2, MidpointRounding.AwayFromZero))
        End If

        If objItemNF.QuantidadeFiscal = 0 Then
            If objNotaFiscal.Pedido.CodigoIndexador = 99 AndAlso objNotaFiscal.Pedido.IndiceFixado > 0 Then
                ValorMoeda = Math.Round(Valor / objNotaFiscal.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
            Else

                If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then

                    Dim indiceDolar As Decimal = New Cotacao(objNotaFiscal.Pedido.CodigoIndexador, objNotaFiscal.DataNota).Indice 'PTAX
                    If indiceDolar > 0 Then ValorMoeda = Math.Round(CType(gridItens.Rows(intIndice).FindControl("txtTotalItem"), TextBox).Text / indiceDolar, 2, MidpointRounding.AwayFromZero)

                Else

                    Dim indiceDolar As Decimal = New Cotacao(objNotaFiscal.Pedido.CodigoIndexador, objNotaFiscal.DataNota).Indice 'PTAX
                    If indiceDolar > 0 Then ValorMoeda = Math.Round(Valor / indiceDolar, 2, MidpointRounding.AwayFromZero)

                End If

            End If
        End If

        objItemNF.IUD = "U"
        objItemNF.Encargos.Clear()
        For Each enc In objItemNF.Encargos
            enc = Nothing
        Next
        objItemNF.Encargos = Nothing

        Dim libera As Boolean = False


        If Left(objNotaFiscal.CodigoEmpresa, 8) = "24450490" AndAlso Not objNotaFiscal.SubOperacao.QuantidadeFiscal AndAlso objNotaFiscal.SubOperacao.FinalidadeDaNota = 2 Then
            libera = True
        End If

        If objNotaFiscal.SubOperacao.NotaCredito OrElse objNotaFiscal.SubOperacao.NotaDebito Then
            libera = True
        End If

        If Not objNotaFiscal.SubOperacao.Devolucao _
          AndAlso objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR _
          AndAlso objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.CONTAEORDEM _
          AndAlso objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS _
          AndAlso Not libera _
          AndAlso (
                   (objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira AndAlso ValorMoeda > objItemNF.SaldoValorMoeda) _
                   OrElse
                   (objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial AndAlso Valor > objItemNF.SaldoValorOficial AndAlso objNotaFiscal.Operacao.CodigoClasse <> "COMPRAS")
                 ) Then
            CType(gridItens.Rows(intIndice).FindControl("txtQuantidadeItem"), TextBox).Text = objItemNF.QuantidadeFiscal
            CType(gridItens.Rows(intIndice).FindControl("txtUnitarioItem"), TextBox).Text = objItemNF.Unitario
            CType(gridItens.Rows(intIndice).FindControl("txtTotalItem"), TextBox).Text = objItemNF.ValorTotal

            If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                MsgBox(Me.Page, "O Valor " & Valor & " informado não pode ser maior que o Valor de Saldo do Pedido " & objItemNF.SaldoValorOficial)
            Else
                MsgBox(Me.Page, "O Valor em Dolar " & ValorMoeda & " informado não pode ser maior que o Valor de Saldo em dolar do Pedido " & objItemNF.SaldoValorMoeda)
            End If

            Exit Sub
            'Devolução a verificação não é sobre o saldo e sim sobre o valor adquirido
        ElseIf IIf(objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, Valor > objItemNF.ValorTotal, ValorMoeda > objItemNF.ValorTotalMoeda) _
          AndAlso objNotaFiscal.SubOperacao.Devolucao = True Then
            CType(gridItens.Rows(intIndice).FindControl("txtQuantidadeItem"), TextBox).Text = objItemNF.QuantidadeFiscal
            CType(gridItens.Rows(intIndice).FindControl("txtUnitarioItem"), TextBox).Text = objItemNF.Unitario
            CType(gridItens.Rows(intIndice).FindControl("txtTotalItem"), TextBox).Text = objItemNF.ValorTotal
            If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                MsgBox(Me.Page, "O Valor " & Valor & " informado não pode ser maior que o Valor de Saldo do Pedido para Devolução " & objItemNF.SaldoValorOficial)
            Else
                MsgBox(Me.Page, "O Valor " & ValorMoeda & " informado não pode ser maior que o Valor de Saldo do Pedido para Devolução " & objItemNF.SaldoValorMoeda)
            End If
            Exit Sub
            'Devolução a verificação não é sobre o saldo e sim sobre o valor adquirido
        ElseIf Not objNotaFiscal.SubOperacao.QuantidadeFiscal _
           AndAlso Not objNotaFiscal.SubOperacao.NotaCredito _
           AndAlso Not objNotaFiscal.SubOperacao.NotaDebito _
           AndAlso IIf(objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, Valor > objItemNF.ValorTotal, ValorMoeda > objItemNF.ValorTotalMoeda) Then
            If (objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR And objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial) _
             Or ((objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS Or objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR) _
                 And objNotaFiscal.SubOperacao.Devolucao = True) Then
                CType(gridItens.Rows(intIndice).FindControl("txtTotalItem"), TextBox).Text = objItemNF.ValorTotal
                If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    MsgBox(Me.Page, "O Valor " & Valor & " informado não pode ser maior que o Valor de Saldo para Ajuste do Pedido " & objItemNF.SaldoValorOficial)
                Else
                    MsgBox(Me.Page, "O Valor " & ValorMoeda & " informado não pode ser maior que o Valor de Saldo para Ajuste do Pedido " & objItemNF.SaldoValorMoeda)
                End If
                Exit Sub
            Else
                If objNotaFiscal.SubOperacao.QuantidadeFisico Then
                    If objNotaFiscal.CriarRomaneio Then
                        objItemNF.QuantidadeFisica = Quantidade
                        If objNotaFiscal.TemNotaTroca AndAlso objNotaFiscal.NotaDeTroca.Codigo > 0 Then
                            objNotaFiscal.Romaneio.PesoBruto = Quantidade
                            objNotaFiscal.Romaneio.PesoLiquido = Quantidade
                            txtPesoRomaneio.Text = objNotaFiscal.Romaneio.PesoLiquido
                        End If
                    ElseIf objNotaFiscal.CodigoRomaneio > 0 Then
                        If objNotaFiscal.Itens(0).Produto.ControlarRomaneio Then
                            objItemNF.QuantidadeFisica = objNotaFiscal.Romaneio.PesoLiquido
                        Else
                            objItemNF.QuantidadeFisica = Quantidade
                        End If
                    Else
                        objItemNF.QuantidadeFisica = Quantidade
                    End If
                Else
                    objItemNF.QuantidadeFisica = 0
                End If

                objItemNF.QuantidadeFiscal = Quantidade
                objItemNF.PesoFiscal = objItemNF.QuantidadeFiscal
                objItemNF.Unitario = Unitario
                objItemNF.ValorTotal = Valor
                objItemNF.ValorTotalMoeda = ValorMoeda
                If CapEmbalagem > 0 Then
                    objItemNF.QuantidadeDeEmbalagem = (Quantidade / CapEmbalagem)
                Else
                    objItemNF.QuantidadeDeEmbalagem = 0
                End If
            End If
        ElseIf Quantidade > objItemNF.SaldoPedidoFiscal And (objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR) Then
            If objNotaFiscal.SubOperacao.Devolucao Then
                CType(gridItens.Rows(intIndice).FindControl("txtQuantidadeItem"), TextBox).Text = objItemNF.QuantidadeFiscal
                CType(gridItens.Rows(intIndice).FindControl("txtUnitarioItem"), TextBox).Text = objItemNF.Unitario
                MsgBox(Me.Page, "A Quantidade informada não pode ser maior que a Quantidade Entregue do Pedido")
                Exit Sub
            Else
                If Quantidade > objItemNF.SaldoPedidoFiscal Then
                    CType(gridItens.Rows(intIndice).FindControl("txtQuantidadeItem"), TextBox).Text = objItemNF.QuantidadeFiscal
                    CType(gridItens.Rows(intIndice).FindControl("txtUnitarioItem"), TextBox).Text = objItemNF.Unitario
                    MsgBox(Me.Page, "A Quantidade informada não pode ser maior que o saldo do Pedido")
                    Exit Sub
                Else
                    CType(gridItens.Rows(intIndice).FindControl("txtTotalItem"), TextBox).Text = Valor.ToString("N2")

                    If objNotaFiscal.SubOperacao.QuantidadeFisico Then
                        If objNotaFiscal.CriarRomaneio Then
                            objItemNF.QuantidadeFisica = Quantidade
                            If objNotaFiscal.TemNotaTroca AndAlso objNotaFiscal.NotaDeTroca.Codigo > 0 Then
                                objNotaFiscal.Romaneio.PesoBruto = Quantidade
                                objNotaFiscal.Romaneio.PesoLiquido = Quantidade
                                txtPesoRomaneio.Text = objNotaFiscal.Romaneio.PesoLiquido
                            End If
                        ElseIf objNotaFiscal.CodigoRomaneio > 0 Then
                            If objNotaFiscal.Itens(0).Produto.ControlarRomaneio Then
                                objItemNF.QuantidadeFisica = objNotaFiscal.Romaneio.PesoLiquido
                            Else
                                objItemNF.QuantidadeFisica = Quantidade
                            End If
                        Else
                            objItemNF.QuantidadeFisica = Quantidade
                        End If
                    Else
                        objItemNF.QuantidadeFisica = 0
                    End If

                    objItemNF.QuantidadeFiscal = Quantidade
                    objItemNF.PesoFiscal = objItemNF.QuantidadeFiscal
                    objItemNF.Unitario = Unitario
                    objItemNF.ValorTotal = Valor
                    objItemNF.ValorTotalMoeda = ValorMoeda
                    If CapEmbalagem > 0 Then
                        objItemNF.QuantidadeDeEmbalagem = (Quantidade / CapEmbalagem)
                    Else
                        objItemNF.QuantidadeDeEmbalagem = 0
                    End If
                End If
            End If
        ElseIf Quantidade > objItemNF.SaldoPedidoFiscal And (objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS Or objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR) And objNotaFiscal.SubOperacao.Devolucao = True Then
            CType(gridItens.Rows(intIndice).FindControl("txtQuantidadeItem"), TextBox).Text = objItemNF.QuantidadeFiscal
            MsgBox(Me.Page, "A Quantidade informada não pode ser maior que o saldo do Pedido para Devolução")
            Exit Sub
        Else
            CType(gridItens.Rows(intIndice).FindControl("txtTotalItem"), TextBox).Text = Valor.ToString("N2")

            If objNotaFiscal.SubOperacao.QuantidadeFisico Then
                If objNotaFiscal.CriarRomaneio Then
                    If objNotaFiscal.CodigoRomaneio = 0 Then objItemNF.QuantidadeFisica = Quantidade
                    If objNotaFiscal.TemNotaTroca AndAlso objNotaFiscal.NotaDeTroca.Codigo > 0 Then
                        objNotaFiscal.Romaneio.PesoBruto = Quantidade
                        objNotaFiscal.Romaneio.PesoLiquido = Quantidade
                        txtPesoRomaneio.Text = objNotaFiscal.Romaneio.PesoLiquido
                    End If
                ElseIf objNotaFiscal.CodigoRomaneio > 0 Then

                    If objNotaFiscal.Itens.Count = 1 AndAlso (objNotaFiscal.Itens(0).Produto.Unidade = "KG" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "KGS" OrElse objNotaFiscal.Itens(0).Produto.Unidade = "TON") Then
                        objItemNF.QuantidadeFisica = objNotaFiscal.Romaneio.PesoLiquido
                    ElseIf objNotaFiscal.Itens(0).Produto.ControlarRomaneio Then
                        objItemNF.QuantidadeFisica = objNotaFiscal.Romaneio.PesoLiquido
                    Else
                        objItemNF.QuantidadeFisica = Quantidade
                    End If
                Else
                    objItemNF.QuantidadeFisica = Quantidade
                End If
            Else
                objItemNF.QuantidadeFisica = 0
            End If

            objItemNF.QuantidadeFiscal = Quantidade
            objItemNF.PesoFiscal = objItemNF.QuantidadeFiscal
            objItemNF.Unitario = Unitario
            objItemNF.ValorTotal = Valor
            objItemNF.ValorTotalMoeda = ValorMoeda
        End If


        If Not objNotaFiscal.Romaneio Is Nothing Then
            objNotaFiscal.PesoBruto = objNotaFiscal.Romaneio.PesoBruto
            objNotaFiscal.PesoLiquido = objNotaFiscal.Romaneio.PesoLiquido
            objNotaFiscal.Quantidade = objNotaFiscal.Romaneio.PesoLiquido
        Else
            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.QuantidadeFisica) > 0 Then
                objNotaFiscal.PesoBruto = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.QuantidadeFisica)
            Else
                objNotaFiscal.PesoBruto = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.QuantidadeFiscal)
            End If

            objNotaFiscal.PesoLiquido = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.QuantidadeFiscal)
            objNotaFiscal.Quantidade = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.QuantidadeFiscal)
        End If

        If objItemNF.EmbalagemProduto IsNot Nothing AndAlso objItemNF.EmbalagemProduto.PesoVariavel Then
            If (Quantidade Mod objItemNF.ProdutoLoteClassificacao.PesoSaco) > 0 Then
                MsgBox(Me.Page, "As Quantidade Informadas devem ser multiplos do Peso do Saco cadastrado no Lote x Peneira que é de: " & objItemNF.ProdutoLoteClassificacao.PesoSaco.ToString("N4"))
                CType(gridItens.Rows(intIndice).FindControl("txtQuantidadeItem"), TextBox).Text = objItemNF.QuantidadeFiscal
                Exit Sub
            End If
            objItemNF.QuantidadeDeEmbalagem = objItemNF.QuantidadeFiscal / objItemNF.ProdutoLoteClassificacao.PesoSaco
            objItemNF.ObservacoesDoProduto &= objItemNF.QuantidadeDeEmbalagem & " " & objItemNF.Embalagem.Descricao & " " & objItemNF.TipoDeEmbalagem.Descricao & " " & IIf(objItemNF.EmbalagemProduto.PesoVariavel, objItemNF.ProdutoLoteClassificacao.PesoSaco, objItemNF.CapacidadeEmbalagem & " " & objItemNF.Produto.Unidade) & " / " & objItemNF.Produto.DescricaoTecnica
        ElseIf objItemNF.CapacidadeEmbalagem > 0 Then
            If (Quantidade Mod CapEmbalagem) > 0 Then
                MsgBox(Me.Page, "As Quantidade Informadas devem ser multiplos da capacidade da embalagem que é de: " & Str(objItemNF.CapacidadeEmbalagem))
                CType(gridItens.Rows(intIndice).FindControl("txtQuantidadeItem"), TextBox).Text = objItemNF.QuantidadeFiscal
                Exit Sub
            End If
            objItemNF.QuantidadeDeEmbalagem = objItemNF.QuantidadeFiscal / objItemNF.CapacidadeEmbalagem
            objItemNF.ObservacoesDoProduto &= objItemNF.QuantidadeDeEmbalagem & " " & objItemNF.Embalagem.Descricao & " " & objItemNF.TipoDeEmbalagem.Descricao & " " & objItemNF.CapacidadeEmbalagem & " " & objItemNF.Produto.Unidade & " / " & objItemNF.Produto.DescricaoTecnica
        End If

        If objItemNF.Lote.Length > 0 Then
            Dim L As New [Lib].Negocio.Lote(objItemNF.Lote, objItemNF.CodigoProduto)
            If L.Tipo = 2 Then
                objItemNF.ObservacoesDoProduto &= IIf(objItemNF.ObservacoesDoProduto.Length = 0, "", ", ") & "Lote: " & objItemNF.Lote & " Validade " & L.DataValidade.ToString("dd-MM-yyyy")
            Else
                objItemNF.ObservacoesDoProduto &= IIf(objItemNF.ObservacoesDoProduto.Length = 0, "", ", ") & "Lote: " & objItemNF.Lote & " GER " & L.Germinacao & IIf(L.Pureza = 0, "", " PUR " & L.Pureza) & " Peneira/Classif. " & objItemNF.Classificacao & " Renasem " & L.Renasem & " Validade " & L.DataValidade.ToString("dd-MM-yyyy")
            End If
        End If

        objNotaFiscal.CarregandoItens = False
        objNotaFiscal.AtualizaTotais()

        'Metodo alimenta informações importantes do IBS e CBS
        objNotaFiscal.AtualizarObservacoes()
        txtObservacoesFiscais.Text = objNotaFiscal.Observacoes
        'Metodo alimenta informações importantes do IBS e CBS

        'Deve carregar os encargos sempre que alterar dados no item - Furlan
        'DgEncargos.DataSource = objItemNF.Encargos.ToArray()
        'DgEncargos.DataBind()
        AtualizaFormulario()

        SessaoSalvaNotaFiscal()

        Dim obsNota() As String = objNotaFiscal.ObservacoesDeEmbarque.Split("|")
        observacaoValores.Value = String.Empty

        If objItemNF.Encargos.Any(Function(s) s.Codigo = "FUNRURAL") OrElse
            objItemNF.Encargos.Any(Function(s) s.Codigo = "FABOV") OrElse
            objItemNF.Encargos.Any(Function(s) s.Codigo = "SENAR") OrElse
            objItemNF.Encargos.Any(Function(s) s.Codigo = "FETHAB") OrElse
            objItemNF.Encargos.Any(Function(s) s.Codigo = "IAGRO") OrElse
            objItemNF.Encargos.Any(Function(s) s.Codigo = "FACS") OrElse
            objItemNF.Encargos.Any(Function(s) s.Codigo = "FUNDEMS") OrElse
            objItemNF.Encargos.Any(Function(s) s.Codigo = "FUNDERSUL") OrElse
            objItemNF.Encargos.Any(Function(s) s.Codigo = "FETHAB GADO") Then

            'Dim ob As Integer
            'For ob = 0 To obsNota.Length - 1
            '    If obsNota(ob).Contains("FUNRURAL") And (obsNota(ob).Contains("por Unidade - R$") Or obsNota(ob).Contains("% - R$")) OrElse
            '       obsNota(ob).Contains("FABOV") And (obsNota(ob).Contains("por Unidade - R$") Or obsNota(ob).Contains("% - R$")) OrElse
            '       obsNota(ob).Contains("SENAR") And (obsNota(ob).Contains("por Unidade - R$") Or obsNota(ob).Contains("% - R$")) OrElse
            '       obsNota(ob).Contains("FETHAB") And (obsNota(ob).Contains("por Unidade - R$") Or obsNota(ob).Contains("% - R$")) OrElse
            '       obsNota(ob).Contains("FUNDEMS") And (obsNota(ob).Contains("por Unidade - R$") Or obsNota(ob).Contains("% - R$")) OrElse
            '       obsNota(ob).Contains("FUNDERSUL") And (obsNota(ob).Contains("por Unidade - R$") Or obsNota(ob).Contains("% - R$")) OrElse
            '       obsNota(ob).Contains("FACS") And (obsNota(ob).Contains("por Unidade - R$") Or obsNota(ob).Contains("% - R$")) Then
            '    Else
            '        observacaoValores.Value = observacaoValores.Value & obsNota(ob)
            '    End If
            'Next

            'If Not String.IsNullOrEmpty(observacaoValores.Value) Then objNotaFiscal.ObservacoesDeEmbarque = observacaoValores.Value

            'observacaoValores.Value = ""

            For Each enc In objItemNF.Encargos
                If enc.Codigo.Contains("FUNRURAL") Then
                    observacaoValores.Value = observacaoValores.Value & "|FUNRURAL - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                ElseIf enc.Codigo.Contains("FABOV") Then
                    observacaoValores.Value = observacaoValores.Value & "|FABOV - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                ElseIf enc.Codigo.Contains("SENAR") Then
                    observacaoValores.Value = observacaoValores.Value & "|SENAR - " & enc.Percentual.ToString("N2").Replace(",", ".") & " % - R$ = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                ElseIf enc.Codigo.Contains("FETHAB") Then
                    observacaoValores.Value = observacaoValores.Value & "|FETHAB - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                ElseIf enc.Codigo.Contains("IAGRO") Then
                    observacaoValores.Value = observacaoValores.Value & "|IAGRO - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                ElseIf enc.Codigo.Contains("FUNDEMS") Then
                    observacaoValores.Value = observacaoValores.Value & "|FUNDEMS - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                ElseIf enc.Codigo.Contains("FUNDERSUL") Then
                    observacaoValores.Value = observacaoValores.Value & "|FUNDERSUL - " & enc.Percentual.ToString("N2").Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & " = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                ElseIf enc.Codigo.Contains("FACS") Then
                    observacaoValores.Value = observacaoValores.Value & "|FACS - " & enc.Percentual.ToString.Replace(",", ".") & " % - R$ = " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                End If
            Next
        End If

        'txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & observacaoValores.Value

        If String.IsNullOrWhiteSpace(observacaoAutorizacao.Value) Then
            txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
        Else
            txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
        End If

        If objNotaFiscal.SubOperacao.Devolucao AndAlso (objItemNF.QuantidadeFiscal <> objItemNF.NotasDevolucao.SomaQtde OrElse Math.Round(objItemNF.ValorTotal, 2, MidpointRounding.AwayFromZero) <> Math.Round(objItemNF.NotasDevolucao.SomaVlr, 2, MidpointRounding.AwayFromZero)) Then
            objItemNF.NotasDevolucao = Nothing

            Session("ssCampo" & HID.Value) = "LivreClasse"
            ucNotaDeDevolucaoXNota.SetarIndice(intIndice)
            ucNotaDeDevolucaoXNota.BindGridView()
            Popup.ConsultaDeNotaDeDevolucaoXNota(Me.Page, "objNotaDeDevolucaoXNota" & HID.Value)

            MsgBox(Me.Page, "A Quantide/Valor da Nota de origem da Devolucao nao conferem com os valores da nota, O sistema ira carregar os valores padroes seguindo o metodo FIFO / PEPS.")
        End If

        'Financeiro
        If Not objNotaFiscal.VencimentosPedido Is Nothing AndAlso objNotaFiscal.VencimentosPedido.Count > 0 Then
            'gridVencimentosPedido.DataSource = objNotaFiscal.VencimentosPedido.ToArray
            'Ajuste para Selecinar o Título e carregar no Financeiro - Furlan 30/07/2025
            gridVencimentosPedido.DataSource = From tit In objNotaFiscal.VencimentosPedido
                                               Select New With {.Codigo = tit.Codigo, .ReceberPagar = tit.ReceberPagar, .DescricaoMoeda = tit.DescricaoMoeda, .DescricaoProvisao = tit.DescricaoProvisao, .Prorrogacao = tit.Prorrogacao, .ValorDoDocumento = tit.ValorDoDocumento, .MoedaValorDoDocumento = tit.MoedaValorDoDocumento, .CodigoCifrado = IIf(tit.ReceberPagar = "P", Funcoes.Cifrar("ContasAPagar-" & tit.Codigo), Funcoes.Cifrar("ContasAReceber-" & tit.Codigo))}
            gridVencimentosPedido.DataBind()

            If objNotaFiscal.VencimentosNota Is Nothing OrElse objNotaFiscal.VencimentosNota.Count = 0 Then
                gridVencimentosNota.DataSource = Nothing
                gridVencimentosNota.DataBind()
            End If
        End If

        If objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
            If Not objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada AndAlso objNotaFiscal.TotalNota <> objNotaFiscal.VencimentosNota.Sum(Function(x) x.ValorDoDocumento) Then
                objNotaFiscal.VencimentosNota.ReajFinanceiro = New ReajusteFinanceiro(objNotaFiscal, False)
                objNotaFiscal.VencimentosNota.ReajFinanceiro.ReajustaNotaDeEntrada()
                'gridVencimentosNota.DataSource = objNotaFiscal.VencimentosNota.ToArray
                'Ajuste para Selecinar o Título e carregar no Financeiro - Furlan 30/07/2025
                gridVencimentosNota.DataSource = From tit In objNotaFiscal.VencimentosNota
                                                 Select New With {.Codigo = tit.Codigo, .ReceberPagar = tit.ReceberPagar, .DescricaoMoeda = tit.DescricaoMoeda, .DescricaoProvisao = tit.DescricaoProvisao, .Prorrogacao = tit.Prorrogacao, .ValorDoDocumento = tit.ValorDoDocumento, .MoedaValorDoDocumento = tit.MoedaValorDoDocumento, .CodigoCifrado = IIf(tit.ReceberPagar = "P", Funcoes.Cifrar("ContasAPagar-" & tit.Codigo), Funcoes.Cifrar("ContasAReceber-" & tit.Codigo))}
                gridVencimentosNota.DataBind()

                'gridVencimentosPedido.DataSource = objNotaFiscal.VencimentosPedido.ToArray
                'Ajuste para Selecinar o Título e carregar no Financeiro - Furlan 30/07/2025
                gridVencimentosPedido.DataSource = From tit In objNotaFiscal.VencimentosPedido
                                                   Select New With {.Codigo = tit.Codigo, .ReceberPagar = tit.ReceberPagar, .DescricaoMoeda = tit.DescricaoMoeda, .DescricaoProvisao = tit.DescricaoProvisao, .Prorrogacao = tit.Prorrogacao, .ValorDoDocumento = tit.ValorDoDocumento, .MoedaValorDoDocumento = tit.MoedaValorDoDocumento, .CodigoCifrado = IIf(tit.ReceberPagar = "P", Funcoes.Cifrar("ContasAPagar-" & tit.Codigo), Funcoes.Cifrar("ContasAReceber-" & tit.Codigo))}
                gridVencimentosPedido.DataBind()
            End If
        Else
            objNotaFiscal.VencimentosNota.Clear()
            objNotaFiscal.VencimentosPedido = Nothing
            gridVencimentosNota.DataSource = Nothing
            gridVencimentosNota.DataBind()
        End If

    End Sub

    Protected Sub btnItem_Click(sender As Object, e As EventArgs)
        Dim btnItemOK As Button = CType(sender, Button)
        Dim row As GridViewRow = CType(btnItemOK.NamingContainer, GridViewRow)
        ItemNotaOK(row.RowIndex)
    End Sub

    Protected Sub gridItens_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gridItens.RowCreated
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim btnItem As Button = CType(e.Row.FindControl("btnItem"), Button)
            btnItem.CommandArgument = e.Row.RowIndex.ToString()
        End If
    End Sub

    Protected Sub gridItens_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridItens.SelectedIndexChanged
        SessaoRecuperaNotaFiscal()

        ''**********************************************************************************************
        ''*******************  Cria a estrutura da Tabela que alimenta o Grid  *************************
        ''**********************************************************************************************
        'Dim dtEncargoItem As New DataTable("EncargoItem")
        'dtEncargoItem.Columns.Add("Codigo", Type.GetType("System.String"))
        'dtEncargoItem.Columns.Add("SituacaoTributaria", Type.GetType("System.String"))
        'dtEncargoItem.Columns.Add("SituacaoTributariaPISCOFINS", Type.GetType("System.String"))
        'dtEncargoItem.Columns.Add("SituacaoTributariaIPI", Type.GetType("System.String"))
        'dtEncargoItem.Columns.Add("CodigoOperacao", Type.GetType("System.String"))
        'dtEncargoItem.Columns.Add("CodigoSubOperacao", Type.GetType("System.String"))
        'dtEncargoItem.Columns.Add("CodigoGrupoProduto", Type.GetType("System.String"))
        'dtEncargoItem.Columns.Add("CodigoProduto", Type.GetType("System.String"))
        'dtEncargoItem.Columns.Add("EstadoOrigem", Type.GetType("System.String"))
        'dtEncargoItem.Columns.Add("EstadoDestino", Type.GetType("System.String"))
        'dtEncargoItem.Columns.Add("Base", Type.GetType("System.Decimal"))
        'dtEncargoItem.Columns.Add("Percentual", Type.GetType("System.Decimal"))
        'dtEncargoItem.Columns.Add("PercentualExibicao", Type.GetType("System.String"))
        'dtEncargoItem.Columns.Add("Valor", Type.GetType("System.Decimal"))
        'dtEncargoItem.Columns.Add("Sinal", Type.GetType("System.String"))
        ''**********************************************************************************************

        'For Each row As [Lib].Negocio.NotaFiscalXItemXEncargo In objNotaFiscal.Itens(gridItens.SelectedIndex).Encargos
        '    Dim drRow As DataRow = dtEncargoItem.NewRow()
        '    drRow("Codigo") = row.Codigo
        '    drRow("SituacaoTributaria") = row.SituacaoTributaria
        '    drRow("SituacaoTributariaPISCOFINS") = row.SituacaoTributariaPISCOFINS
        '    drRow("SituacaoTributariaIPI") = row.SituacaoTributariaIPI
        '    drRow("CodigoOperacao") = row.CodigoOperacao
        '    drRow("CodigoSubOperacao") = row.SubOperacao.Codigo
        '    drRow("CodigoGrupoProduto") = row.CodigoGrupoProduto
        '    drRow("CodigoProduto") = row.CodigoProduto
        '    drRow("EstadoOrigem") = row.EstadoOrigem
        '    drRow("EstadoDestino") = row.EstadoDestino
        '    drRow("Base") = row.Base
        '    drRow("Percentual") = row.Percentual
        '    If row.Percentual = row.PercentualExibicao Then
        '        drRow("PercentualExibicao") = ""
        '    Else
        '        drRow("PercentualExibicao") = row.PercentualExibicao
        '    End If
        '    drRow("Valor") = row.Valor
        '    drRow("Sinal") = row.Sinal
        '    dtEncargoItem.Rows.Add(drRow)
        'Next

        'DgEncargos.DataSource = dtEncargoItem
        'DgEncargos.DataBind()

        'Dim q As Integer
        'For q = 0 To DgEncargos.Rows.Count - 1
        '    If DgEncargos.Rows(q).Cells(0).Text() = "FRETES" OrElse _
        '       DgEncargos.Rows(q).Cells(0).Text() = "DESCONTOS" OrElse _
        '       DgEncargos.Rows(q).Cells(0).Text() = "SEGURO" OrElse _
        '       DgEncargos.Rows(q).Cells(0).Text().Contains("FUNRURAL") OrElse _
        '       DgEncargos.Rows(q).Cells(0).Text().Contains("SENAR") OrElse _
        '       DgEncargos.Rows(q).Cells(0).Text().Contains("FACS") OrElse _
        '       DgEncargos.Rows(q).Cells(0).Text().Contains("FETHAB") OrElse _
        '       DgEncargos.Rows(q).Cells(0).Text().Contains("FETHAB GADO") OrElse _
        '       DgEncargos.Rows(q).Cells(0).Text().Contains("FABOV") OrElse _
        '       DgEncargos.Rows(q).Cells(0).Text().Contains("ADUANEIRAS") Then
        '        CType(DgEncargos.Rows(q).FindControl("txtValorEncargoItem"), TextBox).Enabled = True
        '        CType(DgEncargos.Rows(q).FindControl("btnEncargoItem"), Button).Enabled = True
        '    ElseIf DgEncargos.Rows(q).Cells(0).Text().Contains("ICMS") OrElse _
        '           DgEncargos.Rows(q).Cells(0).Text().Contains("IPI") OrElse _
        '           DgEncargos.Rows(q).Cells(0).Text().Contains("PIS") OrElse _
        '           DgEncargos.Rows(q).Cells(0).Text().Contains("COFINS") Then
        '        CType(DgEncargos.Rows(q).FindControl("txtBaseEncargoItem"), TextBox).Enabled = True
        '        CType(DgEncargos.Rows(q).FindControl("txtPercentualItem"), TextBox).Enabled = True
        '        CType(DgEncargos.Rows(q).FindControl("txtValorEncargoItem"), TextBox).Enabled = True
        '        CType(DgEncargos.Rows(q).FindControl("btnEncargoItem"), Button).Enabled = True
        '    End If
        'Next

        If objNotaFiscal.IUD = "U" AndAlso objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Eletronica Then
            imgConfirmarDI.Enabled = False
        Else
            imgConfirmarDI.Enabled = True
        End If

        'Alterado para que Nf's de devolução tenham nas obs de produtos informações relacionadas à lotes e embalagens.
        If Not objNotaFiscal.SubOperacao.Devolucao AndAlso (objNotaFiscal.Itens(gridItens.SelectedIndex).Produto.ControlarEmbalagem = True OrElse objNotaFiscal.Itens(gridItens.SelectedIndex).Produto.ControlarLote = True) Then
            TabContainer1.ActiveTabIndex = 2
            CarregarLoteEmbalagem()
        Else
            gridEmbalagem.DataSource = Nothing
            gridEmbalagem.DataBind()
            TabContainer1.ActiveTabIndex = 1
        End If
    End Sub

    Protected Sub DgEncargos_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        'If e.Row.RowType = DataControlRowType.DataRow Then
        '    Dim btnEncargoItem As Button = CType(e.Row.FindControl("btnEncargoItem"), Button)
        '    btnEncargoItem.CommandArgument = e.Row.RowIndex.ToString()
        'End If
    End Sub


#End Region

#Region "DDL / CHECK"
    Protected Sub cmbFinalidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFinalidade.SelectedIndexChanged
        SessaoRecuperaNotaFiscal()

        If objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
            If (cmbFinalidade.SelectedValue = 20 OrElse cmbFinalidade.SelectedValue = 22) Then
                If Not objNotaFiscal.SubOperacao.Devolucao AndAlso objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR AndAlso objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS AndAlso objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.COMPRAS Then
                    MsgBox(Me.Page, "As Finalidades de Devolução para Compra ou transferencia só são permitida em operações de Devolucao cuja classe seja Deposito ou Afixar")
                    cmbFinalidade.SelectedValue = objNotaFiscal.CodigoFinalidade
                    Exit Sub
                End If
                objNotaFiscal.CodigoAutorizacao = 0
                ChkTroca.Visible = True
                ChkTroca.Enabled = False
                ChkTroca.Checked = False
                objNotaFiscal.Troca = False

                BtnNotas.Visible = False
            Else
                objNotaFiscal.Troca = False
                ChkTroca.Visible = False
                BtnNotas.Visible = objNotaFiscal.Itens(0).Produto.Agrupar = "N"
            End If
        Else
            If (cmbFinalidade.SelectedValue = 20 OrElse cmbFinalidade.SelectedValue = 22) Then
                MsgBox(Me.Page, "As Finalidades de Devolução para Compra ou transferencia só são permitida em operações de Saida")
                cmbFinalidade.SelectedValue = objNotaFiscal.CodigoFinalidade
                Exit Sub
            End If

            If (cmbFinalidade.SelectedValue = 13 OrElse cmbFinalidade.SelectedValue = 23) Then
                If cmbFinalidade.SelectedValue = 13 And Not objNotaFiscal.SubOperacao.PrecoFixo Then
                    MsgBox(Me.Page, "A Devolução para Compra só é permitida em operações de preço Fixo")
                    cmbFinalidade.SelectedValue = objNotaFiscal.CodigoFinalidade
                    Exit Sub
                End If
                objNotaFiscal.CodigoAutorizacao = 0
                ChkTroca.Visible = False
                ChkTroca.Checked = False
                objNotaFiscal.Troca = False

                BtnNotas.Visible = True
            Else
                objNotaFiscal.Troca = False
                ChkTroca.Visible = True
                BtnNotas.Visible = False
            End If
        End If

        objNotaFiscal.CodigoFinalidade = cmbFinalidade.SelectedValue

        SessaoSalvaNotaFiscal()

        If BtnNotas.Visible AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico Then
            If Not BuscaNotaTroca() Then MensagemControle()
        End If
    End Sub

    Protected Sub chk_NossaEmissao_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Session("Carregando" & HID.Value) = False Then AtualizaNossaEmissaoEletronica()
    End Sub

    Protected Sub chk_nfe_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chk_nfe.CheckedChanged
        If Session("Carregando" & HID.Value) = False Then AtualizaNossaEmissaoEletronica()
        txtChaveNFe.Enabled = True
        txtChaveNFe.Focus()
    End Sub
#End Region

#Region "TextBox"
    Protected Sub txtNumero_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNumero.TextChanged
        If Not String.IsNullOrWhiteSpace(txtNumero.Text) Then
            SessaoRecuperaNotaFiscal()
            objNotaFiscal.Codigo = txtNumero.Text.ToUpper
            objNotaFiscal.NotaProdutor = txtNumero.Text.ToUpper
            SessaoSalvaNotaFiscal()
        End If
    End Sub

    Protected Sub txtSerie_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Trim(txtSerie.Text) = "00" Or Trim(txtSerie.Text) = "000" Then
            MsgBox(Me.Page, "Série invalida.")
            txtSerie.Text = ""
        Else
            If Not String.IsNullOrWhiteSpace(txtSerie.Text) Then
                SessaoRecuperaNotaFiscal()
                objNotaFiscal.Serie = Trim(txtSerie.Text)
                objNotaFiscal.SerieNotaProdutor = txtSerie.Text
                txtSerie.Text = objNotaFiscal.Serie
                SessaoSalvaNotaFiscal()
            End If
        End If
    End Sub

    Protected Sub lnkVerificarChaveNFE_Click(ByVal sender As Object, ByVal e As EventArgs)

        If String.IsNullOrWhiteSpace(txtChaveNFe.Text.Replace(".", "")) Then
            MsgBox(Me.Page, "Chave da Nota Fiscal não foi informada.")
        ElseIf txtChaveNFe.Text.Replace(".", "").Length <> 44 Then
            txtChaveNFe.Text = ""
            MsgBox(Me.Page, "Chave da nota eletrônica deve ter 44 dígitos.")
        Else
            SessaoRecuperaNotaFiscal()
            objNotaFiscal.ChaveNFE = txtChaveNFe.Text.Replace(".", "")

            If NotaJaPreenchida() Then
                'Se alguns campos foram preenchidos então serão validados em relação a chavenfe
                If ValidarPreenchimentoDaNFE() Then
                    RealizarManifestoNFE()
                End If
            Else
                'Caso não tenham sido preenchido acontece o manifesto e a importação do xml com posterior preenchimento dos dados da nfe de entrada
                RealizarManifestoNFE()
            End If

        End If
    End Sub

    Private Sub RealizarManifestoNFE()
        SessaoRecuperaNotaFiscal()
        objNotaFiscal.ChaveNFE = txtChaveNFe.Text.Replace(".", "")

        Dim valida As Boolean = True

        If (Mid(objNotaFiscal.ChaveNFE, 7, 14) = "02935843000105" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "01409655000180" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "78393592000146" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "82951310000156" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "87958674000181" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "58290502000184" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "76416890000189" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "03507415000578") Then
            valida = False
        End If

        Dim msgResultado As String = String.Empty

        Dim ModeloNFe As String = Mid(objNotaFiscal.ChaveNFE, 21, 2)
        'Realiza o manifesto da NFe
        If objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica AndAlso ucFile.Parent.Visible AndAlso ModeloNFe.Equals("55") Then '(Modelo: 55 NFe, 57 CTe)

            ''Download do Arquivo.
            Dim bytes As Byte() = New FilesManager().getFileXml(String.Format("{0}-nfe.xml", objNotaFiscal.ChaveNFE))
            If bytes Is Nothing Then
                MsgBox(Me.Page, "XML não foi encontrado, favor inserir o manualmente.")
                Exit Sub
            Else
                Dim caminhoArquivoFile As String = Server.MapPath(String.Format("~/Files/{0}-nfe.xml", objNotaFiscal.ChaveNFE))
                If Not File.Exists(caminhoArquivoFile) Then
                    System.IO.File.WriteAllBytes(caminhoArquivoFile, bytes)
                End If
            End If

            If valida AndAlso Not objNotaFiscal.TemConfirmacaoDaOperacao Then

                Dim verStatusNFe As String = DocumentoEletronico.ConsultaNFEFornecedor(objNotaFiscal)
                Dim statusNFE As String() = verStatusNFe.Split(";")

                If statusNFE(0) = "100" Then
                    If Not DocumentoEletronico.ManifestoNFe(objNotaFiscal, eTipoManifesto.ConfirmacaoDaOperacao, msgResultado) Then
                        MsgBox(Me.Page, msgResultado)
                        Exit Sub
                    End If
                ElseIf statusNFE(0) = "101" Then
                    MsgBox(Me.Page, "Nota Fiscal Cancelada pelo Fornecedor não pode ser utilizada.")
                    Exit Sub
                Else
                    MsgBox(Me.Page, statusNFE(1) & " " & statusNFE(2))
                    Exit Sub
                End If
            End If

            If bytes IsNot Nothing Then
                Try
                    If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                        Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}-nfe.xml", objNotaFiscal.ChaveNFE))
                        If File.Exists(caminhoArquivo) Then
                            File.Delete(caminhoArquivo)
                        End If
                        System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                    End If

                    'Leitura do Arquivo.
                    Dim DsXml As New DataSet
                    DsXml.ReadXml(Server.MapPath(String.Format("~/Files/{0}-nfe.xml", objNotaFiscal.ChaveNFE)))

                    If Not objNotaFiscal.CodigoEmpresa = DsXml.Tables("dest").Rows(0)("CNPJ").ToString() Then
                        MsgBox(Me.Page, "Empresa do XML está diferente da informanda na Nota Fiscal, verifique. Se o erro persistir entre em contato com o suporte.", eTitulo.Erro)
                        Exit Sub
                    End If

                    Dim temArquivo As Boolean = False

                    If objNotaFiscal.Arquivos IsNot Nothing AndAlso objNotaFiscal.Arquivos.Count > 0 Then
                        For Each arq In objNotaFiscal.Arquivos
                            If arq.Descricao = String.Format("{0}-nfe.xml", objNotaFiscal.ChaveNFE) Then
                                temArquivo = True
                            End If
                        Next
                    End If

                    If Not temArquivo Then
                        objNotaFiscal.Arquivos.Add(New [Lib].Negocio.Arquivo() With {
                                             .IUD = "I",
                                             .Codigo = String.Empty,
                                             .Descricao = String.Format("{0}-nfe.xml", objNotaFiscal.ChaveNFE),
                                             .Arquivo = bytes})
                    End If

                    'Incluir a DANFE automaticamente
                    Dim msgNFE As String = ""
                    Dim bytesDanfe As Byte()

                    If DocumentoEletronico.ImprimirNFeDanfe(objNotaFiscal, 1, msgNFE) Then

                        bytesDanfe = New FilesManager().getFile(String.Format("{0}.pdf", objNotaFiscal.ChaveNFE), eTipoDeDocumento.Nota)

                        If bytesDanfe IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then

                            temArquivo = False

                            Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                            If System.IO.File.Exists(caminhoArquivo) Then
                                System.IO.File.Delete(caminhoArquivo)
                            End If
                            System.IO.File.WriteAllBytes(caminhoArquivo, bytesDanfe)

                            If objNotaFiscal.Arquivos IsNot Nothing AndAlso objNotaFiscal.Arquivos.Count > 0 Then
                                For Each arq In objNotaFiscal.Arquivos
                                    If arq.Descricao = String.Format("{0}.pdf", objNotaFiscal.ChaveNFE) Then
                                        temArquivo = True
                                    End If
                                Next
                            End If

                            If Not temArquivo Then
                                objNotaFiscal.Arquivos.Add(New [Lib].Negocio.Arquivo() With {
                                                     .IUD = "I",
                                                     .Codigo = String.Empty,
                                                     .Descricao = String.Format("{0}.pdf", objNotaFiscal.ChaveNFE),
                                                     .Arquivo = bytesDanfe})
                            End If

                        End If
                    End If

                    ucFile.Bind(objNotaFiscal.Arquivos)
                    SessaoSalvaNotaFiscal()

                    'If objNotaFiscal.CodigoPedido <= 0 AndAlso Not NotaJaPreenchida() Then
                    SessaoRecuperaNotaFiscal()

                    PreencherNFeXML(objNotaFiscal, String.Format("{0}-nfe", objNotaFiscal.ChaveNFE), True, msgResultado)

                    'End If
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message)
                    Exit Sub
                End Try
            Else
                MsgBox(Me.Page, "XML não encontrado.")
                Exit Sub
            End If
        Else
            If Not ModeloNFe.Equals("55") Then
                MsgBox(Me.Page, "Manifesto permitido somente para Nota Fiscal")
            ElseIf ModeloNFe.Equals("55") AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica = False Then
                MsgBox(Me.Page, "Empresa não está setada para Emitir NF-e, não é possível importar o XML!")
            End If
        End If
    End Sub

    Private Function NotaJaPreenchida() As Boolean
        If Not String.IsNullOrWhiteSpace(txtNumero.Text) AndAlso CInt(txtNumero.Text) > 0 _
            AndAlso Not String.IsNullOrWhiteSpace(txtSerie.Text) _
            AndAlso Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function ValidarPreenchimentoDaNFE() As Boolean
        Dim valida As Boolean = True

        If (Mid(objNotaFiscal.ChaveNFE, 7, 14) = "02935843000105" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "01409655000180" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "78393592000146" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "82951310000156" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "87958674000181" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "58290502000184" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "76416890000189" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "03507415000578") Then
            valida = False
        End If

        If String.IsNullOrWhiteSpace(txtNumero.Text) Then
            MsgBox(Me.Page, "Número da Nota Fiscal deve ser Informado.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtSerie.Text) Then
            MsgBox(Me.Page, "Série da Nota Fiscal deve ser Informada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            MsgBox(Me.Page, "Cliente da Nota Fiscal deve ser Informado.")
            Return False
        ElseIf Not ddlTipoDeDocumento.SelectedValue = 1 Then
            MsgBox(Me.Page, "Validação só pode ser feita para Documento do Tipo Nota fiscal.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataDeEmissao.Text) Then
            MsgBox(Me.Page, "Data da Nota Fiscal deve ser Informada.")
            Return False
        ElseIf valida AndAlso objNotaFiscal.CodigoCliente.Length = 11 AndAlso Not Mid(objNotaFiscal.ChaveNFE, 10, 11) = objNotaFiscal.CodigoCliente Then
            MsgBox(Me.Page, "Cliente na Chave Eletrônica diferente do informado na Nota Fiscal.")
            Return False
        ElseIf valida AndAlso objNotaFiscal.CodigoCliente.Length = 14 AndAlso Not Mid(objNotaFiscal.ChaveNFE, 7, 14) = objNotaFiscal.CodigoCliente Then
            MsgBox(Me.Page, "Cliente na Chave Eletrônica diferente do informado na Nota Fiscal.")
            Return False
        ElseIf objNotaFiscal.Cliente.Municipio.EstadoIbge = 0 Then
            MsgBox(Me.Page, "o Código IBGE do cliente não foi informado no cadastro! Favor revisar o cadastro do cliente.")
            Return False
        ElseIf valida AndAlso Not CInt(Left(objNotaFiscal.ChaveNFE, 2)) = objNotaFiscal.Cliente.Municipio.EstadoIbge Then
            MsgBox(Me.Page, "Estado do Cliente na Chave Eletrônica diferente do informado na Nota Fiscal.")
            Return False
        ElseIf Not CInt(Mid(objNotaFiscal.ChaveNFE, 26, 9)) = CInt(txtNumero.Text) Then
            MsgBox(Me.Page, "Número da Nota na Chave Eletrônica diferente do informado na Nota Fiscal.")
            Return False
        ElseIf Not CInt(Mid(objNotaFiscal.ChaveNFE, 23, 3)) = CInt(txtSerie.Text) Then
            MsgBox(Me.Page, "Série da Nota na Chave Eletrônica diferente do informado na Nota Fiscal.")
            Return False
        ElseIf Not Mid(objNotaFiscal.ChaveNFE, 3, 2) = String.Format("{0:yy}", CDate(txtDataDeEmissao.Text)) Then
            MsgBox(Me.Page, "Ano da Nota na Chave Eletrônica diferente do informado na Nota Fiscal.")
            Return False
        End If

        Return True
    End Function

    'Carrega a NFe apartir do user control de arquivo.
    Public Overrides Sub Carregar(ByVal pNomeDoArquivo As String)
        Try
            SessaoRecuperaNotaFiscal()
            PreencherNFeXML(objNotaFiscal, pNomeDoArquivo, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub PreencherNFeXML(ByVal objNotaFiscal As [Lib].Negocio.NotaFiscal, ByVal pNomeArquivo As String, ByVal pOrigem As Boolean, Optional ByVal msgResultado As String = "")
        Dim DsXml As New DataSet

        If pNomeArquivo.Contains("-nfe") Then pNomeArquivo = pNomeArquivo.Replace("-nfe", "")

        DsXml.ReadXml(Server.MapPath(String.Format("~/Files/{0}-nfe.xml", pNomeArquivo)))
        SessaoDsXML = DsXml
        DocumentoEletronico.PreencherNFeComXML(objNotaFiscal, DsXml, False, False, False, "", chkImportarProdutoUnico.Checked, chkAGruparNCM.Checked)
        SessaoSalvaNotaFiscal()

        'objNotaFiscal = New [Lib].Negocio.NotaFiscal

        If objNotaFiscal.Cliente Is Nothing OrElse objNotaFiscal.Cliente.Codigo.Length = 0 Then
            MsgBox(Me.Page, "Fornecedor informado no XML não foi encontrado no Sistema.", eTitulo.Info)
            Exit Sub
        End If

        txtNomeDoCliente.Text = objNotaFiscal.Cliente.Nome
        txtInscricaoDoCliente.Text = objNotaFiscal.Cliente.InscricaoEstadual
        txtCnpjDoCliente.Text = objNotaFiscal.Cliente.CodigoFormatado
        txtEnderecoDoCliente.Text = objNotaFiscal.Cliente.Endereco
        txtComplementoDoCliente.Text = objNotaFiscal.Cliente.Complemento
        txtBairroDoCliente.Text = objNotaFiscal.Cliente.Bairro
        txtCepDoCliente.Text = objNotaFiscal.Cliente.CEP
        txtCidadeDoCliente.Text = objNotaFiscal.Cliente.Cidade
        txtTelefoneDoCliente.Text = objNotaFiscal.Cliente.Telefone
        txtEstadoDoCliente.Text = objNotaFiscal.Cliente.CodigoEstado
        txtInscricaoDoCliente.Text = objNotaFiscal.Cliente.InscricaoEstadual

        txtDataDeEmissao.Text = objNotaFiscal.DataNota.ToString("dd/MM/yyyy")
        txtDataDeEntrada.Text = objNotaFiscal.Movimento.ToString("dd/MM/yyyy")

        txtES.Text = objNotaFiscal.EntradaSaida.ToString.Substring(0, 1)
        txtNumero.Text = objNotaFiscal.Codigo
        txtSerie.Text = objNotaFiscal.Serie
        txtChaveNFe.Text = objNotaFiscal.ChaveNFE

        If objNotaFiscal.TemSegCodBarra Then
            txtTextoSegCodBarra.Visible = True
            txtSegCodBarra.Visible = True
            txtSegCodBarra.ReadOnly = False
        End If

        If objNotaFiscal.CodigoTransportador.Length > 0 Then
            txtCodigoTransportador.Value = objNotaFiscal.Transportador.Codigo
            txtNomeDoTransportador.Text = objNotaFiscal.Transportador.Nome
            txtCidadeDoTransportador.Text = objNotaFiscal.Transportador.Cidade
            txtEstadoDoTransportador.Text = objNotaFiscal.Transportador.CodigoEstado
            txtCnpjDoTransportador.Text = objNotaFiscal.Transportador.Codigo
            txtInscricaoDoTransportador.Text = objNotaFiscal.Transportador.InscricaoEstadual
            txtEnderecoDoTransportador.Text = objNotaFiscal.Transportador.Endereco & ", " & objNotaFiscal.Transportador.Numero
        End If

        Dim temFrete As Boolean = False

        CarregarDDLFrete(True)

        For Each item In ddlFrete.Items
            If item.value = objNotaFiscal.CIFFOB.ToString Then
                temFrete = True
                Exit For
            End If
        Next

        If temFrete Then
            ddlFrete.SelectedValue = objNotaFiscal.CIFFOB.ToString
        Else
            MsgBox(Me.Page, "Modelo do Frete informado no XML difere do informado no Pedido.", eTitulo.Info)
        End If

        If msgResultado.Length > 0 Then MsgBox(Me.Page, msgResultado, eTitulo.Info)

        'VerificarProdutos se a nota existe

        Dim strSQL As String

        strSQL = "  SELECT 1 " & vbCrLf &
                             "  FROM NotasFiscais NF" & vbCrLf &
                             "  INNER JOIN NotasFiscaisxItens NFxI" & vbCrLf &
                             "      ON NF.Empresa_Id        = NFxI.Empresa_Id" & vbCrLf &
                             "      AND NF.EndEmpresa_Id    = NFxI.EndEmpresa_Id" & vbCrLf &
                             "      AND NF.Cliente_Id       = NFxI.Cliente_Id" & vbCrLf &
                             "      AND NF.EndCliente_Id    = NFxI.EndCliente_Id" & vbCrLf &
                             "      AND NF.EntradaSaida_Id  = NFxI.EntradaSaida_Id" & vbCrLf &
                             "      AND NF.Serie_Id         = NFxI.Serie_Id" & vbCrLf &
                             "      AND NF.Nota_Id          = NFxI.Nota_Id" & vbCrLf &
                             "  WHERE NF.Empresa_Id         = '" & objNotaFiscal.Empresa.Codigo & "'" & vbCrLf &
                             "      AND NF.Cliente_Id       = '" & objNotaFiscal.Cliente.Codigo & "'" & vbCrLf &
                             "      AND NF.Nota_Id          = " & objNotaFiscal.Codigo & vbCrLf &
                             "      AND NF.Serie_Id         = '" & objNotaFiscal.Serie & "'" & vbCrLf &
                             "      AND NF.EntradaSaida_Id  = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "NotasFiscaisGerais")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count() > 0 Then
            MsgBox(Me.Page, "XML já possui uma nota fiscal cadastrada, não é possivel importar!", eTitulo.Info, False)
            LimparCampos("", "", False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
            Exit Sub
        End If

        'If objNotaFiscal.NotasTrocaOrigem.Count > 0 Then
        '    grdNotasReferenciais.DataSource = objNotaFiscal.NotasTrocaOrigem
        '    grdNotasReferenciais.DataBind()
        'End If

        If objNotaFiscal.CodigoPedido <= 0 Then ConsultaPedidos(True)
    End Sub

    Protected Sub txtObservacoesDeEmbarque_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        objNotaFiscal.ObservacoesDeEmbarque = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacoesDeEmbarque.Text)
        txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque

        SessaoSalvaNotaFiscal()
    End Sub

    Protected Sub txtObservacoesFiscais_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        objNotaFiscal.Observacoes = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacoesFiscais.Text)
        txtObservacoesFiscais.Text = objNotaFiscal.Observacoes
        SessaoSalvaNotaFiscal()
    End Sub

    Protected Sub txtDataDeEmissao_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsDate(txtDataDeEmissao.Text) Then
            MsgBox(Me.Page, "Data de Emissão inválida!")
            txtDataDeEmissao.Text = Today.ToString("dd/MM/yyyy")
            Exit Sub
        End If

        If CDate(txtDataDeEmissao.Text) > Now.Date Then
            txtDataDeEmissao.Text = Today.ToString("dd/MM/yyyy")
            MsgBox(Me.Page, "Data de Emissão não pode ser maior que a data atual")
        Else

            SessaoRecuperaNotaFiscal()

            If Not objNotaFiscal Is Nothing Then
                objNotaFiscal.DataNota = CDate(txtDataDeEmissao.Text)
            End If

            If gridItens.Rows.Count > 0 Then
                Dim EnableQuantidadeItem As Boolean = False
                Dim EnableUnitarioItem As Boolean = False
                Dim EnableTotalItem As Boolean = False

                EnableQuantidadeItem = CType(gridItens.Rows(0).FindControl("txtQuantidadeItem"), TextBox).Enabled()
                EnableUnitarioItem = CType(gridItens.Rows(0).FindControl("txtUnitarioItem"), TextBox).Enabled()
                EnableTotalItem = CType(gridItens.Rows(0).FindControl("txtTotalItem"), TextBox).Enabled

                AtualizaFormularioComAClasse()

                If gridItens.Rows.Count > 0 Then
                    Dim i As Integer = 0
                    While i < gridItens.Rows.Count
                        CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled() = EnableQuantidadeItem
                        CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled() = EnableUnitarioItem
                        CType(gridItens.Rows(i).FindControl("txtTotalItem"), TextBox).Enabled = EnableTotalItem

                        i += 1
                    End While
                End If
            End If

            SessaoSalvaNotaFiscal()
        End If
    End Sub

    Protected Sub txtDataDeEntrada_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsDate(txtDataDeEntrada.Text) Then
            MsgBox(Me.Page, "Data de Entrada inválida!")
            txtDataDeEntrada.Text = Today.ToString("dd/MM/yyyy")
            Exit Sub
        End If

        If CDate(txtDataDeEntrada.Text) > Now.Date Then
            txtDataDeEntrada.Text = Today.ToString("dd/MM/yyyy")
            MsgBox(Me.Page, "Data da Nota não pode ser maior que a data atual")
        Else
            SessaoRecuperaNotaFiscal()

            objNotaFiscal.Movimento = CDate(txtDataDeEntrada.Text)

            If gridItens.Rows.Count > 0 Then
                Dim EnableQuantidadeItem As Boolean = False
                Dim EnableUnitarioItem As Boolean = False
                Dim EnableTotalItem As Boolean = False

                EnableQuantidadeItem = CType(gridItens.Rows(0).FindControl("txtQuantidadeItem"), TextBox).Enabled()
                EnableUnitarioItem = CType(gridItens.Rows(0).FindControl("txtUnitarioItem"), TextBox).Enabled()
                EnableTotalItem = CType(gridItens.Rows(0).FindControl("txtTotalItem"), TextBox).Enabled

                AtualizaFormularioComAClasse()

                Dim i As Integer = 0
                While i < gridItens.Rows.Count
                    CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled() = EnableQuantidadeItem
                    CType(gridItens.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled() = EnableUnitarioItem
                    CType(gridItens.Rows(i).FindControl("txtTotalItem"), TextBox).Enabled = EnableTotalItem

                    i += 1
                End While
            End If

            SessaoSalvaNotaFiscal()
        End If
    End Sub
#End Region

#Region "Botões"
    Protected Sub btnObservacoesFiscais_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HttpContext.Current.Session("ssCampo" & HID.Value) = "NotaFiscalXItens"
            HttpContext.Current.Session("Observacoes" & HID.Value) = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacoesFiscais.Text)
            ucConsultaObservacoes.BindGridView()
            Popup.ConsultaDeObservacoes(Me, "objObservacoesNXI" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEmpresa.Click
        Try
            SessaoRecuperaNotaFiscal()
            LimparCampos("", "", False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)

            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me, "objEmpresaNXI" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ListarClientes()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPedido.Click
        Try
            ConsultaPedidos(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnDeposito_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.IUD = "U" AndAlso objNotaFiscal.CodigoRomaneio > 0 AndAlso objNotaFiscal.Romaneio.Pesagens.Count > 0 Then
                MsgBox(Me.Page, "O Deposito de Notas com Laudo Nao pode ser Alterado")
                Exit Sub
            Else
                ucConsultaClientes.Limpar()
                Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
                Popup.ConsultaDeClientes(Me, "objDepositoNXI" & HID.Value, txtNome.ClientID)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnEmbarque_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeClientes(Me, "objEmbarqueNXI" & HID.Value, txtNome.ClientID)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnTransbordo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeClientes(Me.Page, "objTransbordoNXI" & HID.Value, txtNome.ClientID)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnOrigemDestino_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeClientes(Me.Page, "objOrigemDestinoNXI" & HID.Value, txtNome.ClientID)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnEntrega_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeClientes(Me.Page, "objEntregaNXI" & HID.Value, txtNome.ClientID)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnRomaneio_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not BuscarRomaneio() Then
                MensagemControle()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnNotas_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnNotas.Click
        Try
            If Not BuscaNotaTroca() Then MensagemControle()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ChkTroca_CheckedChanged(sender As Object, e As EventArgs) Handles ChkTroca.CheckedChanged
        Try
            SessaoRecuperaNotaFiscal()
            If Not ChkTroca.Checked AndAlso objNotaFiscal.IUD = "U" AndAlso objNotaFiscal.TemNotaTroca Then
                ChkTroca.Checked = True
                MsgBox(Me.Page, "Esta nota já foi trocada apague o vinculo das trocas.")
            Else
                objNotaFiscal.Troca = ChkTroca.Checked
            End If
            SessaoSalvaNotaFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPlaca_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPlaca.Click
        Try
            BuscarPlaca()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarDDLFrete(ByVal bCarregarXML As Boolean)
        Try
            'Carregas opções dos tipos de frete em relação ao tipo definido na NF Caso seja CIF então somente estarão disponível CIF e
            ' NEN, caso seja FOB então FOB e NEN. 
            ddlFrete.Items.Clear()
            Dim TipoDeFrete As New Dictionary(Of String, String)

            If objNotaFiscal IsNot Nothing AndAlso (objNotaFiscal.Pedido IsNot Nothing AndAlso objNotaFiscal.CodigoPedido > 0) OrElse bCarregarXML Then
                If objNotaFiscal.CIFFOB = eTiposFrete.CIF Then
                    TipoDeFrete.Add("CIF", "CIF - Emitente")    'Saída CIF - 0 - Por conta do emitente = EMPRESA
                    TipoDeFrete.Add("TER", "TER - Terceiro")
                    TipoDeFrete.Add("NEN", "NEN - Nenhum")
                ElseIf objNotaFiscal.CIFFOB = eTiposFrete.FOB Then
                    TipoDeFrete.Add("FOB", "FOB - Destinatario") 'Entrada FOB - 1 – Por conta do destinatário = EMPRESA
                    TipoDeFrete.Add("TER", "TER - Terceiro")
                    TipoDeFrete.Add("NEN", "NEN - Nenhum")
                ElseIf objNotaFiscal.CIFFOB = eTiposFrete.TER Then
                    TipoDeFrete.Add("TER", "TER - Terceiro") '2 – Por conta de terceiros
                    TipoDeFrete.Add("NEN", "NEN - Nenhum")
                ElseIf objNotaFiscal.CIFFOB = eTiposFrete.NEN AndAlso Not objNotaFiscal.Pedido Is Nothing AndAlso objNotaFiscal.Pedido.FreteCIFFOB = eTiposFrete.CIF Then
                    If objNotaFiscal.SubOperacao.Devolucao Then
                        TipoDeFrete.Add("FOB", "FOB - Destinatario")
                    Else
                        TipoDeFrete.Add("CIF", "CIF - Emitente")
                    End If
                    TipoDeFrete.Add("TER", "TER - Terceiro")
                    TipoDeFrete.Add("NEN", "NEN - Nenhum")
                ElseIf objNotaFiscal.CIFFOB = eTiposFrete.NEN AndAlso Not objNotaFiscal.Pedido Is Nothing AndAlso objNotaFiscal.Pedido.FreteCIFFOB = eTiposFrete.FOB Then
                    If objNotaFiscal.SubOperacao.Devolucao Then
                        TipoDeFrete.Add("CIF", "CIF - Emitente")
                    Else
                        TipoDeFrete.Add("FOB", "FOB - Destinatario")
                    End If
                    TipoDeFrete.Add("TER", "TER - Terceiro")
                    TipoDeFrete.Add("NEN", "NEN - Nenhum")

                Else
                    TipoDeFrete.Add("NEN", "NEN - Nenhum")
                End If

                ddlFrete.DataTextField = "Value"
                ddlFrete.DataValueField = "Key"
                ddlFrete.DataSource = TipoDeFrete
                ddlFrete.DataBind()

                ddlFrete.Enabled = True

                ddlFrete.SelectedValue = objNotaFiscal.CIFFOB.ToString

                If bCarregarXML = False Then
                    'Caso a opção nenhum NEN esteja setado tanto na NF quando no pedido então não será possível escolher outra
                    'Só está sendo feita a validação do tipo na NF para casos que ainda não tenham sido corrigidos já que é necessário fazer carta de correção eletrônica.
                    'A regra é que se no pedido estiver NEN então em suas notas também será NEN sem opção de mudança.
                    If objNotaFiscal.Pedido.FreteCIFFOB = eTiposFrete.NEN AndAlso objNotaFiscal.CIFFOB = eTiposFrete.NEN Then
                        ddlFrete.SelectedValue = eTiposFrete.NEN.ToString
                        ddlFrete.Enabled = False
                    End If
                End If

            Else
                TipoDeFrete.Add("NEN", "NEN - Nenhum")

                ddlFrete.DataTextField = "Value"
                ddlFrete.DataValueField = "Key"
                ddlFrete.DataSource = TipoDeFrete
                ddlFrete.DataBind()

                ddlFrete.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlFrete_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlFrete.SelectedIndexChanged
        Try
            SessaoRecuperaNotaFiscal()
            'If objNotaFiscal.IUD = "U" Then
            '    If Funcoes.VerificaPermissao("NotaFiscalXItens", "ALTERAR") Then
            '        Dim msgNFE As String = String.Empty
            '        If objNotaFiscal.NossaEmissao Then

            '            objNotaFiscal.CIFFOB = ([Enum].Parse(GetType(eTiposFrete), ddlFrete.SelectedValue))

            '            If [Lib].Negocio.DocumentoEletronico.CCeCIFFOB(objNotaFiscal, msgNFE) Then
            '                Dim Sqls As New ArrayList

            '                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-impcarta{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
            '                Sqls.Add(Sql)
            '                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-emailcarta{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
            '                Sqls.Add(Sql)

            '                Banco.GravaBanco(Sqls)

            '                Thread.Sleep(5000)

            '                Dim strChave As String = objNotaFiscal.ChaveNFE & "cce"
            '                Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", strChave), eTipoDeDocumento.Nota)
            '                If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(strChave) Then
            '                    Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", strChave))
            '                    System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
            '                    Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", strChave))
            '                End If
            '            Else
            '                MsgBox(Me.Page, "Erro na Carta de Correção: " & msgNFE & ". Tente novamente ou entre em contato com o Suporte.")
            '                LimparCampos("", "", False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
            '                Exit Sub
            '            End If
            '        End If
            '        If objNotaFiscal.AtualizarCIFFOB Then
            '            SessaoSalvaNotaFiscal()
            '            lnkAtualizar.Parent.Visible = False
            '            lnkExcluir.Parent.Visible = False

            '            If Not String.IsNullOrWhiteSpace(msgNFE) Then
            '                MsgBox(Me.Page, "CIFFOB alterado com sucesso! " & msgNFE, eTitulo.Sucess)
            '            Else
            '                MsgBox(Me.Page, "CIFFOB alterado com Sucesso.", eTitulo.Sucess)
            '            End If
            '        Else
            '            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            '        End If
            '    Else
            '        MsgBox(Me.Page, "Usuário sem permissão para alterar registro!")
            '    End If
            'Else
            objNotaFiscal.CIFFOB = ([Enum].Parse(GetType(eTiposFrete), ddlFrete.SelectedValue))

            If objNotaFiscal.CIFFOB = eTiposFrete.NEN Then
                Dim objPlaca As Placa = New Placa()
                objNotaFiscal.PlacaDetalhes = objPlaca
                objNotaFiscal.PlacaTransportador = ""
                objNotaFiscal.CodigoMotorista = ""
                objNotaFiscal.CodigoTransportador = ""
                Dim objTransportador As Cliente = New Cliente()
                objNotaFiscal.Transportador = objTransportador
                objNotaFiscal.Motorista = objTransportador

                txtNomeDoTransportador.Text = ""
                txtCidadeDoTransportador.Text = ""
                txtEstadoDoTransportador.Text = ""
                txtCnpjDoTransportador.Text = ""
                txtInscricaoDoTransportador.Text = ""
                txtEnderecoDoTransportador.Text = ""
                txtPlacas.Text = ""
                txtCodigoPlaca.Value = ""
                txtEstadoDaPlaca.Text = ""
            End If

            SessaoSalvaNotaFiscal()
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnInutilizar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNotaFiscal()
            ucInutilizacao.CarregarInutilizacao(objNotaFiscal)
            Popup.ConsultaDeInutilizacao(Me.Page, "objInutilizacaoNFe" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnTransportador_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTransportador.Click
        Try

            If ddlFrete.SelectedValue = eTiposFrete.NEN.ToString() Then
                MsgBox(Me.Page, "Transportador não pode ser informado.", eTitulo.Info)
                Exit Sub
            End If

            ucConsultaClientes.Limpar()

            ucConsultaClientes.SetarTipoCliente("7")

            Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeClientes(Me, "objTransportadorNXI" & HID.Value, txtNome.ClientID)
            Popup.CenterDialog(Me, "divConsultaCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnClassificacao_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNotaFiscal()

            If objNotaFiscal.CodigoRomaneio > 0 AndAlso Not objNotaFiscal.Romaneio.Processo = "NOTA FISCAL" Then
                MsgBox(Me.Page, "A Nota Já Possui um Romaneio / Classificação")
            ElseIf txtPedido.Text.Length = 0 Then
                MsgBox(Me.Page, "Para consultar a classificação é necessário informar o Pedido")
            ElseIf objNotaFiscal.CodigoRomaneio > 0 AndAlso Not objNotaFiscal.CriarRomaneio AndAlso Not objNotaFiscal.Romaneio.Processo = "NOTA FISCAL" Then
                MsgBox(Me.Page, "Esta Nota já esta vinculado a um Romaneio/Classificação, fisico = sim")
            Else
                If objNotaFiscal.Itens(0).Produto.ControlarPesagem Then
                    If objNotaFiscal.SubOperacao.QuantidadeFisico Then
                        If Not BuscaClassificacao(False) Then MensagemControle()
                    Else
                        MensagemControle()
                        BuscaProcuracao()
                    End If
                Else
                    MsgBox(Me.Page, "Não é informado a Classificação/Peso Físico para este Produto")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProcuracao_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not BuscaProcuracao() Then MensagemControle()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnUFDesembarqueDI_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.DadosDaImportacao Is Nothing Then
                objNotaFiscal.DadosDaImportacao = New [Lib].Negocio.NotaFiscalXImportacao(objNotaFiscal)
                SessaoSalvaNotaFiscal()
            End If
            Popup.ConsultaDeEstados(Me.Page, "objUFDesembarqueDI" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnUFEmbarqueDI_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.DadosDaImportacao Is Nothing Then
                objNotaFiscal.DadosDaImportacao = New [Lib].Negocio.NotaFiscalXImportacao(objNotaFiscal)
                SessaoSalvaNotaFiscal()
            End If
            Popup.ConsultaDeEstados(Me.Page, "objUFEmbarqueDI" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnFabricanteDI_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo" & HID.Value) = "LivreClasse"
            Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeClientes(Me.Page, "objFabricanteDI" & HID.Value, txtNome.ClientID)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function validar_DI() As Boolean
        If Not Trim(txtDI.Text).Length = 11 Then
            erroMsg = "Declaração de Importação deve ter 11 dígitos."
            Return False
        ElseIf txtDataDI.Text.Length = 0 Then
            erroMsg = "Data da Declaração de Importação deve ser informada."
            Return False
        ElseIf Trim(txtDesembarqueDI.Text).Length = 0 Then
            erroMsg = "Local do Desembarque da Declaração de Importação deve ser informado."
            Return False
        ElseIf Trim(lblUFDesembarqueDI.Text).Length = 0 Then
            erroMsg = "UF do Desembarque da Declaração de Importação deve ser informada."
            Return False
        ElseIf txtDataDesembarqueDI.Text.Length = 0 Then
            erroMsg = "Data do Desembarque da Declaração de Importação deve ser informada."
            Return False
        ElseIf Trim(lblFabricanteDI.Text).Length = 0 Then
            erroMsg = "Fabricante que está na Declaração de Importação deve ser informado."
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlViaDeTransporteDI.SelectedValue) Then
            erroMsg = "Via de Transporte na Declaração de Importação é Obrigatório."
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlTipoDeImportacaoDI.SelectedValue) Then
            erroMsg = "Tipo de Importação na Declaração de Importação é Obrigatório."
            Return False
        ElseIf ddlViaDeTransporteDI.SelectedValue = 1 AndAlso (String.IsNullOrWhiteSpace(txtValorVAFRMMDI.Text) OrElse Not CDec(txtValorVAFRMMDI.Text) > 0) Then
            erroMsg = "Para Via de Transporte igual a Marítima. Campo Valor VAFRMM é de preenchimento obrigatório"
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub BtnCancelarSefaz_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCancelarSefaz.Click
        Try
            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.CodigoSituacao = 2 Then
                MsgBox(Me.Page, "Esta nota já está cancelada!")
                Exit Sub
            ElseIf objNotaFiscal.CodigoSituacao = 7 Then
                MsgBox(Me.Page, "Esta nota está no aguardo da SEFAZ para o seu cancelamento, não se pode alterar a observação!")
                Exit Sub
            ElseIf Not objNotaFiscal.NossaEmissao Then
                MsgBox(Me.Page, "Esta nota não é de nossa emissão, sendo assim não pode ser cancelada!")
                Exit Sub
            ElseIf Not objNotaFiscal.Eletronica Then
                MsgBox(Me.Page, "Esta nota não é uma nota eletrônica!")
                Exit Sub
            ElseIf String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                MsgBox(Me.Page, "Esta nota não possui a chave de NF-e da SEFAZ!")
                Exit Sub
            ElseIf String.IsNullOrWhiteSpace(objNotaFiscal.ProtocoloNota) Then
                MsgBox(Me.Page, "Esta nota não possui o número de protocolo da SEFAZ!")
                Exit Sub
            ElseIf txtObservacaoCancelamento.Text.Trim().Length < 15 Then
                MsgBox(Me.Page, "A observação deve conter pelo menos 15 caracteres!")
                Exit Sub
            ElseIf Not btnModo.Text.ToUpper().Trim() = "MODO NORMAL" Then
                MsgBox(Me.Page, "Cancelamento não pode ser feito em Modo de Contingência. ")
                Exit Sub
            End If

            'VERIFICA SE HÁ PRAZO PARA CANCELAMENTO DA NFE
            'NO BANCO ESTÁ GRAVADO EM HORAS EX(1 DIA(24), 7 DIAS(168)), (-2 MINUTOS) MARGEM DE SEGURANÇA PARA CANCELAMENTO
            'Cadastrar o CancelarForaDoPrazo em Processos e dar a Permissão de ACESSAR apenas para o usuário que pode fazer isso. Furlan 02/09/2014
            If Not Funcoes.VerificaPermissao("CancelarForaDoPrazo", "ACESSAR") Then
                Dim objClienteXEmpresa As New [Lib].Negocio.ClienteXEmpresa(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                If (objNotaFiscal.DataHoraNFE.HasValue AndAlso DateTime.Now > CDate(objNotaFiscal.DataHoraNFE).AddHours(objClienteXEmpresa.PrazoCancelamentoNFE).AddMinutes(-2)) Then
                    MsgBox(Me.Page, "Esta nota fiscal não pode ser cancelada, pois o prazo para cancelamento terminou no dia: " & CDate(objNotaFiscal.DataHoraNFE).AddHours(objClienteXEmpresa.PrazoCancelamentoNFE).AddMinutes(-2).ToString("dd/MM/yyyy HH:mm:ss"))
                    Exit Sub
                End If
            End If

            If Not objNotaFiscal.VencimentosNota Is Nothing AndAlso objNotaFiscal.VencimentosNota.Any(Function(s) s.CodigoSituacao = eSituacao.RemessaBancaria) Then
                MsgBox(Me.Page, "Não é possível excluir a nota fiscal " & objNotaFiscal.Codigo & ", o financeiro está em Cobrança Bancária.")
                Exit Sub
            ElseIf Not objNotaFiscal.VencimentosNota Is Nothing AndAlso objNotaFiscal.VencimentosNota.Any(Function(s) s.CodigoSituacao = eSituacao.EndossoTitulo) Then
                MsgBox(Me.Page, "Não é possível excluir a nota fiscal " & objNotaFiscal.Codigo & ", o financeiro está com Endosso.")
                Exit Sub
            End If

            objNotaFiscal.ObservacaoCancelamento = txtObservacaoCancelamento.Text
            objNotaFiscal.IUD = "C"

            If objNotaFiscal.Salvar() Then
                If VerificarNFe() Then
                    Dim msgNFE As String = String.Empty

                    If objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica AndAlso [Lib].Negocio.DocumentoEletronico.CancelarNFe(objNotaFiscal, msgNFE) Then

                        Dim Sqls As New ArrayList
                        CancelarNotaFiscal(Sqls, objNotaFiscal)

                        'Dim textoFinal As New Text.StringBuilder()

                        'For Each item As String In Sqls
                        '    textoFinal.AppendLine(item)
                        'Next

                        '' Exemplo: exibir em um TextBox
                        'Dim t = textoFinal.ToString()

                        If Not String.IsNullOrWhiteSpace(objNotaFiscal.Empresa.EmailNFE) AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.Cliente.EmailNFE) Then
                            [Lib].Negocio.DocumentoEletronico.SendMailNFe(objNotaFiscal, "", False)
                        End If

                        If [Lib].Negocio.DocumentoEletronico.ImprimirNFe(objNotaFiscal, txtNumViasDeImpressao.Text, msgNFE) Then
                            Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", objNotaFiscal.ChaveNFE), eTipoDeDocumento.Nota)
                            If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                                Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                                System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                                Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                            End If
                        End If

                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-status{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)
                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-nota{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)
                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-email{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)
                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-danfe{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        Else
                            If Not String.IsNullOrEmpty(msgNFE) Then
                                MsgBox(Me.Page, msgNFE)
                            End If
                            LimparCampos("C", objNotaFiscal.Codigo, objNotaFiscal.NossaEmissao, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                        End If
                    Else
                        MsgBox(Me.Page, msgNFE)
                        '690 - Pedido de Cancelamento para NF-e com CT-e - Não pode proceder com o Cancelamento.
                    End If
                End If
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnRetirada_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

            SessaoRecuperaNotaFiscal()

            If Left(objNotaFiscal.CodigoEmpresa, 8) = "24450490" Then
                Exit Sub
            End If

            If Not BuscaAutorizacao() Then
                MensagemControle()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnEmbalagem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnEmbalagem.Click
        Try
            If gridEmbalagem.Rows.Count = 0 Then
                MsgBox(Me.Page, "Este Produto Não Contem Lotes/Embalagens cadastrador ou abertos")
                Exit Sub
            End If

            SessaoRecuperaLoteEmbalagem()
            SessaoRecuperaNotaFiscal()

            If objListaDeCombinacoesLoteEmbalagem.QtdeDeEmbalagens = 0 Then
                MsgBox(Me.Page, "Informe a Qtde do produto")
                Exit Sub
            End If

            For Each row As [Lib].Negocio.SaldoProdutoEstoqueLoteEmbalagem In objListaDeCombinacoesLoteEmbalagem
                If objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso row.QtdeEmbalagem > 0 AndAlso row.QtdeEmbalagem > row.SaldoEmbalagem Then
                    MsgBox(Me.Page, "A Qtde de Embalagem Informada é maior do que o numero em estoque")
                    Exit Sub
                End If
            Next

            Dim ProdutoNota As [Lib].Negocio.NotaFiscalXItem = objNotaFiscal.Itens(gridItens.SelectedIndex)

            If objListaDeCombinacoesLoteEmbalagem.QtdeDeProduto > ProdutoNota.SaldoPedidoFiscal And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS Then
                MsgBox(Me.Page, "A Soma das Qtdes nao pode ser maior que o saldo do produto")
                Exit Sub
            End If

            For k As Integer = objNotaFiscal.Itens.Count - 1 To 0 Step -1
                If objNotaFiscal.Itens(k).CodigoProduto = ProdutoNota.CodigoProduto Then objNotaFiscal.Itens.RemoveAt(k)
            Next k

            Dim i As Integer = 0
            For Each row As [Lib].Negocio.SaldoProdutoEstoqueLoteEmbalagem In objListaDeCombinacoesLoteEmbalagem
                If row.QtdeEmbalagem > 0 AndAlso row.QtdeEmbXCapacidade > 0 Then
                    '**********************************************************************************************
                    '************************* Carrega Item da Nota Fiscal ****************************************
                    '**********************************************************************************************
                    Dim objItemNF As New [Lib].Negocio.NotaFiscalXItem(objNotaFiscal)

                    With objItemNF
                        .CodigoProduto = ProdutoNota.CodigoProduto

                        .Lote = row.Lote
                        .Classificacao = row.Classificacao
                        .CodigoEmbalagem = row.Embalagem
                        .CodigoEmbalagemIndea = row.EmbalagemIndea
                        .CodigoTipoDeEmbalagem = row.TipoDeEmbalagem
                        .CapacidadeEmbalagem = row.CapacidadeEmbalagem
                        .QuantidadeDeEmbalagem = row.QtdeEmbalagem


                        If .Produto.ControlarEmbalagem Then
                            .ObservacoesDoProduto &= row.QtdeEmbalagem & " " & .Embalagem.Descricao & " " & .TipoDeEmbalagem.Descricao & " " & IIf(row.PesoVariavel, "", row.CapacidadeEmbalagem & " " & .Produto.Unidade) & " / " & .Produto.DescricaoTecnica
                        End If

                        If .Produto.ControlarLote Then
                            Dim L As New [Lib].Negocio.Lote(row.Lote, .CodigoProduto)
                            If L.Tipo = 2 Then
                                .ObservacoesDoProduto &= IIf(.ObservacoesDoProduto.Length = 0, "", ", ") & "Lote: " & .Lote & " Validade " & L.DataValidade.ToString("dd-MM-yyyy")
                            Else
                                .ObservacoesDoProduto &= IIf(.ObservacoesDoProduto.Length = 0, "", ", ") & "Lote: " & .Lote & " GER " & L.Germinacao & IIf(L.Pureza = 0, "", " PUR " & L.Pureza) & " Peneira/Classif. " & .Classificacao & " Renasem " & L.Renasem & " Validade " & L.DataValidade.ToString("dd-MM-yyyy")
                            End If

                        End If

                        .Sequencia = i
                        i += 1

                        .CodigoOperacao = objNotaFiscal.CodigoOperacao
                        .CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao

                        .CodigoPedido = ProdutoNota.CodigoPedido
                        .PesoQuantidade = ProdutoNota.PesoQuantidade

                        .PesoFiscal = row.QtdeEmbXCapacidade
                        If objNotaFiscal.SubOperacao.QuantidadeFisico Then
                            .QuantidadeFisica = row.QtdeEmbXCapacidade
                        End If
                        .QuantidadeFiscal = row.QtdeEmbXCapacidade

                        .SaldoValorOficial = ProdutoNota.SaldoValorOficial
                        .SaldoValorMoeda = ProdutoNota.SaldoValorMoeda
                        .SaldoPedidoFiscal = ProdutoNota.SaldoPedidoFiscal
                        .SaldoPedidoFisico = ProdutoNota.SaldoPedidoFisico

                        .UnitarioMoeda = ProdutoNota.UnitarioMoeda
                        .Unitario = ProdutoNota.Unitario
                        '#FimBaseDeCalculo
                        '.ValorTotal = (.QuantidadeFiscal / ProdutoNota.Produto.BaseCalculo) * .Unitario
                        .ValorTotal = Math.Round(.QuantidadeFiscal * .Unitario, 2, MidpointRounding.AwayFromZero)

                        objItemNF.CarregandoEncargos = True
                        .Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objItemNF, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                        objItemNF.CarregandoEncargos = False

                        If .QuantidadeFiscal > 0 Or objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS Then objNotaFiscal.Itens.Add(objItemNF)

                    End With
                End If
            Next

            '**********************************************************************************************
            '*******************  Cria a estrutura da Tabela que alimenta o Grid  *************************
            '**********************************************************************************************
            Dim dtItens As New DataTable("Itens")
            dtItens.Columns.Add("Produto", Type.GetType("System.String"))
            dtItens.Columns.Add("NomeProduto", Type.GetType("System.String"))
            dtItens.Columns.Add("BaseCalculo", Type.GetType("System.String"))
            dtItens.Columns.Add("Lote", Type.GetType("System.String"))
            dtItens.Columns.Add("Classificacao", Type.GetType("System.String"))
            dtItens.Columns.Add("Embalagem", Type.GetType("System.String"))
            dtItens.Columns.Add("Saldo", Type.GetType("System.Decimal"))
            dtItens.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
            dtItens.Columns.Add("QuantidadeFisica", Type.GetType("System.Decimal"))
            dtItens.Columns.Add("Unitario", Type.GetType("System.Decimal"))
            dtItens.Columns.Add("Total", Type.GetType("System.Decimal"))
            '******************************************************************************************************
            '************************** Tabela que Alimenta o Grid Da Tela  ***************************************
            '******************************************************************************************************
            For Each objItemNF As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
                Dim drItem As DataRow = dtItens.NewRow()
                objNotaFiscal.Especie = objItemNF.Produto.Embalagem.Descricao
                drItem("Produto") = objItemNF.CodigoProduto
                If objItemNF.Produto.ControlarLote Or objItemNF.Produto.ControlarEmbalagem Then
                    drItem("NomeProduto") = "$" & objItemNF.Produto.Nome
                Else
                    drItem("NomeProduto") = objItemNF.Produto.Nome
                End If
                drItem("Lote") = objItemNF.Lote
                drItem("Classificacao") = objItemNF.Classificacao
                drItem("Embalagem") = objItemNF.CodigoEmbalagemIndea + "-" + objItemNF.CodigoTipoDeEmbalagem + "-" + objItemNF.CapacidadeEmbalagem.ToString

                '#FimBaseDeCalculo
                'drItem("BaseCalculo") = objItemNF.Produto.BaseCalculo
                drItem("BaseCalculo") = 1

                If Not objItemNF.SubOperacao.QuantidadeFiscal Then
                    drItem("Saldo") = objItemNF.SaldoValorOficial
                Else
                    drItem("Saldo") = objItemNF.SaldoPedidoFiscal
                End If

                drItem("Quantidade") = objItemNF.QuantidadeFiscal
                drItem("QuantidadeFisica") = objItemNF.QuantidadeFisica.ToString("N4")
                drItem("Unitario") = objItemNF.Unitario
                drItem("Total") = objItemNF.ValorTotal
                dtItens.Rows.Add(drItem)
            Next

            objNotaFiscal.AtualizaTotais()
            objNotaFiscal.CarregandoItens = False

            'Importação do xml de notas de entrada
            If SessaoDsXML IsNot Nothing Then
                Dim ObsTemp As String = objNotaFiscal.Observacoes
                objNotaFiscal.AtualizarObservacoes()
                objNotaFiscal.Observacoes &= ObsTemp
            Else
                objNotaFiscal.AtualizarObservacoes()
            End If
            txtObservacoesFiscais.Text = objNotaFiscal.Observacoes
            'txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque
            If String.IsNullOrWhiteSpace(observacaoAutorizacao.Value) Then
                txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoValores.Value
            Else
                txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque & ". " & observacaoAutorizacao.Value & ". " & observacaoValores.Value
            End If

            AtualizaFormulario()

            gridItens.DataSource = dtItens
            gridItens.DataBind()

            i = 0
            While i < gridItens.Rows.Count
                If Mid(gridItens.Rows(i).Cells(2).Text, 1, 1) = "$" Or objNotaFiscal.TemNotaTroca Then
                    Dim Campo As TextBox = CType(gridItens.Rows(i).FindControl("txtQuantidadeItem"), TextBox)
                    Campo.Enabled = False
                    gridItens.Rows(i).Cells(2).Text = gridItens.Rows(i).Cells(2).Text.Replace("$", "")
                End If
                i += 1
            End While
            SessaoSalvaNotaFiscal()
            TabContainer1.ActiveTabIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnReajusta_Click(sender As Object, e As EventArgs) Handles btnReajusta.Click
        Try
            SessaoRecuperaNotaFiscal()

            Dim vlrajustar As Decimal = CDec(txtReajuste.Text)
            For Each item In objNotaFiscal.Itens
                If objNotaFiscal.TotalNota < CDec(txtReajuste.Text) Then
                    item.ValorTotal += CDec(txtReajuste.Text) - objNotaFiscal.TotalNota
                    item.Unitario = Math.Round(item.ValorTotal / item.QuantidadeFiscal, 2, MidpointRounding.AwayFromZero)
                    Exit For
                Else
                    If item.ValorTotal < vlrajustar Then
                        vlrajustar -= item.ValorTotal
                        item.ValorTotal = 0
                        item.ValorLiquido = 0
                    Else
                        item.ValorTotal -= item.ValorTotal - CDec(txtReajuste.Text)
                        item.ValorLiquido -= item.ValorTotal - CDec(txtReajuste.Text)
                        vlrajustar -= CDec(txtReajuste.Text)
                    End If
                    If vlrajustar = 0 Then Exit For
                End If
            Next

            objNotaFiscal.Titulos.ReajFinanceiro = New ReajusteFinanceiro(objNotaFiscal)
            Dim resp As ArrayList = objNotaFiscal.Titulos.ReajFinanceiro.Reajusta()
            If resp(0) Then
                ucFinanceiro.AtualizaGridReajustaFinanceiro(objNotaFiscal.Titulos.ReajFinanceiro.TitulosNotaModificada)
            Else
                MsgBox(Me, resp(1))
            End If
            Dim sqls As New ArrayList
            objNotaFiscal.Titulos.SalvarSql(sqls)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

#Region "Alfasig"

    Private Sub Salvar()

        SessaoRecuperaNotaFiscal()

        objNotaFiscal.Codigo = CInt(txtNumero.Text)
        objNotaFiscal.NossaEmissao = chk_NossaEmissao.Checked
        objNotaFiscal.Eletronica = chk_nfe.Checked

        ''SegCodBarras
        If objNotaFiscal.TemSegCodBarra Then
            objNotaFiscal.SegCodBarra = txtSegCodBarra.Text.Replace(".", " ")
        End If

        If objNotaFiscal.ValidarNotaFiscal(SessaoDsXML, TabImportacao.Visible, erroMsg, False) Then

            'Enviar E-mail caso tenha chego no estoque mínimo - Furlan - 28/10/2020
            'Comentei onde enviava e trouxe aqui hoje dia 21/07/2022 - Furlan
            If CType(Session("objEstoqueMinimo" & HID.Value), DataTable).Rows.Count > 0 Then
                objNotaFiscal.enviarEstoqueMinimo(CType(Session("objEstoqueMinimo" & HID.Value), DataTable))
            End If

            lnkNovo.Parent.Visible = False
            objNotaFiscal.IUD = "I"

            'Solicitação Jonathan via e-mail dia 09/03/2021 - Furlan - 09/03/2021
            'If objNotaFiscal.CodigoEmpresa = "03189063000800" OrElse objNotaFiscal.CodigoEmpresa = "05366261000224" Then 'Colocado por Cáceres ser uma hora menos e estar no servidor da Matriz
            'If objNotaFiscal.Empresa.CodigoEstado = "MT" Then
            '    objNotaFiscal.DataInclusao = Format(Date.Now.AddHours(-1), "yyyy-MM-dd HH:mm:ss")
            'Else
            '    objNotaFiscal.DataInclusao = Format(Date.Now, "yyyy-MM-dd HH:mm:ss")
            'End If

            objNotaFiscal.HoraSefaz()

            'Metodo alimenta informações importantes do IBS e CBS
            objNotaFiscal.AtualizarObservacoes()
            txtObservacoesFiscais.Text = Funcoes.EliminarCaracteresEspeciaisNF(objNotaFiscal.Observacoes)
            'Metodo alimenta informações importantes do IBS e CBS

            'objNotaFiscal.Observacoes = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacoesFiscais.Text)
            objNotaFiscal.ObservacoesDeEmbarque = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacoesDeEmbarque.Text)

            If (objNotaFiscal.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE ICMS") OrElse objNotaFiscal.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE IPI")) AndAlso objNotaFiscal.TotalNota = 0 Then
                'LIBERADO PARA NOTA DE COMPLEMENTACAO DE ICMS OU IPI - FURLAN - 16/06/2025
            Else
                If objNotaFiscal.TotalNota > 0 Then
                    objNotaFiscal.Devolucao()
                End If
            End If

            'gridVencimentosNota.DataSource = objNotaFiscal.VencimentosNota.ToArray
            'Ajuste para Selecinar o Título e carregar no Financeiro - Furlan 30/07/2025
            gridVencimentosNota.DataSource = From tit In objNotaFiscal.VencimentosNota
                                             Select New With {.Codigo = tit.Codigo, .ReceberPagar = tit.ReceberPagar, .DescricaoMoeda = tit.DescricaoMoeda, .DescricaoProvisao = tit.DescricaoProvisao, .Prorrogacao = tit.Prorrogacao, .ValorDoDocumento = tit.ValorDoDocumento, .MoedaValorDoDocumento = tit.MoedaValorDoDocumento, .CodigoCifrado = IIf(tit.ReceberPagar = "P", Funcoes.Cifrar("ContasAPagar-" & tit.Codigo), Funcoes.Cifrar("ContasAReceber-" & tit.Codigo))}
            gridVencimentosNota.DataBind()

            'gridVencimentosPedido.DataSource = objNotaFiscal.VencimentosPedido.ToArray
            'Ajuste para Selecinar o Título e carregar no Financeiro - Furlan 30/07/2025
            gridVencimentosPedido.DataSource = From tit In objNotaFiscal.VencimentosPedido
                                               Select New With {.Codigo = tit.Codigo, .ReceberPagar = tit.ReceberPagar, .DescricaoMoeda = tit.DescricaoMoeda, .DescricaoProvisao = tit.DescricaoProvisao, .Prorrogacao = tit.Prorrogacao, .ValorDoDocumento = tit.ValorDoDocumento, .MoedaValorDoDocumento = tit.MoedaValorDoDocumento, .CodigoCifrado = IIf(tit.ReceberPagar = "P", Funcoes.Cifrar("ContasAPagar-" & tit.Codigo), Funcoes.Cifrar("ContasAReceber-" & tit.Codigo))}
            gridVencimentosPedido.DataBind()

            objNotaFiscal.ObservacoesEmbarque()

            If txtVolumes.Text.Length = 0 OrElse txtVolumes.Text = "0" Then
                objNotaFiscal.Quantidade = 1
            Else
                objNotaFiscal.Quantidade = txtVolumes.Text
            End If

            '<ver com Furlan>
            objNotaFiscal.Especie = txtEspecie.Text
            objNotaFiscal.Marca = txtMarca.Text
            objNotaFiscal.Numero = txtNumeracao.Text
            objNotaFiscal.PesoBruto = CDbl(IIf(txtPesoBruto.Text.Length = 0, 0, txtPesoBruto.Text))
            objNotaFiscal.PesoLiquido = CDbl(IIf(txtPesoLiquido.Text.Length = 0, 0, txtPesoLiquido.Text))

            'DESCOMENTE AS DUAS LINHAS ABAIXO CASO VÁ FAZER UMA NOTA COM DATA RETROATIVA, NÃO ESQUEÇA DE COMENTAR APÓS A EMISSÃO.
            'DEPOIS VÁ NO MÉTODO objNotaFiscal.Salvar em SalvarSql(Sqls) ("Sqls.Add(NumNota.IncrementarNumeradorSql())")
            'objNotaFiscal.DataNota = CDate("2014-11-30")
            'objNotaFiscal.Movimento = CDate("2014-11-30")

            If objNotaFiscal.PedidoBloqueado() Then
                MsgBox(Me.Page, "O pedido " & objNotaFiscal.Pedido.Codigo & " foi bloqueado por outro usuário, por favor recarregue o registro!")
                Exit Sub
            End If

            If Not objNotaFiscal.RealizarEstorno Then
                MsgBox(Me.Page, "Não foi possível realizar estorno no pedido!")
                Exit Sub
            End If

            'Arquivo ****************************************************************************************************************************
            '************************************************************************************************************************************
            If ucFile.Parent.Visible Then
                ucFile.Salvar(objNotaFiscal.Arquivos)
            End If

            objNotaFiscal.ObservacoesControleInterno = "INCLUSÃO PELO IP " & lblAcessoUsuario.Text

            If objNotaFiscal.Salvar() Then
                If chkEspelho.Checked Then
                    Dim espelho As New NotaFiscalEspelho()
                    espelho.ExibirEspelho(Me.Page, objNotaFiscal)
                End If

                If objNotaFiscal.Romaneio IsNot Nothing AndAlso
                    objNotaFiscal.Romaneio.Codigo > 0 AndAlso
                    (objNotaFiscal.Romaneio.Pesagens Is Nothing OrElse objNotaFiscal.Romaneio.Pesagens.Count = 0) Then
                    Funcoes.BindRomaneio(Me, objNotaFiscal.CodigoEmpresa, objNotaFiscal.CodigoRomaneio, objNotaFiscal.Itens(0).CodigoProduto)
                End If

                If objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
                    Dim fm As New FilesManager()
                    If fm.IsConnect() Then
                        Dim msgSefaz As String = String.Empty
                        If btnModo.Text.ToUpper().Trim() = "MODO NORMAL" Then
                            If VerificarNFe() Then
                                If EnviarSEFAZ(msgSefaz) Then
                                    MsgBox(Me.Page, "Nota fiscal " & objNotaFiscal.Codigo & " incluída com sucesso! " & msgSefaz, eTitulo.Sucess)
                                    LimparCampos("NFE", objNotaFiscal.Codigo, objNotaFiscal.NossaEmissao, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                                Else
                                    MsgBox(Me.Page, "Nota fiscal " & objNotaFiscal.Codigo & " incluída, porém ainda não foi HOMOLOGADA. " & msgSefaz)
                                    LimparCampos("NFE", objNotaFiscal.Codigo, objNotaFiscal.NossaEmissao, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                                End If
                            End If
                        Else
                            If EnviarSEFAZ(msgSefaz) Then
                                MsgBox(Me.Page, "Nota fiscal " & objNotaFiscal.Codigo & " incluída com sucesso! " & msgSefaz, eTitulo.Sucess)
                                LimparCampos("NFE", objNotaFiscal.Codigo, objNotaFiscal.NossaEmissao, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                            Else
                                MsgBox(Me.Page, "Nota fiscal " & objNotaFiscal.Codigo & " incluída, porém ainda não foi HOMOLOGADA. " & msgSefaz)
                                LimparCampos("NFE", objNotaFiscal.Codigo, objNotaFiscal.NossaEmissao, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                            End If
                        End If
                    Else
                        MsgBox(Me.Page, "Nota fiscal " & objNotaFiscal.Codigo & " incluída, porém Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor! Nota Fiscal ainda não foi HOMOLOGADA.")
                        LimparCampos("NFE", objNotaFiscal.Codigo, objNotaFiscal.NossaEmissao, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                    End If
                Else
                    LimparCampos("I", objNotaFiscal.Codigo, objNotaFiscal.NossaEmissao, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                End If
            End If
        Else
            MsgBox(Me.Page, erroMsg)
        End If
    End Sub

    Public Sub CancelarNotaFiscal(ByRef Sqls As ArrayList, ByRef nf As [Lib].Negocio.NotaFiscal)
        Sql = " DECLARE " & vbCrLf &
      " @Exist as varchar(1) " & vbCrLf &
      " SET @Exist = (select Case " & vbCrLf &
      "                        WHEN exists(SELECT Romaneios.Romaneio_Id " & vbCrLf &
      "                                      FROM NotasFiscaisXRomaneios nfxr" & vbCrLf &
      "                       			 	INNER JOIN Romaneios" & vbCrLf &
      "                                        ON nfxr.Empresa_Id    = Romaneios.Empresa_Id" & vbCrLf &
      "                                       AND nfxr.EndEmpresa_Id = Romaneios.EndEmpresa_Id" & vbCrLf &
      "                       	              AND nfxr.Romaneio_Id   = Romaneios.Romaneio_Id" & vbCrLf &
      "                       			    WHERE nfxr.Empresa_id      = '" & nf.CodigoEmpresa & "'" & vbCrLf &
      "                       	              AND nfxr.EndEmpresa_id   = " & nf.EnderecoEmpresa & vbCrLf &
      "                       	              AND nfxr.Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
      "                                       AND nfxr.EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
      "                                       AND nfxr.EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
      "                                       AND nfxr.Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
      "                                       AND nfxr.Nota_Id         = " & nf.Codigo & vbCrLf &
      "                                       AND Romaneios.Processo   = 'NOTA FISCAL')" & vbCrLf &
      "                          Then 'S' " & vbCrLf &
      "                          Else 'N' " & vbCrLf &
      "                     end); " & vbCrLf &
      " if @Exist = 'N' " & vbCrLf &
      "      DELETE NotasFiscaisXRomaneios" & vbCrLf &
      "       WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
      "         And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
      "         And Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
      "         And EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
      "         And EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
      "         And Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
      "         And Nota_Id         = " & nf.Codigo & "; " & vbCrLf &
      " ELSE " & vbCrLf &
      "      DELETE NotasFiscaisXRomaneios" & vbCrLf &
      "       WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
      "         And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
      "         And Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
      "         And EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
      "         And EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
      "         And Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
      "         And Nota_Id         = " & nf.Codigo & ";" &
      "      DELETE ROMANEIOSxDESCONTOS " & vbCrLf &
      "       WHERE Empresa_id    ='" & nf.CodigoEmpresa & "'" & vbCrLf &
      "         AND EndEmpresa_id = " & nf.EnderecoEmpresa & vbCrLf &
      "         AND ROMANEIO_ID   = (SELECT Romaneio_Id " & vbCrLf &
      "					               FROM NotasFiscaisXRomaneios " & vbCrLf &
      "					              WHERE Empresa_id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
      "					                AND EndEmpresa_id   = " & nf.EnderecoEmpresa & vbCrLf &
      "					                AND Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
      "					                AND EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
      "					                AND EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
      "					                AND Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
      "					                AND Nota_Id         = " & nf.Codigo & "); " & vbCrLf &
      "      DELETE ROMANEIOS " & vbCrLf &
      "       WHERE Empresa_id    ='" & nf.CodigoEmpresa & "'" & vbCrLf &
      "         And EndEmpresa_id = " & nf.EnderecoEmpresa & vbCrLf &
      "         And ROMANEIO_ID   = (SELECT Romaneio_Id " & vbCrLf &
      "					               FROM NotasFiscaisXRomaneios" & vbCrLf &
      "					              WHERE Empresa_id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
      "					                AND EndEmpresa_id   = " & nf.EnderecoEmpresa & vbCrLf &
      "					                AND Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
      "					                AND EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
      "					                AND EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
      "					                AND Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
      "					                AND Nota_Id         = " & nf.Codigo & ");" & vbCrLf
        Sqls.Add(Sql)

        Sql = "UPDATE NotasFiscaisXEncargos Set" & vbCrLf &
              "    Base       = 0," & vbCrLf &
              "    Percentual = 0," & vbCrLf &
              "    Valor      = 0 " & vbCrLf &
              " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   And Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   And EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   And EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   And Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   And Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        Sql = " UPDATE NotasFiscaisXItens Set" & vbCrLf &
              "    PesoFiscal       = 0," & vbCrLf &
              "    QuantidadeFisica = 0," & vbCrLf &
              "    QuantidadeFiscal = 0," & vbCrLf &
              "    Unitario         = 0," & vbCrLf &
              "    Valor            = 0 " & vbCrLf &
              " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   And Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   And EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   And EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   And Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   And Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        Dim obs As String = nf.ObservacoesControleInterno
        If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
            obs = obs & ". CANCELADO PELO IP " & lblAcessoUsuario.Text
        Else
            obs = "CANCELADO PELO IP " & lblAcessoUsuario.Text
        End If

        Sql = "UPDATE NotasFiscais Set" & vbCrLf &
              "   Situacao = 2 " & vbCrLf &
              "	 ,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
              "	 ,UsuarioAlteracaoData = getdate() " & vbCrLf &
              "  ,ObservacoesControleInterno  = '" & obs & "'" & vbCrLf &
              " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   And Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   And EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   And EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   And Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   And Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        Sql = "DELETE NotasXNotas" & vbCrLf &
              " WHERE OrigemEmpresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   And OrigemEndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   And OrigemCliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   And OrigemEndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   And OrigemEntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   And OrigemSerie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   And OrigemNota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        Sql = "DELETE NotasXNotas " & vbCrLf &
              " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   And Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   And EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   And EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   And Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   And Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        Sql = "DELETE NotaFiscalDevolucaoXNotaFiscal " & vbCrLf &
              " WHERE EmpresaDevolucao_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   And EndEmpresaDevolucao_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   And ClienteDevolucao_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   And EndClienteDevolucao_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   And EntradaSaidaDevolucao_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   And SerieDevolucao_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   And NotaDevolucao_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        Sql = "DELETE FROM NotaFiscalReferencial " & vbCrLf &
              "WHERE EmpresaReferencial_Id      = '" & nf.CodigoEmpresa & "'" & vbCrLf &
                "AND EndEmpresaReferencial_Id   = " & nf.EnderecoEmpresa & vbCrLf &
                "AND ClienteReferencial_Id      = '" & nf.CodigoCliente & "'" & vbCrLf &
                "AND EndClienteReferencial_Id   = " & nf.EnderecoCliente & vbCrLf &
                "AND EntradaSaidaReferencial_Id = '" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                "AND SerieReferencial_Id        = '" & nf.Serie & "'" & vbCrLf &
                "AND NotaReferencial_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        Sql = " Delete NotaFiscalXImportacao" & vbCrLf &
              " Where Empresa_Id      = '" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   and EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   and Cliente_Id      = '" & nf.CodigoCliente & "'" & vbCrLf &
              "   and EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   and EntradaSaida_Id = '" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   and Serie_Id        = '" & nf.Serie & "'" & vbCrLf &
              "   and Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        Sql = " Delete NotaFiscalXExportacao" & vbCrLf &
              " Where Empresa_Id      = '" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   and EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   and Cliente_Id      = '" & nf.CodigoCliente & "'" & vbCrLf &
              "   and EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   and EntradaSaida_Id = '" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   and Serie_Id        = '" & nf.Serie & "'" & vbCrLf &
              "   and Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        Sql = " DELETE NotasFiscaisXPercursos" & vbCrLf &
                "WHERE Empresa_Id      = '" & nf.CodigoEmpresa & "'" & vbCrLf &
                "  AND EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
                "  AND Cliente_Id      = '" & nf.CodigoCliente & "'" & vbCrLf &
                "  AND EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
                "  AND EntradaSaida_Id = '" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                "  AND Serie_Id        = '" & nf.Serie & "' " & vbCrLf &
                "  AND Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        If Not nf.SubOperacao.Devolucao Then

            Sql = "UPDATE ContasAPagar SET" & vbCrLf &
              "    ContasAPagar.Provisao = 3" & vbCrLf &
              "  FROM ContasApagar" & vbCrLf &
              " INNER JOIN (select Empresa_Id, EndEmpresa_Id, Cliente_id, Endcliente_Id, Titulo_Id" & vbCrLf &
              "               from NotaFiscalXTitulo " & vbCrLf &
              "				 where Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "				   and EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "				   and Cliente_id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "    			   and Endcliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "				   and EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "				   and Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "				   and Nota_id = " & nf.Codigo & vbCrLf &
              "             ) AS TxN " & vbCrLf &
              "    ON ContasAPagar.Empresa     = TxN.Empresa_Id " & vbCrLf &
              "	  AND ContasAPagar.EndEmpresa  = TxN.EndEmpresa_Id " & vbCrLf &
              "	  AND ContasAPagar.Cliente     = TxN.Cliente_id " & vbCrLf &
              "	  AND ContasAPagar.EndCliente  = TxN.Endcliente_Id " & vbCrLf &
              "	  AND ContasAPagar.Registro_Id = TxN.Titulo_Id " & vbCrLf &
              " WHERE ContasApagar.Provisao = 2; "
            Sqls.Add(Sql)

            Sql = "UPDATE ContasAReceber SET" & vbCrLf &
                  "    ContasAReceber.Provisao = 3" & vbCrLf &
                  "  FROM ContasAReceber " & vbCrLf &
                  " INNER JOIN (select Empresa_Id, EndEmpresa_Id, Cliente_id, Endcliente_Id, Titulo_Id" & vbCrLf &
                  "               from NotaFiscalXTitulo " & vbCrLf &
                  "				 where Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
                  "				   and EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
                  "				   and Cliente_id      ='" & nf.CodigoCliente & "'" & vbCrLf &
                  "    			   and Endcliente_Id   = " & nf.EnderecoCliente & vbCrLf &
                  "				   and EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                  "				   and Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
                  "				   and Nota_id         = " & nf.Codigo & vbCrLf &
                  "             ) AS TxN " & vbCrLf &
                  "    ON ContasAReceber.Empresa     = TxN.Empresa_Id " & vbCrLf &
                  "	  AND ContasAReceber.EndEmpresa  = TxN.EndEmpresa_Id " & vbCrLf &
                  "	  AND ContasAReceber.Cliente     = TxN.Cliente_id " & vbCrLf &
                  "	  AND ContasAReceber.EndCliente  = TxN.Endcliente_Id " & vbCrLf &
                  "	  AND ContasAReceber.Registro_Id = TxN.Titulo_Id " & vbCrLf &
                  " WHERE ContasAReceber.Provisao = 2; "
            Sqls.Add(Sql)

            Sql = "DELETE NotaFiscalXTitulo " & vbCrLf &
                  " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
                  "   And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
                  "   And Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
                  "   And EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
                  "   And EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                  "   And Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
                  "   And Nota_Id         = " & nf.Codigo & ";"
            Sqls.Add(Sql)
        Else
            'Retorna o valor aos titulos que foram alterados pela emissão da NF de Devolução
            If nf.NossaEmissao AndAlso (nf.VencimentosPedido.Any() OrElse nf.VencimentosNota.Any()) AndAlso Not nf.SubOperacao.Classe = eClassesOperacoes.REMESSAS Then
                Dim ReajFinanceiro = New ReajusteFinanceiro()
                ReajFinanceiro.CancelamentoNotaDeDevolucao(nf)
            End If

            'Cancelamento de NF de devolução.
            If nf.VencimentosNota.Any() AndAlso Not nf.SubOperacao.Classe = eClassesOperacoes.REMESSAS Then
                For Each tit In nf.VencimentosNota
                    Dim numerador = (tit.IUD = "I")
                    tit.SalvarSql(Sqls, numerador)
                Next
            End If
        End If

        Sql = "DELETE RAZAO" & vbCrLf &
          " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
          "   And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
          "   And Cliente_nf      ='" & nf.CodigoCliente & "'" & vbCrLf &
          "   And EndCliente_nf   = " & nf.EnderecoCliente & vbCrLf &
          "   And EntradaSaida_nf ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
          "   And Serie_nf        ='" & nf.Serie & "'" & vbCrLf &
          "   And Numero_nf       = " & nf.Codigo & vbCrLf &
          "   And Lote_Id        in (9,10,11); "
        Sqls.Add(Sql)

        Sql = "DELETE NFEPendencias " & vbCrLf &
              " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   And Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   And EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   And EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   And Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   And Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        Sql = "UPDATE NFERealizadas SET" & vbCrLf &
              "    Operacao   = 'Cancelado'," & vbCrLf &
              "    Retorno    = '101'," & vbCrLf &
              "    MsgRetorno = 'Cancelamento de NF-e homologado'" & vbCrLf &
              " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   AND EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   AND Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   AND EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   AND EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   AND Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   AND Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        'REMOVER TRANSFERÊNCIA NA ENTRADA - FURLAN - 05-07-2024
        If nf.EntradaSaida = eEntradaSaida.Saida AndAlso nf.SubOperacao.Classe = eClassesOperacoes.TRANSFERENCIAS Then

            Dim nfTransf As New NotaFiscal

            nfTransf.CodigoEmpresa = nf.CodigoCliente
            nfTransf.EnderecoEmpresa = nf.EnderecoCliente
            nfTransf.CodigoCliente = nf.CodigoEmpresa
            nfTransf.EnderecoCliente = nf.EnderecoEmpresa
            nfTransf.EntradaSaida = eEntradaSaida.Entrada
            nfTransf.Serie = nf.Serie
            nfTransf.Codigo = nf.Codigo
            nfTransf = New NotaFiscal(nfTransf)

            If nfTransf.CodigoOperacao = 31 Then
                nfTransf.IUD = "D"
                nfTransf.Salvar(Sqls)
            End If
        End If

        'INSERIR FINANCEIRO NOVO
        If FinanceiroNovo Then
            'Dim nfc As New NotaFiscal
            'nfc.CodigoEmpresa = nf.CodigoEmpresa
            'nfc.EnderecoEmpresa = nf.EnderecoEmpresa
            'nfc.CodigoCliente = nf.CodigoCliente
            'nfc.EnderecoCliente = nf.EnderecoCliente
            'nfc.EntradaSaida = nf.EntradaSaida
            'nfc.Serie = nf.Serie
            'nfc.Codigo = nf.Codigo

            'Dim nff As New NotaFiscal(nfc)
            'nff.IUD = "D"
            'Dim objRF As New ListReconstroiFinanceiroPedido(nff)
            'objRF.SalvarSql(Sqls)
        End If
    End Sub

    Private Function VerificarNFe() As Boolean
        Dim Sqls As New ArrayList
        Dim aux As Boolean = True
        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("status{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa)
        obj.Texto = String.Empty
        obj.SalvarSql(Sqls)

        If Not Banco.GravaBanco(Sqls) Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Return False
        End If

        'AGUARDAR RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Format("resp-status{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa)

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(90)

        While resp Is Nothing AndAlso Now < tempoLimite
            resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
            System.Threading.Thread.Sleep(3000)
        End While

        If resp IsNot Nothing Then
            Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

            Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            Sqls.Clear()

            If String.IsNullOrWhiteSpace(strCodigo) Then
                aux = False
                MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE.")
            ElseIf strCodigo = "107" Then
                aux = True

                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-status{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                Sqls.Add(Sql)
            ElseIf strCodigo = "4036" Then
                aux = False

                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-status{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                Sqls.Add(Sql)

                MsgBox(Me.Page, "Guardian em Modo de Contingência não pode ser usado para Emissão em Modo Normal.")
            Else
                aux = False
                MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
            End If

            If Sqls.Count > 0 Then Banco.GravaBanco(Sqls)
        Else
            aux = False
            MsgBox(Me.Page, "Verifique o Guardian no Servidor, não foi recebida nenhuma resposta do Modo de Operação.")
        End If

        Return aux
    End Function

    Private Function getDataSetEncargosXNotas(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal DrProdutosXNotas As DataRow) As DataSet
        Sql = "SELECT nfe.Encargo_Id AS Encargo," & vbCrLf &
              "       nfe.Base," & vbCrLf &
              "       nfe.Percentual," & vbCrLf &
              "       nfe.Valor," & vbCrLf &
              "       OE.STICMS as SituacaoTributaria," & vbCrLf &
              "       OE.Operacao," & vbCrLf &
              "       isnull(OEE.Sinal, '') AS Sinal," & vbCrLf &
              "       OE.STIPI as SituacaoTributariaIPI," & vbCrLf &
              "       OE.STPISCOFINS as SituacaoTributariaPISCOFINS" & vbCrLf &
              "  FROM NotasFiscaisXItens nfi" & vbCrLf &
              " INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
              "    ON nfi.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
              "   AND nfi.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
              "   AND nfi.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
              "   AND nfi.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
              "   AND nfi.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
              "   AND nfi.Serie_Id        = nfe.Serie_Id" & vbCrLf &
              "   AND nfi.Nota_Id         = nfe.Nota_Id" & vbCrLf &
              "   AND nfi.Produto_Id      = nfe.Produto_Id" & vbCrLf &
              "   AND nfi.CFOP_Id         = nfe.CFOP_Id" & vbCrLf &
              "   AND nfi.Sequencia_Id    = nfe.Sequencia_Id" & vbCrLf &
              " INNER JOIN OperacaoXEstado OE" & vbCrLf &
              "    on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
              "  LEFT JOIN OperacaoXEstadoxEncargo OEE" & vbCrLf &
              "    ON OEE.Codigo_Id  = nfi.OperacaoXEstado" & vbCrLf &
              "   and OEE.Encargo_Id = nfe.Encargo_Id" & vbCrLf &
              " WHERE (nfi.Empresa_Id      ='" & nf.CodigoEmpresa & "')" & vbCrLf &
              "   AND (nfi.EndEmpresa_Id   = " & nf.EnderecoEmpresa & ")" & vbCrLf &
              "   AND (nfi.Cliente_Id      ='" & nf.CodigoCliente & "')" & vbCrLf &
              "   AND (nfi.EndCliente_Id   = " & nf.EnderecoCliente & ")" & vbCrLf &
              "   AND (nfi.EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "')" & vbCrLf &
              "   AND (nfi.Serie_Id        ='" & nf.Serie & "')" & vbCrLf &
              "   AND (nfi.Nota_Id         = " & nf.Codigo & ")" & vbCrLf &
              "   AND (nfi.Produto_Id      ='" & DrProdutosXNotas("Produto") & "')" & vbCrLf &
              "   AND (nfi.CFOP_Id         = " & DrProdutosXNotas("CFOP") & ")" & vbCrLf &
              "   AND (nfi.Sequencia_Id    = " & DrProdutosXNotas("Sequencia_Id") & ")" & vbCrLf

        Return Banco.ConsultaDataSet(Sql, "ConsultaEncargos")
    End Function

    Private Function getDataSetNotasReferenciais(ByVal nf As [Lib].Negocio.NotaFiscal) As DataSet
        'INFORMAÇÕES DAS NOTAS FISCAIS DE DEVOLUÇÃO E NOTAS REFERENCIAIS DE EXPORTAÇÃO
        Sql = "SELECT * FROM " & vbCrLf &
                " ( " & vbCrLf &
                " SELECT ISNULL(RE.ChaveNFE,'') AS REFNFE, M.Estadoibge AS CUF,   " & vbCrLf &
                        " CONVERT(INT, RIGHT(CONVERT(VARCHAR, YEAR(NF.Movimento)),2) + RIGHT('00' + CONVERT(VARCHAR, MONTH(NF.Movimento)),2)) AS AAMM,  " & vbCrLf &
                        " '01' AS MOD, LEFT(RTRIM(NF.Empresa_Id), 14) AS CNPJ, NFR.Nota_Id AS NNF, NFR.Serie_Id AS Serie, Eletronica, Nf.TipoDeDocumento  " & vbCrLf &
                    " FROM NotasFiscais NF " & vbCrLf &
                    " INNER JOIN NotaFiscalDevolucaoXNotaFiscal NFR " & vbCrLf &
                    "   ON NFR.EmpresaDevolucao_Id      = NF.Empresa_Id  " & vbCrLf &
                    "   AND NFR.EndEmpresaDevolucao_Id  = NF.EndEmpresa_Id  " & vbCrLf &
                    "   AND NFR.ClienteDevolucao_Id      = NF.Cliente_Id   " & vbCrLf &
                    "   AND NFR.EndClienteDevolucao_Id   = NF.EndCliente_Id  " & vbCrLf &
                    "   AND NFR.EntradaSaida_Id = NF.EntradaSaida_Id  " & vbCrLf &
                    "   AND NFR.Serie_Id        = NF.Serie_Id  " & vbCrLf &
                    "   AND NFR.Nota_Id         = NF.Nota_Id  " & vbCrLf &
                    " INNER JOIN Clientes E " & vbCrLf &
                    "   ON NFR.EmpresaDevolucao_Id = E.Cliente_Id " & vbCrLf &
                    "   AND NFR.EndEmpresaDevolucao_Id = E.Endereco_Id " & vbCrLf &
                    " INNER JOIN Municipios M " & vbCrLf &
                    "   ON E.CodigoDoMunicipio = M.Codigo_id " & vbCrLf &
                    "   AND E.Estado            = M.Estado_id " & vbCrLf &
                    "   AND E.Cidade            = M.Municipio_id " & vbCrLf &
                    " LEFT JOIN NFERealizadas RE  " & vbCrLf &
                    "   ON NFR.EmpresaDevolucao_Id      = RE.Empresa_Id  " & vbCrLf &
                    "   AND NFR.EndEmpresaDevolucao_Id   = RE.EndEmpresa_Id  " & vbCrLf &
                    "   AND NFR.ClienteDevolucao_Id      = RE.Cliente_Id  " & vbCrLf &
                    "   AND NFR.EndClienteDevolucao_Id   = RE.EndCliente_Id  " & vbCrLf &
                    "   AND NFR.EntradaSaida_Id = RE.EntradaSaida_Id  " & vbCrLf &
                    "   AND NFR.Serie_Id        = RE.Serie_Id  " & vbCrLf &
                    "   AND NFR.Nota_Id         = RE.Nota_Id  " & vbCrLf &
                    " WHERE NFR.EmpresaDevolucao_Id = '" & nf.CodigoEmpresa & "'" & vbCrLf &
                    "   AND NFR.EndEmpresaDevolucao_Id = " & nf.EnderecoEmpresa & vbCrLf &
                    "   AND NFR.ClienteDevolucao_Id = '" & nf.CodigoCliente & "'" & vbCrLf &
                    "   AND NFR.EndClienteDevolucao_Id =" & nf.EnderecoCliente & vbCrLf &
                    "   AND NFR.NotaDevolucao_Id = " & nf.Codigo & vbCrLf &
                    "   AND NFR.EntradaSaidaDevolucao_Id = '" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                    "   AND NFR.SerieDevolucao_Id = '" & nf.Serie & "'" & vbCrLf &
                    " UNION ALL " & vbCrLf &
                    " SELECT ISNULL(RE.ChaveNFE,'') AS REFNFE, M.Estadoibge AS CUF,   " & vbCrLf &
                        " CONVERT(INT, RIGHT(CONVERT(VARCHAR, YEAR(NF.Movimento)),2) + RIGHT('00' + CONVERT(VARCHAR, MONTH(NF.Movimento)),2)) AS AAMM,  " & vbCrLf &
                        " '01' AS MOD, LEFT(RTRIM(NF.Empresa_Id), 14) AS CNPJ, NFR.Nota_Id AS NNF, NFR.Serie_Id AS Serie, Eletronica, Nf.TipoDeDocumento  " & vbCrLf &
                    " FROM NotaFiscalReferencial NFR  " & vbCrLf &
                    " INNER JOIN NotasFiscais NF  " & vbCrLf &
                    "   ON NFR.Empresa_Id      = NF.Empresa_Id   " & vbCrLf &
                    "   AND NFR.EndEmpresa_Id   = NF.EndEmpresa_Id   " & vbCrLf &
                    "   AND NFR.Cliente_Id      = NF.Cliente_Id   " & vbCrLf &
                    "   AND NFR.EndCliente_Id   = NF.EndCliente_Id   " & vbCrLf &
                    "   AND NFR.EntradaSaida_Id = NF.EntradaSaida_Id   " & vbCrLf &
                    "   AND NFR.Serie_Id        = NF.Serie_Id   " & vbCrLf &
                    "   AND NFR.Nota_Id         = NF.Nota_Id   " & vbCrLf &
                    " INNER JOIN Clientes E " & vbCrLf &
                    "   ON NFR.Empresa_Id = E.Cliente_Id " & vbCrLf &
                    "   AND NFR.EndEmpresa_Id = E.Endereco_Id " & vbCrLf &
                    " INNER JOIN Municipios M " & vbCrLf &
                    "   ON E.CodigoDoMunicipio = M.Codigo_id " & vbCrLf &
                    "   AND E.Estado            = M.Estado_id " & vbCrLf &
                    "   AND E.Cidade            = M.Municipio_id " & vbCrLf &
                    " LEFT JOIN NFERealizadas RE  " & vbCrLf &
                    "   ON NFR.Empresa_Id      = RE.Empresa_Id  " & vbCrLf &
                    "   AND NFR.EndEmpresa_Id   = RE.EndEmpresa_Id  " & vbCrLf &
                    "   AND NFR.Cliente_Id      = RE.Cliente_Id  " & vbCrLf &
                    "   AND NFR.EndCliente_Id   = RE.EndCliente_Id  " & vbCrLf &
                    "   AND NFR.EntradaSaida_Id = RE.EntradaSaida_Id  " & vbCrLf &
                    "   AND NFR.Serie_Id        = RE.Serie_Id  " & vbCrLf &
                    "   AND NFR.Nota_Id         = RE.Nota_Id  " & vbCrLf &
                    " WHERE NFR.EmpresaReferencial_Id = '" & nf.CodigoEmpresa & "'" & vbCrLf &
                    "   AND NFR.EndEmpresaReferencial_Id = " & nf.EnderecoEmpresa & vbCrLf &
                    "   AND NFR.ClienteReferencial_Id = '" & nf.CodigoCliente & "'" & vbCrLf &
                    "   AND NFR.EndClienteReferencial_Id =" & nf.EnderecoCliente & vbCrLf &
                    "   AND NFR.NotaReferencial_Id = " & nf.Codigo & vbCrLf &
                    "   AND NFR.EntradaSaidaReferencial_Id = '" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                    "   AND NFR.SerieReferencial_Id = '" & nf.Serie & "'" & vbCrLf &
                    "   AND NFR.TipoReferencial_Id = 'EXP'" & vbCrLf &
                " ) as t  " & vbCrLf &
                " GROUP BY t.REFNFE, t.CUF, t.AAMM, t.MOD, t.CNPJ, t.NNF, t.Serie, t.Eletronica, t.TipoDeDocumento " & vbCrLf &
                " ORDER BY t.Eletronica DESC" & vbCrLf
        Return Banco.ConsultaDataSet(Sql, "NotaFiscalReferencial")
    End Function

    Private Function getDataSetEncargos(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal DrProdutosXNotas As DataRow) As DataSet
        Sql = "SELECT NotasFiscaisXEncargos.Encargo_Id AS Encargo, NotasFiscaisXEncargos.Base, NotasFiscaisXEncargos.Percentual, isnull(NotasFiscaisXEncargos.PercentualExibicao,0) AS PercentualExibicao, " & vbCrLf &
              "       NotasFiscaisXEncargos.Valor, NotasFiscaisXEncargos.SituacaoTributaria, NotasFiscaisXEncargos.Operacao, " & vbCrLf &
              "       NotasFiscaisXEncargos.SituacaoTributariaIPI, NotasFiscaisXEncargos.SituacaoTributariaPISCOFINS " & vbCrLf &
              "  FROM NotasFiscaisXItens INNER JOIN " & vbCrLf &
              "       NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
              "       NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf &
              "       NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf &
              "       NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND " & vbCrLf &
              "       NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id AND NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id " & vbCrLf &
              " WHERE (NotasFiscaisXItens.Empresa_Id = '" & nf.CodigoEmpresa & "') AND (NotasFiscaisXItens.EndEmpresa_Id = " & nf.EnderecoEmpresa & ") AND " & vbCrLf &
              "	      (NotasFiscaisXItens.Cliente_Id = '" & nf.CodigoCliente & "') AND (NotasFiscaisXItens.EndCliente_Id = " & nf.EnderecoCliente & ") AND " & vbCrLf &
              "	      (NotasFiscaisXItens.EntradaSaida_Id = '" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "') AND (NotasFiscaisXItens.Serie_Id = '" & nf.Serie & "') AND " & vbCrLf &
              "       (NotasFiscaisXItens.Nota_Id = " & nf.Codigo & ") AND (NotasFiscaisXItens.Produto_Id = '" & DrProdutosXNotas("Produto") & "') AND " & vbCrLf &
              "       (NotasFiscaisXItens.CFOP_Id = " & DrProdutosXNotas("CFOP") & ") AND (NotasFiscaisXItens.Sequencia_Id = " & DrProdutosXNotas("Sequencia_Id") & ")" & vbCrLf
        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Function getTextoEmail(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal lstClientes As [Lib].Negocio.ListCliente, ByVal strAssunto As String, ByVal strMensagem As String, Optional ByVal compactado As Boolean = False) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)
        For Each objCliente As [Lib].Negocio.Cliente In lstClientes.Where(Function(s) Not String.IsNullOrWhiteSpace(s.EmailNFE)).ToList()
            sb.Append("DESTINATARIO=" & objCliente.EmailNFE & ControlChars.CrLf)
        Next
        sb.Append("ASSUNTO=" & strAssunto & " - NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        sb.Append("MENSAGEM=" & strMensagem & " - Envio NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        If Not String.IsNullOrWhiteSpace(nf.Empresa.EmailNFE) Then
            Dim strEmpresa As String = nf.Empresa.EmailNFE.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)(0).Trim()
            sb.Append("EMAILEMITENTE=" & strEmpresa & ControlChars.CrLf)
            sb.Append("NOMEEMITENTE=" & nf.Empresa.Nome & ControlChars.CrLf)
        End If
        sb.Append("ANEXOPDF=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOXML=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOPROTOCOLO=" & "SIM" & ControlChars.CrLf)
        sb.Append("COMPACTADO=" & IIf(compactado, "SIM", "NAO") & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Function GravarNFeXpress(ByVal nf As [Lib].Negocio.NotaFiscal) As Boolean
        Dim aux As Boolean = True
        Try
            Dim Sqls As New ArrayList
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"
            obj.NomeArquivo = String.Format("nota{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)

            obj.Texto = [Lib].Negocio.DocumentoEletronico.getTextoNFe4G(nf, HomologAlfasig, IIf(btnModo.Text.ToUpper().Trim() = "MODO NORMAL", 1, 2))

            If String.IsNullOrWhiteSpace(obj.Texto) Then
                MsgBox(Me.Page, "Não foi possível construir o arquivo texto para emissão da nota fiscal!")
                Return False
            End If
            obj.SalvarSql(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Return False
            End If
        Catch ex As Exception
            aux = False
            Throw New Exception(ex.Message)
        End Try
        Return aux
    End Function

    Private Function EnviarSEFAZ(ByRef msgNFE As String) As Boolean
        If objNotaFiscal Is Nothing Then
            msgNFE = "É necessário consultar um documento NF-e para realizar o envio para a SEFAZ!"
            Return False
        End If

        If GravarNFeXpress(objNotaFiscal) Then
            Dim obj As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = String.Format("resp-nota{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa)

            Dim tempoLimite As DateTime
            tempoLimite = Now.AddSeconds(90)

            While obj Is Nothing AndAlso Now < tempoLimite
                obj = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                System.Threading.Thread.Sleep(3000)
            End While

            If obj IsNot Nothing Then
                Dim lstMsg As String() = obj.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
                Dim chave As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("CHAVE")).FirstOrDefault()
                Dim recibo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RECIBO")).FirstOrDefault()
                Dim protocolo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NPROTOCOLO")).FirstOrDefault()
                Dim lote As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NUMEROLOTE")).FirstOrDefault()

                Dim strCodigo As String = String.Empty
                If resultado IsNot Nothing AndAlso resultado.Length > 0 Then
                    strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                Dim strMsg As String = String.Empty
                If mensagem IsNot Nothing AndAlso mensagem.Length > 0 Then
                    strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                Dim strChave As String = String.Empty
                If chave IsNot Nothing AndAlso chave.Length > 0 Then
                    strChave = chave.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                Dim strRecibo As String = String.Empty
                If recibo IsNot Nothing AndAlso recibo.Length > 0 Then
                    strRecibo = recibo.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                Dim strProtocolo As String = String.Empty
                If protocolo IsNot Nothing AndAlso protocolo.Length > 0 Then
                    strProtocolo = protocolo.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                Dim strLote As String = String.Empty
                If lote IsNot Nothing AndAlso lote.Length > 0 Then
                    strLote = lote.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                'ATUALIZAR NFE PENDENCIAS COM OS DADOS DO RETORNO
                Dim Sqls As New ArrayList


                If strCodigo = "4017" Then
                    'DANFE EM REGIME DE CONTIGÊNCIA
                ElseIf (Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "100" AndAlso strCodigo <> "110" AndAlso strCodigo <> "124" AndAlso strCodigo <> "302") _
                    Or Not strChave.Length = 44 _
                    Or Not strProtocolo.Length = 15 Then

                    lnkNovo.Parent.Visible = False
                    lnkEnviarSEFAZ.Parent.Visible = True
                    'lnkVisualizar.Parent.Parent.Parent.Visible = True
                    IdVisualizarNFe.Visible = False
                    IdModelNota.Visible = False
                    IdVisualizar.Visible = True

                    Sql = "UPDATE NFEPendencias " & vbCrLf &
                          "   SET ObservacoesFiscais = '" & objNotaFiscal.Observacoes & "', " & vbCrLf &
                          "       Retorno = '" & strCodigo & "', " & vbCrLf &
                          "       MsgRetorno = '" & strMsg.Replace("'", "") & "' " & vbCrLf
                    If strRecibo.Length > 0 Then Sql &= ", Recibo = '" & strRecibo & "' " & vbCrLf
                    If strChave.Length > 0 AndAlso strChave.Length = 44 Then Sql &= ", ChaveNfe = '" & strChave & "' " & vbCrLf
                    If strLote.Length > 0 Then Sql &= ", Lote = '" & strLote & "' " & vbCrLf
                    If strProtocolo.Length > 0 AndAlso strProtocolo.Length = 15 Then Sql &= ", Protocolo = '" & strProtocolo & "' " & vbCrLf
                    Sql &= "WHERE Empresa_Id = '" & objNotaFiscal.CodigoEmpresa & "' " & vbCrLf &
                           "  AND EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                           "  AND Cliente_Id = '" & objNotaFiscal.CodigoCliente & "' " & vbCrLf &
                           "  AND EndCliente_Id = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                           "  AND EntradaSaida_Id = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                           "  AND Serie_Id = '" & objNotaFiscal.Serie & "' " & vbCrLf &
                           "  AND Nota_Id = " & objNotaFiscal.Codigo & "; " & vbCrLf
                    Sqls.Add(Sql)

                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-status{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                    Sqls.Add(Sql)
                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-nota{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                    Sqls.Add(Sql)

                    If Banco.GravaBanco(Sqls) Then
                        msgNFE = String.Format("{0} - {1}", strCodigo, strMsg)
                    Else
                        msgNFE = HttpContext.Current.Session("ssMessage")
                    End If

                    Return False
                End If

                msgNFE = String.Format("{0} - {1}", strCodigo, strMsg)

                Sqls.Clear()

                Sql = "DELETE NFEContingencia " & vbCrLf &
                      " WHERE Empresa_Id = '" & objNotaFiscal.CodigoEmpresa & "' " & vbCrLf &
                      "   AND EndEmpresa_Id = '" & objNotaFiscal.EnderecoEmpresa & "' " & vbCrLf &
                      "   AND Cliente_Id = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                      "   AND EndCliente_Id = '" & objNotaFiscal.EnderecoCliente & "'" & vbCrLf &
                      "   AND EntradaSaida_Id = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                      "   AND Serie_Id = '" & objNotaFiscal.Serie & "' " & vbCrLf &
                      "   AND Nota_Id = '" & objNotaFiscal.Codigo & "'; " & vbCrLf
                Sqls.Add(Sql)

                If btnModo.Text.ToUpper().Trim() <> "MODO NORMAL" AndAlso (strCodigo = "4017" OrElse strCodigo = "124") Then
                    Sql = "INSERT INTO NFEContingencia "
                    Sql &= "(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, "
                    Sql &= "NfExpress, Operacao, Retorno, MsgRetorno, Recibo, ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, Protocolo, RegDPEC) "
                    Sql &= "VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ", '" & objNotaFiscal.CodigoCliente & "', '"
                    Sql &= objNotaFiscal.EnderecoCliente & "', '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "', '" & objNotaFiscal.Serie & "', '" & objNotaFiscal.Codigo & "', '"
                    Sql &= objNotaFiscal.DataInclusao.ToSqlDate() & "', '" & Format(objNotaFiscal.DataInclusao, "HH:mm:ss") & "', '" & UsuarioServidor.NomeUsuario & "', '"
                    Sql &= String.Format("nota{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "', 'INCLUIR', '"
                    Sql &= strCodigo & "', '" & strMsg & "', '" & strRecibo & "', '" & strChave & "', '" & strLote & "', '2', '" & objNotaFiscal.Observacoes & "', '', '" & strProtocolo & "', ''); "
                    Sqls.Add(Sql)
                Else
                    Sql = "INSERT INTO NFERealizadas "
                    Sql &= "(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, "
                    Sql &= "NfExpress, Operacao, Retorno, MsgRetorno, Recibo, ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, Protocolo, RegDPEC) "
                    Sql &= "VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ", '" & objNotaFiscal.CodigoCliente & "', '"
                    Sql &= objNotaFiscal.EnderecoCliente & "', '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "', '" & objNotaFiscal.Serie & "', '" & objNotaFiscal.Codigo & "', '"
                    Sql &= objNotaFiscal.DataInclusao.ToSqlDate() & "', '" & Format(objNotaFiscal.DataInclusao, "HH:mm:ss") & "', '" & UsuarioServidor.NomeUsuario & "', '"
                    Sql &= String.Format("nota{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "', 'INCLUIR', '"
                    Sql &= strCodigo & "', '" & strMsg & "', '" & strRecibo & "', '" & strChave & "', '" & strLote & "', '1', '" & objNotaFiscal.Observacoes & "', '', '" & strProtocolo & "', ''); "
                    Sqls.Add(Sql)
                End If

                Sql = "DELETE NFEPendencias " & vbCrLf &
                      " WHERE Empresa_Id = '" & objNotaFiscal.CodigoEmpresa & "' " & vbCrLf &
                      "   AND EndEmpresa_Id = '" & objNotaFiscal.EnderecoEmpresa & "'" & vbCrLf &
                      "   AND Cliente_Id = '" & objNotaFiscal.CodigoCliente & "' " & vbCrLf &
                      "   AND EndCliente_Id = '" & objNotaFiscal.EnderecoCliente & "' " & vbCrLf &
                      "   AND EntradaSaida_Id = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                      "   AND Serie_Id = '" & objNotaFiscal.Serie & "'" & vbCrLf &
                      "   AND Nota_Id = '" & objNotaFiscal.Codigo & "'; " & vbCrLf
                Sqls.Add(Sql)

                If strCodigo = "110" OrElse strCodigo = "302" Then
                    Sql = "UPDATE NotasFiscais Set Situacao = 10 "
                Else
                    Sql = "UPDATE NotasFiscais Set Situacao = 1 "
                End If

                Sql &= "WHERE Empresa_Id = '" & objNotaFiscal.CodigoEmpresa & "' " & vbCrLf &
                       "  AND EndEmpresa_Id = '" & objNotaFiscal.EnderecoEmpresa & "' " & vbCrLf &
                       "  AND Cliente_Id = '" & objNotaFiscal.CodigoCliente & "' " & vbCrLf &
                       "  AND EndCliente_Id = '" & objNotaFiscal.EnderecoCliente & "' " & vbCrLf &
                       "  AND EntradaSaida_Id = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                       "  AND Serie_Id = '" & objNotaFiscal.Serie & "' " & vbCrLf &
                       "  AND Nota_Id = '" & objNotaFiscal.Codigo & "'; " & vbCrLf
                Sqls.Add(Sql)

                If strCodigo = "110" OrElse strCodigo = "301" OrElse strCodigo = "302" OrElse strCodigo = "303" Then
                    objNotaFiscal.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)

                    If objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                        For Each tit In objNotaFiscal.VencimentosNota
                            If tit.CodigoSituacao = eSituacao.Normal AndAlso tit.CodigoProvisao = eProvisao.Previsao Then
                                Sql = "Update ContasAPagar Set Provisao = 3 " & vbCrLf &
                                        "Where Registro_id = " & tit.Codigo
                                Sqls.Add(Sql)

                                Sql = "Update ContasAReceber Set Provisao = 3 " & vbCrLf &
                                        "Where Registro_id = " & tit.Codigo
                                Sqls.Add(Sql)
                            End If
                        Next
                    End If

                    If objNotaFiscal.NotasReferenciais IsNot Nothing AndAlso objNotaFiscal.NotasReferenciais.Count > 0 Then
                        For Each item In objNotaFiscal.NotasReferenciais
                            item.IUD = "D"
                            item.SalvarSql(Sqls)
                        Next
                    End If

                    Sql = "Delete NotasFiscaisXTransportadores" & vbCrLf &
                          " Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                          "   and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                          "   and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                          "   and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                          "   and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                          "   and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                          "   and Nota_Id         = " & objNotaFiscal.Codigo
                    Sqls.Add(Sql)


                    Sql = " Delete NotasFiscaisXRomaneios" & vbCrLf &
                          "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                          "    and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                          "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                          "    and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                          "    and Romaneio_Id     = " & objNotaFiscal.CodigoRomaneio
                    Sqls.Add(Sql)

                End If

                If Not Banco.GravaBanco(Sqls) Then
                    msgNFE = HttpContext.Current.Session("ssMessage")
                    Return False
                End If
                'FIM DA ROTINA

                'Pega a Chave para Envio de Email
                objNotaFiscal.ChaveNFE = strChave

                SessaoSalvaNotaFiscal()

                Dim usoDenegado As Boolean = False

                If strCodigo = "110" OrElse strCodigo = "302" Then
                    msgNFE = msgNFE & "Uso Denegado : Irregularidade fiscal do destinatario. " & objNotaFiscal.Codigo & "-" & objNotaFiscal.Serie & "!"
                    usoDenegado = True
                ElseIf btnModo.Text.ToUpper().Trim() = "MODO NORMAL" Then

                    If Not HomologAlfasig AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.Empresa.EmailNFE) AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.Cliente.EmailNFE) Then
                        [Lib].Negocio.DocumentoEletronico.SendMailNFe(objNotaFiscal, msgNFE, False)
                    End If

                    If [Lib].Negocio.DocumentoEletronico.ImprimirNFe(objNotaFiscal, txtNumViasDeImpressao.Text, msgNFE, True) Then
                        Try
                            Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", objNotaFiscal.ChaveNFE), eTipoDeDocumento.Nota)
                            If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                                Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                                System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                                Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                            End If
                        Catch ex As Exception
                            msgNFE = msgNFE & "Não foi possível encontrar o arquivo DANFE da nota fiscal " & objNotaFiscal.Codigo & "-" & objNotaFiscal.Serie & "!"
                        End Try
                    End If
                Else
                    Try
                        Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", strChave), eTipoDeDocumento.Nota)
                        If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(strChave) Then
                            Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", strChave))
                            System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                            Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", strChave))
                        End If
                    Catch ex As Exception
                        msgNFE = msgNFE & "Não foi possível encontrar o arquivo DANFE da nota fiscal " & objNotaFiscal.Codigo & "-" & objNotaFiscal.Serie & "!"
                    End Try
                End If

                Sqls.Clear()
                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-status{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                Sqls.Add(Sql)
                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-nota{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                Sqls.Add(Sql)
                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-email{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                Sqls.Add(Sql)
                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-danfe{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                Sqls.Add(Sql)

                'Excluir tíitulo caso Nota seja Denegada - Furlan - 03/09/2021
                If usoDenegado Then
                    If objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                        For Each dr As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosNota
                            dr.IUD = "D"
                            dr.SalvarSql(Sqls, False)
                        Next

                        objNotaFiscal.VencimentosNota.NF = objNotaFiscal
                        'Atualiza a provisão do pedido.
                        For Each dr As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosPedido.Where(Function(t) t.CodigoProvisao = eProvisao.Provisao)
                            dr.IUD = "U"
                            dr.SalvarSql(Sqls, False)
                        Next
                    End If
                End If

                If Not Banco.GravaBanco(Sqls) Then
                    msgNFE = msgNFE & HttpContext.Current.Session("ssMessage")
                    Return False
                ElseIf Not String.IsNullOrEmpty(msgNFE) AndAlso Not strCodigo = "100" Then
                    Return False
                Else
                    Return True
                End If
            Else
                MsgBox(Me.Page, "Falha no retorno da Sefaz ref. a Homologação da NF-e, entre nas Pêndencias e verifique para realizar o reenvio à SEFAZ!")
                Return False
            End If
        End If
    End Function

    Private Function IsEnabled(ByVal nf As [Lib].Negocio.NotaFiscal) As Boolean
        Sql = "SELECT * FROM NFEPendencias " & vbCrLf &
            "WHERE 1=1 " & vbCrLf &
            "AND Empresa_Id = '" & nf.CodigoEmpresa & "' " & vbCrLf &
            "AND EndEmpresa_Id = '" & nf.EnderecoEmpresa & "' " & vbCrLf &
            "AND Cliente_Id = '" & nf.CodigoCliente & "' " & vbCrLf &
            "AND EndCliente_Id = '" & nf.EnderecoCliente & "' " & vbCrLf &
            "AND EntradaSaida_Id = '" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
            "AND Serie_Id = '" & nf.Serie & "' " & vbCrLf &
            "AND Nota_Id = '" & nf.Codigo & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "NFEPendencias")
        Return ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0
    End Function

    Private Sub EmContingencia(ByVal modo As [Lib].Negocio.eModo)
        Try
            SessaoRecuperaNotaFiscal()
            Dim Sqls As New ArrayList
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"
            obj.NomeArquivo = "contingencia.txt"
            obj.Texto = getTextoContingencia(objNotaFiscal, modo)
            obj.SalvarSql(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Exit Sub
            End If

            'AGUARDAR RESPOSTA SEFAZ
            Dim resp As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = "resp-contingencia.txt"

            Dim tempoLimite As DateTime
            tempoLimite = Now.AddSeconds(90)

            While resp Is Nothing AndAlso Now < tempoLimite
                resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                System.Threading.Thread.Sleep(3000)
            End While

            If resp IsNot Nothing Then
                Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                If Not String.IsNullOrWhiteSpace(strCodigo) Then
                    MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                    If modo = eModo.Contingencia AndAlso strCodigo = "4021" AndAlso strMsg.Contains("contingencia2") Then
                        btnModo.BackColor = Drawing.Color.Red
                        btnModo.Text = "MODO CONTINGÊNCIA"
                    ElseIf modo = eModo.Normal AndAlso strCodigo = "4021" AndAlso strMsg.Contains("normal") Then
                        btnModo.BackColor = Drawing.Color.Green
                        btnModo.Text = "MODO NORMAL"
                    End If
                End If
            Else
                MsgBox(Me.Page, "Falha no retorno ref. o Modo de Operação, verifique novamente.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getTextoContingencia(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal modo As [Lib].Negocio.eModo) As String
        Dim sb As New StringBuilder()
        sb.Append("MODALIDADE=" & IIf(modo = eModo.Normal, "0", "2") & ControlChars.CrLf)
        sb.Append("CNPJ=" & nf.CodigoEmpresa & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Public Function VerificarModo(ByVal Codigo As String, ByVal Empresa As String, ByVal NossaEmissao As Boolean, ByVal verMensagem As Boolean) As Boolean
        Dim aux As Boolean = True

        Try
            Dim Sqls As New ArrayList
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"
            obj.NomeArquivo = String.Format("status{0:000000000}#{1}.txt", Codigo, Empresa)
            obj.Texto = String.Empty
            obj.SalvarSql(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Return False
            End If

            'AGUARDAR RESPOSTA SEFAZ
            Dim resp As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = String.Format("resp-status{0:000000000}#{1}.txt", Codigo, Empresa)

            Dim tempoLimite As DateTime
            tempoLimite = Now.AddSeconds(90)

            While resp Is Nothing AndAlso Now < tempoLimite
                resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                System.Threading.Thread.Sleep(3000)
            End While

            If resp IsNot Nothing Then
                Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                Sqls.Clear()

                If String.IsNullOrWhiteSpace(strCodigo) Then
                    MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE.")
                    aux = False
                ElseIf strCodigo = "107" Then
                    btnModo.BackColor = Drawing.Color.Green
                    btnModo.Text = "MODO NORMAL"
                    aux = True
                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-status{0:000000000}#{1}.txt", Codigo, Empresa) & "';"
                    Sqls.Add(Sql)
                    If verMensagem Then MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg) & " - MODO NORMAL")
                ElseIf strCodigo = "4036" Then
                    btnModo.BackColor = Drawing.Color.Red
                    btnModo.Text = "MODO CONTINGÊNCIA"
                    aux = True
                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-status{0:000000000}#{1}.txt", Codigo, Empresa) & "';"
                    Sqls.Add(Sql)
                    If verMensagem Then MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg) & " - MODO CONTINGÊNCIA")
                Else
                    MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                    aux = False
                End If

                If Sqls.Count > 0 Then Banco.GravaBanco(Sqls)
            Else
                MsgBox(Me.Page, "Verifique o Guardian no Servidor, não foi recebida nenhuma resposta do Modo de Operação. ")
                aux = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            aux = False
        End Try

        Return aux
    End Function

    Private Function TemNFeBloqueadas() As Boolean
        Sql = "SELECT COALESCE(COUNT(*),0) as qtde " & vbCrLf &
              "  FROM NFEPendencias nfp " & vbCrLf &
              " INNER JOIN NotasFiscais nf " & vbCrLf &
              "    ON  nfp.Empresa_Id = nf.Empresa_Id " & vbCrLf &
              "   AND nfp.EndEmpresa_Id = nf.EndEmpresa_Id " & vbCrLf &
              "   AND nfp.Cliente_Id = nf.Cliente_Id " & vbCrLf &
              "   AND nfp.EndCliente_Id = nf.EndCliente_Id " & vbCrLf &
              "   AND nfp.EntradaSaida_Id = nf.EntradaSaida_Id " & vbCrLf &
              "   AND nfp.Serie_Id = nf.Serie_Id " & vbCrLf &
              "   AND nfp.Nota_Id = nf.Nota_Id " & vbCrLf &
              " WHERE nf.Situacao = 4"
        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "NFEPendencias")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            If CInt(ds.Tables(0).Rows(0)("qtde")) > 0 Then
                Return True
            End If
        End If
        Return False
    End Function

    Public Sub VerContingencia()
        SessaoRecuperaNotaFiscal()
        If VerificarModo(1, objNotaFiscal.CodigoEmpresa, False, False) Then
            If btnModo.Text.ToUpper().Trim() = "MODO NORMAL" Then VerificarContingencia()
        End If
    End Sub

    Private Sub VerificarContingencia()
        Try
            Dim Sqls As New ArrayList

            Sql = "SELECT NF.Empresa_Id, NF.EndEmpresa_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.EntradaSaida_Id, NF.Serie_id, NF.Nota_Id, " & vbCrLf &
                  "       NF.UsuarioInclusao, NF.UsuarioInclusaoData, NF.Observacoes, NFC.ChaveNfe " & vbCrLf &
                  "  FROM NFEContingencia NFC " & vbCrLf &
                  " INNER JOIN NotasFiscais AS NF" & vbCrLf &
                  "    ON NFC.Empresa_Id      = NF.Empresa_Id " & vbCrLf &
                  "   AND NFC.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf &
                  "   AND NFC.Cliente_Id      = NF.Cliente_Id " & vbCrLf &
                  "   AND NFC.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf &
                  "   AND NFC.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf &
                  "   AND NFC.Serie_Id        = NF.Serie_Id " & vbCrLf &
                  "   AND NFC.Nota_Id         = NF.Nota_Id " & vbCrLf &
                  " WHERE NF.Eletronica = 'S' " & vbCrLf &
                  "   AND NF.NossaEmissao = 'S'" & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Contingencia")
            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    Dim nf As New [Lib].Negocio.NotaFiscal()
                    nf.CodigoEmpresa = row("Empresa_Id")
                    nf.EnderecoEmpresa = row("EndEmpresa_Id")
                    nf.CodigoCliente = row("Cliente_Id")
                    nf.EnderecoCliente = row("EndCliente_Id")
                    nf.EntradaSaida = IIf(row("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                    nf.Serie = row("Serie_Id")
                    nf.Codigo = row("Nota_Id")
                    nf.UsuarioInclusao = row("UsuarioInclusao")
                    nf.DataInclusao = row("UsuarioInclusaoData")
                    nf.Observacoes = row("Observacoes")
                    nf.ChaveNFE = row("ChaveNfe")

                    Dim obj As New [Lib].Negocio.Fil()
                    obj.IUD = "I"
                    obj.NomeArquivo = String.Format("consulta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                    obj.Texto = [Lib].Negocio.DocumentoEletronico.getTextoConsultar(nf)
                    obj.SalvarSql(Sqls)

                    Dim Banco As New AcessaBanco
                    If Not Banco.GravaBanco(Sqls) Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        Exit Sub
                    End If

                    Dim resp As [Lib].Negocio.Resp = Nothing
                    Dim fileName As String = String.Format("resp-consulta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)

                    Dim tempoLimite As DateTime
                    tempoLimite = Now.AddSeconds(90)

                    While resp Is Nothing AndAlso Now < tempoLimite
                        resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                        System.Threading.Thread.Sleep(3000)
                    End While

                    If resp Is Nothing Then
                        fileName = String.Format("resp-consulta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                        tempoLimite = Now.AddSeconds(90)

                        While resp Is Nothing AndAlso Now < tempoLimite
                            resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                            System.Threading.Thread.Sleep(3000)
                        End While
                    End If

                    If obj IsNot Nothing Then
                        Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                        Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                        Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
                        Dim chave As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("CHAVE")).FirstOrDefault()
                        Dim recibo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RECIBO")).FirstOrDefault()
                        Dim protocolo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NPROTOCOLO")).FirstOrDefault()
                        Dim lote As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NUMEROLOTE")).FirstOrDefault()

                        Dim strCodigo As String = String.Empty
                        If resultado IsNot Nothing AndAlso resultado.Length > 0 Then
                            strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                        End If

                        Dim strMsg As String = String.Empty
                        If mensagem IsNot Nothing AndAlso mensagem.Length > 0 Then
                            strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                        End If

                        Dim strChave As String = String.Empty
                        If chave IsNot Nothing AndAlso chave.Length > 0 Then
                            strChave = chave.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                        End If

                        Dim strRecibo As String = String.Empty
                        If recibo IsNot Nothing AndAlso recibo.Length > 0 Then
                            strRecibo = recibo.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                        End If

                        Dim strProtocolo As String = String.Empty
                        If protocolo IsNot Nothing AndAlso protocolo.Length > 0 Then
                            strProtocolo = protocolo.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                        End If

                        Dim strLote As String = String.Empty
                        If lote IsNot Nothing AndAlso lote.Length > 0 Then
                            strLote = lote.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                        End If

                        If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo = "100" Then
                            Sqls.Clear()

                            Sql = "INSERT INTO NFERealizadas " & vbCrLf &
                                  "       (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, " & vbCrLf &
                                  "        NfExpress, Operacao, Retorno, MsgRetorno, Recibo, ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, Protocolo, RegDPEC) " & vbCrLf &
                                  "VALUES ('" & nf.CodigoEmpresa & "', " & nf.EnderecoEmpresa & ", '" & nf.CodigoCliente & "'," & vbCrLf &
                                  nf.EnderecoCliente & ", '" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "', '" & nf.Serie & "', " & nf.Codigo & "," & vbCrLf &
                                  "'" & nf.DataInclusao.ToSqlDate() & "', '" & Format(nf.DataInclusao, "HH:mm:ss") & "', '" & nf.UsuarioInclusao & "'," & vbCrLf &
                                  "'" & String.Format("nota{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa) & "', 'INCLUIR'," & vbCrLf &
                                  "'" & strCodigo & "', '" & strMsg & "', '" & strRecibo & "', '" & strChave & "', " & strLote & ", 2, '" & nf.Observacoes & "', '', '" & strProtocolo & "', ''); " & vbCrLf
                            Sqls.Add(Sql)

                            Sql = "DELETE FROM NFEContingencia " & vbCrLf &
                                  " WHERE Empresa_Id = '" & nf.CodigoEmpresa & "' AND EndEmpresa_Id = " & nf.EnderecoEmpresa & " AND " & vbCrLf &
                                  "       Cliente_Id = '" & nf.CodigoCliente & "' AND EndCliente_Id = " & nf.EnderecoCliente & " AND " & vbCrLf &
                                  "       EntradaSaida_Id = '" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND Serie_Id = '" & nf.Serie & "' AND " & vbCrLf &
                                  "       Nota_Id = " & nf.Codigo & "; " & vbCrLf
                            Sqls.Add(Sql)

                            Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-nota{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa) & "';"
                            Sqls.Add(Sql)

                            Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-consulta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa) & "';"
                            Sqls.Add(Sql)

                            If Not Banco.GravaBanco(Sqls) Then
                                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                                Exit Sub
                            End If
                        End If
                    Else
                        MsgBox(Me.Page, "Verifique o Guardian no Servidor, não foi recebida nenhuma resposta.")
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function EnviarNFePendentes() As Boolean
        Dim pendenciaOK As Boolean = True

        Try
            Sql = "SELECT nf.Empresa_Id, cast(nf.Nota_Id as nvarchar) + '-' + cast(nf.Serie_Id as nvarchar) as nf, nf.UsuarioInclusao, nf.UsuarioAlteracao, nfeP.Retorno, nfeP.MsgRetorno " & vbCrLf &
                  " FROM NotasFiscais nf  " & vbCrLf &
                  "		inner join NFEPendencias nfeP" & vbCrLf &
                  "				ON nfeP.Empresa_Id       = nf.Empresa_Id" & vbCrLf &
                  "		 		and nfeP.EndEmpresa_iD   = nf.EndEmpresa_Id" & vbCrLf &
                  "		 		and nfeP.Cliente_Id      = nf.Cliente_Id" & vbCrLf &
                  "		 		and nfeP.EndCliente_Id   = nf.EndCliente_Id" & vbCrLf &
                  "		 		and nfeP.EntradaSaida_Id = nf.EntradaSaida_Id" & vbCrLf &
                  "		 		and nfeP.Serie_Id        = nf.Serie_Id" & vbCrLf &
                  "		 		and nfeP.Nota_Id         = nf.Nota_Id" & vbCrLf &
                  "WHERE nf.Eletronica = 'S' " & vbCrLf &
                  "  AND nf.NossaEmissao = 'S' " & vbCrLf &
                  "  AND nf.TipoDeDocumento = 1 " & vbCrLf &
                  "  AND nf.Situacao in(4,7) " & vbCrLf &
                  "  AND datepart(hh, GETDATE() - nf.UsuarioInclusaoData) > 0 " & vbCrLf &
                  "  AND datepart(mi, GETDATE() - nf.UsuarioInclusaoData) > 1 " & vbCrLf &
                  "ORDER BY nf.Movimento DESC"

            Dim lstMail As New List(Of String)
            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "NotasFiscais")
            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                Dim Filial As String = String.Empty
                Dim bodyHTML = "Sr(s), <br/><br/> as notas fiscais a seguir estão pendentes de envio para a SEFAZ: <br/>" & vbCrLf
                bodyHTML &= "<ul>"
                For Each row As DataRow In ds.Tables(0).Rows
                    If String.IsNullOrEmpty(Filial) Then Filial = row("Empresa_Id")

                    bodyHTML &= "<li>"
                    bodyHTML &= "<label>" & row("Empresa_Id") & " - " & row("nf") & " (Usuário Inclusão: " & row("UsuarioInclusao") & ")" & "</label>"
                    bodyHTML &= "</li>"

                    bodyHTML &= "<li>"
                    bodyHTML &= "<label>RETORNO DA SEFAZ: " & row("Retorno") & " - " & row("MsgRetorno") & ")" & "</label>"
                    bodyHTML &= "</li>"

                    If Not IsDBNull(row("UsuarioInclusao")) AndAlso Not String.IsNullOrWhiteSpace(row("UsuarioInclusao")) Then
                        Dim objUsuario As New [Lib].Negocio.Usuario(row("UsuarioInclusao"))
                        If objUsuario IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objUsuario.Email) AndAlso objUsuario.Email.Trim() <> "&nbsp;" Then
                            lstMail.Add(objUsuario.Email)
                        End If
                    End If

                    If Not IsDBNull(row("UsuarioAlteracao")) AndAlso Not String.IsNullOrWhiteSpace(row("UsuarioAlteracao")) Then
                        Dim objUsuario As New [Lib].Negocio.Usuario(row("UsuarioAlteracao"))
                        If objUsuario IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objUsuario.Email) AndAlso objUsuario.Email.Trim() <> "&nbsp;" Then
                            lstMail.Add(objUsuario.Email)
                        End If
                    End If
                Next
                bodyHTML &= "</ul>"

                Dim Empresa = New [Lib].Negocio.Cliente(Filial, 0)
                Dim errorMsg As String = String.Empty
                Dim subject As String = String.Format("Empresa " & Empresa.CodigoFormatado & " (" & Empresa.Cidade & "/" & Empresa.CodigoEstado & ") - NF-e pendentes de envio para SEFAZ - {0:dd/MM/yyyy HH:mm}", DateTime.Now)
                Dim smtp = Funcoes.GetSmtpSettings()
                Dim fromMail = Funcoes.GetFromMail()

                Sql = "SELECT u.Email " & vbCrLf &
                      "  FROM ConfiguracaoXUsuario cxu " & vbCrLf &
                      " INNER JOIN Usuarios u " & vbCrLf &
                      "    ON (u.Usuario_Id = cxu.Usuario_Id) " & vbCrLf &
                      " WHERE cxu.Etapa_Id = " & eEtapa.NFePendencias

                Dim dsMail As DataSet = Banco.ConsultaDataSet(Sql, "ConfiguracaoXUsuario")
                If dsMail IsNot Nothing AndAlso dsMail.Tables IsNot Nothing AndAlso dsMail.Tables.Count > 0 AndAlso dsMail.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In dsMail.Tables(0).Rows
                        If Not String.IsNullOrWhiteSpace(row("Email")) Then
                            lstMail.Add(row("Email"))
                        End If
                    Next
                End If

                If lstMail IsNot Nothing AndAlso lstMail.Count > 0 Then
                    If Not Funcoes.SendMail(fromMail, "NGS SOLUÇÕES", lstMail, subject, bodyHTML, smtp, errorMsg) Then
                        MsgBox(Me.Page, "Erro ao tententar enviar E-Mail das Notas Pendentes. " & errorMsg, eTitulo.Erro)
                        pendenciaOK = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            pendenciaOK = False
        End Try

        Return pendenciaOK
    End Function

#End Region

#Region "Toolbar"
    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("NotaFiscalXItens", "GRAVAR") Then
                If Not String.IsNullOrWhiteSpace(txtES.Text) AndAlso Not String.IsNullOrWhiteSpace(txtNumero.Text) AndAlso Not String.IsNullOrWhiteSpace(txtSerie.Text) Then

                    SessaoRecuperaNotaFiscal()

                    If objNotaFiscal.Empresa.Empresa.ControlaDataMovimentoNFG AndAlso objNotaFiscal.Movimento <> DateTime.Today Then
                        If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Retroativa Then
                            'LIBERA EMISSÃO PARA NOTA RETROATIVA - FURLAN 03/04/2025
                        Else
                            MsgBox(Me.Page, "A Data de Movimento não pode ser diferente da Data de Hoje.")
                            Exit Sub
                        End If
                    End If

                    If objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
                        Dim sql As String = String.Empty

                        sql = "select N.Nota_Id, N.Movimento, N.Situacao " & vbCrLf &
                                "from NFEPendencias NFE " & vbCrLf &
                                "	INNER JOIN NotasFiscais N " & vbCrLf &
                                "		   ON N.Empresa_Id    = NFE.Empresa_Id" & vbCrLf &
                                "		AND N.EndEmpresa_Id   = NFE.EndEmpresa_Id" & vbCrLf &
                                "		AND N.Cliente_Id      = NFE.Cliente_Id" & vbCrLf &
                                "		AND N.EndCliente_Id   = NFE.EndCliente_Id" & vbCrLf &
                                "		AND N.EntradaSaida_Id = NFE.EntradaSaida_Id" & vbCrLf &
                                "		AND N.Serie_Id        = NFE.Serie_Id" & vbCrLf &
                                "		AND N.Nota_Id         = NFE.Nota_Id" & vbCrLf &
                                "WHERE N.Movimento < CONVERT(DATE, GETDATE())" & vbCrLf &
                                "AND N.TipoDeDocumento = 1"

                        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ConsultaPendencia")

                        If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                            MsgBox(Me.Page, "Existe Nota Fiscal com data inferior pendente de Homologação, verifique antes de proceder com essa emissão.")
                        Else
                            Salvar()
                        End If
                    Else
                        Salvar()
                    End If
                Else
                    MsgBox(Me.Page, "E/S, Número ou Série da Nota Fiscal não estão em conformidade, verifique.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

        lnkNovo.Visible = True
        lnkNovo.Enabled = True

    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("NotaFiscalXItens", "ALTERAR") Then
                If Not String.IsNullOrWhiteSpace(txtES.Text) AndAlso Not String.IsNullOrWhiteSpace(txtNumero.Text) AndAlso Not String.IsNullOrWhiteSpace(txtSerie.Text) Then
                    IniciarAlteracaoNotaFiscal()
                Else
                    MsgBox(Me.Page, "E/S, Número ou Série da Nota Fiscal não estão em conformidade, verifique.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("NotaFiscalXItens", "EXCLUIR") Then
                If Not String.IsNullOrWhiteSpace(txtES.Text) AndAlso Not String.IsNullOrWhiteSpace(txtNumero.Text) AndAlso Not String.IsNullOrWhiteSpace(txtSerie.Text) Then
                    IniciarExclusaoNotaFiscal()
                Else
                    MsgBox(Me.Page, "E/S, Número ou Série da Nota Fiscal não estão em conformidade, verifique.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("NotaFiscalXItens", "LEITURA") Then
                Dim operacao As String = ""
                If cmbSubOperacao.SelectedIndex > 0 Then
                    operacao = cmbSubOperacao.SelectedValue
                End If

                If Not String.IsNullOrWhiteSpace(txtCnpjDaEmpresa.Text) Then
                    Session("ssCnpjDaEmpresa" & HID.Value) = txtCnpjDaEmpresa.Text
                End If

                If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                    Dim cliente As String() = txtCodigoCliente.Value.Split("-")
                    If cliente IsNot Nothing AndAlso cliente.Length > 0 Then
                        Session("txtCnpjDoCliente" & HID.Value) = cliente(0)
                        Session("txtEndDoCliente" & HID.Value) = cliente(1)
                    End If
                End If

                Session("txtDataDeEmissao" & HID.Value) = txtDataDeEmissao.Text
                Session("txtDataDeEntrada" & HID.Value) = txtDataDeEntrada.Text
                Session("txtPedido" & HID.Value) = txtPedido.Text
                Session("txtOperacao" & HID.Value) = operacao
                Session("txtES" & HID.Value) = txtES.Text
                Session("txtSerie" & HID.Value) = Trim(txtSerie.Text)
                Session("txtNumero" & HID.Value) = txtNumero.Text
                Session("ssCampo" & HID.Value) = "NotaXItens"

                ucConsultaPedidosXNotas.BindGridView()
                Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaNXI" & HID.Value)

                lnkEnviarEmail.Parent.Visible = True
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            SessaoRecuperaNotaFiscal()
            LimparCampos("", "", False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEspelho_Click(sender As Object, e As EventArgs) Handles lnkEspelho.Click
        Try
            ImprimirEspelho()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lkn_ModeloNota_Click(sender As Object, e As EventArgs) Handles lkn_ModeloNota.Click
        Try
            Try

                'If Funcoes.VerificaPermissao("RelatorioDeTitulos", "RELATORIO") Then

                Dim Parametros As String = ""
                Dim crystal As String = ""
                Dim ds As DataSet

                ds = ConstroiDataSetNota()


                crystal = "Cr_NotaFiscal"


                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Bloco1", bloco1)
                parameters.Add("Bloco2", bloco2)
                parameters.Add("Bloco3", bloco3)
                'ucVencimentos.Limpar()
                'ucVencimentos.BindGridView(parameters)

                Funcoes.BindReport(Me.Page, ds, crystal, eExportType.PDF, parameters)
                'Else
                '    MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
                'End If

            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkServicoSefaz_Click(sender As Object, e As EventArgs) Handles lnkServicoSefaz.Click
        Try
            SessaoRecuperaNotaFiscal()
            If VerificarModo(1, objNotaFiscal.CodigoEmpresa, False, True) Then

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEnviarSEFAZ_Click(sender As Object, e As EventArgs) Handles lnkEnviarSEFAZ.Click
        Try
            Dim fm As New FilesManager()
            If fm.IsConnect() Then
                SessaoRecuperaNotaFiscal()

                Dim msgSefaz As String = String.Empty

                If EnviarSEFAZ(msgSefaz) Then
                    MsgBox(Me.Page, msgSefaz)
                    LimparCampos("NFE", objNotaFiscal.Codigo, objNotaFiscal.NossaEmissao, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                Else
                    MsgBox(Me.Page, msgSefaz)
                End If
            Else
                MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEnviarEmail_Click(sender As Object, e As EventArgs) Handles lnkEnviarEmail.Click
        Try
            SessaoRecuperaNotaFiscal()
            If Not String.IsNullOrWhiteSpace(objNotaFiscal.Empresa.EmailNFE) Then

                Dim parameters As New Dictionary(Of String, Object)
                parameters("Empresa") = objNotaFiscal.Empresa.Nome
                parameters("NF") = objNotaFiscal.Codigo
                parameters("ChaveNFE") = objNotaFiscal.ChaveNFE
                parameters("Emissao") = objNotaFiscal.DataNota.ToString("dd/MM/yyyy")
                parameters("Produto") = objNotaFiscal.Itens(0).Produto.Nome
                parameters("Quantidade") = objNotaFiscal.Quantidade & " " & objNotaFiscal.Especie
                parameters("ValorUnitario") = objNotaFiscal.TotalProduto.ToString("C")
                parameters("ValorTotal") = objNotaFiscal.TotalNota.ToString("C")
                parameters("Cliente") = objNotaFiscal.Cliente.Nome
                If Not String.IsNullOrWhiteSpace(objNotaFiscal.Pedido.PedidoEfetivo) Then
                    parameters("Pedido") = objNotaFiscal.Pedido.PedidoEfetivo
                Else
                    parameters("Pedido") = objNotaFiscal.Pedido.Codigo
                End If
                parameters("EmailNFE") = objNotaFiscal.Cliente.EmailNFE
                parameters("Placa") = objNotaFiscal.PlacaTransportador

                ucEmailNFe.Limpar()
                ucEmailNFe.EmailNFE(parameters)

                Dim txtDestinatario As TextBox = CType(ucEmailNFe.FindControlRecursive("txtDestinatario"), TextBox)
                Popup.ConsultaDeEmailNFe(Me.Page, "objEmailNFe" & HID.Value, txtDestinatario.ClientID, 100)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(sender As Object, e As EventArgs) Handles lnkPdf.Click
        Try
            Dim fm As New FilesManager()
            If fm.IsConnect() Then
                SessaoRecuperaNotaFiscal()
                Dim msgNFE As String = String.Empty
                If [Lib].Negocio.DocumentoEletronico.ImprimirNFeDanfe(objNotaFiscal, txtNumViasDeImpressao.Text, msgNFE) Then
                    Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", objNotaFiscal.ChaveNFE), eTipoDeDocumento.Nota)
                    If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                        Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                        System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                        Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                    End If

                    Dim Sqls As New ArrayList

                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-danfe{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                    Sqls.Add(Sql)

                    If Not Banco.GravaBanco(Sqls) Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If

                Else
                    MsgBox(Me.Page, msgNFE)
                End If
            Else
                MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkImpressora_Click(sender As Object, e As EventArgs) Handles lnkImpressora.Click
        Try
            Dim fm As New FilesManager()
            If fm.IsConnect() Then
                SessaoRecuperaNotaFiscal()
                Dim msgNFE As String = String.Empty
                If [Lib].Negocio.DocumentoEletronico.ImprimirNFeImpressora(objNotaFiscal, txtNumViasDeImpressao.Text, msgNFE) Then

                    Dim Sqls As New ArrayList

                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-danfe{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                    Sqls.Add(Sql)

                    If Banco.GravaBanco(Sqls) Then
                        MsgBox(Me.Page, "Danfe impressa com sucesso, verifique na impressora!")
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                Else
                    MsgBox(Me.Page, msgNFE)
                End If
            Else
                MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub bntConsultarNaviosXInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntConsultarNaviosXInvoice.Click
        Try
            SessaoRecuperaNotaFiscal()

            If objNotaFiscal.Itens.Count = 0 Then
                MsgBox(Me.Page, "Produto não foi selecionado.")
            Else
                ucConsultarNaviosXInvoice.Limpar()
                ucConsultarNaviosXInvoice.carregarNavios(objNotaFiscal.Itens(0).CodigoProduto)
                Popup.ConsultarNaviosXInvoice(Me.Page, "objNavioXInvoice" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkVisualizar_Click(sender As Object, e As EventArgs) Handles lnkVisualizar.Click
        Try
            SessaoRecuperaNotaFiscal()
            If objNotaFiscal IsNot Nothing Then
                Dim fileName As String = Server.MapPath(String.Format("~/Files/nota{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa))
                If System.IO.File.Exists(fileName) Then
                    System.IO.File.Delete(fileName)
                End If

                Dim str As String = String.Empty

                str = [Lib].Negocio.DocumentoEletronico.getTextoNFe4G(objNotaFiscal, HomologAlfasig, IIf(btnModo.Text.ToUpper().Trim() = "MODO NORMAL", 1, 2))

                Using sw As New StreamWriter(fileName)
                    sw.WriteLine(str)
                    sw.Close()
                End Using

                Response.Clear()
                Response.ClearHeaders()
                Response.AddHeader("content-length", str.Length.ToString())
                Response.ContentType = "text/plain"
                Response.AppendHeader("content-disposition", "attachment;filename=" & String.Format("nota{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa))
                Response.Write(str)
                Response.End()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Protected Sub btnModo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnModo.Click
    '    Try
    '        Dim fm As New FilesManager()
    '        If fm.IsConnect() Then
    '            If btnModo.Text.ToUpper().Trim() = "MODO NORMAL" Then
    '                EmContingencia(eModo.Contingencia)
    '            Else
    '                EmContingencia(eModo.Normal)
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    Protected Sub btnMonitorDeNotas_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMonitorDeNotas.Click
        Try
            Session("ssCampo" & HID.Value) = "NotaXItens"
            Popup.ConsultaMonitorDeNotas(Me.Page, "objNFConsultaNXI" & HID.Value)
            ucMonitorDeNotas.BindGridView()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "NotaFiscalXItens")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
#End Region

End Class
