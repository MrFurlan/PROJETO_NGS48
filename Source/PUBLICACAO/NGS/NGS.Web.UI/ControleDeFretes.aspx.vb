Imports System.Data
Imports System.Runtime.InteropServices.ComTypes
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ControleDeFretes
    Inherits BasePage

#Region "Variáveis"

    Private EmpresaPF As String()
    Private EmpresaProg As String()
    Private EmpresaLan As String()
    Private ConveniadoPF As String()
    Private ConveniadoProg As String()
    Private ClienteProg As String()
    Private ConveniadoLan As String()
    Private FaturasDeFreteLan As String()

    Dim iCons As Integer = 0
    Dim Nf As New [Lib].Negocio.NotaFiscal()
    Dim ItemFF As New [Lib].Negocio.FaturaDeFreteXItens

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ControleDeFretes", "ACESSAR") Then
                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "", True)
                ddl.Carregar(DdlEmpresaProg, CarregarDDL.Tabela.Empresas, "", True)
                ddl.Carregar(ddlFilial, CarregarDDL.Tabela.Empresas, "", True)
                ddl.Carregar(ddlCarteiraPF, CarregarDDL.Tabela.CarteiraFinanceira, " Classificacao = 'P'", True)
                Limpar(True)
                lnkLimparPR_Click(Nothing, Nothing)
                ddlEmpresa.SelectedIndex = ddlEmpresa.Items.IndexOf(ddlEmpresa.Items.FindByValue(UsuarioServidor.CodigoEmpresa & "-" & UsuarioServidor.EnderecoEmpresa))
                DdlEmpresaProg.SelectedIndex = DdlEmpresaProg.Items.IndexOf(DdlEmpresaProg.Items.FindByValue(UsuarioServidor.CodigoEmpresa & "-" & UsuarioServidor.EnderecoEmpresa))

                Dim objEmpresa As New [Lib].Negocio.Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
                Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objEmpresa)
                txtCodigoEmpresa.Value = itemEmpresa.Value
                txtEmpresa.Text = itemEmpresa.Text

                TabLancamento.ActiveTabIndex = 1
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub btnDadosBancarios_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDadosBancarios.Click
        If Not String.IsNullOrWhiteSpace(CodigoConveniadoPF.Value) Then
            Dim strCliente() = CodigoConveniadoPF.Value.Split("-")
            Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
            ucConsultaDadosBancarios.CarregaGrid(objCliente.Codigo, objCliente.CodigoEndereco)
            Popup.ConsultaDeDadosBancarios(Me.Page, "objContaBancariaFRT" & HID.Value)
        Else
            MsgBox(Me.Page, "É necessário selecionar um cliente!")
        End If
    End Sub

    Protected Sub btnConveniadoPF_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConveniadoPF.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objConveniadoPF" & HID.Value, "txtNome")
    End Sub

    Protected Sub lnkNovoPF_Click(sender As Object, e As EventArgs) Handles lnkNovoPF.Click
        Try
            SalvarPreFatura()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizarPF_Click(sender As Object, e As EventArgs) Handles lnkAtualizarPF.Click
        Try
            If ValidaCamposPreFatura() Then
                Dim SalvarFatura As [Lib].Negocio.FaturaDeFrete = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex)
                Dim Sqls As New ArrayList
                SalvarFatura.IUD = "U"
                SalvarFatura.ValorDaFatura = CDec(TxtValorFaturaPF.Text)
                CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).ValorDaFatura = CDec(TxtValorFaturaPF.Text)

                If SalvarFatura.JaFaturada Then
                    Dim objCotacao As New [Lib].Negocio.Cotacao(3, CDate(txtVencimentoPF.Text).ToString("yyyy-MM-dd"))
                    If Not FinanceiroNovo Then
                        SalvarFatura.ListTituloFatura(0).Titulo.Movimento = CDate(txtMovimentoPF.Text).ToString("yyyy-MM-dd")
                        SalvarFatura.ListTituloFatura(0).Titulo.Vencimento = CDate(txtVencimentoPF.Text).ToString("yyyy-MM-dd")
                        SalvarFatura.ListTituloFatura(0).Titulo.Prorrogacao = SalvarFatura.ListTituloFatura(0).Titulo.Vencimento
                        SalvarFatura.ListTituloFatura(0).Titulo.DataMoeda = CDate(txtVencimentoPF.Text).ToString("yyyy-MM-dd")
                        SalvarFatura.ListTituloFatura(0).Titulo.ValorDoDocumento = TxtValorFaturaPF.Text
                        SalvarFatura.ListTituloFatura(0).Titulo.MoedaValorDoDocumento = CDec(TxtValorFaturaPF.Text) / objCotacao.Indice
                        SalvarFatura.ListTituloFatura(0).Titulo.CodigoBancoCliente = txtCodBanco.Text.Trim
                        SalvarFatura.ListTituloFatura(0).Titulo.CodigoAgenciaCliente = txtCodAgencia.Text.Trim
                        SalvarFatura.ListTituloFatura(0).Titulo.DigitoAgenciaCliente = txtDigitoAgencia.Text.Trim
                        SalvarFatura.ListTituloFatura(0).Titulo.ContaCliente = txtConta.Text.Trim
                        SalvarFatura.ListTituloFatura(0).Titulo.DigitoContaCliente = txtDigitoConta.Text.Trim
                        SalvarFatura.ListTituloFatura(0).Titulo.UsuarioAlteracao = HttpContext.Current.Session("ssNomeUsuario")
                        SalvarFatura.ListTituloFatura(0).Titulo.UsuarioAlteracaoData = DateTime.Now
                        SalvarFatura.ListTituloFatura(0).Titulo.IUD = "U"
                        SalvarFatura.ListTituloFatura(0).Titulo.SalvarSql(Sqls, False)
                    Else
                        SalvarFatura.ListTituloFatura(0).TituloNovo.Movimento = CDate(txtMovimentoPF.Text).ToString("yyyy-MM-dd")
                        SalvarFatura.ListTituloFatura(0).TituloNovo.Vencimento = CDate(txtVencimentoPF.Text).ToString("yyyy-MM-dd")
                        SalvarFatura.ListTituloFatura(0).TituloNovo.Reprogramacao = SalvarFatura.ListTituloFatura(0).TituloNovo.Vencimento
                        SalvarFatura.ListTituloFatura(0).TituloNovo.DataMoeda = CDate(txtVencimentoPF.Text).ToString("yyyy-MM-dd")
                        SalvarFatura.ListTituloFatura(0).TituloNovo.Valores.EncargoValorDocumento.ValorOficial = CDec(TxtValorFaturaPF.Text)
                        SalvarFatura.ListTituloFatura(0).TituloNovo.Valores.EncargoValorLiquido.ValorMoeda = CDec(TxtValorFaturaPF.Text) / objCotacao.Indice
                        SalvarFatura.ListTituloFatura(0).TituloNovo.Valores.AtualizaLiquido()

                        SalvarFatura.ListTituloFatura(0).TituloNovo.CodigoBancoCliFor = txtCodBanco.Text.Trim
                        SalvarFatura.ListTituloFatura(0).TituloNovo.CodigoAgenciaCliFor = txtCodAgencia.Text.Trim
                        SalvarFatura.ListTituloFatura(0).TituloNovo.DigitoAgenciaCliFor = txtDigitoAgencia.Text.Trim
                        SalvarFatura.ListTituloFatura(0).TituloNovo.ContaCliFor = txtConta.Text.Trim
                        SalvarFatura.ListTituloFatura(0).TituloNovo.DigitoContaCliFor = txtDigitoConta.Text.Trim
                        SalvarFatura.ListTituloFatura(0).TituloNovo.IUD = "U"
                        SalvarFatura.ListTituloFatura(0).TituloNovo.SalvarSql(Sqls, False)
                    End If

                    SalvarFatura.SalvarSql(Sqls)

                    If Banco.GravaBanco(Sqls) Then
                        LimparPreFaturas()
                        DgPreFaturas.DataSource = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)
                        DgPreFaturas.DataBind()
                        MsgBox(Me.Page, "Pré-fatura alterada com Sucesso.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, "Ocorreu um erro ao alterar pré-fatura, verifique: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)))
                    End If
                Else
                    MsgBox(Me.Page, "Pré-fatura não existe!")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluirPF_Click(sender As Object, e As EventArgs) Handles lnkExcluirPF.Click
        Try
            If ValidaCamposPreFatura() Then
                Dim SalvarFatura As [Lib].Negocio.FaturaDeFrete = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex)
                Dim Sqls As New ArrayList
                SalvarFatura.IUD = "D"
                If SalvarFatura.JaFaturada Then
                    If Not FinanceiroNovo Then
                        Dim t As [Lib].Negocio.Titulo = SalvarFatura.ListTituloFatura(0).Titulo
                        t.UsuarioCancelamento = HttpContext.Current.Session("ssNomeUsuario")
                        t.UsuarioCancelamentoData = DateTime.Now
                        t.IUD = "D"
                        t.SalvarSql(Sqls, False)
                    Else
                        Dim t As [Lib].Negocio.Novo.TituloNovo = SalvarFatura.ListTituloFatura(0).TituloNovo
                        t.IUD = "D"
                        t.SalvarSql(Sqls, False)
                    End If

                    SalvarFatura.SalvarSql(Sqls)

                    If Banco.GravaBanco(Sqls) Then
                        LimparPreFaturas()
                        For Each dr As [Lib].Negocio.FaturaDeFrete In CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)
                            If dr.CodigoFatura = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).CodigoFatura Then
                                CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete).Remove(dr)
                                Exit For
                            End If
                        Next
                        DgPreFaturas.DataSource = Session("ssFaturasPF" & HID.Value)
                        DgPreFaturas.DataBind()
                        MsgBox(Me.Page, "Pré-fatura excluída com Sucesso.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, "Ocorreu um erro ao excluída pré-fatura, verifique: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)))
                    End If
                Else
                    MsgBox(Me.Page, "Pré-fatura não existe!")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultarPF_Click(sender As Object, e As EventArgs) Handles lnkConsultarPF.Click
        Try
            CarregaPreFaturas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorioPF_Click(sender As Object, e As EventArgs) Handles lnkRelatorioPF.Click

    End Sub

    Protected Sub lnkLimparPF_Click(sender As Object, e As EventArgs) Handles lnkLimparPF.Click
        Try
            LimparPreFaturas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjudaPF_Click(sender As Object, e As EventArgs) Handles lnkAjudaPF.Click

    End Sub

    Protected Sub DgPreFaturas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not Session("ssFaturasPF" & HID.Value) Is Nothing Then
            Dim objEmpresa As New [Lib].Negocio.Cliente(CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).CodigoEmpresa, CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).EnderecoEmpresa)
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objEmpresa)
            txtCodigoEmpresa.Value = itemEmpresa.Value
            txtEmpresa.Text = itemEmpresa.Text

            Dim Conveniado As New [Lib].Negocio.Cliente(CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).CodigoConveniado, CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).EnderecoConveniado)
            Dim itemConveniado As ListItem = Funcoes.FormatarListItemCliente(Conveniado)
            txtConveniadoPF.Text = itemConveniado.Text
            CodigoConveniadoPF.Value = itemConveniado.Value

            TxtFaturaPF.Text = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).CodigoFatura
            TxtValorFaturaPF.Text = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).ValorDaFatura
            txtMovimentoPF.Text = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).Movimento
            txtVencimentoPF.Text = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).Vencimento
            ddlCarteiraPF.Visible = Not FinanceiroNovo
            If Not FinanceiroNovo Then
                ddlCarteiraPF.SelectedValue = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).ListTituloFatura(0).Titulo.CodigoCarteira
            End If


            Dim ClienteXConta As New [Lib].Negocio.ListClienteXContaBancaria(Conveniado)
            For Each row As [Lib].Negocio.ClienteXContaBancaria In ClienteXConta
                txtCodBanco.Text = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).ListTituloFatura(0).Titulo.CodigoBancoCliente
                If CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).ListTituloFatura(0).Titulo.CodigoBancoCliente = row.CodigoBanco Then txtNomeBanco.Text = row.NomeBanco
                txtCodAgencia.Text = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).ListTituloFatura(0).Titulo.CodigoAgenciaCliente
                txtDigitoAgencia.Text = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).ListTituloFatura(0).Titulo.DigitoAgenciaCliente
                If CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).ListTituloFatura(0).Titulo.CodigoAgenciaCliente = CStr(row.CodigoAgencia) _
                   And CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).ListTituloFatura(0).Titulo.DigitoAgenciaCliente = row.DigitoAgencia Then txtPracaAgencia.Text = row.Praca
                txtConta.Text = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).ListTituloFatura(0).Titulo.ContaContabilCliente
                txtDigitoConta.Text = CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).ListTituloFatura(0).Titulo.DigitoContaCliente
            Next

            If CType(Session("ssFaturasPF" & HID.Value), [Lib].Negocio.ListFaturaDeFrete)(DgPreFaturas.SelectedIndex).ListTituloFatura(0).Titulo.CodigoProvisao = 3 Then
                lnkAtualizarPF.Enabled = True
                lnkExcluirPF.Enabled = True
                lnkNovoPF.Enabled = False
            Else
                lnkAtualizarPF.Enabled = False
                lnkExcluirPF.Enabled = False
                lnkNovoPF.Enabled = True
                MsgBox(Me.Page, "Pré-fatura não pode ser alterada/excluída!")
            End If
        End If
    End Sub

    Protected Sub btnConveniadoProg_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConveniadoProg.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objConveniadoPROG" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnClienteProg_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClienteProg.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClientePROG" & HID.Value, "txtNome")
    End Sub

    Protected Sub lnkConsultarPR_Click(sender As Object, e As EventArgs) Handles lnkConsultarPR.Click
        Try
            ConsultarProg()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparPR_Click(sender As Object, e As EventArgs) Handles lnkLimparPR.Click
        Try
            CodigoConveniadoProg.Value = String.Empty
            txtConveniadoProg.Text = String.Empty
            CodigoClienteProg.Value = String.Empty
            txtClienteProg.Text = String.Empty
            txtNumeroFatura.Text = String.Empty
            txtVencimentoDe.Text = "01/" & DateTime.Now.ToString("MM/yyyy")
            txtVencimentoAte.Text = DateTime.Now.ToString("dd/MM/yyyy")
            rdAtivos.Checked = True
            DgProgFaturas.DataSource = Nothing
            DgProgFaturas.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjudaPR_Click(sender As Object, e As EventArgs) Handles lnkAjudaPR.Click

    End Sub

    Protected Sub btnConveniado_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConveniadoLan.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objConveniadoLan" & HID.Value, "txtNome")
    End Sub

    Protected Sub imgFatura_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If ValidaCamposLanFatura() Then
            Me.LimparParcial()
            If ddlEmpresa.SelectedIndex > 0 Then
                EmpresaLan = ddlEmpresa.SelectedValue.ToString.Split("-")
            End If
            ucConsultaFaturasDeFrete.Limpar()
            Popup.ConsultaDeFaturasDeFrete(Me.Page, "ConsultaFatura" & HID.Value)
            Dim parameters As New Dictionary(Of String, Object)
            parameters.Add("Num", txtFatura.Text.Trim())
            parameters.Add("Emp", EmpresaLan(0))
            parameters.Add("EndEmp", EmpresaLan(1))
            If CodigoConveniadoLan.Value.ToString.Length > 0 Then
                ConveniadoLan = CodigoConveniadoLan.Value.ToString.Split("-")
                parameters.Add("Conv", ConveniadoLan(0))
                parameters.Add("EndConv", ConveniadoLan(1))
            Else
                parameters.Add("Conv", "")
                parameters.Add("EndConv", 0)
            End If
            ucConsultaFaturasDeFrete.BindGridView(parameters)
        End If
    End Sub

    Protected Sub BtnNumeroNotaFrete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnNumeroNotaFrete.Click
        If ValidaCamposLanFatura() Then
            If String.IsNullOrWhiteSpace(TxtNrFrete.Text) Then
                MsgBox(Me.Page, "É necessário informar o número do conhecimento de transporte!")
                Exit Sub
            End If

            If Session("Fatura" & HID.Value) Is Nothing Then
                MsgBox(Me.Page, "É necessário selecionar a fatura!")
                Exit Sub
            End If

            Dim pTipo As String = ""
            Dim NumNota As String = ""

            'If RbRecibo.Checked = True Then
            '    pTipo = "3"
            'ElseIf RbCtrc.Checked = True Then
            '    pTipo = "2"
            'End If

            pTipo = 57

            If ddlEmpresa.SelectedIndex > 0 Then
                EmpresaLan = ddlEmpresa.SelectedValue.ToString.Split("-")
            End If

            If Not String.IsNullOrWhiteSpace(CodigoConveniadoLan.Value) Then
                ConveniadoLan = CodigoConveniadoLan.Value.ToString.Split("-")
            End If

            If Not String.IsNullOrWhiteSpace(TxtNrFrete.Text) Then
                NumNota = TxtNrFrete.Text
            End If

            Dim parameters As New Dictionary(Of String, Object)
            parameters.Add("Tipo", pTipo)
            parameters.Add("Emp", EmpresaLan(0))
            parameters.Add("EndEmp", EmpresaLan(1))
            parameters.Add("Nota", NumNota)
            parameters.Add("PagarReceber", IIf(rdbPagarLancamento.Checked, "P", "R"))
            Popup.ConsultaDeNotasDeFrete(Me.Page, "ConsultaNotasDeFrete" & HID.Value)
            ucConsultaNotasDeFrete.BindGridView(parameters)
        End If
    End Sub

    Protected Sub btnOk_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SalvarFatura()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnExcluir_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        RecuperaItemFF()
        Dim Fatura As [Lib].Negocio.FaturaDeFrete = Session("Fatura" & HID.Value)
        If Fatura Is Nothing Then
            MsgBox(Me.Page, "É necessário selecionar uma fatura!")
            Exit Sub
        End If

        If Fatura.LancamentoManual = 0 AndAlso ItemFF.JaFaturadaRegistroMestre() Then
            MsgBox(Me.Page, Session("ssMessage"))
            Exit Sub
        End If

        If ItemFF IsNot Nothing Then

            Dim Sqls As New ArrayList
            ItemFF.IUD = "D"
            ItemFF.SalvarSql(Sqls)

            For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In ItemFF.Nota.Itens(0).Encargos
                If enc.Encargo.Etapa > 0 Then
                    enc.IUD = "D"
                    enc.SalvarSql(Sqls)
                End If
            Next

            Fatura.AtualizarVinculos(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, Session("ssMessage"))
            Else
                Fatura.Itens.Remove(ItemFF)
                Sqls.Clear()
                Dim razao As New [Lib].Negocio.Razao
                razao.ExcluirFretesRazao(ItemFF.Nota)
                razao.ContabilizarFretesNoRazao(ItemFF.Nota)
                Dim ListNfDeFrete As [Lib].Negocio.ListFaturasDeFretesXItens
                ListNfDeFrete = New [Lib].Negocio.ListFaturasDeFretesXItens(Fatura)
                DgComposicaoFatura.DataSource = ListNfDeFrete
                DgComposicaoFatura.DataBind()
                btnEncerrar.Enabled = ListNfDeFrete IsNot Nothing AndAlso ListNfDeFrete.Count > 0
                LimparParcial()

            End If

        End If

    End Sub

    Protected Sub DgComposicaoFatura_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgComposicaoFatura.SelectedIndexChanged
        Dim objFatura As [Lib].Negocio.FaturaDeFrete = Session("Fatura" & HID.Value)
        ItemFF = objFatura.Itens(DgComposicaoFatura.SelectedIndex)
        CarregarCTE(ItemFF, True)
    End Sub

    Protected Sub btnLimparLancamento_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLimparLancamento.Click
        LimparParcial()
    End Sub

    Protected Sub txtPesoChegada_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            RecuperaItemFF()
            Dim NfDeFrete As [Lib].Negocio.NotaFiscal = Session("NotasDeFrete" & HID.Value)
            ItemFF.PesoDeChegada = txtPesoChegada.Text

            If ItemFF.PesoDeChegada > 0 AndAlso (Math.Abs(ItemFF.Peso - ItemFF.PesoDeChegada) <= CDec(txtTolerancia.Text) OrElse Math.Abs(ItemFF.PesoDeChegada - ItemFF.Peso) <= CDec(txtTolerancia.Text)) Then
                ItemFF.FreteChegada = ItemFF.ValorFrete
            ElseIf Left(Session("ssEmpresa"), 8) = "62780383" AndAlso ItemFF.Nota.CodigoTransportador = "63358210000176" Then
                ItemFF.FreteChegada = ItemFF.ValorFrete
            End If

            SalvarItemFF()

            txtFreteChegada.Text = String.Format("{0:N2}", ItemFF.FreteChegada)
            txtDiferenca.Text = String.Format("{0:N2}", ItemFF.Diferenca)

            For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In NfDeFrete.Itens(0).Encargos
                If ItemFF.PesoDeChegada > 0 AndAlso (Math.Abs(ItemFF.Peso - ItemFF.PesoDeChegada) > CDec(txtTolerancia.Text) OrElse Math.Abs(ItemFF.PesoDeChegada - ItemFF.Peso) > CDec(txtTolerancia.Text)) Then
                    If enc.Codigo.Trim = IIf(CInt(txtDiferenca.Text) > 0, "SOBRAS", "QUEBRAS") Then
                        enc.Valor = Math.Abs(CDec(txtDiferenca.Text))
                        Exit For
                    End If
                End If
            Next

            If String.IsNullOrWhiteSpace(TxtValorAdiantamento.Text) Then
                Dim Adiantamento As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
                Dim lstEncargos As List(Of NotaFiscalXItemXEncargo) = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(NfDeFrete.Itens(0), New List(Of eEtapaEncago) From {eEtapaEncago.Adiantamento})
                For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In lstEncargos
                    If enc.Codigo = "ADTODEFRETE" Then
                        Adiantamento = enc
                        Exit For
                    End If
                Next

                If Not Adiantamento Is Nothing Then
                    txtPercAdiantamento.Text = Adiantamento.Percentual.ToString + "%"
                    TxtValorAdiantamento.Text = Adiantamento.Valor
                Else
                    txtPercAdiantamento.Text = "0%"
                    TxtValorAdiantamento.Text = 0
                End If
            End If

            NfDeFrete.Itens(0).Encargos.AtualizaLiquido(Math.Abs(ItemFF.ValorFrete - TxtValorAdiantamento.Text))

            If ItemFF.Parent.CodigoEmpresa <> NfDeFrete.CodigoEmpresa Then
                Dim TRANSFCARTA As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
                Dim LIQUIDOAPAGAR As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
                For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In NfDeFrete.Itens(0).Encargos
                    If enc.Codigo.Trim = "TRANSFCARTA" Then
                        TRANSFCARTA = enc
                    End If
                    If enc.Codigo.Trim = "LIQUIDOAPAGAR" Then
                        LIQUIDOAPAGAR = enc
                    End If
                Next
                TRANSFCARTA.Base = LIQUIDOAPAGAR.Base
            End If

            Session("NotasDeFrete" & HID.Value) = NfDeFrete
            CarregaDgEncargos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getParam() As String
        Dim param As String = "Parametros:" & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Dim objEmpresa = New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            param &= "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtConveniadoLan.Text) Then
            param &= "Código: " & txtConveniadoLan.Text
        End If

        Return param
    End Function

    Protected Sub btnCalcular_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Calcular()
    End Sub

    Protected Sub btnEncerrar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Encerrar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatório_Click(sender As Object, e As EventArgs) Handles lnkRelatório.Click
        Try
            If ValidaImpressao() Then
                Dim ds As DataSet
                Dim sql As String = GetSql()

                ds = Banco.ConsultaDataSet(sql, "FaturasDeFretes")

                'LOGOTIPO
                Dim dt As DataTable = ds.Tables.Add("Logotipo")
                dt.Columns.Add("path", GetType(String))
                dt.Columns.Add("Imagem", GetType(System.Byte()))
                dt.Columns.Add("Nome", GetType(String))
                dt.Columns.Add("Cidade", GetType(String))
                dt.Columns.Add("Estado_Id", GetType(String))
                Dim drImagem As DataRow = dt.NewRow()
                Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
                drImagem("path") = strCaminhoImagem
                drImagem("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
                drImagem("Nome") = HttpContext.Current.Session("ssNomeEmpresa")
                drImagem("Cidade") = HttpContext.Current.Session("ssCidadeEmpresa")
                drImagem("Estado_Id") = HttpContext.Current.Session("ssEstadoEmpresa")
                dt.Rows.Add(drImagem)

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Parametros", getParam())

                Funcoes.BindReport(Me.Page, ds, "Cr_FaturasDeFrete", eExportType.PDF, parameters)
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

    Protected Sub btnEmpresa_Click(sender As Object, e As EventArgs) Handles btnEmpresa.Click
        ucConsultaEmpresas.Limpar()
        Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaCDF" & HID.Value)
    End Sub

    Protected Sub optCartaFrete_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If FinanceiroNovo Then
                ucFinanceiro.Limpar()
                ucFinanceiro.AtualizaFaturaFrete(CDec(txtValorFrete.Text), CDec(txtFreteChegada.Text), CDate(txtVencimento.Text), False)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível programar o financeiro da fatura!")
        End Try
    End Sub

    Protected Sub optBaixaAdiantamento_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If FinanceiroNovo Then
                ucFinanceiro.Limpar()
                ucFinanceiro.AtualizaFaturaFrete(CDec(txtValorFrete.Text), CDec(txtFreteChegada.Text), CDate(txtVencimento.Text))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível programar o financeiro da fatura!")
        End Try
    End Sub

#End Region

#Region "Métodos"

    Public Sub RecuperaItemFF()
        ItemFF = Session("ItemFF" & HID.Value)
    End Sub

    Public Sub SalvarItemFF()
        Session("ItemFF" & HID.Value) = ItemFF
    End Sub

    Private Function GetSql() As String
        Dim sql As String = " SELECT FFxI.EmpresaPagadora_Id + '-' + CAST(FFxI.EndEmpresaPagadora_Id as nvarchar) as EmpresaFatura," & vbCrLf &
                            "        FFxI.Conveniado_Id + '-' + CAST(FFxI.EndConveniado_Id as nvarchar) as ConveniadoFatura," & vbCrLf &
                            "   	 FF.Movimento," & vbCrLf &
                            "   	 FFxI.Fatura_Id," & vbCrLf &
                            "   	 FFxT.Titulo_Id AS Titulo," & vbCrLf &
                            "        p.Descricao as Provisao," & vbCrLf &
                            "   	 FF.ValorDaFatura," & vbCrLf &
                            "   	 CAST(n.Nota_Id as nvarchar) + ' - ' + CAST(n.Serie_Id as nvarchar) as Conhecimento," & vbCrLf &
                            "   	 tp.Descricao as  TipoDeDocumento," & vbCrLf &
                            "   	 CAST(sub1.Operacao_Id as nvarchar) + '-' + CAST(sub1.SubOperacoes_Id as nvarchar) as Operacao," & vbCrLf &
                            "   	 CAST(ori.Nota_Id as nvarchar) + ' - ' + CAST(ori.Serie_Id as nvarchar) as OrigemNota," & vbCrLf &
                            "   	 CAST(sub2.Operacao_Id as nvarchar) + '-' + CAST(sub2.SubOperacoes_Id as nvarchar) as OperacaoNota" & vbCrLf &
                            "   FROM FaturasDeFretesXItens FFxI" & vbCrLf &
                            "   JOIN FaturasDeFretes FF" & vbCrLf &
                            "     ON FFxI.EmpresaPagadora_Id    = FF.Empresa_Id" & vbCrLf &
                            "    AND FFxI.EndEmpresaPagadora_Id = FF.EndEmpresa_Id" & vbCrLf &
                            "    AND FFxI.Conveniado_Id         = FF.Conveniado_Id" & vbCrLf &
                            "    AND FFxI.EndConveniado_Id      = FF.EndConveniado_Id" & vbCrLf &
                            "    AND FFxI.Fatura_Id             = FF.Fatura_Id" & vbCrLf &
                            "   JOIN FaturaDeFreteXTitulo FFxT" & vbCrLf &
                            "     ON FF.Empresa_Id       = FFxT.Empresa_Id" & vbCrLf &
                            "    AND FF.EndEmpresa_Id    = FFxT.EndEmpresa_Id" & vbCrLf &
                            "    AND FF.Conveniado_Id    = FFxT.Conveniado_Id" & vbCrLf &
                            "    AND FF.EndConveniado_Id = FFxT.EndConveniado_Id" & vbCrLf &
                            "    AND FF.Fatura_Id        = FFxT.Fatura_Id" & vbCrLf &
                            "  INNER JOIN NotasFiscais n" & vbCrLf &
                            "     ON N.Empresa_Id      = FFxI.Empresa_Id" & vbCrLf &
                            "    AND N.EndEmpresa_Id   = FFxI.EndEmpresa_Id" & vbCrLf &
                            "    AND N.Cliente_Id      = FFxI.Cliente_Id" & vbCrLf &
                            "    AND N.EndCliente_Id   = FFxI.EndCliente_Id" & vbCrLf &
                            "    AND N.EntradaSaida_Id = FFxI.EntradaSaida_Id" & vbCrLf &
                            "    AND N.Serie_Id        = FFxI.Serie_Id" & vbCrLf &
                            "    AND N.Nota_Id         = FFxI.Nota_Id" & vbCrLf &
                            "  INNER JOIN NotasXNotas nn " & vbCrLf &
                            " 	  ON N.Empresa_Id      = nn.Empresa_Id" & vbCrLf &
                            "    AND N.EndEmpresa_Id   = nn.EndEmpresa_Id" & vbCrLf &
                            "    AND N.Cliente_Id      = nn.Cliente_Id" & vbCrLf &
                            "    AND N.EndCliente_Id   = nn.EndCliente_Id" & vbCrLf &
                            "    AND N.EntradaSaida_Id = nn.EntradaSaida_Id" & vbCrLf &
                            "    AND N.Serie_Id        = nn.Serie_Id" & vbCrLf &
                            "    AND N.Nota_Id         = nn.Nota_Id" & vbCrLf &
                            "  INNER JOIN NotasFiscais ori" & vbCrLf &
                            "  	  ON ori.Empresa_Id      = nn.OrigemEmpresa_Id" & vbCrLf &
                            "    AND ori.EndEmpresa_Id   = nn.OrigemEndEmpresa_Id" & vbCrLf &
                            "    AND ori.Cliente_Id      = nn.OrigemCliente_Id" & vbCrLf &
                            "    AND ori.EndCliente_Id   = nn.OrigemEndCliente_Id" & vbCrLf &
                            "    AND ori.EntradaSaida_Id = nn.OrigemEntradaSaida_Id" & vbCrLf &
                            "    AND ori.Serie_Id        = nn.OrigemSerie_Id" & vbCrLf &
                            "    AND ori.Nota_Id         = nn.OrigemNota_Id" & vbCrLf &
                            "  INNER JOIN ContasAPagar cp " & vbCrLf &
                            "     ON cp.Registro_Id = FFxT.Titulo_Id" & vbCrLf &
                            "  INNER JOIN SubOperacoes sub1 " & vbCrLf &
                            "     ON sub1.Operacao_Id = n.Operacao" & vbCrLf &
                            "    AND sub1.SubOperacoes_Id = n.SubOperacao" & vbCrLf &
                            "  INNER JOIN SubOperacoes sub2 " & vbCrLf &
                            "     ON sub2.Operacao_Id = ori.Operacao" & vbCrLf &
                            "    AND sub2.SubOperacoes_Id = ori.SubOperacao" & vbCrLf &
                            "  INNER JOIN TipoDeDocumento tp" & vbCrLf &
                            "     ON n.TipoDeDocumento = tp.Codigo_Id" & vbCrLf &
                            "  INNER JOIN Provisoes p" & vbCrLf &
                            "     ON p.Provisao_Id = cp.Provisao" & vbCrLf &
                            "  WHERE 1=1" & vbCrLf &
                            "    AND FFxI.EmpresaPagadora_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                            "    AND FFxI.EndEmpresaPagadora_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
                            "    AND FFxI.Conveniado_Id = '" & CodigoConveniadoLan.Value.Split("-")(0) & "'" & vbCrLf &
                            "    AND FFxI.EndConveniado_Id = " & CodigoConveniadoLan.Value.Split("-")(1) & vbCrLf
        If Not String.IsNullOrWhiteSpace(txtFatura.Text) Then
            sql &= "    AND (FFxI.Fatura_Id = " & txtFatura.Text.Trim() & " OR FFxI.Nota_Id = " & txtFatura.Text.Trim() & ") " & vbCrLf
        End If
        sql &= "   ORDER BY FFxI.Fatura_Id"
        Return sql
    End Function

    Private Function ValidaImpressao() As Boolean
        If ddlEmpresa.SelectedIndex < 1 Then
            MsgBox(Me.Page, "Informe a Empresa para gerar o Relatório.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(CodigoConveniadoLan.Value) Then
            MsgBox(Me.Page, "Informe o Conveniado para gerar o Relatório.")
            Return False
        End If

        Return True
    End Function

    Private Sub ConsultarProg()
        If ValidaCamposProgFatura() Then
            CarregaProgFaturas()
        End If
    End Sub

    Private Sub Encerrar()
        RecuperaItemFF()

        Dim CodigoTitulo As Integer
        Dim Sinistro As Boolean = False


        Dim Sqls As New ArrayList
        Dim fatura As [Lib].Negocio.FaturaDeFrete = CType(Session("Fatura" & HID.Value), [Lib].Negocio.FaturaDeFrete)

        If fatura Is Nothing Then
            MsgBox(Me.Page, "É necessário selecionar uma fatura!")
            Exit Sub
        End If

        If String.IsNullOrWhiteSpace(txtVencimento.Text) Then
            MsgBox(Me.Page, "É necessário informar o campo data de vencimento do título!")
            Exit Sub
        End If

        Dim lst As [Lib].Negocio.ListFaturasDeFretesXItens = New [Lib].Negocio.ListFaturasDeFretesXItens(fatura)

        If Not FinanceiroNovo Then

            Dim t As [Lib].Negocio.Titulo = New [Lib].Negocio.Titulo(fatura.ListTituloFatura.Where(Function(s) s.Titulo.CodigoSituacao = eSituacao.Normal)(0).CodigoTitulo, IIf(rdbReceberLancamento.Checked, "R", "P"))

            If t.Codigo = 0 Then
                MsgBox(Me.Page, "Não foi possivel encontrar o título!")
                Exit Sub
            End If

            If t IsNot Nothing Then
                'Caso exista algum título parcial que esteja na mesma fatura o mesmo será cancelado 
                'já que liberando o título principal será atualizado com o valor total normal. 

                If fatura.ListTituloFatura.Where(Function(s) s.Titulo.CodigoSituacao = eSituacao.Normal).Count > 1 Then
                    For i As Integer = 1 To fatura.ListTituloFatura.Count - 1
                        If fatura.ListTituloFatura(i).Titulo.CodigoProvisao <> eProvisao.Baixa Then
                            fatura.ListTituloFatura(i).IUD = "D"
                            fatura.ListTituloFatura(i).Salvar()
                        Else
                            MsgBox(Me.Page, "O título parcial " & fatura.ListTituloFatura(i).Titulo.Codigo & " vinculado à esta fatura está baixado, verifique!")
                            Exit Sub
                        End If
                    Next
                End If

                t.IUD = "U"
                t.CodigoProvisao = eProvisao.Previsao
                t.CodigoMoeda = 1
                t.IndiceTitulo = 0
                t.IndiceFixo = True
                t.CodigoIndexador = 3
                t.ReceberPagar = IIf(rdbPagarLancamento.Checked, "P", "R")
                t.Prorrogacao = CDate(txtVencimento.Text)
                t.Movimento = CDate(txtVencimento.Text)

                t.MoedaValorDoDocumento = 0
                t.MoedaJuros = 0
                t.MoedaAcrescimos = 0
                t.MoedaDescontos = 0
                t.MoedaDeducoes = 0

                If (fatura.Itens.All(Function(s) s.CodigoEncargo = "LIQUIDOAPAGAR")) Then
                    If fatura.Itens(0).Nota.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao Then
                        Dim vlrTarifaSeguro As Decimal = ItemFF.Nota.Itens(0).Encargos.Where(Function(s) s.Codigo.ToUpper().Trim() = "TARIFA SEGURO").Sum(Function(s) s.Valor)
                        t.ValorDoDocumento = CDec(txtValorFrete.Text) - IIf(String.IsNullOrWhiteSpace(TxtValorAdiantamento.Text), Decimal.Zero, CDec(TxtValorAdiantamento.Text)) - vlrTarifaSeguro
                        t.MoedaValorDoDocumento = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(t.ValorDoDocumento, 3, CDate(txtVencimento.Text)), 2, MidpointRounding.AwayFromZero)
                    Else
                        t.ValorDoDocumento = CDec(txtValorFrete.Text) - IIf(String.IsNullOrWhiteSpace(TxtValorAdiantamento.Text), Decimal.Zero, CDec(TxtValorAdiantamento.Text))
                        t.MoedaValorDoDocumento = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(t.ValorDoDocumento, 3, CDate(txtVencimento.Text)), 2, MidpointRounding.AwayFromZero)
                    End If
                ElseIf (fatura.Itens.All(Function(s) s.CodigoEncargo = "BAIXAADTO")) AndAlso TxtValorAdiantamento.Text.Length > 0 Then
                    t.ValorDoDocumento = CDec(TxtValorAdiantamento.Text)
                    t.MoedaValorDoDocumento = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(t.ValorDoDocumento, 3, CDate(txtVencimento.Text)), 2, MidpointRounding.AwayFromZero)
                End If

                If (fatura.Itens.All(Function(s) s.CodigoEncargo = "LIQUIDOAPAGAR")) Then
                    Dim vlrSaida As Decimal = Decimal.Zero

                    Dim pesoDeChegada As Decimal = Decimal.Zero
                    If Not String.IsNullOrWhiteSpace(txtPesoChegada.Text) Then pesoDeChegada = CDec(txtPesoChegada.Text)
                    fatura.Itens(0).PesoDeChegada = pesoDeChegada
                    fatura.Itens(0).IUD = "U"
                    fatura.Itens(0).Salvar()

                    If Not String.IsNullOrWhiteSpace(txtValorFrete.Text) Then vlrSaida = CDec(txtValorFrete.Text)

                    Dim vlrChegada As Decimal = Decimal.Zero
                    If Not String.IsNullOrWhiteSpace(txtFreteChegada.Text) Then
                        vlrChegada = CDec(txtFreteChegada.Text)
                    End If

                    t.Deducoes = 0
                    t.MoedaDeducoes = 0
                    t.Acrescimos = 0
                    t.MoedaAcrescimos = 0

                    'Sinistro = fatura.Itens(0).Nota.NotasOrigemDestino.Any(Function(s) s.PesoDeChegada.Sinistro)

                    '#Complemento
                    If CDec(txtValorCteComplementar.Text) > 0 Then
                        If vlrChegada > 0 AndAlso (CDec(txtValorTotalFrete.Text) - vlrChegada) > 0 Then  'QUEBRA

                            'SE A DIF DO PESO DE SAIDA MENOS O PESO DE CHEGADA MAIOR QUE TOLERANCIA
                            If (CDec(TxtPeso.Text) - CDec(txtPesoChegada.Text)) > CDec(txtTolerancia.Text) Then
                                Dim vlrLiquido As NotaFiscalXItemXEncargo = ItemFF.Nota.Itens(0).Encargos.Where(Function(s) s.Codigo.ToUpper().Trim().Contains("LIQUIDOAPAGAR")).FirstOrDefault()
                                If vlrLiquido IsNot Nothing Then
                                    t.Deducoes = t.ValorDoDocumento - Decimal.Round(vlrLiquido.Valor, 2)
                                End If
                            Else
                                t.Deducoes = Math.Abs(CDec(txtValorTotalFrete.Text) - vlrChegada)
                            End If

                        ElseIf vlrChegada > 0 AndAlso (vlrChegada - CDec(txtValorTotalFrete.Text)) > 0 Then 'SOBRA
                            t.Acrescimos = vlrChegada - CDec(txtValorTotalFrete.Text)
                        End If

                    ElseIf vlrChegada > 0 AndAlso (vlrSaida - vlrChegada) > 0 Then  'QUEBRA
                        'SE A DIF DO PESO DE SAIDA MENOS O PESO DE CHEGADA MAIOR QUE TOLERANCIA
                        If (CDec(TxtPeso.Text) - CDec(txtPesoChegada.Text)) > CDec(txtTolerancia.Text) Then
                            Dim vlrLiquido As NotaFiscalXItemXEncargo = ItemFF.Nota.Itens(0).Encargos.Where(Function(s) s.Codigo.ToUpper().Trim().Contains("LIQUIDOAPAGAR")).FirstOrDefault()
                            If vlrLiquido IsNot Nothing Then
                                t.Deducoes = t.ValorDoDocumento - Decimal.Round(vlrLiquido.Valor, 2)
                            End If
                        Else
                            t.Deducoes = vlrSaida - vlrChegada
                        End If
                    ElseIf vlrChegada > 0 AndAlso (vlrChegada - vlrSaida) > 0 AndAlso Not CDec(TxtPeso.Text) = CDec(txtPesoChegada.Text) Then 'SOBRA
                        t.Acrescimos = vlrChegada - vlrSaida
                    End If

                End If

                If lst.All(Function(s) s.CodigoEncargo.Trim().Equals("BAIXAADTO")) Then
                    Dim nf As String = ""
                    If lst.FirstOrDefault() IsNot Nothing Then
                        nf = lst(0).NumeroNota
                    End If
                    t.Historico = "PGTO ADIANTAMENTO CONFORME FATURA NÚMERO " & fatura.CodigoFatura & " - CTRC NÚMERO " & nf
                ElseIf lst.All(Function(s) s.CodigoEncargo.Trim().Equals("LIQUIDOAPAGAR")) Then
                    Dim nf As String = ""
                    If lst.FirstOrDefault() IsNot Nothing Then
                        nf = lst(0).NumeroNota
                    End If
                    t.Historico = "PGTO SALDO CONFORME FATURA NÚMERO " & fatura.CodigoFatura & " - CTRC NÚMERO " & nf
                End If

                t.SalvarSql(Sqls, False)

                CodigoTitulo = t.Codigo

            End If
        Else
            Dim t As [Lib].Negocio.Novo.TituloNovo = New [Lib].Negocio.Novo.TituloNovo(fatura.ListTituloFatura.Where(Function(s) s.TituloNovo.CodigoSituacao = eSituacao.Normal)(0).CodigoTitulo)

            If t IsNot Nothing Then
                t.IUD = "U"
                t.CodigoProvisao = eProvisao.Provisao
                t.CodigoIndexador = 3

                If fatura.Itens(0).EntradaSaida = "E" Then
                    t.ReceberPagar = "P"
                Else
                    t.ReceberPagar = "R"
                End If

                t.Reprogramacao = CDate(txtVencimento.Text)
                t.Movimento = CDate(txtVencimento.Text)
                t.CodigoMoeda = 1

                If (fatura.Itens.All(Function(s) s.CodigoEncargo = "LIQUIDOAPAGAR")) Then
                    'CONTA - VALOR DO DOCUMENTO
                    t.CodigoContaContabilCliFor = ItemFF.Nota.SubOperacao.CodigoGrupoContas
                    'CONTA - VALOR LÍQUIDO
                    t.CodigoContaContabilRecPag = t.Empresa.Empresa.CodigoContaGrupoBanco
                    'VALOR DO DOCUMENTO
                    If fatura.Itens(0).Nota.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao Then
                        Dim vlrTarifaSeguro As Decimal = ItemFF.Nota.Itens(0).Encargos.Where(Function(s) s.Codigo.ToUpper().Trim() = "TARIFA SEGURO").Sum(Function(s) s.Valor)
                        t.Valores.EncargoValorDocumento.ValorOficial = CDec(txtValorFrete.Text) - IIf(String.IsNullOrWhiteSpace(TxtValorAdiantamento.Text), Decimal.Zero, CDec(TxtValorAdiantamento.Text)) - vlrTarifaSeguro
                        t.Valores.EncargoValorDocumento.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(t.Valores.EncargoValorDocumento.ValorOficial, 3, CDate(txtVencimento.Text))
                    Else
                        Dim vlrLiquido As Decimal = ItemFF.Nota.Itens(0).Encargos.Where(Function(s) s.Codigo.ToUpper().Trim() = "LIQUIDOAPAGAR").Sum(Function(s) s.Valor)
                        t.Valores.EncargoValorDocumento.ValorOficial = CDec(txtValorFrete.Text)
                        t.Valores.EncargoValorDocumento.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(t.Valores.EncargoValorDocumento.ValorOficial, 3, CDate(txtVencimento.Text))
                    End If
                    'VALOR LÍQUIDO
                    t.Valores.AtualizaLiquido()
                ElseIf (fatura.Itens.All(Function(s) s.CodigoEncargo = "BAIXAADTO")) Then
                    'CONTA - VALOR DO DOCUMENTO
                    t.CodigoContaContabilCliFor = ItemFF.Nota.SubOperacao.CodigoGrupoContas
                    'CONTA - VALOR LÍQUIDO
                    t.CodigoContaContabilRecPag = t.Empresa.Empresa.CodigoContaGrupoBanco
                    'VALOR DO DOCUMENTO
                    t.Valores.EncargoValorDocumento.ValorOficial = CDec(TxtValorAdiantamento.Text)
                    t.Valores.EncargoValorDocumento.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(t.Valores.EncargoValorDocumento.ValorOficial, 3, CDate(txtVencimento.Text))
                    'VALOR LÍQUIDO
                    t.Valores.AtualizaLiquido()
                End If

                If (fatura.Itens.All(Function(s) s.CodigoEncargo = "LIQUIDOAPAGAR")) Then
                    Dim pesoDeChegada As Decimal = Decimal.Zero
                    If Not String.IsNullOrWhiteSpace(txtPesoChegada.Text) Then pesoDeChegada = CDec(txtPesoChegada.Text)
                    fatura.Itens(0).PesoDeChegada = pesoDeChegada
                    fatura.Itens(0).IUD = "U"
                    fatura.Itens(0).Salvar()

                    Dim vlrSaida As Decimal = Decimal.Zero
                    If Not String.IsNullOrWhiteSpace(txtValorFrete.Text) Then
                        vlrSaida = CDec(txtValorFrete.Text)
                    End If

                    Dim vlrChegada As Decimal = Decimal.Zero
                    If Not String.IsNullOrWhiteSpace(txtFreteChegada.Text) Then
                        vlrChegada = CDec(txtFreteChegada.Text)
                    End If
                    '
                    If (vlrSaida - vlrChegada) > 0 Then
                        'VALOR DE DEDUÇÃO 
                        Dim enc As New [Lib].Negocio.Novo.TituloXContaContabil(t)
                        enc.CodigoContaEncargo = t.Empresa.Empresa.CodigoContaAdiantamentoDeFrete
                        enc.ValorOficial = vlrSaida - vlrChegada
                        enc.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(enc.ValorOficial, 3, CDate(txtVencimento.Text))
                        t.Valores.Add(enc)
                        t.Valores.AtualizaLiquido()
                    Else
                        'VALOR DE ACRÉSCIMO 
                        Dim enc As New [Lib].Negocio.Novo.TituloXContaContabil(t)
                        enc.CodigoContaEncargo = t.Empresa.Empresa.CodigoContaPedagioDeFrete
                        enc.ValorOficial = vlrChegada - vlrSaida
                        enc.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(enc.ValorOficial, 3, CDate(txtVencimento.Text))
                        t.Valores.Add(enc)
                        t.Valores.AtualizaLiquido()
                    End If
                End If

                If lst.All(Function(s) s.CodigoEncargo.Trim().Equals("BAIXAADTO")) Then
                    Dim nf As String = ""
                    If lst.FirstOrDefault() IsNot Nothing Then
                        nf = lst(0).NumeroNota
                    End If
                    t.Historico = "PGTO " & "ADIANTAMENTO CONFORME FATURA NÚMERO " & fatura.CodigoFatura & " - CTRC NÚMERO " & nf
                ElseIf lst.All(Function(s) s.CodigoEncargo.Trim().Equals("LIQUIDOAPAGAR")) Then
                    Dim nf As String = ""
                    If lst.FirstOrDefault() IsNot Nothing Then
                        nf = lst(0).NumeroNota
                    End If
                    t.Historico = "PGTO SALDO CONFORME FATURA NÚMERO " & fatura.CodigoFatura & " - CTRC NÚMERO " & nf
                End If

                fatura.ListTituloFatura.SalvarSql(Sqls)
                CodigoTitulo = t.Codigo
            End If
        End If

        If ItemFF.Nota IsNot Nothing AndAlso ItemFF.Nota.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao Then
            Dim quebra = ItemFF.Nota.Itens(0).Encargos.Where(Function(s) s.Codigo = "QUEBRAS - P. CH").FirstOrDefault()
            Dim sobra = ItemFF.Nota.Itens(0).Encargos.Where(Function(s) s.Codigo = "SOBRAS - P. CH").FirstOrDefault()

            Dim objCirculacao As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal(ItemFF.Nota)
            If objCirculacao IsNot Nothing AndAlso objCirculacao.NotasXNotas IsNot Nothing Then
                Dim lstNotas As New [Lib].Negocio.ListNotasFiscais()
                Dim lstNotasXNotas As New [Lib].Negocio.ListNotasXNotas(objCirculacao, True, True)
                For Each item As [Lib].Negocio.NotasXNotas In lstNotasXNotas
                    Dim nf As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                    nf.CodigoEmpresa = item.EmpresaCnpj
                    nf.EnderecoEmpresa = item.EndEmpresa
                    nf.CodigoCliente = item.ClienteCnpj
                    nf.EnderecoCliente = item.EndCliente
                    nf.EntradaSaida = IIf(item.EntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                    nf.Codigo = item.NumeroNota
                    nf.Serie = item.Serie
                    nf = New [Lib].Negocio.NotaFiscal(nf)
                    lstNotas.Add(nf)
                Next

                If lstNotas IsNot Nothing AndAlso lstNotas.Count > 0 Then
                    Dim objComprovacao As [Lib].Negocio.NotaFiscal = lstNotas.Where(Function(s) s.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Comprovacao).FirstOrDefault()
                    If objComprovacao IsNot Nothing Then
                        Dim vlrConvertido As Decimal = 0
                        If quebra IsNot Nothing AndAlso quebra.Valor > 0 Then
                            If objComprovacao.CodigoPedido > 0 Then
                                vlrConvertido = Funcoes.ConverteParaMoedaExtrangeira(quebra.Valor, 2, objComprovacao.DataNota, True, False, 2)
                            Else
                                vlrConvertido = Funcoes.ConverteMoeda(quebra.Valor, objComprovacao.Pedido.IndiceFixado, eTiposMoeda.MoedaEstrangeira, True, False, 2)
                            End If
                            ContabilizarCTe(objComprovacao, quebra, "D", quebra.Valor, vlrConvertido, Sqls, 1)
                        ElseIf sobra IsNot Nothing AndAlso sobra.Valor > 0 Then
                            If objComprovacao.CodigoPedido > 0 Then
                                vlrConvertido = Funcoes.ConverteParaMoedaExtrangeira(sobra.Valor, 2, objComprovacao.DataNota, True, False, 2)
                            Else
                                vlrConvertido = Funcoes.ConverteMoeda(sobra.Valor, objComprovacao.Pedido.IndiceFixado, eTiposMoeda.MoedaEstrangeira, True, False, 2)
                            End If
                            ContabilizarCTe(objComprovacao, sobra, "C", sobra.Valor, vlrConvertido, Sqls, 1)
                        End If
                    End If
                End If
            End If
        End If

        If Not Banco.GravaBanco(Sqls) Then
            MsgBox(Me.Page, Session("ssMessage"))
            Exit Sub
        Else
            MsgBox(Me.Page, "Título " & CodigoTitulo & "  liberado com Sucesso.", eTitulo.Sucess)
        End If
        Limpar(False)
    End Sub

    Private Sub CarregaProgFaturas(Optional ByVal Consulta As Boolean = False)
        Dim Fatura As New [Lib].Negocio.FaturaDeFrete()
        If DdlEmpresaProg.SelectedIndex > 0 Then
            EmpresaProg = DdlEmpresaProg.SelectedValue.ToString.Split("-")
            Fatura.CodigoEmpresa = EmpresaProg(0)
            Fatura.EnderecoEmpresa = EmpresaProg(1)
        End If

        If CodigoConveniadoProg.Value.ToString.Length > 0 Then
            ConveniadoProg = CodigoConveniadoProg.Value.ToString.Split("-")
            Fatura.CodigoConveniado = ConveniadoProg(0)
            Fatura.EnderecoConveniado = ConveniadoProg(1)
        End If

        If Not String.IsNullOrWhiteSpace(txtNumeroFatura.Text) AndAlso CInt(txtNumeroFatura.Text) > 0 Then
            Fatura.CodigoFatura = txtNumeroFatura.Text
        End If

        Dim tipo As String = String.Empty
        If rdAtivos.Checked Then
            tipo = "3"
        ElseIf rdLiberados.Checked Then
            tipo = "2"
        ElseIf rdBaixados.Checked Then
            tipo = "1"
        End If

        Dim clienteFatura As Cliente
        If CodigoClienteProg.Value.ToString.Length > 0 Then
            ClienteProg = CodigoClienteProg.Value.ToString.Split("-")
            clienteFatura = New Cliente(ClienteProg(0), ClienteProg(1))
        End If

        Dim Faturas As New [Lib].Negocio.ListFaturaDeFrete(Fatura, chkUsarPeriodo.Checked, clienteFatura, txtVencimentoDe.Text, txtVencimentoAte.Text, tipo)

        If Faturas.Count > 0 Then
            Session("ssFaturasProg" & HID.Value) = Faturas
            DgProgFaturas.DataSource = From f In Faturas.SelectMany(Function(s) s.ListTituloFatura)
                                       Where f.Titulo.Provisao.Codigo = tipo
                                       Select f.FaturaDeFrete.CodigoFatura, f.FaturaDeFrete.CodigoConveniado, f.FaturaDeFrete.EnderecoConveniado,
                                              f.FaturaDeFrete.ConveniadoNome, f.FaturaDeFrete.Movimento, f.FaturaDeFrete.ValorDaFatura, f.FaturaDeFrete.ValorLancadoFatura,
                                              f.FaturaDeFrete.ValorSaldoFatura, f.Titulo
            DgProgFaturas.DataBind()
        Else
            DgProgFaturas.DataSource = Nothing
            DgProgFaturas.DataBind()
            MsgBox(Me.Page, "Registro(s) não encontrado(s) para esta seleção/período!")
            Session.Remove("ssFaturasProg" & HID.Value)
        End If
    End Sub

    Protected Sub DgProgFaturas_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles DgProgFaturas.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                If e.Row.DataItem IsNot Nothing Then
                    Dim imgDelete As ImageButton = CType(e.Row.FindControl("imgDelete"), ImageButton)
                    imgDelete.Visible = rdLiberados.Checked
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgDelete_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgDelete As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgDelete.NamingContainer, GridViewRow)

            Dim Sqls As New ArrayList
            Dim t As New [Lib].Negocio.Titulo(CType(row.FindControl("lblTitulo"), Label).Text())
            t.CodigoProvisao = 3
            t.ValorDoDocumento = CDec(DgProgFaturas.Rows(row.RowIndex).Cells(7).Text)
            t.Deducoes = 0
            t.Descontos = 0
            t.Acrescimos = 0
            t.Juros = 0
            t.MoedaValorDoDocumento = 0
            t.MoedaDeducoes = 0
            t.MoedaDescontos =
            t.MoedaAcrescimos = 0
            t.Juros = 0
            t.IUD = "U"
            t.UsuarioAlteracao = HttpContext.Current.Session("ssNomeUsuario")
            t.UsuarioAlteracaoData = DateTime.Now
            t.SalvarSql(Sqls, False)

            If Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, "Título bloqueado com Sucesso.", eTitulo.Sucess)

                ConsultarProg()
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaCamposLanFatura() As Boolean
        EmpresaPF = txtCodigoEmpresa.Value.Split("-")
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa pagadora é obrigatória!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtConveniadoLan.Text) AndAlso (String.IsNullOrWhiteSpace(txtFatura.Text) OrElse CInt(txtFatura.Text) = 0) Then
            MsgBox(Me.Page, "Conveniado ou Número do Fatura/CTRC é obrigatório!")
            Return False
        Else
            Return True
        End If
    End Function

    Private Function ValidaCamposDadosLanFatura() As Boolean
        EmpresaPF = txtCodigoEmpresa.Value.ToString.Split("-")
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa pagadora obrigatória!")
            Return False
        ElseIf Trim(txtConveniadoLan.Text) = "" Then
            MsgBox(Me.Page, "Conveniado obrigatório!")
            Return False
        ElseIf txtFatura.Text = 0 Or txtFatura.Text.Trim = "" Then
            MsgBox(Me.Page, "Numero da fatura obrigatório!")
            txtFatura.Focus()
            Return False
        ElseIf txtValorFatura.Text = 0 Then
            MsgBox(Me.Page, "Valor da fatura obrigatório!")
            txtFatura.Focus()
            Return False
        ElseIf RbRecibo.Checked = False And RbCtrc.Checked = False Then
            MsgBox(Me.Page, "Obrigatório informar tipo de documento (CTRC ou recibo de frete)!")
            Return False
        ElseIf optAdiantamento.Checked = False And optAmortizacao.Checked = False And optCartaFrete.Checked = False Then
            MsgBox(Me.Page, "Obrigatório informar via (adiantamento, saldo ou amortização)!")
            Return False
        End If
        Return True
    End Function

    Private Sub CarregaNotaDeFrete()

        'Dim objPesoDeChegada As New [Lib].Negocio.ListNotaFiscalXDestino()
        Dim NfDeFrete As New [Lib].Negocio.NotaFiscal(CType(Session("NotasDeFrete" & HID.Value), [Lib].Negocio.NotaFiscal))

        If NfDeFrete IsNot Nothing AndAlso Not NfDeFrete.NossaEmissao AndAlso optCartaFrete.Checked Then
            If NfDeFrete IsNot Nothing AndAlso NfDeFrete.NotasXNotas IsNot Nothing Then
                Dim nf As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                nf.CodigoEmpresa = NfDeFrete.NotasXNotas.OrigemEmpresaCnpj
                nf.EnderecoEmpresa = NfDeFrete.NotasXNotas.OrigemEndEmpresa
                nf.CodigoCliente = NfDeFrete.NotasXNotas.OrigemClienteCnpj
                nf.EnderecoCliente = NfDeFrete.NotasXNotas.OrigemEndCliente
                nf.EntradaSaida = IIf(NfDeFrete.NotasXNotas.OrigemEntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                nf.Codigo = NfDeFrete.NotasXNotas.OrigemNota
                nf.Serie = NfDeFrete.NotasXNotas.OrigemSerie
                nf = New [Lib].Negocio.NotaFiscal(nf)

                If (NfDeFrete.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.Estadia) AndAlso NfDeFrete.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.ComplementoDeFrete)) AndAlso nf.CodigoTipoDeDocumento = eTipoDeDocumento.CTRC Then
                    MsgBox(Me.Page, "Não é possível realizar o lançamento do saldo, pois já existe um conhecimento de transporte vinculado!")
                    Exit Sub
                ElseIf nf.CodigoTipoDeDocumento = eTipoDeDocumento.Nota Then
                    Dim origem As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                    origem.CodigoEmpresa = NfDeFrete.NotasXNotas.OrigemEmpresaCnpj
                    origem.EnderecoEmpresa = NfDeFrete.NotasXNotas.OrigemEndEmpresa
                    origem.CodigoCliente = NfDeFrete.NotasXNotas.OrigemClienteCnpj
                    origem.EnderecoCliente = NfDeFrete.NotasXNotas.OrigemEndCliente
                    origem.EntradaSaida = IIf(NfDeFrete.NotasXNotas.OrigemEntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                    origem.Codigo = NfDeFrete.NotasXNotas.OrigemNota
                    origem.Serie = NfDeFrete.NotasXNotas.OrigemSerie
                    origem = New [Lib].Negocio.NotaFiscal(origem)
                    If (NfDeFrete.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.Estadia) AndAlso NfDeFrete.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.ComplementoDeFrete)) AndAlso nf.CodigoTipoDeDocumento = eTipoDeDocumento.CTRC AndAlso NfDeFrete.Codigo <> nf.Codigo Then
                        MsgBox(Me.Page, "Não é possível realizar o lançamento do saldo, pois já existe um conhecimento de transporte vinculado!")
                        Exit Sub
                    End If

                End If

                Dim lstNotasXNotas As New [Lib].Negocio.ListNotasXNotas(NfDeFrete)
                Dim objNota As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                objNota.CodigoEmpresa = lstNotasXNotas(0).OrigemEmpresaCnpj
                objNota.EnderecoEmpresa = lstNotasXNotas(0).OrigemEndEmpresa
                objNota.CodigoCliente = lstNotasXNotas(0).OrigemClienteCnpj
                objNota.EnderecoCliente = lstNotasXNotas(0).OrigemEndCliente
                objNota.EntradaSaida = IIf(lstNotasXNotas(0).OrigemEntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                objNota.Codigo = lstNotasXNotas(0).OrigemNota
                objNota.Serie = lstNotasXNotas(0).OrigemSerie
                objNota = New [Lib].Negocio.NotaFiscal(objNota)
                If NfDeFrete.CodigoTipoDeDocumentoFrete <> eTipoDeDocumentoFrete.Estadia AndAlso NfDeFrete.CodigoTipoDeDocumentoFrete <> eTipoDeDocumentoFrete.Complemento Then
                    If NfDeFrete.PesoDeChegada Is Nothing Then
                        MsgBox(Me.Page, String.Format("É necessário existir um lançamento no peso de chegada para o CTRC {0}-{1} e NF {2}-{3}!", NfDeFrete.Codigo, NfDeFrete.Serie, objNota.Codigo, objNota.Serie))
                        Exit Sub
                    End If
                End If
            End If
        Else
            If NfDeFrete IsNot Nothing AndAlso NfDeFrete.NossaEmissao AndAlso optCartaFrete.Checked Then
                Dim objCirculacao As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal(NfDeFrete)
                If objCirculacao IsNot Nothing AndAlso objCirculacao.NotasXNotas IsNot Nothing Then
                    Dim lstNotas As New [Lib].Negocio.ListNotasFiscais()
                    Dim lst As New [Lib].Negocio.ListNotasXNotas(objCirculacao, True, True)
                    For Each item As [Lib].Negocio.NotasXNotas In lst
                        Dim nf As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                        nf.CodigoEmpresa = objCirculacao.NotasXNotas.EmpresaCnpj
                        nf.EnderecoEmpresa = objCirculacao.NotasXNotas.EndEmpresa
                        nf.CodigoCliente = objCirculacao.NotasXNotas.ClienteCnpj
                        nf.EnderecoCliente = objCirculacao.NotasXNotas.EndCliente
                        nf.EntradaSaida = IIf(objCirculacao.NotasXNotas.EntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                        nf.Codigo = objCirculacao.NotasXNotas.NumeroNota
                        nf.Serie = objCirculacao.NotasXNotas.Serie
                        nf = New [Lib].Negocio.NotaFiscal(nf)
                        lstNotas.Add(nf)
                    Next

                    If lstNotas Is Nothing OrElse Not lstNotas.Count > 0 OrElse Not lstNotas.Any(Function(s) s.CodigoTipoDeDocumento = CInt(eTipoDeDocumentoFrete.Comprovacao)) Then
                        MsgBox(Me.Page, "É necessário possuir um conhecimento de comprovação para realizar o lançamento de saldo!")
                        Exit Sub
                    End If
                End If

                Dim lstNotasXNotas As New [Lib].Negocio.ListNotasXNotas(objCirculacao)
                Dim objNota As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                objNota.CodigoEmpresa = lstNotasXNotas(0).OrigemEmpresaCnpj
                objNota.EnderecoEmpresa = lstNotasXNotas(0).OrigemEndEmpresa
                objNota.CodigoCliente = lstNotasXNotas(0).OrigemClienteCnpj
                objNota.EnderecoCliente = lstNotasXNotas(0).OrigemEndCliente
                objNota.EntradaSaida = IIf(lstNotasXNotas(0).OrigemEntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                objNota.Codigo = lstNotasXNotas(0).OrigemNota
                objNota.Serie = lstNotasXNotas(0).OrigemSerie
                objNota = New [Lib].Negocio.NotaFiscal(objNota)
                If NfDeFrete.CodigoTipoDeDocumentoFrete <> eTipoDeDocumentoFrete.Estadia AndAlso NfDeFrete.CodigoTipoDeDocumentoFrete <> eTipoDeDocumentoFrete.Complemento Then
                    If NfDeFrete.PesoDeChegada Is Nothing Then
                        MsgBox(Me.Page, String.Format("É necessário existir um lançamento no peso de chegada para o CTRC {0}-{1} e NF {2}-{3}!", NfDeFrete.Codigo, NfDeFrete.Serie, objNota.Codigo, objNota.Serie))
                        Exit Sub
                    End If
                End If
            End If
        End If

        Dim objFatura As [Lib].Negocio.FaturaDeFrete = Session("Fatura" & HID.Value)
        If objFatura Is Nothing Then
            MsgBox(Me.Page, "É necessário selecionar uma fatura de frete!")
            Exit Sub
        End If

        ItemFF = New [Lib].Negocio.FaturaDeFreteXItens(objFatura)
        ItemFF.EmpresaCnpj = NfDeFrete.CodigoEmpresa
        ItemFF.EmpresaEnd = NfDeFrete.EnderecoEmpresa
        ItemFF.ClienteCnpj = NfDeFrete.CodigoCliente
        ItemFF.ClienteEnd = NfDeFrete.EnderecoCliente
        ItemFF.Serie = NfDeFrete.Serie
        ItemFF.NumeroNota = NfDeFrete.Codigo
        ItemFF.Peso = NfDeFrete.PesoBruto
        ItemFF.FreteCombinado = NfDeFrete.Itens(0).Unitario
        ItemFF.ValorFrete = NfDeFrete.Itens(0).ValorTotal
        ItemFF.ViaAdiantamento = optAdiantamento.Checked
        ItemFF.ViaAmortizacao = optAmortizacao.Checked
        ItemFF.ViaCartaFrete = optCartaFrete.Checked
        ItemFF.EntradaSaida = IIf(NfDeFrete.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
        ItemFF.CTRC = RbCtrc.Checked
        ItemFF.ReciboFrete = RbRecibo.Checked

        If optCartaFrete.Checked Then
            If NfDeFrete.PesoDeChegada IsNot Nothing Then
                ItemFF.PesoDeChegada = NfDeFrete.PesoDeChegada.PesoBruto
            End If
        ElseIf optAdiantamento.Checked OrElse optAmortizacao.Checked Then
            ItemFF.PesoDeChegada = IIf(ItemFF.EntradaSaida = "E" And NfDeFrete.NotasXNotas.PesoBrutoRomaneio > 0, NfDeFrete.NotasXNotas.PesoBrutoRomaneio, NfDeFrete.NotasXNotas.PesoFiscal)
        End If
        SalvarItemFF()

        If objFatura.LancamentoManual = 1 Then
            CarregarCTE(ItemFF, False)
        Else
            CarregarFormItemFF()
        End If

    End Sub

    Private Sub CarregarCTE(ByVal ItemFaturaItens As [Lib].Negocio.FaturaDeFreteXItens, bConsulta As Boolean)

        Dim objFatura As [Lib].Negocio.FaturaDeFrete = Session("Fatura" & HID.Value)
        ItemFF = ItemFaturaItens
        Session("NotasDeFrete" & HID.Value) = ItemFF.Nota

        Dim campo() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
        Dim objEmpresa = New ClienteXEmpresa(campo(0), campo(1))

        If ItemFF.NumeroNota > 0 Then
            Dim objTransportadora As New [Lib].Negocio.Cliente(ItemFF.ClienteCnpj, ItemFF.ClienteEnd)
            ddlFilial.SelectedValue = ItemFF.EmpresaCnpj & "-" & ItemFF.EmpresaEnd
            txtTransportador.Text = objTransportadora.Nome
            TxtNrFrete.Text = ItemFF.NumeroNota
            TxtPeso.Text = String.Format("{0:N0}", ItemFF.Peso)
            txtFreteCombinado.Text = String.Format("{0:N2}", ItemFF.FreteCombinado)
            txtValorFrete.Text = String.Format("{0:N2}", ItemFF.ValorFrete)

            txtValorKg.Text = String.Format("{0:N4}", Decimal.Zero)
            txtTolerancia.Text = String.Format("{0:N2}", Decimal.Zero)
            txtPesoChegada.Text = String.Format("{0:N0}", Decimal.Zero)
            txtFreteChegada.Text = String.Format("{0:N2}", Decimal.Zero)
            txtDiferenca.Text = String.Format("{0:N2}", Decimal.Zero)

            Dim objAdto As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
            ItemFF.Nota.CarregandoNota = True
            Dim lstEncargos As New [Lib].Negocio.ListNotaFiscalXItemXEncargo(ItemFF.Nota.Itens(0), New List(Of eEtapaEncago) From {eEtapaEncago.Normal, eEtapaEncago.Adiantamento})
            ItemFF.Nota.CarregandoNota = False
            For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In lstEncargos.Where(Function(s) s.Existe AndAlso s.Valor > 0)
                If enc.Codigo.Trim().Contains("ADTODEFRETE") OrElse enc.Codigo.Trim().Contains("ADIANTAMENTO") Then
                    objAdto = enc
                    objAdto.Percentual = (CDec(enc.Valor) * CDec(100)) / CDec(txtValorFrete.Text)
                    Exit For
                End If
            Next

            If objAdto IsNot Nothing Then
                txtPercAdiantamento.Text = String.Format("{0:N2}%", objAdto.Percentual)
                TxtValorAdiantamento.Text = String.Format("{0:N2}", objAdto.Valor)
            Else
                txtPercAdiantamento.Text = "0%"
                TxtValorAdiantamento.Text = 0
            End If

            Dim objNotaDeFrete As [Lib].Negocio.NotaFiscal = Session("NotasDeFrete" & HID.Value)
            Dim objNotaFiscal As New [Lib].Negocio.NotaFiscal(objNotaDeFrete)
            Dim lstNotasOrigem As New [Lib].Negocio.ListNotasFiscais()
            Dim TemComplemento As Boolean = False

            Dim PesoDeChegada As Integer = 0
            Dim Sinistro As Boolean = False

            Dim ComplementoPeso As Decimal = 0
            Dim ComplementoValor As Decimal = 0
            Dim NFxMultiplosCTEs As Boolean = False

            If objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.NotasXNotas IsNot Nothing Then
                Dim lstNotasXNotas As New [Lib].Negocio.ListNotasXNotas(objNotaFiscal)

                For Each objNotasXNotas As [Lib].Negocio.NotasXNotas In lstNotasXNotas
                    Dim origemNota As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                    origemNota.CodigoEmpresa = objNotasXNotas.OrigemEmpresaCnpj
                    origemNota.EnderecoEmpresa = objNotasXNotas.OrigemEndEmpresa
                    origemNota.CodigoCliente = objNotasXNotas.OrigemClienteCnpj
                    origemNota.EnderecoCliente = objNotasXNotas.OrigemEndCliente
                    origemNota.EntradaSaida = IIf(objNotasXNotas.OrigemEntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                    origemNota.Codigo = objNotasXNotas.OrigemNota
                    origemNota.Serie = objNotasXNotas.OrigemSerie
                    origemNota = New [Lib].Negocio.NotaFiscal(origemNota)
                    lstNotasOrigem.Add(origemNota)
                Next

                If lstNotasOrigem.Count = 1 Then

                    'Busca as Notas vinculadas
                    Dim NotasDeOrigem As New [Lib].Negocio.ListNotasXNotas(lstNotasOrigem(0), True, True)
                    'Quantidade de ctes vinculados a essa nota
                    Dim QtdCtesVinculados As Integer = NotasDeOrigem.Where(Function(s) CType(s.NotaFiscal.CodigoTipoDeDocumento, eTipoDeDocumento) = eTipoDeDocumento.CTRC).Count()

                    'Caso tenha mais de um cte vinculado significa que é uma forma especial de lançamento onde se emitem vários ctes para a mesma nota.
                    If QtdCtesVinculados > 1 Then
                        PesoDeChegada = objNotaDeFrete.TotalQuantidadeFiscal
                        NFxMultiplosCTEs = True
                    Else
                        'Soma dos pesos de chegada de cada uma das notas de origem do CTE
                        PesoDeChegada = lstNotasOrigem.Sum(Function(s) s.PesoDeChegada.PesoBruto)
                    End If
                Else
                    'Soma dos pesos de chegada de cada uma das notas de origem do CTE
                    PesoDeChegada = lstNotasOrigem.Sum(Function(s) s.PesoDeChegada.PesoBruto)
                End If

                'Verifica se existe pelo menos um dos lançamentos de peso de chegada onde esteja definido sinistro.
                Sinistro = lstNotasOrigem.Any(Function(s) s.PesoDeChegada.Sinistro)

                '#Complemento
                If ItemFF.Nota.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) Then
                    If Not ItemFF.ViaAdiantamento AndAlso Not ItemFF.ViaAmortizacao Then
                        'Verifica se Quantidade Fiscal da Nota é igual ao do Conhecimento, caso seja menor libera encargos para ajuste porque vai ter Complemento
                        If lstNotasOrigem.SelectMany(Function(s) s.Itens) IsNot Nothing AndAlso lstNotasOrigem.SelectMany(Function(s) s.Itens).Count > 0 Then
                            If ItemFF.Nota.Itens(0).QuantidadeFiscal < lstNotasOrigem.SelectMany(Function(s) s.Itens).Sum(Function(s) s.QuantidadeFiscal) Then
                                ComplementoPeso = lstNotasOrigem.SelectMany(Function(s) s.Itens).Sum(Function(s) s.QuantidadeFiscal) - ItemFF.Nota.Itens(0).QuantidadeFiscal
                                ComplementoValor = ItemFF.Nota.NotasOrigemDestino.Where(Function(x) x.CodigoTipoDeDocumento = eTipoDeDocumento.ComplementoDeFrete).SelectMany(Function(s) s.Itens).Sum(Function(s) s.ValorTotal)
                            End If
                        End If
                    End If
                End If

                txtValorCteComplementar.Text = ComplementoValor.ToString("N2")
                txtValorTotalFrete.Text = (ItemFF.ValorFrete + ComplementoValor).ToString("N2")

                If PesoDeChegada = 0 Then
                    PesoDeChegada = ItemFF.PesoDeChegada
                End If

                If PesoDeChegada > 0 Then
                    txtPesoChegada.Text = String.Format("{0:N0}", PesoDeChegada)
                    Dim objClienteXEmpresa As New [Lib].Negocio.ClienteXEmpresa(objNotaDeFrete.CodigoEmpresa, objNotaDeFrete.EnderecoEmpresa)

                    If objNotaDeFrete.Itens IsNot Nothing AndAlso objNotaDeFrete.Itens.Count > 0 AndAlso objNotaDeFrete.Itens(0).CodigoProduto <> objClienteXEmpresa.CodigoProdutoDeFrete Then
                        txtFreteChegada.Text = String.Format("{0:N2}", ItemFF.FreteCombinado)
                        txtValorKg.Text = String.Format("{0:N4}", ItemFF.FreteCombinado)
                    Else
                        If Left(Session("ssEmpresa"), 8) = "62780383" AndAlso ItemFF.Nota.CodigoTransportador = "63358210000176" Then
                            'NÃO FAZ NADA
                            txtFreteChegada.Text = String.Format("{0:N2}", (ItemFF.Peso * CDec(ItemFF.FreteCombinado)) / 1000)
                            txtValorKg.Text = String.Format("{0:N4}", lstNotasOrigem.SelectMany(Function(s) s.Itens).Sum(Function(s) s.ValorTotal) / CDec(TxtPeso.Text))
                        Else
                            txtFreteChegada.Text = String.Format("{0:N2}", (CDec(PesoDeChegada) * CDec(ItemFF.FreteCombinado)) / 1000)
                            txtValorKg.Text = String.Format("{0:N4}", lstNotasOrigem.SelectMany(Function(s) s.Itens).Sum(Function(s) s.ValorTotal) / CDec(TxtPeso.Text))
                        End If
                    End If

                    '#Complemento
                    If (CDec(TxtPeso.Text) + ComplementoPeso) > CDec(txtPesoChegada.Text) Then
                        txtTolerancia.Text = String.Format("{0:N2}", ItemFF.Peso * CDec("0,0025"))
                    Else
                        txtTolerancia.Text = String.Format("{0:N2}", Decimal.Zero)
                    End If

                    Dim pesoSaida As Decimal = Decimal.Zero

                    If Not String.IsNullOrWhiteSpace(TxtPeso.Text) Then
                        '#Complemento
                        pesoSaida = CDec(TxtPeso.Text) + ComplementoPeso
                    End If

                    Dim pesoChegada As Decimal = Decimal.Zero
                    If Not String.IsNullOrWhiteSpace(txtPesoChegada.Text) Then
                        pesoChegada = CDec(txtPesoChegada.Text)
                    End If

                    Dim pesoTolerancia As Decimal = Decimal.Zero
                    If Not String.IsNullOrWhiteSpace(txtTolerancia.Text) Then
                        pesoTolerancia = CDec(txtTolerancia.Text)
                    End If

                    Dim pesoValorKg As Decimal = Decimal.Zero
                    If Not String.IsNullOrWhiteSpace(txtValorKg.Text) Then
                        pesoValorKg = CDec(txtValorKg.Text)
                    End If

                    If lstNotasOrigem.SelectMany(Function(s) s.Itens) IsNot Nothing AndAlso lstNotasOrigem.SelectMany(Function(s) s.Itens).Count > 0 AndAlso Not lstNotasOrigem.SelectMany(Function(s) s.Itens).Any(Function(s) s.Produto.Agrupar = "N") Then
                        txtDiferenca.Text = String.Format("{0:N2}", Decimal.Zero)
                    Else
                        If Math.Abs(pesoSaida - pesoChegada) > pesoTolerancia Then
                            Dim pesoDiff As Decimal
                            If objEmpresa.ApenasExcedenteTolerancia Then
                                pesoDiff = (Math.Abs(pesoSaida - pesoChegada) - pesoTolerancia) * pesoValorKg
                            Else
                                pesoDiff = (Math.Abs(pesoSaida - pesoChegada)) * pesoValorKg
                            End If

                            If PesoDeChegada > 0 AndAlso Not Sinistro AndAlso pesoSaida > pesoChegada Then
                                txtDiferenca.Text = String.Format("{0:N2}", CDec(pesoDiff.ToString.Substring("0", "5")))
                            ElseIf PesoDeChegada > 0 AndAlso Not Sinistro AndAlso pesoSaida < pesoChegada Then
                                'Dim vlr As Decimal = CDec(txtFreteChegada.Text) - CDec(txtValorFrete.Text)
                                Dim vlr As Decimal = CDec(txtFreteChegada.Text) - CDec(txtValorTotalFrete.Text)
                                txtDiferenca.Text = String.Format("{0:N2}", vlr)
                            End If
                        Else
                            txtDiferenca.Text = String.Format("{0:N2}", Decimal.Zero)
                            txtFreteChegada.Text = String.Format("{0:N2}", (CDec(pesoSaida) * CDec(ItemFF.FreteCombinado)) / 1000)
                        End If
                    End If
                End If
            Else
                txtFreteChegada.Text = String.Format("{0:N2}", ItemFF.FreteChegada)
                txtDiferenca.Text = String.Format("{0:N2}", ItemFF.Diferenca)
            End If

            If PesoDeChegada > 0 AndAlso Sinistro Then
                lblMsg.Text = String.Format("DACTE {0} - FATURA {1} possui sinistro vinculado ao peso de chegada!", objFatura.CodigoNota, objFatura.CodigoFatura)
                pnlMsg.Visible = True
            Else
                lblMsg.Text = String.Empty
                pnlMsg.Visible = False
            End If

            RbCtrc.Checked = ItemFF.CTRC
            RbRecibo.Checked = ItemFF.ReciboFrete
            optAdiantamento.Checked = ItemFF.ViaAdiantamento
            optAmortizacao.Checked = ItemFF.ViaAmortizacao
            optCartaFrete.Checked = ItemFF.ViaCartaFrete
            txtVencimento.Text = DateTime.Now.ToShortDateString()

            If ItemFF.Nota.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) Or ItemFF.Nota.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) Then
                If optAdiantamento.Checked Then
                    ItemFF.Nota.Itens(0).Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(ItemFF.Nota.Itens(0), New List(Of eEtapaEncago) From {eEtapaEncago.Normal, eEtapaEncago.Adiantamento})
                ElseIf optAmortizacao.Checked Then
                    ItemFF.Nota.Itens(0).Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(ItemFF.Nota.Itens(0), New List(Of eEtapaEncago) From {eEtapaEncago.Amortizacao})
                Else
                    ItemFF.Nota.Itens(0).Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(ItemFF.Nota.Itens(0), New List(Of eEtapaEncago) From {eEtapaEncago.Liquidacao})

                    'Verifica se Quantidade Fiscal da Nota é igual ao do Conhecimento, caso seja menor libera encargos para ajuste porque vai ter Complemento
                    If Not NFxMultiplosCTEs AndAlso lstNotasOrigem.SelectMany(Function(s) s.Itens) IsNot Nothing AndAlso lstNotasOrigem.SelectMany(Function(s) s.Itens).Count > 0 Then
                        If ItemFF.Nota.Itens(0).QuantidadeFiscal < lstNotasOrigem.SelectMany(Function(s) s.Itens).Sum(Function(s) s.QuantidadeFiscal) Then
                            TemComplemento = True
                        End If
                    End If
                End If
            ElseIf ItemFF.Nota.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) OrElse ItemFF.Nota.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ComplementoDeFrete) Then
                ItemFF.Nota.Itens(0).Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(ItemFF.Nota.Itens(0), New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
            End If

            Dim vlrSobra As Decimal = Decimal.Zero
            Dim vlrQuebra As Decimal = Decimal.Zero

            Dim vlrSaida As Decimal = Decimal.Zero
            If Not String.IsNullOrWhiteSpace(TxtPeso.Text) Then
                '#Complemento
                vlrSaida = CDec(TxtPeso.Text) + ComplementoPeso
            End If

            Dim vlrChegada As Decimal = Decimal.Zero
            If Not String.IsNullOrWhiteSpace(txtPesoChegada.Text) Then
                vlrChegada = CDec(txtPesoChegada.Text)
            End If

            Dim vlrAdto As Decimal = Decimal.Zero
            If Not String.IsNullOrWhiteSpace(TxtValorAdiantamento.Text) Then
                vlrAdto = CDec(TxtValorAdiantamento.Text)
            End If

            Dim vlrTolerancia As Decimal = Decimal.Zero
            If Not String.IsNullOrWhiteSpace(txtTolerancia.Text) Then
                vlrTolerancia = CDec(txtTolerancia.Text)
            End If

            Dim vlrValorKg As Decimal = Decimal.Zero
            If Not String.IsNullOrWhiteSpace(txtValorKg.Text) Then
                vlrValorKg = CDec(txtValorKg.Text)
            End If

            If ItemFF.Nota.Itens IsNot Nothing AndAlso ItemFF.Nota.Itens.Count > 0 AndAlso Not ItemFF.Nota.Itens.Any(Function(s) s.Produto.Agrupar = "N") Then
                Dim vlrDiff As Decimal
                If objEmpresa.ApenasExcedenteTolerancia Then
                    vlrDiff = ((Math.Abs(vlrSaida - vlrChegada) - vlrTolerancia) * vlrValorKg).ToString.Substring("0", "5")
                Else
                    vlrDiff = ((Math.Abs(vlrSaida - vlrChegada)) * vlrValorKg).ToString.Substring("0", "5")
                End If

                If ItemFF.CodigoEncargo = "BAIXAADTO" Then
                    For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In ItemFF.Nota.Itens(0).Encargos
                        If enc.Codigo.Trim().Contains("LIQUIDO") Then
                            ItemFF.Nota.Itens(0).CarregandoEncargos = True
                            enc.Valor = vlrAdto
                            ItemFF.Nota.Itens(0).CarregandoEncargos = False
                        End If
                    Next
                Else
                    For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In ItemFF.Nota.Itens(0).Encargos
                        If PesoDeChegada > 0 AndAlso Not Sinistro AndAlso enc.Codigo.Trim() = "QUEBRAS" AndAlso vlrSaida > vlrChegada Then
                            If Left(Session("ssEmpresa"), 8) = "62780383" AndAlso ItemFF.Nota.CodigoTransportador = "63358210000176" Then
                                'NÃO FAZ NADA
                            Else
                                If Math.Abs(vlrSaida - vlrChegada) > CDec(txtTolerancia.Text) Then
                                    enc.Valor = Math.Abs(vlrDiff)
                                    vlrQuebra = Math.Abs(vlrDiff)
                                End If
                            End If
                        ElseIf PesoDeChegada > 0 AndAlso Not Sinistro AndAlso enc.Codigo.Trim() = "SOBRAS" AndAlso vlrSaida < vlrChegada Then
                            If Left(Session("ssEmpresa"), 8) = "62780383" AndAlso ItemFF.Nota.CodigoTransportador = "63358210000176" Then
                                'NÃO FAZ NADA
                            Else
                                Dim vlr As Decimal = CDec(txtFreteChegada.Text) - CDec(txtValorFrete.Text)
                                enc.Valor = Math.Abs(vlr)
                                vlrSobra = Math.Abs(vlr)
                            End If
                        ElseIf enc.Codigo.Trim().Contains("QUEBRAS - P. CH") AndAlso vlrSaida > vlrChegada Then
                            '#Complemento
                            If Left(Session("ssEmpresa"), 8) = "62780383" AndAlso ItemFF.Nota.CodigoTransportador = "63358210000176" Then
                                'NÃO FAZ NADA
                            Else
                                If Math.Abs(vlrSaida - vlrChegada) > CDec(txtTolerancia.Text) Then
                                    Dim vlr As Decimal = 0
                                    vlr = Math.Abs(CDec(txtFreteChegada.Text) - CDec(txtValorFrete.Text))
                                    enc.Valor = Math.Abs(vlr)
                                End If
                            End If
                        ElseIf enc.Codigo.Trim().Contains("SOBRAS - P. CH") AndAlso vlrSaida < vlrChegada Then
                            If Left(Session("ssEmpresa"), 8) = "62780383" AndAlso ItemFF.Nota.CodigoTransportador = "63358210000176" Then
                                'NÃO FAZ NADA
                            Else
                                Dim vlr As Decimal = CDec(txtFreteChegada.Text) - CDec(txtValorFrete.Text)
                                enc.Valor = Math.Abs(vlr)
                            End If
                        ElseIf enc.Codigo.Trim().Contains("LIQUIDOAPAGAR") Then
                            ItemFF.Nota.Itens(0).CarregandoEncargos = True
                            Dim vlrEncargo As Decimal = objNotaFiscal.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "TAXA CADASTRO" OrElse s.Codigo.Trim = "TARIFA SEGURO").Sum(Function(s) s.Valor)
                            '#Complemento
                            enc.Valor = CDec(txtFreteChegada.Text) - vlrAdto - vlrQuebra - vlrEncargo - ComplementoValor
                            ItemFF.Nota.Itens(0).CarregandoEncargos = False
                        End If
                    Next
                End If
            Else
                If ItemFF.CodigoEncargo = "BAIXAADTO" Then
                    For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In ItemFF.Nota.Itens(0).Encargos
                        If enc.Codigo.Trim().Contains("LIQUIDO") Then
                            ItemFF.Nota.Itens(0).CarregandoEncargos = True
                            enc.Valor = vlrAdto
                            ItemFF.Nota.Itens(0).CarregandoEncargos = False
                        End If
                    Next
                Else
                    For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In ItemFF.Nota.Itens(0).Encargos
                        If enc.Codigo.Trim().Contains("QUEBRAS - P. CH") AndAlso vlrSaida > vlrChegada Then
                            If Left(Session("ssEmpresa"), 8) = "62780383" AndAlso ItemFF.Nota.CodigoTransportador = "63358210000176" Then
                                'NÃO FAZ NADA
                            Else
                                Dim vlr As Decimal = CDec(txtFreteChegada.Text) - CDec(txtValorFrete.Text)
                                enc.Valor = Math.Abs(vlr)
                            End If
                        ElseIf enc.Codigo.Trim().Contains("SOBRAS - P. CH") AndAlso vlrSaida < vlrChegada Then
                            If Left(Session("ssEmpresa"), 8) = "62780383" AndAlso ItemFF.Nota.CodigoTransportador = "63358210000176" Then
                                'NÃO FAZ NADA
                            Else
                                Dim vlr As Decimal = CDec(txtFreteChegada.Text) - CDec(txtValorFrete.Text)
                                enc.Valor = Math.Abs(vlr)
                            End If
                        ElseIf enc.Codigo.Trim().Contains("LIQUIDOAPAGAR") Then
                            ItemFF.Nota.Itens(0).CarregandoEncargos = True
                            Dim vlrEncargo As Decimal = objNotaFiscal.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "TAXA CADASTRO" OrElse s.Codigo.Trim = "TARIFA SEGURO").Sum(Function(s) s.Valor)
                            enc.Valor = CDec(txtFreteChegada.Text) - vlrAdto - vlrQuebra - vlrEncargo
                            ItemFF.Nota.Itens(0).CarregandoEncargos = False
                        End If
                    Next
                End If
            End If

            RbCtrc.Enabled = False
            RbRecibo.Enabled = False
            optAdiantamento.Enabled = False
            optAmortizacao.Enabled = False

            TxtNrFrete.Enabled = False
            BtnNumeroNotaFrete.Enabled = False
            txtPesoChegada.Enabled = True

            If txtPesoChegada.Text.Length > 0 Then
                If CType(txtPesoChegada.Text, Decimal) > 0 Then
                    txtPesoChegada_TextChanged(Nothing, Nothing)
                End If
            End If

            If bConsulta Then
                btnCalcular.Enabled = False
            Else
                btnCalcular.Enabled = True
            End If

            SalvarItemFF()
            CarregaDgEncargos(lstEncargos.Where(Function(s) s.Codigo = "TARIFA SEGURO").ToList(), TemComplemento)

            'Apresenta o Financeiro do Frete
            If FinanceiroNovo Then
                ucFinanceiro.Visible = True
                ucFinanceiro.Limpar()
                ucFinanceiro.AtualizaFaturaFrete(CDec(txtValorFrete.Text), CDec(txtFreteChegada.Text), CDate(txtVencimento.Text), False)
            End If

            If TemComplemento Then
                If ComplementoValor = 0 Then
                    MsgBox(Me.Page, "Atenção, esse conhecimento está com a quantidade menor do que o peso fiscal da Nota. Peça a emissão do complemento pela Transportadora." _
                           & " A diferença será atribuída como sobra. Caso necessário ajuste os valores e clique em Calcular.")
                Else
                    MsgBox(Me.Page, "Atenção, esse Conhecimento têm complemento lançado. Lembre-se de fazer o encerramento da fatura do complemento também.")
                End If
            End If

            TabOutros.ActiveTabIndex = 0

        End If

        If objFatura.Itens.Where(Function(x) x.EmpresaCnpj = ItemFF.Nota.CodigoEmpresa And x.ClienteCnpj = ItemFF.Nota.CodigoCliente And x.Serie = ItemFF.Nota.Serie And x.NumeroNota = ItemFF.Nota.Codigo).Count = 0 Then
            btnOk.Enabled = True
            btnExcluir.Enabled = False
        Else
            btnOk.Enabled = False
            btnExcluir.Enabled = True
        End If

    End Sub

    Private Sub CarregarFormItemFF()
        Dim NfDeFrete As [Lib].Negocio.NotaFiscal = Session("NotasDeFrete" & HID.Value)
        Dim Fatura As [Lib].Negocio.FaturaDeFrete = Session("Fatura" & HID.Value)
        RecuperaItemFF()

        If ItemFF.NumeroNota <> 0 Then
            Dim ConsTrans As New [Lib].Negocio.Cliente(ItemFF.ClienteCnpj, ItemFF.ClienteEnd)
            TxtNrFrete.Text = ItemFF.NumeroNota
            ddlFilial.SelectedValue = ItemFF.EmpresaCnpj & "-" & ItemFF.EmpresaEnd
            txtTransportador.Text = ConsTrans.Nome
            TxtPeso.Text = ItemFF.Peso
            txtFreteCombinado.Text = ItemFF.FreteCombinado
            txtValorFrete.Text = ItemFF.ValorFrete
            txtPesoChegada.Text = ItemFF.PesoDeChegada

            If ItemFF.Nota.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) Or ItemFF.Nota.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) Then
                Dim pEtapa As eEtapaEncago = IIf(optAdiantamento.Checked, eEtapaEncago.Adiantamento, IIf(optAmortizacao.Checked, eEtapaEncago.Amortizacao, eEtapaEncago.Liquidacao))

                NfDeFrete.Itens(0).Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(NfDeFrete.Itens(0), New List(Of eEtapaEncago) From {pEtapa})
            ElseIf ItemFF.Nota.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) OrElse ItemFF.Nota.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ComplementoDeFrete) Then
                NfDeFrete.Itens(0).Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(NfDeFrete.Itens(0), New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
            End If

            Dim Adiantamento As [Lib].Negocio.NotaFiscalXItemXEncargo = Nothing
            Dim lstEncargos As List(Of NotaFiscalXItemXEncargo) = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(NfDeFrete.Itens(0), New List(Of eEtapaEncago) From {eEtapaEncago.Adiantamento})
            For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In lstEncargos
                If enc.Codigo = "ADTODEFRETE" Then
                    Adiantamento = enc
                    Exit For
                End If
            Next

            If Not Adiantamento Is Nothing Then
                txtPercAdiantamento.Text = Adiantamento.Percentual.ToString + "%"
                TxtValorAdiantamento.Text = Adiantamento.Valor
            Else
                txtPercAdiantamento.Text = "0%"
                TxtValorAdiantamento.Text = 0
                If optAdiantamento.Checked = True Then
                    MsgBox(Me.Page, "Não existe adiantamento para este frete!")
                    Exit Sub
                End If
            End If

            If optAdiantamento.Checked Or optAmortizacao.Checked Then
                For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In NfDeFrete.Itens(0).Encargos
                    If enc.Codigo = "ADTODEFRETE" Then
                        Adiantamento.Base = enc.Valor
                        Exit For
                    End If
                Next

                txtPesoChegada.Enabled = False
            Else
                For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In NfDeFrete.Itens(0).Encargos
                    If enc.Codigo = "ADTODEFRETE" Then
                        Adiantamento.Base = 0
                        Exit For
                    End If
                Next

                txtPesoChegada.Enabled = True
            End If

            Dim TRANSFCARTA As [Lib].Negocio.NotaFiscalXItemXEncargo
            If Fatura.CodigoEmpresa = NfDeFrete.CodigoEmpresa Then
                For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In NfDeFrete.Itens(0).Encargos
                    If enc.Codigo = "TRANSFCARTA" Then
                        TRANSFCARTA = enc
                        Exit For
                    End If
                Next
            End If

            If ItemFF.PesoDeChegada > 0 Then
                txtPesoChegada_TextChanged(Nothing, Nothing)
            End If

            If Fatura.LancamentoManual = 0 Then

                If ItemFF.JaFaturada Then

                    MsgBox(Me.Page, Session("ssMessage"))
                    dgEncargos.DataSource = Nothing
                    btnCalcular.Enabled = False
                    LimparParcial()
                    Exit Sub

                Else

                    ItemFF.IUD = "I"
                    optAdiantamento.Enabled = False
                    optAmortizacao.Enabled = False
                    'optCartaFrete.Enabled = False
                    RbCtrc.Enabled = False
                    RbRecibo.Enabled = False
                    TxtNrFrete.Enabled = False

                End If

            Else

                If ItemFF.JaFaturadaRegistroMestre Then

                    MsgBox(Me.Page, Session("ssMessage"))
                    dgEncargos.DataSource = Nothing
                    btnCalcular.Enabled = False
                    LimparParcial()
                    Exit Sub

                Else

                    ItemFF.IUD = "I"
                    optAdiantamento.Enabled = False
                    optAmortizacao.Enabled = False
                    'optCartaFrete.Enabled = False
                    RbCtrc.Enabled = False
                    RbRecibo.Enabled = False
                    TxtNrFrete.Enabled = False

                End If

            End If

        End If

    End Sub

    Private Function ValidaCamposProgFatura() As Boolean
        EmpresaProg = DdlEmpresaProg.SelectedValue.ToString.Split("-")
        If DdlEmpresaProg.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa pagadora obrigatória para programar/consultar fatura!")
            Return False
        ElseIf Trim(txtVencimentoDe.Text) = "" Or Trim(txtVencimentoAte.Text) = "" Then
            MsgBox(Me.Page, "Datas inicial e final são obrigatórias!")
            Return False
        ElseIf CDate(txtVencimentoAte.Text) < CDate(txtVencimentoDe.Text) Then
            MsgBox(Me.Page, "Data de final não pode ser inferior a data inicial!")
            Return False
        End If
        Return True
    End Function

    Public Sub CarregarNotasdeFrete()
        If Not Session("ConsultaNotasDeFrete" & HID.Value) Is Nothing Then
            Session("NotasDeFrete" & HID.Value) = Session("ConsultaNotasDeFrete" & HID.Value)
            CarregaNotaDeFrete()
            Session.Remove("ConsultaNotasDeFrete" & HID.Value)
        End If
    End Sub

    Public Sub CarregarFatura()
        If Session("ConsultaFatura" & HID.Value) IsNot Nothing Then
            Dim aux As Boolean = True
            Dim RetConsultaFaturas As [Lib].Negocio.FaturaDeFrete = Session("ConsultaFatura" & HID.Value)

            If FinanceiroNovo Then
                aux = RetConsultaFaturas.ListTituloFatura.Where(Function(s) s.TituloNovo.CodigoSituacao = eSituacao.Normal)(0).TituloNovo.CodigoProvisao <> 3
            Else
                aux = RetConsultaFaturas.ListTituloFatura.Where(Function(s) s.Titulo.CodigoSituacao = eSituacao.Normal)(0).Titulo.CodigoProvisao <> 2
            End If

            If aux Then
                If RetConsultaFaturas.JaFaturada Then
                    RecuperaItemFF()
                    If ItemFF Is Nothing AndAlso RetConsultaFaturas IsNot Nothing AndAlso RetConsultaFaturas.Itens IsNot Nothing AndAlso RetConsultaFaturas.Itens.Count > 0 Then
                        Session("NotasDeFrete" & HID.Value) = RetConsultaFaturas.Itens(0).Nota
                        ItemFF = RetConsultaFaturas.Itens(0)
                        SalvarItemFF()
                    End If

                    If Not Session("NotasDeFrete" & HID.Value) Is Nothing Then

                        Dim NfDeFrete As New [Lib].Negocio.NotaFiscal(CType(Session("NotasDeFrete" & HID.Value), [Lib].Negocio.NotaFiscal))
                        If NfDeFrete IsNot Nothing AndAlso Not NfDeFrete.NossaEmissao _
                                AndAlso Not String.IsNullOrWhiteSpace(ItemFF.CodigoEncargo) _
                                AndAlso ItemFF.CodigoEncargo.Contains("LIQUIDOAPAGAR") Then
                            If NfDeFrete IsNot Nothing AndAlso NfDeFrete.NotasXNotas IsNot Nothing Then
                                Dim nf As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                                nf.CodigoEmpresa = NfDeFrete.NotasXNotas.OrigemEmpresaCnpj
                                nf.EnderecoEmpresa = NfDeFrete.NotasXNotas.OrigemEndEmpresa
                                nf.CodigoCliente = NfDeFrete.NotasXNotas.OrigemClienteCnpj
                                nf.EnderecoCliente = NfDeFrete.NotasXNotas.OrigemEndCliente
                                nf.EntradaSaida = IIf(NfDeFrete.NotasXNotas.OrigemEntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                                nf.Codigo = NfDeFrete.NotasXNotas.OrigemNota
                                nf.Serie = NfDeFrete.NotasXNotas.OrigemSerie
                                nf = New [Lib].Negocio.NotaFiscal(nf)

                                If (NfDeFrete.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.Estadia) AndAlso NfDeFrete.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.ComplementoDeFrete)) AndAlso nf.CodigoTipoDeDocumento = eTipoDeDocumento.CTRC Then
                                    MsgBox(Me.Page, "Não é possível realizar o lançamento do saldo, pois já existe um conhecimento de transporte vinculado!")
                                    Exit Sub
                                ElseIf nf.CodigoTipoDeDocumento = eTipoDeDocumento.Nota Then
                                    Dim origem As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                                    origem.CodigoEmpresa = NfDeFrete.NotasXNotas.OrigemEmpresaCnpj
                                    origem.EnderecoEmpresa = NfDeFrete.NotasXNotas.OrigemEndEmpresa
                                    origem.CodigoCliente = NfDeFrete.NotasXNotas.OrigemClienteCnpj
                                    origem.EnderecoCliente = NfDeFrete.NotasXNotas.OrigemEndCliente
                                    origem.EntradaSaida = IIf(NfDeFrete.NotasXNotas.OrigemEntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                                    origem.Codigo = NfDeFrete.NotasXNotas.OrigemNota
                                    origem.Serie = NfDeFrete.NotasXNotas.OrigemSerie
                                    origem = New [Lib].Negocio.NotaFiscal(origem)
                                    If (NfDeFrete.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.Estadia) AndAlso NfDeFrete.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.ComplementoDeFrete)) AndAlso nf.CodigoTipoDeDocumento = eTipoDeDocumento.CTRC AndAlso NfDeFrete.Codigo <> nf.Codigo Then
                                        MsgBox(Me.Page, "Não é possível realizar o lançamento do saldo, pois já existe um conhecimento de transporte vinculado!")
                                        Exit Sub
                                    End If
                                End If

                                Dim lstNotasXNotas As New [Lib].Negocio.ListNotasXNotas(NfDeFrete)
                                Dim objNota As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                                objNota.CodigoEmpresa = lstNotasXNotas(0).OrigemEmpresaCnpj
                                objNota.EnderecoEmpresa = lstNotasXNotas(0).OrigemEndEmpresa
                                objNota.CodigoCliente = lstNotasXNotas(0).OrigemClienteCnpj
                                objNota.EnderecoCliente = lstNotasXNotas(0).OrigemEndCliente
                                objNota.EntradaSaida = IIf(lstNotasXNotas(0).OrigemEntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                                objNota.Codigo = lstNotasXNotas(0).OrigemNota
                                objNota.Serie = lstNotasXNotas(0).OrigemSerie
                                objNota = New [Lib].Negocio.NotaFiscal(objNota)
                                If NfDeFrete.CodigoTipoDeDocumentoFrete <> eTipoDeDocumentoFrete.Estadia AndAlso NfDeFrete.CodigoTipoDeDocumentoFrete <> eTipoDeDocumentoFrete.Complemento Then
                                    If NfDeFrete.PesoDeChegada Is Nothing Then
                                        MsgBox(Me.Page, String.Format("É necessário existir um lançamento no peso de chegada para o CTRC {0}-{1} e NF {2}-{3}!", NfDeFrete.Codigo, NfDeFrete.Serie, objNota.Codigo, objNota.Serie))
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Else
                            If NfDeFrete IsNot Nothing AndAlso NfDeFrete.NossaEmissao _
                                    AndAlso Not String.IsNullOrWhiteSpace(ItemFF.CodigoEncargo) _
                                    AndAlso ItemFF.CodigoEncargo.Contains("LIQUIDOAPAGAR") Then
                                Dim objCirculacao As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal(NfDeFrete)
                                If objCirculacao IsNot Nothing AndAlso objCirculacao.NotasXNotas IsNot Nothing Then
                                    Dim lstNotas As New [Lib].Negocio.ListNotasFiscais()
                                    Dim lst As New [Lib].Negocio.ListNotasXNotas(objCirculacao, True, True)
                                    For Each item As [Lib].Negocio.NotasXNotas In lst
                                        Dim nf As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                                        nf.CodigoEmpresa = objCirculacao.NotasXNotas.EmpresaCnpj
                                        nf.EnderecoEmpresa = objCirculacao.NotasXNotas.EndEmpresa
                                        nf.CodigoCliente = objCirculacao.NotasXNotas.ClienteCnpj
                                        nf.EnderecoCliente = objCirculacao.NotasXNotas.EndCliente
                                        nf.EntradaSaida = IIf(objCirculacao.NotasXNotas.EntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                                        nf.Codigo = objCirculacao.NotasXNotas.NumeroNota
                                        nf.Serie = objCirculacao.NotasXNotas.Serie
                                        nf = New [Lib].Negocio.NotaFiscal(nf)
                                        lstNotas.Add(nf)
                                    Next

                                    If lstNotas Is Nothing OrElse Not lstNotas.Count > 0 OrElse Not lstNotas.Any(Function(s) s.CodigoTipoDeDocumento = CInt(eTipoDeDocumentoFrete.Comprovacao)) Then
                                        MsgBox(Me.Page, "É necessário possuir um conhecimento de comprovação para realizar o lançamento de saldo!")
                                        Exit Sub
                                    End If
                                End If

                                Dim lstNotasXNotas As New [Lib].Negocio.ListNotasXNotas(objCirculacao)
                                Dim objNota As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                                objNota.CodigoEmpresa = lstNotasXNotas(0).OrigemEmpresaCnpj
                                objNota.EnderecoEmpresa = lstNotasXNotas(0).OrigemEndEmpresa
                                objNota.CodigoCliente = lstNotasXNotas(0).OrigemClienteCnpj
                                objNota.EnderecoCliente = lstNotasXNotas(0).OrigemEndCliente
                                objNota.EntradaSaida = IIf(lstNotasXNotas(0).OrigemEntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                                objNota.Codigo = lstNotasXNotas(0).OrigemNota
                                objNota.Serie = lstNotasXNotas(0).OrigemSerie
                                objNota = New [Lib].Negocio.NotaFiscal(objNota)
                                If NfDeFrete.CodigoTipoDeDocumentoFrete <> eTipoDeDocumentoFrete.Estadia AndAlso NfDeFrete.CodigoTipoDeDocumentoFrete <> eTipoDeDocumentoFrete.Complemento Then
                                    If NfDeFrete.PesoDeChegada Is Nothing Then
                                        MsgBox(Me.Page, String.Format("É necessário existir um lançamento no peso de chegada para o CTRC {0}-{1} e NF {2}-{3}!", NfDeFrete.Codigo, NfDeFrete.Serie, objNota.Codigo, objNota.Serie))
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If

                    End If

                    Session("Fatura" & HID.Value) = RetConsultaFaturas
                    txtFatura.Text = RetConsultaFaturas.CodigoFatura
                    txtDataMovimento.Text = RetConsultaFaturas.Movimento
                    txtDataVencimento.Text = RetConsultaFaturas.Vencimento
                    txtValorFatura.Text = RetConsultaFaturas.ValorDaFatura
                    txtValorLancado.Text = RetConsultaFaturas.ValorLancadoFatura
                    txtValorSaldo.Text = RetConsultaFaturas.ValorSaldoFatura
                    txtConveniadoLan.Text = RetConsultaFaturas.Conveniado.Nome
                    CodigoConveniadoLan.Value = RetConsultaFaturas.CodigoConveniado & "-" & RetConsultaFaturas.Conveniado.CodigoEndereco

                    'If FinanceiroNovo Then
                    ucFinanceiro.CarregarControle()
                    'End If

                    If RetConsultaFaturas.PagarReceber = "P" Then
                        rdbPagarLancamento.Checked = True
                    Else
                        rdbReceberLancamento.Checked = True
                    End If

                    rdbPagarLancamento.Enabled = False
                    rdbReceberLancamento.Enabled = False

                    Dim Fatura As [Lib].Negocio.FaturaDeFrete = CType(Session("Fatura" & HID.Value), [Lib].Negocio.FaturaDeFrete)
                    Dim ListNfDeFrete As New [Lib].Negocio.ListFaturasDeFretesXItens(Fatura)
                    DgComposicaoFatura.DataSource = ListNfDeFrete
                    DgComposicaoFatura.DataBind()
                    btnEncerrar.Enabled = ListNfDeFrete IsNot Nothing AndAlso ListNfDeFrete.Count > 0
                    TabOutros.ActiveTabIndex = IIf(ListNfDeFrete IsNot Nothing AndAlso ListNfDeFrete.Count > 0, 1, 0)
                    Session.Remove("ConsultaFatura" & HID.Value)
                End If
            Else
                txtFatura.Text = ""
                Session.Remove("ConsultaFatura" & HID.Value)
                MsgBox(Me.Page, "Esta fatura já foi previsionada!")
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Not Session("objConveniadoPF" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objConveniadoPF" & HID.Value), [Lib].Negocio.Cliente))
            txtConveniadoPF.Text = itemCliente.Text
            CodigoConveniadoPF.Value = itemCliente.Value
            Session.Remove("objConveniadoPF" & HID.Value)
        ElseIf Not Session("objConveniadoPROG" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objConveniadoPROG" & HID.Value), [Lib].Negocio.Cliente))
            txtConveniadoProg.Text = itemCliente.Text
            CodigoConveniadoProg.Value = itemCliente.Value
            Session.Remove("objConveniadoPROG" & HID.Value)
        ElseIf Not Session("objClientePROG" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClientePROG" & HID.Value), [Lib].Negocio.Cliente))
            txtClienteProg.Text = itemCliente.Text
            CodigoClienteProg.Value = itemCliente.Value
            Session.Remove("objClientePROG" & HID.Value)
        ElseIf Not Session("objConveniadoLan" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objConveniadoLan" & HID.Value), [Lib].Negocio.Cliente))
            txtConveniadoLan.Text = itemCliente.Text
            CodigoConveniadoLan.Value = itemCliente.Value
            Session.Remove("objConveniadoLan" & HID.Value)
        ElseIf Not Session("objEmpresaCDF" & HID.Value) Is Nothing Then
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objEmpresaCDF" & HID.Value), [Lib].Negocio.Cliente))
            txtEmpresa.Text = itemEmpresa.Text
            txtCodigoEmpresa.Value = itemEmpresa.Value
            If CType(obj, [Lib].Negocio.Cliente).UnidadeDeNegocio IsNot Nothing Then
                txtUnNegocio.Value = CType(obj, [Lib].Negocio.Cliente).UnidadeDeNegocio.Codigo
            End If
            Session.Remove("objEmpresaCDF" & HID.Value)
        ElseIf Not Session("objContaBancariaFRT" & HID.Value) Is Nothing Then
            txtCodBanco.Text = CType(Session("objContaBancariaFRT" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).CodigoBanco
            txtNomeBanco.Text = CType(Session("objContaBancariaFRT" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).NomeBanco
            txtCodAgencia.Text = CType(Session("objContaBancariaFRT" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).CodigoAgencia
            txtDigitoAgencia.Text = CType(Session("objContaBancariaFRT" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).DigitoAgencia
            txtPracaAgencia.Text = CType(Session("objContaBancariaFRT" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).Praca
            txtConta.Text = CType(Session("objContaBancariaFRT" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).ContaCorrente
            txtDigitoConta.Text = CType(Session("objContaBancariaFRT" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).DigitoConta
            Session.Remove("objContaBancariaFRT" & HID.Value)
        ElseIf Session("ConsultaFatura" & HID.Value) IsNot Nothing Then
            CarregarFatura()
        End If
    End Sub

    Private Sub CarregaPreFaturas(Optional ByVal Consulta As Boolean = False)

        Dim Fatura As New [Lib].Negocio.FaturaDeFrete()
        If Not String.IsNullOrWhiteSpace(txtCodigoEmpresa.Value) Then
            EmpresaPF = txtCodigoEmpresa.Value.Split("-")
            Fatura.CodigoEmpresa = EmpresaPF(0)
            Fatura.EnderecoEmpresa = EmpresaPF(1)
        End If

        If CodigoConveniadoPF.Value.ToString.Length > 0 Then
            ConveniadoPF = CodigoConveniadoPF.Value.ToString.Split("-")
            Fatura.CodigoConveniado = ConveniadoPF(0)
            Fatura.EnderecoConveniado = ConveniadoPF(1)
        End If

        If TxtFaturaPF.Text.Length > 0 Then
            Fatura.CodigoFatura = TxtFaturaPF.Text
        End If

        Dim Faturas As New [Lib].Negocio.ListFaturaDeFrete(Fatura, True, New Cliente, txtConsultaDePF.Text, txtConsultaAtePF.Text, IIf(rdBaixadosPF.Checked, "1", "3"))

        If Faturas.Count > 0 Then
            Session("ssFaturasPF" & HID.Value) = Faturas
            'DgPreFaturas.DataSource = Faturas.ToArray

            DgPreFaturas.DataSource = From f In Faturas.SelectMany(Function(s) s.ListTituloFatura)
                                      Select f.FaturaDeFrete.CodigoFatura, f.FaturaDeFrete.CodigoConveniado, f.FaturaDeFrete.EnderecoConveniado,
                                             f.FaturaDeFrete.ConveniadoNome, f.FaturaDeFrete.Movimento, f.FaturaDeFrete.ValorDaFatura, f.FaturaDeFrete.ValorLancadoFatura,
                                             f.Titulo, f.Titulo.Prorrogacao
            DgPreFaturas.DataBind()
        Else
            DgPreFaturas.DataBind()
            MsgBox(Me.Page, "Registro(s) não encontrado(s) para esta seleção/período!")
            Session.Remove("ssFaturasPF" & HID.Value)
        End If
    End Sub

    Private Sub LimparPreFaturas()
        txtConveniadoPF.Text = ""
        CodigoConveniadoPF.Value = ""
        TxtFaturaPF.Text = ""
        TxtValorFaturaPF.Text = ""
        txtMovimentoPF.Text = Now.Date.ToString("dd/MM/yyyy")
        txtVencimentoPF.Text = Now.Date.AddDays(3).ToString("dd/MM/yyyy")
        ddlCarteiraPF.SelectedValue = ""
        txtCodBanco.Text = ""
        txtNomeBanco.Text = ""
        txtCodAgencia.Text = ""
        txtDigitoAgencia.Text = ""
        txtConta.Text = ""
        txtPracaAgencia.Text = ""
        txtDigitoConta.Text = ""
        txtConsultaDePF.Text = ""
        txtConsultaAtePF.Text = ""
        lnkNovoPF.Enabled = True
        lnkAtualizarPF.Enabled = False
        lnkExcluirPF.Enabled = False
        LiberaEmpresa()
    End Sub

    Private Sub Limpar(Optional ByVal clearAll As Boolean = True)
        If clearAll Then
            'ddlEmpresa.SelectedValue = ""
            CodigoConveniadoLan.Value = ""
            txtConveniadoLan.Text = ""
        End If

        rdbPagarLancamento.Enabled = True
        rdbReceberLancamento.Enabled = True

        lblMsg.Text = ""
        pnlMsg.Visible = False
        txtFatura.Text = ""
        txtValorFatura.Text = ""
        txtValorLancado.Text = ""
        txtValorSaldo.Text = ""
        txtDataMovimento.Text = ""
        txtDataVencimento.Text = ""
        txtFatura.Enabled = True
        ddlEmpresa.Enabled = True
        imgFatura.Enabled = True
        btnConveniadoLan.Enabled = True
        btnEncerrar.Enabled = False
        txtMovimentoPF.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtVencimentoPF.Text = DateTime.Now.AddDays(3).ToString("dd/MM/yyyy")
        txtVencimento.Text = DateTime.Now.ToShortDateString()
        DgComposicaoFatura.DataSource = Nothing
        DgComposicaoFatura.DataBind()
        TabOutros.ActiveTabIndex = 0
        LimparParcial()
        LiberaEmpresa()

        DgComposicaoFatura.DataSource = Nothing
        DgComposicaoFatura.DataBind()

        Session.Remove("Fatura" & HID.Value)
        HID.Value = Guid.NewGuid().ToString()

    End Sub

    Private Sub LimparParcial()

        RbCtrc.Enabled = True
        RbCtrc.Checked = True
        optAdiantamento.Enabled = True
        optCartaFrete.Enabled = True
        optAdiantamento.Checked = True
        optBaixaAdiantamento.Checked = False
        optAmortizacao.Checked = False
        TxtNrFrete.Enabled = True
        TxtNrFrete.Text = ""
        ddlFilial.SelectedValue = ""
        txtTransportador.Text = ""
        TxtPeso.Text = ""
        txtFreteCombinado.Text = ""
        txtValorFrete.Text = ""
        txtValorCteComplementar.Text = String.Empty
        txtValorTotalFrete.Text = String.Empty
        txtPesoChegada.Text = ""
        txtFreteChegada.Text = ""
        txtDiferenca.Text = ""
        txtPercAdiantamento.Text = ""
        TxtValorAdiantamento.Text = ""
        txtVencimento.Text = DateTime.Now.ToShortDateString()
        txtValorKg.Text = ""
        txtTolerancia.Text = ""

        txtPesoChegada.Enabled = True
        BtnNumeroNotaFrete.Enabled = True
        btnOk.Enabled = False
        btnCalcular.Enabled = False
        dgEncargos.DataSource = Nothing
        dgEncargos.DataBind()

        Session.Remove("NotasDeFrete" & HID.Value)
        Session.Remove("ItemFF" & HID.Value)
        Session.Remove("ssFaturasPF" & HID.Value)
        Session.Remove("ssFaturasProg" & HID.Value)
        Session.Remove("objContaBancariaFRT" & HID.Value)
        Session.Remove("ConsultaNotasDeFrete" & HID.Value)
        Session.Remove("ConsultaFatura" & HID.Value)
        Session.Remove("objConveniadoPF" & HID.Value)
        Session.Remove("objConveniadoPROG" & HID.Value)
        Session.Remove("objClientePROG" & HID.Value)
        Session.Remove("objConveniadoLan" & HID.Value)
        ucConsultaEmpresas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaFaturasDeFrete.SetarHID(HID.Value)
        ucConsultaNotasDeFrete.SetarHID(HID.Value)
        ucConsultaDadosBancarios.SetarHID(HID.Value)
        'If FinanceiroNovo Then
        ucFinanceiro.Limpar()
        ucFinanceiro.SetarHID(HID.Value)
        ucFinanceiro.SetarOrigem = "FRETE"
        ucFinanceiro.Visible = True
        'End If
        LiberaEmpresa()

    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            btnEmpresa.Enabled = False
            ddlEmpresa.Enabled = False
            DdlEmpresaProg.Enabled = False
        End If
    End Sub

    Private Function ValidaCamposPreFatura() As Boolean
        EmpresaPF = txtCodigoEmpresa.Value.Split("-")
        If String.IsNullOrWhiteSpace(txtCodigoEmpresa.Value) Then
            MsgBox(Me.Page, "Empresa pagadora obrigatória para pré-fatura!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtConveniadoPF.Text) Then
            MsgBox(Me.Page, "Conveniado é obrigatório!")
            txtConveniadoPF.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtCodBanco.Text) Then
            MsgBox(Me.Page, "Banco é obrigatório!")
            btnDadosBancarios.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtCodAgencia.Text) Then
            MsgBox(Me.Page, "Agência é obrigatório!")
            btnDadosBancarios.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtConta.Text) Then
            MsgBox(Me.Page, "Conta é obrigatório!")
            btnDadosBancarios.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(TxtValorFaturaPF.Text) Then
            MsgBox(Me.Page, "Valor da fatura obrigatório para pré-fatura!")
            TxtValorFaturaPF.Focus()
            Return False
        ElseIf CDate(txtMovimentoPF.Text) > CDate(txtVencimentoPF.Text) Then
            MsgBox(Me.Page, "Data de vencimento tem que ser maior que data movimento!")
            Return False
        ElseIf CDate(txtVencimentoPF.Text) < Funcoes.ValidaDataUtil(EmpresaPF(0), EmpresaPF(1), CDate(txtMovimentoPF.Text).AddDays(3)) Then
            MsgBox(Me.Page, "Data de vencimento não pode ser inferior a 72 hrs do movimento!")
            Return False
        ElseIf ddlCarteiraPF.SelectedIndex = 0 AndAlso Not FinanceiroNovo Then
            MsgBox(Me.Page, "Carteira é obrigatório!")
            Return False
        End If
        Return True
    End Function

    Private Sub SalvarPreFatura()
        If ValidaCamposPreFatura() Then
            Dim fatura As New [Lib].Negocio.FaturaDeFrete
            Dim numerador As New [Lib].Negocio.Numerador(1)
            Dim Sqls As New ArrayList
            EmpresaPF = txtCodigoEmpresa.Value.Split("-")
            ConveniadoPF = CodigoConveniadoPF.Value.ToString.Split("-")

            fatura.IUD = "I"
            fatura.CodigoEmpresa = EmpresaPF(0)
            fatura.EnderecoEmpresa = EmpresaPF(1)
            fatura.CodigoConveniado = ConveniadoPF(0)
            fatura.EnderecoConveniado = ConveniadoPF(1)
            'Fatura lançada manualmente
            fatura.LancamentoManual = 1
            fatura.CodigoFatura = [Lib].Negocio.Numerador.PegarNumero(HttpContext.Current.Session("ssNomeServidor"), eTiposNumerador.FaturaDeFrete)
            fatura.Movimento = CDate(txtMovimentoPF.Text).ToString("yyyy-MM-dd")
            fatura.ValorDaFatura = CDec(TxtValorFaturaPF.Text)

            If Not fatura.JaFaturada Then

                Dim objCotacao As New [Lib].Negocio.Cotacao(3, CDate(txtVencimentoPF.Text).ToString("yyyy-MM-dd"))

                If Not FinanceiroNovo Then

                    Dim t As New [Lib].Negocio.Titulo
                    t.CodigoProvisao = 3
                    t.CodigoCarteira = ddlCarteiraPF.SelectedValue
                    t.CodigoTipoPgto = 1
                    t.CodigoSituacao = 1

                    If rdbPagar.Checked Then
                        t.ReceberPagar = "P"
                    Else
                        t.ReceberPagar = "R"
                    End If

                    t.CodigoMoeda = 1
                    t.Movimento = CDate(txtMovimentoPF.Text).ToString("yyyy-MM-dd")
                    t.Vencimento = CDate(txtVencimentoPF.Text).ToString("yyyy-MM-dd")
                    t.Prorrogacao = t.Vencimento
                    t.DataMoeda = CDate(txtVencimentoPF.Text).ToString("yyyy-MM-dd")
                    t.CodigoUnidadeDeNegocio = txtUnNegocio.Value
                    t.CodigoEmpresa = EmpresaPF(0)
                    t.EnderecoEmpresa = EmpresaPF(1)
                    t.CodigoEmpresaPagadora = EmpresaPF(0)
                    t.EndEmpresaPagadora = EmpresaPF(1)

                    Dim objEncargo As New [Lib].Negocio.Encargo("LIQUIDOAPAGAR")
                    t.ContaContabilCliente = IIf(objEncargo.ContaCredito = "", "2010103", objEncargo.ContaCredito)
                    t.CodigoCliente = ConveniadoPF(0)
                    t.EndCliente = ConveniadoPF(1)
                    t.ValorDoDocumento = TxtValorFaturaPF.Text
                    t.MoedaValorDoDocumento = CDec(TxtValorFaturaPF.Text) / objCotacao.Indice
                    t.CodigoBancoCliente = txtCodBanco.Text.Trim
                    t.CodigoAgenciaCliente = txtCodAgencia.Text.Trim
                    t.DigitoAgenciaCliente = txtDigitoAgencia.Text.Trim
                    t.ContaCliente = txtConta.Text.Trim
                    t.DigitoContaCliente = txtDigitoConta.Text.Trim
                    t.Historico = "FATURA FRETE NÚMERO " & TxtFaturaPF.Text
                    t.CodigoDestinatario = ConveniadoPF(0)
                    t.EndDestinatario = ConveniadoPF(1)
                    t.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
                    t.UsuarioInclusaoData = DateTime.Now
                    t.CodigoIndexador = 3
                    t.Observacoes = ""
                    t.IUD = "I"
                    t.Agrupado = 3
                    t.SalvarSql(Sqls, True)
                    fatura.ListTituloFatura.Add(New FaturaDeFretexTitulo(fatura) With {.CodigoTitulo = t.Codigo})
                Else
                    Dim t As New [Lib].Negocio.Novo.TituloNovo
                    t.CodigoProvisao = 3
                    t.CodigoTipoPgto = 1
                    t.CodigoSituacao = 1
                    t.ReceberPagar = "P"
                    t.CodigoMoeda = 1
                    t.Movimento = CDate(txtMovimentoPF.Text).ToString("yyyy-MM-dd")
                    t.Vencimento = CDate(txtVencimentoPF.Text).ToString("yyyy-MM-dd")
                    t.Reprogramacao = t.Vencimento
                    t.DataMoeda = CDate(txtVencimentoPF.Text).ToString("yyyy-MM-dd")
                    t.CodigoUnidadeDeNegocio = txtUnNegocio.Value
                    t.CodigoEmpresa = EmpresaPF(0)
                    t.EnderecoEmpresa = EmpresaPF(1)
                    t.CodigoEmpresaRecPag = EmpresaPF(0)
                    t.EndEmpresaRecPag = EmpresaPF(1)

                    Dim objEncargo As New [Lib].Negocio.Encargo("LIQUIDOAPAGAR")
                    t.CodigoClienteRecPag = ConveniadoPF(0)
                    t.EndClienteRecPag = ConveniadoPF(1)

                    'CONTA - VALOR DO DOCUMENTO
                    t.CodigoContaContabilCliFor = t.Empresa.Empresa.CodigoContaFornecedorFrete '(objEncargo.ContaCredito = "", "2010103", objEncargo.ContaCredito)
                    'CONTA - VALOR LÍQUIDO
                    t.CodigoContaContabilRecPag = t.Empresa.Empresa.CodigoContaGrupoBanco
                    'VALOR DO DOCUMENTO
                    t.Valores.EncargoValorDocumento.ValorOficial = CDec(TxtValorFaturaPF.Text)
                    t.Valores.EncargoValorDocumento.ValorMoeda = CDec(TxtValorFaturaPF.Text) / objCotacao.Indice
                    'VALOR LÍQUIDO
                    t.Valores.AtualizaLiquido()

                    t.CodigoBancoCliFor = txtCodBanco.Text.Trim
                    t.CodigoAgenciaCliFor = txtCodAgencia.Text.Trim
                    t.DigitoAgenciaCliFor = txtDigitoAgencia.Text.Trim
                    t.ContaCliFor = txtConta.Text.Trim
                    t.DigitoContaCliFor = txtDigitoConta.Text.Trim
                    t.Historico = "FATURA FRETE NÚMERO " & TxtFaturaPF.Text
                    t.TituloDestinacao.CodigoDestinatario = ConveniadoPF(0)
                    t.TituloDestinacao.EndDestinatario = ConveniadoPF(1)
                    t.CodigoIndexador = 3

                    t.IUD = "I"
                    t.SalvarSql(Sqls, True)
                    fatura.ListTituloFatura.Add(New FaturaDeFretexTitulo(fatura) With {.CodigoTitulo = t.Codigo})
                End If

                fatura.SalvarSql(Sqls)

                If Banco.GravaBanco(Sqls) Then
                    LimparPreFaturas()
                    MsgBox(Me.Page, "Pré-fatura incluída com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Pré-fatura já cadastrada!")
            End If
        End If
    End Sub

    Private Sub SalvarFatura()
        Calcular()
        Dim Fatura As [Lib].Negocio.FaturaDeFrete = Session("Fatura" & HID.Value)
        Dim NfDeFrete As [Lib].Negocio.NotaFiscal = Session("NotasDeFrete" & HID.Value)
        RecuperaItemFF()

        If Fatura Is Nothing OrElse NfDeFrete Is Nothing OrElse ItemFF Is Nothing Then
            MsgBox(Me.Page, "É necessário preencher todos os campos obrigatórios!")
            Exit Sub
        End If

        If Not ItemFF.JaFaturada() Then
            Select Case True
                Case ItemFF.ViaAdiantamento
                    For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In NfDeFrete.Itens(0).Encargos
                        If enc.Codigo.Trim = "BAIXAADTO" Then
                            ItemFF.CodigoEncargo = "BAIXAADTO"
                            ItemFF.ValorLancadoNota = enc.Valor
                            ItemFF.ViaAdiantamento = True
                            Exit For
                        End If
                    Next
                Case ItemFF.ViaAmortizacao
                    For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In NfDeFrete.Itens(0).Encargos
                        If enc.Codigo.Trim = "AMORTIZAADTO" Then
                            ItemFF.CodigoEncargo = "AMORTIZAADTO"
                            ItemFF.ValorLancadoNota = enc.Valor
                            ItemFF.ViaAmortizacao = True
                            Exit For
                        End If
                    Next
                Case ItemFF.ViaCartaFrete
                    For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In NfDeFrete.Itens(0).Encargos
                        If enc.Codigo.Trim = "LIQUIDOAPAGAR" Then
                            ItemFF.CodigoEncargo = "LIQUIDOAPAGAR"
                            ItemFF.ValorLancadoNota = enc.Valor
                            ItemFF.ViaCartaFrete = True
                            Exit For
                        End If
                    Next
            End Select
        End If

        Dim sqls As New ArrayList
        ItemFF.IUD = "I"
        ItemFF.SalvarSql(sqls)

        Fatura.AtualizarVinculos(sqls)

        If Not Banco.GravaBanco(sqls) Then
            MsgBox(Me.Page, Session("ssMessage"))
        End If

        Fatura.Itens.Add(ItemFF)
        LimparParcial()

        Dim ListNfDeFrete As [Lib].Negocio.ListFaturasDeFretesXItens
        ListNfDeFrete = New [Lib].Negocio.ListFaturasDeFretesXItens(Fatura)
        DgComposicaoFatura.DataSource = ListNfDeFrete
        DgComposicaoFatura.DataBind()
        btnEncerrar.Enabled = ListNfDeFrete IsNot Nothing AndAlso ListNfDeFrete.Count > 0

        Dim razao As New [Lib].Negocio.Razao
        razao.ContabilizarFretesNoRazao(NfDeFrete)
        Session("Fatura" & HID.Value) = Fatura
        TabOutros.ActiveTabIndex = 1
    End Sub

    Private Sub Calcular()
        Try
            Dim NfDeFrete As [Lib].Negocio.NotaFiscal = Session("NotasDeFrete" & HID.Value)
            NfDeFrete.Itens(0).CarregandoEncargos = True
            Dim BrutoFrete = CDec(txtValorFrete.Text)
            Dim i As Integer
            While i < dgEncargos.Rows.Count
                Dim j As Integer = 0
                While j < NfDeFrete.Itens(0).Encargos.Count
                    If NfDeFrete.Itens(0).Encargos(j).Codigo = dgEncargos.Rows(i).Cells(1).Text Then
                        If NfDeFrete.Itens(0).Encargos(j).Codigo = "LIQUIDOAPAGAR" Then
                            NfDeFrete.Itens(0).Encargos(j).Valor = BrutoFrete
                        Else
                            If (dgEncargos.Rows(i).Cells(4).Text = "-" Or dgEncargos.Rows(i).Cells(1).Text.Contains("QUEBRAS")) AndAlso
                                CDec(CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Text) > 0 Then
                                BrutoFrete -= CDec(CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Text)
                            ElseIf (dgEncargos.Rows(i).Cells(4).Text = "+" Or dgEncargos.Rows(i).Cells(1).Text.Contains("SOBRAS")) AndAlso
                                CDec(CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Text) > 0 Then
                                BrutoFrete += CDec(CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Text)
                            End If
                            NfDeFrete.Itens(0).Encargos(j).Valor = CDec(CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Text)
                        End If
                        Exit While
                    End If
                    j += 1
                End While
                i += 1
            End While
            'NfDeFrete.Itens(0).Encargos.AtualizaLiquido()
            Session("NotasDeFrete" & HID.Value) = NfDeFrete
            CarregaDgEncargos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregaDgEncargos(Optional ByVal lstEncargos As List(Of NotaFiscalXItemXEncargo) = Nothing, Optional ByVal temComplemento As Boolean = False)
        Try
            Dim aux = False
            RecuperaItemFF()
            Dim NfDeFrete As [Lib].Negocio.NotaFiscal = Session("NotasDeFrete" & HID.Value)
            For Each row As GridViewRow In dgEncargos.Rows
                Dim txt As TextBox = CType(row.FindControl("txtValorEncargo"), TextBox)
                If txt.Enabled Then
                    aux = True
                End If
            Next

            If temComplemento OrElse Funcoes.VerificaPermissao("LiberaAjusteFrete", "ALTERAR") Then
                btnCalcular.Enabled = True
            ElseIf btnCalcular.Enabled = aux AndAlso NfDeFrete.Itens(0).Encargos IsNot Nothing AndAlso NfDeFrete.Itens(0).Encargos.Count > 0 Then
            End If

            If ItemFF IsNot Nothing AndAlso ItemFF.CodigoEncargo = "BAIXAADTO" Then
                dgEncargos.DataSource = NfDeFrete.Itens(0).Encargos.Where(Function(s) s.Codigo = "PRODUTO" OrElse s.Codigo = "ADIANTAMENTO" OrElse s.Codigo = "LIQUIDO")
                dgEncargos.DataBind()

                Dim i As Integer
                While i < dgEncargos.Rows.Count

                    If NfDeFrete.Itens(0).Encargos(i).Encargo.Codigo.Trim.Equals("LIQUIDOAPAGAR") OrElse
                         NfDeFrete.Itens(0).Encargos(i).Encargo.Codigo.Trim.Equals("TAXA CADASTRO") OrElse
                         NfDeFrete.Itens(0).Encargos(i).Encargo.Codigo.Trim.Equals("TARIFA SEGURO") Then
                        CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Enabled = False
                    ElseIf temComplemento OrElse Funcoes.VerificaPermissao("LiberaAjusteFrete", "ALTERAR") Then
                        CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Enabled = True
                    ElseIf CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Enabled = NfDeFrete.Itens(0).Encargos(i).Encargo.Atualizacao Then
                    End If
                    i += 1
                End While
            ElseIf ItemFF IsNot Nothing AndAlso ItemFF.CodigoEncargo = "LIQUIDOAPAGAR" Then
                Dim index As Integer = 0
                If lstEncargos IsNot Nothing AndAlso lstEncargos.Count > 0 Then lstEncargos.ForEach(Function(s)
                                                                                                        index += 1
                                                                                                        s.Sequencia = index
                                                                                                        s.Sinal = "-"
                                                                                                        Return True
                                                                                                    End Function)

                NfDeFrete.Itens(0).Encargos.ForEach(Function(s)
                                                        index += 1
                                                        s.Sequencia = index
                                                        Return True
                                                    End Function)

                If lstEncargos IsNot Nothing AndAlso lstEncargos.Count > 0 Then
                    NfDeFrete.Itens(0).Encargos.AddRange(lstEncargos)
                End If

                dgEncargos.DataSource = NfDeFrete.Itens(0).Encargos.OrderBy(Function(s) s.Sequencia)
                dgEncargos.DataBind()

                Dim i As Integer
                While i < dgEncargos.Rows.Count
                    If NfDeFrete.Itens(0).Encargos(i).Encargo.Codigo.Trim.Equals("LIQUIDOAPAGAR") OrElse
                         NfDeFrete.Itens(0).Encargos(i).Encargo.Codigo.Trim.Equals("TAXA CADASTRO") OrElse
                         NfDeFrete.Itens(0).Encargos(i).Encargo.Codigo.Trim.Equals("TARIFA SEGURO") Then
                        CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Enabled = False
                    ElseIf temComplemento OrElse Funcoes.VerificaPermissao("LiberaAjusteFrete", "ALTERAR") Then
                        CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Enabled = True
                    ElseIf CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Enabled = NfDeFrete.Itens(0).Encargos.OrderBy(Function(s) s.Sequencia)(i).Encargo.Atualizacao Then
                    End If
                    i += 1
                End While
            Else
                dgEncargos.DataSource = NfDeFrete.Itens(0).Encargos
                dgEncargos.DataBind()

                Dim i As Integer
                While i < dgEncargos.Rows.Count
                    If NfDeFrete.Itens(0).Encargos(i).Encargo.Codigo.Trim.Equals("LIQUIDOAPAGAR") OrElse
                         NfDeFrete.Itens(0).Encargos(i).Encargo.Codigo.Trim.Equals("TAXA CADASTRO") OrElse
                         NfDeFrete.Itens(0).Encargos(i).Encargo.Codigo.Trim.Equals("TARIFA SEGURO") Then
                        CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Enabled = False
                    ElseIf temComplemento OrElse Funcoes.VerificaPermissao("LiberaAjusteFrete", "ALTERAR") Then
                        CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Enabled = True
                    ElseIf CType(dgEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Enabled = NfDeFrete.Itens(0).Encargos(i).Encargo.Atualizacao Then
                    End If
                    i += 1
                End While
            End If
            Session("NotasDeFrete" & HID.Value) = NfDeFrete
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ContabilizarCTe(ByVal objNotaFiscal As [Lib].Negocio.NotaFiscal, ByVal objEncargo As [Lib].Negocio.NotaFiscalXItemXEncargo, ByVal LiqDC As String, ByVal LiqConta As String, ByVal LiqValorConvertido As Decimal, ByRef Sqls As ArrayList, ByRef i As Integer)
        Dim sql As String = ""
        Sqls.Add("DELETE r FROM Razao r " & vbCrLf &
                       "INNER JOIN NotasFiscais nf " & vbCrLf &
                               " ON nf.Cliente_Id = r.Cliente_Nf " & vbCrLf &
                               "AND nf.EndCliente_Id = r.EndCliente_Nf " & vbCrLf &
                               "AND nf.EntradaSaida_Id = r.EntradaSaida_Nf " & vbCrLf &
                               "AND nf.Serie_Id = r.Serie_Nf " & vbCrLf &
                               "AND nf.Nota_Id = r.Numero_Nf " & vbCrLf &
                               "AND nf.TipoDeDocumento = 2 " & vbCrLf &
                               "AND nf.TipoDeDocumentoFrete = 2 " & vbCrLf &
                       "WHERE 1=1 " & vbCrLf &
                       "AND r.Cliente_Nf = '" & objNotaFiscal.CodigoCliente & "'  " & vbCrLf &
                       "AND r.EndCliente_Nf = '" & objNotaFiscal.EnderecoCliente & "'  " & vbCrLf &
                       "AND r.EntradaSaida_Nf = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'  " & vbCrLf &
                       "AND r.Serie_Nf = '" & objNotaFiscal.Serie & "'  " & vbCrLf &
                       "AND r.Numero_Nf = '" & objNotaFiscal.Codigo & "' " & vbCrLf &
                       "AND r.Encargo_Nf = '" & objEncargo.Codigo & "'")

        sql = "DECLARE " & vbCrLf &
               " @SequenciaOriginal int" & vbCrLf &
               ",@SequenciaServidor int" & vbCrLf
        sql &= ",@SequenciaTransferenciaCliente int" & vbCrLf &
                  " SET @SequenciaOriginal = (SELECT Sequencia" & vbCrLf &
                  "                             FROM numerador" & vbCrLf &
                  "                            WHERE UPPER(Empresa_id) = '" & HttpContext.Current.Session("ssNomeServidor").ToUpper() & "'" & vbCrLf &
                  "                              AND Numerador_id      = 60)" & vbCrLf &
                  " SET @SequenciaServidor = (SELECT isnull(Max(Sequencia_Id),isnull(@SequenciaOriginal,0))" & vbCrLf &
                  "                             FROM Razao (nolock) " & vbCrLf &
                  "                            WHERE Lote_Id      = " & "21" & vbCrLf &
                  "                              AND Movimento_Id = '" & objNotaFiscal.Movimento.ToSqlDate() & "'" & vbCrLf &
                  "                              AND Sequencia_id between isnull(@SequenciaOriginal,0) and isnull(@SequenciaOriginal,0) + 99999)" & vbCrLf &
                  " SET @SequenciaTransferenciaCliente = (SELECT isnull(Max(Sequencia_Id),isnull(@SequenciaOriginal,0))" & vbCrLf &
                  "                                         FROM razao (nolock) " & vbCrLf &
                  "                                        WHERE Lote_Id      = 11" & vbCrLf &
                  "                                          AND Movimento_Id = '" & objNotaFiscal.Movimento.ToSqlDate() & "'" & vbCrLf &
                  "                                          AND Sequencia_id between isnull(@SequenciaOriginal,0) and isnull(@SequenciaOriginal,0) + 99999)" & vbCrLf

        sql &= "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, DataMoeda, " & vbCrLf &
               "                   DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda," & vbCrLf &
               "                   Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Produto_NF, CFOP_NF, Sequencia_NF, Encargo_NF," & vbCrLf &
               "                   Pedido, PrevistoRealizado, Custo, Produto, Rateado, Deposito, EndDeposito)" & vbCrLf &
               " VALUES ('" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
               ", " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
               ",'" & IIf(LiqDC = "D", objEncargo.ContaDeDebito, objEncargo.ContaDeCredito) & "'" & vbCrLf &
               ",'" & objNotaFiscal.CodigoDeposito & "'" & vbCrLf &
               ", " & objNotaFiscal.EnderecoDeposito & vbCrLf &
               ",'" & objNotaFiscal.Movimento.ToSqlDate() & "'" & vbCrLf &
               ",21" & vbCrLf &
               ", @SequenciaTransferenciaCliente + " & i & vbCrLf &
               ",0" & vbCrLf &
               IIf(objNotaFiscal.CodigoPedido = 0, ", ''", ", '" & objNotaFiscal.Pedido.CodigoUnidadeNegocio & "'") & vbCrLf &
               IIf(objNotaFiscal.CodigoPedido = 0, ", 2", "," & objNotaFiscal.Pedido.CodigoIndexador) & vbCrLf &
               ",'" & objNotaFiscal.Movimento.ToSqlDate() & "'" & vbCrLf

        If LiqDC = "D" Then
            sql &= ",0.0, " & Str(objEncargo.Valor) & ", 0.0," & Str(LiqValorConvertido) & vbCrLf
        Else
            sql &= "," & Str(objEncargo.Valor) & ", 0.0," & Str(LiqValorConvertido) & ",0.0" & vbCrLf
        End If

        Dim Historico As String = "DIF FRETE DE SAIDA E CHEGADA - " & objEncargo.Codigo & " OP " & objNotaFiscal.Itens(0).CodigoOperacao & "-" & objNotaFiscal.Itens(0).CodigoSubOperacao & " NF " & objNotaFiscal.Codigo & "-" & objNotaFiscal.Serie & " Pedido " & objNotaFiscal.CodigoPedido & " - " & objNotaFiscal.Cliente.Nome
        If objNotaFiscal.Itens(0).CodigoOperacao > 68 AndAlso objNotaFiscal.Observacoes.Length > 0 Then
            Historico &= ". " & objNotaFiscal.Observacoes
        End If

        Dim objPlanoDeConta As New [Lib].Negocio.PlanoDeConta("", 0, objEncargo.OperacaoEncargo.CodigoCreditaConta)
        sql &= ",'" & Historico & " '" & vbCrLf &
                  ",'" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                  ", " & objNotaFiscal.EnderecoCliente & vbCrLf &
                  ",'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                  ",'" & objNotaFiscal.Serie & "'" & vbCrLf &
                  ", " & objNotaFiscal.Codigo & vbCrLf &
                  ",'" & objNotaFiscal.Itens(0).CodigoProduto & "'" & vbCrLf &
                  ", " & objNotaFiscal.Itens(0).CFOP & vbCrLf &
                  ", " & objNotaFiscal.Itens(0).Sequencia & vbCrLf &
                  ",'" & objEncargo.Codigo & "'" & vbCrLf &
                  ", " & objNotaFiscal.CodigoPedido.ToSqlNULL & vbCrLf &
                  ",'P'" & vbCrLf &
                  ",'" & objEncargo.CentroDeCusto & "'," & vbCrLf &
                  IIf(objPlanoDeConta.TemProduto, objNotaFiscal.Itens(0).CodigoProduto, "NULL") & "," & vbCrLf &
                  IIf(objNotaFiscal.Itens(0).Rateado, "1", "0") & vbCrLf

        If objNotaFiscal.Itens(0).SubOperacao.EntradaSaida = eEntradaSaida.Saida And objNotaFiscal.Itens(0).SubOperacao.Deposito Then
            sql &= ",'" & objNotaFiscal.CodigoDestino & "'," & objNotaFiscal.EnderecoDestino & ")" & vbCrLf
        Else
            sql &= ",'" & objNotaFiscal.CodigoDeposito & "'," & objNotaFiscal.EnderecoDeposito & ")" & vbCrLf
        End If
        Sqls.Add(sql)
    End Sub

#End Region

End Class