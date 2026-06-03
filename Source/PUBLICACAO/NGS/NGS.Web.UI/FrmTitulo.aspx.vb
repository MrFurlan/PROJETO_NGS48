Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Xml

Partial Class FrmTitulo
    Inherits BasePage

#Region "Variáveis Locais"
    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim ds As DataSet

    Private ObjTitulo As Novo.TituloNovo
    Private ObjTituloAgrupado As Novo.TituloNovo
    Private ObjTituloBordero As Novo.TituloNovo
    Private ObjTitulosAcao As Novo.ListTituloNovo
#End Region

#Region "Eventos"

#Region "Page_Load"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'CarregarXML()
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ContasAReceber", "ACESSAR") Then
                Session.Remove("ObjTitulo" + HID.Value)

                If Funcoes.VerificaPermissao("ContasAReceber", "LIBERAR") Then
                    chkPrevisao.Visible = True
                    chkProvisao.Visible = True
                End If

                PaineisInvisiveis()

                ddl.Carregar(DdlUnidadeDeNegocioEmpresaCliente, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                ddl.Carregar(ddlUnidadeDeNegocioRecPag, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                ddl.Carregar(DdlUnidadeConsultaTitulos, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                ddl.Carregar(ddlUnidadeDeNegocioBaixa, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                'ddl.Carregar(DdlEmpresaRecPag, CarregarDDL.Tabela.ClientesXEmpresas, "", False)
                'ddl.Carregar(ddlEmpresaBaixaTodos, CarregarDDL.Tabela.ClientesXEmpresas, "", True)

                ddl.Carregar(ddlMoeda, CarregarDDL.Tabela.Moeda, "")
                ddl.Carregar(ddlIndexador, CarregarDDL.Tabela.Indexador, "")

                Dim Parametros As New Hashtable
                Parametros.Clear()
                Parametros.Add("listarTudo", "N")

                ddl.Carregar(DdlTiposDeRecebimentos, CarregarDDL.Tabela.TipoDePagamento, "", True, Parametros)
                ddl.Carregar(DdlTiposDeRecebimentosBaixarTodos, CarregarDDL.Tabela.TipoDePagamento, "", True, Parametros)
                ddl.Carregar(DdlProvisoes, CarregarDDL.Tabela.Provisoes, "")
                ddl.Carregar(ddlCarteiraDoTitulo, CarregarDDL.Tabela.CarteiraDoTitulo, "")

                ddl.Carregar(ddlBancoDuplicataCobrancaSimples, CarregarDDL.Tabela.BancosXContaContabil, "", False)
                ddl.Carregar(ddlCarteiraDuplicataCobrancaSimples, CarregarDDL.Tabela.CarteiraDoTitulo, "isnull(C.EmiteDuplicata,0) = 1", False)
                txtDataEnvioBorderoCobrancaDuplicataSimples.Text = Date.Now.ToString("dd/MM/yyyy")
                CarregarLote()
                Limpar()
                MudouRP("R")
                SessaoSalvaTitulo()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        Else
            If Not Request("__EVENTARGUMENT") Is Nothing Then
                'If Request("__EVENTARGUMENT") = "ExcluirTitulo" And (txtRegistro.Text <> "" And txtRegistro.Text <> "0") Then Excluir_Titulo()
            End If
        End If

        AtribuirValoresCampos()
    End Sub

#End Region

#Region "DropDownList"

    Protected Sub DdlUnidadeDeNegocioEmpresaCliente_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlUnidadeDeNegocioEmpresaCliente.SelectedIndexChanged
        Try
            SessaoRecuperaTitulo()
            ddl.Carregar(DdlEmpresaCliente, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocioEmpresaCliente.SelectedValue, True)
            ObjTitulo.CodigoUnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocioRecPag_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeDeNegocioRecPag.SelectedIndexChanged
        Try
            SessaoRecuperaTitulo()
            ddl.Carregar(DdlEmpresaRecPag, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocioRecPag.SelectedValue, True)
            ObjTitulo.CodigoUnidadeDeNegocioRecPag = ddlUnidadeDeNegocioRecPag.SelectedValue
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlProvisoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlProvisoes.SelectedIndexChanged
        Try
            SessaoRecuperaTitulo()

            If String.IsNullOrWhiteSpace(DdlProvisoes.SelectedValue) Then Exit Sub
            If DdlProvisoes.SelectedValue = 1 And ObjTitulo.CodigoProvisao = eProvisao.Previsao AndAlso Not ObjTitulo.ContaContabilCliFor.Adiantamento Then
                MsgBox(Me.Page, "Não é permitido realizar baixa de título em (2) - previsão.")
                DdlProvisoes.SelectedValue = ObjTitulo.CodigoProvisao
                Exit Sub
            End If
            If ObjTitulo.CodigoProvisao = eProvisao.Previsao AndAlso ObjTitulo.Pedido IsNot Nothing AndAlso Not ObjTitulo.ContaContabilCliFor.Adiantamento Then
                MsgBox(Me.Page, "Não é permitido alterar previsão associada ao pedido: " & ObjTitulo.CodigoPedido)
                DdlProvisoes.SelectedValue = ObjTitulo.CodigoProvisao
                Exit Sub
            End If
            If DdlProvisoes.SelectedValue = 4 Then
                MsgBox(Me.Page, "Compensaçoes são feitas automaticamente.")
                DdlProvisoes.SelectedValue = ObjTitulo.CodigoProvisao
                Exit Sub
            End If
            If ObjTitulo.ReceberPagar <> "C" And ObjTitulo.CodigoCliFor.Length = 0 And ObjTitulo.CodigoContaContabilCliFor.Length <> 9 Then
                MsgBox(Me.Page, "Informe um cliente antes de mudar o status do titulo.")
                Exit Sub
            End If
            If ObjTitulo.CodigoProvisao = eProvisao.Baixa AndAlso Not ObjTitulo.Adiantamento Is Nothing AndAlso ObjTitulo.Adiantamento.Baixas.Count > 0 Then
                MsgBox(Me.Page, "O adiantamento vinculado a este titulo ja contem baixas.")
                Exit Sub
            End If

            ObjTitulo.CodigoProvisao = DdlProvisoes.SelectedIndex

            If ObjTitulo.CodigoProvisao = eProvisao.Baixa Then
                ObjTitulo.DataBaixa = Date.Now
                txtDataBaixa.Text = ObjTitulo.DataBaixa.ToString("dd/MM/yyyy")
            End If

            SessaoSalvaTitulo()
            DatadaBaixa()
            If Not ObjTitulo.CliFor Is Nothing AndAlso ObjTitulo.ReceberPagar <> "C" AndAlso (ObjTitulo.DesdobrarFornecedor Is Nothing OrElse ObjTitulo.DesdobrarFornecedor.CodigoCliente.Length = 0) AndAlso ObjTitulo.CodigoProvisao = eProvisao.Baixa Then
                If ObjTitulo.CliFor.DesdobrarFornecedor Then
                    Dim strJavaScript As String = ""
                    strJavaScript = "var x = (screen.height / 2) - 70; "
                    strJavaScript += "var y = (screen.width / 2) - 400; "
                    strJavaScript += "window.open(""DestinoContabil.aspx?tipo=R&hid=" & HID.Value & """, """", ""resizable=no, menubar=no, scrollbars=yes, width=800, height=140, top="" + x + "", left="" + y + """");"
                    ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), "DestinoContabil", strJavaScript, True)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlMoeda_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaTitulo()
            ObjTitulo.CodigoMoeda = ddlMoeda.SelectedValue
            SessaoSalvaTitulo()
            AtualizaValoresNoForm(ObjTitulo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlEmpresaCliente_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaTitulo()
            If DdlEmpresaCliente.SelectedIndex = 0 Then
                ObjTitulo.CodigoEmpresa = ""
                ObjTitulo.EnderecoEmpresa = 0
            Else
                Dim emp As String() = DdlEmpresaCliente.SelectedValue.Split("-")
                ObjTitulo.CodigoEmpresa = emp(0)
                ObjTitulo.EnderecoEmpresa = emp(1)
            End If
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlTiposDeRecebimentos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaTitulo()
            If DdlTiposDeRecebimentos.SelectedValue = 0 Then
                ObjTitulo.CodigoTipoPgto = 0
            Else
                ObjTitulo.CodigoTipoPgto = DdlTiposDeRecebimentos.SelectedValue
                AdicionaContaTEDDOC(ObjTitulo)
            End If
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlCarteiraDoTitulo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaTitulo()
            ObjTitulo.CodigoCarteiraDoTitulo = ddlCarteiraDoTitulo.SelectedValue
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlIndexador_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaTitulo()
            ObjTitulo.CodigoIndexador = ddlIndexador.SelectedIndex
            SessaoSalvaTitulo()
            AtualizaValoresNoForm(ObjTitulo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlContaContabilEmpresaRecPag_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlContaContabilEmpresaRecPag.SelectedIndexChanged
        Try
            If Not String.IsNullOrWhiteSpace(ddlContaContabilEmpresaRecPag.SelectedValue) Then
                SessaoRecuperaTitulo()
                ObjTitulo.CodigoContaContabilRecPag = ddlContaContabilEmpresaRecPag.SelectedValue
                SelecionarContaContabilRecPag()
                SessaoSalvaTitulo()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlEmpresaRecPag_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlEmpresaRecPag.SelectedIndexChanged
        Try
            SessaoRecuperaTitulo()
            If DdlEmpresaRecPag.SelectedIndex = 0 Then
                ObjTitulo.CodigoEmpresaRecPag = ""
                ObjTitulo.EndEmpresaRecPag = 0
            Else
                Dim emp As String() = DdlEmpresaRecPag.SelectedValue.Split("-")
                ObjTitulo.CodigoEmpresaRecPag = emp(0)
                ObjTitulo.EndEmpresaRecPag = emp(1)
                CarregarContaContabilEmpresaRecPag(ObjTitulo.ReceberPagar)
            End If
            SessaoSalvaTitulo()
            DdlTiposDeRecebimentos.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlCentroDeCusto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim ddlgrid As DropDownList = CType(sender, DropDownList)
            Dim row As GridViewRow = CType(ddlgrid.NamingContainer, GridViewRow)
            SessaoRecuperaTitulo()
            ObjTitulo.Valores(row.RowIndex).CodigoCentroDeCusto = ddlgrid.SelectedValue
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlUnidadeConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlUnidadeConsultaTitulos.SelectedIndexChanged
        Try
            ddl.Carregar(DdlEmpresaConsultaTitulos, CarregarDDL.Tabela.Empresas, DdlUnidadeConsultaTitulos.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocioBaixa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUnidadeDeNegocioBaixa.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresaBaixaTodos, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocioBaixa.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEmpresaBaixaTodos_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlEmpresaBaixaTodos.SelectedIndexChanged
        Try
            CarregarContaContabilEmpresaRecPagBaixarTodos(IIf(chkReceber.Checked, "R", "P"))
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Button"

#Region "Botões Principais"

#Region "Cadastro Titulo"

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("ContasAReceber", "GRAVAR") Then
                SessaoRecuperaTitulo()
                'Recupera os valores da tela para salvar o título
                ObjTitulo.CodigoDeBarras = txtCodigoDeBarras.Text
                ObjTitulo.CodigoDeBarrasDigitado = CkbCodigoDeBarras.Checked
                ObjTitulo.CodigoDeBarrasPreImpresso = ckPreImpresso.Checked
                ObjTitulo.Observacoes = txtObservacoes.Text
                ObjTitulo.Historico = txtHistorico.Text

                If RdContabil.Checked Then
                    ObjTitulo.Lote = ddlLote.SelectedValue 'Lote Contábil
                    ObjTitulo.Sequencia = IIf(String.IsNullOrWhiteSpace(ddlSequencia.SelectedValue), 1, ddlSequencia.SelectedValue)
                Else
                    ObjTitulo.Lote = 70 'lote para contas a pagar e receber
                    ObjTitulo.Sequencia = 0
                End If

                If Not ObjTitulo.Bordero Is Nothing AndAlso ObjTitulo.Bordero.TitulosDoBordero.Count > 0 Then
                    ObjTitulo.Bordero.Contrato = txtBordero.Text
                End If
                If ObjTitulo.isBaixaAdiantamento AndAlso ObjTitulo.AdiantamentosAbertos.Count > 1 Then
                    ObjTitulo.CodigoPedido = ObjTitulo.AdiantamentosAbertos(0).Titulo.CodigoPedido
                End If

                'Realiza as baixas de adiantamento
                If gridAdiantamentosDisponiveis.Visible Then
                    For Each row As GridViewRow In gridAdiantamentosDisponiveis.Rows
                        Dim vlrBaixa As Decimal = CType(row.FindControl("txtVlrBaixa"), TextBox).Text
                        If vlrBaixa > 0 Then
                            ObjTitulo.AdiantamentosAbertos(row.RowIndex).VlrBaixa = vlrBaixa
                            ObjTitulo.AdiantamentosAbertos.Titulo = ObjTitulo
                        End If
                    Next
                End If

                'Adiciona e Remove o agrupamento
                If rdAdTituloAgrupamento.Checked Then
                    SessaoRecuperaTituloAgrupado()
                    Dim sqls As New ArrayList
                    ObjTituloAgrupado.SalvarSql(sqls)
                    If Banco.GravaBanco(sqls) Then
                        MsgBox(Me.Page, ObjTituloAgrupado.Observacoes & " Alterado com Sucesso.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, "Erro ao alterar agrupamento.")
                    End If
                    SessaoSalvaTituloAgrupado()
                End If


                If Not ValidaCampos() Then Exit Sub
                If Not String.IsNullOrWhiteSpace(txtRegistro.Text) Then ObjTitulo.IUD = "U" 'Atualiza o Titulo
                If ObjTitulo.TitulosAgrupados.Count > 0 Then
                    AgruparValoresTituloXContaContabil(ObjTitulo)
                End If

                'kitio coloquei o not afrente do ObjTitulo.Pedido.Bloquear()
                If FinanceiroNovo AndAlso ObjTitulo.Pedido IsNot Nothing AndAlso Not ObjTitulo.Pedido.Bloquear() Then
                    MsgBox(Me.Page, "O pedido " & ObjTitulo.Pedido.Codigo & " foi bloqueado por outro usuário, por favor recarregue o registro!")
                    Exit Sub
                End If
                If ObjTitulo.Salvar() Then
                    'Emitir Recibo.
                    If chkEmitirRecibo.Checked Then EmitirRecibo(ObjTitulo.Codigo.ToString)
                    MsgBox(Me.Page, IIf(ObjTitulo.MsgControle.Length = 0, "Titulo: " & ObjTitulo.Codigo & IIf(Not String.IsNullOrWhiteSpace(txtRegistro.Text), " Atualizado com sucesso.", " Cadastrado com Sucesso."), ObjTitulo.MsgControle), eTitulo.Sucess)
                    ObjTitulo.Codigo = 0 'Novo Titulo Contábil
                    Limpar()
                    LimparConsulta()
                Else
                    MsgBox(Me.Page, "Erro ao salvar o titulo.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("ContasAReceber", "LEITURA") Then
                Dim Registro As String = txtRegistro.Text
                Dim Bordero As String = txtBordero.Text
                Dim Empresa As String = ""

                If IsNumeric(txtBordero.Text) AndAlso (DdlEmpresaCliente.SelectedIndex = 0 Or DdlEmpresaCliente.SelectedIndex = -1) Then
                    MsgBox(Me.Page, "Para carregar o titulo pelo numero do bordero, seleciona a empresa do bordero.")
                    Exit Sub
                Else
                    Empresa = DdlEmpresaCliente.SelectedValue.Split("-")(0)
                End If
                Limpar()
                If IsNumeric(Bordero) Then
                    ObjTitulo = New Novo.TituloNovo(0, Empresa, Bordero)
                Else
                    ObjTitulo = New Novo.TituloNovo(Registro)
                End If
                If ObjTitulo.Codigo = 0 Then
                    MsgBox(Me.Page, "Título não existe.")
                    Exit Sub
                End If
                ObjTitulo.IUD = "U"
                CarregarFormularioComAClasse(ObjTitulo)
                SessaoSalvaTitulo()
                lnkExcluir.Parent.Visible = True
                LimparConsulta()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            chkManterLancamento.Checked = False
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkDuplicata_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkDuplicata.Click
        Try
            SessaoRecuperaTitulo()
            If ObjTitulo.IUD = "U" Then
                Novo.ImpressaoDuplicata.ExibirImpressao(Me, ObjTitulo, 1)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnImprimirBordero_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnImprimirBordero.Click
        Try
            Dim Bordero As New Novo.Bordero
            Dim rpt As New ReportDocument()
            Dim UrlArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
            Dim NomeArquivo As String = HttpContext.Current.Server.MapPath(UrlArquivo)

            rpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/Cr_Bordero.rpt")
            rpt.Load(rpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)
            rpt.SetDataSource(Bordero.ImprimirBordero())

            If Dir(NomeArquivo).Length > 0 Then
                Kill(NomeArquivo)
            End If

            Try
                rpt.ExportToDisk(ExportFormatType.PortableDocFormat, NomeArquivo)
                If (IO.File.Exists(NomeArquivo)) Then
                    Funcoes.AbrirArquivo(Page, UrlArquivo)
                End If
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Finally
                rpt.Close()
                rpt.Dispose()
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
        'SessaoRecuperaTitulo()
        'Dim strJavaScript As String = String.Empty
        'Novo.ImpressaoBordero.ExibirImpressao(Me.Page, ObjTitulo.CodigoEmpresa, ObjTitulo.Bordero.CodigoBordero, Not chkImprimirDuplicatas.Checked, strJavaScript)
        'If (chkImprimirDuplicatas.Checked) Then
        '    Dim lst As New Novo.ListTitulo
        '    ObjTitulo.Bordero.TitulosDoBordero.Select(Function(s) s.Titulo).ToList().ForEach(Function(x)
        '                                                                                         lst.Add(x)
        '                                                                                         Return True
        '                                                                                     End Function)
        '    Novo.ImpressaoDuplicata.ExibirImpressao(Me.Page, lst, 1, Not chkImprimirDuplicatas.Checked, strJavaScript)
        '    Funcoes.OpenManyPDF(Me.Page, strJavaScript)
        'End If
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            SessaoRecuperaTitulo()

            If ObjTitulo.ContaContabilCliFor.Adiantamento Then
                If ObjTitulo.Baixas_AdiantamentoEfetuadas.Count > 0 Then
                    MsgBox(Me.Page, "Não foi possível excluir, adiantamento com baixas realizadas:" & ObjTitulo.RegistroMestre)
                    Exit Sub
                End If
                ObjTitulo.IUD = "C"
            Else
                If ObjTitulo.RegistroMestre > 0 AndAlso ObjTitulo.RegistroMestre <> ObjTitulo.Codigo Then
                    MsgBox(Me.Page, "Não foi possível excluir, título agrupado ao título:" & ObjTitulo.RegistroMestre)
                    Exit Sub
                ElseIf ObjTitulo.CodigoProvisao = eProvisao.Baixa Then
                    MsgBox(Me.Page, "Não foi possível excluir, título baixado.")
                    Exit Sub
                ElseIf ObjTitulo.CodigoPedido > 0 Then
                    MsgBox(Me.Page, "Não foi possível excluir, título com pedido.")
                    Exit Sub
                ElseIf ObjTitulo.NotaTitulo IsNot Nothing AndAlso _
                    ObjTitulo.NotaTitulo.NotaFiscal IsNot Nothing AndAlso _
                    ObjTitulo.NotaTitulo.NotaFiscal.Codigo > 0 Then
                    MsgBox(Me.Page, "Não é possível excluir Titulo associado a Nota Fiscal")
                    Exit Sub
                End If
                ObjTitulo.IUD = "D"
            End If
            If ObjTitulo.Salvar Then
                MsgBox(Me.Page, "Título excluído com Sucesso.", eTitulo.Sucess)
                Limpar()
            Else
                MsgBox(Me.Page, "Erro ao excluir Título.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRecibo_Click(sender As Object, e As EventArgs) Handles lnkRecibo.Click
        Try
            EmitirRecibo(txtRegistro.Text)
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível emitir recibo!")
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "FrmTitulo")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "Consulta"

    Protected Sub lnkConsultarAbaConsulta_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultarAbaConsulta.Click
        Try
            If Not ValidaConsulta() Then Exit Sub
            GridConsultaTitulos.DataSource = Nothing
            GridConsultaTitulos.DataBind()
            Dim TemRegistro As Boolean
            If rdConsultarBordero.Checked Then
                TemRegistro = TitulosConsultaBordero()
            Else
                TemRegistro = TitulosConsulta()
            End If

            If (rdAgrupar.Checked AndAlso GridConsultaTitulos.Rows.Count > 0) Then
                Dim chkSelectAll As CheckBox = CType(GridConsultaTitulos.HeaderRow.FindControl("chkSelectAll"), CheckBox)
                chkSelectAll.Visible = False
            End If

            If TemRegistro Then
                divAcoes.Visible = True
                divParametros.Visible = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparAbaConsulta_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparAbaConsulta.Click
        Try
            LimparConsulta()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAcao_Click(sender As Object, e As EventArgs) Handles lnkAcao.Click
        Try
            If (divAcaoAgrupar.Visible) Then
                Agrupar()
            ElseIf (divAcaoBaixar.Visible) Then
                Baixar()
            ElseIf (divAcaoReprogramacao.Visible) Then
                Reprogramar()
            ElseIf (divAcaoAdicionar.Visible) Then
                Adicionar()
            ElseIf (divAcaoDuplicatas.Visible) Then
                Bordero()
            ElseIf (divAcaoDuplicatasCobrancaSimples.Visible) Then
                GerarBorderoCobrancaSimples()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkSlip_Click(sender As Object, e As EventArgs) Handles lnkSlip.Click
        Try
            EmitirSplip()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#End Region

    Protected Sub cmdPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdPedido.Click
        Try
            SessaoRecuperaTitulo()
            If ObjTitulo.CodigoEmpresa.Length = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.")
            ElseIf ObjTitulo.CodigoCliFor.Length = 0 Then
                MsgBox(Me.Page, "Cliente não foi selecionado.")
                'ElseIf ObjTitulo.Valores.EncargoValorDocumento.ValorOficial = 0 Or ObjTitulo.Valores.EncargoValorLiquido.ValorOficial = 0 Then
                '    Popup.Mensagem(Me, "Valor Do Documento em R$ ou U$ e Valor Liquido para Pgto são obrigatórios ...")
            Else
                ViewState.Add("campo", CType(sender, Button).ID)
                HttpContext.Current.Session("ssCampo" & HID.Value) = "Pedidos"
                Dim parametros As New Dictionary(Of String, Object)()
                parametros("unidade") = ObjTitulo.CodigoUnidadeDeNegocio
                parametros("empresa") = ObjTitulo.CodigoEmpresa
                parametros("enderecoEmpresa") = ObjTitulo.EnderecoEmpresa
                parametros("cliente") = ObjTitulo.CodigoCliFor
                parametros("enderecoCliente") = ObjTitulo.EnderecoCliFor
                parametros("situacao") = eSituacao.Normal
                If ObjTitulo.ContaContabilCliFor.Adiantamento AndAlso ObjTitulo.TituloOriginal Is Nothing Then
                    parametros("ContaContabil") = ObjTitulo.ContaContabilCliFor.Conta
                End If
                Popup.ConsultaDePedidos(Me.Page, "objPedidoTitulo" & HID.Value.ToString)
                ucConsultaPedidos.BindGridView(parametros)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnClientesTitulo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClientesTitulo.Click
        Try
            SessaoRecuperaTitulo()
            If ObjTitulo.ReceberPagar = "C" And ObjTitulo.CodigoContaContabilCliFor.Length = 0 Then
                MsgBox(Me.Page, "Nos lancamentos contabeis a selecao da conta define se o lancamento tera ou nao cliente, selecione primeiro a conta contabil do lancamento.")
                Exit Sub
            End If
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCR" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultaClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultaClientes.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteConsulta" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnClienteLancamentoContabil_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.SetarHID(HID.Value)
            Popup.ConsultaDeClientes(Me.Page, "objClienteLC" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnDadosBancarios_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaTitulo()
            If ObjTitulo.CodigoProvisao = eProvisao.Baixa AndAlso ObjTitulo.Codigo > 0 Then Exit Sub
            If ObjTitulo.CodigoCliFor.Length = 0 Then
                MsgBox(Me.Page, "Selecione um cliente para continuar.")
                Exit Sub
            End If
            ucConsultaDadosBancarios.SetarHID(HID.Value)
            ucConsultaDadosBancarios.CarregaGrid(ObjTitulo.CodigoCliFor, ObjTitulo.EnderecoCliFor)
            Popup.ConsultaDeDadosBancarios(Me.Page, "objBancoFRMTIT" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnAdicionarConta_Click(ByVal sender As Object, ByVal e As EventArgs) Handles BtnAdicionarConta.Click
        Try
            SessaoRecuperaTitulo()
            Session.Remove("EncargosPlanoDeContas" & HID.Value)
            'If (ObjTitulo.ContaContabilCliFor IsNot Nothing AndAlso ObjTitulo.ContaContabilCliFor.EncargosPlanoDeContas IsNot Nothing) Then
            ucConsultaEncargosPlanoDeContas.BindGridView(ObjTitulo.ContaContabilCliFor.EncargosPlanoDeContas)
            Session("EncargosPlanoDeContas" & HID.Value) = ObjTitulo.ContaContabilCliFor.EncargosPlanoDeContas
            Popup.ConsultaDeEncargosPlanoDeContas(Me.Page, "objEncargosPlanoDeContas" & HID.Value, "btnSelecionar")
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnDuplicataGrid_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim btn As Button = CType(sender, Button)
            Dim row As GridViewRow = CType(btn.NamingContainer, GridViewRow)
            Dim tit As Novo.TituloNovo

            tit = New Novo.TituloNovo(GridConsultaTitulos.Rows(row.RowIndex).Cells(0).Text)

            Novo.ImpressaoDuplicata.ExibirImpressao(Me, tit, 1)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCalcular_Click(sender As Object, e As EventArgs) Handles btnCalcular.Click
        Try
            SessaoRecuperaTitulo()
            AtualizaTituloXContaContabil(ObjTitulo)
            'AgruparValoresTituloXContaContabil(ObjTitulo)
            SessaoSalvaTitulo()
            AtualizaValoresNoForm(ObjTitulo)
            AtualizaTituloXContaContabil(ObjTitulo, False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnDepositoCliente_Click(sender As Object, e As EventArgs) Handles btnDepositoCliente.Click
        Try
            ucConsultaClientes.SetarTipoCliente(eTipoCliente.Depositos)
            Popup.ConsultaDeClientes(Me.Page, "objDepositoCliFor" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnDepositoRecPag_Click(sender As Object, e As EventArgs) Handles btnDepositoRecPag.Click
        Try
            ucConsultaClientes.SetarTipoCliente(eTipoCliente.Depositos)
            Popup.ConsultaDeClientes(Me.Page, "objDepositoRecPag" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnRecomprar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRecomprar.Click
        Try
            'Transformando o Titulo em um Titulo de Compra
            SessaoRecuperaTitulo()
            ObjTitulo.Codigo = 0
            ObjTitulo.IUD = "I"
            ObjTitulo.ReceberPagar = "P"
            ObjTitulo.Observacoes &= " Ref. a Recompra do Titulo " & ObjTitulo.TituloOriginal.Codigo
            ObjTitulo.CodigoCliFor = ObjTitulo.UltimoBordero.Bordero.TituloDoBordero.CodigoCliFor
            ObjTitulo.EnderecoCliFor = ObjTitulo.UltimoBordero.Bordero.TituloDoBordero.EnderecoCliFor
            ObjTitulo.CodigoContaContabilCliFor = ObjTitulo.UltimoBordero.Bordero.TituloDoBordero.CodigoContaContabilCliFor
            ObjTitulo.CodigoEmpresaRecPag = ObjTitulo.UltimoBordero.Bordero.TituloDoBordero.CodigoEmpresaRecPag
            ObjTitulo.EndEmpresaRecPag = ObjTitulo.UltimoBordero.Bordero.TituloDoBordero.EndEmpresaRecPag
            ObjTitulo.CodigoContaContabilRecPag = ObjTitulo.UltimoBordero.Bordero.TituloDoBordero.CodigoContaContabilRecPag
            ObjTitulo.CodigoCarteiraDoTitulo = ObjTitulo.UltimoBordero.Bordero.TituloDoBordero.CodigoCarteiraDoTitulo
            ObjTitulo.DataBaixa = Date.Now
            ObjTitulo.DataMoeda = Date.Now
            ObjTitulo.Reprogramacao = Date.Now
            ObjTitulo.Vencimento = Date.Now
            ObjTitulo.CodigoProvisao = eProvisao.Baixa

            DdlProvisoes.Enabled = False

            ObjTitulo.TituloBorderos = Nothing
            Dim TxB As New Novo.BorderoXTitulo()
            TxB.CodigoBordero = ObjTitulo.UltimoBordero.CodigoBordero
            TxB.Bordero = ObjTitulo.UltimoBordero.Bordero
            ObjTitulo.TituloBorderos.Add(TxB)
            ObjTitulo.TituloOriginal.UltimoBordero.IUD = "U"
            ObjTitulo.TituloOriginal.UltimoBordero.CodigoTituloRecompra = -1
            If ObjTitulo.Valores.Count = 0 Then 'So pra instanciar a lista de valores
            End If

            CarregarFormularioComAClasse(ObjTitulo)
            SessaoSalvaTitulo()
            lnkNovo.Enabled = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProduto_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaProduto.Limpar()
            ucConsultaProduto.SetarHID(HID.Value)
            Session("Where" & HID.Value) = "Situacao = 1"
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoPRD" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProdutoExcluir_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaTitulo()
            ObjTitulo.CodigoProduto = ""
            ObjTitulo.Quantidade = 0
            txtProduto.Text = ""
            txtQuantidade.Text = ""
            txtQuantidade.Enabled = False
            divPedido.Visible = True
            divPedidoRecPag.Visible = True
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPedidoRecPag_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaTitulo()
            If ObjTitulo.CodigoEmpresaRecPag.Length = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.")
            ElseIf ObjTitulo.CodigoClienteRecPag.Length = 0 Then
                MsgBox(Me.Page, "Cliente não foi selecionado.")
            Else
                'HttpContext.Current.Session("ssCampo") = "Pedidos"
                'Dim strJavaScript As String = ""
                'strJavaScript = "var x = (screen.height / 2) - 200; "
                'strJavaScript &= "var y = (screen.width / 2) - 400; "
                'strJavaScript &= "window.open(""ConsultaPedidos.aspx?url=ContasAReceber&tipo=RECPAG"
                'strJavaScript &= "&ue=" & ObjTitulo.CodigoUnidadeDeNegocio
                'strJavaScript &= "&emp=" & ObjTitulo.CodigoEmpresaRecPag
                'strJavaScript &= "&ende=" & ObjTitulo.EndEmpresaRecPag
                'If ObjTitulo.CodigoClienteRecPag.Length > 0 Then
                '    strJavaScript &= "&cli=" & ObjTitulo.CodigoClienteRecPag
                '    strJavaScript &= "&endc=" & ObjTitulo.EndClienteRecPag
                'End If
                'strJavaScript &= """, """", ""resizable=no, menubar=no, scrollbars=Yes, width=800, height=400, top="" + x + "", left="" + y + """");"
                'ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), "BuscaPedido", strJavaScript, True)
                SessaoRecuperaTitulo()
                HttpContext.Current.Session("ssCampo" & HID.Value) = "Pedidos"
                ViewState.Add("campo", CType(sender, Button).ID)
                ucConsultaPedidos.SetarHID(HID.Value)
                Dim parameters As New Dictionary(Of String, Object)
                parameters("empresa") = ObjTitulo.CodigoEmpresaRecPag
                parameters("enderecoEmpresa") = ObjTitulo.EndEmpresaRecPag
                parameters("cliente") = ObjTitulo.CodigoClienteRecPag
                parameters("enderecoCliente") = ObjTitulo.EndClienteRecPag
                If Not String.IsNullOrWhiteSpace(txtProduto.Text) Then
                    parameters("produto") = txtProduto.Text.Split("-").ToArray(0).Trim
                End If
                ucConsultaPedidos.BindGridView(parameters)
                Popup.ConsultaDePedidos(Me.Page, "objPedidoTitulo" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPlanoDeContas_Click(sender As Object, e As EventArgs) Handles btnPlanoDeContas.Click
        Try
            Dim btn As Button = CType(sender, Button)
            ViewState.Add("campo", btn.ID)
            ucConsultaPlanoDeContas.Limpar()
            Dim strReceberPagar As String = String.Empty
            If RdPagar.Checked Then strReceberPagar = "P"
            If RdReceber.Checked Then strReceberPagar = "R"
            If RdContabil.Checked Then
                ucConsultaPlanoDeContas.BindGridView(True)
            Else
                ucConsultaPlanoDeContas.BindGridViewTemEncargo(True, strReceberPagar)
            End If

            SessaoRecuperaTitulo()
            If ObjTitulo.CodigoPedido > 0 Then
                If ObjTitulo.Pedido.Titulos.Where(Function(s) s.CodigoProvisao = eProvisao.Provisao).Count > 0 Then
                    MsgBox(Me.Page, "Não é permitido fazer adiantamento em pedido que tenha títulos em provisão.")
                    Exit Sub
                End If
                ucConsultaPlanoDeContas.CarregarContaAdiantamentoSubOperacao(ObjTitulo.Pedido.SubOperacao.CodigoContaAdiantamento)
            End If
            Popup.ConsultaDePlanoDeContas(Me, "objTituloPDC" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "TextBox"

    Protected Sub txtRegistro_TextChanged(sender As Object, e As EventArgs) Handles txtRegistro.TextChanged
        Try
            If IsNumeric(txtRegistro.Text) Then
                SessaoRecuperaTitulo()
                ObjTitulo = New Novo.TituloNovo(txtRegistro.Text)
                If (ObjTitulo.Codigo > 0 AndAlso ObjTitulo.CodigoSituacao = eSituacao.Normal) Then
                    ObjTitulo.IUD = "U"
                    CarregarFormularioComAClasse(ObjTitulo)
                    SessaoSalvaTitulo()
                    LimparConsulta()
                Else
                    MsgBox(Me.Page, "Título não encontrado.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtDataBaixa_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaTitulo()
            ObjTitulo.DataBaixa = CDate(txtDataBaixa.Text)
            SessaoSalvaTitulo()
            AtualizaValoresNoForm(ObjTitulo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtProrrogacao_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not IsDate(txtProrrogacao.Text) Then Exit Sub
            SessaoRecuperaTitulo()
            ObjTitulo.Reprogramacao = CDate(txtProrrogacao.Text)
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtCotacao_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not IsNumeric(txtCotacao) Then
                MsgBox(Me.Page, "Informe um valor valido para o indice do titulo.")
                Exit Sub
            End If
            SessaoRecuperaTitulo()
            ObjTitulo.CodigoIndexador = 99
            ObjTitulo.IndiceTitulo = CDec(txtCotacao.Text)
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtValorOficial_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim txtValorOficial As TextBox = CType(sender, TextBox)
            Dim row As GridViewRow = CType(txtValorOficial.NamingContainer, GridViewRow)

            If Not IsNumeric(txtValorOficial.Text) OrElse CDec(txtValorOficial.Text) < 0 Then txtValorOficial.Text = "0,00"

            SessaoRecuperaTitulo()

            If ObjTitulo.CodigoPedido > 0 AndAlso ObjTitulo.CodigoContaContabilCliFor = ObjTitulo.Pedido.SubOperacao.CodigoGrupoContas AndAlso Not ObjTitulo.TituloOriginal Is Nothing AndAlso CDec(txtValorOficial.Text) > ObjTitulo.TituloOriginal.Valores.EncargoValorDocumento.ValorOficial Then
                MsgBox(Me.Page, "Não é permitido aumentar o valor do capital de um titulo vinculado a um pedido com a conta contabil do pedido.")
                AtualizaValoresNoForm(ObjTitulo)
                Exit Sub
            End If
            'Remove a ContaXContabil do titulo em caso de alteração do valor
            If ObjTitulo.TitulosAgrupados.Count > 0 Then RemoverContaXContabilAgrupamento(ObjTitulo)

            If ObjTitulo.isBaixaAdiantamento AndAlso ObjTitulo.AdiantamentosAbertos.Count > 0 Then
                ObjTitulo.AdiantamentosAbertos.Titulo = ObjTitulo
                If ObjTitulo.AdiantamentosAbertos.ValorTotalDisponivelParaBaixa < CDec(txtValorOficial.Text) Then
                    MsgBox(Me.Page, "O Valor Maximo de Adiantamentos abertos é de " & Format(ObjTitulo.AdiantamentosAbertos.ValorTotalDisponivelParaBaixa, "N2"))
                    ObjTitulo.Valores(row.RowIndex).ValorOficial = ObjTitulo.AdiantamentosAbertos.ValorTotalDisponivelParaBaixa
                Else
                    ObjTitulo.Valores(row.RowIndex).ValorOficial = CDec(txtValorOficial.Text)
                End If
                'ObjTitulo.AdiantamentosAbertos.DistribuirValorParaBaixarAdiantamentos(IIf(ObjTitulo.Moeda.Classificacao = eTiposMoeda.Oficial, ObjTitulo.Valores.EncargoValorLiquido.ValorOficial, ObjTitulo.Valores.EncargoValorLiquido.ValorMoeda), ObjTitulo.Moeda.Classificacao, True)

                If String.IsNullOrWhiteSpace(txtPedido.Text) Then
                    gridAdiantamentosDisponiveis.DataSource = ObjTitulo.AdiantamentosAbertos.ToArray
                Else
                    gridAdiantamentosDisponiveis.DataSource = ObjTitulo.AdiantamentosAbertos.Where(Function(A) A.Titulo.CodigoPedido = txtPedido.Text)
                End If

                gridAdiantamentosDisponiveis.DataBind()
            Else
                ObjTitulo.Valores(row.RowIndex).ValorOficial = CDec(txtValorOficial.Text)
            End If

            'Rateio dos valores do titulo mestre para os filhos
            'If ObjTitulo.TitulosAgrupados.Count > 0 Then AgruparValoresTituloXContaContabil(ObjTitulo)

            SessaoSalvaTitulo()

            AtualizaValoresNoForm(ObjTitulo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtValorMoeda_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim txtValorMoeda As TextBox = CType(sender, TextBox)
            Dim row As GridViewRow = CType(txtValorMoeda.NamingContainer, GridViewRow)

            If Not IsNumeric(txtValorMoeda.Text) OrElse CDec(txtValorMoeda.Text) < 0 Then txtValorMoeda.Text = "0,00"

            SessaoRecuperaTitulo()

            If ObjTitulo.CodigoPedido > 0 AndAlso ObjTitulo.CodigoContaContabilCliFor = ObjTitulo.Pedido.SubOperacao.CodigoGrupoContas AndAlso Not ObjTitulo.TituloOriginal Is Nothing AndAlso CDec(txtValorMoeda.Text) > ObjTitulo.TituloOriginal.Valores.EncargoValorDocumento.ValorMoeda Then
                MsgBox(Me.Page, "Não é permitido aumentar o valor do capital de um titulo vinculado a um pedido com a conta contabil do pedido.")
                AtualizaValoresNoForm(ObjTitulo)
                Exit Sub
            End If

            If ObjTitulo.isBaixaAdiantamento Then
                If ObjTitulo.AdiantamentosAbertos.ValorTotalDisponivelParaBaixa < CDec(txtValorMoeda.Text) Then
                    MsgBox(Me.Page, "O Valor Maximo de Adiantamentos abertos é de " & Format(ObjTitulo.AdiantamentosAbertos.ValorTotalDisponivelParaBaixa, "N2"))
                    ObjTitulo.Valores(row.RowIndex).ValorMoeda = ObjTitulo.AdiantamentosAbertos.ValorTotalDisponivelParaBaixa
                Else
                    ObjTitulo.Valores(row.RowIndex).ValorMoeda = CDec(txtValorMoeda.Text)
                End If
                ObjTitulo.AdiantamentosAbertos.DistribuirValorParaBaixarAdiantamentos(IIf(ObjTitulo.Moeda.Classificacao = eTiposMoeda.Oficial, ObjTitulo.Valores.EncargoValorLiquido.ValorOficial, ObjTitulo.Valores.EncargoValorLiquido.ValorMoeda), ObjTitulo.Moeda.Classificacao)
                gridAdiantamentosDisponiveis.DataSource = ObjTitulo.AdiantamentosAbertos.ToArray
                gridAdiantamentosDisponiveis.DataBind()
            Else
                ObjTitulo.Valores(row.RowIndex).ValorMoeda = CDec(txtValorMoeda.Text)
            End If

            SessaoSalvaTitulo()

            AtualizaValoresNoForm(ObjTitulo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtVencimentoAdiantamento_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If CDate(txtVencimentoAdiantamento.Text) < Date.Now Then
                MsgBox(Me.Page, "A Data de Vencimento nao pode ser menor que a data atual")
                txtVencimentoAdiantamento.Text = Date.Now.ToString("dd/MM/yyyy")
                SessaoRecuperaTitulo()
                ObjTitulo.Adiantamento.Vencimento = Date.Now
                SessaoSalvaTitulo()
                Exit Sub
            End If
            SessaoRecuperaTitulo()
            ObjTitulo.Adiantamento.Vencimento = CDate(txtVencimentoAdiantamento.Text)
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Exit Sub
        End Try
    End Sub

    Protected Sub txtTaxaJuroAdiantamento_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxaJuroAdiantamento.TextChanged
        Try
            If Not IsNumeric(txtTaxaJuroAdiantamento.Text) Then
                MsgBox(Me.Page, "Informe uma taxa de juro mensal para o adiantamento.")
                Exit Sub
            End If
            SessaoRecuperaTitulo()
            ObjTitulo.Adiantamento.Taxa = txtTaxaJuroAdiantamento.Text
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Protected Sub txtVlrBaixa_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
    '    Try
    '        Dim txt As TextBox = CType(sender, TextBox)
    '        Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)

    '        SessaoRecuperaTitulo()
    '        If Not IsNumeric(txt.Text) Then txt.Text = 0
    '        ObjTitulo.AdiantamentosAbertos(row.RowIndex).VlrBaixa = Math.Abs(CDec(txt.Text))
    '        ObjTitulo.AdiantamentosAbertos.Titulo = ObjTitulo

    '        SessaoSalvaTitulo()
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    Protected Sub txtContaContabilCliFor_TextChanged(sender As Object, e As EventArgs) Handles txtContaContabilCliFor.TextChanged
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtQuantidade_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaTitulo()
            ObjTitulo.Quantidade = txtQuantidade.Text
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtJurosBordero_TextChanged(sender As Object, e As EventArgs) Handles txtJurosBordero.TextChanged
        Try
            SessaoRecuperaTitulo()
            CalcularDesagioBordero(ObjTitulo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "ImageButton"

    Protected Sub imgBloqueio_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("ContasAReceber", "LIBERAR") Then
                DdlProvisoes.Enabled = True
                lnkNovo.Parent.Visible = True
                BtnAdicionarConta.Visible = True
                SessaoRecuperaTitulo()
                If ObjTitulo.ContaContabilCliFor.Adiantamento Then
                    If String.IsNullOrWhiteSpace(txtPedido.Text) Then
                        cmdPedido.Enabled = True
                    End If
                    lnkExcluir.Parent.Visible = True
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExtratoPedido_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            SessaoRecuperaTitulo()

            If ObjTitulo.Codigo = 0 Then
                MsgBox(Me.Page, "Consulte o registro para visualização do extrato.")
            ElseIf ObjTitulo.CodigoEmpresa.Length = 0 Then
                MsgBox(Me.Page, "Empresa do registro não encontrada.")
            ElseIf ObjTitulo.CodigoCliFor.Length = 0 Then
                MsgBox(Me.Page, "Cliente do registro não encontrado.")
            ElseIf ObjTitulo.CodigoPedido = 0 Then
                MsgBox(Me.Page, "Registro sem pedido não pode ser visualizado.")
            Else
                Extrato.Emitir(Me.Page, FinanceiroNovo, ObjTitulo.Pedido.CodigoEmpresa, ObjTitulo.Pedido.EnderecoEmpresa, "T", ObjTitulo.Pedido.Codigo, ObjTitulo.Pedido.CodigoCliente, ObjTitulo.Pedido.EnderecoCliente)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgRecompraPorBordero_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgRecompraPorBordero.Click
        Try
            If gridRecompra.Rows.Count = 0 Then
                divRecompraPorBordero.Visible = False
                Exit Sub
            End If
            divRecompraPorBordero.Visible = Not divRecompraPorBordero.Visible
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imbLimparPedido_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imbLimparPedido.Click
        Try
            SessaoRecuperaTitulo()
            ObjTitulo.CodigoPedido = 0
            SessaoSalvaTitulo()
            txtPedido.Text = String.Empty
            divProduto.Visible = True
            divQuantidade.Visible = True
            txtDepositoCliente.Parent.Visible = True
            txtDepositoRecPag.Parent.Visible = True

            If gridAdiantamentosDisponiveis.Visible Then
                ObjTitulo.AdiantamentosAbertos.Clear()
                AtualizaValoresAdiantamentoNoForm(ObjTitulo)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgLimparPedidoRecPag_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If Funcoes.VerificaPermissao("FrmTitulo", "LEITURA") Then
            SessaoRecuperaTitulo()
            txtPedidoRecPag.Text = "0"
            btnPedidoRecPag.Enabled = True
            ObjTitulo.CodigoPedidoRecPag = ""
            SessaoSalvaTitulo()
        Else
            MsgBox(Me.Page, "Usuario sem permissao para remover o pedido.")
        End If
    End Sub

    Protected Sub imgExtratoPedidoRecPag_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgExtratoPedidoRecPag.Click
        Try
            SessaoRecuperaTitulo()

            If ObjTitulo.CodigoEmpresaRecPag.Length = 0 Then
                MsgBox(Me.Page, "Empresa do registro não encontrada.")
            ElseIf ObjTitulo.CodigoClienteRecPag.Length = 0 Then
                MsgBox(Me.Page, "Cliente do registro não encontrado.")
            ElseIf ObjTitulo.CodigoPedidoRecPag = 0 Then
                MsgBox(Me.Page, "Registro sem pedido não pode ser visualizado.")
            Else
                Extrato.Emitir(Me.Page, FinanceiroNovo, ObjTitulo.CodigoEmpresaRecPag, ObjTitulo.EndEmpresaRecPag.ToString, "T", _
                          ObjTitulo.CodigoPedidoRecPag)
                'Dim strQueryString As String
                'strQueryString = "?fim=" & DateTime.Now.ToString("dd/MM/yyyy") & vbCrLf & _
                '                 "&empresa=" & ObjTitulo.CodigoEmpresaRecPag + "-" + ObjTitulo.EndEmpresaRecPag.ToString & vbCrLf & _
                '                 "&cliente=" & ObjTitulo.CodigoClienteRecPag + "-" + ObjTitulo.EndClienteRecPag.ToString & vbCrLf & _
                '                 "&pedido=" & ObjTitulo.CodigoPedidoRecPag & vbCrLf & _
                '                 "&es=ES" & vbCrLf

                'For Each row As ClientexTipo In ObjTitulo.Empresa.Tipos
                '    If row.CodigoTipo = eTipoCliente.Revenda Then
                '        strQueryString &= "&desprd=S"
                '    End If
                'Next
                'ScriptManager.RegisterClientScriptBlock(Me, Me.GetType, "CarregarHTML", "window.open(""ExtratoPedidoHTML.aspx" & strQueryString & """, ""relatorio"", ""status=no, toolbar=yes, location=no, menubar=yes, resizable=yes, height=600, width=800, scrollbars=yes"");", True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "CheckBox"

    Protected Sub chkRecompraPorBordero_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkRecompraPorBordero.CheckedChanged
        Try
            If chkRecompraPorBordero.Checked Then
                Dim ds As DataSet
                Sql = getSqlConsulta(Nothing, True)
                ds = Banco.ConsultaDataSet(Sql, "RecompraPorBordero")
                If ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Nenhum registro encontrado.")
                    chkRecompraPorBordero.Checked = False
                    Exit Sub
                End If
                gridRecompra.DataSource = ds
                gridRecompra.DataBind()
                divRecompraPorBordero.Visible = True
            Else
                ObjTitulosAcao = Nothing
                SessaoSalvaTitulosAcao()
                gridRecompra.DataSource = Nothing
                gridRecompra.DataBind()
                divRecompraPorBordero.Visible = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ChkGridRecompra_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim ChkGridRecompra As CheckBox = CType(sender, CheckBox)
            Dim row As GridViewRow = CType(ChkGridRecompra.NamingContainer, GridViewRow)

            Dim tit As Novo.TituloNovo

            If ChkGridRecompra.Checked Then
                tit = New Novo.TituloNovo(gridRecompra.Rows(row.RowIndex).Cells(2).Text)
                SessaoRecuperaTitulo()
                ObjTitulo.Bordero.TitulosRecomprados.Add(tit.UltimoBordero)
                SessaoSalvaTitulo()

                If ObjTitulo.Bordero.TitulosRecomprados.Count = 1 Then
                    Sql = getSqlConsulta(Nothing, True)
                    gridRecompra.DataSource = Banco.ConsultaDataSet(Sql, "RecompraPorBordero")
                    gridRecompra.DataBind()
                    Dim chk As CheckBox
                    chk = gridRecompra.Rows(0).FindControl("ChkGridRecompra")
                    chk.Checked = True
                End If
            Else
                SessaoRecuperaTitulo()
                tit = ObjTitulosAcao.Find(Function(s) s.Codigo = gridRecompra.Rows(row.RowIndex).Cells(2).Text)
                ObjTitulo.Bordero.TitulosRecomprados.Remove(tit.UltimoBordero)
                SessaoSalvaTitulo()
            End If

            gridTitulosAcao.DataSource = ObjTitulo.Bordero.TitulosRecomprados.ToList
            gridTitulosAcao.DataBind()
            ''updFrmTitulo.Update()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkTodosRecompra_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim chkTodosRecompra As CheckBox = CType(sender, CheckBox)
            Dim tit As Novo.TituloNovo

            If chkTodosRecompra.Checked Then
                For Each row As GridViewRow In gridRecompra.Rows
                    tit = New Novo.TituloNovo(row.Cells(2).Text)
                    SessaoRecuperaTitulo()
                    ObjTitulo.Bordero.TitulosRecomprados.Add(tit.UltimoBordero)
                    SessaoSalvaTitulo()

                    If ObjTitulo.Bordero.TitulosRecomprados.Count = 1 Then
                        Dim Banco As New AcessaBanco
                        Sql = getSqlConsulta(Nothing, True)
                        gridRecompra.DataSource = Banco.ConsultaDataSet(Sql, "RecompraPorBordero")
                        gridRecompra.DataBind()
                    End If
                    Dim chk As CheckBox
                    chk = row.FindControl("ChkGridRecompra")
                    chk.Checked = False
                    'updFrmTitulo.Update()
                Next
            Else
                For Each row As GridViewRow In gridRecompra.Rows
                    SessaoRecuperaTitulo()
                    Dim Titulo As String = row.Cells(2).Text
                    tit = ObjTitulosAcao.Find(Function(s) s.Codigo = Titulo)
                    ObjTitulo.Bordero.TitulosRecomprados.Remove(tit.UltimoBordero)
                    SessaoSalvaTitulo()

                Next
            End If

            gridTitulosAcao.DataSource = ObjTitulo.Bordero.TitulosRecomprados.ToList
            gridTitulosAcao.DataBind()
            'updFrmTitulo.Update()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkJuros_CheckedChanged(sender As Object, e As EventArgs) Handles chkJuros.CheckedChanged
        Try
            txtJuros.Parent.Visible = True
            txtTaxa.Parent.Visible = False
            chkTaxa.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkTaxa_CheckedChanged(sender As Object, e As EventArgs) Handles chkTaxa.CheckedChanged
        Try
            txtJuros.Parent.Visible = False
            txtTaxa.Parent.Visible = True
            chkJuros.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ChkGridTitulos_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim ChkGridTitulos As CheckBox = CType(sender, CheckBox)
            Dim row As GridViewRow = CType(ChkGridTitulos.NamingContainer, GridViewRow)
            Dim tit As Novo.TituloNovo
            Dim valor As Decimal = 0

            If ChkGridTitulos.Checked Then
                tit = New Novo.TituloNovo(GridConsultaTitulos.Rows(row.RowIndex).Cells(0).Text)

                If rdAgrupar.Checked Then
                    If tit.RegistroMestre > 0 Then
                        ChkGridTitulos.Checked = False
                        Exit Sub
                    End If
                End If

                If rdAdTituloAgrupamento.Checked Then
                    If tit.Codigo = tit.RegistroMestre Then
                        ChkGridTitulos.Checked = False
                        Exit Sub
                    End If
                End If

                If rdAdTituloAgrupamento.Checked Then
                    SessaoRecuperaTituloAgrupado()
                    Dim TitExiste As Novo.TituloNovo = ObjTituloAgrupado.TitulosAgrupados.Find(Function(s) s.Codigo = row.Cells(0).Text)
                    If Not TitExiste Is Nothing Then
                        TitExiste.IUD = "U"
                    Else
                        tit.IUD = "I"
                        ObjTituloAgrupado.TitulosAgrupados.Add(tit)
                    End If
                    SessaoSalvaTituloAgrupado()
                ElseIf rdDuplicatasDescontadas.Checked Then
                    SessaoRecuperaTitulo()
                    Dim BxT As New Novo.BorderoXTitulo
                    BxT.CodigoTitulo = tit.Codigo
                    BxT.Titulo = tit
                    ObjTitulo.Bordero.TitulosDoBordero.Add(BxT)
                    SessaoSalvaTitulo()
                Else
                    SessaoRecuperaTitulosAcao()
                    ObjTitulosAcao.Add(tit)
                    If (ObjTitulosAcao.Count > 0) Then
                        lblTotalRegistroAgrupado.Parent.Visible = True
                        lblTotalRegistroAgrupado.Text = ObjTitulosAcao.Count & " Título(s) selecionado(s) no valor total de: " & ObjTitulosAcao.Sum(Function(T) T.Valores.EncargoValorLiquido.ValorOficial).ToString("N2")
                    End If
                    SessaoSalvaTitulosAcao()
                    If rdAgrupar.Checked And ObjTitulosAcao.Count = 1 Then TitulosConsulta()
                End If
            Else
                If rdAdTituloAgrupamento.Checked Then
                    SessaoRecuperaTituloAgrupado()
                    tit = ObjTituloAgrupado.TitulosAgrupados.Find(Function(s) s.Codigo = GridConsultaTitulos.Rows(row.RowIndex).Cells(0).Text)

                    If tit Is Nothing Then
                        ChkGridTitulos.Checked = True
                        Exit Sub
                    End If

                    If tit.IUD = "I" Then
                        ObjTituloAgrupado.TitulosAgrupados.Remove(tit)
                    Else
                        tit.IUD = "D"
                    End If
                    SessaoSalvaTituloAgrupado()
                ElseIf rdDuplicatasDescontadas.Checked Then
                    SessaoRecuperaTitulo()
                    Dim BxT As New Novo.BorderoXTitulo
                    BxT = ObjTitulo.Bordero.TitulosDoBordero.Find(Function(s) s.CodigoTitulo = GridConsultaTitulos.Rows(row.RowIndex).Cells(0).Text)
                    ObjTitulo.Bordero.TitulosDoBordero.Remove(BxT)
                    SessaoSalvaTitulo()
                Else
                    SessaoRecuperaTitulosAcao()
                    tit = ObjTitulosAcao.Find(Function(s) s.Codigo = GridConsultaTitulos.Rows(row.RowIndex).Cells(0).Text)
                    ObjTitulosAcao.Remove(tit)
                    If (ObjTitulosAcao.Count > 0) Then
                        lblTotalRegistroAgrupado.Parent.Visible = True
                        lblTotalRegistroAgrupado.Text = ObjTitulosAcao.Count & " Título(s) selecionado(s)  no valor total de: " & ObjTitulosAcao.Sum(Function(T) T.Valores.EncargoValorDocumento.ValorOficial).ToString("N2")
                    End If
                    SessaoSalvaTitulosAcao()
                End If
            End If
            ''updFrmTitulo.Update()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkSelectAll_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim chkSelectAll As CheckBox = CType(sender, CheckBox)
            Dim chk As CheckBox
            SessaoRecuperaTitulosAcao()
            SessaoRecuperaTitulo()
            ObjTitulosAcao.Clear()
            For Each Row As GridViewRow In GridConsultaTitulos.Rows
                chk = CType(Row.FindControl("ChkGridTitulos"), CheckBox)
                chk.Checked = chkSelectAll.Checked
                If (chkSelectAll.Checked) Then
                    'Adiciona os títulos na lista
                    Dim Titulo As Novo.TituloNovo = New Novo.TituloNovo(Row.Cells(0).Text)
                    'Selecionar Todos p/ Borderô
                    If (rdDuplicatasDescontadas.Checked) Then
                        Dim BxT As New Novo.BorderoXTitulo
                        BxT.CodigoTitulo = Titulo.Codigo
                        BxT.Titulo = Titulo
                        ObjTitulo.Bordero.TitulosDoBordero.Add(BxT)
                    Else
                        ObjTitulosAcao.Add(Titulo)
                    End If
                End If
            Next
            'Quando o "chkSelectAll" estiver desmarcado = limpa a lista de titulos
            If Not chkSelectAll.Checked Then
                ObjTitulosAcao.Clear()
                'Limpa a lista de título do Borderô
                If rdDuplicatasDescontadas.Checked Then ObjTitulo.Bordero.TitulosDoBordero.Clear()
            End If

            If (ObjTitulosAcao.Count > 0) Then
                lblTotalRegistroAgrupado.Parent.Visible = True
                lblTotalRegistroAgrupado.Text = ObjTitulosAcao.Count & " Título(s) selecionado(s) no valor total de: " & ObjTitulosAcao.Sum(Function(T) T.Valores.EncargoValorLiquido.ValorOficial).ToString("N2")
            End If

            SessaoSalvaTitulosAcao()
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "GridView"

    Protected Sub GridParcelas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridParcelas.SelectedIndexChanged
        Try
            CarregarContabilizacao(GridParcelas.SelectedRow.Cells(1).Text)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridValores_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                SessaoRecuperaTitulo()
                If ObjTitulo.ReceberPagar = "C" Then
                    'Popula Centro de Custo se a conta tiver
                    Dim ddlgridCC As DropDownList = e.Row.FindControl("ddlCentroDeCusto")
                    If ObjTitulo.Valores(e.Row.RowIndex).ContaEncargo.TemCentroDeCusto Then
                        ddl.Carregar(ddlgridCC, CarregarDDL.Tabela.CentroDeCusto, "")
                        ddlgridCC.SelectedValue = ObjTitulo.Valores(e.Row.RowIndex).CodigoCentroDeCusto
                    Else
                        ddlgridCC.Visible = False
                    End If

                    'Habilita o Debito e credito para as conta menos para principal "Valor do produto" e o liquido
                    Dim ddlgridDC As DropDownList = e.Row.FindControl("ddlDebitoCredito")
                    If ObjTitulo.Valores(e.Row.RowIndex).CodigoContaEncargo <> ObjTitulo.CodigoContaContabilCliFor And ObjTitulo.Valores(e.Row.RowIndex).CodigoContaEncargo <> ObjTitulo.CodigoContaContabilRecPag Then
                        ddlgridDC.Enabled = True
                    End If
                End If
                'Variacao Passiva/Ativa
                Dim txtValorOficial As TextBox = e.Row.FindControl("txtValorOficial")
                Dim txtValorMoeda As TextBox = e.Row.FindControl("txtValorMoeda")
                If Not ObjTitulo.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    txtValorOficial.Enabled = False
                    If e.Row.Cells(0).Text = ObjTitulo.Empresa.Empresa.CodigoContaVariacaoMonetariaPassiva OrElse _
                       e.Row.Cells(0).Text = ObjTitulo.Empresa.Empresa.CodigoContaVariacaoMonetariaAtiva Then
                        txtValorMoeda.Enabled = False
                    End If
                End If
                'Para agrupamento de títulos não permite alterar o valor 
                If e.Row.Cells(0).Text = ObjTitulo.ContaContabilCliFor.Conta AndAlso _
                    ObjTitulo.TitulosAgrupados.Count > 0 Then
                    txtValorOficial.Enabled = False
                End If

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridBordero_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles gridBordero.SelectedIndexChanged
        Try
            Limpar()
            ObjTitulo = New Novo.TituloNovo(gridBordero.SelectedRow.Cells(1).Text())
            ObjTitulo.IUD = "U"
            CarregarFormularioComAClasse(ObjTitulo)
            SessaoSalvaTitulo()
            tbcFrmTitulo.ActiveTabIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridConsultaTitulos_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridConsultaTitulos.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                If rdAdTituloAgrupamento.Checked And Not String.IsNullOrWhiteSpace(Server.HtmlDecode(e.Row.Cells(7).Text)) Then
                    Dim ChkGridTitulos As CheckBox = e.Row.FindControl("ChkGridTitulos")
                    Dim row As GridViewRow = CType(ChkGridTitulos.NamingContainer, GridViewRow)
                    ChkGridTitulos.Checked = True
                End If

                Dim btn As Button
                Dim lbl As Label = e.Row.FindControl("lblEmiteDuplicatasGrid")
                If lbl.Text = "False" Then
                    btn = e.Row.FindControl("btnDuplicataGrid")
                    btn.Visible = False
                End If
                lbl.Visible = False


                If e.Row.Cells(17).Text = "0" Then e.Row.Cells(17).Text = ""
                If e.Row.Cells(4).Text = "0" Then e.Row.Cells(4).Text = ""
                e.Row.Cells(5).Text = ""

                If e.Row.RowType = DataControlRowType.DataRow Then
                    If (e.Row.Cells(7).Text.Equals("F")) Then
                        Dim lnk As LinkButton = CType(e.Row.FindControl("lnkSelecionarAgrupado"), LinkButton)
                        lnk.Visible = True
                    Else
                        Dim lnk As LinkButton = CType(e.Row.FindControl("lnkSelecionar"), LinkButton)
                        lnk.Visible = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridConsultaTitulos.SelectedIndexChanged
        Try
            Limpar()
            ObjTitulo = New Novo.TituloNovo(GridConsultaTitulos.SelectedRow.Cells(0).Text())
            ObjTitulo.IUD = "U"
            CarregarFormularioComAClasse(ObjTitulo)
            SessaoSalvaTitulo()
            tbcFrmTitulo.ActiveTabIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "RadioButton"

    Protected Sub RbGeral_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            chkPrevisao.Checked = False
            chkProvisao.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RbAtivo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If chkPrevisao.Visible = True Then
                chkPrevisao.Checked = True
                chkProvisao.Checked = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RbBaixado_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            chkPrevisao.Checked = False
            chkProvisao.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdConsultar_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdConsultar.CheckedChanged
        Try
            PaineisInvisiveis()
            lblAcao.Text = "Consultar"
            chkBaixa.Enabled = True
            chkPrevisao.Enabled = True
            chkProvisao.Enabled = True
            chkContabil.Enabled = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdConsultarBordero_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdConsultarBordero.CheckedChanged
        Try
            PaineisInvisiveis()
            lblAcao.Text = "Consultar Borderôs"
            divAcaoConsultaBordero.Visible = True
            chkBaixa.Enabled = True
            chkPrevisao.Enabled = True
            chkProvisao.Enabled = True
            chkContabil.Enabled = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdAgrupar_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdAgrupar.CheckedChanged
        Try
            PaineisInvisiveis()
            divAcaoAgrupar.Visible = True
            lnkAcao.Parent.Visible = True
            lnkAcao.Text = "Agrupar"
            lblAcao.Text = "Agrupar Contas a " & IIf(chkReceber.Checked, "Receber", "Pagar")

            chkBaixa.Enabled = False
            chkBaixa.Checked = False
            chkPrevisao.Enabled = False
            chkPrevisao.Checked = False
            chkProvisao.Enabled = False
            chkProvisao.Checked = True
            chkContabil.Enabled = False
            chkContabil.Checked = False

            If (rdAgrupar.Checked AndAlso GridConsultaTitulos.Rows.Count > 0) Then
                Dim chkSelectAll As CheckBox = CType(GridConsultaTitulos.HeaderRow.FindControl("chkSelectAll"), CheckBox)
                chkSelectAll.Visible = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdBaixar_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdBaixar.CheckedChanged
        Try
            PaineisInvisiveis()
            divAcaoBaixar.Visible = True
            lnkAcao.Parent.Visible = True
            lnkAcao.Text = "Baixar"
            lblAcao.Text = "Baixar Contas a " & IIf(chkReceber.Checked, "Receber", "Pagar")

            chkBaixa.Enabled = False
            chkBaixa.Checked = False
            chkPrevisao.Enabled = False
            chkPrevisao.Checked = False
            chkProvisao.Enabled = False
            chkProvisao.Checked = True
            chkContabil.Enabled = False
            chkContabil.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdReprogramar_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdReprogramar.CheckedChanged
        Try
            PaineisInvisiveis()
            divAcaoReprogramacao.Visible = True
            lnkAcao.Parent.Visible = True
            lnkAcao.Text = "Reprogramar"
            lblAcao.Text = "Reprogramar"

            chkBaixa.Enabled = False
            chkBaixa.Checked = False
            chkPrevisao.Enabled = True
            chkPrevisao.Checked = True
            chkProvisao.Enabled = True
            chkProvisao.Checked = True
            chkContabil.Enabled = False
            chkContabil.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdAdTituloAgrupamento_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdAdTituloAgrupamento.CheckedChanged
        Try
            PaineisInvisiveis()
            divParametrosAdicionarTituloAgrupado.Visible = True
            divAcaoAdicionar.Visible = True
            lnkAcao.Parent.Visible = True
            lnkAcao.Text = "Agrupar"
            lblAcao.Text = "Adicionar/Remover Título do Agrupamento"

            chkBaixa.Enabled = False
            chkBaixa.Checked = False
            chkPrevisao.Enabled = False
            chkPrevisao.Checked = False
            chkProvisao.Enabled = False
            chkProvisao.Checked = True
            chkContabil.Enabled = False
            chkContabil.Checked = False
            chkPagar.Checked = True
            chkReceber.Checked = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdDuplicatasDescontadas_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdDuplicatasDescontadas.CheckedChanged
        Try
            PaineisInvisiveis()
            lblAcao.Text = "Duplicatas Descontadas - Gera Título para o Borderô"
            divAcaoDuplicatas.Visible = True
            lnkAcao.Parent.Visible = True
            lnkAcao.Text = "Borderô"

            chkBaixa.Enabled = False
            chkBaixa.Checked = False
            chkPrevisao.Enabled = False
            chkPrevisao.Checked = False
            chkProvisao.Enabled = False
            chkProvisao.Checked = True
            chkContabil.Enabled = False
            chkContabil.Checked = False
            chkPagar.Enabled = False
            chkPagar.Checked = False
            chkReceber.Enabled = False
            chkReceber.Checked = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdGeraDuplicatasCobrancaSimples_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdGeraDuplicatasCobrancaSimples.CheckedChanged
        Try
            PaineisInvisiveis()
            lblAcao.Text = "Duplicatas Cobrança Simples Banco - (Não gera título para o Borderô)"
            divAcaoDuplicatasCobrancaSimples.Visible = True
            lnkAcao.Parent.Visible = True
            lnkAcao.Text = "Borderô"

            chkBaixa.Enabled = False
            chkBaixa.Checked = False
            chkPrevisao.Enabled = False
            chkPrevisao.Checked = False
            chkProvisao.Enabled = False
            chkProvisao.Checked = True
            chkContabil.Enabled = False
            chkContabil.Checked = False
            chkPagar.Enabled = False
            chkPagar.Checked = False
            chkReceber.Enabled = False
            chkReceber.Checked = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rbCheque_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rbCheque.CheckedChanged
        Try
            PaineisInvisiveis()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RdReceber_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RdReceber.CheckedChanged
        Try
            SessaoRecuperaTitulo()
            MudouRP("R")
            SessaoSalvaTitulo()
            pnlBordero.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RdPagar_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaTitulo()
            MudouRP("P")
            SessaoSalvaTitulo()
            pnlBordero.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RdContabil_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaTitulo()
            MudouRP("C")
            SessaoSalvaTitulo()
            pnlBordero.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#End Region

#Region "Métodos"

    Private Sub Limpar(Optional ByVal ParametroLimpeza As Integer = 0)
        Try
            If Not chkManterLancamento.Checked Then
                Session.Remove("ObjTitulo" & HID.Value)
                Session.Remove("CONTARECPAG" & HID.Value)
                Session.Remove("objClienteCR" & HID.Value)
                Session.Remove("objBancoFRMTIT" & HID.Value)
                Session.Remove("objTituloPDC" & HID.Value)
                Session.Remove("objPedidoTitulo" & HID.Value.ToString)

                If ParametroLimpeza <> 1 Then
                    ObjTitulosAcao = Nothing
                    SessaoSalvaTitulosAcao()
                End If

                HID.Value = Guid.NewGuid().ToString
                ucConsultaClientes.SetarHID(HID.Value)
                ucConsultaEncargosPlanoDeContas.SetarHID(HID.Value)
                ucConsultaPedidos.SetarHID(HID.Value)
                ucConsultaProduto.SetarHID(HID.Value)

                ObjTitulo = New Novo.TituloNovo
                ObjTitulo.DesligarControles()
                ObjTitulo.IUD = "I"
                ObjTitulo.CodigoSituacao = eSituacao.Normal

                ObjTitulo.Movimento = Now
                ObjTitulo.Vencimento = Now
                ObjTitulo.Reprogramacao = Now
                ObjTitulo.DataMoeda = Now
                ObjTitulo.DataBaixa = Now

                txtMovimento.Text = Format(Today, "dd/MM/yyyy")
                txtProrrogacao.Text = Format(Today, "dd/MM/yyyy")
                txtPeriodoInicialConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")
                txtPeriodoFinalConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")

                imgBloqueio.Visible = False
                imgExtratoPedido.Visible = False
                txtLiberarPedido.Value = "N"
                txtUsuarioLiberarPedido.Value = ""
                txtUsuarioLiberarPedidoData.Value = ""
                lblContabilizacao.Text = "Contabilização"
                txtContaContabilRecPag.Text = String.Empty

                DdlUnidadeDeNegocioEmpresaCliente.ClearSelection()
                ddl.Carregar(DdlEmpresaCliente, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocioEmpresaCliente.SelectedValue, False)
                txtCliente.Text = ""

                txtContaContabilCliFor.Text = String.Empty
                hdnCodigoContaContabilCliFor.Value = String.Empty
                divAdiantamento.Visible = False

                lblBanco.Text = "Banco"
                lblAgencia.Text = "Agência"
                lblContaCorrente.Text = "Conta"

                txtVencimentoAdiantamento.Text = ""
                txtTaxaJuroAdiantamento.Text = ""

                ddlMoeda.SelectedIndex = 1
                ObjTitulo.CodigoMoeda = ddlMoeda.SelectedValue

                DdlProvisoes.SelectedIndex = 0
                ObjTitulo.CodigoProvisao = eProvisao.Provisao

                ddlIndexador.SelectedIndex = 3
                ObjTitulo.CodigoIndexador = ddlIndexador.SelectedValue

                DdlTiposDeRecebimentos.ClearSelection()
                ddlUnidadeDeNegocioRecPag.ClearSelection()
                ddlContaContabilEmpresaRecPag.ClearSelection()

                ddlCarteiraDoTitulo.SelectedIndex = 0
                lnkDuplicata.Parent.Visible = False

                txtRegistro.Enabled = True

                txtPedido.Text = String.Empty
                txtRegistro.Text = ""

                txtProduto.Text = ""
                txtQuantidade.Text = ""
                txtQuantidade.Enabled = False

                txtHistorico.Text = ""
                txtProrrogacao.Text = Now.ToString("dd/MM/yyyy")
                LblVencOriginal.Text = ""

                txtContaContabilRecPag.Visible = False
                btnPlanoDeContaRecPag.Visible = False
                ddlContaContabilEmpresaRecPag.Visible = True

                cmdPedido.Enabled = True

                lblCotacao.Text = "Cotação"
                txtCotacao.Enabled = True

                txtMestre.Text = ""
                txtObservacoes.Text = ""

                divBaixaAdiantamento.Visible = False
                gridAdiantamentosDisponiveis.DataSource = Nothing
                gridAdiantamentosDisponiveis.DataBind()

                gridBaixasAdiantamentos.DataSource = Nothing
                gridBaixasAdiantamentos.DataBind()

                gridValores.DataSource = Nothing
                gridValores.DataBind()

                GridParcelas.DataSource = Nothing
                GridParcelas.DataBind()

                gridRazao.DataSource = Nothing
                gridRazao.DataBind()

                gdvDesmembramentoValores.DataSource = Nothing
                gdvDesmembramentoValores.DataBind()

                GridConsultaTitulos.Visible = True

                DdlUnidadeDeNegocioEmpresaCliente.Enabled = True
                DdlEmpresaCliente.Enabled = True
                btnClientesTitulo.Enabled = True
                DdlProvisoes.Enabled = True
                txtHistorico.Enabled = True

                ddlMoeda.Enabled = True
                ddlIndexador.Enabled = True

                lnkNovo.Parent.Visible = True
                lnkExcluir.Parent.Visible = False
                lnkRecibo.Parent.Visible = False
                lblAcaoTitulo.Visible = False
                BtnAdicionarConta.Visible = True

                VerificaUnidade()
                MudouRP("R")

                pnlBordero.Visible = True
                txtBordero.Text = ""
                lblBorderoMestre.Text = ""
                pnlBorderoDuplicatasImpressao.Visible = False

                chkBaixa.Enabled = True
                chkPrevisao.Enabled = True
                chkProvisao.Enabled = True
                chkPagar.Enabled = True
                chkReceber.Enabled = True
                chkContabil.Enabled = True

                ObjTitulo.VerificarIndice()
                AtualizaValoresNoForm(ObjTitulo)
                ObjTitulo.LigarControles()
                SessaoSalvaTitulo()
                DatadaBaixa()
                RdPagar.Enabled = True
                RdReceber.Enabled = True
                RdContabil.Enabled = True
                lblTotalDeCreditos.Visible = False
                lblTotalDeDebitos.Visible = False
                txtContaContabilCliFor.Enabled = True
                btnPlanoDeContas.Visible = True
                txtCliente.Enabled = True
                btnClientesTitulo.Visible = True
                ddlMoeda.Enabled = True
                txtCodigoDeBarras.Text = String.Empty
            Else
                SessaoRecuperaTitulo()
                ObjTitulo.Valores.EncargoValorDocumento.Valor = 0
                ObjTitulo.Valores.EncargoValorLiquido.Valor = 0
                gridValores.DataSource = ObjTitulo.Valores
                gridValores.DataBind()
            End If

        Catch ex As Exception
            HttpContext.Current.Session("ssMessage") = ex.Message
            MsgBox(Me.Page, ex.Message.ToString)
        End Try

    End Sub

    Private Sub LimparConsulta()
        DdlUnidadeConsultaTitulos.ClearSelection()
        DdlEmpresaConsultaTitulos.Items.Clear()
        txtClienteConsulta.Text = String.Empty
        txtPeriodoInicialConsultaTitulos.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtPeriodoFinalConsultaTitulos.Text = DateTime.Now.ToString("dd/MM/yyyy")
        GridConsultaTitulos.DataSource = Nothing
        GridConsultaTitulos.DataBind()
        lnkAcao.Parent.Visible = False
        divAcoes.Visible = False
        SessaoRecuperaTitulosAcao()
        ObjTitulosAcao.Clear()
        SessaoSalvaTitulo()
        lblTotalRegistroAgrupado.Text = String.Empty
        rdConsultar.Checked = True
        rdConsultarBordero.Checked = False
        rdAgrupar.Checked = False
        rdBaixar.Checked = False
        rdReprogramar.Checked = False
        rdAdTituloAgrupamento.Checked = False
        rdDuplicatasDescontadas.Checked = False
        rbCheque.Checked = False
        divParametrosAdicionarTituloAgrupado.Visible = False
        txtMestre.Text = String.Empty
    End Sub

#Region "Sessão"

    Private Sub SessaoSalvaTitulo()
        Session("ObjTitulo" + HID.Value) = ObjTitulo
    End Sub

    Private Sub SessaoRecuperaTitulo()
        If Not ObjTitulo Is Nothing Then Exit Sub
        If Session("ObjTitulo" + HID.Value) Is Nothing Then
            ObjTitulo = New Novo.TituloNovo
        Else
            ObjTitulo = CType(Session("ObjTitulo" + HID.Value), Novo.TituloNovo)
        End If
    End Sub

    'Usado no Procedimento de Adicionar um titulo ao agrupamento
    Private Sub SessaoSalvaTituloAgrupado()
        Session("ObjTituloAgrupado" + HID.Value) = ObjTituloAgrupado
    End Sub

    Private Sub SessaoRecuperaTituloAgrupado()
        If Not ObjTituloAgrupado Is Nothing Then Exit Sub
        If Session("ObjTituloAgrupado" + HID.Value) Is Nothing Then
            ObjTituloAgrupado = New Novo.TituloNovo
        Else
            ObjTituloAgrupado = CType(Session("ObjTituloAgrupado" + HID.Value), Novo.TituloNovo)
        End If
    End Sub

    'Usado no Procedimento de Geracao de Duplicatas "BORDERO"
    Private Sub SessaoSalvaTituloBordero()
        Session("ObjTituloBordero" + HID.Value) = ObjTituloBordero
    End Sub

    Private Sub SessaoRecuperaTituloBordero()
        If Not ObjTituloBordero Is Nothing Then Exit Sub
        If Session("ObjTituloBordero" + HID.Value) Is Nothing Then
            ObjTituloBordero = New Novo.TituloNovo
        Else
            ObjTituloBordero = CType(Session("ObjTituloBordero" + HID.Value), Novo.TituloNovo)
        End If
    End Sub

    'Usado no Procedimento de Reprogramacao e Baixa em Lote
    Private Sub SessaoSalvaTitulosAcao()
        Session("ObjTitulosAcao" + HID.Value) = ObjTitulosAcao
    End Sub

    Private Sub SessaoRecuperaTitulosAcao(Optional ByVal pHID As HiddenField = Nothing)
        Dim Ponteiro As String
        If Not pHID Is Nothing Then
            Ponteiro = pHID.Value
        Else
            Ponteiro = HID.Value
        End If

        If Not ObjTitulosAcao Is Nothing Then Exit Sub
        If Session("ObjTitulosAcao" + Ponteiro) Is Nothing Then
            ObjTitulosAcao = New Novo.ListTituloNovo
        Else
            ObjTitulosAcao = CType(Session("ObjTitulosAcao" + Ponteiro), Novo.ListTituloNovo)
        End If
    End Sub

#End Region

    Private Sub AtribuirValoresCampos()
        If Not Session("ClienteLC" + HID.Value) Is Nothing Then
            SessaoRecuperaTitulo()
            Dim cli As Cliente = Session("ClienteLC" + HID.Value)
            ObjTitulo.ClienteRecPag = cli
            SessaoSalvaTitulo()
            Session.Remove("ClienteLC" + HID.Value)
            Exit Sub
        End If

        If Not Session("objContaBancaria" + HID.Value) Is Nothing Then
            SessaoRecuperaTitulo()
            Dim cXb As ClienteXContaBancaria = Session("objContaBancaria" + HID.Value)
            ObjTitulo.CodigoBancoCliFor = cXb.CodigoBanco
            ObjTitulo.CodigoAgenciaCliFor = cXb.CodigoAgencia
            ObjTitulo.DigitoAgenciaCliFor = cXb.DigitoAgencia
            ObjTitulo.ContaCliFor = cXb.ContaCorrente
            ObjTitulo.DigitoContaCliFor = cXb.DigitoConta
            SessaoSalvaTitulo()
            AtualizaDadosBancariosNoForm(ObjTitulo)
            Session.Remove("objContaBancaria" + HID.Value)
            Exit Sub
        End If
    End Sub

    Private Sub VerificaUnidade()
        Dim sql As String = ""
        sql = " SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              "   from Usuarios" & vbCrLf & _
              "  where Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"

        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Consulta").Tables(0).Rows
            DdlUnidadeDeNegocioEmpresaCliente.SelectedValue = Dr("AcessoUnidade")
            ObjTitulo.CodigoUnidadeDeNegocio = Dr("AcessoUnidade")
            ddl.Carregar(DdlEmpresaCliente, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocioEmpresaCliente.SelectedValue, False)
            DdlEmpresaCliente.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
            ObjTitulo.CodigoEmpresa = Dr("AcessoEmpresa")
            ObjTitulo.EnderecoEmpresa = Dr("AcessoEndEmpresa")
            'Pagadora/Recebora
            ObjTitulo.CodigoUnidadeDeNegocioRecPag = Dr("AcessoUnidade")
            ddlUnidadeDeNegocioRecPag.SelectedValue = ObjTitulo.CodigoUnidadeDeNegocioRecPag
            ddl.Carregar(DdlEmpresaRecPag, CarregarDDL.Tabela.Empresas, ObjTitulo.CodigoUnidadeDeNegocioRecPag, True)
            ObjTitulo.CodigoEmpresaRecPag = Dr("AcessoEmpresa")
            ObjTitulo.EndEmpresaRecPag = Dr("AcessoEndEmpresa")
            DdlEmpresaRecPag.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Function ValidaCampos() As Boolean
        If ObjTitulo.Valores.TotalDebitos <> ObjTitulo.Valores.TotalCreditos Then
            MsgBox(Me.Page, "A soma dos debitos e a soma dos creditos devem ser iguais.")
            Return False
        End If
        If (ObjTitulo.Valores.TotalDebitos + ObjTitulo.Valores.TotalCreditos) = 0 Then
            MsgBox(Me.Page, "Não é possível gravar título com valor 0,00.")
            Return False
        End If
        If ObjTitulo.isBaixaAdiantamento Then
            ObjTitulo.AdiantamentosAbertos.Titulo = ObjTitulo
            If ObjTitulo.AdiantamentosAbertos.ValorTotalInformadoParaBaixa > ObjTitulo.Valores.EncargoValorDocumento.Valor Then
                MsgBox(Me.Page, "A soma da baixas é maior que o saldo disponivel para baixa.")
                Return False
            End If
        End If

        'PARA PRODUÇÃO COLOCAR REGRA ABAIXO.
        'If Not ObjTitulo.Bloqueio Then
        '    If Not (ObjTitulo.CodigoContaContabilRecPag = "101010101" AndAlso ObjTitulo.Valores.EncargoValorDocumento.Valor < 150) Then
        '        MsgBox(Me.Page, "Título Bloqueado!")
        '        Return False
        '    End If
        'End If

        '******   Moeda e Indexador  **************************************
        If ObjTitulo.CodigoUnidadeDeNegocio.Length = 0 Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório.")
            Return False
        ElseIf ObjTitulo.CodigoEmpresa.Length = 0 Then
            MsgBox(Me.Page, "Empresa do Cliente é obrigatório.")
            Return False
        ElseIf ObjTitulo.CodigoCliFor.Length = 0 And ObjTitulo.CodigoContaContabilCliFor.Length <> 9 Then
            MsgBox(Me.Page, "Cliente é obrigatório.")
            Return False
        ElseIf ObjTitulo.CodigoProvisao = 0 Then
            MsgBox(Me.Page, "Previsao é obrigatório.")
            Return False
        ElseIf ObjTitulo.CodigoIndexador = 0 Then
            MsgBox(Me.Page, "Indexador é obrigatório.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlTiposDeRecebimentos.SelectedValue) Then
            MsgBox(Me.Page, lblTipoRecPag.Text & " é obrigatório.")
            Return False
        End If
        If Not String.IsNullOrWhiteSpace(DdlProvisoes.SelectedValue) AndAlso (DdlTiposDeRecebimentos.SelectedValue = 4 OrElse Not String.IsNullOrWhiteSpace(txtCodigoDeBarras.Text)) Then '  Validação Boleto Bancário
            If ckPreImpresso.Checked = False Then
                If DdlProvisoes.Text <> "" And DdlTiposDeRecebimentos.Text <> "" Then

                    If Trim(txtCodigoDeBarras.Text) <> "" Then
                        If CkbCodigoDeBarras.Checked Then txtCodigoDeBarras.Text = Funcoes.FormataLinhaDigitavelOriginal(txtCodigoDeBarras.Text)
                        If Not Funcoes.ValidaCodigoBarras(txtCodigoDeBarras.Text, CkbCodigoDeBarras.Checked, txtProrrogacao.Text, ObjTitulo.Valores.EncargoValorDocumento.ValorOficial, ObjTitulo.CodigoEmpresa, ObjTitulo.EnderecoEmpresa, Banco) Then
                            MsgBox(Me.Page, "Código de barras Inválido.")
                            Return False
                        End If
                    End If
                Else
                    MsgBox(Me.Page, "Preenchimento do código de barras somente aceito para boletos bancarios.")
                    Return False
                End If
                'Else
                '    MsgBox(Me.Page, "Tipo de pagto e previsao são obrigatórios para validação do código de barras.")
                '    Return False
            End If
        End If
        'Para Boleto não pode desmembrar o titulo
        If ObjTitulo.TituloOriginal IsNot Nothing AndAlso _
           ObjTitulo.Valores.EncargoValorDocumento.Valor <> ObjTitulo.TituloOriginal.Valores.EncargoValorDocumento.Valor AndAlso _
           ObjTitulo.CodigoTipoPgto = 4 AndAlso Not String.IsNullOrWhiteSpace(txtCodigoDeBarras.Text) Then
            MsgBox(Me.Page, "Para alterar o tipo de pagamento: boleto é necessario limpar o campo código e barras.")
            Return False
        End If
        'Caso seja uma baixa valida os dados da Empresa Pagadora ou Recebedora
        If (ObjTitulo.CodigoProvisao = eProvisao.Baixa AndAlso Not RdContabil.Checked) Then
            If (String.IsNullOrWhiteSpace(ObjTitulo.CodigoUnidadeDeNegocioRecPag)) Then
                MsgBox(Me.Page, "Unidade de negócio do Cliente é obrigatória.")
                Return False
            ElseIf (String.IsNullOrWhiteSpace(ObjTitulo.CodigoEmpresaRecPag)) Then
                MsgBox(Me.Page, "Empresa Recebedora é obrigatório.")
                Return False
            ElseIf ObjTitulo.CodigoTipoPgto = 0 Then
                MsgBox(Me.Page, lblTipoRecPag.Text & " é obrigatório.")
                Return False
            ElseIf ObjTitulo.CodigoContaContabilRecPag.Length = 0 Then
                MsgBox(Me.Page, "Banco é obrigatório.")
                Return False
            End If
            'Caso a conta contábil não possua cliente e tenha somente 7 digitos
            If (Not ObjTitulo.ContaContabilRecPag.TemCliente And ObjTitulo.ContaContabilRecPag.Conta.Length.Equals(7)) Then
                MsgBox(Me.Page, "Não é possível associar título nesta Conta Banco: " & ddlContaContabilEmpresaRecPag.SelectedValue)
                Return False
            End If
            'Para a conta contábil com 7 digitos que possua cliente = 'S' o campo cliente é obrigatório
            If ObjTitulo.ContaContabilRecPag.TemCliente And ObjTitulo.ContaContabilRecPag.Conta.Length.Equals(7) And String.IsNullOrWhiteSpace(ObjTitulo.CodigoCliFor) Then
                MsgBox(Me.Page, "Para a Conta Banco: " & ddlContaContabilEmpresaRecPag.SelectedValue & " o campo " & lblCliFor.Text & " é obrigatório.")
                Return False
            End If

            'Acrescenta o nome do cliente ao Histórico do Título
            If ObjTitulo.ContaContabilCliFor.TemCliente AndAlso ObjTitulo.CodigoPedido = 0 Then
                ObjTitulo.Historico = txtHistorico.Text & IIf(ObjTitulo.CodigoProvisao = eProvisao.Baixa, " - " & ObjTitulo.CliFor.Nome, String.Empty)
            End If
            'Tiutlo com valor de Doc = 0, tem q ser baixado na mesma empresa do titulo
            If ObjTitulo.Valores.EncargoValorDocumento.Valor = 0 AndAlso Not ObjTitulo.CodigoEmpresa = ObjTitulo.CodigoEmpresaRecPag Then
                MsgBox(Me.Page, "Titulo com valor 0 não pode ser baixado em empresa diferente do titulo.")
                Return False
            End If
        End If
        If ObjTitulo.Historico.Trim.Length = 0 Then
            MsgBox(Me.Page, "Histórico é obrigatório.")
            Return False
        End If

        'If (ObjTitulo.TitulosAgrupados.Count AndAlso ObjTitulo.CodigoProvisao <> eProvisao.Baixa) Then
        '    MsgBox(Me.Page, "Para Agrupamento de Títulos é obrigatório selecionar Baixa no campo Provisão.")
        '    Return False
        'End If

        'Caso seja um adiantamento avulso e esteja sendo associado a um pedido
        If ObjTitulo.TituloOriginal IsNot Nothing AndAlso _
           ObjTitulo.TituloOriginal.CodigoPedido = 0 AndAlso _
           ObjTitulo.ContaContabilCliFor.Adiantamento AndAlso _
           ObjTitulo.Pedido IsNot Nothing AndAlso _
           ObjTitulo.Pedido.Titulos.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao).Sum(Function(t) t.Valores.EncargoValorDocumento.Valor) < ObjTitulo.Valores.EncargoValorDocumento.Valor Then
            MsgBox(Me.Page, "Não há saldo suficiente para associar um adiantamento no pedido.")
            Return False
        End If

        If ObjTitulo.NotaTitulo IsNot Nothing AndAlso _
            ObjTitulo.NotaTitulo.NotaFiscal IsNot Nothing AndAlso _
            ObjTitulo.DataBaixa < ObjTitulo.NotaTitulo.NotaFiscal.DataNota Then
            MsgBox(Me.Page, "Não é possivel baixar titulo com Data anterior a Data de Emissão da Nota Fiscal")
            Return False
        End If
        'Agrupamento de titulo
        If ObjTitulo.TitulosAgrupados.Count > 0 And ObjTitulo.CodigoProvisao = eProvisao.Previsao Then
            MsgBox(Me.Page, "Agrupamento de título não pode ser salvo em previsão")
            Return False
        End If

        Return True
    End Function

#Region "Atualização do Formulário"

    Private Sub CarregarFormularioComAClasse(ByRef pObjTitulo As Novo.TituloNovo)
        Try
            txtRegistro.Text = IIf(pObjTitulo.Codigo > 0, pObjTitulo.Codigo, String.Empty)

            'Associa o titulo como o titulo original - recompra
            If pObjTitulo.Codigo > 0 Then pObjTitulo.TituloOriginal = New Novo.TituloNovo(pObjTitulo.Codigo)

            'MudouRP(pObjTitulo.TituloOriginal.ReceberPagar)
            MudouRP(pObjTitulo.ReceberPagar)

            hdnValorTitulo.Value = pObjTitulo.Valores.EncargoValorDocumento.ValorOficial
            DdlUnidadeDeNegocioEmpresaCliente.SelectedValue = pObjTitulo.CodigoUnidadeDeNegocio
            ddl.Carregar(DdlEmpresaCliente, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocioEmpresaCliente.SelectedValue, False)
            DdlEmpresaCliente.SelectedValue = pObjTitulo.CodigoEmpresa & "-" & pObjTitulo.EnderecoEmpresa

            'Para o Borderô
            If Not pObjTitulo.Bordero Is Nothing AndAlso pObjTitulo.Bordero.TitulosDoBordero.Count > 0 Then
                divBordero.Visible = True
                divCodigoDeBarras.Visible = False
                divBanco.Visible = False
                divPedido.Visible = False
                txtDataEnvioBordero.Text = pObjTitulo.Bordero.DataEnvio.ToString("dd/MM/yyyy")
                txtHistorico.Text = pObjTitulo.Historico
                lblDiasBanco.Visible = True

                Dim BxC As New [Lib].Negocio.BancosXContas(pObjTitulo.CodigoContaContabilRecPag)
                If BxC IsNot Nothing AndAlso BxC.Banco IsNot Nothing Then
                    lblDiasBanco.Text = BxC.Banco.LiquidacaoDias.ToString & " Dia(s)"
                End If

                Try
                    txtContaContabilCliFor.Text = IIf(pObjTitulo.CodigoContaContabilCliFor.Length = 0, "", pObjTitulo.CodigoContaContabilCliFor & " - " & pObjTitulo.ContaContabilCliFor.Titulo)
                Catch ex As Exception
                    pObjTitulo.CodigoContaContabilCliFor = txtContaContabilCliFor.Text
                End Try
            Else
                txtContaContabilCliFor.Text = IIf(pObjTitulo.CodigoContaContabilCliFor.Length = 0, "", pObjTitulo.CodigoContaContabilCliFor & " - " & pObjTitulo.ContaContabilCliFor.Titulo)
                lblDiasBanco.Visible = False
            End If

            '************************************************************************************************************
            '*****************************************  Adiantamento  ***************************************************
            '************************************************************************************************************
            'Carrega a lista de adiantamentos abertos somente para adiantamentos
            If pObjTitulo.AdiantamentosAbertos IsNot Nothing AndAlso (pObjTitulo.ContaContabilCliFor.Adiantamento OrElse pObjTitulo.CodigoProvisao = eProvisao.Provisao) Then
                AtualizaValoresAdiantamentoNoForm(pObjTitulo)
            End If


            AtualizaDadosBancariosNoForm(pObjTitulo)

            '***********************************************************************************
            ' Cliente / Pagar Receber Debito
            '***********************************************************************************
            If pObjTitulo.ReceberPagar = "C" And Not pObjTitulo.ContaContabilCliFor.TemCliente Then
                txtCliente.Text = "CONTA SEM CLIENTE"
                btnClientesTitulo.Enabled = False
                txtCliente.Enabled = False
            Else
                btnClientesTitulo.Enabled = True
                txtCliente.Enabled = True
                txtCliente.Text = pObjTitulo.CliFor.Codigo & " - " & pObjTitulo.CliFor.CodigoEndereco & " ..." & pObjTitulo.CliFor.Nome & " / " & pObjTitulo.CliFor.Cidade & "-" & pObjTitulo.CliFor.CodigoEstado
            End If

            '***********************************************************************************
            ' Empresa Pagadora / Pagar Receber Credito
            '***********************************************************************************
            If pObjTitulo.CodigoEmpresaRecPag.Length > 0 Then
                ddlUnidadeDeNegocioRecPag.SelectedValue = pObjTitulo.CodigoUnidadeDeNegocioRecPag
                ddl.Carregar(DdlEmpresaRecPag, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocioRecPag.SelectedValue, False)
                DdlEmpresaRecPag.SelectedValue = pObjTitulo.CodigoEmpresaRecPag & "-" & pObjTitulo.EndEmpresaRecPag
                DdlTiposDeRecebimentos.SelectedValue = IIf(pObjTitulo.CodigoTipoPgto = 0, "", pObjTitulo.CodigoTipoPgto)
                If Not pObjTitulo.CodigoProvisao = eProvisao.Baixa Then
                    AdicionaContaTEDDOC(pObjTitulo)
                End If
                ddlContaContabilEmpresaRecPag.SelectedValue = IIf(pObjTitulo.CodigoContaContabilRecPag.Length = 0, "", pObjTitulo.CodigoContaContabilRecPag)
            End If

            If pObjTitulo.CodigoMoeda > 0 Then ddlMoeda.SelectedValue = pObjTitulo.CodigoMoeda
            If pObjTitulo.CodigoIndexador > 0 Then ddlIndexador.SelectedValue = pObjTitulo.CodigoIndexador

            ddlCarteiraDoTitulo.SelectedValue = pObjTitulo.CodigoCarteiraDoTitulo
            lnkDuplicata.Parent.Visible = pObjTitulo.CarteiraDoTitulo.EmiteDuplicata
            DdlProvisoes.SelectedValue = pObjTitulo.CodigoProvisao

            '****************************************************************************************************************************************************************
            If pObjTitulo.RegistroMestre > 0 AndAlso pObjTitulo.Codigo <> pObjTitulo.RegistroMestre Then
                lnkNovo.Enabled = False
                lnkExcluir.Enabled = False
            End If

            If pObjTitulo.RegistroMestre > 0 Then
                txtMestre.Text = "Mestre " & pObjTitulo.RegistroMestre
            Else
                txtMestre.Text = ""
            End If

            'If pObjTitulo.Codigo = pObjTitulo.RegistroMestre Or pObjTitulo.Codigo = pObjTitulo.Bordero.CodigoTituloBordero Then
            '    lnkNovo.Enabled = False
            'End If
            '****************************************************************************************************************************************************************

            If pObjTitulo.Bordero.CodigoTituloBordero > 0 AndAlso pObjTitulo.Codigo <> pObjTitulo.Bordero.CodigoTituloBordero Then
                lnkNovo.Enabled = False
                lnkExcluir.Enabled = False
            End If

            pnlBordero.Visible = False
            pnlBorderoDuplicatasImpressao.Visible = False

            'pObjTitulo.Bordero.CodigoBordero estara com o valor de -1 na hora da criacao
            If pObjTitulo.IUD <> "I" AndAlso (pObjTitulo.Bordero.CodigoBordero <> 0 OrElse (pObjTitulo.TituloBorderos.Where(Function(s) s.Bordero.CodigoSituacao = eSituacao.Normal).Count > 0) OrElse (Not pObjTitulo.UltimoBordero Is Nothing AndAlso pObjTitulo.UltimoBordero.CodigoTituloRecompra = 0)) Then
                pnlBordero.Visible = True
                If pObjTitulo.TituloBorderos.Count > 0 Then
                    txtBordero.Text = pObjTitulo.UltimoBordero.CodigoBordero
                    lblBorderoMestre.Visible = True
                    lblBorderoMestre.Text = "Titulo Bordero " & pObjTitulo.UltimoBordero.Bordero.CodigoTituloBordero
                Else
                    txtBordero.Text = pObjTitulo.Bordero.CodigoBordero
                    lblBorderoMestre.Visible = False
                    lblBorderoMestre.Text = ""
                End If

                btnImprimirBordero.Visible = False
                chkImprimirDuplicatas.Visible = False
                btnRecomprar.Visible = False

                pnlBorderoDuplicatasImpressao.Visible = True

                If pObjTitulo.Codigo = pObjTitulo.Bordero.CodigoTituloBordero Then
                    btnImprimirBordero.Visible = True
                    chkImprimirDuplicatas.Visible = True
                ElseIf pObjTitulo.TituloBorderos.Where(Function(s) s.Bordero.CodigoSituacao = eSituacao.Normal).Count > 0 Then
                    btnRecomprar.Visible = pObjTitulo.UltimoBordero.CodigoTituloRecompra = 0
                End If
            End If
            '****************************************************************************************************************************************************************

            '*******************  Datas ******************************
            txtMovimento.Text = pObjTitulo.Movimento.ToString("dd/MM/yyyy")
            txtProrrogacao.Text = pObjTitulo.Reprogramacao.ToString("dd/MM/yyyy") '*** Vencimento/Prorrogacao
            LblVencOriginal.Text = pObjTitulo.Vencimento.ToString("dd/MM/yyyy") '*** Vencimento Original
            txtDataBaixa.Text = pObjTitulo.DataBaixa.ToString("dd/MM/yyyy")
            DatadaBaixa()
            '*********************************************************

            txtHistorico.Text = pObjTitulo.Historico
            txtObservacoes.Text = pObjTitulo.Observacoes
            txtCodigoDeBarras.Text = pObjTitulo.CodigoDeBarras
            If (CBool(pObjTitulo.CodigoDeBarrasDigitado)) Then CkbCodigoDeBarras.Checked = True
            If (CBool(pObjTitulo.CodigoDeBarrasPreImpresso)) Then ckPreImpresso.Checked = True
            txtPedido.Text = IIf(pObjTitulo.CodigoPedido.Equals(0), String.Empty, pObjTitulo.CodigoPedido)

            If pObjTitulo.CodigoPedido > 0 Then
                txtPedido.Enabled = False
                cmdPedido.Enabled = False
                imgExtratoPedido.Visible = True
                ddlMoeda.Enabled = False
                ddlIndexador.Enabled = False
            Else
                txtPedido.Enabled = True
                cmdPedido.Enabled = True
            End If

            '** LANCAMENTOS CONTABEIS ***********************
            If pObjTitulo.ReceberPagar = "C" Then
                If Not pObjTitulo.ContaContabilRecPag.TemCliente Then
                    txtClienteLancContabil.Text = "CONTA SEM CLIENTE"
                    btnClienteLancamentoContabil.Enabled = False
                    txtClienteLancContabil.Enabled = False
                    divPedido.Visible = False
                Else
                    btnClienteLancamentoContabil.Enabled = True
                    txtClienteLancContabil.Enabled = True
                    If ObjTitulo.ClienteRecPag IsNot Nothing Then txtClienteLancContabil.Text = pObjTitulo.ClienteRecPag.Codigo & " - " & pObjTitulo.ClienteRecPag.CodigoEndereco & " ..." & pObjTitulo.ClienteRecPag.Nome & " / " & pObjTitulo.ClienteRecPag.Cidade & "-" & pObjTitulo.ClienteRecPag.CodigoEstado
                    divPedido.Visible = True
                End If

                If pObjTitulo.CodigoPedidoRecPag > 0 Then
                    txtPedidoRecPag.Text = pObjTitulo.CodigoPedidoRecPag
                    txtPedidoRecPag.Enabled = False
                    btnPedidoRecPag.Enabled = False
                    imgExtratoPedidoRecPag.Visible = True
                Else
                    txtPedidoRecPag.Text = ""
                    txtPedidoRecPag.Enabled = True
                    btnPedidoRecPag.Enabled = True
                End If

                If pObjTitulo.CodigoProduto.Length > 0 Then
                    txtProduto.Text = pObjTitulo.CodigoProduto & " - " & pObjTitulo.Produto.Descricao
                    txtQuantidade.Text = pObjTitulo.Quantidade.ToString("N2")
                End If

                txtContaContabilRecPag.Text = pObjTitulo.ContaContabilRecPag.Conta & " - " & pObjTitulo.ContaContabilRecPag.Titulo
                ddlLote.SelectedValue = pObjTitulo.Contabilizacoes.FirstOrDefault.Lote
                TotalizaLote()
                ddlSequencia.SelectedValue = pObjTitulo.Contabilizacoes.FirstOrDefault.Sequencia
            End If
            '*************************************************
            pObjTitulo.Valores.AtualizaValores()
            AtualizaValoresNoForm(pObjTitulo)

            txtCotacao.Text = pObjTitulo.IndiceTitulo

            'Titulo Baixado, compensado e título de pedido de troca não permite alteração.
            If (pObjTitulo.CodigoProvisao = eProvisao.Baixa OrElse pObjTitulo.CodigoProvisao = eProvisao.Compensado OrElse (pObjTitulo.Pedido IsNot Nothing AndAlso pObjTitulo.Pedido.Troca)) AndAlso pObjTitulo.Codigo > 0 Then
                DdlProvisoes.Enabled = False
                txtDataBaixa.Enabled = False
                imgBloqueio.Visible = True
                BtnAdicionarConta.Visible = False
                txtMovimento.Text = ObjTitulo.DataBaixa.ToString("dd/MM/yyyy")
                imbLimparPedido.Visible = False
                imgBloqueio.Visible = Not pObjTitulo.CodigoProvisao = eProvisao.Compensado
                imgBloqueio.Visible = Not (pObjTitulo.Pedido IsNot Nothing AndAlso pObjTitulo.Pedido.Troca)
                lnkNovo.Parent.Visible = False
            Else
                DdlProvisoes.Enabled = True
                lnkNovo.Parent.Visible = True
                lnkExcluir.Parent.Visible = True
                txtDataBaixa.Enabled = True
                imgBloqueio.Visible = False
                BtnAdicionarConta.Visible = True
                imbLimparPedido.Visible = True
                'Adiciona as contas de juros e desconto
                If Not pObjTitulo.ContaContabilCliFor.Adiantamento AndAlso Not RdContabil.Checked Then
                    AdicionaContaJurosDesconto(pObjTitulo)
                End If
            End If

            CarregarParcelasDoPedido(pObjTitulo.CodigoPedido, pObjTitulo.Codigo)
            CarregarContabilizacao(pObjTitulo.Codigo)
            txtRegistro.Enabled = False
            lnkRecibo.Parent.Visible = True
            'Na consulta não permite a troca lançamento
            If ((pObjTitulo.Codigo > 0 AndAlso (pObjTitulo.CodigoProvisao = eProvisao.Baixa OrElse pObjTitulo.CodigoProvisao = eProvisao.Compensado)) Or pObjTitulo.TitulosAgrupados.Count > 1) Then
                Select Case pObjTitulo.ReceberPagar
                    Case "R"
                        RdReceber.Checked = True
                    Case "P"
                        txtContaContabilCliFor.Enabled = False
                        btnPlanoDeContas.Visible = False
                        txtCliente.Enabled = False
                        btnClientesTitulo.Visible = False
                        RdPagar.Checked = True
                    Case "C"
                        RdContabil.Checked = True
                End Select
            End If
            'Não Permite a troca do tipo do título
            RdPagar.Enabled = False
            RdContabil.Enabled = False
            RdReceber.Enabled = False
            'Caso o titulo tenha pedido vínculado não pode alterar o fornecedor
            If (pObjTitulo.Pedido IsNot Nothing AndAlso pObjTitulo.Pedido.Codigo > 0) Then
                txtCliente.Enabled = False
                btnClientesTitulo.Visible = False
            End If

            'PARA PRODUÇÃO COLOCAR REGRA ABAIXO.
            'If Not pObjTitulo.Bloqueio Then
            '    lblAcaoTitulo.Visible = True
            'Else
            '    lblAcaoTitulo.Visible = False
            'End If
            'Dim TituloXHistorico As New Novo.ListTituloXHistorico(pObjTitulo)

            'If TituloXHistorico.Count > 0 AndAlso Not TituloXHistorico.First().Acao.Equals("LIBERAR") Then
            '    lnkNovo.Parent.Visible = False
            '    lnkExcluir.Parent.Visible = False
            '    lblAcaoTitulo.Visible = True
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message.ToString)
        End Try
    End Sub

    Public Sub CarregarParcelasDoPedido(ByVal pCodigoPedido As Integer, ByVal pCodigoDoTitulo As Integer)
        Try
            Sql = "SELECT CR.Titulo_Id AS Registro, " & vbCrLf & _
                          "       CR.Reprogramacao as Vencimento," & vbCrLf & _
                          "       case When CR.Provisao = 1" & vbCrLf & _
                          "             then CR.DataBaixa" & vbCrLf & _
                          "             else CR.Reprogramacao" & vbCrLf & _
                          "       end Baixa," & vbCrLf & _
                          "       CR.Historico," & vbCrLf & _
                          "       Valores.ValorDoDocumento," & vbCrLf & _
                          "       Valores.Deducoes," & vbCrLf & _
                          "       Valores.Acrescimos," & vbCrLf & _
                          "       Valores.ValorLiquido," & vbCrLf & _
                          "       0.0 as Saldo," & vbCrLf & _
                          "       CR.Provisao," & vbCrLf & _
                          "       P.Descricao as Situacao," & vbCrLf & _
                          "       M.Descricao as Moeda" & vbCrLf & _
                          "  FROM Titulos CR" & vbCrLf & _
                          " INNER JOIN Moedas M" & vbCrLf & _
                          "    on M.Moeda_id = CR.Moeda" & vbCrLf & _
                          " INNER JOIN Provisoes P" & vbCrLf & _
                          "    on P.Provisao_id = CR.Provisao" & vbCrLf & _
                          " Inner Join(" & vbCrLf & _
                          "             Select T.Titulo_Id," & vbCrLf & _
                          "                    SUM(case" & vbCrLf & _
                          "                          when Tc.Conta_Id  = T.ContacontabilCliFor and T.RecPag in ('C','P') and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
                          "                          when Tc.Conta_Id  = T.ContacontabilCliFor and T.RecPag = 'R'        and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
                          "                          else 0" & vbCrLf & _
                          "                        end" & vbCrLf & _
                          "                        ) as ValorDoDocumento," & vbCrLf & _
                          "                    SUM(case" & vbCrLf & _
                          "                          when Tc.Conta_Id  not in (T.ContacontabilCliFor,T.ContaContabilRecPag) and T.RecPag in ('C','P') and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
                          "                          when Tc.Conta_Id  not in (T.ContacontabilCliFor,T.ContaContabilRecPag) and T.RecPag = 'R'        and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
                          "                          else 0" & vbCrLf & _
                          "                        end" & vbCrLf & _
                          "                        ) as Acrescimos," & vbCrLf & _
                          "                     SUM(case" & vbCrLf & _
                          "                          when Tc.Conta_Id  not in (T.ContacontabilCliFor,T.ContaContabilRecPag) and T.RecPag in ('C','P') and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
                          "                          when Tc.Conta_Id  not in (T.ContacontabilCliFor,T.ContaContabilRecPag) and T.RecPag = 'R'        and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
                          "                          else 0" & vbCrLf & _
                          "                        end" & vbCrLf & _
                          "                        ) as Deducoes," & vbCrLf & _
                          "                     SUM(case" & vbCrLf & _
                          "                          when Tc.Conta_Id  = T.ContaContabilRecPag and T.RecPag in ('C','P') and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
                          "                          when Tc.Conta_Id  = T.ContaContabilRecPag and T.RecPag = 'R'        and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
                          "                          else 0" & vbCrLf & _
                          "                        end" & vbCrLf & _
                          "                        ) as ValorLiquido" & vbCrLf & _
                          "               from Titulos T" & vbCrLf & _
                          "              inner Join TitulosxContaContabil Tc" & vbCrLf & _
                          "                 on Tc.Titulo_Id = T.Titulo_Id" & vbCrLf & _
                          "              INNER JOIN Moedas M" & vbCrLf & _
                          "                 on M.Moeda_id = T.Moeda" & vbCrLf & _
                          "              Group by  T.Titulo_Id" & vbCrLf & _
                          "            ) Valores" & vbCrLf & _
                          "    on Valores.Titulo_Id = CR.Titulo_Id" & vbCrLf & _
                          " WHERE (" & IIf(pCodigoPedido = 0, "", "CR.Pedido  = " & pCodigoPedido & " or ") & "CR.Titulo_Id = " & pCodigoDoTitulo & ") " & vbCrLf & _
                          "   and CR.Situacao = 1" & vbCrLf & _
                          " ORDER BY CR.Provisao, CR.Reprogramacao" & vbCrLf

            Dim ds As New DataSet
            ds = Banco.ConsultaDataSet(Sql, "Pedidos")

            Dim TotalParcelas As Decimal = 0
            For Each dra As DataRow In ds.Tables(0).Rows
                TotalParcelas += dra("ValorDoDocumento")
            Next

            For Each dra As DataRow In ds.Tables(0).Rows
                If dra("Provisao") = 1 Then
                    TotalParcelas -= dra("ValorDoDocumento")
                    dra("Saldo") = TotalParcelas
                Else
                    dra("Saldo") = TotalParcelas
                End If
            Next

            GridParcelas.DataSource = ds
            GridParcelas.DataBind()

            DdlUnidadeDeNegocioEmpresaCliente.Enabled = False
            DdlEmpresaCliente.Enabled = False
            btnClientesTitulo.Enabled = False
            cmdPedido.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub AtualizaValoresNoForm(ByRef pObjTitulo As Novo.TituloNovo)
        Try
            gridValores.Columns(2).Visible = True
            gridValores.Columns(3).Visible = True
            gridValores.Columns(5).Visible = True

            If pObjTitulo.ReceberPagar <> "C" Then gridValores.Columns(5).Visible = False
            Dim Cotacao As New [Lib].Negocio.Cotacao(pObjTitulo.CodigoIndexador, DateTime.Now)
            If Not pObjTitulo.Valores Is Nothing Then
                If pObjTitulo.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    gridValores.Columns(3).Visible = False
                Else
                    If pObjTitulo.Pedido IsNot Nothing AndAlso pObjTitulo.Pedido.TemVariacao Then
                        'Variação Passiva.
                        Dim vlrDolar As Decimal = Cotacao.Indice
                        Dim vlrTituloAtualizado = ObjTitulo.Valores.EncargoValorDocumento.ValorOficial - (pObjTitulo.Valores.EncargoValorDocumento.ValorMoeda * vlrDolar)
                        Dim TituloXContaContabil As New Novo.TituloXContaContabil(pObjTitulo)

                        TituloXContaContabil.IUD = "I"
                        TituloXContaContabil.ValorMoeda = 0

                        If vlrTituloAtualizado > 0 Then
                            TituloXContaContabil.DC = IIf(ObjTitulo.ReceberPagar = "P", "C", "D")
                            TituloXContaContabil.CodigoContaEncargo = IIf(ObjTitulo.ReceberPagar = "P", pObjTitulo.Empresa.Empresa.CodigoContaVariacaoMonetariaAtiva, pObjTitulo.Empresa.Empresa.CodigoContaVariacaoMonetariaPassiva)
                            TituloXContaContabil.ValorOficial = vlrTituloAtualizado
                        Else
                            TituloXContaContabil.DC = IIf(ObjTitulo.ReceberPagar = "P", "D", "C")
                            TituloXContaContabil.CodigoContaEncargo = IIf(ObjTitulo.ReceberPagar = "P", pObjTitulo.Empresa.Empresa.CodigoContaVariacaoMonetariaPassiva, pObjTitulo.Empresa.Empresa.CodigoContaVariacaoMonetariaAtiva)
                            TituloXContaContabil.ValorOficial = vlrTituloAtualizado * (-1)
                        End If

                        pObjTitulo.Valores.RemoveAll(Function(s) s.CodigoContaEncargo = TituloXContaContabil.CodigoContaEncargo)
                        'If Not pObjTitulo.Valores.Where(Function(s) s.CodigoContaEncargo = TituloXContaContabil.CodigoContaEncargo).Count = 1 Then
                        pObjTitulo.Valores.Add(TituloXContaContabil)
                        pObjTitulo.Valores.AtualizaLiquido()
                        'End If
                    End If
                End If

                lblTotalDeCreditos.Text = "Total de Débitos: " & pObjTitulo.Valores.Where(Function(E) E.DC = "D").Sum(Function(V) V.ValorOficial).ToString("N2")
                lblTotalDeDebitos.Text = "Total de Créditos: " & pObjTitulo.Valores.Where(Function(E) E.DC = "C").Sum(Function(V) V.ValorOficial).ToString("N2")
                gridValores.DataSource = pObjTitulo.Valores
                lblTotalDeCreditos.Visible = True
                lblTotalDeDebitos.Visible = True
            Else
                gridValores.DataSource = Nothing
            End If
            gridValores.DataBind()

            'txtCotacao.Text = pObjTitulo.IndiceTitulo
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub AtualizaValoresAdiantamentoNoForm(ByRef pObjtitulo As Novo.TituloNovo)
        Try
            divAdiantamento.Visible = False
            divBaixaAdiantamento.Visible = False

            gridBaixasAdiantamentos.DataSource = Nothing
            gridBaixasAdiantamentos.DataBind()
            gridAdiantamentosDisponiveis.Visible = False
            gridAdiantamentosDisponiveis.DataSource = Nothing
            gridAdiantamentosDisponiveis.DataBind()
            gridAdiantamentosDisponiveis.Visible = False

            If pObjtitulo.CodigoProvisao = eProvisao.Baixa Then
                txtVencimentoAdiantamento.Text = pObjtitulo.Adiantamento.Vencimento.ToString("dd/MM/yyyy")
                txtTaxaJuroAdiantamento.Text = pObjtitulo.Adiantamento.Taxa
            End If

            If (pObjtitulo.ContaContabilRecPag Is Nothing) Or pObjtitulo.ReceberPagar = "C" Then Exit Sub

            Select Case pObjtitulo.ReceberPagar
                Case "R"
                    Select Case pObjtitulo.CodigoContaContabilCliFor.Substring(0, 1)
                        Case "1" : divBaixaAdiantamento.Visible = True
                        Case "2" : divAdiantamento.Visible = False 'Solicitação Ezer(Verificar com o Edson(NGS))
                    End Select
                    If pObjtitulo.ContaContabilRecPag.Adiantamento Then divBaixaAdiantamento.Visible = True
                Case "P"
                    Select Case pObjtitulo.CodigoContaContabilCliFor.Substring(0, 1)
                        Case "1" : divAdiantamento.Visible = True
                        Case "2" : divBaixaAdiantamento.Visible = True
                    End Select
                    If pObjtitulo.ContaContabilRecPag.Adiantamento Then divBaixaAdiantamento.Visible = True
            End Select

            If divBaixaAdiantamento.Visible Then
                If pObjtitulo.Codigo > 0 AndAlso pObjtitulo.Baixas_AdiantamentoEfetuadas.Count > 0 Then
                    lblAdiantamento.Text = "Baixas de Adiantamentos Efetuadas"

                    gridBaixasAdiantamentos.DataSource = pObjtitulo.Baixas_AdiantamentoEfetuadas.ToArray
                    gridBaixasAdiantamentos.DataBind()
                    gridAdiantamentosDisponiveis.Visible = True
                    gridBaixasAdiantamentos.Columns(3).Visible = True
                    gridBaixasAdiantamentos.Columns(4).Visible = True

                    If pObjtitulo.Moeda.Classificacao = eTiposMoeda.Oficial Then
                        gridBaixasAdiantamentos.Columns(4).Visible = False
                    Else
                        gridBaixasAdiantamentos.Columns(3).Visible = False
                    End If
                Else
                    lblAdiantamento.Text = "Adiantamentos com Saldo para Baixar"
                    If pObjtitulo.CodigoCliFor.Length = 0 Then
                        gridAdiantamentosDisponiveis.DataSource = Nothing
                    Else
                        If (String.IsNullOrWhiteSpace(txtPedido.Text)) Then
                            gridAdiantamentosDisponiveis.DataSource = pObjtitulo.AdiantamentosAbertos
                        Else
                            gridAdiantamentosDisponiveis.DataSource = pObjtitulo.AdiantamentosAbertos.Where(Function(s) s.Titulo.CodigoPedido = 0)
                        End If

                        If Not pObjtitulo.AdiantamentosAbertos.AtualizarJuroEVariacaoDosAdiantamentos Then
                            MsgBox(Me.Page, "Erro ao Atualizar os Juros e Variacoes dos Adiantamentos")
                        End If
                    End If

                    gridAdiantamentosDisponiveis.DataBind()
                    gridAdiantamentosDisponiveis.Visible = True
                    divBaixaAdiantamento.Visible = pObjtitulo.AdiantamentosAbertos.Count > 0
                End If
            Else
                gridAdiantamentosDisponiveis.DataSource = Nothing
                gridAdiantamentosDisponiveis.DataBind()
                gridBaixasAdiantamentos.DataSource = Nothing
                gridBaixasAdiantamentos.DataBind()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub AtualizaDadosBancariosNoForm(ByRef pObjTitulo As Novo.TituloNovo)
        Try
            If pObjTitulo.CodigoBancoCliFor > 0 And Not pObjTitulo.AgenciaCliFor Is Nothing Then
                lblBanco.Text = pObjTitulo.CodigoBancoCliFor.ToString & " | " & pObjTitulo.BancoCliFor.Descricao
                lblAgencia.Text = pObjTitulo.CodigoAgenciaCliFor & "-" & pObjTitulo.DigitoAgenciaCliFor & " | " & pObjTitulo.AgenciaCliFor.Praca
                lblContaCorrente.Text = pObjTitulo.ContaCliFor & "-" & pObjTitulo.DigitoContaCliFor
            Else
                lblBanco.Text = "Banco"
                lblAgencia.Text = "Agencia"
                lblContaCorrente.Text = "Conta"
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub ConsultarTitulo(ByVal Codigo As Integer)
        Try
            If Codigo = 0 Then
                MsgBox(Me.Page, "Informe o numero do registro para consulta.")
                Exit Sub
            End If
            ObjTitulo = New Novo.TituloNovo(Codigo)

            If ObjTitulo.Empresa Is Nothing Then
                MsgBox(Me.Page, "Registro não encontrado.")
                Exit Sub
            Else
                ObjTitulo.TituloOriginal = New Novo.TituloNovo(Codigo)
                CarregarFormularioComAClasse(ObjTitulo)
                SessaoSalvaTitulo()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

    Function ValidaData(ByVal Data As String, ByVal Tipo As String, ByVal Empresa As String, ByVal EndEmpresa As String) As String
        If Not IsDate(Data) Then Return " Data de " & Tipo & " inválida..."
        If CDate(Data).DayOfWeek = 6 Then Return "Sábado - Data Inválida para " & Tipo & "."
        If CDate(Data).DayOfWeek = 0 Then Return "Domingo - Data Inválida para " & Tipo & "."

        If Funcoes.VerificaDatasNaoProgramaveis("99999999999999", 0, Data) Then
            Return "Data de " & Tipo & " não programável, feriado nacional!"
        End If

        If Empresa <> "" Then
            If Funcoes.VerificaDatasNaoProgramaveis(Empresa, EndEmpresa, Data) Then
                Return "Data de " & Tipo & " não programável, feriado municipal!"
            End If
        End If

        If Tipo = "Movimento" Then
            If Not Funcoes.VerificaAcesso(Empresa, EndEmpresa, Data, "Financeiro") Then
                Return "Movimento já fechado para esta data."
            End If
        End If

        Return ""
    End Function

    Public Sub DatadaBaixa()
        Try
            If ObjTitulo.CodigoProvisao = eProvisao.Baixa Then
                divDataBaixa.Visible = True
            Else
                divDataBaixa.Visible = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#Region "Rádio Receber / Pagar"

    Private Sub CarregarContaContabilEmpresaRecPag(ByVal RP As String)
        Dim Where As String = ""
        Where = " exists(" & vbCrLf & _
                       "        select 1" & vbCrLf & _
                       "          from BancosXContas BC" & vbCrLf & _
                       "         where BC.ContaContabil = planodecontas.Conta_Id" & vbCrLf & _
                       "           and BC.Empresa_Id    ='" & ObjTitulo.CodigoEmpresaRecPag & "'" & vbCrLf & _
                       "           and BC.EndEmpresa_Id = " & ObjTitulo.EndEmpresaRecPag & vbCrLf & _
                       "           and ISNULL(BC.Ativo,0) = 1" & vbCrLf & _
                       "         ) OR Conta_Id = (select ContaGrupoBanco" & vbCrLf & _
                       "                            from ClientesXEmpresas" & vbCrLf & _
                       "                           where Empresa_Id    ='" & ObjTitulo.CodigoEmpresaRecPag & "'" & vbCrLf & _
                       "                             and EndEmpresa_Id = " & ObjTitulo.EndEmpresaRecPag & ")"

        If ObjTitulo.Codigo > 0 AndAlso ObjTitulo.ContaContabilRecPag IsNot Nothing Then Where &= " OR Conta_Id = " & ObjTitulo.ContaContabilRecPag.Conta
        'Where = " AND LEN(Conta_Id) in(7,9) "
        ddl.Carregar(ddlContaContabilEmpresaRecPag, CarregarDDL.Tabela.PlanoDeContas, Where, True)
    End Sub

    Private Sub CarregarContaContabilEmpresaRecPagBaixarTodos(ByVal RP As String)
        Dim emp As String() = ddlEmpresaBaixaTodos.SelectedValue.Split("-")

        Dim Where As String = ""
        Select Case RP
            Case "R"
                Where = " exists(" & vbCrLf & _
                        "        select 1" & vbCrLf & _
                        "          from BancosXContas BC" & vbCrLf & _
                        "         where BC.ContaContabil = planodecontas.Conta_Id" & vbCrLf & _
                        "           and BC.Empresa_Id    ='" & emp(0) & "'" & vbCrLf & _
                        "           and BC.EndEmpresa_Id = " & emp(1) & vbCrLf & _
                        "           and len(BC.ContaContabil) = 9 " & vbCrLf & _
                        "       )" & vbCrLf
            Case "P"
                Where = " exists(" & vbCrLf & _
                        "        select 1" & vbCrLf & _
                        "          from BancosXContas BC" & vbCrLf & _
                        "         where BC.ContaContabil      = planodecontas.Conta_Id" & vbCrLf & _
                        "           and BC.Empresa_Id         ='" & emp(0) & "'" & vbCrLf & _
                        "           and BC.EndEmpresa_Id      = " & emp(1) & vbCrLf & _
                        "           and len(BC.ContaContabil) = 9 " & vbCrLf & _
                        "       )" & vbCrLf
        End Select
        ddl.Carregar(ddlContaContabilEmpresaRecPagBaixarTodos, CarregarDDL.Tabela.PlanoDeContas, Where)
    End Sub

    Public Sub MudouRP(ByVal RP As String)
        Try
            CarregarContaContabilEmpresaRecPag(RP)
            lblCliFor.Parent.Visible = True
            divBanco.Visible = True
            divPedido.Visible = True
            RdReceber.Checked = False
            RdPagar.Checked = False
            RdContabil.Checked = False
            txtContaContabilRecPag.Visible = False
            btnPlanoDeContaRecPag.Visible = False
            ddlContaContabilEmpresaRecPag.Visible = True
            txtCotacao.Parent.Visible = True
            ddlContaContabilEmpresaRecPag.Visible = True
            divTipoDeRecebimento.Visible = True
            divCarteiraDoTitulo.Visible = True
            divPrevisaoBaixa.Visible = True
            divVencOriginal.Visible = True
            DdlProvisoes.ClearSelection()
            DdlTiposDeRecebimentos.ClearSelection()
            ddlCarteiraDoTitulo.ClearSelection()
            If RP = "R" Then
                RdReceber.Checked = True
                ObjTitulo.ReceberPagar = "R"
                LblRecPag.Text = "Contas A Receber"
                LblCliForTit.Text = "Dados Do Cliente"
                lblCliFor.Text = "Cliente:"
                lblEmpresaRecPag.Text = "Empresa Recebedora"
                divLancamentoContabil.Visible = False
                divProduto.Visible = False
                divQuantidade.Visible = False
                ddlLote.Parent.Visible = False
                ddlSequencia.Parent.Visible = False
                txtClienteLancContabil.Parent.Visible = False
                divPedidoRecPag.Visible = False
                divCodigoDeBarras.Visible = True
                DdlProvisoes.Enabled = True
                DdlProvisoes.ClearSelection()
                txtDepositoCliente.Parent.Visible = False
                txtDepositoRecPag.Parent.Visible = False
                lblTipoRecPag.Text = "Tipo Recebimento:"
                divBanco.Visible = False
                divAdiantamento.Visible = False
            ElseIf RP = "P" Then
                RdPagar.Checked = True
                ObjTitulo.ReceberPagar = "P"
                LblRecPag.Text = "Contas A Pagar"
                LblCliForTit.Text = "Dados Do Fornecedor"
                lblCliFor.Text = "Fornecedor:"
                lblEmpresaRecPag.Text = "Empresa Pagadora"
                divLancamentoContabil.Visible = False
                divProduto.Visible = False
                divQuantidade.Visible = False
                ddlLote.Parent.Visible = False
                ddlSequencia.Parent.Visible = False
                txtClienteLancContabil.Parent.Visible = False
                divPedidoRecPag.Visible = False
                divBanco.Visible = True
                divCodigoDeBarras.Visible = True
                DdlProvisoes.Enabled = True
                DdlProvisoes.ClearSelection()
                txtDepositoCliente.Parent.Visible = False
                txtDepositoRecPag.Parent.Visible = False
                lblTipoRecPag.Text = "Tipo Pagamento:"
            ElseIf RP = "C" Then
                RdContabil.Checked = True
                ObjTitulo.ReceberPagar = "C"
                LblRecPag.Text = "Lançamento Contábil"
                LblCliForTit.Text = "Débito"
                lblCliFor.Text = "Cliente:"
                lblEmpresaRecPag.Text = "Crédito"
                divLancamentoContabil.Visible = True
                divProduto.Visible = True
                divQuantidade.Visible = True
                ddlLote.Parent.Visible = True
                ddlSequencia.Parent.Visible = True
                txtClienteLancContabil.Parent.Visible = True
                divPedidoRecPag.Visible = True
                divBanco.Visible = False
                divCodigoDeBarras.Visible = False
                DdlProvisoes.SelectedValue = 1 'Lançamento Contábil é sempre Baixa
                DdlTiposDeRecebimentos.SelectedValue = 1 'Lançamento Contábil é sempre NORMAL
                ObjTitulo.CodigoTipoPgto = DdlTiposDeRecebimentos.SelectedValue
                ddlCarteiraDoTitulo.SelectedValue = 0 'lançamento Contábil e sempre NENHUM
                DdlProvisoes.Enabled = False
                SessaoRecuperaTitulo()
                ObjTitulo.CodigoProvisao = eProvisao.Baixa
                ObjTitulo.DataBaixa = Date.Now
                txtDataBaixa.Text = ObjTitulo.DataBaixa.ToString("dd/MM/yyyy")
                DatadaBaixa()
                SessaoSalvaTitulo()
                lblTipoRecPag.Text = "Tipo:"
                txtDepositoCliente.Parent.Visible = True
                txtDepositoRecPag.Parent.Visible = True
                txtContaContabilRecPag.Visible = True
                btnPlanoDeContaRecPag.Visible = True
                txtCotacao.Parent.Visible = False
                ddlContaContabilEmpresaRecPag.Visible = False
                divTipoDeRecebimento.Visible = False
                divCarteiraDoTitulo.Visible = False
                divPrevisaoBaixa.Visible = False
                divVencOriginal.Visible = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Consulta Títulos"

    Public Function TitulosConsultaBordero() As Boolean
        Try
            Dim sql As String
            sql = GetSqlBordero()

            gridBordero.DataSource = Banco.ConsultaDataSet(sql, "Bordero")
            gridBordero.DataBind()

            If gridBordero.Rows.Count = 0 Then
                MsgBox(Me.Page, "Nenhum registro encontrado.")
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Function

    Public Function TitulosConsulta(Optional ByVal Where As Hashtable = Nothing) As Boolean
        Try
            Sql = getSqlConsulta(Where)
            ds = Banco.ConsultaDataSet(Sql, "Contas")

            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Nenhum Registro encontrado")
                Return False
            Else
                GridConsultaTitulos.DataSource = ds
                GridConsultaTitulos.DataBind()
                'checa o primeiro titulo selecionado no agrupamento no topo da selecao
                If rdAgrupar.Checked AndAlso ObjTitulosAcao.Count > 0 Then
                    Dim chk As CheckBox
                    chk = GridConsultaTitulos.Rows(0).FindControl("ChkGridTitulos")
                    chk.Checked = True
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
        Return True
    End Function

    Private Function getSqlConsulta(Optional ByVal Where As Hashtable = Nothing, Optional ByVal isRecompra As Boolean = False) As String
        Dim Cliente As String
        Dim Campo() As String
        Dim i As Integer = 0
        Dim Unidade As String = DdlUnidadeConsultaTitulos.SelectedValue
        Dim Provisao As String = ""
        Dim RecPagCont As String = ""

        Sql = "SELECT Case " & vbCrLf & _
              "          when T.Titulo_Id = isnull(T.RegistroMestre,0)                                                                      then (SELECT Observacoes FROM Titulos WHERE Titulo_Id = T.RegistroMestre) " & vbCrLf & _
              "          when T.Titulo_Id = isnull(BorderoMestre.TituloBordero,0)                                                           then (SELECT Observacoes FROM Titulos WHERE Titulo_Id = BorderoMestre.TituloBordero) " & vbCrLf & _
              "          When isnull(T.RegistroMestre,0) > 0                                                                                then 'Título Agrupado ao Registro ' + CONVERT(varchar,T.RegistroMestre) " & vbCrLf & _
              "          When isnull(B.Bordero_Id,0)     > 0 and (isnull(BxTa.TituloRecompra,0) = 0 or isnull(BxTa.BorderoRecompra,0) = 0 ) then 'Título no Borderô: ' + + CONVERT(varchar,B.Bordero_Id) " & vbCrLf & _
              "          When isnull(T.RegistroMestre,0) = 0                                                                                then '' " & vbCrLf & _
              "       End as Agrupamento," & vbCrLf & _
              "       ISNULL(RegistroMestre,0) AS RegistroMestre," & vbCrLf & _
              "       case T.RecPag" & vbCrLf & _
              "          when 'P' then 'PAGAR'" & vbCrLf & _
              "          when 'R' then 'RECEBER'" & vbCrLf & _
              "          when 'C' then 'CONTABIL'" & vbCrLf & _
              "       End as Doc," & vbCrLf & _
              "       T.Titulo_Id," & vbCrLf & _
              "       case" & vbCrLf & _
              "          when T.Titulo_Id = isnull(BxTa.TituloRecompra,0)" & vbCrLf & _
              "            then 0" & vbCrLf & _
              "            else isnull(BxTa.TituloRecompra,0)" & vbCrLf & _
              "       end as TituloRecompra," & vbCrLf & _
              "       isnull(BxTa.BorderoRecompra,0) as BorderoRecompra," & vbCrLf & _
              "       T.ContaContabilCliFor," & vbCrLf & _
              "       T.ContaContabilRecPag, " & vbCrLf & _
              "       E.Nome AS Empresa, " & vbCrLf & _
              "       E.Cliente_Id AS EndEmpresa, " & vbCrLf & _
              "       E.Cidade, " & vbCrLf & _
              "       E.Estado, " & vbCrLf & _
              "       C.Nome, " & vbCrLf & _
              "       C.Cliente_Id, " & vbCrLf & _
              "       T.Reprogramacao as Vencimento," & vbCrLf & _
              "       case When T.Provisao = 1" & vbCrLf & _
              "             then T.DataBaixa" & vbCrLf & _
              "             else T.Reprogramacao" & vbCrLf & _
              "       end Baixa," & vbCrLf & _
              "       T.Historico," & vbCrLf & _
              "       Valores.ValorDoDocumento," & vbCrLf & _
              "       Valores.Deducoes," & vbCrLf & _
              "       Valores.Acrescimos," & vbCrLf & _
              "       Valores.ValorLiquido," & vbCrLf & _
              "       P.Descricao as Situacao," & vbCrLf & _
              "       M.Descricao as Moeda," & vbCrLf & _
              "       isnull(T.Cheque,0) as Cheque," & vbCrLf & _
              "       isnull(BorderoMestre.Bordero_Id,0) + isnull(B.Bordero_Id,0) as Bordero," & vbCrLf & _
              "       isnull(Ca.EmiteDuplicata,0) as EmiteDuplicata" & vbCrLf & _
              "  FROM Titulos T" & vbCrLf & _
              " INNER JOIN Clientes C " & vbCrLf & _
              "    ON T.CliFor         = C.Cliente_Id " & vbCrLf & _
              "   AND T.EnderecoCliFor = C.Endereco_Id" & vbCrLf & _
              "INNER JOIN Clientes E " & vbCrLf & _
              "      ON T.Empresa         = E.Cliente_Id " & vbCrLf & _
              "      AND T.EndEmpresa = E.Endereco_Id " & vbCrLf & _
              "  Left Join(" & vbCrLf & _
              "            Select BxT.Empresa_Id, BxT.EndEmpresa_Id, BxT.Titulo_id, max(BxT.Bordero_Id) as Bordero_Id" & vbCrLf & _
              "              from BorderoXTitulo BxT" & vbCrLf & _
              "             inner join Bordero B" & vbCrLf & _
              "				   on B.Empresa_Id    = BxT.Empresa_Id" & vbCrLf & _
              "			      and B.EndEmpresa_Id = BxT.EndEmpresa_Id" & vbCrLf & _
              "			      and B.Bordero_Id    = BxT.Bordero_id" & vbCrLf

        If Not Where Is Nothing AndAlso Where.ContainsKey("Bordero") Then
            Sql &= "             where BxT.Empresa_Id    ='" & Where("Empresa").ToString & "'" & vbCrLf & _
                   "               And BxT.EndEmpresa_Id = " & Where("EndEmpresa").ToString & vbCrLf & _
                   "               And BxT.Bordero_Id    = " & Where("Bordero").ToString & vbCrLf
        Else
            Sql &= "             where B.Situacao = 1 " & vbCrLf
        End If

        Sql &= "             group by BxT.Empresa_Id, BxT.EndEmpresa_Id, BxT.Titulo_id " & vbCrLf & _
              "            ) BxT" & vbCrLf & _
              "    on T.Titulo_Id = BxT.Titulo_Id" & vbCrLf & _
              "  left join Bordero B " & vbCrLf & _
              "    on B.Empresa_Id    = BxT.Empresa_Id" & vbCrLf & _
              "   and B.EndEmpresa_Id = BxT.EndEmpresa_Id" & vbCrLf & _
              "   and B.Bordero_Id    = BxT.Bordero_Id" & vbCrLf & _
              "  left Join Bordero BorderoMestre" & vbCrLf & _
              "    on BorderoMestre.Empresa_Id    = T.Empresa" & vbCrLf & _
              "   and BorderoMestre.EndEmpresa_Id = T.EndEmpresa" & vbCrLf & _
              "   and BorderoMestre.TituloBordero = T.Titulo_Id" & vbCrLf & _
              "  left Join BorderoXTitulo BxTa" & vbCrLf & _
              "    on B.Empresa_Id    = BxTa.Empresa_Id" & vbCrLf & _
              "   and B.EndEmpresa_Id = BxTa.EndEmpresa_Id" & vbCrLf & _
              "   and B.Bordero_Id    = BxTa.Bordero_Id" & vbCrLf & _
              "   and T.Titulo_Id     = BxTa.Titulo_Id" & vbCrLf & _
              " Inner Join Carteira Ca" & vbCrLf & _
              "    on Ca.Carteira_Id = T.CarteiraDoTitulo" & vbCrLf & _
              " INNER JOIN Moedas M" & vbCrLf & _
              "    on M.Moeda_id = T.Moeda" & vbCrLf & _
              " INNER JOIN Provisoes P" & vbCrLf & _
              "    on P.Provisao_id = T.Provisao" & vbCrLf & _
              " Inner Join(" & vbCrLf & _
              "             Select Tp.Titulo_Id," & vbCrLf & _
              "                    SUM(case" & vbCrLf & _
              "                          when Tc.Conta_Id  = Tp.ContacontabilCliFor and Tp.RecPag in ('C','P') and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
              "                          when Tc.Conta_Id  = Tp.ContacontabilCliFor and Tp.RecPag = 'R'        and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
              "                          else 0" & vbCrLf & _
              "                        end" & vbCrLf & _
              "                        ) as ValorDoDocumento," & vbCrLf & _
              "                    SUM(case" & vbCrLf & _
              "                          when Tc.Conta_Id  not in (Tp.ContacontabilCliFor,Tp.ContaContabilRecPag) and Tp.RecPag in ('C','P') and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
              "                          when Tc.Conta_Id  not in (Tp.ContacontabilCliFor,Tp.ContaContabilRecPag) and Tp.RecPag = 'R'        and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
              "                          else 0" & vbCrLf & _
              "                        end" & vbCrLf & _
              "                        ) as Acrescimos," & vbCrLf & _
              "                     SUM(case" & vbCrLf & _
              "                          when Tc.Conta_Id  not in (Tp.ContacontabilCliFor,Tp.ContaContabilRecPag) and Tp.RecPag in ('C','P') and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
              "                          when Tc.Conta_Id  not in (Tp.ContacontabilCliFor,Tp.ContaContabilRecPag) and Tp.RecPag = 'R'        and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
              "                          else 0" & vbCrLf & _
              "                        end" & vbCrLf & _
              "                        ) as Deducoes," & vbCrLf & _
              "                     SUM(case" & vbCrLf & _
              "                          when Tc.Conta_Id  = Tp.ContaContabilRecPag and Tp.RecPag in ('C','P') and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
              "                          when Tc.Conta_Id  = Tp.ContaContabilRecPag and Tp.RecPag = 'R'        and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end" & vbCrLf & _
              "                          else 0" & vbCrLf & _
              "                        end" & vbCrLf & _
              "                        ) as ValorLiquido" & vbCrLf & _
              "               from Titulos Tp" & vbCrLf & _
              "              inner Join TitulosxContaContabil Tc" & vbCrLf & _
              "                 on Tc.Titulo_Id = Tp.Titulo_Id" & vbCrLf & _
              "              INNER JOIN Moedas M" & vbCrLf & _
              "                 on M.Moeda_id = Tp.Moeda" & vbCrLf & _
              "              Group by Tp.Titulo_Id" & vbCrLf & _
              "            ) Valores" & vbCrLf & _
              "    on Valores.Titulo_Id = T.Titulo_Id" & vbCrLf

        If Not Where Is Nothing AndAlso Where.ContainsKey("Bordero") Then
            Sql &= "             Where T.Empresa    ='" & Where("Empresa").ToString & "'" & vbCrLf & _
                   "               And T.EndEmpresa = " & Where("EndEmpresa").ToString & vbCrLf & _
                   "               And isnull(BorderoMestre.Bordero_Id,0) + isnull(B.Bordero_Id,0) = " & Where("Bordero").ToString & vbCrLf
            ds = Banco.ConsultaDataSet(Sql, "Contas")
            If Not ds.Tables(0).Rows.Count = 0 Then
                Return Sql
                '    GridConsultaTitulos.DataBind()
                '    MsgBox(Me.Page, "Nenhum Registro encontrado")
                '    Return False
                'Else
                '    GridConsultaTitulos.DataSource = ds
                '    GridConsultaTitulos.DataBind()
            End If
            Return True
        ElseIf Not Where Is Nothing AndAlso Where.ContainsKey("Slip") Then
            Sql &= "       WHERE T.Slips = 0 " & vbCrLf & _
                    "       And T.Empresa    = '" & Where("Empresa").ToString & "'" & vbCrLf & _
                    "       And T.EndEmpresa = " & Where("EndEmpresa").ToString & vbCrLf
            Return Sql
        End If

        Sql &= " Where T.Situacao = 1" & vbCrLf

        SessaoRecuperaTitulosAcao()
        If rdAdTituloAgrupamento.Checked Then
            SessaoRecuperaTituloAgrupado()
            Sql &= " and T.UnidadeDeNegocio    ='" & ObjTituloAgrupado.CodigoUnidadeDeNegocio & "'" & vbCrLf & _
                   " and T.Empresa             ='" & ObjTituloAgrupado.CodigoEmpresa & "'" & vbCrLf & _
                   " and T.EndEmpresa          = " & ObjTituloAgrupado.EnderecoEmpresa & vbCrLf & _
                   " and T.CliFor              ='" & ObjTituloAgrupado.CodigoCliFor & "'" & vbCrLf & _
                   " and T.EnderecoCliFor      = " & ObjTituloAgrupado.EnderecoCliFor & vbCrLf & _
                   " and T.ContaContabilCliFor ='" & ObjTituloAgrupado.CodigoContaContabilCliFor & "'" & vbCrLf & _
                   " and T.Moeda               = " & ObjTituloAgrupado.CodigoMoeda & vbCrLf & _
                   " and T.Provisao            = " & ObjTituloAgrupado.CodigoProvisao & vbCrLf & _
                   " and T.RecPag              = '" & ObjTituloAgrupado.ReceberPagar & "'"
        ElseIf rdAgrupar.Checked AndAlso ObjTitulosAcao.Count > 0 Then
            Sql &= " and T.UnidadeDeNegocio    ='" & ObjTitulosAcao(0).CodigoUnidadeDeNegocio & "'" & vbCrLf & _
                   " and T.Empresa             ='" & ObjTitulosAcao(0).CodigoEmpresa & "'" & vbCrLf & _
                   " and T.EndEmpresa          = " & ObjTitulosAcao(0).EnderecoEmpresa & vbCrLf & _
                   " and T.CliFor              ='" & ObjTitulosAcao(0).CodigoCliFor & "'" & vbCrLf & _
                   " and T.EnderecoCliFor      = " & ObjTitulosAcao(0).EnderecoCliFor & vbCrLf
            If ObjTitulosAcao(0).CodigoPedido = 0 Then
                Sql &= " and T.ContaContabilCliFor LIKE '" & Left(ObjTitulosAcao(0).CodigoContaContabilCliFor, 7) & "%'" & vbCrLf
            Else
                Sql &= " and T.ContaContabilCliFor ='" & ObjTitulosAcao(0).CodigoContaContabilCliFor & "'" & vbCrLf
            End If

            Sql &= " and T.Moeda               = " & ObjTitulosAcao(0).CodigoMoeda & vbCrLf & _
                   " and T.Provisao            = " & ObjTitulosAcao(0).CodigoProvisao & vbCrLf & _
                   " and T.RecPag              = '" & ObjTitulosAcao(0).ReceberPagar & "'"
        Else
            If rdDuplicatasDescontadas.Checked Then
                Sql &= " and T.RegistroMestre IS NULL" & vbCrLf & _
                       IIf(Not isRecompra, " and isnull(BorderoMestre.Bordero_Id,0) + isnull(B.Bordero_Id,0) = 0", "")
            End If

            Cliente = DdlUnidadeConsultaTitulos.SelectedValue
            If Cliente <> "" Then
                Sql &= " and T.UnidadeDeNegocio = '" & Cliente & "'"
            End If

            Cliente = DdlEmpresaConsultaTitulos.SelectedValue
            Campo = Cliente.Split("-")
            If Campo(0) <> "" Then
                Sql &= " and T.Empresa    ='" & Campo(0) & "'" 'Empresa
                Sql &= " and T.EndEmpresa = " & Campo(1)       'Endereco da Empresa
            End If

            Cliente = txtCodigoClienteConsulta.Value
            Campo = Cliente.Split("-")
            If Campo(0) <> "" Then
                Sql &= " and T.CliFor         ='" & Campo(0) & "'"   'Cliente
                Sql &= " and T.EnderecoCliFor = " & Campo(1)         'Cliente da Empresa
            End If

            Provisao = IIf(chkBaixa.Checked, "1", "")
            Provisao &= IIf(chkProvisao.Checked, IIf(Provisao.Length > 0, ",3", "3"), "")
            Provisao &= IIf(chkPrevisao.Checked, IIf(Provisao.Length > 0, ",2", "2"), "")
            Sql &= "   and T.Provisao in (" & Provisao & ")" & vbCrLf

            RecPagCont = IIf(chkReceber.Checked, "'R'", "")
            RecPagCont &= IIf(chkPagar.Checked, IIf(RecPagCont.Length > 0, ",'P'", "'P'"), "")
            RecPagCont &= IIf(chkContabil.Checked, IIf(RecPagCont.Length > 0, ",'C'", "'C'"), "")
            Sql &= "   and T.RecPag in (" & RecPagCont & ")" & vbCrLf

            If chkRecompraPorBordero.Checked Then
                Sql &= " and T.Reprogramacao < '" & Date.Now.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                       " and isnull(B.Bordero_Id,0) > 0" & vbCrLf

                If ObjTitulosAcao.Count > 0 Then
                    Sql &= "   and T.CarteiraDoTitulo = " & ObjTitulosAcao(0).CodigoCarteiraDoTitulo & vbCrLf & _
                           " order by case when T.Titulo_id = " & ObjTitulosAcao(0).Codigo & " then 0 else 1 end, isnull(B.Bordero_Id,0), T.Titulo_Id"
                Else
                    Sql &= " order by isnull(B.Bordero_Id,0), T.Titulo_Id"
                End If
                Return Sql
            End If
        End If

        If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
            Sql &= " and T.Reprogramacao between '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "' "
        End If

        'Coloca o primeiro titulo selecionado no agrupamento no topo da selecao
        If rdAgrupar.Checked AndAlso ObjTitulosAcao.Count > 0 Then
            Sql &= "order by case when T.Titulo_id = " & ObjTitulosAcao(0).Codigo & " then 0 else 1 end"
        ElseIf rdAdTituloAgrupamento.Checked Then
            'Para adcionar ou remover um titulo no Agrupamento pesquisa os titulos que não estão agrupados em outro agrupamento.
            Sql &= " AND T.RegistroMestre is null or T.RegistroMestre = " & txtTituloConsulta.Text & " ORDER BY RegistroMestre DESC, Titulo_Id ASC"
        Else
            Sql &= " ORDER BY T.Reprogramacao DESC "
        End If

        Return Sql
    End Function

    Private Function GetSqlBordero()
        Dim sql As String = String.Empty
        sql = " SELECT  " & vbCrLf & _
          " T.Empresa, " & vbCrLf & _
          " EndEmpresa," & vbCrLf & _
          " B.TituloBordero," & vbCrLf & _
          " B.Contrato,  " & vbCrLf & _
          " C.Nome, " & vbCrLf & _
          " C.Endereco, " & vbCrLf & _
          " C.Numero, " & vbCrLf & _
          " C.Complemento,  " & vbCrLf & _
          " C.Cidade, " & vbCrLf & _
          " C.Estado, " & vbCrLf & _
          " B.DataEnvio, " & vbCrLf & _
          " B.JurosTaxa, 	 " & vbCrLf & _
          " FCT.Titulo AS ContaBordero, " & vbCrLf & _
          " BCO.Titulo AS TituloBanco, " & vbCrLf & _
          " T.Historico, " & vbCrLf & _
          " TP.Descricao " & vbCrLf & _
          " FROM Bordero B " & vbCrLf & _
          " INNER JOIN Titulos T  " & vbCrLf & _
          "   ON B.TituloBordero = T.Titulo_Id " & vbCrLf & _
          " INNER JOIN Clientes C " & vbCrLf & _
          "   ON T.Empresa = C.Cliente_Id " & vbCrLf & _
          "   AND T.EndEmpresa = C.Endereco_Id " & vbCrLf & _
          " INNER JOIN PlanoDeContas AS FCT " & vbCrLf & _
          "   ON T.ContaContabilCliFor = FCT.Conta_Id " & vbCrLf & _
          " INNER JOIN PlanoDeContas AS BCO " & vbCrLf & _
          "   ON T.ContaContabilRecPag = BCO.Conta_Id " & vbCrLf & _
          " INNER JOIN TiposDePagamentos TP  " & vbCrLf & _
          "   ON T.TipoPagto = TP.TipoDePagamento_Id " & vbCrLf & _
          " WHERE " & vbCrLf & _
          " (B.TituloBordero = " & txtBorderoConsulta.Text & " OR B.Contrato = " & txtBorderoConsulta.Text & ") " & vbCrLf

        If (String.IsNullOrWhiteSpace(txtBorderoConsulta.Text)) Then
            sql &= "   AND DataEnvio BETWEEN '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf
        End If

        If (Not String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue)) Then
            sql &= " AND B.Empresa_Id = '07090163000151'" & vbCrLf & _
                   " AND B.EndEmpresa_Id = '0' "
        End If
        Return sql
    End Function

    Private Function ValidaConsulta() As Boolean
        If rdAdTituloAgrupamento.Checked Then
            If Not IsNumeric(txtTituloConsulta.Text) Then
                MsgBox(Me.Page, "Informe no código do título o número do título mestre no qual os títulos selecionados serao adicionados.")
                Return False
            Else
                Dim TitAgrupado As New Novo.TituloNovo(txtTituloConsulta.Text)
                If TitAgrupado.Codigo <> TitAgrupado.RegistroMestre Then
                    MsgBox(Me.Page, "Título informado não é um título agrupado *Mestre*")
                    Return False
                End If
                If TitAgrupado.CodigoSituacao <> eSituacao.Normal Then
                    MsgBox(Me.Page, "Título inválido.")
                    Return False
                End If
                If TitAgrupado.CodigoProvisao <> eProvisao.Provisao Then
                    MsgBox(Me.Page, "O título mestre deve estar em provisão para aceitar novos títulos.")
                    Return False
                End If

                ObjTituloAgrupado = TitAgrupado
                For Each row In TitAgrupado.TitulosAgrupados
                    row.IUD = "U"
                Next
                SessaoSalvaTituloAgrupado()
            End If
        End If
        If Not chkPagar.Checked AndAlso Not chkReceber.Checked AndAlso Not chkContabil.Checked Then
            MsgBox(Me.Page, "O campo Tipo Título é Obrigatório!")
            Return False
        End If

        If Not chkBaixa.Checked AndAlso Not chkProvisao.Checked AndAlso Not chkPrevisao.Checked Then
            MsgBox(Me.Page, "O campo Sítuação é Obrigatório!")
            Return False
        End If
        'Pode fazer borderô com titulos de empresas diferentes!
        'If rdDuplicatasDescontadas.Checked Or rdGeraDuplicatasCobrancaSimples.Checked Or rdConsultarBordero.Checked Then
        '    If DdlEmpresaConsultaTitulos.SelectedIndex = 0 Or DdlEmpresaConsultaTitulos.SelectedIndex = -1 Then
        '        MsgBox(Me.Page, "Selecione a Empresa para selecao dos Titulos que farao parte do Bordero")
        '        Return False
        '    End If
        'End If

        If Not IsDate(txtPeriodoFinalConsultaTitulos.Text) Then
            MsgBox(Me.Page, "Data final do período de: inválida")
            Return False
        End If
        If Not IsDate(txtPeriodoInicialConsultaTitulos.Text) Then
            MsgBox(Me.Page, "Data incial do período de: inválida")
            Return False
        End If

        'If (rdAgrupar.Checked AndAlso String.IsNullOrWhiteSpace(txtClienteConsulta.Text)) Then
        '    MsgBox(Me.Page, "Para realizar agrupamento o cliente é obrigatório!")
        '    Return False
        'End If

        Return True
    End Function

    Private Sub PaineisInvisiveis()
        divAcoes.Visible = True
        divAcaoAgrupar.Visible = False
        divAcaoBaixar.Visible = False
        divAcaoReprogramacao.Visible = False
        divAcaoAdicionar.Visible = False

        divAcaoDuplicatas.Visible = False
        divAcaoDuplicatasCobrancaSimples.Visible = False
        divParametrosAdicionarTituloAgrupado.Visible = False
        divAcaoConsultaBordero.Visible = False
        divRecompraPorBordero.Visible = False
    End Sub

    Protected Sub Bordero()
        Try
            Dim obsTitulosBordero As String = "Títulos do bordero: "
            Dim obsTitulosRecompra As String = "Títulos recomprados: "
            SessaoRecuperaTitulo()
            Session("TitTemp") = ObjTitulo
            Limpar()
            ObjTitulo = Session("TitTemp")
            Session.Remove("TitTemp")

            ObjTitulo.DesligarControles()
            ObjTitulo.ReceberPagar = "R"
            ObjTitulo.IUD = "I"
            ObjTitulo.Observacoes = ""
            ObjTitulo.Movimento = Date.Now
            ObjTitulo.Vencimento = Date.Now
            ObjTitulo.Reprogramacao = Date.Now

            ObjTitulo.Bordero.IUD = "I"
            ObjTitulo.Bordero.DataEnvio = Date.Now
            ObjTitulo.Bordero.CodigoEmpresa = ObjTitulo.CodigoEmpresa
            ObjTitulo.Bordero.EnderecoEmpresa = ObjTitulo.EnderecoEmpresa
            ObjTitulo.Bordero.CodigoSituacao = eSituacao.Normal 'Ativo

            For Each TituloFilho In ObjTitulo.Bordero.TitulosDoBordero
                obsTitulosBordero &= " - " & TituloFilho.CodigoTitulo
            Next

            For Each TituloRecompra In ObjTitulo.Bordero.TitulosRecomprados
                obsTitulosRecompra &= " - " & TituloRecompra.CodigoTitulo
            Next

            ObjTitulo.Historico = obsTitulosBordero & IIf(ObjTitulo.Bordero.TitulosRecomprados.Count > 1, vbCrLf & obsTitulosRecompra, "")

            '****** Alimenta principais campos do Mestre com as informacoes do primeiro titulo agrupado ******'
            ObjTitulo.ReceberPagar = ObjTitulo.Bordero.TitulosDoBordero(0).Titulo.ReceberPagar
            ObjTitulo.CodigoUnidadeDeNegocio = ObjTitulo.Bordero.TitulosDoBordero(0).Titulo.CodigoUnidadeDeNegocio
            ObjTitulo.CodigoEmpresa = ObjTitulo.Bordero.TitulosDoBordero(0).Titulo.CodigoEmpresa
            ObjTitulo.EnderecoEmpresa = ObjTitulo.Bordero.TitulosDoBordero(0).Titulo.EnderecoEmpresa

            ObjTitulo.CodigoCliFor = ObjTitulo.Bordero.TitulosDoBordero(0).Titulo.CodigoEmpresa
            ObjTitulo.EnderecoCliFor = ObjTitulo.Bordero.TitulosDoBordero(0).Titulo.EnderecoEmpresa

            'ddl.Carregar(ddlContaContabilCliFor, CarregarDDL.Tabela.PlanoDeContas, "Conta_id like '" & ObjTitulo.Bordero.TitulosDoBordero(0).Titulo.Empresa.Empresa.CodigoContaGrupoDuplicatasDescontada & "%' --and len(Conta_Id) = 9", False)
            ObjTitulo.CodigoContaContabilCliFor = ObjTitulo.Bordero.TitulosDoBordero(0).Titulo.Empresa.Empresa.CodigoContaGrupoDuplicatasDescontada

            ObjTitulo.CodigoMoeda = ObjTitulo.Bordero.TitulosDoBordero(0).Titulo.CodigoMoeda

            ObjTitulo.CodigoProvisao = eProvisao.Provisao

            ObjTitulo.AtualizaValoresBordero()

            CarregarFormularioComAClasse(ObjTitulo)
            SessaoSalvaTitulo()
            divParametros.Visible = True
            divAcoes.Visible = False
            tbcFrmTitulo.ActiveTabIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GerarBorderoCobrancaSimples()
        Try
            SessaoRecuperaTitulosAcao()

            Dim B As New Novo.Bordero()
            B.IUD = "I"
            B.DataEnvio = CDate(txtDataEnvioBorderoCobrancaDuplicataSimples.Text)
            B.CodigoEmpresa = ObjTitulosAcao(0).CodigoEmpresa
            B.EnderecoEmpresa = ObjTitulosAcao(0).EnderecoEmpresa
            B.CodigoCarteiraDoTitulo = ddlCarteiraDuplicataCobrancaSimples.SelectedValue
            B.CodigoSituacao = eSituacao.Normal

            For Each TituloFilho In ObjTitulosAcao
                Dim BxT As New Novo.BorderoXTitulo
                BxT.CodigoTitulo = TituloFilho.Codigo
                BxT.Titulo = TituloFilho

                TituloFilho.IUD = "U"
                TituloFilho.CodigoContaContabilRecPag = ddlBancoDuplicataCobrancaSimples.SelectedValue
                TituloFilho.CodigoCarteiraDoTitulo = ddlCarteiraDuplicataCobrancaSimples.SelectedValue

                B.TitulosDoBordero.Add(BxT)
            Next

            If B.Salvar Then
                MsgBox(Me.Page, "Bordero salvo com Sucesso.", eTitulo.Sucess)
                divParametros.Visible = True
                divAcoes.Visible = False

                If chkImprimirDuplicatasCobrancaSimples.Checked Then
                    Dim lst As New Novo.ListTituloNovo
                    B.TitulosDoBordero.Select(Function(s) s.Titulo).ToList().ForEach(Function(x)
                                                                                         lst.Add(x)
                                                                                         Return True
                                                                                     End Function)
                    Novo.ImpressaoDuplicata.ExibirImpressao(Me.Page, lst, 1)
                End If
            Else
                MsgBox(Me.Page, "Erro ao salvar o bordero.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

    Protected Sub AtualizaTituloXContaContabil(ByVal Titulo As Novo.TituloNovo, Optional ByVal IncluirEncargo As Boolean = True)
        Dim Percentual As Decimal = 0
        Dim ValorTitulo As Decimal = 0
        Dim Soma As Decimal = 0
        Dim Conta = lblContaBordero.Text.Split("\")
        Dim Encargo As New Novo.TituloXContaContabil()

        Try
            If (Titulo.Bordero.TitulosDoBordero.Count > 0) Then
                'Bordero
                Encargo = Titulo.Valores.Find(Function(T) T.CodigoContaEncargo = Conta(0))
                For Each TitFilho As Novo.BorderoXTitulo In Titulo.Bordero.TitulosDoBordero
                    Dim Resultado As Decimal = 0
                    'Juros
                    For Each _Row As GridViewRow In gdvDesmembramentoValores.Rows()
                        If TitFilho.Titulo.Codigo = _Row.Cells(0).Text Then
                            If chkJuros.Checked Then
                                Dim Juros As Decimal = CDec(txtJuros.Text)
                                Dim Valor As Decimal = _Row.Cells(2).Text
                                Dim Dias As Integer = Convert.ToDateTime(txtProrrogacao.Text).Subtract(Convert.ToDateTime(_Row.Cells(3).Text)).Days
                                Resultado = (Valor * (CDec(txtJuros.Text) / 100) * (Dias / 30)).ToString("N2")
                                Titulo.Bordero.JurosTaxa = CDec(txtJuros.Text)
                            End If
                            If chkTaxa.Checked Then
                                If Not chkProporcional.Checked Then
                                    Dim Taxa As Decimal = CDec(txtTaxa.Text)
                                    Dim Valor As Decimal = CDec(txtTaxa.Text) / gdvDesmembramentoValores.Rows.Count
                                    If (_Row.RowIndex + 1) = gdvDesmembramentoValores.Rows.Count Then
                                        Resultado = (Taxa - Soma).ToString("N2")
                                    Else
                                        Resultado = Valor.ToString("N2")
                                        Soma += Resultado
                                    End If
                                Else
                                    Dim ValorBordero = CType(gridValores.Rows(0).FindControl("txtValorOficial"), TextBox).Text
                                    ValorTitulo = _Row.Cells(2).Text
                                    Percentual = (100 * ValorTitulo) / ValorBordero
                                    If (_Row.RowIndex + 1) = gdvDesmembramentoValores.Rows.Count Then
                                        Resultado = CDec(txtTaxa.Text) - Soma
                                    Else
                                        Resultado = ((Percentual * CDec(txtTaxa.Text)) / 100).ToString("N2")
                                        Soma += Resultado
                                    End If

                                End If
                            End If
                            CType(_Row.FindControl("txtValor"), TextBox).Text = Resultado
                            Exit For
                        End If
                    Next
                    If (IncluirEncargo) Then
                        If Encargo Is Nothing Then
                            Encargo = New Novo.TituloXContaContabil(Titulo)
                            Encargo.IUD = "I"
                            Encargo.CodigoContaEncargo = Conta(0)
                            Encargo.DC = Conta(2).ToString.Trim
                            Encargo.Descricao = Conta(1)
                            Encargo.ValorMoeda = IIf(Titulo.Moeda.Classificacao = eTiposMoeda.Oficial, Resultado / Titulo.IndiceTitulo, Resultado * Titulo.IndiceTitulo)
                            Encargo.ValorOficial = Resultado
                        Else
                            Encargo.IUD = "I"
                            Encargo.CodigoContaEncargo = Conta(0)
                            Encargo.DC = Conta(2).ToString.Trim
                            Encargo.Descricao = Conta(1)
                            Encargo.ValorMoeda = IIf(Titulo.Moeda.Classificacao = eTiposMoeda.Oficial, (Encargo.ValorOficial + Resultado) / Titulo.IndiceTitulo, Resultado * Titulo.IndiceTitulo)
                            Encargo.ValorOficial = (Encargo.ValorOficial + Resultado)
                        End If
                    End If
                Next
                If (IncluirEncargo) Then
                    Titulo.Valores.RemoveAll(Function(T) T.CodigoContaEncargo = Conta(0))
                    Titulo.Valores.Add(Encargo)
                    Titulo.AtualizaValoresBordero()
                End If
            End If
            Titulo.Valores.EncargoValorDocumento.IUD = "U"
            Titulo.Valores.EncargoValorLiquido.IUD = "U"
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub SelecionarTitulo()
        For Each row As GridViewRow In GridConsultaTitulos.Rows
            Dim tit As Novo.TituloNovo
            Dim index As Integer = row.RowIndex
            Dim ChkGridTitulos As CheckBox = CType(row.FindControl("ChkGridTitulos"), CheckBox)

            Try
                If ChkGridTitulos.Checked Then
                    tit = New Novo.TituloNovo(GridConsultaTitulos.Rows(row.RowIndex).Cells(0).Text)

                    'TODO: passar para RowDataBound da GridView e trazer desabilitado!!!
                    'If rdAgrupar.Checked Then
                    '    If tit.RegistroMestre > 0 Then
                    '        ChkGridTitulos.Checked = False
                    '        Exit Sub
                    '    End If
                    'End If

                    'If rdAdTituloAgrupamento.Checked Then
                    '    If tit.Codigo = tit.RegistroMestre Then
                    '        ChkGridTitulos.Checked = False
                    '        Exit Sub
                    '    End If
                    'End If

                    If rdAdTituloAgrupamento.Checked Then
                        SessaoRecuperaTituloAgrupado()
                        Dim codigo As String = row.Cells(0).Text
                        Dim TitExiste As Novo.TituloNovo = ObjTituloAgrupado.TitulosAgrupados.Find(Function(s) s.Codigo = codigo)
                        If Not TitExiste Is Nothing Then
                            TitExiste.IUD = "U"
                        Else
                            tit.IUD = "I"
                            ObjTituloAgrupado.TitulosAgrupados.Add(tit)
                        End If
                        SessaoSalvaTituloAgrupado()
                    ElseIf rdDuplicatasDescontadas.Checked Then
                        SessaoRecuperaTitulo()
                        Dim BxT As New Novo.BorderoXTitulo
                        BxT.CodigoTitulo = tit.Codigo
                        BxT.Titulo = tit
                        ObjTitulo.Bordero.TitulosDoBordero.Add(BxT)
                        SessaoSalvaTitulo()
                    Else
                        SessaoRecuperaTitulosAcao()
                        ObjTitulosAcao.Add(tit)
                        SessaoSalvaTitulosAcao()
                        If rdAgrupar.Checked And ObjTitulosAcao.Count = 1 Then TitulosConsulta()
                    End If
                Else
                    If rdAdTituloAgrupamento.Checked Then
                        SessaoRecuperaTituloAgrupado()
                        tit = ObjTituloAgrupado.TitulosAgrupados.Find(Function(s) s.Codigo = GridConsultaTitulos.Rows(index).Cells(0).Text)

                        'If tit Is Nothing Then
                        '    ChkGridTitulos.Checked = True
                        '    Exit Sub
                        'End If

                        If tit.IUD = "I" Then
                            ObjTituloAgrupado.TitulosAgrupados.Remove(tit)
                        Else
                            tit.IUD = "D"
                        End If
                        SessaoSalvaTituloAgrupado()
                    ElseIf rdDuplicatasDescontadas.Checked Then
                        SessaoRecuperaTitulo()
                        Dim BxT As New Novo.BorderoXTitulo
                        BxT = ObjTitulo.Bordero.TitulosDoBordero.Find(Function(s) s.CodigoTitulo = GridConsultaTitulos.Rows(index).Cells(0).Text)
                        ObjTitulo.Bordero.TitulosDoBordero.Remove(BxT)
                        SessaoSalvaTitulo()
                    Else
                        SessaoRecuperaTitulosAcao()
                        tit = ObjTitulosAcao.Find(Function(s) s.Codigo = GridConsultaTitulos.Rows(index).Cells(0).Text)
                        ObjTitulosAcao.Remove(tit)
                        SessaoSalvaTitulosAcao()
                    End If
                End If
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        Next
    End Sub

    Public Sub TextBoxAutoSuggest(ByVal strChave As String, txtSuggest As TextBox, ByVal lstLista As List(Of String))

    End Sub

#Region "Popup"

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Session("objClienteCR" & HID.Value) IsNot Nothing Then
                SessaoRecuperaTitulo()
                Dim cli As Cliente = Session("objClienteCR" & HID.Value)
                ObjTitulo.CliFor = cli
                SessaoSalvaTitulo()
                btnClientesTitulo.Enabled = True
                txtCliente.Enabled = True
                txtCliente.Text = ObjTitulo.CliFor.Codigo & " - " & ObjTitulo.CliFor.CodigoEndereco & " ..." & ObjTitulo.CliFor.Nome & " / " & ObjTitulo.CliFor.Cidade & "-" & ObjTitulo.CliFor.CodigoEstado
                Session.Remove("objClienteCR" & HID.Value)
                AtualizaValoresAdiantamentoNoForm(ObjTitulo)
            ElseIf Session("objClienteConsulta" & HID.Value) IsNot Nothing Then
                Dim cli As Cliente = Session("objClienteConsulta" & HID.Value)
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
                txtCodigoClienteConsulta.Value = itemCliente.Value
                txtClienteConsulta.Text = itemCliente.Text
                Session.Remove("objClienteConsulta" & HID.Value)
            ElseIf Session("objClienteLC" & HID.Value) IsNot Nothing Then
                Dim cli As Cliente = Session("objClienteLC" & HID.Value)
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
                SessaoRecuperaTitulo()
                ObjTitulo.CodigoClienteRecPag = itemCliente.Value.Split("-").ToArray(0)
                ObjTitulo.EndClienteRecPag = itemCliente.Value.Split("-").ToArray(1)
                SessaoSalvaTitulo()
                txtCodigoClienteConsulta.Value = itemCliente.Value
                txtClienteLancContabil.Text = itemCliente.Text
                Session.Remove("objClienteLC" & HID.Value)
            ElseIf Session("objDepositoCliFor" & HID.Value) IsNot Nothing Then
                Dim cli As Cliente = Session("objDepositoCliFor" & HID.Value)
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
                SessaoRecuperaTitulo()
                ObjTitulo.CodigoDepositoCliFor = itemCliente.Value.Split("-").ToArray(0)
                ObjTitulo.EndDepositoCliFor = itemCliente.Value.Split("-").ToArray(1)
                SessaoSalvaTitulo()
                txtDepositoCliente.Text = itemCliente.Text
                Session.Remove("objDepositoCliFor" & HID.Value)
            ElseIf Session("objDepositoRecPag" & HID.Value) IsNot Nothing Then
                Dim cli As Cliente = Session("objDepositoRecPag" & HID.Value)
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
                SessaoRecuperaTitulo()
                ObjTitulo.CodigoDepositoRecPag = itemCliente.Value.Split("-").ToArray(0)
                ObjTitulo.EndDepositoRecPag = itemCliente.Value.Split("-").ToArray(1)
                SessaoSalvaTitulo()
                txtDepositoRecPag.Text = itemCliente.Text
                Session.Remove("objDepositoRecPag" & HID.Value)
                ''updFrmTitulo.Update()
            ElseIf Session("objProdutoPRD" & HID.Value) IsNot Nothing Then
                SessaoRecuperaTitulo()
                ObjTitulo.CodigoProduto = CType(Session("objProdutoPRD" & HID.Value), Produto).Codigo
                txtProduto.Text = ObjTitulo.Produto.Codigo.ToString & " - " & ObjTitulo.Produto.Descricao
                txtQuantidade.Enabled = True
                SessaoSalvaTitulo()
                Session.Remove("objProdutoPRD" & HID.Value)
                divPedido.Visible = False
                divPedidoRecPag.Visible = False
            ElseIf Session("EncargosPlanoDeContas" & HID.Value) IsNot Nothing Then
                Dim LContas As ListEncargosPlanoDeContas = CType(Session("EncargosPlanoDeContas" & HID.Value), ListEncargosPlanoDeContas)
                If LContas.Count = 0 Then Exit Sub
                Dim Msg As String = ""
                SessaoRecuperaTitulo()
                For Each row In LContas.Where(Function(s) s.Selecionado AndAlso Not ObjTitulo.Valores.Select(Function(x) x.CodigoContaEncargo).Contains(s.CodigoContaEncargo))
                    Dim TxC As New Novo.TituloXContaContabil(ObjTitulo)
                    TxC.IUD = "I"
                    TxC.CodigoContaEncargo = row.CodigoContaEncargo
                    TxC.Descricao = row.TituloEncargo
                    TxC.DC = IIf(ObjTitulo.ReceberPagar = "C", "I", (IIf(ObjTitulo.ReceberPagar = "R", row.ContaEncargo.Receber, row.ContaEncargo.Pagar)))
                    If String.IsNullOrWhiteSpace(TxC.DC) Then
                        Msg &= IIf(Not String.IsNullOrWhiteSpace(Msg), ", ", "") & row.CodigoContaEncargo & "-" & row.TituloEncargo
                    Else
                        'If ObjTitulo.Bordero.TitulosDoBordero.Count > 0 Then
                        '    lblContaBordero.Text = row.CodigoContaEncargo & " \ " & row.TituloEncargo & " \ " & TxC.DC
                        '    divtblBordero.Visible = True
                        'Else
                        ObjTitulo.Valores.Insert(1, TxC)
                        gridValores.DataSource = ObjTitulo.Valores
                        gridValores.DataBind()
                        'End If
                    End If
                Next
                ObjTitulo.Valores.AtualizaValores()
                SessaoSalvaTitulo()
                If Not String.IsNullOrWhiteSpace(Msg) Then MsgBox(Me.Page, "Verifique no cadastro de Encargos x Plano de contas o comportamento da conta selecionada no contas a Pagar/Receber")
            ElseIf Not Session("objPedidoTitulo" & HID.Value) Is Nothing Then
                Dim Pedido As [Lib].Negocio.Pedido = CType(Session("objPedidoTitulo" & HID.Value), [Lib].Negocio.Pedido)
                SessaoRecuperaTitulo()
                If ViewState("campo").ToString.Equals(btnPedidoRecPag.ID) Then
                    txtPedidoRecPag.Text = Pedido.Codigo
                    SessaoRecuperaTitulo()
                    ObjTitulo.CodigoPedidoRecPag = Pedido.Codigo
                    SessaoSalvaTitulo()
                Else
                    txtPedido.Text = Pedido.Codigo
                    ObjTitulo.CodigoPedido = Pedido.Codigo
                    divProduto.Visible = False
                    divQuantidade.Visible = False
                    ddlLote.Parent.Visible = False
                    ddlSequencia.Parent.Visible = False
                    txtDepositoCliente.Parent.Visible = False
                    txtDepositoRecPag.Parent.Visible = False
                    ddlMoeda.SelectedValue = Pedido.CodigoMoeda.ToString
                    'ObjTitulo.AdiantamentosAbertos.Clear()
                    ddlMoeda.Enabled = False
                    AtualizaValoresAdiantamentoNoForm(ObjTitulo)
                End If
                SessaoSalvaTitulo()
            ElseIf Session("objBancoFRMTIT" & HID.Value) IsNot Nothing Then
                SessaoRecuperaTitulo()
                Dim Banco As [Lib].Negocio.ClienteXContaBancaria = CType(Session("objBancoFRMTIT" & HID.Value), [Lib].Negocio.ClienteXContaBancaria)
                ObjTitulo.CodigoBancoCliFor = Banco.CodigoBanco
                ObjTitulo.CodigoAgenciaCliFor = Banco.CodigoAgencia
                ObjTitulo.DigitoAgenciaCliFor = Banco.DigitoAgencia
                ObjTitulo.ContaCliFor = Banco.ContaCorrente
                ObjTitulo.DigitoContaCliFor = Banco.DigitoConta
                AtualizaDadosBancariosNoForm(ObjTitulo)
                Session.Remove("objBancoFRMTIT" & HID.Value)
                SessaoSalvaTitulo()
            ElseIf Not Session("objTituloPDC" & HID.Value) Is Nothing Then
                SessaoRecuperaTitulo()
                Dim PlanoDeConta As [Lib].Negocio.PlanoDeConta = CType(HttpContext.Current.Session("objTituloPDC" & HID.Value), [Lib].Negocio.PlanoDeConta)
                If ViewState("campo").ToString.Equals(btnPlanoDeContas.ID) Then
                    txtContaContabilCliFor.Text = PlanoDeConta.Conta & " - " & PlanoDeConta.Titulo
                    hdnCodigoContaContabilCliFor.Value = PlanoDeConta.Conta

                    SelecionarContaContabilCliFor()

                    If PlanoDeConta.TemCliente And Not ObjTitulo.CodigoPedido > 0 Then
                        btnClientesTitulo_Click(btnClientesTitulo, Nothing)
                    Else
                        txtCliente.Enabled = False
                    End If

                    If PlanoDeConta.Conta.Length = 9 AndAlso RdContabil.Checked Then 'Conta com 9 digitos não há necessidade de cliente, banco e pedido
                        lblCliFor.Parent.Visible = False
                        divBanco.Visible = False
                        divPedido.Visible = False
                    End If

                    If PlanoDeConta.Adiantamento AndAlso RdPagar.Checked Then
                        divPedido.Visible = False
                    Else
                        divPedido.Visible = True
                    End If

                Else
                    txtContaContabilRecPag.Text = PlanoDeConta.Conta & " - " & PlanoDeConta.Titulo
                    ObjTitulo.CodigoContaContabilRecPag = PlanoDeConta.Conta
                    If PlanoDeConta.TemCliente And Not ObjTitulo.CodigoPedido > 0 Then
                        btnClienteLancamentoContabil_Click(btnClienteLancamentoContabil, Nothing)
                    End If
                    SelecionarContaContabilRecPag()
                    'ddlContaContabilEmpresaRecPag.SelectedValue = PlanoDeConta.Conta
                End If
                Session.Remove("objTituloPDC" & HID.Value)
                SessaoSalvaTitulo()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Splip"

    Protected Sub EmitirSplip()
        Try
            'Dataset para report
            Dim dtRecibo = New DataTable("ReciboAPagar")
            dtRecibo.Columns.Add("ENome", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ECidade", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EEstado", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ECnpj", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CNome", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CCnpj", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TNumtit", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TValor", Type.GetType("System.Decimal"))
            dtRecibo.Columns.Add("THistorico", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TVencimento", Type.GetType("System.DateTime"))
            dtRecibo.Columns.Add("TContaContabilPagadora", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TMovimento", Type.GetType("System.String"))

            Dim Cliente = DdlEmpresaConsultaTitulos.SelectedValue.Split("-")

            Dim where As New Hashtable
            where.Add("UndadeDeNegocio", DdlUnidadeConsultaTitulos.SelectedValue)
            where.Add("Empresa", Cliente(0))
            where.Add("EndEmpresa", Cliente(1))
            where.Add("Slip", "Slip")

            Dim ds As DataSet = Banco.ConsultaDataSet(getSqlConsulta(where), "Slip")
            GridConsultaTitulos.DataSource = ds
            GridConsultaTitulos.DataBind()

            Dim Row As DataRow
            For Each _Row As DataRow In ds.Tables(0).Rows
                Row = dtRecibo.NewRow()
                'Carrega o DataSet
                Row("ENome") = _Row("Empresa")
                Row("ECidade") = _Row("Cidade")
                Row("EEstado") = _Row("Estado")
                Row("ECnpj") = _Row("EndEmpresa")
                Row("CNome") = _Row("Nome")
                Row("CCnpj") = _Row("Cliente_Id")
                Row("TNumtit") = _Row("Titulo_Id")
                Row("TValor") = _Row("ValorLiquido")
                Row("THistorico") = _Row("Historico")
                Row("TVencimento") = _Row("Vencimento")
                Row("TContaContabilPagadora") = _Row("ContaContabilCliFor")
                Row("TMovimento") = _Row("Vencimento")

                dtRecibo.Rows.Add(Row)
            Next
            Dim dsRecibo As New DataSet
            dsRecibo.Tables.Add(dtRecibo)
            Dim dtImagem As New DataTable("Images")
            dtImagem.Columns.Add("path", GetType(String))
            dtImagem.Columns.Add("image", GetType(System.Byte()))

            Dim drImagem As DataRow = dtImagem.NewRow()
            Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))

            drImagem("path") = strCaminhoImagem
            drImagem("image") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
            dtImagem.Rows.Add(drImagem)

            dsRecibo.Tables.Add(dtImagem)

            Dim crpt As New ReportDocument()

            Try
                crpt.FileName = Server.MapPath("~/Reports/Cr_slip.rpt")
                crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

                Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
                Dim arquivo As String = NomeArquivo

                crpt.SetDataSource(dsRecibo)

                Dim crparametervalues As ParameterValues
                Dim crparameterdiscretevalue As ParameterDiscreteValue
                Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

                crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

                crparameterfielddefinition = crparameterfielddefinitions.Item("XNome")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = String.Empty
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
                ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo2 & "');", True)
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Finally
                crpt.Close()
                crpt.Dispose()
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Agrupamento"

    Protected Sub AgruparValoresTituloXContaContabil(ByVal Titulo As Novo.TituloNovo)
        Try
            Dim i As Integer = 1
            Dim vlrEncargo As Decimal = 0
            For Each TitAgrupado As Novo.TituloNovo In Titulo.TitulosAgrupados
                For Each Valores As Novo.TituloXContaContabil In Titulo.Valores
                    If (Valores.CodigoContaEncargo.Substring(0, 7) <> Titulo.Valores.EncargoValorDocumento.CodigoContaEncargo.Substring(0, 7) And Valores.CodigoContaEncargo.Substring(0, 7) <> Titulo.Valores.EncargoValorLiquido.CodigoContaEncargo.Substring(0, 7)) Then

                        'Realiza o cálculo percentual
                        Dim Valor = Titulo.Valores.EncargoValorDocumento.ValorOficial
                        Dim Percentual = Math.Round((100 * TitAgrupado.Valores.EncargoValorDocumento.ValorOficial) / Valor)
                        Dim Resultado = ((Percentual * CDec(Valores.ValorOficial) / 100)).ToString("N2")
                        If i = Titulo.TitulosAgrupados.Count Then
                            Resultado = Valores.ValorOficial - vlrEncargo
                        Else
                            vlrEncargo += Resultado
                        End If

                        Dim CodigoConta = Valores.CodigoContaEncargo

                        Dim Encargo As New Novo.TituloXContaContabil()
                        Encargo = TitAgrupado.Valores.Find(Function(T) T.CodigoContaEncargo = CodigoConta)

                        If Encargo Is Nothing Then
                            Encargo = New Novo.TituloXContaContabil(TitAgrupado)
                            Encargo.IUD = "I"
                            Encargo.CodigoContaEncargo = Valores.CodigoContaEncargo
                            Encargo.DC = Valores.DC
                            Encargo.Descricao = Valores.Descricao
                            Encargo.ValorMoeda = IIf(TitAgrupado.Moeda.Classificacao = eTiposMoeda.Oficial, Resultado / TitAgrupado.IndiceTitulo, Resultado * TitAgrupado.IndiceTitulo)
                            Encargo.ValorOficial = Resultado
                            TitAgrupado.Valores.Add(Encargo)
                        Else
                            Encargo.IUD = "I"
                            Encargo.CodigoContaEncargo = Valores.CodigoContaEncargo
                            Encargo.DC = Valores.DC
                            Encargo.Descricao = Valores.Descricao
                            Encargo.ValorMoeda = IIf(TitAgrupado.Moeda.Classificacao = eTiposMoeda.Oficial, (Encargo.ValorOficial + Resultado) / TitAgrupado.IndiceTitulo, Resultado * TitAgrupado.IndiceTitulo)
                            Encargo.ValorOficial = (Encargo.ValorOficial + Resultado)
                        End If
                    End If
                Next
                i += 1
            Next
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RemoverContaXContabilAgrupamento(ByVal Titulo As Novo.TituloNovo)
        Try
            For Each Valores As Novo.TituloXContaContabil In Titulo.Valores
                If (Valores.CodigoContaEncargo <> Titulo.Valores.EncargoValorDocumento.CodigoContaEncargo And Valores.CodigoContaEncargo <> Titulo.Valores.EncargoValorLiquido.CodigoContaEncargo) Then
                    For Each TitAgrupado As Novo.TituloNovo In Titulo.TitulosAgrupados
                        'Realiza o cálculo percentual
                        Dim Valor = Titulo.Valores.EncargoValorDocumento.ValorOficial
                        Dim Percentual = Math.Round((100 * TitAgrupado.Valores.EncargoValorDocumento.ValorOficial) / Valor)
                        Dim Resultado = Math.Round((Percentual * CDec(Valores.ValorOficial)) / 100).ToString("N2")
                        Dim CodigoConta = Valores.CodigoContaEncargo

                        Dim Encargo As New Novo.TituloXContaContabil()
                        Encargo = TitAgrupado.Valores.Find(Function(T) T.CodigoContaEncargo = CodigoConta)

                        If Not Encargo Is Nothing Then
                            If ((Encargo.ValorOficial - Resultado) = 0) Then
                                TitAgrupado.Valores.Remove(Encargo)
                            Else
                                Encargo.ValorOficial = Encargo.ValorOficial - Resultado
                            End If
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Agrupar()
        Try
            SessaoRecuperaTitulosAcao()
            Limpar(1)

            ObjTitulo.DesligarControles()
            ObjTitulo.IUD = "I"
            ObjTitulo.RegistroMestre = ObjTitulo.Codigo
            ObjTitulo.Observacoes = "Agrupamento dos Registros"
            ObjTitulo.Movimento = Today
            ObjTitulo.Vencimento = Funcoes.ValidaDataUtil(ObjTitulo.CodigoEmpresa, ObjTitulo.EnderecoCliFor, Date.Now).ToString("dd/MM/yyyy")
            ObjTitulo.Reprogramacao = Funcoes.ValidaDataUtil(ObjTitulo.CodigoEmpresa, ObjTitulo.EnderecoCliFor, Date.Now).ToString("dd/MM/yyyy")

            For Each row In ObjTitulosAcao
                Dim TituloFilho As Novo.TituloNovo = row
                TituloFilho.RegistroMestre = ObjTitulo.Codigo
                TituloFilho.TituloMestre = ObjTitulo
                TituloFilho.Reprogramacao = Format(Today, "dd/MM/yyyy")
                TituloFilho.IUD = "U"
                ObjTitulo.TitulosAgrupados.Add(TituloFilho)
                ObjTitulo.Observacoes &= " - " & TituloFilho.Codigo
            Next

            If ObjTitulo.TitulosAgrupados.Count < 2 Then
                MsgBox(Me.Page, "Selecione ao Menos 2 Titulos para efetuar o Agrupamento")
                Exit Sub
            End If

            '****** Alimenta principais campos do Mestre com as informacoes do primeiro titulo agrupado ******'
            ObjTitulo.ReceberPagar = ObjTitulo.TitulosAgrupados(0).ReceberPagar
            ObjTitulo.CodigoUnidadeDeNegocio = ObjTitulo.TitulosAgrupados(0).CodigoUnidadeDeNegocio
            ObjTitulo.CodigoEmpresa = ObjTitulo.TitulosAgrupados(0).CodigoEmpresa
            ObjTitulo.EnderecoEmpresa = ObjTitulo.TitulosAgrupados(0).EnderecoEmpresa
            ObjTitulo.CodigoCliFor = ObjTitulo.TitulosAgrupados(0).CodigoCliFor
            ObjTitulo.EnderecoCliFor = ObjTitulo.TitulosAgrupados(0).EnderecoCliFor
            ObjTitulo.CodigoContaContabilCliFor = ObjTitulo.TitulosAgrupados(0).CodigoContaContabilCliFor
            'ObjTitulo.CodigoContaContabilRecPag = ObjTitulo.TitulosAgrupados(0).CodigoContaContabilRecPag
            ObjTitulo.CodigoBancoCliFor = ObjTitulo.TitulosAgrupados(0).CodigoBancoCliFor
            ObjTitulo.CodigoAgenciaCliFor = ObjTitulo.TitulosAgrupados(0).CodigoAgenciaCliFor
            ObjTitulo.DigitoAgenciaCliFor = ObjTitulo.TitulosAgrupados(0).DigitoAgenciaCliFor
            ObjTitulo.ContaCliFor = ObjTitulo.TitulosAgrupados(0).ContaCliFor
            ObjTitulo.DigitoContaCliFor = ObjTitulo.TitulosAgrupados(0).DigitoContaCliFor

            ObjTitulo.CodigoMoeda = ObjTitulo.TitulosAgrupados(0).CodigoMoeda
            ObjTitulo.CodigoProvisao = eProvisao.Baixa
            ObjTitulo.AtualizaValoresAgrupados()

            CarregarFormularioComAClasse(ObjTitulo)

            SessaoSalvaTitulo()
            divParametros.Visible = True
            divAcoes.Visible = True
            tbcFrmTitulo.ActiveTabIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Adicionar()
        Try
            SessaoRecuperaTituloAgrupado()

            If (From x In ObjTituloAgrupado.TitulosAgrupados Where x.IUD = "I" Or x.IUD = "U" Select x).Count < 2 Then
                MsgBox(Me.Page, "O agrupamento deve conter ao menos dois títulos.")
                Exit Sub
            End If

            Dim obs As String = ""
            Dim sqls As New ArrayList
            RemoverContaXContabilAgrupamento(ObjTituloAgrupado)
            For Each row In ObjTituloAgrupado.TitulosAgrupados
                If row.IUD = "D" Then
                    row.IUD = "U"
                    row.RegistroMestre = 0
                    row.SalvarSql(sqls)
                    row.IUD = "D"
                Else
                    row.IUD = "U"
                    obs &= " - " & row.Codigo
                End If
            Next

            'ObjTituloAgrupado.TitulosAgrupados.RemoveAll(Function(s) s.IUD = "D")
            ObjTituloAgrupado.AtualizaValoresAgrupados()
            AgruparValoresTituloXContaContabil(ObjTituloAgrupado)
            ObjTituloAgrupado.IUD = "U"

            ObjTituloAgrupado.Observacoes = "Agrupamento dos Registros" & obs

            SessaoRecuperaTitulo()
            ObjTitulo = ObjTituloAgrupado
            SessaoSalvaTitulo()

            CarregarFormularioComAClasse(ObjTituloAgrupado)
            tbcFrmTitulo.ActiveTabIndex = 0
            SessaoSalvaTituloAgrupado()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Reprogramação"

    Protected Sub Reprogramar()
        SessaoRecuperaTitulosAcao()
        'Valida a Reprogramação
        If ObjTitulosAcao.Count <= 0 Then
            MsgBox(Me.Page, "Não há títulos selecionados para realizar a reprogramação.")
            Exit Sub
        ElseIf (Not IsDate(txtDataReprogramacao.Text)) Then
            MsgBox(Me.Page, "Campo data reprogramação inválido.")
            Exit Sub
        ElseIf (CDate(txtDataReprogramacao.Text) < Date.Today) Then
            MsgBox(Me.Page, "Data de reprogramação não pode ser anterior a data de hoje.")
            Exit Sub
        End If
        If ObjTitulosAcao.Reprogramar(CDate(txtDataReprogramacao.Text)) Then
            MsgBox(Me.Page, "Títulos reprogramados com Sucesso.", eTitulo.Sucess)
            LimparConsulta()
        Else
            MsgBox(Me.Page, "Erro ao reprogramar títulos.")
        End If
    End Sub

#End Region

#Region "Baixar"

    Protected Sub Baixar()
        'Validar Baixa
        Dim msgErro As String = String.Empty

        If Not IsDate(txtDataBaixarTodos.Text) Then
            MsgBox(Me.Page, "Informe uma data de baixa válida.")
        End If
        If (String.IsNullOrWhiteSpace(ddlEmpresaBaixaTodos.SelectedValue)) Then
            MsgBox(Me.Page, "O campo empresa é obrigatório.")
        End If
        If String.IsNullOrWhiteSpace(DdlTiposDeRecebimentosBaixarTodos.SelectedValue) Then
            MsgBox(Me.Page, "O campo tipo é obrigatório.")
        End If
        If String.IsNullOrWhiteSpace(ddlContaContabilEmpresaRecPagBaixarTodos.SelectedValue) Then
            MsgBox(Me.Page, "O campo banco é obrigatório.")
        End If
        If (String.IsNullOrWhiteSpace(ddlUnidadeDeNegocioBaixa.SelectedValue)) Then
            MsgBox(Me.Page, "O campo unidade de negócio é obrigatório.")
        End If
        If Not (String.IsNullOrWhiteSpace(msgErro)) Then
            MsgBox(Me.Page, msgErro)
            Exit Sub
        End If

        Dim Emp As String() = ddlEmpresaBaixaTodos.SelectedValue.Split("-")
        Dim DataProgramavel As Date = Funcoes.ValidaDataUtil(Emp(0), Emp(1), CDate(txtDataBaixarTodos.Text))
        ''Verificar se a data esta aberta
        If DataProgramavel > CDate(txtDataBaixarTodos.Text) Then
            MsgBox(Me.Page, "Data de Baixa informada " & CDate(txtDataBaixarTodos.Text).ToString("dd/MM/yyyy") & " é uma data não programavel, confirme a utilizacao da proxima data util valida " & DataProgramavel.ToString("dd/MM/yyyy"))
            txtDataBaixarTodos.Text = DataProgramavel.ToString("dd/MM/yyyy")
            Exit Sub
        End If

        SessaoRecuperaTitulosAcao()
        For Each row In ObjTitulosAcao
            row.IUD = "U"
            row.CodigoProvisao = eProvisao.Baixa
            row.DataBaixa = DataProgramavel
            row.CodigoUnidadeDeNegocioRecPag = ddlUnidadeDeNegocioBaixa.SelectedValue
            row.CodigoEmpresaRecPag = Emp(0)
            row.EndEmpresaRecPag = Emp(1)
            row.CodigoTipoPgto = DdlTiposDeRecebimentosBaixarTodos.SelectedValue
            row.CodigoContaContabilRecPag = ddlContaContabilEmpresaRecPagBaixarTodos.SelectedValue
        Next

        If ObjTitulosAcao.Salvar Then
            MsgBox(Me.Page, "Títulos baixados com Sucesso.", eTitulo.Sucess)
            LimparConsulta()
        Else
            MsgBox(Me.Page, "Erro ao baixar títulos.")
        End If
    End Sub

#End Region

#Region "Contabilização"

    Public Sub CarregarContabilizacao(ByVal CodigoTitulo As Integer)
        Try
            If CodigoTitulo = 0 Then
                lblContabilizacao.Text = "Contabilizacao"
                gridRazao.DataSource = Nothing
                gridRazao.DataBind()
                Exit Sub
            End If

            lblContabilizacao.Text = "Contabilizacao do Titulo " & CodigoTitulo
            Sql = "SELECT case when Razao.Titulo = " & CodigoTitulo & " then 1 else 0 end Pintar," & vbCrLf & _
                  "       (SELECT Reduzido FROM Clientes AS C WHERE C.Cliente_Id = Razao.Empresa_Id and C.Endereco_Id = Razao.EndEmpresa_Id) AS Reduzido, " & vbCrLf & _
                  "       Razao.Conta_Id, " & vbCrLf & _
                  "       Razao.Cliente_Id," & vbCrLf & _
                  "       Razao.EndCliente_Id," & vbCrLf & _
                  "       PlanoDeContas.Titulo AS Titulo, " & vbCrLf & _
                  "       Razao.Movimento_Id," & vbCrLf & _
                  "       Razao.Lote_Id, " & vbCrLf & _
                  "       isnull(Razao.Produto, '') AS Produto," & vbCrLf & _
                  "       Razao.Custo," & vbCrLf & _
                  "       Razao.Historico," & vbCrLf & _
                  "       Razao.DebitoOficial," & vbCrLf & _
                  "       Razao.CreditoOficial" & vbCrLf & _
                  "  FROM Razao" & vbCrLf & _
                  " INNER JOIN PlanoDeContas" & vbCrLf & _
                  "    ON Razao.Conta_Id = PlanoDeContas.Conta_Id" & vbCrLf & _
                  " Inner Join Titulos T" & vbCrLf & _
                  "    on Razao.Titulo = T.Titulo_Id" & vbCrLf & _
                  " where (Razao.Titulo in (" & vbCrLf & _
                  "                         Select Titulo_id " & vbCrLf & _
                  "                           from titulos" & vbCrLf & _
                  "                          where registromestre in (Select RegistroMestre" & vbCrLf & _
                  "                                                     from titulos" & vbCrLf & _
                  "                                                    where titulo_Id = " & CodigoTitulo & ")" & vbCrLf & _
                  "                            and RegistroMestre <> 0" & vbCrLf & _
                  "                        ) or Razao.Titulo =  " & CodigoTitulo & ")" & vbCrLf & _
                  " order by case when Razao.Titulo = T.RegistroMestre then 1 else 0 end, Reduzido, Razao.Titulo, Razao.conta_Id "

            Dim dsRazao As New DataSet
            dsRazao = Banco.ConsultaDataSet(Sql, "Razao")
            gridRazao.DataSource = Nothing

            If Not dsRazao Is Nothing AndAlso dsRazao.Tables(0).Rows.Count > 0 Then
                Dim dtRazao As New DataTable("razao")
                Dim saldo As Double = 0

                dtRazao.Columns.Add("Conta", Type.GetType("System.String"))
                dtRazao.Columns.Add("Cliente", Type.GetType("System.String"))
                dtRazao.Columns.Add("Movimento", Type.GetType("System.DateTime"))
                dtRazao.Columns.Add("Titulo", Type.GetType("System.String"))
                dtRazao.Columns.Add("Lote", Type.GetType("System.String"))
                dtRazao.Columns.Add("Produto", Type.GetType("System.String"))
                dtRazao.Columns.Add("Custo", Type.GetType("System.String"))
                dtRazao.Columns.Add("Historico", Type.GetType("System.String"))
                dtRazao.Columns.Add("Debito", Type.GetType("System.Double"))
                dtRazao.Columns.Add("Credito", Type.GetType("System.Double"))
                dtRazao.Columns.Add("Saldo", Type.GetType("System.Double"))
                dtRazao.Columns.Add("Reduzido", Type.GetType("System.String"))

                For Each row As DataRow In dsRazao.Tables(0).Rows
                    Dim drRazao As DataRow = dtRazao.NewRow()

                    drRazao("Conta") = row("Conta_Id")
                    drRazao("Cliente") = row("Cliente_Id") & "-" & row("EndCliente_Id")
                    drRazao("Movimento") = row("Movimento_Id")
                    drRazao("Titulo") = row("Titulo")
                    drRazao("Lote") = row("Lote_Id")
                    drRazao("Produto") = row("Produto")
                    drRazao("Custo") = row("Custo")
                    drRazao("Historico") = row("Historico")
                    drRazao("Debito") = row("DebitoOficial")
                    drRazao("Credito") = row("CreditoOficial")
                    saldo = Math.Round(saldo + (row("DebitoOficial") - row("CreditoOficial")), 2)
                    drRazao("Saldo") = saldo
                    drRazao("Reduzido") = row("Reduzido")
                    dtRazao.Rows.Add(drRazao)
                Next
                gridRazao.DataSource = dtRazao
            End If
            gridRazao.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Borderô"

    Protected Sub CalcularDesagioBordero(ByVal Titulo As Novo.TituloNovo)
        Try
            Dim Resultado As Decimal = 0
            For Each TitFilho As Novo.BorderoXTitulo In Titulo.Bordero.TitulosDoBordero
                Dim Juros As Decimal = CDec(txtJurosBordero.Text)
                Dim Valor As Decimal = TitFilho.Titulo.Valores.EncargoValorLiquido.ValorOficial
                Dim Dias As Integer = Convert.ToDateTime(TitFilho.Titulo.Reprogramacao).Subtract(Convert.ToDateTime(txtDataEnvioBordero.Text)).Days
                Resultado += (Valor * (Juros / 100) * (Dias / 30)).ToString("N2")
                Titulo.Bordero.JurosTaxa = Juros
            Next
            'Inclui a conta Contábil de Deságio no Borderô
            Dim Encargo As New Novo.TituloXContaContabil(Titulo)
            Encargo.IUD = "I"
            Encargo.CodigoContaEncargo = "402020115"
            Encargo.DC = "D"
            Encargo.Descricao = "ENCARGOS S/ TITULOS DESCONTADOS"
            Encargo.ValorMoeda = IIf(Titulo.Moeda.Classificacao = eTiposMoeda.Oficial, (Encargo.ValorOficial + Resultado) / Titulo.IndiceTitulo, Resultado * Titulo.IndiceTitulo)
            Encargo.ValorOficial = (Encargo.ValorOficial + Resultado)
            Titulo.Valores.RemoveAll(Function(T) T.CodigoContaEncargo = "402020115")
            Titulo.Valores.Add(Encargo)
            Titulo.Valores.EncargoValorDocumento.IUD = "U"
            Titulo.Valores.EncargoValorLiquido.IUD = "U"
            Titulo.AtualizaValoresBordero()
            AtualizaValoresNoForm(Titulo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

    Protected Sub SelecionarContaContabilCliFor()
        Try
            SessaoRecuperaTitulo()
            ObjTitulo.CodigoContaContabilCliFor = hdnCodigoContaContabilCliFor.Value
            CarregarContaContabilEmpresaRecPag(ObjTitulo.ReceberPagar)
            'Atualiza os valores
            AtualizaValoresNoForm(ObjTitulo)
            'Atualiza os Adiantamentos
            If ObjTitulo.ContaContabilCliFor.Adiantamento Then AtualizaValoresAdiantamentoNoForm(ObjTitulo)

            If ObjTitulo.ContaContabilCliFor.Adiantamento Then
                BtnAdicionarConta.Visible = False
            Else
                BtnAdicionarConta.Visible = True
            End If

            ObjTitulo.CodigoContaContabilRecPag = "1010102"

            If ObjTitulo.ReceberPagar = "C" Then
                If Not ObjTitulo.ContaContabilCliFor.TemCliente Then
                    ObjTitulo.CodigoCliFor = ""
                    ObjTitulo.EnderecoCliFor = 0
                    txtCliente.Text = "CONTA SEM CLIENTE"
                    btnClientesTitulo.Enabled = False
                    txtCliente.Enabled = False
                    divPedido.Visible = False
                    divPedidoRecPag.Visible = False
                Else
                    ucConsultaClientes.Limpar()
                    Popup.ConsultaDeClientes(Me.Page, "objClienteCR" & HID.Value, "txtNome")
                    divPedido.Visible = True
                    divPedidoRecPag.Visible = True
                End If
            Else
                btnClientesTitulo.Enabled = True
                txtCliente.Enabled = True
            End If
            'Conta Contabil de Adiantamento não tem Juros ou Descontos
            If Not ObjTitulo.ContaContabilCliFor.Adiantamento AndAlso Not RdContabil.Checked Then
                AdicionaContaJurosDesconto(ObjTitulo)
            Else
                ObjTitulo.Valores.RemoveAll(Function(s) Not s.CodigoContaEncargo = ObjTitulo.Valores.EncargoValorDocumento.CodigoContaEncargo AndAlso Not s.CodigoContaEncargo = ObjTitulo.Valores.EncargoValorLiquido.CodigoContaEncargo)
                gridValores.DataSource = ObjTitulo.Valores
                gridValores.DataBind()
            End If
            SessaoSalvaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub SelecionarContaContabilRecPag()
        Try
            AtualizaValoresNoForm(ObjTitulo)

            If ObjTitulo.ReceberPagar = "C" Then
                If Not ObjTitulo.ContaContabilRecPag.TemCliente Then
                    btnClienteLancamentoContabil.Enabled = False
                    txtClienteLancContabil.Enabled = False
                    txtClienteLancContabil.Text = "CONTA SEM CLIENTE"
                    ObjTitulo.CodigoClienteRecPag = ""
                    ObjTitulo.EndClienteRecPag = 0
                Else
                    btnClienteLancamentoContabil.Enabled = True
                    txtClienteLancContabil.Enabled = True
                    txtClienteLancContabil.Text = ""

                    ObjTitulo.CodigoClienteRecPag = ""
                    ObjTitulo.EndClienteRecPag = 0
                End If

                If Not ObjTitulo.ContaContabilRecPag.TemProduto Then
                    btnProduto.Enabled = False
                    btnProdutoExcluir.Enabled = False
                    txtProduto.Enabled = False
                    txtProduto.Text = "CONTA SEM PRODUTO"
                    ObjTitulo.CodigoProduto = ""
                    divProduto.Visible = False
                    txtQuantidade.Text = "0,00"
                    txtQuantidade.Enabled = False
                    ObjTitulo.Quantidade = 0
                Else
                    btnProduto.Enabled = True
                    btnProdutoExcluir.Enabled = True
                    txtProduto.Enabled = True
                    txtProduto.Text = ""
                    divProduto.Visible = True
                    ObjTitulo.CodigoProduto = ""
                    txtQuantidade.Text = "0,00"
                    txtQuantidade.Enabled = True
                    ObjTitulo.Quantidade = 0
                End If
            End If

            'Carregar Unidade e Empresa pela conta banco
            'Dim BxC As New [Lib].Negocio.BancosXContas(ddlContaContabilEmpresaRecPag.SelectedValue)

            'If (BxC IsNot Nothing) Then
            '    ddlUnidadeDeNegocioRecPag.SelectedValue = BxC.CodigoUnidadeDeNegocio
            '    ddl.Carregar(DdlEmpresaRecPag, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocioRecPag.SelectedValue, True)
            '    ObjTitulo.CodigoUnidadeDeNegocioRecPag = ddlUnidadeDeNegocioRecPag.SelectedValue
            '    DdlEmpresaRecPag.SelectedValue = BxC.CodigoEmpresa & "-" & BxC.EndEmpresa
            'End If

            Dim BxC As New [Lib].Negocio.BancosXContas(ObjTitulo.CodigoContaContabilRecPag)
            If BxC IsNot Nothing AndAlso BxC.Banco IsNot Nothing Then
                lblDiasBanco.Text = BxC.Banco.LiquidacaoDias.ToString & " Dia(s)"
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub AdicionaContaTEDDOC(ByVal pTitulo As Novo.TituloNovo)
        Try
            'Pagamento em TED/DOC adiciona a conta contábil de desconto TED/DOC
            If pTitulo.Valores.Where(Function(s) s.CodigoContaEncargo = pTitulo.Empresa.Empresa.CodigoContaGrupoTedDoc).Count = 0 Then
                If pTitulo.CodigoTipoPgto = 6 OrElse pTitulo.CodigoTipoPgto = 11 AndAlso pTitulo.ReceberPagar = "P" Then
                    Dim TituloXContaContabil As New Novo.TituloXContaContabil(pTitulo)
                    TituloXContaContabil.IUD = "I"
                    TituloXContaContabil.DC = "D"
                    TituloXContaContabil.CodigoContaEncargo = pTitulo.Empresa.Empresa.CodigoContaGrupoTedDoc
                    pTitulo.Valores.Add(TituloXContaContabil)
                End If
                pTitulo.Valores.AtualizaLiquido()
                AtualizaValoresNoForm(pTitulo)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao adicionar a conta TED/DOC no cadastro da empresa \n" & ex.Message, eTitulo.Info, False)
        End Try
    End Sub

    Protected Sub AdicionaContaJurosDesconto(ByVal pTitulo As Novo.TituloNovo)
        Try
            If RdPagar.Checked Then
                'Juros
                If Not pTitulo.Valores.Where(Function(s) s.CodigoContaEncargo = pTitulo.Empresa.Empresa.CodigoContaJuroPago).Count > 0 Then
                    Dim TitXCCJurosPagos As New Novo.TituloXContaContabil(pTitulo)
                    TitXCCJurosPagos.IUD = "I"
                    TitXCCJurosPagos.DC = "D"
                    TitXCCJurosPagos.CodigoContaEncargo = pTitulo.Empresa.Empresa.CodigoContaJuroPago
                    pTitulo.Valores.Add(TitXCCJurosPagos)
                End If

                'Desconto
                If Not pTitulo.Valores.Where(Function(s) s.CodigoContaEncargo = pTitulo.Empresa.Empresa.CodigoContaDescontoObtido).Count > 0 Then
                    pTitulo.Valores.RemoveAll(Function(s) s.CodigoContaEncargo = pTitulo.Empresa.Empresa.CodigoContaDescontoObtido)
                    Dim TitXCCDescontoObtido As New Novo.TituloXContaContabil(pTitulo)
                    TitXCCDescontoObtido.IUD = "I"
                    TitXCCDescontoObtido.DC = "C"
                    TitXCCDescontoObtido.CodigoContaEncargo = ObjTitulo.Empresa.Empresa.CodigoContaDescontoObtido
                    pTitulo.Valores.Add(TitXCCDescontoObtido)
                End If
            Else
                'Juros
                pTitulo.Valores.RemoveAll(Function(s) s.CodigoContaEncargo = pTitulo.Empresa.Empresa.CodigoContaJuroPago)
                Dim TitXCCJurosAuferido As New Novo.TituloXContaContabil(pTitulo)
                TitXCCJurosAuferido.IUD = "I"
                TitXCCJurosAuferido.DC = "C"
                TitXCCJurosAuferido.CodigoContaEncargo = pTitulo.Empresa.Empresa.CodigoContaJuroRecebido
                pTitulo.Valores.Add(TitXCCJurosAuferido)
                'Desconto
                pTitulo.Valores.RemoveAll(Function(s) s.CodigoContaEncargo = pTitulo.Empresa.Empresa.CodigoContaDescontoObtido)
                Dim TitXCCDescontoConcedido As New Novo.TituloXContaContabil(pTitulo)
                TitXCCDescontoConcedido.IUD = "I"
                TitXCCDescontoConcedido.DC = "D"
                TitXCCDescontoConcedido.CodigoContaEncargo = pTitulo.Empresa.Empresa.CodigoContaDescontoConcedido
                pTitulo.Valores.Add(TitXCCDescontoConcedido)
            End If

            pTitulo.Valores.AtualizaLiquido()
            AtualizaValoresNoForm(pTitulo)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlLote_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlLote.SelectedIndexChanged
        Try
            If Not String.IsNullOrWhiteSpace(ddlSequencia.Text) Then
                ddlSequencia.SelectedIndex = 0
            End If
            TotalizaLote()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub TotalizaLote()
        SessaoRecuperaTitulo()
        Sql = "SELECT DISTINCT Sequencia_Id " & vbCrLf & _
               " FROM Razao" & vbCrLf & _
               " WHERE Empresa_Id = '" & ObjTitulo.CodigoEmpresa & "' AND EndEmpresa_Id = " & ObjTitulo.EnderecoEmpresa & vbCrLf & _
               " And Movimento_Id = '" & txtMovimento.Text.ToSqlDate() & "'" & vbCrLf & _
               " and Lote_Id = " & ddlLote.SelectedValue & vbCrLf & _
               " Order by Sequencia_Id" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            ddlSequencia.Items.Add(New ListItem(Dr("Sequencia_Id"), Dr("Sequencia_Id")))
        Next

        ddlSequencia.Items.Insert(0, "")
        ddlSequencia.SelectedIndex = 0
    End Sub

    Private Sub CarregarLote()
        ddlLote.Items.Clear()
        Sql = "SELECT     Lotes.Lote_Id as Codigo, convert(varchar,Lotes.Lote_Id) + '-' + Lotes.Descricao  as Descricao " & vbCrLf & _
              " FROM     Lotes RIGHT OUTER JOIN " & vbCrLf & _
              " Sistemas ON Lotes.Sistema_Id = Sistemas.Sistema_Id " & vbCrLf & _
              " Where Sistemas.Sistema_Id = 2" & vbCrLf

        ddlLote.DataValueField = "Codigo"
        ddlLote.DataTextField = "Descricao"
        ddlLote.DataSource = Banco.ConsultaDataSet(Sql, "Lotes")
        ddlLote.DataBind()
        ddlLote.Items.Insert(0, "")
        ddlLote.ClearSelection()
        ddlLote.SelectedIndex = 1
    End Sub

#Region "Recibo"
    Protected Sub EmitirRecibo(ByVal pCodigo As String)

        Sql = " SELECT T.Titulo_Id AS TNumtit, " & vbCrLf & _
        " 0 AS TRecibo, " & vbCrLf & _
        " 0 AS TRegistro, " & vbCrLf & _
        " CONVERT(VARCHAR,T.TipoPagto) + ' - ' + CONVERT(VARCHAR,TP.Descricao) AS TFormaPagto, " & vbCrLf & _
        " T.Reprogramacao AS TVencimento, " & vbCrLf & _
        " T.DataBaixa AS TBaixa, " & vbCrLf & _
        " CONVERT(VARCHAR, DAY(T.DataBaixa)) AS TDia, " & vbCrLf & _
        " CONVERT(VARCHAR, MONTH(T.DataBaixa)) AS TMes, " & vbCrLf & _
        " CONVERT(VARCHAR, YEAR(T.DataBaixa)) AS TAno, " & vbCrLf & _
        " E.Cliente_Id AS CCnpj, " & vbCrLf & _
        " E.Nome AS ENome,  " & vbCrLf & _
        " E.Cidade AS ECidade,  " & vbCrLf & _
        " E.Estado AS EEstado, " & vbCrLf & _
        " E.Endereco AS EEndereco, " & vbCrLf & _
        " E.Cep AS ECep, " & vbCrLf & _
        " E.Inscricao AS EInscricao, " & vbCrLf & _
        " E.Telefone AS EFone, " & vbCrLf & _
        " E.Bairro AS EBairro, " & vbCrLf & _
        " E.Complemento AS EComplemento, " & vbCrLf & _
        " E.Numero AS ENumero, " & vbCrLf & _
        " C.Cliente_Id AS CCnpj, " & vbCrLf & _
        " C.Nome AS CNome,  " & vbCrLf & _
        " C.Cidade AS CCidade, " & vbCrLf & _
        " C.Estado AS CEstado, " & vbCrLf & _
        " C.Endereco AS CEndereco, " & vbCrLf & _
        " C.Cep AS CCep, " & vbCrLf & _
        " C.Inscricao AS CInscricao, " & vbCrLf & _
        " C.Telefone AS CFone, " & vbCrLf & _
        " C.Bairro AS CBairro, " & vbCrLf & _
        " C.Complemento AS CComplemento, " & vbCrLf & _
        " C.Numero AS CNumero, " & vbCrLf & _
        " CONVERT(VARCHAR, CxCB.Banco_Id) AS TBanco, " & vbCrLf & _
        " CONVERT(VARCHAR, CxCB.Agencia_Id) AS TAgencia, " & vbCrLf & _
        " CONVERT(VARCHAR, CxCB.DigitoAgencia_Id) AS TDigitoAgencia, " & vbCrLf & _
        " CONVERT(VARCHAR, CxCB.ContaCorrente_Id) AS TConta,  " & vbCrLf & _
        " CONVERT(VARCHAR, CxCB.DigitoConta_Id) AS TDigito, " & vbCrLf & _
        " T.Cheque AS TNumeroDoCheque, " & vbCrLf & _
        " Valores.ValorDoDocumento AS TValorDoDocumento, " & vbCrLf & _
        " 0 AS TDescontos, " & vbCrLf & _
        " Valores.Deducoes AS TDeducoes, " & vbCrLf & _
        " 0 AS TJuros, " & vbCrLf & _
        " Valores.Acrescimos AS TAcrescimos, " & vbCrLf & _
        " Valores.ValorLiquido AS TValor, " & vbCrLf & _
        " '' AS TExtenso, " & vbCrLf & _
        " (case when len(convert(nvarchar(4000),T.Observacoes)) > 0 then Historico + ' ' + convert(nvarchar(4000),T.Observacoes ) else T.Historico end) as THistorico " & vbCrLf & _
        " FROM Titulos T " & vbCrLf & _
        " INNER JOIN Clientes C  " & vbCrLf & _
        "    ON T.CliFor         = C.Cliente_Id " & vbCrLf & _
        "    AND T.EnderecoCliFor = C.Endereco_Id " & vbCrLf & _
        " INNER JOIN Clientes E  " & vbCrLf & _
        "    ON T.Empresa         = E.Cliente_Id " & vbCrLf & _
        "    AND T.EndEmpresa = E.Endereco_Id " & vbCrLf & _
        " Inner Join Carteira Ca " & vbCrLf & _
        "    on Ca.Carteira_Id = T.CarteiraDoTitulo " & vbCrLf & _
        " INNER JOIN Moedas M " & vbCrLf & _
        "   on M.Moeda_id = T.Moeda " & vbCrLf & _
        " INNER JOIN Provisoes P " & vbCrLf & _
        "    on P.Provisao_id = T.Provisao " & vbCrLf & _
        " Inner Join( " & vbCrLf & _
              " Select Tp.Titulo_Id, " & vbCrLf & _
                     " SUM(case " & vbCrLf & _
                           " when Tc.Conta_Id  = Tp.ContacontabilCliFor and Tp.RecPag in ('C','P') and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                           " when Tc.Conta_Id  = Tp.ContacontabilCliFor and Tp.RecPag = 'R'        and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                           " else 0 " & vbCrLf & _
                    " End " & vbCrLf & _
                         " ) as ValorDoDocumento, " & vbCrLf & _
                    " SUM(case " & vbCrLf & _
                           " when Tc.Conta_Id  not in (Tp.ContacontabilCliFor,Tp.ContaContabilRecPag) and Tp.RecPag in ('C','P') and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                           " when Tc.Conta_Id  not in (Tp.ContacontabilCliFor,Tp.ContaContabilRecPag) and Tp.RecPag = 'R'        and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                           " else 0 " & vbCrLf & _
                    " End " & vbCrLf & _
                         " ) as Acrescimos, " & vbCrLf & _
                      " SUM(case " & vbCrLf & _
                           " when Tc.Conta_Id  not in (Tp.ContacontabilCliFor,Tp.ContaContabilRecPag) and Tp.RecPag in ('C','P') and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                           " when Tc.Conta_Id  not in (Tp.ContacontabilCliFor,Tp.ContaContabilRecPag) and Tp.RecPag = 'R'        and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                           " else 0 " & vbCrLf & _
                    " End " & vbCrLf & _
                         " ) as Deducoes, " & vbCrLf & _
                      " SUM(case " & vbCrLf & _
                           " when Tc.Conta_Id  = Tp.ContaContabilRecPag and Tp.RecPag in ('C','P') and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                           " when Tc.Conta_Id  = Tp.ContaContabilRecPag and Tp.RecPag = 'R'        and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                           " else 0 " & vbCrLf & _
                    " End " & vbCrLf & _
                         " ) as ValorLiquido " & vbCrLf & _
        " from Titulos Tp " & vbCrLf & _
        " inner Join TitulosxContaContabil Tc " & vbCrLf & _
        " on Tc.Titulo_Id = Tp.Titulo_Id " & vbCrLf & _
        "INNER JOIN Moedas M " & vbCrLf & _
        " on M.Moeda_id = Tp.Moeda " & vbCrLf & _
        " Group by Tp.Titulo_Id " & vbCrLf & _
        " ) Valores " & vbCrLf & _
        " on Valores.Titulo_Id = T.Titulo_Id " & vbCrLf & _
        " LEFT JOIN TituloXDestinacao TxD " & vbCrLf & _
        "	ON T.Titulo_Id = TxD.Titulo_Id " & vbCrLf & _
        " LEFT JOIN ClientesXContasBancarias CxCB " & vbCrLf & _
        "   ON T.BancoCliFor = CxCB.Banco_Id " & vbCrLf & _
        "   and T.AgenciaCliFor = CxCB.Agencia_Id " & vbCrLf & _
        "   and T.DigitoAgenciaCliFor = CxCB.DigitoAgencia_Id " & vbCrLf & _
        "   and T.ContaCliFor = CxCB.ContaCorrente_Id " & vbCrLf & _
        "   and T.DigitoContaCliFor = CxCB.DigitoConta_Id " & vbCrLf & _
        "   and T.CliFor = CxCB.Cliente_Id " & vbCrLf & _
        " LEFT JOIN TiposDePagamentos AS TP " & vbCrLf & _
        "   On T.TipoPagto = TP.TipoDePagamento_Id " & vbCrLf & _
        " Where T.Titulo_Id = " & pCodigo

        Dim ds As New DataSet
        Dim empresa As String = String.Empty
        Dim xextenso As String
        Dim yextenso As String
        Dim mes As Integer = 0
        Dim j As Integer = 0
        Dim k As Integer = 0
        Dim Mesx As String = ""
        Dim dsRecibo As New DataSet
        Dim row As DataRow
        Dim RegistroI As String = ""
        Dim RegistroS As String = ""

        'IIf(RdPagar.Checked, "ReciboAPagar", "ReciboAReceber")
        ds = Banco.ConsultaDataSet(Sql, "ReciboAPagar")

        For Each row In ds.Tables("ReciboAPagar").Rows
            empresa = row("ENome")
            'Valor por Extenso
            yextenso = "("
            yextenso &= UCase(Funcoes.Extenso(row("TValor"), "Real", "Reais"))
            yextenso &= " *"
            xextenso = yextenso
            For j = 1 To (120 - Len(xextenso))
                xextenso &= " *"
            Next
            xextenso &= ")"
            row("TExtenso") = xextenso
            ''* Rotina de extenso fim 
            If txtObservacoes.Text.Length > 0 Then
                row("THistorico") = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim) & ". " & Funcoes.EliminarCaracteresEspeciais(txtObservacoes.Text.Trim) & ". "
            Else
                row("THistorico") = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim)
            End If
            'Mês Por Extenso
            row("TMes") = Funcoes.MesPorExtenso(row("TMes"))
            'Calculo de resgistro
            RegistroI = txtRegistro.Text()
            RegistroS = " "
            RegistroS = Trim(RegistroS)
            If Len(RegistroS) < 6 Then
                For k = 1 To (6 - Len(Trim(RegistroI)))
                    RegistroS &= "0"
                Next
            End If
            row("TRecibo") = Trim(RegistroS) & (Trim(RegistroI))
            row("TRegistro") = Trim(RegistroS) & (Trim(RegistroI))
        Next

        Dim param As New Dictionary(Of String, Object)
        param.Add("XNome", empresa)

        Funcoes.BindReport(Me.Page, ds, IIf(RdPagar.Checked, "Cr_ReciboPagar", "Cr_ReciboReceber"), eExportType.PDF, param)
    End Sub
#End Region

#End Region

End Class