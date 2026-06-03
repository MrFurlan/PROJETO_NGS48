Imports System.Globalization
Imports System.IO
Imports Boleto2Net
Imports BoletoNet
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports ArquivoRemessa = Boleto2Net.ArquivoRemessa
Imports Boleto = Boleto2Net.Boleto
Imports Sacado = Boleto2Net.Sacado
Imports TipoArquivo = Boleto2Net.TipoArquivo

Public Class BoletoBancario
    Inherits BasePage

    Private mensagemErro As String
    Private empresa() As String
    Private cliente() As String

#Region "Barra Botão"

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Not Funcoes.VerificaPermissao("BoletoBancario", "LEITURA") Then
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
                Exit Sub
            End If

            If ValidarSelecao() Then

                empresa = ddlEmpresa.SelectedValue.ToString.Split("-")

                If txtCodigoCliente.Value.ToString.Length = 0 Then
                    txtCodigoCliente.Value = "" & "-" & 0
                End If

                cliente = txtCodigoCliente.Value.ToString.Split("-")

                Dim listNotasFiscais As New [Lib].Negocio.ListNotasFiscais(empresa(0), empresa(1), txtDataInicial.Text, txtDataFinal.Text, cliente(0), cliente(1), "S", "", 1)

                listNotasFiscais.ForEach(Function(x) x.VencimentosNota.RemoveAll(Function(v) v.Agrupado = eAgrupamentoFinanceiro.Agrupado Or v.BoletoBancario Or v.CodigoProvisao = eProvisao.Baixa))

                If listNotasFiscais.Count > 0 Then
                    Dim ds As New DataSet
                    Dim tbNotas As New DataTable("Notas")
                    tbNotas.Columns.Add("Empresa", Type.GetType("System.String"))
                    tbNotas.Columns.Add("Cliente", Type.GetType("System.String"))
                    tbNotas.Columns.Add("Nome", Type.GetType("System.String"))
                    tbNotas.Columns.Add("E/S", Type.GetType("System.String"))
                    tbNotas.Columns.Add("Serie", Type.GetType("System.String"))
                    tbNotas.Columns.Add("Nota", Type.GetType("System.String"))
                    ds.Tables.Add(tbNotas)

                    For Each nf As [Lib].Negocio.NotaFiscal In listNotasFiscais.Where(Function(x) x.VencimentosNota.Any())
                        If txtNotaFiscal.Text.Length > 0 Then
                            If nf.Codigo = CInt(txtNotaFiscal.Text) Then
                                Dim drRow As DataRow = ds.Tables(0).NewRow()
                                drRow("Empresa") = nf.CodigoEmpresa & "-" & nf.EnderecoEmpresa
                                drRow("Cliente") = nf.CodigoCliente & "-" & nf.EnderecoCliente
                                drRow("Nome") = nf.Cliente.Nome
                                drRow("E/S") = IIf(nf.EntradaSaida = eEntradaSaida.Saida, "S", "E")
                                drRow("Serie") = nf.Serie
                                drRow("Nota") = nf.Codigo

                                ds.Tables(0).Rows.Add(drRow)
                            End If
                        Else
                            Dim drRow As DataRow = ds.Tables(0).NewRow()
                            drRow("Empresa") = nf.CodigoEmpresa & "-" & nf.EnderecoEmpresa
                            drRow("Cliente") = nf.CodigoCliente & "-" & nf.EnderecoCliente
                            drRow("Nome") = nf.Cliente.Nome
                            drRow("E/S") = IIf(nf.EntradaSaida = eEntradaSaida.Saida, "S", "E")
                            drRow("Serie") = nf.Serie
                            drRow("Nota") = nf.Codigo

                            ds.Tables(0).Rows.Add(drRow)
                        End If
                    Next

                    gridBoletoNotaFiscal.DataSource = ds
                    gridBoletoNotaFiscal.DataBind()

                    TabContainer1.ActiveTabIndex = 1
                Else
                    MsgBox(Me.Page, "Registro(s) não encontrado(s) para a seleção.")
                    Limpar()
                End If
            Else
                MsgBox(Me.Page, mensagemErro)
                Limpar()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Limpar()
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
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkGerarRemessa_Click(sender As Object, e As EventArgs) Handles lnkGerarRemessa.Click
        Try
            If Not Funcoes.VerificaPermissao("BoletoBancario", "GRAVAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro!")
                Exit Sub
            End If

            If ddlBanco.SelectedValue.Length = 0 Then
                MsgBox(Me.Page, "Banco não foi selecionado.", eTitulo.Info)
                Exit Sub
            End If

            Dim i = 0
            For Each row As GridViewRow In gridVencimentoBoleto.Rows
                Dim chk As CheckBox = row.Cells(0).FindControl("chkVencimento")
                If chk.Checked Then
                    i = i + 1
                End If
            Next

            If i = 0 Then
                MsgBox(Me.Page, "Selecione o(s) vencimento(s) para gerar o arquivo remessa", eTitulo.Erro)
                Exit Sub
            End If

            Dim ListaDeTitulos = New List(Of Titulo)
            Dim Boletos = New Boleto2Net.Boletos()

            For Each row As GridViewRow In gridVencimentoBoleto.Rows

                Dim chk As CheckBox = row.Cells(0).FindControl("chkVencimento")

                If chk.Checked Then
                    Dim codigo As String = TryCast(row.Cells(1), DataControlFieldCell).Text
                    CType(row.FindControl("imgBoletoPDFAntesDoEnvio"), ImageButton).Visible = True

                    Dim Titulos = From t In CType(Session("objVencimentos"), List(Of Titulo))
                                  Where t.Codigo = CInt(codigo)
                                  Select t

                    Dim Titulo = Titulos(0)
                    ListaDeTitulos.Add(Titulo)
                End If
            Next

            Dim strConta As String() = ddlBanco.SelectedValue.Split(";")
            Dim where = "Where Ativo = 1 AND Banco_Id = " & strConta(0) & " AND Agencia_Id = " & strConta(1) & " AND Conta_Id = " & strConta(3) & " AND Convenio = " & strConta(5)
            Dim parametros As New ParametrosDaCarteiraDeCobranca(where)

            If strConta(0) = "707" Then
                GerarRemessaDaycoval(ListaDeTitulos, parametros)
                AtualizarTitulos(ListaDeTitulos, parametros, where)

                For Each tit In ListaDeTitulos
                    Dim TitulosBoleto = New List(Of Titulo)
                    TitulosBoleto.Add(tit)

                    Dim daycovalBoletos = BoletosDaycoval(TitulosBoleto, parametros)
                    Dim _nomePdf = NomeNoPdf(TitulosBoleto)
                    GerarBoletos(daycovalBoletos, _nomePdf, False)
                Next

                MostrarBotaoDownload()

                Exit Sub
            End If

            If strConta(0) = "748" Then
                GerarRemessaSicredi(ListaDeTitulos, parametros)
                AtualizarTitulos(ListaDeTitulos, parametros, where)

                For Each tit In ListaDeTitulos
                    Dim TitulosBoleto = New List(Of Titulo)
                    TitulosBoleto.Add(tit)

                    Dim sicrediBoletos = BoletosSicredi(TitulosBoleto, parametros)
                    Dim _nomePdf = NomeNoPdf(TitulosBoleto)
                    GerarBoletos(sicrediBoletos, _nomePdf, False)
                Next

                MostrarBotaoDownload()

                Exit Sub
            End If

            If strConta(0) = "756" Then
                GerarRemessaSicoob(ListaDeTitulos, parametros)
                AtualizarTitulos(ListaDeTitulos, parametros, where)

                For Each tit In ListaDeTitulos
                    Dim TitulosBoleto = New List(Of Titulo)
                    TitulosBoleto.Add(tit)

                    Dim sicoobBoletos = BoletosSicoob(TitulosBoleto, parametros)
                    Dim _nomePdf = NomeNoPdf(TitulosBoleto)
                    GerarBoletos(sicoobBoletos, _nomePdf, False)
                Next

                MostrarBotaoDownload()

                Exit Sub
            End If

            If strConta(0) = "274" Then
                GerarRemessaMoneyPlus(ListaDeTitulos, parametros)
                AtualizarTitulos(ListaDeTitulos, parametros, where)

                For Each tit In ListaDeTitulos
                    Dim TitulosBoleto = New List(Of Titulo)
                    TitulosBoleto.Add(tit)

                    Dim moneyPLusBoletos = BoletosMoneyPlus(TitulosBoleto, parametros)
                    Dim _nomePdf = NomeNoPdf(TitulosBoleto)
                    GerarBoletos(moneyPLusBoletos, _nomePdf, False)
                Next

                MostrarBotaoDownload()

                Exit Sub
            End If

            If strConta(0) = "460" Then
                GerarRemessaUnavanti(ListaDeTitulos, parametros)
                AtualizarTitulos(ListaDeTitulos, parametros, where)

                For Each tit In ListaDeTitulos
                    Dim TitulosBoleto = New List(Of Titulo)
                    TitulosBoleto.Add(tit)

                    Dim unavantiBoletos = BoletosUnavanti(TitulosBoleto, parametros)
                    Dim _nomePdf = NomeNoPdf(TitulosBoleto)
                    GerarBoletos(unavantiBoletos, _nomePdf, False)
                Next

                MostrarBotaoDownload()

                Exit Sub
            End If

            GerarBoletos(Boletos, ListaDeTitulos, ddlBanco.SelectedValue)

            If Not Directory.Exists(Server.MapPath("~/RemessaBancaria/" & parametros.Convenio.ToString)) Then Directory.CreateDirectory(Server.MapPath("~/RemessaBancaria/" & parametros.Convenio.ToString))

            'Gerar Arquivo de Remessa
            Dim stream = New MemoryStream()
            Dim remessa = New ArquivoRemessa(Boletos.Banco, parametros.TipoArquivoCNAB, (parametros.SequenciaDeRemessa + 1))
            remessa.GerarArquivoRemessa(Boletos, stream)

            Dim NomeArquivo2 As String = "RemessaBancaria/" & parametros.Convenio.ToString & "/AR" & (parametros.SequenciaDeRemessa + 1).ToString("00000") & parametros.NomeNoArquivo & ".rem"

            Session("NomeArquivoDownload") = NomeArquivo2

            Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)

            Dim fileInbites As Byte()

            Using stream
                fileInbites = stream.ToArray()
            End Using

            Dim oFileStream As System.IO.FileStream
            oFileStream = New System.IO.FileStream(NomeArquivo, FileMode.Create)
            oFileStream.Write(fileInbites, 0, fileInbites.Length)
            oFileStream.Close()

            'Gerar Boletos PDF
            For Each b In Boletos

                Try

                    Dim BoletoBancario = New Boleto2Net.BoletoBancario With {.Boleto = b}
                    Dim pdf = BoletoBancario.MontaBytesPDF(False)

                    Dim tit = (From t In ListaDeTitulos
                               Where t.Codigo = CInt(b.NumeroControleParticipante)
                               Select t).First()

                    'Dim historico As String() = tit.Historico.Split(",")
                    'Dim nomePDF = historico(0).Replace("REF.", "").Replace("-1-S", "").Replace(Environment.NewLine, "")
                    'If Not nomePDF.Contains("AGRUPAMENTO") Then
                    '    If historico.Length > 1 Then
                    '        nomePDF += " " & historico(1).Replace("/", " de ").Replace(Environment.NewLine, "") & " " & tit.Cliente.Nome.Replace("/", "").Replace(".", "") & " - " & tit.Prorrogacao.ToString("dd-MM-yyyy")
                    '    Else
                    '        nomePDF += " " & historico(0).Replace("/", " de ").Replace(Environment.NewLine, "") & " " & tit.Cliente.Nome.Replace("/", "").Replace(".", "") & " - " & tit.Prorrogacao.ToString("dd-MM-yyyy")
                    '    End If
                    'Else
                    '    nomePDF += " " & tit.Cliente.Nome.Replace("/", "").Replace(".", "") & " - " & tit.Prorrogacao.ToString("dd-MM-yyyy")
                    'End If

                    Dim nomePDF = NomeNoPdf(ListaDeTitulos)
                    Dim pathPDF = Server.MapPath("Boletos/" & nomePDF & ".pdf")
                    File.WriteAllBytes(pathPDF, pdf)

                Catch ex As Exception
                    Throw New Exception(ex.Message)
                End Try

            Next

            AtualizarTitulos(ListaDeTitulos, parametros, where)

            MostrarBotaoDownload()
            'CarregarTitulosEnviadosAoBanco()
            MsgBox(Me.Page, "Arquivo de remessa gerado com sucesso!", eTitulo.Erro)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConfirmarVencimentos_Click(sender As Object, e As EventArgs) Handles lnkConfirmarVencimentos.Click
        Try

            If Not Funcoes.VerificaPermissao("BoletoBancario", "GRAVAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro!")
                Exit Sub
            End If

            If CType(Session("objVencimentos"), List(Of Titulo)).Any() Then
                CarregarVencimentos()
                TabContainer1.ActiveTabIndex = 2
            Else
                MsgBox(Me.Page, "Para confimar é necessário selecionar ao menos uma NF", eTitulo.Erro)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkProcessarRetorno_Click(sender As Object, e As EventArgs) Handles lnkProcessarRetorno.Click
        Try
            If Not Funcoes.VerificaPermissao("BoletoBancario", "GRAVAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro!")
                Exit Sub
            End If

            If ddlBancoRetorno.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Selecione o Convênio.", eTitulo.Info)
                Exit Sub
            End If
            ProcessarRetornoRemessa()
            TabContainer1.ActiveTabIndex = 3
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível processar o retorno da remessa bancária.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkConsultarRetorno_Click(sender As Object, e As EventArgs) Handles lnkConsultarRetorno.Click
        Try
            If Not Funcoes.VerificaPermissao("BoletoBancario", "LEITURA") Then
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
                Exit Sub
            End If

            If ddlBancoRetorno.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Selecione o Convênio.", eTitulo.Info)
                Exit Sub
            End If
            CarregarTitulosEnviadosAoBanco()
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível consultar os Títulos enviados ao Banco.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkAgruparPagamento_Click(sender As Object, e As EventArgs) Handles lnkAgruparPagamento.Click
        Try
            If Not Funcoes.VerificaPermissao("BoletoBancario", "GRAVAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro!")
                Exit Sub
            End If

            lnkAgruparPagamento.Visible = False
            lnkGerarRemessa.Visible = False
            lnkNovo.Visible = True
            lnkLimparAgrupamento.Visible = True
            divArquivo.Visible = False
            divAgrupamento.Visible = True
            divVencimentoAgrupamento.Visible = True

            Dim vencimentos = CType(Session("objVencimentos"), List(Of Titulo))
            Dim clientes = From t In vencimentos
                           Where t.Agrupado = eAgrupamentoFinanceiro.NaoAgrupado
                           Group t By t.Cliente.Codigo Into c = Group
                           Select c.First().Cliente.Codigo, c.First().Cliente.Nome

            ddlAgrupamento.Items.Clear()
            ddlAgrupamento.Items.Add(New ListItem With {.Text = "Selecione o cliente", .Value = -1})

            For Each c In clientes
                ddlAgrupamento.Items.Add(New ListItem With {.Text = c.Nome, .Value = c.Codigo})
            Next

            For j = 0 To gridVencimentoBoleto.Rows.Count - 1
                CType(gridVencimentoBoleto.Rows(j).FindControl("chkVencimento"), CheckBox).Checked = False
            Next

            If clientes.Count() = 1 Then
                ddlAgrupamento.SelectedValue = clientes.First().Codigo
            End If

            gridVencimentoBoleto.DataSource = vencimentos.Where(Function(t) Not t.Agrupado = eAgrupamentoFinanceiro.Agrupado).ToList()
            gridVencimentoBoleto.DataBind()

        Catch ex As Exception
            Dim message = ex.Message
        End Try
    End Sub

    Protected Sub lnkLimparAgrupamento_Click(sender As Object, e As EventArgs)
        Try
            LimparAgrupamento()
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs)
        Try
            If (ValidarAgrupamento()) Then
                Dim tituloMestre As Titulo = Nothing
                Dim ListaDeTitulos As New List(Of Titulo)
                Dim vencimentos = CType(Session("objVencimentos"), List(Of Titulo))
                Dim historicoDosTitulos As String = ""
                For Each row As GridViewRow In gridVencimentoBoleto.Rows

                    Dim chk As CheckBox = row.Cells(0).FindControl("chkVencimento")

                    If chk.Checked Then

                        Dim codigo As String = TryCast(row.Cells(1), DataControlFieldCell).Text

                        'Titulo Mestre
                        If tituloMestre Is Nothing Then
                            tituloMestre = New Titulo(codigo)
                            Dim numTitulo As New [Lib].Negocio.Numerador(1)
                            tituloMestre.Codigo = numTitulo.Sequencia + 1
                            tituloMestre.IUD = "I"
                            tituloMestre.Historico = "AGRUPAMENTO DAS"
                            tituloMestre.Vencimento = CDate(txtVencimentoAgrupamento.Text)
                            tituloMestre.Prorrogacao = tituloMestre.Vencimento
                            tituloMestre.Baixa = tituloMestre.Vencimento
                            tituloMestre.ValorDoDocumento = CDec(txtValorAgrupamento.Text.Replace(".", ""))
                            tituloMestre.Deducoes = 0
                            tituloMestre.Juros = 0
                            tituloMestre.Descontos = 0
                            tituloMestre.Acrescimos = 0
                            tituloMestre.Agrupado = eAgrupamentoFinanceiro.Mestre
                            ListaDeTitulos.Add(tituloMestre)
                        End If

                        Dim tit As New Titulo(codigo)
                        tit.IUD = "U"
                        tit.Agrupado = eAgrupamentoFinanceiro.Agrupado
                        tit.RegistroMestre = tituloMestre.Codigo
                        tit.Vencimento = tituloMestre.Vencimento
                        tit.Prorrogacao = tituloMestre.Vencimento
                        tit.Baixa = tituloMestre.Vencimento
                        Dim historico As String() = tit.Historico.Split(",")
                        Dim nota = historico(0).Replace("REF.", "").Replace("-1-S", "").Trim()

                        If Not tituloMestre.Historico.Contains(nota) Then
                            tituloMestre.Historico += " " & nota
                        End If

                        historicoDosTitulos += Environment.NewLine & tit.Codigo & " - " & tit.Historico
                        ListaDeTitulos.Add(tit)
                    End If
                Next

                'Add histórico completo
                tituloMestre.Historico += "," & historicoDosTitulos

                For Each tit In ListaDeTitulos
                    Dim result = tit.Salvar()
                    If result AndAlso tit.Agrupado.Equals(eAgrupamentoFinanceiro.Agrupado) Then
                        vencimentos.RemoveAll(Function(ByVal c) c.Codigo = tit.Codigo)
                    Else
                        vencimentos.Add(tit)
                    End If
                Next

                Session("objVencimentos") = vencimentos
                LimparAgrupamento()
            Else
                MsgBox(Me.Page, mensagemErro)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("BoletoBancario", "ACESSAR") Then
                BuncarUnidadeDeNegocio()
                txtDataInicial.Text = Now.ToString("dd/MM/yyyy")
                txtDataFinal.Text = Now.ToString("dd/MM/yyyy")
                Limpar()
                CarregarBancos()
                If Not Directory.Exists(Server.MapPath("~/Boletos")) Then Directory.CreateDirectory(Server.MapPath("~/Boletos"))
                If Not Directory.Exists(Server.MapPath("~/RemessaBancaria")) Then Directory.CreateDirectory(Server.MapPath("~/RemessaBancaria"))
                If Not Directory.Exists(Server.MapPath("~/RetornoBancario")) Then Directory.CreateDirectory(Server.MapPath("~/RetornoBancario"))
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If

        If IsPostBack And fup.HasFile Then
            Dim strConta As String() = ddlBancoRetorno.SelectedValue.Split(";")
            Dim where = "Where Ativo = 1 AND Banco_Id = " & strConta(0) & " AND Agencia_Id = " & strConta(1) & " AND Conta_Id = " & strConta(3) & " AND Convenio = " & strConta(5)
            Dim parametros As New ParametrosDaCarteiraDeCobranca(where)

            If Not Directory.Exists(Server.MapPath("~/RetornoBancario/" & parametros.Convenio.ToString)) Then Directory.CreateDirectory(Server.MapPath("~/RetornoBancario/" & parametros.Convenio.ToString))

            Dim strCaminho As String = Server.MapPath("RetornoBancario/" & parametros.Convenio.ToString & "/") & Path.GetFileName(fup.FileName)
            If Dir(strCaminho).Length > 0 Then Kill(strCaminho)

            'Salvamos o mesmo 
            lblArquivoRetorno.Text = fup.FileName
            divRetorno.Visible = True
            fup.PostedFile.SaveAs(strCaminho)

            If strConta(0) = 104 Then
                CarregarRetornoCEF(parametros)
            ElseIf strConta(0) = 756 Then
                CarregarRetornoSicoob(parametros)
            Else
                CarregarRetornoRemessa(parametros)
            End If

            TabContainer1.ActiveTabIndex = 3
            fup.Visible = False
        End If
    End Sub

    Protected Sub ddlBancoRetorno_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarTitulosEnviadosAoBanco()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkNotaFiscal_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim chkNotaFiscal As CheckBox = CType(sender, CheckBox)
            Dim row As GridViewRow = CType(chkNotaFiscal.NamingContainer, GridViewRow)
            hidBoletoNotaFiscal.Value = row.RowIndex

            gridVencimentoBoleto.DataSource = Nothing
            gridVencimentoBoleto.DataBind()

            empresa = gridBoletoNotaFiscal.Rows(row.RowIndex).Cells(1).Text().Split("-")
            cliente = gridBoletoNotaFiscal.Rows(row.RowIndex).Cells(2).Text().Split("-")

            Dim NfConsulta As New [Lib].Negocio.NotaFiscal
            NfConsulta.CodigoEmpresa = empresa(0)
            NfConsulta.EnderecoEmpresa = empresa(1)
            NfConsulta.CodigoCliente = cliente(0)
            NfConsulta.EnderecoCliente = cliente(1)
            NfConsulta.EntradaSaida = IIf(gridBoletoNotaFiscal.Rows(row.RowIndex).Cells(4).Text() = "S", eEntradaSaida.Saida, eEntradaSaida.Entrada)
            NfConsulta.Serie = gridBoletoNotaFiscal.Rows(row.RowIndex).Cells(5).Text()
            NfConsulta.Codigo = gridBoletoNotaFiscal.Rows(row.RowIndex).Cells(6).Text()

            NfConsulta = New [Lib].Negocio.NotaFiscal(NfConsulta)

            If chkNotaFiscal.Checked Then
                Dim listaNf = CType(Session("objNfs"), List(Of Integer))
                Dim vencimentos = CType(Session("objVencimentos"), List(Of Titulo))
                If NfConsulta.CodigoPedido > 0 Then
                    If NfConsulta.VencimentosNota.Count > 0 Then

                        Dim qtdeTitAgrupado = NfConsulta.VencimentosNota.Where(Function(t) t.Agrupado = eAgrupamentoFinanceiro.Agrupado).Count()

                        If qtdeTitAgrupado = NfConsulta.VencimentosNota.Count Then
                            MsgBox(Me.Page, "Nota Fiscal com vencimento(s) agrupados(s)")
                            chkNotaFiscal.Checked = False
                            Exit Sub
                        End If

                        listaNf.Add(NfConsulta.Codigo)
                        NfConsulta.VencimentosNota.RemoveAll(Function(t) t.Agrupado = eAgrupamentoFinanceiro.Agrupado Or t.CodigoProvisao = eProvisao.Baixa)
                        vencimentos.AddRange(NfConsulta.VencimentosNota)
                        Session("objNfs") = listaNf
                        Session("objVencimentos") = vencimentos
                    Else
                        MsgBox(Me.Page, "Vencimento(s) não encontrado(s).")
                    End If

                End If

                Dim i As Integer

                For i = 0 To gridBoletoNotaFiscal.Rows.Count - 1
                    If listaNf.Contains(CInt(gridBoletoNotaFiscal.Rows(i).Cells(6).Text)) Then
                        CType(gridBoletoNotaFiscal.Rows(i).FindControl("chkNotaFiscal"), CheckBox).Checked = True
                    Else
                        CType(gridBoletoNotaFiscal.Rows(i).FindControl("chkNotaFiscal"), CheckBox).Checked = False
                    End If
                Next
            Else
                Dim vencimentos = CType(Session("objVencimentos"), List(Of Titulo))
                For Each tit In NfConsulta.VencimentosNota
                    vencimentos.RemoveAll(Function(t As Titulo) t.Codigo = tit.Codigo)
                Next
                Session("objVencimentos") = vencimentos
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnDownload_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDownload.Click
        Try
            Dim strConta As String() = ddlBanco.SelectedValue.Split(";")
            Dim where = "Where Ativo = 1 AND Banco_Id = " & strConta(0) & " AND Agencia_Id = " & strConta(1) & " AND Conta_Id = " & strConta(3) & " AND Convenio = " & strConta(5)
            Dim parametros As New ParametrosDaCarteiraDeCobranca(where)
            'Para o Daycoval
            If strConta(0) = "707" Then
                Download("RemessaBancaria/" & parametros.Convenio.ToString, parametros.NomeNoArquivo + ".txt")
                Exit Sub
            End If

            If strConta(0) = "274" Then
                'Gerar Arquivo de Remessa
                Download("RemessaBancaria/" & parametros.Convenio.ToString, parametros.NomeNoArquivo + ".REM")
                Exit Sub
            End If

            If strConta(0) = "460" Then
                'Gerar Arquivo de Remessa
                Download("RemessaBancaria/" & parametros.Convenio.ToString, parametros.NomeNoArquivo + ".REM")
                Exit Sub
            End If

            If strConta(0) = "748" OrElse strConta(0) = "756" Then
                'Codigo Beneficiario
                Dim nomeArq = Integer.Parse(parametros.CodigoDaEmpresaNoBanco).ToString("00000")
                'Codigo do Mês e Dia
                Dim mes = DateTime.Now.Month
                Dim dia = DateTime.Now.Day

                If (mes <= 9) Then
                    nomeArq += mes.ToString()
                ElseIf (mes = 10) Then
                    nomeArq += "O"
                ElseIf (mes = 11) Then
                    nomeArq += "N"
                ElseIf (mes = 12) Then
                    nomeArq += "D"
                End If

                nomeArq += dia.ToString("00")

                Download("RemessaBancaria/" & parametros.Convenio.ToString, nomeArq + ".txt")
                Exit Sub
            End If
            'Outros Bancos
            Download("RemessaBancaria/" & parametros.Convenio.ToString, "AR" & (parametros.SequenciaDeRemessa).ToString("00000") & parametros.NomeNoArquivo & ".REM")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub btnBoleto_Click(ByVal sender As Object, ByVal e As CommandEventArgs)
        Try
            Dim ListaDeTitulos = New List(Of Titulo)
            Dim Boletos = New Boleto2Net.Boletos()
            Dim tit = New Titulo(CInt(e.CommandName))
            ListaDeTitulos.Add(tit)

            Dim strConta As String() = ddlBanco.SelectedValue.Split(";")
            Dim where = "Where Ativo = 1 AND Banco_Id = " & strConta(0) & " AND Agencia_Id = " & strConta(1) & " AND Conta_Id = " & strConta(3) & " AND Convenio = " & strConta(5)
            Dim parametros As New ParametrosDaCarteiraDeCobranca(where)

            If tit.CodigoBancoPagador = 707 Then
                Dim daycovalBoletos = BoletosDaycoval(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoletos(daycovalBoletos, _nomePdf, True)
                Exit Sub
            End If

            If tit.CodigoBancoPagador = 748 Then
                Dim sicrediBoletos = BoletosSicredi(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoletos(sicrediBoletos, _nomePdf, True)
                Exit Sub
            End If

            If tit.CodigoBancoPagador = 756 Then
                Dim sicoobBoletos = BoletosSicoob(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoletos(sicoobBoletos, _nomePdf, True)
                Exit Sub
            End If

            If tit.CodigoBancoPagador = 274 Then
                Dim moneyPlusBoletos = BoletosMoneyPlus(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoletos(moneyPlusBoletos, _nomePdf, True)
                Exit Sub
            End If

            If tit.CodigoBancoPagador = 460 Then
                Dim unavantiBoletos = BoletosUnavanti(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoletos(unavantiBoletos, _nomePdf, True)
                Exit Sub
            End If

            GerarBoletos(Boletos, ListaDeTitulos, ddlBanco.SelectedValue)

            Dim Titulo = ListaDeTitulos.FirstOrDefault(Function(x) x.Codigo = CInt(e.CommandName))

            Dim nomePDF As String = ""

            'Gerar Boletos PDF
            For Each b In Boletos
                Dim BoletoBancario = New Boleto2Net.BoletoBancario With {.Boleto = b}
                Dim pdf = BoletoBancario.MontaBytesPDF(False)
                nomePDF = NomeNoPdf(ListaDeTitulos)

                Dim pathPDF = Server.MapPath("Boletos/" & nomePDF & ".pdf")
                File.WriteAllBytes(pathPDF, pdf)
            Next

            Download("Boletos", nomePDF & ".pdf", "application/pdf")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        Finally
            'Kill(Server.MapPath("~/Boletos/" & e.CommandName & ".pdf"))
        End Try
    End Sub

    Protected Sub btnBoletoRetorno_Click(ByVal sender As Object, ByVal e As CommandEventArgs)
        Try
            Dim ListaDeTitulos = New List(Of Titulo)
            Dim Boletos = New Boleto2Net.Boletos()
            Dim tit = New Titulo(CInt(e.CommandName))
            ListaDeTitulos.Add(tit)

            Dim strConta As String() = ddlBancoRetorno.SelectedValue.Split(";")
            Dim where = "Where Ativo = 1 AND Banco_Id = " & strConta(0) & " AND Agencia_Id = " & strConta(1) & " AND Conta_Id = " & strConta(3) & " AND Convenio = " & strConta(5)
            Dim parametros As New ParametrosDaCarteiraDeCobranca(where)

            If tit.CodigoBancoPagador = 707 Then
                Dim daycovalBoletos = BoletosDaycoval(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoletos(daycovalBoletos, _nomePdf, True)
                Exit Sub
            End If

            If tit.CodigoBancoPagador = 748 Then
                Dim sicredilBoletos = BoletosSicredi(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoletos(sicredilBoletos, _nomePdf, True)
                Exit Sub
            End If

            If tit.CodigoBancoPagador = 756 Then
                Dim sicredilBoletos = BoletosSicoob(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoletos(sicredilBoletos, _nomePdf, True)
                Exit Sub
            End If

            If tit.CodigoBancoPagador = 274 Then
                Dim sicredilBoletos = BoletosMoneyPlus(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoletos(sicredilBoletos, _nomePdf, True)
                Exit Sub
            End If

            If tit.CodigoBancoPagador = 460 Then
                Dim unavantiBoletos = BoletosUnavanti(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoletos(unavantiBoletos, _nomePdf, True)
                Exit Sub
            End If

            GerarBoletos(Boletos, ListaDeTitulos, ddlBancoRetorno.SelectedValue)

            Dim Titulo = ListaDeTitulos.FirstOrDefault(Function(x) x.Codigo = CInt(e.CommandName))

            Dim nomePDF As String = ""

            'Gerar Boletos PDF
            For Each b In Boletos
                Dim BoletoBancario = New Boleto2Net.BoletoBancario With {.Boleto = b}
                Dim pdf = BoletoBancario.MontaBytesPDF(False)
                nomePDF = NomeNoPdf(ListaDeTitulos)

                Dim pathPDF = Server.MapPath("Boletos/" & nomePDF & ".pdf")
                File.WriteAllBytes(pathPDF, pdf)
            Next

            Download("Boletos", nomePDF & ".pdf", "application/pdf")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        Finally
            'Kill(Server.MapPath("~/Boletos/" & e.CommandName & ".pdf"))
        End Try
    End Sub

    Protected Sub btnBoletos_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim strConta As String() = ddlBancoRetorno.SelectedValue.Split(";")
            Dim where = "Where Ativo = 1 AND Banco_Id = " & strConta(0) & " AND Agencia_Id = " & strConta(1) & " AND Conta_Id = " & strConta(3) & " AND Convenio = " & strConta(5)
            Dim Boletos = New Boleto2Net.Boletos()

            Dim parametros As New ParametrosDaCarteiraDeCobranca(where)
            Dim empresa As New Cliente(parametros.CodigoEmpresa, parametros.EnderecoEmpresa)

            'Daycoval 707 - Sicredi 748 - SICOOB 756
            If parametros.CodigoBanco = 707 Or parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 756 Or parametros.CodigoBanco = 274 Or parametros.CodigoBanco = 460 Then
                Dim listaDeTitulos As New List(Of Titulo)
                For Each row As GridViewRow In gridRetornoTitulos.Rows
                    Dim chk As CheckBox = row.Cells(0).FindControl("chkImpBoletos")
                    If chk.Checked Then
                        Dim tit As New Titulo(CInt(CType(row.FindControl("hpTitulo"), HyperLink).Text))
                        listaDeTitulos.Add(tit)
                    End If
                Next

                If listaDeTitulos.Count < 2 Then
                    MsgBox(Me.Page, "Para impressão em bloco deve ser selecionado no mínimo 2 registros.", eTitulo.Info)
                    Exit Sub
                End If

                If parametros.CodigoBanco = 707 Then
                    Dim daycovalBoletos = BoletosDaycoval(listaDeTitulos, parametros)
                    GerarBoletos(daycovalBoletos, "Boletos", True)
                    Exit Sub
                ElseIf parametros.CodigoBanco = 756 Then
                    Dim sicoobBoletos = BoletosSicoob(listaDeTitulos, parametros)
                    GerarBoletos(sicoobBoletos, "Boletos", True)
                    Exit Sub
                ElseIf parametros.CodigoBanco = 274 Then
                    Dim moneyPlusBoletos = BoletosMoneyPlus(listaDeTitulos, parametros)
                    GerarBoletos(moneyPlusBoletos, "Boletos", True)
                    Exit Sub
                ElseIf parametros.CodigoBanco = 460 Then
                    Dim unavantiBoletos = BoletosUnavanti(listaDeTitulos, parametros)
                    GerarBoletos(unavantiBoletos, "Boletos", True)
                    Exit Sub
                Else
                    'Sicredi
                    Dim sicrediBoletos = BoletosSicredi(listaDeTitulos, parametros)
                    GerarBoletos(sicrediBoletos, "Boletos", True)
                    Exit Sub
                End If

            End If

            'Cabeçalho
            Boletos.Banco = Boleto2Net.Banco.Instancia(parametros.CodigoBanco)
            Boletos.Banco.Cedente = New Boleto2Net.Cedente With
            {
                .CPFCNPJ = parametros.CodigoEmpresa,
            .Nome = parametros.NomeDaEmpresaNoBanco,
            .ContaBancaria = New Boleto2Net.ContaBancaria With
            {
                .Agencia = parametros.CodigoAgencia,
                .DigitoAgencia = parametros.DigitoAgencia,
                .Conta = IIf(parametros.CodigoBanco = Boleto2Net.Bancos.Bradesco, parametros.Conta, parametros.Convenio),
                .DigitoConta = parametros.DigitoConta,
                .CarteiraPadrao = parametros.CarteiraPadrao,
                .TipoCarteiraPadrao = TipoCarteira.CarteiraCobrancaSimples,
                .TipoFormaCadastramento = TipoFormaCadastramento.ComRegistro,
                .TipoImpressaoBoleto = TipoImpressaoBoleto.Empresa,
                .TipoDocumento = TipoDocumento.Tradicional
            },
            .Endereco = New Boleto2Net.Endereco With
            {
                .LogradouroEndereco = empresa.Endereco,
                .LogradouroNumero = empresa.Numero,
                .LogradouroComplemento = empresa.Complemento,
                .Bairro = empresa.Bairro,
                .Cidade = empresa.Cidade,
                .UF = empresa.Estado.Codigo,
                .CEP = empresa.CEP
             },
            .Codigo = parametros.Convenio,
            .CodigoDV = parametros.DigitoConta,
            .CodigoTransmissao = String.Empty
            }

            Boletos.Banco.FormataCedente()

            For Each row As GridViewRow In gridRetornoTitulos.Rows
                Dim chk As CheckBox = row.Cells(0).FindControl("chkImpBoletos")
                If chk.Checked Then

                    Dim Titulo As New Titulo(CInt(CType(row.FindControl("hpTitulo"), HyperLink).Text))

                    'Adiciona os Títulos
                    Dim Boleto As New Boleto(Boletos.Banco)
                    Boleto.Sacado = New Sacado With
                    {
                        .CPFCNPJ = Titulo.Cliente.Codigo,
                        .Nome = Titulo.Cliente.Nome,
                        .Endereco = New Boleto2Net.Endereco With
                        {
                            .LogradouroEndereco = Titulo.Cliente.Endereco,
                            .LogradouroNumero = Titulo.Cliente.Numero,
                            .LogradouroComplemento = Titulo.Cliente.Complemento,
                            .Bairro = Titulo.Cliente.Bairro,
                            .Cidade = Titulo.Cliente.Cidade,
                            .UF = Titulo.Cliente.Estado.Codigo,
                            .CEP = Titulo.Cliente.CEP
                        }
                    }

                    Boleto.CodigoOcorrencia = "01"
                    Boleto.DescricaoOcorrencia = "Remessa Registrar"

                    Boleto.NumeroDocumento = Titulo.Codigo.ToString()
                    Boleto.NumeroControleParticipante = Titulo.Codigo.ToString()
                    Boleto.NossoNumero = Titulo.Codigo.ToString()

                    Boleto.DataEmissao = DateTime.Now.Date

                    'Furlan - 26/07/2021 - Alterado para prorrogação pois podem ter mudado a data antes de gerar os Boletos
                    'Boleto.DataVencimento = Titulo.Vencimento
                    Boleto.DataVencimento = Titulo.Prorrogacao

                    Boleto.ValorTitulo = Titulo.ValorLiquido
                    Boleto.Aceite = "N"
                    Boleto.EspecieDocumento = TipoEspecieDocumento.DM
                    Boleto.MensagemInstrucoesCaixa = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."

                    If (parametros.JurosMes > 0) Then
                        Boleto.DataMulta = Titulo.Prorrogacao.AddDays(1)
                        Boleto.PercentualMulta = parametros.JurosMes
                        Boleto.ValorMulta = CDec(Boleto.ValorTitulo * Boleto.PercentualMulta / 100)

                        Dim Instrucao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & Boleto.ValorMulta.ToString("C2", CultureInfo.CurrentCulture)

                        If (String.IsNullOrEmpty(Boleto.MensagemInstrucoesCaixa)) Then
                            Boleto.MensagemInstrucoesCaixa = Instrucao
                        Else
                            Boleto.MensagemInstrucoesCaixa += Environment.NewLine + Instrucao
                        End If
                    End If

                    If (parametros.MoraDiaria > 0) Then
                        Boleto.DataJuros = Titulo.Prorrogacao.AddDays(1)

                        'Furlan - 26-07-2021 - Química é diferente pois cobra 3,50% ao mês dividido por 30. 
                        If parametros.CodigoEmpresa = "05272759000147" Then
                            Boleto.PercentualJurosDia = (parametros.MoraDiaria / 30)
                            Boleto.ValorJurosDia = Boleto.ValorTitulo * Boleto.PercentualJurosDia / 100
                        Else
                            Boleto.PercentualJurosDia = parametros.MoraDiaria
                            Boleto.ValorJurosDia = (Boleto.ValorTitulo / 30) * Boleto.PercentualJurosDia / 100
                        End If

                        Dim Instrucao = "<br /> APOS VENCIMENTO COBRAR R$ " & Boleto.ValorJurosDia.ToString("C2", CultureInfo.CurrentCulture) & " POR DIA DE ATRASO"

                        If (String.IsNullOrEmpty(Boleto.MensagemInstrucoesCaixa)) Then
                            Boleto.MensagemInstrucoesCaixa = Instrucao
                        Else
                            Boleto.MensagemInstrucoesCaixa += Environment.NewLine + Instrucao
                        End If
                    End If

                    'Boleto.CodigoProtesto = TipoCodigoProtesto.NaoProtestar
                    'Boleto.CodigoBaixaDevolucao = TipoCodigoBaixaDevolucao.NaoBaixarNaoDevolver
                    'Boleto.DiasBaixaDevolucao = 0

                    'Protesto boleto
                    Boleto.CodigoProtesto = IIf(parametros.DiasParaProtesto > 0, TipoCodigoProtesto.ProtestarDiasUteis, TipoCodigoProtesto.NaoProtestar)
                    If parametros.DiasParaProtesto > 0 Then
                        Boleto.CodigoInstrucao1 = "06"
                        Boleto.CodigoInstrucao2 = parametros.DiasParaProtesto.ToString()
                        Boleto.DiasProtesto = parametros.DiasParaProtesto

                        If Left(parametros.CodigoEmpresa, 8) = "40938762" OrElse Left(parametros.CodigoEmpresa, 8) = "49673784" OrElse Left(parametros.CodigoEmpresa, 8) = "05366261" OrElse Left(parametros.CodigoEmpresa, 8) = "62747840" OrElse Left(parametros.CodigoEmpresa, 8) = "62780383" Then
                            Boleto.MensagemInstrucoesCaixa += Environment.NewLine + "<br /> PROTESTAR APÓS " + parametros.DiasParaProtesto.ToString() + " DIAS DO VENCIMENTO."
                        Else
                            Boleto.MensagemInstrucoesCaixa += Environment.NewLine + "<br /> PROTESTAR APÓS " + parametros.DiasParaProtesto.ToString() + " DIAS ÚTEIS."
                        End If
                    End If

                    Boleto.CodigoBaixaDevolucao = IIf(parametros.CodigoBanco = 104, TipoCodigoBaixaDevolucao.BaixarDevolver, TipoCodigoBaixaDevolucao.NaoBaixarNaoDevolver)
                    'Boleto.DiasBaixaDevolucao = 0

                    Boleto.ValidarDados()
                    Boletos.Add(Boleto)
                End If
            Next

            If Boletos.Count < 2 Then
                MsgBox(Me.Page, "Para impressão em bloco deve ser selecionado no mínimo 2 registros.", eTitulo.Info)
                Exit Sub
            End If


            Dim html = New StringBuilder()

            'Gerar Boletos HTML
            For Each b In Boletos
                Using BoletoBancario As Boleto2Net.BoletoBancario = New Boleto2Net.BoletoBancario With {.Boleto = b, .OcultarInstrucoes = False, .MostrarComprovanteEntrega = False, .MostrarEnderecoCedente = True}
                    html.Append("<div style=""page-break-after: always;"">")
                    html.Append(BoletoBancario.MontaHtml())
                    html.Append("</div>")
                End Using
            Next

            Dim pathPDF = Server.MapPath("Boletos/Boletos.pdf")

            Dim pdf = New NReco.PdfGenerator.HtmlToPdfConverter().GeneratePdf(html.ToString)

            Using fs As FileStream = New FileStream(pathPDF, FileMode.Create)
                fs.Write(pdf, 0, pdf.Length)
                fs.Close()
            End Using

            Download("Boletos", "Boletos.pdf", "application/pdf")

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub chkTodosNotaFiscal_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim chkTodosNotaFiscal As CheckBox = CType(sender, CheckBox)

        For Each row As GridViewRow In gridBoletoNotaFiscal.Rows
            Dim chk As CheckBox = row.FindControl("chkNotaFiscal")
            chk.Checked = chkTodosNotaFiscal.Checked
            If (chk.Checked) Then
                chkNotaFiscal_CheckedChanged(chk, Nothing)
            Else
                Session("objVencimentos") = New List(Of Titulo)
            End If
        Next

    End Sub

    Protected Sub chkTodosVencimento_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim chkTodosNotaFiscal As CheckBox = CType(sender, CheckBox)
        Dim valor As Decimal = 0
        For Each row As GridViewRow In gridVencimentoBoleto.Rows
            Dim chk As CheckBox = row.FindControl("chkVencimento")
            chk.Checked = chkTodosNotaFiscal.Checked

            If lnkNovo.Visible AndAlso chk.Checked Then
                Dim v = row.Cells(4).Text.Replace(".", "")
                valor += CDec(v)
            End If
        Next

        txtValorAgrupamento.Text = String.Format("{0:n2}", valor)

    End Sub

    Protected Sub imbExcluirRetorno_Click(sender As Object, e As ImageClickEventArgs)
        Try
            fup.Visible = True
            divRetorno.Visible = False
            lblArquivoRetorno.Text = ""
            CarregarTitulosEnviadosAoBanco()
        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao excluir arquivo de retorno")
        Finally
            'Kill(Server.MapPath("~/Boletos/" & e.CommandName & ".pdf"))
        End Try
    End Sub

    Protected Sub ddlAgrupamento_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            Dim ddl = CType(sender, DropDownList)
            If ddl.SelectedValue <> -1 Then

                Dim vencimentos = From t In CType(Session("objVencimentos"), List(Of Titulo))
                                  Where t.Cliente.Codigo = ddl.SelectedValue
                                  Select t

                gridVencimentoBoleto.DataSource = vencimentos
                gridVencimentoBoleto.DataBind()
            Else
                gridVencimentoBoleto.DataSource = CType(Session("objVencimentos"), List(Of Titulo))
                gridVencimentoBoleto.DataBind()
            End If

            For j = 0 To gridVencimentoBoleto.Rows.Count - 1
                CType(gridVencimentoBoleto.Rows(j).FindControl("chkVencimento"), CheckBox).Checked = False
            Next

        Catch ex As Exception
            Dim message = ex.Message
        End Try
    End Sub

    Protected Sub chkVencimento_CheckedChanged(sender As Object, e As EventArgs)
        Dim valor As Decimal = 0
        If lnkNovo.Visible Then
            For Each row As GridViewRow In gridVencimentoBoleto.Rows
                Dim chk As CheckBox = row.FindControl("chkVencimento")
                If chk.Checked Then
                    Dim v = row.Cells(4).Text.Replace(".", "")
                    valor += CDec(v)
                End If

            Next
            txtValorAgrupamento.Text = String.Format("{0:n2}", valor)
        End If
    End Sub

#End Region

#Region "Methods"

    Private Sub BuncarUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)
    End Sub

    Private Sub Limpar()
        txtNotaFiscal.Text = ""

        lnkGerarRemessa.Enabled = False
        lnkAgruparPagamento.Enabled = False
        ddlBanco.Enabled = False

        gridBoletoNotaFiscal.DataSource = Nothing
        gridBoletoNotaFiscal.DataBind()

        gridVencimentoBoleto.DataSource = Nothing
        gridVencimentoBoleto.DataBind()

        gridRetornoTitulos.DataSource = Nothing
        gridRetornoTitulos.DataBind()

        Session.Remove("objNotaFiscalBoleto")
        Dim nfs As New List(Of Integer)
        Session("objNfs") = nfs
        Dim vencimentos As New List(Of Titulo)
        Session("objVencimentos") = vencimentos
    End Sub

    Function ValidarSelecao() As Boolean
        If ddlUnidadeNegocio.SelectedIndex = 0 Then
            mensagemErro = "Unidade de Negócio não foi selecionada."
            Return False
        ElseIf ddlEmpresa.SelectedIndex = 0 Then
            mensagemErro = "Empresa não foi selecionada."
            Return False
        Else
            Return True
        End If
    End Function

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objCliente" & HID.Value) IsNot Nothing Then
            Dim cli As Cliente = Session("objCliente" & HID.Value)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objCliente" & HID.Value)
        End If
    End Sub

    Private Sub Download(pPath As String, pNomeArquivo As String, Optional pType As String = "text/plain")
        Try
            Response.Clear()
            Response.ContentType = pType
            Response.AppendHeader("Content-Disposition", "attachment; filename=" & pNomeArquivo)
            Response.TransmitFile(Server.MapPath("~/" & pPath & "/" & pNomeArquivo))
            Response.Flush()
            Response.End()
        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao fazer o download do arquivo! contate o suporte")
        End Try
    End Sub

    Protected Sub CarregarVencimentos()
        Try
            'CÓDIGO GERADO PARA IMPRIMIR BOLETOS ANTIGOS
            'Dim venc As New ListTitulo(" UsuarioBoletoBancarioDate BETWEEN '2021-04-01' AND '2021-05-17' and len(HistoricoRemessa) = 0 and Situacao = 101")
            'Dim vencimentos As New List(Of Titulo)
            'For Each t As Titulo In venc
            '    vencimentos.Add(t)
            'Next
            'Session("objVencimentos") = vencimentos
            '----------------------- ATÉ AQUI ------------------------------------------

            Dim vencimentos = CType(Session("objVencimentos"), List(Of Titulo))

            gridVencimentoBoleto.DataSource = vencimentos
            gridVencimentoBoleto.DataBind()

            Dim j As Integer
            For j = 0 To gridVencimentoBoleto.Rows.Count - 1
                Dim codigo = CInt(gridVencimentoBoleto.Rows(j).Cells(1).Text)
                Dim tit = vencimentos.Find(Function(t As Titulo) t.Codigo = codigo)
                If (tit.BoletoBancario) Then
                    CType(gridVencimentoBoleto.Rows(j).FindControl("chkVencimento"), CheckBox).Checked = False
                    CType(gridVencimentoBoleto.Rows(j).FindControl("chkVencimento"), CheckBox).Enabled = False
                End If

            Next

            lnkGerarRemessa.Enabled = True
            lnkAgruparPagamento.Enabled = True
            ddlBanco.Enabled = True

        Catch ex As Exception
            Dim e = ex.Message
        End Try
    End Sub

    Public Sub ProcessarRetornoRemessa()

        Dim i = 0
        For Each row As GridViewRow In gridRetornoTitulos.Rows
            Dim chk As CheckBox = row.Cells(0).FindControl("chkVencimento")
            If chk.Checked Then
                i = i + 1
            End If
        Next

        If i = 0 Then
            MsgBox(Me.Page, "Selecione o(s) vencimento(s) para gerar o arquivo remessa", eTitulo.Erro)
            Exit Sub
        End If

        Dim resultado = False
        Dim listaDeTitulos As New List(Of Titulo)

        Try
            For Each row As GridViewRow In gridRetornoTitulos.Rows

                Dim chk As CheckBox = row.Cells(0).FindControl("chkVencimento")
                Dim _registroId As String = CType(row.FindControl("hpTitulo"), HyperLink).Text

                If chk.Checked Then

                    Dim tit As New Titulo(_registroId)

                    If tit.CodigoProvisao <> eProvisao.Baixa Then

                        Dim retorno As String() = TryCast(row.Cells(4), DataControlFieldCell).Text.Split("-")
                        Dim enumerador As eSituacaoBancaria = CType(CInt(retorno(0)), eSituacaoBancaria)

                        If (tit.Agrupado = eAgrupamentoFinanceiro.Mestre) Then
                            Dim titulosAgrupados As New ListTitulo("RegistroMestre = " & tit.Codigo)

                            Dim primeiro As Boolean = True 'Server para pegar o Juros do Mestre e colocar no primeiro título dos filhos - Furlan - 29/10/2021
                            Dim temJuros As Decimal = 0

                            If CDec(row.Cells(7).Text) > 0 Then temJuros = CDec(row.Cells(7).Text)

                            For Each t As Titulo In titulosAgrupados

                                If primeiro AndAlso temJuros > 0 Then
                                    primeiro = False
                                Else
                                    temJuros = 0
                                End If

                                AtualizarTitulos(t, enumerador, CDate(row.Cells(5).Text), temJuros)
                                listaDeTitulos.Add(t)
                            Next
                        End If

                        AtualizarTitulos(tit, enumerador, CDate(row.Cells(5).Text), CDec(row.Cells(7).Text))
                        listaDeTitulos.Add(tit)
                    End If
                End If
            Next

            For Each t As Titulo In listaDeTitulos
                resultado = t.Salvar()
                If Not resultado Then
                    Exit For
                End If
            Next

            If resultado = True Then
                CarregarTitulosEnviadosAoBanco()
                MsgBox(Me.Page, "Arquivo de retorno processado com sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, "Erro ao processar arquivo de retorno.", eTitulo.Erro)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message.ToString(), eTitulo.Erro)
        Finally

        End Try
    End Sub

    Public Sub CarregarRetornoRemessa(ByRef parametros As ParametrosDaCarteiraDeCobranca)
        Dim arquivo As String = Server.MapPath("~/RetornoBancario/" & parametros.Convenio.ToString & "/" & fup.FileName)
        Dim objArquivo As New StreamReader(arquivo)
        Dim strLinha As String
        'Dim campos() As String
        Dim intLinha As Integer

        Dim ErroMSG As String = String.Empty

        Dim titulosArqRetorno = New List(Of Object)
        Dim titLiquidados = New List(Of Integer)
        Try
            Do While objArquivo.Peek >= 0
                intLinha += 1
                strLinha = objArquivo.ReadLine()
                If intLinha <> 1 Then
                    If Mid(strLinha, 1, 1).ToString().Equals("1") Then

                        Dim codigo = IIf(parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 460, Mid(strLinha, 117, 10).Trim(), Mid(strLinha, 38, 25).Trim())
                        Dim codRetorno = Mid(strLinha, 109, 2)
                        Dim DataPgto = Mid(strLinha, 111, 2) & "/" & Mid(strLinha, 113, 2) & "/" & Mid(strLinha, 115, 2)
                        Dim ValorJuros = Mid(strLinha, 267, 13)

                        Dim valorTotal = Mid(strLinha, 254, 13)

                        Dim numeroInt As Integer
                        'Se a captura do numero do titulo for maior que o padrão do sistema
                        If Not Integer.TryParse(codigo, numeroInt) Then
                            codigo = Mid(strLinha, 117, 10).Trim()
                        End If

                        DataPgto = DateTime.ParseExact(DataPgto, "dd/MM/yy", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy")

                        ErroMSG = "TITULO " & codigo.ToString & " - PAGAMENTO " & DataPgto.ToString & " - VER RETORNO " & codRetorno.ToString

                        For i = 0 To gridRetornoTitulos.Rows.Count - 1

                            Dim codigoNaGrid = CType(gridRetornoTitulos.Rows(i).FindControl("hpTitulo"), HyperLink).Text
                            If codigo.ToString.Length > 0 AndAlso CInt(codigoNaGrid) = CInt(codigo) AndAlso Not titLiquidados.Any(Function(x) x = CInt(codigo)) Then

                                Dim enumerador As eSituacaoBancaria = CType(CInt(codRetorno), eSituacaoBancaria)
                                gridRetornoTitulos.Rows(i).Cells(4).Text = CInt(enumerador).ToString("00") & " - " & Textos.GetEnumDescription(enumerador)
                                gridRetornoTitulos.Rows(i).Cells(5).Text = DataPgto

                                If enumerador = eSituacaoBancaria.LiquidacaoNormal Or enumerador = eSituacaoBancaria.BaixaSimples Then
                                    titLiquidados.Add(CInt(codigo))

                                    'Banco Unavanti tem somente o valor total pago(Valor do tiutlo + juros/mora)
                                    If (parametros.CodigoBanco = 460) Then
                                        ValorJuros = CDec(valorTotal.Insert(11, ",")) - CDec(gridRetornoTitulos.Rows(i).Cells(6).Text)
                                    Else
                                        ValorJuros = (CDec(ValorJuros) / 100)
                                    End If
                                End If

                                titulosArqRetorno.Add(New With
                                                     {
                                                       .Codigo = CInt(codigo),
                                                       .Cidade = gridRetornoTitulos.Rows(i).Cells(2).Text,
                                                       .Historico = CType(gridRetornoTitulos.Rows(i).FindControl("lblHistorico"), Label).Text,
                                                       .Observacoes = gridRetornoTitulos.Rows(i).Cells(4).Text,
                                                       .Prorrogacao = DataPgto,
                                                       .Juros = CDec(ValorJuros),
                                    .ValorLiquido = CDec(gridRetornoTitulos.Rows(i).Cells(6).Text) + CDec(ValorJuros)
                                                     })

                            End If
                        Next
                    End If
                End If
            Loop

        Catch ex As Exception
            'Dim e = ex.Message
            MsgBox(Me.Page, ErroMSG & ". " & ex.Message.ToString, eTitulo.Erro)
        Finally
            fup.PostedFile.InputStream.Dispose()
            'fup.Dispose()
            objArquivo.Dispose()
            TabContainer1.ActiveTabIndex = 3

            'gridRetornoTitulos.DataSource = From tit In titulosArqRetorno
            '                                Select tit.Codigo, tit.Prorrogacao, tit.Observacoes, tit.Historico, tit.ValorLiquido, tit.Juros, CodigoCifrado = Funcoes.Cifrar("ContasAReceber-" & tit.Codigo)

            gridRetornoTitulos.DataSource = From tit In titulosArqRetorno
                                            Select New With {.Codigo = tit.Codigo, .Cidade = tit.Cidade, .Prorrogacao = tit.Prorrogacao, .Observacoes = tit.Observacoes, .Historico = tit.Historico, .ValorLiquido = tit.ValorLiquido, .Juros = tit.Juros, .CodigoCifrado = Funcoes.Cifrar("ContasAReceber-" & tit.Codigo)}

            gridRetornoTitulos.DataBind()

            For i = 0 To gridRetornoTitulos.Rows.Count - 1
                CType(gridRetornoTitulos.Rows(i).FindControl("chkVencimento"), CheckBox).Checked = True
            Next
        End Try
    End Sub

    Public Sub CarregarRetornoCEF(ByRef parametros As ParametrosDaCarteiraDeCobranca)
        Dim arquivo As String = Server.MapPath("~/RetornoBancario/" & parametros.Convenio.ToString & "/" & fup.FileName)
        Dim objArquivo As New StreamReader(arquivo)
        Dim strLinha As String

        Dim ErroMSG As String = String.Empty

        Dim ListTitulo As New [Lib].Negocio.ListTitulo
        Dim titulosArqRetorno = New List(Of Object)
        Try
            Do While objArquivo.Peek >= 0
                strLinha = objArquivo.ReadLine()
                If Mid(strLinha, 8, 1).Trim() = 3 AndAlso Mid(strLinha, 14, 1).Trim() = "T" Then
                    If Mid(strLinha, 1, 1).ToString().Equals("1") Then
                        'Segmento T
                        Dim codigo = Mid(strLinha, 59, 11).Trim()
                        Dim codRetorno = Mid(strLinha, 16, 2)
                        'Segmento U
                        strLinha = objArquivo.ReadLine()
                        Dim DataPgto = Mid(strLinha, 138, 2) & "/" & Mid(strLinha, 140, 2) & "/" & Mid(strLinha, 142, 4)

                        ErroMSG = "TITULO " & codigo.ToString & " - PAGAMENTO " & DataPgto.ToString & " - VER RETORNO " & codRetorno.ToString

                        For i = 0 To gridRetornoTitulos.Rows.Count - 1
                            Dim codigoNaGrid = CType(gridRetornoTitulos.Rows(i).FindControl("hpTitulo"), HyperLink).Text
                            If codigoNaGrid = codigo And CType(gridRetornoTitulos.Rows(i).FindControl("chkVencimento"), CheckBox).Checked = False Then
                                CType(gridRetornoTitulos.Rows(i).FindControl("chkVencimento"), CheckBox).Checked = True
                                Dim enumerador As eSituacaoBancaria = CType(CInt(codRetorno), eSituacaoBancaria)
                                gridRetornoTitulos.Rows(i).Cells(4).Text = CInt(enumerador).ToString("00") & " - " & Textos.GetEnumDescription(enumerador)
                                gridRetornoTitulos.Rows(i).Cells(5).Text = DataPgto

                                titulosArqRetorno.Add(New With
                                                     {
                                                       .Codigo = codigo,
                                                       .Cidade = gridRetornoTitulos.Rows(i).Cells(2).Text,
                                                       .Historico = CType(gridRetornoTitulos.Rows(i).FindControl("lblHistorico"), Label).Text,
                                                       .Observacoes = gridRetornoTitulos.Rows(i).Cells(4).Text,
                                                       .Prorrogacao = DataPgto,
                                                       .Juros = 0,
                                                       .ValorLiquido = CDec(gridRetornoTitulos.Rows(i).Cells(6).Text)
                                                     })

                            End If
                        Next
                    End If
                End If
            Loop

        Catch ex As Exception
            'Dim e = ex.Message
            MsgBox(Me.Page, ErroMSG & ". " & ex.Message.ToString, eTitulo.Erro)
        Finally
            fup.PostedFile.InputStream.Dispose()
            'fup.Dispose()
            objArquivo.Dispose()
            TabContainer1.ActiveTabIndex = 3
            'gridRetornoTitulos.DataSource = From tit In titulosArqRetorno
            '                                Select tit.Codigo, tit.Prorrogacao, tit.Observacoes, tit.Historico, tit.ValorLiquido, tit.Juros, CodigoCifrado = Funcoes.Cifrar("ContasAReceber-" & tit.Codigo)

            gridRetornoTitulos.DataSource = From tit In titulosArqRetorno
                                            Select New With {.Codigo = tit.Codigo, .Cidade = tit.Cidade, .Prorrogacao = tit.Prorrogacao, .Observacoes = tit.Observacoes, .Historico = tit.Historico, .ValorLiquido = tit.ValorLiquido, .Juros = tit.Juros, .CodigoCifrado = Funcoes.Cifrar("ContasAReceber-" & tit.Codigo)}

            gridRetornoTitulos.DataBind()

            For i = 0 To gridRetornoTitulos.Rows.Count - 1
                CType(gridRetornoTitulos.Rows(i).FindControl("chkVencimento"), CheckBox).Checked = True
            Next
        End Try
    End Sub

    Public Sub CarregarRetornoSicoob(ByRef parametros As ParametrosDaCarteiraDeCobranca)
        Dim arquivo As String = Server.MapPath("~/RetornoBancario/" & parametros.Convenio.ToString & "/" & fup.FileName)
        Dim objArquivo As New StreamReader(arquivo)
        Dim strLinha As String

        Dim ErroMSG As String = String.Empty

        Dim ListTitulo As New [Lib].Negocio.ListTitulo
        Dim titulosArqRetorno = New List(Of Object)
        Try
            Do While objArquivo.Peek >= 0
                strLinha = objArquivo.ReadLine()
                If Mid(strLinha, 8, 1).Trim() = 3 AndAlso Mid(strLinha, 14, 1).Trim() = "T" Then

                    'Segmento T
                    Dim codigo = Mid(strLinha, 106, 25).Trim()
                    'Segmento U
                    strLinha = objArquivo.ReadLine()
                    Dim DataPgto = Mid(strLinha, 138, 2) & "/" & Mid(strLinha, 140, 2) & "/" & Mid(strLinha, 142, 4)
                    Dim codRetorno = Mid(strLinha, 16, 2)

                    ErroMSG = "TITULO " & codigo.ToString & " - PAGAMENTO " & DataPgto.ToString & " - VER RETORNO " & codRetorno.ToString

                    For i = 0 To gridRetornoTitulos.Rows.Count - 1
                        Dim codigoNaGrid = CType(gridRetornoTitulos.Rows(i).FindControl("hpTitulo"), HyperLink).Text
                        If codigoNaGrid = codigo And CType(gridRetornoTitulos.Rows(i).FindControl("chkVencimento"), CheckBox).Checked = False Then
                            CType(gridRetornoTitulos.Rows(i).FindControl("chkVencimento"), CheckBox).Checked = True
                            Dim enumerador As eSituacaoBancaria = CType(CInt(codRetorno), eSituacaoBancaria)
                            gridRetornoTitulos.Rows(i).Cells(4).Text = CInt(enumerador).ToString("00") & " - " & Textos.GetEnumDescription(enumerador)
                            gridRetornoTitulos.Rows(i).Cells(5).Text = DataPgto

                            titulosArqRetorno.Add(New With
                                                 {
                                                   .Codigo = codigo,
                                                   .Cidade = gridRetornoTitulos.Rows(i).Cells(2).Text,
                                                   .Historico = CType(gridRetornoTitulos.Rows(i).FindControl("lblHistorico"), Label).Text,
                                                   .Observacoes = gridRetornoTitulos.Rows(i).Cells(4).Text,
                                                   .Prorrogacao = DataPgto,
                                                   .Juros = 0,
                                                   .ValorLiquido = CDec(gridRetornoTitulos.Rows(i).Cells(6).Text)
                                                 })

                        End If
                    Next

                End If
            Loop

        Catch ex As Exception
            'Dim e = ex.Message
            MsgBox(Me.Page, ErroMSG & ". " & ex.Message.ToString, eTitulo.Erro)
        Finally
            fup.PostedFile.InputStream.Dispose()
            'fup.Dispose()
            objArquivo.Dispose()
            TabContainer1.ActiveTabIndex = 3
            'gridRetornoTitulos.DataSource = From tit In titulosArqRetorno
            '                                Select tit.Codigo, tit.Prorrogacao, tit.Observacoes, tit.Historico, tit.ValorLiquido, tit.Juros, CodigoCifrado = Funcoes.Cifrar("ContasAReceber-" & tit.Codigo)

            gridRetornoTitulos.DataSource = From tit In titulosArqRetorno
                                            Select New With {.Codigo = tit.Codigo, .Cidade = tit.Cidade, .Prorrogacao = tit.Prorrogacao, .Observacoes = tit.Observacoes, .Historico = tit.Historico, .ValorLiquido = tit.ValorLiquido, .Juros = tit.Juros, .CodigoCifrado = Funcoes.Cifrar("ContasAReceber-" & tit.Codigo)}

            gridRetornoTitulos.DataBind()

            For i = 0 To gridRetornoTitulos.Rows.Count - 1
                CType(gridRetornoTitulos.Rows(i).FindControl("chkVencimento"), CheckBox).Checked = True
            Next
        End Try
    End Sub

    Public Sub CarregarTitulosEnviadosAoBanco()

        Dim strConta As String() = ddlBancoRetorno.SelectedValue.Split(";")
        Dim where = "BancoPagador = " & strConta(0) & " AND AgenciaPagadora = " & strConta(1) & " AND ContaPagadora = " & strConta(3)
        Dim Boletos = New Boleto2Net.Boletos()

        gridRetornoTitulos.DataSource = Nothing

        where += " AND Situacao = " & CInt(eSituacao.RemessaBancaria) & " AND Provisao <> " & CInt(eProvisao.Baixa)
        Dim titulos As New ListTitulo(where)

        For Each tit In titulos
            If tit.SituacaoBancaria = CInt(eSituacaoBancaria.AguardandoArquivoDeRetorno) Then
                tit.Observacoes = CInt(eSituacaoBancaria.AguardandoArquivoDeRetorno).ToString("00") & " - " & Textos.GetEnumDescription(eSituacaoBancaria.AguardandoArquivoDeRetorno)
            End If
        Next

        gridRetornoTitulos.DataSource = From tit In titulos
                                        Select New With {.Codigo = tit.Codigo, .Cidade = tit.Empresa.Cidade, .Prorrogacao = tit.Prorrogacao, .Observacoes = tit.Observacoes, .Historico = tit.Historico, .ValorLiquido = tit.ValorLiquido, .Juros = tit.Juros, .CodigoCifrado = Funcoes.Cifrar("ContasAReceber-" & tit.Codigo)}
        gridRetornoTitulos.DataBind()
    End Sub

    Function ValidarAgrupamento() As Boolean
        Dim i = 0
        For Each row As GridViewRow In gridVencimentoBoleto.Rows
            Dim chk As CheckBox = row.Cells(0).FindControl("chkVencimento")
            If chk.Checked Then
                i = i + 1
            End If
        Next

        If i = 0 Then
            mensagemErro = "Selecione o(s) vencimento(s) para gerar o agrupamento"
            Return False
        ElseIf i < 2 Then
            mensagemErro = "Para fazer o agrupamento é necessário selecionar 2 ou mais títulos"
            Return False
        ElseIf ddlAgrupamento.SelectedIndex = 0 Then
            mensagemErro = "Cliente não foi selecionado."
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtVencimentoAgrupamento.Text) Then
            mensagemErro = "Vencimento não foi preenchido."
            Return False
        Else
            Return True
        End If
    End Function

    Public Sub LimparAgrupamento()
        lnkAgruparPagamento.Visible = True
        lnkGerarRemessa.Visible = True
        lnkNovo.Visible = False
        lnkLimparAgrupamento.Visible = False
        divArquivo.Visible = True
        divAgrupamento.Visible = False
        divVencimentoAgrupamento.Visible = False
        txtVencimentoAgrupamento.Text = ""
        txtValorAgrupamento.Text = ""
        CarregarVencimentos()
    End Sub

    Public Sub AtualizarTitulos(ByRef tit As Titulo, ByVal enumerador As eSituacaoBancaria, ByVal dataDeBaixa As Date, ByVal Juros As Decimal)

        Dim strConta As String() = ddlBancoRetorno.SelectedValue.Split(";")
        Dim where = "Where Ativo = 1 AND Banco_Id = " & strConta(0) & " AND Agencia_Id = " & strConta(1) & " AND Conta_Id = " & strConta(3) & " AND Convenio = " & strConta(5)
        Dim carteria As New ParametrosDaCarteiraDeCobranca(where) 'TODO: Codigo do convenio bancario
        tit.IUD = "U"

        Dim vlrdoc As Decimal
        Dim vlrliq As Decimal

        Select Case enumerador
            Case eSituacaoBancaria.LiquidacaoNormal
                tit.CodigoEmpresaPagadora = carteria.CodigoEmpresa
                tit.EndEmpresaPagadora = carteria.EnderecoEmpresa

                tit.CodigoProvisao = eProvisao.Baixa

                dataDeBaixa = dataDeBaixa.AddDays(1)

                tit.Baixa = Funcoes.ValidaDataUtil(tit.CodigoEmpresa, tit.EnderecoEmpresa, dataDeBaixa)

                tit.CodigoSituacao = eSituacao.Normal
                tit.CodigoBancoPagador = carteria.CodigoBanco
                tit.CodigoAgenciaPagadora = carteria.CodigoAgencia
                tit.TipoContaPagadora = "C"
                tit.ContaPagadora = carteria.Conta.ToString()
                tit.DigitoContaPagadora = carteria.DigitoConta
                tit.ContaContabilPagadora = carteria.ContaContabil
                tit.UsuarioBaixa = Session("ssNomeUsuario")
                tit.UsuarioBaixaData = Today()

                tit.CodigoTipoPgto = 4 'Boleto Bançario

                'Para atualizar valor em Dólar - Furlan - 04/06/2020
                vlrdoc = tit.ValorDoDocumento

                'Ajustando valor Líquido caso tenha Juros - Furlan - 29/09/2021
                If Juros > 0 Then
                    If tit.Agrupado = eAgrupamentoFinanceiro.Mestre Then
                        vlrdoc += Juros
                    Else
                        tit.Juros = Juros
                    End If
                End If

                tit.IndiceTitulo = tit.DolarizaBaixa(tit.Baixa.ToString("yyyy/MM/dd"), IIf(tit.Agrupado = eAgrupamentoFinanceiro.Mestre, vlrdoc, tit.ValorDoDocumento), IIf(tit.CodigoIndexador = 99, 2, tit.CodigoIndexador))

                If tit.IndiceTitulo > 0 OrElse Juros > 0 Then tit.ValorDoDocumento = vlrdoc

            Case eSituacaoBancaria.LiquidacaoEmCartorio
                tit.CodigoEmpresaPagadora = carteria.CodigoEmpresa
                tit.EndEmpresaPagadora = carteria.EnderecoEmpresa

                tit.CodigoProvisao = eProvisao.Baixa

                dataDeBaixa = dataDeBaixa.AddDays(1)

                tit.Baixa = Funcoes.ValidaDataUtil(tit.CodigoEmpresa, tit.EnderecoEmpresa, dataDeBaixa)

                tit.CodigoSituacao = eSituacao.Normal
                tit.CodigoBancoPagador = carteria.CodigoBanco
                tit.CodigoAgenciaPagadora = carteria.CodigoAgencia
                tit.TipoContaPagadora = "C"
                tit.ContaPagadora = carteria.Conta.ToString()
                tit.DigitoContaPagadora = carteria.DigitoConta
                tit.ContaContabilPagadora = carteria.ContaContabil
                tit.UsuarioBaixa = Session("ssNomeUsuario")
                tit.UsuarioBaixaData = Today()

                tit.CodigoTipoPgto = 4 'Boleto Bançario

                'Para atualizar valor em Dólar - Furlan - 04/06/2020
                vlrdoc = tit.ValorDoDocumento

                'Ajustando valor Líquido caso tenha Juros - Furlan - 29/09/2021
                If Juros > 0 Then
                    If tit.Agrupado = eAgrupamentoFinanceiro.Mestre Then
                        vlrdoc += Juros
                    Else
                        tit.Juros = Juros
                    End If
                End If

                tit.IndiceTitulo = tit.DolarizaBaixa(tit.Baixa.ToString("yyyy/MM/dd"), IIf(tit.Agrupado = eAgrupamentoFinanceiro.Mestre, vlrdoc, tit.ValorDoDocumento), IIf(tit.CodigoIndexador = 99, 2, tit.CodigoIndexador))

                If tit.IndiceTitulo > 0 OrElse Juros > 0 Then tit.ValorDoDocumento = vlrdoc

            Case eSituacaoBancaria.BaixaSimples
                tit.CodigoSituacao = eSituacao.Normal
                tit.BoletoBancario = False

                tit.CodigoEmpresaPagadora = String.Empty
                tit.EndEmpresaPagadora = 0
                tit.CodigoBancoPagador = 0
                tit.CodigoAgenciaPagadora = String.Empty
                tit.DigitoAgenciaPagadora = String.Empty
                tit.ContaPagadora = String.Empty
                tit.DigitoContaPagadora = String.Empty
                tit.ContaContabilPagadora = String.Empty

                tit.CodigoTipoPgto = 1

            Case eSituacaoBancaria.LiquidacaoParcial
                tit.CodigoSituacao = eSituacao.Normal
                tit.BoletoBancario = False

                tit.CodigoEmpresaPagadora = String.Empty
                tit.EndEmpresaPagadora = 0
                tit.CodigoBancoPagador = 0
                tit.CodigoAgenciaPagadora = String.Empty
                tit.DigitoAgenciaPagadora = String.Empty
                tit.ContaPagadora = String.Empty
                tit.DigitoContaPagadora = String.Empty
                tit.ContaContabilPagadora = String.Empty

                tit.CodigoTipoPgto = 1


            Case eSituacaoBancaria.BaixaPorLiquidado
                'Para Nutri e Baxi liberar título pois foi descontado para cobrança - Furlan - 20-08-2021
                If Left(tit.CodigoEmpresaPagadora, 8) = "05366261" OrElse Left(tit.CodigoEmpresaPagadora, 8) = "38198213" OrElse Left(tit.CodigoEmpresaPagadora, 8) = "40938762" OrElse Left(tit.CodigoEmpresaPagadora, 8) = "49673784" OrElse Left(tit.CodigoEmpresaPagadora, 8) = "62747840" OrElse Left(tit.CodigoEmpresaPagadora, 8) = "62780383" Then
                    tit.CodigoSituacao = eSituacao.Normal
                    tit.BoletoBancario = False

                    tit.CodigoEmpresaPagadora = String.Empty
                    tit.EndEmpresaPagadora = 0
                    tit.CodigoBancoPagador = 0
                    tit.CodigoAgenciaPagadora = String.Empty
                    tit.DigitoAgenciaPagadora = String.Empty
                    tit.ContaPagadora = String.Empty
                    tit.DigitoContaPagadora = String.Empty
                    tit.ContaContabilPagadora = String.Empty

                    tit.CodigoTipoPgto = 1
                End If

            Case eSituacaoBancaria.EntradaRejeitada
                tit.CodigoSituacao = eSituacao.Normal
                tit.BoletoBancario = False

                tit.CodigoEmpresaPagadora = String.Empty
                tit.EndEmpresaPagadora = 0
                tit.CodigoBancoPagador = 0
                tit.CodigoAgenciaPagadora = String.Empty
                tit.DigitoAgenciaPagadora = String.Empty
                tit.ContaPagadora = String.Empty
                tit.DigitoContaPagadora = String.Empty
                tit.ContaContabilPagadora = String.Empty

                tit.CodigoTipoPgto = 1

        End Select

        tit.SituacaoRemessaBancaria = enumerador
        tit.HistoricoRemessa = tit.HistoricoRemessa & Environment.NewLine & "ARQUIVO DE RETORNO - DATA: " & DateTime.Now.ToString("dd/MM/yyyy") & " " & CInt(enumerador).ToString("00") & " - " & Textos.GetEnumDescription(enumerador)

        If enumerador = eSituacaoBancaria.BaixaPorLiquidado Then tit.HistoricoRemessa = tit.HistoricoRemessa & " - FOI PARA DESCONTO"

    End Sub

    Private Sub CarregarBancos()
        Dim Empresa As String() = ddlEmpresa.SelectedValue.ToString.Split("-")
        Dim where As String = "Where isnull(FolhaDePagamento,0) = 0 AND Ativo = 1 AND left(Empresa_Id,8) = '" & Left(HttpContext.Current.Session("ssEmpresa"), 8) + "'"
        Dim objReceber As New [Lib].Negocio.ListParametrosDaCarteiraDeCobranca(where)

        ddlBanco.Items.Clear()

        For Each row As [Lib].Negocio.ParametrosDaCarteiraDeCobranca In objReceber

            Dim banco As New [Lib].Negocio.Banco(row.CodigoBanco)

            Dim agencia = IIf(row.DigitoAgencia.Length > 0, Funcoes.AlinharDireita(row.CodigoAgencia, 4, " ") & "-" & row.DigitoAgencia, Funcoes.AlinharDireita(row.CodigoAgencia, 6, "0"))
            Dim conta = IIf(row.DigitoConta.Length > 0, Funcoes.AlinharDireita(row.Conta, 10, " ") & "-" & row.DigitoConta, Funcoes.AlinharDireita(row.Conta, 10, " "))

            ddlBanco.Items.Add(New ListItem(Funcoes.AlinharEsquerda(banco.Descricao, 35, ".") & " - AG: " & agencia & " - CTA: " & conta & " CONVÊNIO: " & row.Convenio,
                                 row.CodigoBanco & ";" & row.CodigoAgencia & ";" & row.DigitoAgencia & ";" & row.Conta & ";" & row.DigitoConta & ";" & row.Convenio))

            ddlBancoRetorno.Items.Add(New ListItem(Funcoes.AlinharEsquerda(banco.Descricao, 35, ".") & " - AG: " & agencia & " - CTA: " & conta & " CONVÊNIO: " & row.Convenio,
                                 row.CodigoBanco & ";" & row.CodigoAgencia & ";" & row.DigitoAgencia & ";" & row.Conta & ";" & row.DigitoConta & ";" & row.Convenio))
        Next

        Funcoes.InserirLinhaEmBranco(ddlBanco)
        Funcoes.InserirLinhaEmBranco(ddlBancoRetorno)
    End Sub

    Public Function RetornarBoletoBancario(ByRef tit As Titulo, Optional ByVal nBanco As String = "") As String
        Dim pathPDF As String = String.Empty
        Dim nomePDF As String = String.Empty

        Dim nomearquivo As String = String.Empty
        Dim ListaDeTitulos = New List(Of Titulo)
        Dim Boletos = New Boleto2Net.Boletos()

        ListaDeTitulos.Add(tit)

        Dim where = "Where Ativo = 1 AND Empresa_Id = " & tit.CodigoEmpresaPagadora & " AND Banco_Id = " & tit.CodigoBancoPagador & " AND Agencia_Id = " & tit.CodigoAgenciaPagadora & " AND Conta_Id = " & tit.ContaPagadora
        Dim parametros As New ParametrosDaCarteiraDeCobranca(where)

        Dim parametrosCobranca As String = parametros.CodigoBanco & ";" & parametros.CodigoAgencia & ";" & parametros.DigitoAgencia & ";" & parametros.Conta & ";" & parametros.DigitoConta & ";" & parametros.Convenio

        If tit.CodigoBancoPagador = 707 Then
            Dim daycovalBoletos = BoletosDaycoval(ListaDeTitulos, parametros)
            nomePDF = NomeNoPdf(ListaDeTitulos)
            GerarBoletos(daycovalBoletos, nomePDF, False, nBanco)
        ElseIf tit.CodigoBancoPagador = 748 Then
            Dim sicrediBoletos = BoletosSicredi(ListaDeTitulos, parametros)
            nomePDF = NomeNoPdf(ListaDeTitulos)
            GerarBoletos(sicrediBoletos, nomePDF, False, nBanco)
        ElseIf tit.CodigoBancoPagador = 756 Then
            Dim sicrediBoletos = BoletosSicredi(ListaDeTitulos, parametros)
            nomePDF = NomeNoPdf(ListaDeTitulos)
            GerarBoletos(sicrediBoletos, nomePDF, False, nBanco)
        ElseIf tit.CodigoBancoPagador = 274 Then
            Dim sicrediBoletos = BoletosMoneyPlus(ListaDeTitulos, parametros)
            nomePDF = NomeNoPdf(ListaDeTitulos)
            GerarBoletos(sicrediBoletos, nomePDF, False, nBanco)
        ElseIf tit.CodigoBancoPagador = 460 Then
            Dim unavantiBoletos = BoletosUnavanti(ListaDeTitulos, parametros)
            nomePDF = NomeNoPdf(ListaDeTitulos)
            GerarBoletos(unavantiBoletos, nomePDF, False, nBanco)
        Else
            GerarBoletos(Boletos, ListaDeTitulos, parametrosCobranca)

            'Gerar Boletos PDF
            For Each b In Boletos
                Dim BoletoBancario = New Boleto2Net.BoletoBancario With {.Boleto = b}
                Dim pdf = BoletoBancario.MontaBytesPDF(False)
                nomePDF = NomeNoPdf(ListaDeTitulos)

                pathPDF = Server.MapPath("Boletos/" & nomePDF & ".pdf")
                File.WriteAllBytes(pathPDF, pdf)
            Next
        End If

        Return nomePDF

    End Function

    Private Sub GerarBoletos(ByRef Boletos As Boleto2Net.Boletos, ByRef ListaDeTitulos As List(Of Titulo), ByVal parametrosDoWhere As String)

        Dim strConta As String() = parametrosDoWhere.Split(";")
        Dim where = "Where Ativo = 1 AND Banco_Id = " & strConta(0) & " AND Agencia_Id = " & strConta(1) & " AND Conta_Id = " & strConta(3) & " AND Convenio = " & strConta(5) & " AND left(Empresa_Id,8) = " & Left(ListaDeTitulos(0).CodigoEmpresa, 8)

        Dim parametros As New ParametrosDaCarteiraDeCobranca(where)


        Dim empresa As New Cliente()
        If parametros.CodigoBanco = 237 AndAlso parametros.BeneficiarioFinal.Length > 0 Then
            empresa = New Cliente(parametros.BeneficiarioFinal, parametros.EnderecoEmpresa)
        Else
            empresa = New Cliente(parametros.CodigoEmpresa, parametros.EnderecoEmpresa)
        End If

        'Cabeçalho
        Boletos.Banco = Boleto2Net.Banco.Instancia(parametros.CodigoBanco)
        Boletos.Banco.Cedente = New Boleto2Net.Cedente With
        {
            .CPFCNPJ = parametros.CodigoEmpresa,
            .Nome = parametros.NomeDaEmpresaNoBanco,
            .ContaBancaria = New Boleto2Net.ContaBancaria With
            {
                .Agencia = parametros.CodigoAgencia,
                .DigitoAgencia = parametros.DigitoAgencia,
                .Conta = IIf(parametros.CodigoBanco = Boleto2Net.Bancos.Bradesco, parametros.Conta, parametros.Convenio),
                .DigitoConta = parametros.DigitoConta,
                .CarteiraPadrao = IIf(parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 756, parametros.CarteiraPadrao.ToArray()(0).ToString(), parametros.CarteiraPadrao),
                .TipoCarteiraPadrao = TipoCarteira.CarteiraCobrancaSimples,
                .TipoFormaCadastramento = TipoFormaCadastramento.ComRegistro,
                .TipoImpressaoBoleto = TipoImpressaoBoleto.Empresa,
                .TipoDocumento = TipoDocumento.Tradicional,
                .VariacaoCarteiraPadrao = IIf(parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 756, parametros.CarteiraPadrao.ToArray()(1).ToString(), ""),
                .OperacaoConta = IIf(parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 756, String.Format("{0:00}", Integer.Parse(parametros.SubConta)), "")
            },
            .Endereco = New Boleto2Net.Endereco With
            {
                .LogradouroEndereco = empresa.Endereco,
                .LogradouroNumero = empresa.Numero,
                .LogradouroComplemento = empresa.Complemento,
                .Bairro = empresa.Bairro,
                .Cidade = empresa.Cidade,
                .UF = empresa.Estado.Codigo,
                .CEP = empresa.CEP
             },
            .Codigo = parametros.Convenio,
            .CodigoDV = parametros.DigitoConta,
            .CodigoTransmissao = String.Empty
        }

        Boletos.Banco.FormataCedente()

        Dim cedente = New Boleto2Net.Cedente()

        For Each Titulo In ListaDeTitulos
            'Adiciona os Títulos

            Dim Boleto As New Boleto(Boletos.Banco)

            If parametros.CodigoBanco = 237 AndAlso parametros.BeneficiarioFinal.Length > 0 Then
                'Colocar um If para somente qdo for bradesco

                Boleto.Sacado = New Sacado With
                {
                    .CPFCNPJ = Titulo.Cliente.Codigo,
                    .Nome = Titulo.Cliente.Nome,
                    .Endereco = New Boleto2Net.Endereco With
                    {
                        .LogradouroEndereco = Titulo.Cliente.Endereco,
                        .LogradouroNumero = Titulo.Cliente.Numero,
                        .LogradouroComplemento = Titulo.Cliente.Complemento,
                        .Bairro = Titulo.Cliente.Bairro,
                        .Cidade = Titulo.Cliente.Cidade,
                        .UF = Titulo.Cliente.Estado.Codigo,
                        .CEP = Titulo.Cliente.CEP
                    }
                }

                'Inicio Avalista/Beneficiario final
                Boleto.Avalista = New Sacado With
                {
                    .CPFCNPJ = Titulo.Empresa.Codigo,
                    .Nome = Titulo.Empresa.Nome,
                    .Endereco = New Boleto2Net.Endereco With
                    {
                        .LogradouroEndereco = Titulo.Empresa.Endereco,
                        .LogradouroNumero = Titulo.Empresa.Numero,
                        .LogradouroComplemento = Titulo.Empresa.Complemento,
                        .Bairro = Titulo.Empresa.Bairro,
                        .Cidade = Titulo.Empresa.Cidade,
                        .UF = Titulo.Empresa.Estado.Codigo,
                        .CEP = Titulo.Empresa.CEP
                    }
                }
                'Fim avalista/benefeciario final

            Else

                Boleto.Sacado = New Sacado With
                {
                    .CPFCNPJ = Titulo.Cliente.Codigo,
                    .Nome = Titulo.Cliente.Nome,
                    .Endereco = New Boleto2Net.Endereco With
                    {
                        .LogradouroEndereco = Titulo.Cliente.Endereco,
                        .LogradouroNumero = Titulo.Cliente.Numero,
                        .LogradouroComplemento = Titulo.Cliente.Complemento,
                        .Bairro = Titulo.Cliente.Bairro,
                        .Cidade = Titulo.Cliente.Cidade,
                        .UF = Titulo.Cliente.Estado.Codigo,
                        .CEP = Titulo.Cliente.CEP
                    }
                }
            End If

            Boleto.CodigoOcorrencia = "01"
            Boleto.DescricaoOcorrencia = "Remessa Registrar"

            Boleto.NumeroDocumento = Titulo.Codigo.ToString()
            Boleto.NumeroControleParticipante = Titulo.Codigo.ToString()
            Boleto.NossoNumero = IIf(parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 756, String.Format("{0:00000}", (parametros.SequenciaDeRegistroDePagamento + 1)), Titulo.Codigo.ToString())

            Boleto.DataEmissao = DateTime.Now.Date

            'Furlan - 26/07/2021 - Alterado para prorrogação pois podem ter mudado a data antes de gerar os Boletos
            'Boleto.DataVencimento = Titulo.Vencimento
            Boleto.DataVencimento = Titulo.Prorrogacao

            Boleto.ValorTitulo = Titulo.ValorLiquido
            Boleto.Aceite = "N"
            Boleto.EspecieDocumento = IIf(parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 756, TipoEspecieDocumento.DMI, TipoEspecieDocumento.DM)

            'Boleto.CodigoInstrucao1 = IIf(parametros.CodigoBanco = 104, "91", "05")

            Boleto.MensagemInstrucoesCaixa = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."

            If (parametros.JurosMes > 0) Then
                Boleto.DataMulta = Titulo.Prorrogacao.AddDays(1)
                Boleto.PercentualMulta = parametros.JurosMes
                Boleto.ValorMulta = CDec(Boleto.ValorTitulo * Boleto.PercentualMulta / 100)

                Dim Instrucao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & Boleto.ValorMulta.ToString("C2", CultureInfo.CurrentCulture) & "."

                If (String.IsNullOrEmpty(Boleto.MensagemInstrucoesCaixa)) Then
                    Boleto.MensagemInstrucoesCaixa = Instrucao
                Else
                    Boleto.MensagemInstrucoesCaixa += Environment.NewLine + Instrucao
                End If
            End If

            If (parametros.MoraDiaria > 0) Then
                Boleto.DataJuros = Titulo.Prorrogacao.AddDays(1)

                'Furlan - 26-07-2021 - Química é diferente pois cobra 3,50% ao mês dividido por 30.  
                If parametros.CodigoEmpresa = "05272759000147" Then
                    Boleto.PercentualJurosDia = (parametros.MoraDiaria / 30)
                    Boleto.ValorJurosDia = Boleto.ValorTitulo * Boleto.PercentualJurosDia / 100
                Else
                    Boleto.PercentualJurosDia = parametros.MoraDiaria
                    Boleto.ValorJurosDia = (Boleto.ValorTitulo / 30) * Boleto.PercentualJurosDia / 100
                End If

                Dim Instrucao = "<br /> APOS VENCIMENTO COBRAR R$ " & Boleto.ValorJurosDia.ToString("C2", CultureInfo.CurrentCulture) & " POR DIA DE ATRASO."

                If (String.IsNullOrEmpty(Boleto.MensagemInstrucoesCaixa)) Then
                    Boleto.MensagemInstrucoesCaixa = Instrucao
                Else
                    Boleto.MensagemInstrucoesCaixa += Environment.NewLine + Instrucao
                End If
            End If

            'Protesto boleto
            Boleto.CodigoProtesto = IIf(parametros.DiasParaProtesto > 0, TipoCodigoProtesto.ProtestarDiasUteis, TipoCodigoProtesto.NaoProtestar)
            If parametros.DiasParaProtesto > 0 Then
                Boleto.CodigoInstrucao1 = "06"
                Boleto.CodigoInstrucao2 = parametros.DiasParaProtesto.ToString()
                Boleto.DiasProtesto = parametros.DiasParaProtesto

                If Left(parametros.CodigoEmpresa, 8) = "40938762" OrElse Left(parametros.CodigoEmpresa, 8) = "49673784" OrElse Left(parametros.CodigoEmpresa, 8) = "05366261" OrElse Left(parametros.CodigoEmpresa, 8) = "62747840" OrElse Left(parametros.CodigoEmpresa, 8) = "62780383" Then
                    Boleto.MensagemInstrucoesCaixa += Environment.NewLine + "<br /> PROTESTAR APÓS " + parametros.DiasParaProtesto.ToString() + " DIAS DO VENCIMENTO."
                Else
                    Boleto.MensagemInstrucoesCaixa += Environment.NewLine + "<br /> PROTESTAR APÓS " + parametros.DiasParaProtesto.ToString() + " DIAS ÚTEIS."
                End If
            End If

            Boleto.CodigoBaixaDevolucao = IIf(parametros.CodigoBanco = 104, TipoCodigoBaixaDevolucao.BaixarDevolver, TipoCodigoBaixaDevolucao.NaoBaixarNaoDevolver)
            'Boleto.DiasBaixaDevolucao = 0

            Boleto.ValidarDados()
            Boletos.Add(Boleto)

        Next
    End Sub

#Region "Daycoval"

    Protected Function BoletosDaycoval(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca) As List(Of BoletoNet.Boleto)
        Dim boletos As New List(Of BoletoNet.Boleto)
        For Each tit In pTitulos
            Dim b As New BoletoNet.Boleto
            b.Banco = New BoletoNet.Banco(param.CodigoBanco)
            b.Cedente = New BoletoNet.Cedente With
            {
                .CPFCNPJ = param.CodigoEmpresa,
                .Nome = param.NomeDaEmpresaNoBanco,
                .ContaBancaria = New BoletoNet.ContaBancaria With
                {
                    .Agencia = param.CodigoAgencia,
                    .DigitoAgencia = param.DigitoAgencia,
                    .Conta = param.Conta,
                    .DigitoConta = param.DigitoConta
                },
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Empresa.Endereco,
                    .Numero = tit.Empresa.Numero,
                    .Complemento = tit.Empresa.Complemento,
                    .Bairro = tit.Empresa.Bairro,
                    .Cidade = tit.Empresa.Cidade,
                    .UF = tit.Empresa.Estado.Codigo,
                    .CEP = tit.Empresa.CEP
                 },
                .Codigo = param.CodigoDaEmpresaNoBanco,
                .Convenio = param.Convenio
            }
            b.Sacado = New BoletoNet.Sacado With
            {
                .CPFCNPJ = tit.Cliente.Codigo,
                .Nome = tit.Cliente.Nome,
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Cliente.Endereco,
                    .Numero = tit.Cliente.Numero,
                    .Complemento = tit.Cliente.Complemento,
                    .Bairro = tit.Cliente.Bairro,
                    .Cidade = tit.Cliente.Cidade,
                    .UF = tit.Cliente.Estado.Codigo,
                    .CEP = tit.Cliente.CEP
                }
            }

            b.NumeroDocumento = tit.Codigo.ToString()
            b.NumeroControle = param.SubConta
            b.NossoNumero = param.SequenciaDePagamentos
            b.DataDocumento = DateTime.Now.Date
            b.DataVencimento = tit.Prorrogacao
            b.ValorBoleto = tit.ValorLiquido
            b.Aceite = "N"
            b.EspecieDocumento = New BoletoNet.EspecieDocumento(param.CodigoBanco, "12")
            b.Carteira = param.CarteiraPadrao

            Dim instrucao1 = New BoletoNet.Instrucao(param.CodigoBanco)
            instrucao1.Descricao = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."
            b.Instrucoes.Add(instrucao1)

            If (param.JurosMes > 0) Then
                b.DataMulta = tit.Prorrogacao.AddDays(1)
                b.PercMulta = param.JurosMes
                b.ValorMulta = CDec(b.ValorBoleto * b.PercMulta / 100)

                Dim instrucao2 = New BoletoNet.Instrucao(param.CodigoBanco)
                instrucao2.Descricao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & b.ValorMulta.ToString("C2", CultureInfo.CurrentCulture) & "."
                b.Instrucoes.Add(instrucao2)

            End If

            boletos.Add(b)
            param.SequenciaDePagamentos += 1
        Next
        Return boletos
    End Function

    Protected Sub GerarRemessaDaycoval(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca)
        Dim arDaycoval = New NGS.Lib.Negocio.BoletoBancario()
        Dim boletos = BoletosDaycoval(pTitulos, param)

        Dim sb = arDaycoval.DaycovalArquivoCNAB400(boletos)

        If Not Directory.Exists(Server.MapPath("~/RemessaBancaria/" & param.Convenio.ToString)) Then Directory.CreateDirectory(Server.MapPath("~/RemessaBancaria/" & param.Convenio.ToString))

        Dim seq = 1
        Dim arr = param.NomeNoArquivo.Replace("4CK", "").ToArray()
        Dim dia = (arr(0) + arr(1))
        Dim mes = (arr(2) + arr(3))
        If (DateTime.Today = New DateTime(DateTime.Now.Year, Convert.ToInt32(mes), Convert.ToInt32(dia))) Then
            seq += 1
        End If

        'Gerar Arquivo de Remessa
        Dim NomeArq = "4CK" + DateTime.Now.ToString("ddMM") + seq.ToString()
        param.NomeNoArquivo = NomeArq
        Dim Path As String = "RemessaBancaria/" & param.Convenio.ToString & "/4CK" + DateTime.Now.ToString("ddMM") + seq.ToString() + ".txt"

        Dim bytes = Encoding.ASCII.GetBytes(sb.ToString().ToCharArray())

        'File.WriteAllBytes(Server.MapPath(NomeArquivo2), bytes)

        Dim strm = New StreamWriter(Server.MapPath(Path))

        For Each line In sb
            strm.WriteLine(line)
        Next

        strm.Dispose()
        strm.Close()

        MostrarBotaoDownload()

        'Gerar Pdf
        'DownloadBoletoDaycoval(boletos)
    End Sub

    Protected Sub GerarBoletos(ByRef boletos As List(Of BoletoNet.Boleto), ByVal nomePDF As String, ByVal baixar As Boolean, Optional ByVal nBanco As String = "")
        Dim html = New StringBuilder()
        Dim remessa As New NGS.Lib.Negocio.BoletoBancario
        Dim codigoBanco = 0
        Dim strConta As String()

        If nBanco.Length > 0 Then
            strConta = nBanco.Split(";")
        Else
            strConta = IIf(String.IsNullOrEmpty(ddlBanco.SelectedValue), ddlBancoRetorno.SelectedValue.Split(";"), ddlBanco.SelectedValue.Split(";"))
        End If

        'Gerar Boletos HTML
        For Each b In boletos

            html.Append("<div style=""page-break-after: always;"">")
            Dim htmlAux As String = ""

            If b.Banco.Codigo = 748 Then
                htmlAux = remessa.SicrediPdf(b)
                codigoBanco = 748
            ElseIf b.Banco.Codigo = 756 Then
                htmlAux = remessa.SicoobPdf(b)
                codigoBanco = 756
            ElseIf b.Banco.Codigo = 237 And strConta(0) = 274 Then 'Money Plus us o codigo do Bradesco
                htmlAux = remessa.MoneyPlusPdf(b)
                codigoBanco = 274
            ElseIf b.Banco.Codigo = 237 And strConta(0) = 460 Then 'Unavanti us o codigo do Bradesco
                htmlAux = remessa.UnavantiPdf(b)
                codigoBanco = 460
            Else
                htmlAux = remessa.DaycovalPdf(b)
                codigoBanco = 707
            End If

            htmlAux = htmlAux.Replace("@URLIMAGEMLOGO", Server.MapPath("Images/" + codigoBanco.ToString() + ".png"))
            htmlAux = htmlAux.Replace("@URLIMGCEDENTE", Server.MapPath("Images/" + codigoBanco.ToString() + ".png"))
            html.Append(htmlAux)
            html.Append("</div>")
        Next

        Dim pathPDF = Server.MapPath("Boletos/" + nomePDF + ".pdf")

        Dim pdf = New NReco.PdfGenerator.HtmlToPdfConverter().GeneratePdf(html.ToString)

        Using fs As FileStream = New FileStream(pathPDF, FileMode.Create)
            fs.Write(pdf, 0, pdf.Length)
            fs.Close()
        End Using

        If baixar Then Download("Boletos", nomePDF & ".pdf", "application/pdf")
    End Sub

    Protected Sub GerarBoletoSemDownload(ByRef boletos As List(Of BoletoNet.Boleto), ByVal nomePDF As String)
        Dim html = New StringBuilder()
        Dim remesaDaycoval As New NGS.Lib.Negocio.BoletoBancario
        'Gerar Boletos HTML
        For Each b In boletos

            html.Append("<div style=""page-break-after: always;"">")
            Dim htmlAux As String = ""

            If b.Banco.Codigo = 748 Or b.Banco.Codigo = 756 Then
                htmlAux = remesaDaycoval.SicrediPdf(b)
            Else
                htmlAux = remesaDaycoval.DaycovalPdf(b)
            End If

            htmlAux = htmlAux.Replace("@URLIMAGEMLOGO", Server.MapPath("Images/" + b.Banco.Codigo.ToString() + ".png"))
            htmlAux = htmlAux.Replace("@URLIMGCEDENTE", Server.MapPath("Images/" + b.Banco.Codigo.ToString() + ".png"))
            html.Append(htmlAux)
            html.Append("</div>")
        Next

        Dim pathPDF = Server.MapPath("Boletos/" + nomePDF + ".pdf")

        Dim pdf = New NReco.PdfGenerator.HtmlToPdfConverter().GeneratePdf(html.ToString)

        Using fs As FileStream = New FileStream(pathPDF, FileMode.Create)
            fs.Write(pdf, 0, pdf.Length)
            fs.Close()
        End Using
    End Sub

#End Region

#Region "Sicredi"
    Protected Function BoletosSicredi(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca) As List(Of BoletoNet.Boleto)
        Dim boletos As New List(Of BoletoNet.Boleto)
        For Each tit In pTitulos
            Dim b As New BoletoNet.Boleto
            b.Banco = New BoletoNet.Banco(param.CodigoBanco)
            b.Cedente = New BoletoNet.Cedente With
            {
                .CPFCNPJ = param.CodigoEmpresa,
                .Nome = param.NomeDaEmpresaNoBanco,
                .ContaBancaria = New BoletoNet.ContaBancaria With
                {
                    .Agencia = param.CodigoAgencia,
                    .DigitoAgencia = param.DigitoAgencia,
                    .Conta = param.Conta,
                    .DigitoConta = param.DigitoConta,
                    .OperacaConta = param.SubConta
                },
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Empresa.Endereco,
                    .Numero = tit.Empresa.Numero,
                    .Complemento = tit.Empresa.Complemento,
                    .Bairro = tit.Empresa.Bairro,
                    .Cidade = tit.Empresa.Cidade,
                    .UF = tit.Empresa.Estado.Codigo,
                    .CEP = tit.Empresa.CEP
                 },
                .Codigo = param.CodigoDaEmpresaNoBanco,
                .Convenio = param.Convenio,
                .CodigoTransmissao = param.SequenciaDeRemessa
            }
            b.Sacado = New BoletoNet.Sacado With
            {
                .CPFCNPJ = tit.Cliente.Codigo,
                .Nome = tit.Cliente.Nome,
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Cliente.Endereco,
                    .Numero = tit.Cliente.Numero,
                    .Complemento = tit.Cliente.Complemento,
                    .Bairro = tit.Cliente.Bairro,
                    .Cidade = tit.Cliente.Cidade,
                    .UF = tit.Cliente.Estado.Codigo,
                    .CEP = tit.Cliente.CEP
                }
            }

            b.NumeroDocumento = tit.Codigo.ToString()
            b.NossoNumero = IIf(tit.NossoNumero = "0", param.SequenciaDePagamentos, tit.NossoNumero)
            b.DataDocumento = tit.UsuarioInclusaoData
            b.DataVencimento = tit.Prorrogacao
            b.ValorBoleto = tit.ValorLiquido
            b.Aceite = "N"
            b.Carteira = param.CarteiraPadrao

            Dim instrucao1 = New BoletoNet.Instrucao(param.CodigoBanco)
            instrucao1.Descricao = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."
            b.Instrucoes.Add(instrucao1)

            If (param.JurosMes > 0) Then
                b.DataMulta = tit.Prorrogacao.AddDays(1)
                b.PercMulta = param.JurosMes
                b.ValorMulta = CDec(b.ValorBoleto * b.PercMulta / 100)

                Dim instrucao2 = New BoletoNet.Instrucao(param.CodigoBanco)
                instrucao2.Descricao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & b.ValorMulta.ToString("C2", CultureInfo.CurrentCulture) & "."
                b.Instrucoes.Add(instrucao2)

            End If
            'Mora dia
            If (param.MoraDiaria > 0) Then
                b.JurosMora = (b.ValorBoleto * param.MoraDiaria) / 100
                b.PercJurosMora = param.MoraDiaria
                Dim instrucao3 = New BoletoNet.Instrucao(param.CodigoBanco)
                instrucao3.Descricao = "<br /> APOS VENCIMENTO COBRAR R$ " & b.JurosMora.ToString("C2", CultureInfo.CurrentCulture) & " POR DIA DE ATRASO"
                b.Instrucoes.Add(instrucao3)

            End If
            'Precisa gravar no titulo o nosso numero para gerar o pdf
            tit.NossoNumero = param.SequenciaDePagamentos
            boletos.Add(b)
            param.SequenciaDePagamentos += 1
        Next
        Return boletos
    End Function

    Protected Sub GerarRemessaSicredi(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca)
        Dim sicredi = New NGS.Lib.Negocio.BoletoBancario()
        Dim boletos = BoletosSicredi(pTitulos, param)

        Dim sb = sicredi.SicrediArquivoCNAB400(boletos)

        If Not Directory.Exists(Server.MapPath("~/RemessaBancaria/" & param.Convenio.ToString)) Then Directory.CreateDirectory(Server.MapPath("~/RemessaBancaria/" & param.Convenio.ToString))

        'Codigo Beneficiario
        Dim nomeArq = Integer.Parse(param.CodigoDaEmpresaNoBanco).ToString("00000")
        'Codigo do Mês e Dia
        Dim mes = DateTime.Now.Month
        Dim dia = DateTime.Now.Day
        'Select Case mes
        '    Case mes <= 9
        '        nomeArq += mes.ToString()
        '    Case 10
        '        nomeArq += ("O" + dia.ToString("00"))
        '    Case 11
        '        nomeArq += ("N" + dia.ToString("00"))
        '    Case 12
        '        nomeArq += ("D" + dia.ToString("00"))
        'End Select

        If (mes <= 9) Then
            nomeArq += mes.ToString()
        ElseIf (mes = 10) Then
            nomeArq += "O"
        ElseIf (mes = 11) Then
            nomeArq += "N"
        ElseIf (mes = 12) Then
            nomeArq += "D"
        End If

        nomeArq += dia.ToString("00")

        'Sequencia de registro de remessa na extensão do arquivo
        'nomeArq += ("." + (param.SequenciaDeRemessa + 1).ToString("0000"))
        'Gerar Arquivo de Remessa
        param.NomeNoArquivo = nomeArq
        Dim path As String = "RemessaBancaria/" & param.Convenio.ToString & "/" + nomeArq + ".txt"

        Dim bytes = Encoding.ASCII.GetBytes(sb.ToString().ToCharArray())

        'File.WriteAllBytes(Server.MapPath(NomeArquivo2), bytes)

        Dim strm = New StreamWriter(Server.MapPath(path))

        For Each line In sb
            strm.WriteLine(line)
        Next

        strm.Dispose()
        strm.Close()

        MostrarBotaoDownload()

        'Gerar Pdf
        'DownloadBoletoDaycoval(boletos)
    End Sub
#End Region

#Region "Sicoob"
    Protected Function BoletosSicoob(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca) As List(Of BoletoNet.Boleto)
        Dim boletos As New List(Of BoletoNet.Boleto)
        Dim i As Integer = 0
        For Each tit In pTitulos
            Dim b As New BoletoNet.Boleto
            b.Banco = New BoletoNet.Banco(param.CodigoBanco)
            b.Cedente = New BoletoNet.Cedente With
            {
                .CPFCNPJ = param.CodigoEmpresa,
                .Nome = param.NomeDaEmpresaNoBanco,
                .ContaBancaria = New BoletoNet.ContaBancaria With
                {
                    .Agencia = param.CodigoAgencia,
                    .DigitoAgencia = param.DigitoAgencia,
                    .Conta = param.Conta,
                    .DigitoConta = param.DigitoConta,
                    .OperacaConta = param.SubConta
                },
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Empresa.Endereco,
                    .Numero = tit.Empresa.Numero,
                    .Complemento = tit.Empresa.Complemento,
                    .Bairro = tit.Empresa.Bairro,
                    .Cidade = tit.Empresa.Cidade,
                    .UF = tit.Empresa.Estado.Codigo,
                    .CEP = tit.Empresa.CEP
                 },
                .Codigo = param.CodigoDaEmpresaNoBanco,
                .Convenio = param.Convenio,
                .CodigoTransmissao = param.SequenciaDeRemessa
            }
            b.Sacado = New BoletoNet.Sacado With
            {
                .CPFCNPJ = tit.Cliente.Codigo,
                .Nome = tit.Cliente.Nome,
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Cliente.Endereco,
                    .Numero = tit.Cliente.Numero,
                    .Complemento = tit.Cliente.Complemento,
                    .Bairro = tit.Cliente.Bairro,
                    .Cidade = tit.Cliente.Cidade,
                    .UF = tit.Cliente.Estado.Codigo,
                    .CEP = tit.Cliente.CEP
                }
            }

            b.NumeroDocumento = tit.Codigo.ToString()

            If (tit.NossoNumero <> "0" AndAlso (Not String.IsNullOrWhiteSpace(tit.NossoNumero))) Then
                b.NossoNumero = tit.NossoNumero
            Else
                b.NossoNumero = (param.SequenciaDePagamentos + i)
                i = (i + 1)
            End If

            b.DataDocumento = tit.UsuarioInclusaoData
            b.DataVencimento = tit.Prorrogacao
            b.ValorBoleto = tit.ValorLiquido
            b.Aceite = "N"
            b.Carteira = param.CarteiraPadrao

            Dim instrucao1 = New BoletoNet.Instrucao(param.CodigoBanco)
            instrucao1.Descricao = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."
            b.Instrucoes.Add(instrucao1)

            If (param.JurosMes > 0) Then
                b.DataMulta = tit.Prorrogacao
                b.PercMulta = param.JurosMes
                b.ValorMulta = Math.Round(CDec(b.ValorBoleto * b.PercMulta / 100), 2)

                Dim instrucao2 = New BoletoNet.Instrucao(param.CodigoBanco)
                instrucao2.Descricao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & b.ValorMulta.ToString("C2", CultureInfo.CurrentCulture) & "."
                b.Instrucoes.Add(instrucao2)

            End If
            'Mora dia
            If (param.MoraDiaria > 0) Then
                'Furlan - 26-07-2021 - Química é diferente pois cobra 3,50% ao mês dividido por 30.  
                If param.CodigoEmpresa = "05272759000147" Then
                    b.PercJurosMora = Math.Round((param.MoraDiaria / 30), 2)
                    b.JurosMora = Math.Round(b.ValorBoleto * b.PercJurosMora / 100, 2)
                Else
                    b.PercJurosMora = param.MoraDiaria
                    b.JurosMora = (b.ValorBoleto / 30) * b.PercJurosMora / 100
                End If
                b.DataJurosMora = tit.Prorrogacao
                Dim instrucao3 = New BoletoNet.Instrucao(param.CodigoBanco)
                instrucao3.Descricao = "<br /> APOS VENCIMENTO COBRAR R$ " & b.JurosMora.ToString("C2", CultureInfo.CurrentCulture) & " POR DIA DE ATRASO"
                b.Instrucoes.Add(instrucao3)

            End If
            'Precisa gravar no titulo o nosso numero para gerar o pdf
            'tit.NossoNumero = param.SequenciaDePagamentos
            boletos.Add(b)
        Next
        Return boletos
    End Function

    Protected Sub GerarRemessaSicoob(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca)
        Dim sicoob = New NGS.Lib.Negocio.BoletoBancario()
        Dim boletos = BoletosSicoob(pTitulos, param)

        Dim sb = sicoob.SicoobArquivoCNAB240(boletos)

        If Not Directory.Exists(Server.MapPath("~/RemessaBancaria/" & param.Convenio.ToString)) Then Directory.CreateDirectory(Server.MapPath("~/RemessaBancaria/" & param.Convenio.ToString))

        'Codigo Beneficiario
        Dim nomeArq = Integer.Parse(param.CodigoDaEmpresaNoBanco).ToString("00000")
        'Codigo do Mês e Dia
        Dim mes = DateTime.Now.Month
        Dim dia = DateTime.Now.Day

        If (mes <= 9) Then
            nomeArq += mes.ToString()
        ElseIf (mes = 10) Then
            nomeArq += "O"
        ElseIf (mes = 11) Then
            nomeArq += "N"
        ElseIf (mes = 12) Then
            nomeArq += "D"
        End If

        nomeArq += dia.ToString("00")

        param.NomeNoArquivo = nomeArq
        Dim path As String = "RemessaBancaria/" & param.Convenio.ToString & "/" + nomeArq + ".txt"

        Dim bytes = Encoding.ASCII.GetBytes(sb.ToString().ToCharArray())

        Dim strm = New StreamWriter(Server.MapPath(path))

        For Each line In sb
            strm.WriteLine(line)
        Next

        strm.Dispose()
        strm.Close()

        MostrarBotaoDownload()
    End Sub
#End Region

#Region "Banco Plus"
    Protected Function BoletosMoneyPlus(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca) As List(Of BoletoNet.Boleto)
        Dim boletos As New List(Of BoletoNet.Boleto)
        For Each tit In pTitulos
            Dim b As New BoletoNet.Boleto
            b.Banco = New BoletoNet.Banco(237) 'Money Plus usa o codigo do Bradesco
            b.Cedente = New BoletoNet.Cedente With
            {
                .CPFCNPJ = param.CodigoEmpresa,
                .Nome = param.NomeDaEmpresaNoBanco,
                .ContaBancaria = New BoletoNet.ContaBancaria With
                {
                    .Agencia = param.CodigoAgencia,
                    .DigitoAgencia = param.DigitoAgencia,
                    .Conta = param.Conta,
                    .DigitoConta = param.DigitoConta,
                    .OperacaConta = param.SubConta
                },
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Empresa.Endereco,
                    .Numero = tit.Empresa.Numero,
                    .Complemento = tit.Empresa.Complemento,
                    .Bairro = tit.Empresa.Bairro,
                    .Cidade = tit.Empresa.Cidade,
                    .UF = tit.Empresa.Estado.Codigo,
                    .CEP = tit.Empresa.CEP
                 },
                .Codigo = param.CodigoDaEmpresaNoBanco,
                .Convenio = param.Convenio,
                .CodigoTransmissao = param.SequenciaDeRemessa
            }
            b.Sacado = New BoletoNet.Sacado With
            {
                .CPFCNPJ = tit.Cliente.Codigo,
                .Nome = tit.Cliente.Nome,
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Cliente.Endereco,
                    .Numero = tit.Cliente.Numero,
                    .Complemento = tit.Cliente.Complemento,
                    .Bairro = tit.Cliente.Bairro,
                    .Cidade = tit.Cliente.Cidade,
                    .UF = tit.Cliente.Estado.Codigo,
                    .CEP = tit.Cliente.CEP
                }
            }

            b.NumeroDocumento = tit.Codigo.ToString()
            b.NossoNumero = IIf(tit.NossoNumero = "0", param.SequenciaDePagamentos, tit.NossoNumero)
            b.DataDocumento = tit.UsuarioInclusaoData
            b.DataVencimento = tit.Prorrogacao
            b.ValorBoleto = tit.ValorLiquido
            b.Aceite = "N"
            b.Carteira = param.CarteiraPadrao

            Dim instrucao1 = New BoletoNet.Instrucao(237) 'Usa igual ao Bradesco
            instrucao1.Descricao = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."
            b.Instrucoes.Add(instrucao1)

            If (param.JurosMes > 0) Then
                b.DataMulta = tit.Prorrogacao.AddDays(1)
                b.PercMulta = param.JurosMes
                b.ValorMulta = CDec(b.ValorBoleto * b.PercMulta / 100)

                Dim instrucao2 = New BoletoNet.Instrucao(237) 'Usa igual ao Bradesco
                instrucao2.Descricao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & b.ValorMulta.ToString("C2", CultureInfo.CurrentCulture) & "."
                b.Instrucoes.Add(instrucao2)

            End If
            'Mora dia
            If (param.MoraDiaria > 0) Then
                b.PercJurosMora = param.MoraDiaria
                b.JurosMora = (b.ValorBoleto / 30) * b.PercJurosMora / 100
                Dim instrucao3 = New BoletoNet.Instrucao(237) 'Usa igual ao Bradesco
                instrucao3.Descricao = "<br /> APOS VENCIMENTO COBRAR R$ " & b.JurosMora.ToString("C2", CultureInfo.CurrentCulture) & " POR DIA DE ATRASO"
                b.Instrucoes.Add(instrucao3)

            End If
            'Precisa gravar no titulo o nosso numero para gerar o pdf
            tit.NossoNumero = param.SequenciaDePagamentos
            boletos.Add(b)
            param.SequenciaDePagamentos += 1
        Next
        Return boletos
    End Function

    Protected Sub GerarRemessaMoneyPlus(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca)
        Dim moneyPlus = New NGS.Lib.Negocio.BoletoBancario()
        Dim boletos = BoletosMoneyPlus(pTitulos, param)

        Dim sb = moneyPlus.MoneyPlusCNAB400(boletos)

        'CBDDMMSSSSSSS.REM
        'CB -Cobrança de Boleto
        'DD – O Dia geração Do arquivo
        'MM – O Mês da geração Do Arquivo
        'SSSSSSS – Sequencial
        'Numérica Ex.:  0000001,0000002, 0000003, - .
        '.Rem – Extensão do arquivo.
        'Exemplo: CB01050000001.REM

        If Not Directory.Exists(Server.MapPath("~/RemessaBancaria/" & param.Convenio.ToString)) Then Directory.CreateDirectory(Server.MapPath("~/RemessaBancaria/" & param.Convenio.ToString))

        'Codigo do Mês e Dia
        Dim dia = DateTime.Now.Day
        Dim mes = DateTime.Now.Month

        'Sequencia de registro de remessa na extensão do arquivo
        'Gerar Arquivo de Remessa
        Dim nomeArq = String.Format("CB{0}{1}{2}", dia, mes, Integer.Parse(param.SequenciaDeRemessa).ToString("0000000"))
        param.NomeNoArquivo = nomeArq
        Dim path As String = "RemessaBancaria/" & param.Convenio.ToString & "/" + nomeArq + ".REM"

        Dim bytes = Encoding.ASCII.GetBytes(sb.ToString().ToCharArray())

        Dim strm = New StreamWriter(Server.MapPath(path))

        For Each line In sb
            strm.WriteLine(line)
        Next

        strm.Dispose()
        strm.Close()

        MostrarBotaoDownload()

        'Gerar Pdf
        'DownloadBoletoDaycoval(boletos)
    End Sub
#End Region

#Region "Unavanti"
    Protected Function BoletosUnavanti(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca) As List(Of BoletoNet.Boleto)
        Dim boletos As New List(Of BoletoNet.Boleto)
        For Each tit In pTitulos
            Dim b As New BoletoNet.Boleto
            b.Banco = New BoletoNet.Banco(237) 'Unavanti usa o codigo do Bradesco
            b.Cedente = New BoletoNet.Cedente With
            {
                .CPFCNPJ = param.CodigoEmpresa,
                .Nome = param.NomeDaEmpresaNoBanco,
                .ContaBancaria = New BoletoNet.ContaBancaria With
                {
                    .Agencia = param.CodigoAgencia,
                    .DigitoAgencia = param.DigitoAgencia,
                    .Conta = param.Conta,
                    .DigitoConta = param.DigitoConta,
                    .OperacaConta = param.SubConta
                },
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Empresa.Endereco,
                    .Numero = tit.Empresa.Numero,
                    .Complemento = tit.Empresa.Complemento,
                    .Bairro = tit.Empresa.Bairro,
                    .Cidade = tit.Empresa.Cidade,
                    .UF = tit.Empresa.Estado.Codigo,
                    .CEP = tit.Empresa.CEP
                 },
                .Codigo = param.CodigoDaEmpresaNoBanco,
                .Convenio = param.Convenio,
                .CodigoTransmissao = param.SequenciaDeRemessa
            }
            b.Sacado = New BoletoNet.Sacado With
            {
                .CPFCNPJ = tit.Cliente.Codigo,
                .Nome = tit.Cliente.Nome,
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Cliente.Endereco,
                    .Numero = tit.Cliente.Numero,
                    .Complemento = tit.Cliente.Complemento,
                    .Bairro = tit.Cliente.Bairro,
                    .Cidade = tit.Cliente.Cidade,
                    .UF = tit.Cliente.Estado.Codigo,
                    .CEP = tit.Cliente.CEP
                }
            }

            b.NumeroDocumento = tit.Codigo.ToString()
            b.NossoNumero = IIf(tit.NossoNumero = "0", param.SequenciaDePagamentos, tit.NossoNumero)
            b.DataDocumento = tit.UsuarioInclusaoData
            b.DataVencimento = tit.Prorrogacao
            b.ValorBoleto = tit.ValorLiquido
            b.Aceite = "N"
            b.Carteira = param.CarteiraPadrao

            Dim instrucao1 = New BoletoNet.Instrucao(237) 'Usa igual ao Bradesco
            instrucao1.Descricao = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."
            b.Instrucoes.Add(instrucao1)

            If (param.JurosMes > 0) Then
                b.DataMulta = tit.Prorrogacao.AddDays(1)
                b.PercMulta = param.JurosMes
                b.ValorMulta = CDec(b.ValorBoleto * b.PercMulta / 100)

                Dim instrucao2 = New BoletoNet.Instrucao(237) 'Usa igual ao Bradesco
                instrucao2.Descricao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & b.ValorMulta.ToString("C2", CultureInfo.CurrentCulture) & "."
                b.Instrucoes.Add(instrucao2)

            End If
            'Mora dia
            If (param.MoraDiaria > 0) Then
                b.PercJurosMora = param.MoraDiaria
                b.JurosMora = (b.ValorBoleto / 30) * b.PercJurosMora / 100
                Dim instrucao3 = New BoletoNet.Instrucao(237) 'Usa igual ao Bradesco
                instrucao3.Descricao = "<br /> APOS VENCIMENTO COBRAR R$ " & b.JurosMora.ToString("C2", CultureInfo.CurrentCulture) & " POR DIA DE ATRASO"
                b.Instrucoes.Add(instrucao3)

            End If

            Dim instrucao4 = New BoletoNet.Instrucao(237) 'Usa igual ao Bradesco
            instrucao4.Descricao = "<br /> PROTESTAR APÓS " + param.DiasParaProtesto.ToString() + " DIAS DO VENCIMENTO."
            b.Instrucoes.Add(instrucao4)

            Dim instrucao5 = New BoletoNet.Instrucao(237) 'Usa igual ao Bradesco
            'Número da NF
            Dim titXNF = New NotaFiscalXTitulo(New TituloV With {.Codigo = tit.Codigo})
            instrucao5.Descricao = "<br /> NOTA FISCAL: " + titXNF.NotaFiscal.Codigo.ToString()
            b.Instrucoes.Add(instrucao5)

            'Precisa gravar no titulo o nosso numero para gerar o pdf
            tit.NossoNumero = param.SequenciaDePagamentos
            boletos.Add(b)
            param.SequenciaDePagamentos += 1
        Next
        Return boletos
    End Function

    Protected Sub GerarRemessaUnavanti(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca)
        Dim unavanti = New NGS.Lib.Negocio.BoletoBancario()
        Dim boletos = BoletosUnavanti(pTitulos, param)

        Dim sb = unavanti.UnavantiCNAB400(boletos)

        'CBDDMMSSSSSSS.REM
        'CB -Cobrança de Boleto
        'DD – O Dia geração Do arquivo
        'MM – O Mês da geração Do Arquivo
        'SS – Sequencial
        'Numérica Ex.:  0000001,0000002, 0000003, - .
        '.Rem – Extensão do arquivo.
        'Exemplo: CB01050000001.REM

        If Not Directory.Exists(Server.MapPath("~/RemessaBancaria/" & param.Convenio.ToString)) Then Directory.CreateDirectory(Server.MapPath("~/RemessaBancaria/" & param.Convenio.ToString))

        'Codigo do Mês e Dia
        Dim diaMes = String.Format("{0}{1}", DateTime.Now.Day.ToString("00"), DateTime.Now.Month.ToString("00"))

        'Sequencia de registro de remessa na extensão do arquivo
        If (diaMes <> param.NomeNoArquivo.Substring(2, 4)) Then
            param.SequenciaDeRemessa = 1
        End If

        'Gerar Arquivo de Remessa
        Dim nomeArq = String.Format("CB{0}{1}", diaMes, Integer.Parse(param.SequenciaDeRemessa).ToString("00"))
        param.NomeNoArquivo = nomeArq
        Dim path As String = "RemessaBancaria/" & param.Convenio.ToString & "/" + nomeArq + ".REM"

        Dim bytes = Encoding.ASCII.GetBytes(sb.ToString().ToCharArray())

        Dim strm = New StreamWriter(Server.MapPath(path))

        For Each line In sb
            strm.WriteLine(line)
        Next

        strm.Dispose()
        strm.Close()

        MostrarBotaoDownload()

        'Gerar Pdf
        'DownloadBoletoDaycoval(boletos)
    End Sub
#End Region

#Region "Metodos Auxiliares do Boleto"
    Public Sub MostrarBotaoDownload()
        lnkGerarRemessa.Enabled = False
        lnkAgruparPagamento.Enabled = False
        ddlBanco.Enabled = False

        gridBoletoNotaFiscal.DataSource = Nothing
        gridBoletoNotaFiscal.DataBind()

        btnDownload.Visible = True
        btnDownload.Text = "Download do arquivo bancário com os boletos gerados"
    End Sub

    Public Function AtualizarTitulos(ByRef ListaDeTitulos As List(Of Titulo), ByVal parametros As ParametrosDaCarteiraDeCobranca, ByVal where As String) As Boolean

        Try
            'Atualiza os campos do boleto bancario
            Dim i As Integer = 0
            For Each tit In ListaDeTitulos
                tit.IUD = "U"
                tit.CodigoEmpresaPagadora = parametros.CodigoEmpresa
                tit.EndEmpresaPagadora = parametros.EnderecoEmpresa
                tit.CodigoBancoPagador = parametros.CodigoBanco
                tit.CodigoAgenciaPagadora = parametros.CodigoAgencia
                tit.DigitoAgenciaPagadora = parametros.DigitoAgencia
                tit.ContaPagadora = parametros.Conta
                tit.DigitoContaPagadora = parametros.DigitoConta
                tit.ContaContabilPagadora = parametros.ContaContabil
                tit.NossoNumero = IIf(tit.NossoNumero = "0", (parametros.SequenciaDePagamentos + i), tit.NossoNumero)
                tit.TipoContaPagadora = "C"
                tit.BoletoBancario = True
                tit.UsuarioBoletoBancario = Session("ssNomeUsuario")
                tit.UsuarioBoletoBancarioDate = Today()
                tit.CodigoSituacao = eSituacao.RemessaBancaria
                tit.CodigoTipoPgto = 4
                tit.SituacaoBancaria = eSituacaoBancaria.AguardandoArquivoDeRetorno
                tit.Salvar()
                i = (i + 1)
            Next
            parametros.SequenciaDePagamentos = (parametros.SequenciaDePagamentos + i)
            parametros.AtualizarParametrosDaCarteira(where, parametros)
        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao atualizar titulos")
        End Try
    End Function

    Protected Function NomeNoPdf(ByRef ListaDeTitulos As List(Of Titulo)) As String
        Dim nomePDF As String = ""
        If ListaDeTitulos.Count() = 1 Then

            Dim tit = ListaDeTitulos(0)

            If tit.Empresa.Codigo = "05272759000147" Then

                Dim historico As String() = tit.Historico.Split(",")

                nomePDF = historico(0).Replace("REF.", "").Replace("-1-S", "").Replace(Environment.NewLine, "")

                If Not nomePDF.Contains("AGRUPAMENTO") Then

                    If historico.Length > 1 Then
                        nomePDF += " " & historico(1).Replace("/", " de ").Replace(Environment.NewLine, "") & " " & tit.Cliente.Nome.Replace("/", "").Replace(".", "") & " - " & tit.Prorrogacao.ToString("dd-MM-yyyy")
                    Else
                        nomePDF += " " & historico(0).Replace("/", " de ").Replace(Environment.NewLine, "") & " " & tit.Cliente.Nome.Replace("/", "").Replace(".", "") & " - " & tit.Prorrogacao.ToString("dd-MM-yyyy")
                    End If

                Else

                    nomePDF += " " & tit.Cliente.Nome.Replace("/", "").Replace(".", "") & " - " & tit.Prorrogacao.ToString("dd-MM-yyyy")

                End If

            Else

                nomePDF = String.Format("{0}_{1}_{2}", tit.Cliente.Nome.Replace("/", "").Replace(".", ""), tit.Codigo, DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss"))

            End If

        End If

        Return nomePDF

    End Function

#End Region

#End Region

End Class