Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports Boleto2Net
Imports BoletoNet
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class FolhaPagamento
    Inherits BasePage

    Private mensagemErro As String
    Private empresa() As String
    Private cliente() As String
    Private ocorrenciasRetornoDictionary As New Dictionary(Of String, String)



#Region "Barra Botão"

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

            GerarRemessa()

            'CarregarTitulosEnviadosAoBanco()
            MsgBox(Me.Page, "Arquivo de remessa gerado com sucesso!", eTitulo.Erro)
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

            If ddlBancoRetorno.SelectedValue.Length = 0 Then
                MsgBox(Me.Page, "Banco não foi selecionado.", eTitulo.Info)
                Exit Sub
            End If
            CarregarTitulosEnviadosAoBanco()
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível consultar os Títulos enviados ao Banco.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        If ddlBanco.SelectedValue.Length = 0 Then
            MsgBox(Me.Page, "Banco não foi selecionado.", eTitulo.Info)
            Exit Sub
        End If

        CarregarVencimentos()
    End Sub


#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("BoletoBancario", "ACESSAR") Then
                BuncarUnidadeDeNegocio()
                CarregarBancos()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If

        If IsPostBack And fup.HasFile Then

            If gridRetornoTitulos.Rows.Count = 0 Then
                MsgBox(Me.Page, "Necessário pesquisar títulos antes de selecionar o arquivo de retorno.", eTitulo.Info)
                Exit Sub
            End If

            If ddlBancoRetorno.SelectedValue.Length = 0 Then
                MsgBox(Me.Page, "Banco não foi selecionado.", eTitulo.Info)
                Exit Sub
            End If

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

            TabContainer1.ActiveTabIndex = 1
            fup.Visible = False

            CarregarRetornoRemessa(parametros)
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

    Protected Sub btnDownload_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDownload.Click
        Try
            Dim where As String = "Where FolhaDePagamento = 1 AND Ativo = 1 AND LEFT(Empresa_Id, 8) = '" & Left(HttpContext.Current.Session("ssEmpresa"), 8) + "'"

            Dim param As New ParametrosDaCarteiraDeCobranca(where)

            If Not Directory.Exists(Server.MapPath("~/FolhaPagamento/" & param.Convenio.ToString)) Then Directory.CreateDirectory(Server.MapPath("~/FolhaPagamento/" & param.Convenio.ToString))

            'Gerar Arquivo de Remessa
            Dim nomeNoArquivo = param.NomeNoArquivo + ".txt"
            Dim path As String = "FolhaPagamento/" & param.Convenio.ToString

            'Outros Bancos
            Download(path, nomeNoArquivo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
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


#End Region

#Region "Methods"

    Protected Sub CarregarVencimentos()
        Try
            Dim Vencimentos As New ListTitulo("Situacao = 105 and FolhaDePagamento = 0")

            If (Vencimentos.Count > 0) Then
                ViewState("Vencimentos") = Vencimentos

                gridFuncionarios.DataSource = From tit In Vencimentos
                                              Select New With
                                              {
                                                  .Codigo = tit.Codigo,
                                                  .Empresa = tit.Empresa.Nome,
                                                  .UnidadeDeNegocio = tit.UnidadeDeNegocio.Nome,
                                                  .Movimento = tit.Movimento,
                                                  .Carteira = tit.Carteira.Descricao,
                                                  .Cliente = tit.Cliente.Codigo,
                                                  .Nome = tit.Cliente.Nome,
                                                  .Historico = tit.Historico,
                                                  .ValorLiquido = tit.ValorLiquido,
                                                  .CodigoCifrado = Funcoes.Cifrar("ContasAPagar-" & tit.Codigo)
                                              }

                gridFuncionarios.DataBind()

                lnkGerarRemessa.Enabled = True
                ddlBanco.Enabled = True
            Else
                MsgBox(Me.Page, "Registro(s) não encontrado(s) para a seleção.")
                Limpar()
            End If

        Catch ex As Exception
            Dim e = ex.Message
        End Try
    End Sub

    Private Sub BuncarUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)
    End Sub

    Private Sub CarregarBancos()

        Dim where As String = "Where isnull(FolhaDePagamento,0) = 1 AND Ativo = 1 AND LEFT(Empresa_Id, 8) = '" & Left(HttpContext.Current.Session("ssEmpresa"), 8) + "'"
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

    Private Sub Limpar()
        lnkGerarRemessa.Enabled = False
        ddlBanco.Enabled = True

        gridRetornoTitulos.DataSource = Nothing
        gridRetornoTitulos.DataBind()
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

    Public Sub ProcessarRetornoRemessa()

        Dim resultado = False

        Try
            For Each row As GridViewRow In gridRetornoTitulos.Rows

                Dim _registroId As String = CType(row.FindControl("hpTitulo"), HyperLink).Text

                Dim tit As New Titulo(_registroId)

                If tit.CodigoProvisao <> eProvisao.Baixa Then

                    Dim retorno As String() = TryCast(row.Cells(3), DataControlFieldCell).Text.Split("-")
                    Dim enumerador = GetEnumValueByDescriptionContains(Of eSituacaoFolhaPagamento)(retorno(0))

                    resultado = AtualizarTitulo(tit, enumerador, CDate(row.Cells(6).Text), CDec(row.Cells(7).Text))

                    If Not resultado Then
                        Exit For
                    End If

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

    Public Sub CarregarTitulosEnviadosAoBanco()

        Dim strConta As String() = ddlBancoRetorno.SelectedValue.Split(";")
        Dim where = "BancoPagador = " & strConta(0) & " AND AgenciaPagadora = " & strConta(1) & " AND ContaPagadora = " & strConta(3)

        gridRetornoTitulos.DataSource = Nothing

        where += " AND Situacao = " & CInt(eSituacao.RemessaBancariaFuncionarios) & " AND Provisao <> " & CInt(eProvisao.Baixa)
        Dim titulos As New ListTitulo(where)

        For Each tit In titulos

            If tit.SituacaoRemessaBancaria <> CInt(eSituacaoFolhaPagamento.AguardandoArquivoDeRetorno) OrElse tit.SituacaoRemessaBancaria <> 0 Then
                Dim enumerador As eSituacaoFolhaPagamento = CType(CInt(tit.SituacaoRemessaBancaria), eSituacaoFolhaPagamento)
                tit.Observacoes = Textos.GetEnumDescription(enumerador)
            Else
                tit.Observacoes = Textos.GetEnumDescription(eSituacaoFolhaPagamento.AguardandoArquivoDeRetorno)
            End If
        Next

        ViewState("Vencimentos_Retorno") = titulos

        gridRetornoTitulos.DataSource = From tit In titulos
                                        Select New With
                                        {
                                          .Codigo = tit.Codigo,
                                          .Empresa = tit.Empresa.Nome,
                                          .UnidadeDeNegocio = tit.UnidadeDeNegocio.Nome,
                                          .Movimento = tit.Observacoes,
                                          .Carteira = tit.Carteira.Descricao,
                                          .Cliente = tit.Cliente.Codigo,
                                          .Nome = tit.Cliente.Nome,
                                          .Historico = tit.Historico,
                                          .ValorLiquido = tit.ValorLiquido,
                                          .Prorrogacao = tit.Prorrogacao,
                                          .CodigoCifrado = Funcoes.Cifrar("ContasAPagar-" & tit.Codigo)
                                        }

        gridRetornoTitulos.DataBind()
    End Sub

    Public Function AtualizarTitulo(ByRef tit As Titulo, ByVal enumerador As eSituacaoFolhaPagamento, ByVal dataDeBaixa As Date, ByVal Juros As Decimal) As Boolean
        tit.SituacaoRemessaBancaria = enumerador
        tit.SituacaoBancaria = 105
        tit.IUD = "U"
        tit.HistoricoRemessa = tit.HistoricoRemessa & Environment.NewLine & "ARQUIVO DE RETORNO - DATA: " & DateTime.Now.ToString("dd/MM/yyyy") & " " & CInt(enumerador).ToString("00") & " - " & Textos.GetEnumDescription(enumerador)

        If enumerador = eSituacaoFolhaPagamento.Liquidacao OrElse enumerador = eSituacaoFolhaPagamento.LiquidacaoEmHomologacao Then

            tit.CodigoProvisao = eProvisao.Baixa

            tit.Baixa = Funcoes.ValidaDataUtil(tit.CodigoEmpresa, tit.EnderecoEmpresa, dataDeBaixa)

            tit.CodigoSituacao = eSituacao.Normal
            tit.UsuarioBaixa = Session("ssNomeUsuario")
            tit.UsuarioBaixaData = Today()

            'tit.CodigoTipoPgto = 4 'Boleto Bançario

            tit.IndiceTitulo = tit.DolarizaBaixa(tit.Baixa.ToString("yyyy/MM/dd"), tit.ValorDoDocumento, IIf(tit.CodigoIndexador = 99, 2, tit.CodigoIndexador))

        End If

        Return tit.Salvar()

    End Function

    Protected Sub GerarRemessa()

        Dim where As String = "Where FolhaDePagamento = 1 AND Ativo = 1 AND LEFT(Empresa_Id, 8) = '" & Left(HttpContext.Current.Session("ssEmpresa"), 8) + "'"

        Dim parametros As New ParametrosDaCarteiraDeCobranca(where)

        Dim folhaPagamento = New NGS.Lib.Negocio.FolhaDePagamento()

        Dim vencimentos = CType(ViewState("Vencimentos"), ListTitulo)

        Dim sb = folhaPagamento.BradescoCNAB240(vencimentos, parametros)

        'Gerar Arquivo de Remessa
        parametros.NomeNoArquivo = parametros.SequenciaDeRemessa.ToString("000000")

        AtualizarTitulos(vencimentos, parametros, where)

        If Not Directory.Exists(Server.MapPath("~/FolhaPagamento/" & parametros.Convenio.ToString)) Then Directory.CreateDirectory(Server.MapPath("~/FolhaPagamento/" & parametros.Convenio.ToString))

        Dim Path As String = "FolhaPagamento/" & parametros.Convenio.ToString & "/" + parametros.SequenciaDeRemessa.ToString("000000") + ".txt"

        Dim bytes = Encoding.ASCII.GetBytes(sb.ToString().ToCharArray())

        Dim strm = New StreamWriter(Server.MapPath(Path))

        For Each line In sb
            strm.WriteLine(line)
        Next

        strm.Dispose()
        strm.Close()

        MostrarBotaoDownload()

    End Sub

    Public Sub MostrarBotaoDownload()
        lnkGerarRemessa.Enabled = False
        ddlBanco.Enabled = False

        btnDownload.Visible = True
    End Sub

    Public Function AtualizarTitulos(ByRef ListaDeTitulos As List(Of Titulo), ByVal parametros As ParametrosDaCarteiraDeCobranca, ByVal where As String) As Boolean

        Try
            'Atualiza os campos do boleto bancario
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
                tit.TipoContaPagadora = "C"
                tit.BoletoBancario = True
                tit.UsuarioBoletoBancario = Session("ssNomeUsuario")
                tit.UsuarioBoletoBancarioDate = Today()
                tit.CodigoSituacao = eSituacao.RemessaBancariaFuncionarios
                'tit.CodigoTipoPgto = 4
                tit.SituacaoBancaria = 105
                tit.SituacaoRemessaBancaria = eSituacaoFolhaPagamento.AguardandoArquivoDeRetorno
                tit.FolhaDePagamento = 1
                tit.Salvar()
            Next

            parametros.AtualizarParametrosDaCarteira(where, parametros)
        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao atualizar titulos")
        End Try
    End Function

    Public Sub CarregarRetornoRemessa(ByRef parametros As ParametrosDaCarteiraDeCobranca)

        Dim arquivo As String = Server.MapPath("~/RetornoBancario/" & parametros.Convenio.ToString & "/" & fup.FileName)
        Dim objArquivo As New StreamReader(arquivo)
        Dim strLinha As String
        'Dim campos() As String
        Dim intLinha As Integer

        Dim ErroMSG As String = String.Empty

        Dim titulosArqRetorno = New List(Of Object)

        Try
            Do While objArquivo.Peek >= 0
                intLinha += 1
                strLinha = objArquivo.ReadLine()
                If intLinha <> 1 Then
                    If Mid(strLinha, 14, 1).ToString().Equals("A") Then

                        Dim codigo = Mid(strLinha, 74, 20).Trim()
                        Dim codRetorno = Mid(strLinha, 231, 240).Trim()
                        Dim tipoMovimentoRetorno = Mid(strLinha, 16, 2)
                        Dim DataPgto = Mid(strLinha, 94, 2) & "/" & Mid(strLinha, 96, 2) & "/" & Mid(strLinha, 98, 4)

                        Dim valorTotal = Mid(strLinha, 163, 15).Insert(13, ",")

                        Dim enumerador = GetEnumValueByDescriptionContains(Of eSituacaoFolhaPagamento)(codRetorno)

                        Dim retorno = IIf(enumerador <> 0, Textos.GetEnumDescription(enumerador), Textos.GetEnumDescription(eSituacaoFolhaPagamento.Erro))


                        ErroMSG = "TITULO " & codigo.ToString & " - PAGAMENTO " & DataPgto.ToString & " - VER RETORNO " & retorno

                        For i = 0 To gridRetornoTitulos.Rows.Count - 1

                            Dim codigoNaGrid = CType(gridRetornoTitulos.Rows(i).FindControl("hpTitulo"), HyperLink).Text
                            If codigo.ToString.Length > 0 AndAlso CInt(codigoNaGrid) = CInt(codigo) AndAlso Not titulosArqRetorno.Any(Function(x) x.Codigo = CInt(codigo)) Then

                                Dim isValid As Boolean = IsDate(DataPgto)

                                Dim prorrogacaoDataPgtoEfetiva = gridRetornoTitulos.Rows(i).Cells(6).Text

                                If isValid Then
                                    prorrogacaoDataPgtoEfetiva = DateTime.ParseExact(DataPgto, "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy")
                                End If


                                titulosArqRetorno.Add(
                                    New With
                                    {
                                        .Codigo = CInt(codigo),
                                        .Empresa = gridRetornoTitulos.Rows(i).Cells(1).Text,
                                        .UnidadeDeNegocio = gridRetornoTitulos.Rows(i).Cells(2).Text,
                                        .Movimento = codRetorno.ToString() & " - " & retorno.ToString,
                                        .Cliente = gridRetornoTitulos.Rows(i).Cells(4).Text,
                                        .Nome = gridRetornoTitulos.Rows(i).Cells(5).Text,
                                        .Historico = gridRetornoTitulos.Rows(i).Cells(6).Text,
                                        .ValorLiquido = CDec(valorTotal),
                                        .Prorrogacao = prorrogacaoDataPgtoEfetiva,
                                        .CodigoCifrado = Funcoes.Cifrar("ContasAPagar-" & codigo)
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

            gridRetornoTitulos.DataSource = titulosArqRetorno
            gridRetornoTitulos.DataBind()

        End Try
    End Sub

    Public Function GetEnumValueByDescriptionContains(Of TEnum As Structure)(descriptionSubstring As String) As TEnum
        If Not GetType(TEnum).IsEnum Then
            Throw New ArgumentException("TEnum must be an Enum type.")
        End If

        For Each field As FieldInfo In GetType(TEnum).GetFields(BindingFlags.Public Or BindingFlags.Static)
            Dim attributes() As DescriptionAttribute = DirectCast(field.GetCustomAttributes(GetType(DescriptionAttribute), False), DescriptionAttribute())
            If attributes.Length > 0 AndAlso attributes(0).Description.Contains(descriptionSubstring) Then
                Return DirectCast(field.GetValue(Nothing), TEnum)
            End If
        Next

        Throw New Exception("Enumerdor não encontrado")
    End Function


#End Region

End Class
