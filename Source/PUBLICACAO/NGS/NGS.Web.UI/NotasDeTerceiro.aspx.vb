Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Threading.Tasks
Imports System.Web.Services.Description
Imports System.Xml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class NotasDeTerceiro
    Inherits BasePage

#Region "Variáveis Locais"

    Private objNotaFiscal As NotaFiscal

    Private Const ApiUrlPattern As String = "{0}/api/NotaFiscal/EmitirNotaFiscal?xml={1}&usarProdutoXML=false&notaDeTerceiro=true&pedido=&empresa="
    Private Const ApiLocal As String = "http://localhost:44323"
    'Private Const ApiServidor As String = "https://baxi.nocti.com.br/ngsApi"

    Public verURL As String

    Private ReadOnly Property ApiServidor As String
        Get
            Dim request = HttpContext.Current.Request
            Dim dominio = request.Url.Host ' exemplo: "nuba.nocti.com.br"
            Dim protocolo = If(request.IsSecureConnection, "https", "http")
            Return $"{protocolo}://{dominio}/ngsApi"
        End Get
    End Property

#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If IsConnect AndAlso Not IsPostBack Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Expedicao.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("NotasDeTerceiro", "ACESSAR") Then
                ddl.Carregar(ddlTipoDeDocumento, CarregarDDL.Tabela.TipoDeDocumento, " Codigo_Id = 1", True)
                ddl.Carregar(ddlSituacao, CarregarDDL.Tabela.Situacao, "")

                LimparCampos()

                If Not Directory.Exists("C:\NGS\Log") Then Directory.CreateDirectory("C:\NGS\Log")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)


    End Sub

    Public Sub AdicionarArquivoBD()

        lnkImportar.Parent.Visible = True

    End Sub

    Private Function IsAmbienteLocal() As Boolean
        Dim url As String = HttpContext.Current.Request.Url.Host.ToLower()
        Return url.Contains("localhost") OrElse url.Contains("127.0.0.1")
    End Function

    Private Sub ImportarNotaDeTerceiro()

        objNotaFiscal = New NotaFiscal
        ucFile.Salvar(objNotaFiscal.Arquivos)

        If objNotaFiscal.Arquivos.Count() = 0 Then
            MsgBox(Me.Page, "É necessário escolher um arquivo para importar!")
            Exit Sub
        ElseIf objNotaFiscal.Arquivos.Count() > 1 Then
            MsgBox(Me.Page, "É necessário escolher apenas um arquivo para importar!")
            Exit Sub
        End If

        'Dim baseUrl As String = If(IsAmbienteLocal(), ApiLocal, ApiServidor)
        'Dim apiUrl As String = String.Format(ApiUrlPattern, baseUrl, objNotaFiscal.Arquivos.FirstOrDefault().Descricao)
        'HTTP://LOCALHOST/NGSAPI/api/NotaFiscal/EmitirNotaFiscal?xml=41241024450490000277550060000019751607615517-nfe.xml&usarProdutoXML=false

        'Dim xmlFileName As String = arquivoMIC.Name
        'Dim apiUrl As String = $"http://localhost/NGSAPI/api/NotaFiscal/EmitirNotaFiscal?xml={Uri.EscapeDataString(xmlFileName)}"

        'Para importar notas ex. NUBA
        'Para executar manualmente em localhost, descomentar
        'apiUrl = String.Format("http://localhost:44323{0}", apiUrl)

        Dim baseUrl As String = If(IsAmbienteLocal(), ApiLocal, ApiServidor)

        Dim nomeArquivo As String = HttpUtility.UrlEncode(objNotaFiscal.Arquivos.First().Descricao)
        Dim pedido As String = IIf(objNotaFiscal.CodigoPedido = 0, "", objNotaFiscal.CodigoPedido)

        Dim apiUrl As String = String.Format("{0}/api/NotaFiscal/EmitirNotaFiscal?xml={1}&usarProdutoXML=false&notaDeTerceiro=true&pedido={2}&empresa={3}",
                                     baseUrl, nomeArquivo, pedido, Session("ssEmpresa"))

        verURL = apiUrl
        ' Chama a API de emissão
        Dim responseData As String = EmitirNotaFiscal(apiUrl)

        Dim sMensagem As String

        If Not responseData.Contains("Erro") Then

            sMensagem = "Nota fiscal emitida com sucesso. Resposta da API: " & responseData

            Dim dados = ExtrairSerieENumeroDaChaveArquivo(objNotaFiscal.Arquivos.FirstOrDefault().Descricao)
            txtSerie.Text = dados.Serie
            txtNumeroNota.Text = dados.Numero

            Consultar()

            If Session("objNFConsultaTerceiro" & HID.Value) IsNot Nothing Then

                Dim objNota As NotaFiscal = CType(Session("objNFConsultaTerceiro" & HID.Value), [Lib].Negocio.NotaFiscal)

                If Not objNotaFiscal Is Nothing Then

                    Dim Sqls As New ArrayList

                    For Each arq In objNotaFiscal.Arquivos

                        arq.CodigoEmpresa = objNota.CodigoEmpresa
                        arq.EnderecoEmpresa = objNota.EnderecoEmpresa
                        arq.CodigoCliente = objNota.CodigoCliente
                        arq.EnderecoCliente = objNota.EnderecoCliente
                        arq.CodigoNota = objNota.Codigo
                        arq.Serie = objNota.Serie
                        arq.CodigoPedido = objNota.CodigoPedido
                        arq.SalvarSql(Sqls)

                    Next

                    If Sqls.Count > 0 Then

                        If Banco.GravaBanco(Sqls) Then

                            ucFile.Clear()

                            If ucFile.Parent.Visible Then
                                ucFile.Bind(objNotaFiscal.Arquivos)
                            End If

                        Else
                            MsgBox(Me.Page, sMensagem)
                        End If

                    End If

                End If

            End If

        Else

            Dim errorMessage As ErrorMessage = JsonResponse.GetErrorMessage(responseData)

            If Not errorMessage.exceptionMessage Is Nothing AndAlso errorMessage.exceptionMessage.Length > 0 Then
                sMensagem = errorMessage.exceptionMessage
                MsgBox(Me.Page, sMensagem)
                LimparCampos()
            End If

        End If

    End Sub

    Public Function ExtrairSerieENumeroDaChaveArquivo(nomeArquivo As String) As DadosNotaFiscal
        Dim chave As String = nomeArquivo.ToUpper.Replace("-NFE.XML", "").Trim()

        If chave.Length <> 44 Then
            Throw New ArgumentException("Chave de acesso inválida. Deve conter exatamente 44 dígitos.")
        End If

        Dim serieRaw As String = chave.Substring(22, 3)
        Dim numeroRaw As String = chave.Substring(25, 9)

        Dim serieFinal As String = If(IsNumeric(serieRaw), serieRaw.TrimStart("0"c), serieRaw)
        Dim numeroFinal As String = If(IsNumeric(numeroRaw), numeroRaw.TrimStart("0"c), numeroRaw)

        Return New DadosNotaFiscal With {
        .Serie = serieFinal,
        .Numero = numeroFinal
    }
    End Function

    Public Function EmitirNotaFiscal(apiUrl As String) As String
        Try
            Dim request As HttpWebRequest = CType(WebRequest.Create(apiUrl), HttpWebRequest)
            request.Method = "GET"
            request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
            request.Headers.Add("Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3")
            request.Headers.Add("Accept-Encoding", "gzip, deflate")
            request.Headers.Add("Cache-Control", "no-cache")
            request.Headers.Add("Pragma", "no-cache")
            request.KeepAlive = True
            request.AllowAutoRedirect = True
            request.Timeout = 900000

            Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Using reader As New StreamReader(response.GetResponseStream())
                    Return reader.ReadToEnd()
                End Using
            End Using

        Catch ex As WebException
            If ex.Response IsNot Nothing Then
                Using reader As New StreamReader(ex.Response.GetResponseStream())
                    Return "Erro HTTP: " & reader.ReadToEnd()
                End Using
            Else
                Return "Erro de conexão: " & ex.Message
            End If

        Catch ex As Exception
            Return "Erro geral: " & ex.Message
        End Try
    End Function


    Public Sub CarregarNotaFiscal()
        If Session("objNFConsultaTerceiro" & HID.Value) IsNot Nothing Then
            Dim objNota As NotaFiscal = CType(Session("objNFConsultaTerceiro" & HID.Value), [Lib].Negocio.NotaFiscal)
            CarregarFormComAClasse(objNota)
        End If
    End Sub

    Private Sub CarregarFormComAClasse(ByVal objNota As NotaFiscal)

        hdfCodigoEmpresa.Value = objNota.CodigoEmpresa & "-" & objNota.EnderecoEmpresa
        If objNota.Cliente IsNot Nothing Then Funcoes.FormatarClienteTXT(txtNomeEmpresa, objNota.Empresa)
        hdfCodigoCliente.Value = objNota.CodigoCliente & "-" & objNota.EnderecoCliente
        If objNota.Cliente IsNot Nothing Then Funcoes.FormatarClienteTXT(txtNomeCliente, objNota.Cliente)

        txtChaveNFe.Text = Funcoes.FormatarChaveNFe(objNota.ChaveNFE)
        txtObservacao.Text = objNota.Observacoes
        ddlSituacao.SelectedValue = objNota.CodigoSituacao
        ddlSituacao.Enabled = False
        ddlTipoDeDocumento.SelectedValue = objNota.CodigoTipoDeDocumento

        lnkAdicionarItem.Parent.Parent.Visible = True
        grdProdutos.Columns(0).Visible = True
        grdProdutos.Columns(13).Visible = True
        lblPedido.Text = objNota.CodigoPedido
        txtNumeroNota.Text = objNota.Codigo
        txtSerie.Text = objNota.Serie
        txtMovimento.Text = objNota.Movimento.ToShortDateString()
        txtDataNota.Text = objNota.DataNota.ToShortDateString()

        'Arquivo
        If ucFile.Parent.Visible Then
            ucFile.Bind(objNota.Arquivos)
        End If

        grdProdutos.DataSource = objNota.Itens.Where(Function(s) s.IUD <> "D").ToList.OrderBy(Function(s) s.Sequencia)
        grdProdutos.DataBind()

        gridEncargosGerais.DataSource = From nfEnc In objNota.Itens.Where(Function(s) s.IUD <> "D").SelectMany(Function(s) s.Encargos)
                                        Group By nfEnc.Codigo Into ValorOficial = Sum(nfEnc.Valor)
                                        Order By IIf(Codigo = "PRODUTO", 1, IIf(Codigo = "LIQUIDO", 3, 2))
                                        Select Codigo, ValorOficial
        gridEncargosGerais.DataBind()

        TabPesoDeChegada.Visible = True

        Dim objNFDestino As New [Lib].Negocio.NotaFiscalXDestino(CType(Session("objNFConsultaTerceiro" & HID.Value), [Lib].Negocio.NotaFiscal), True)

        If objNFDestino.PesoLiquido > 0 Then
            txtBrutoDeChegada.Text = objNFDestino.PesoBruto
            txtDesconto.Text = objNFDestino.Desconto
            txtLiquido.Text = objNFDestino.PesoLiquido
            txtDataDeChegada.Text = objNFDestino.Movimento.ToString("dd/MM/yyyy")
            txtTarifaFrete.Text = objNFDestino.TarifaFrete
            chkSinistro.Checked = objNFDestino.Sinistro
        Else
            txtBrutoDeSaida.Text = String.Format("{0:N0}", objNota.Itens.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.QuantidadeFiscal))
            txtSaldo.Text = String.Format("{0:N0}", Decimal.Zero)
        End If

        objNotaFiscal = objNota
        SessaoSalvaNotaFiscal()

        lnkExcluir.Parent.Visible = True

    End Sub

    Private Sub IniciarExclusaoNotaFiscal()
        SessaoRecuperaNotaFiscal()

        objNotaFiscal.IUD = "D"

        If objNotaFiscal.Salvar Then
            MsgBox(Me.Page, "Nota fiscal " & objNotaFiscal.Codigo & " removida com Sucesso.", eTitulo.Sucess)

            LimparCampos()
        Else
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage" & HID.Value))
        End If
    End Sub

    Private Sub LimparCampos()

        Session.Remove("objNotaFiscal" & HID.Value)
        Session.Remove("objNFConsultaTerceiro" & HID.Value)

        lblUsuario.Text = Session("ssNomeUsuario")

        txtBrutoDeSaida.Text = String.Empty
        txtBrutoDeChegada.Text = String.Empty
        txtDataDeChegada.Text = String.Empty
        txtSaldo.Text = String.Empty
        txtDesconto.Text = String.Empty
        txtLiquido.Text = String.Empty
        txtTarifaFrete.Text = String.Empty
        txtValorDoFrete.Text = String.Empty
        chkSinistro.Checked = False

        TabPesoDeChegada.Visible = False

        lnkExcluir.Parent.Visible = False

        lblPedido.Text = "0"
        txtNumeroNota.Text = String.Empty
        txtSerie.Text = String.Empty
        txtChaveNFe.Text = String.Empty
        txtObservacao.Text = String.Empty
        txtMovimento.Text = String.Empty
        txtDataNota.Text = String.Empty

        txtMovimento.Enabled = True
        grdProdutos.DataSource = Nothing
        grdProdutos.DataBind()

        gridEncargosGerais.DataSource = Nothing
        gridEncargosGerais.DataBind()

        txtDataNota.Enabled = True
        txtNumeroNota.Enabled = True
        txtSerie.Enabled = True
        txtChaveNFe.Enabled = True
        ddlSituacao.Enabled = True

        imgExtratoPedido.Visible = False
        lnkAdicionarItem.Parent.Parent.Visible = True
        grdProdutos.Columns(0).Visible = True
        grdProdutos.Columns(13).Visible = True

        ddlUsuarios.Items.Clear()
        ddlUsuarios.Items.Add("Inc.- " & Session("ssNomeUsuario"))

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucFile.Clear()

        txtNomeEmpresa.Text = String.Empty
        hdfCodigoEmpresa.Value = String.Empty

        txtNomeCliente.Text = String.Empty
        hdfCodigoCliente.Value = String.Empty

        ddlTipoDeDocumento.SelectedIndex = 0
        ddlSituacao.SelectedValue = 1

    End Sub

    Private Sub Consultar()
        'If String.IsNullOrWhiteSpace(txtNumeroNota.Text) And (String.IsNullOrWhiteSpace(txtMovimento.Text) Or String.IsNullOrWhiteSpace(txtDataNota.Text)) And String.IsNullOrWhiteSpace(hdfCodigoCliente.Value) Then
        '    MsgBox(Me.Page, "Informe o número da Nota Fiscal ou período: Data Inicial(Movimento) até Data Final(Data da Nota) ou então um cliente!")
        '    Exit Sub
        'End If

        If Not String.IsNullOrWhiteSpace(hdfCodigoEmpresa.Value) Then
            Dim cliente As String() = hdfCodigoEmpresa.Value.Split("-")
            If cliente IsNot Nothing AndAlso cliente.Length > 0 Then
                Session("ssCnpjDaEmpresa" & HID.Value) = cliente(0)
                Session("ssEndDaEmpresa" & HID.Value) = cliente(1)
            End If
        Else
            Session.Remove("ssCnpjDaEmpresa" & HID.Value)
            Session.Remove("ssEndDaEmpresa" & HID.Value)
        End If

        If Not String.IsNullOrWhiteSpace(hdfCodigoCliente.Value) Then
            Dim cliente As String() = hdfCodigoCliente.Value.Split("-")
            If cliente IsNot Nothing AndAlso cliente.Length > 0 Then
                Session("txtCnpjDoCliente" & HID.Value) = cliente(0)
                Session("txtEndDoCliente" & HID.Value) = cliente(1)
            End If
        Else
            Session.Remove("txtCnpjDoCliente" & HID.Value)
            Session.Remove("txtEndDoCliente" & HID.Value)
        End If

        If Not String.IsNullOrWhiteSpace(txtMovimento.Text) Then
            Session("txtDataDeEmissao" & HID.Value) = txtMovimento.Text
        Else
            Session.Remove("txtDataDeEmissao" & HID.Value)
        End If

        If Not String.IsNullOrWhiteSpace(txtDataNota.Text) Then
            Session("txtDataDeEntrada" & HID.Value) = txtDataNota.Text
        Else
            Session.Remove("txtDataDeEntrada" & HID.Value)
        End If

        Session("txtSerie" & HID.Value) = txtSerie.Text
        Session("txtNumero" & HID.Value) = txtNumeroNota.Text
        Session("ssCampo" & HID.Value) = "NotasDeTerceiro"

        Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaNFG" & HID.Value)
        Dim numberRows As Integer = ucConsultaPedidosXNotas.BindGridView()
        If numberRows = 1 Then
            Popup.CloseDialog(Me.Page, "objNFConsultaNFG" & HID.Value)
        End If
    End Sub

    Protected Sub txtBrutoDeChegada_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If txtBrutoDeChegada.Text.Length = 0 Then txtBrutoDeChegada.Text = 0
        txtDesconto.Text = 0

        If Not String.IsNullOrWhiteSpace(txtBrutoDeChegada.Text) Then
            txtLiquido.Text = String.Format("{0:N0}", CDec(txtBrutoDeChegada.Text))
        End If

        Dim vlrBrutoDeSaida As Decimal = 0
        If Not String.IsNullOrWhiteSpace(txtBrutoDeSaida.Text) Then
            vlrBrutoDeSaida = CDec(txtBrutoDeSaida.Text)
        End If

        Dim vlrBrutoDeChegada As Decimal = 0
        If Not String.IsNullOrWhiteSpace(txtBrutoDeChegada.Text) Then
            vlrBrutoDeChegada = CDec(txtBrutoDeChegada.Text)
        End If

        txtSaldo.Text = String.Format("{0:N0}", vlrBrutoDeChegada - vlrBrutoDeSaida)

    End Sub

    Protected Sub txtDesconto_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim vlrBrutoDeChegada As Decimal = 0
        If Not String.IsNullOrWhiteSpace(txtBrutoDeChegada.Text) Then
            vlrBrutoDeChegada = CDec(txtBrutoDeChegada.Text)
        End If

        Dim vlrDesconto As Decimal = 0
        If Not String.IsNullOrWhiteSpace(txtDesconto.Text) Then
            vlrDesconto = CDec(txtDesconto.Text)
        End If

        txtLiquido.Text = String.Format("{0:N2}", vlrBrutoDeChegada - vlrDesconto)
    End Sub

    Private Sub InclusaoPesoDeChegada()
        If Funcoes.VerificaPermissao("PesosDeChegada", "GRAVAR") Then

            If String.IsNullOrWhiteSpace(txtDataDeChegada.Text) OrElse String.IsNullOrWhiteSpace(txtBrutoDeChegada.Text) OrElse String.IsNullOrWhiteSpace(txtLiquido.Text) Then
                MsgBox(Me.Page, "É necessário informar os campos data de chegada, bruto de chegada e líquido!")
                Exit Sub
            End If

            If Not String.IsNullOrWhiteSpace(txtBrutoDeChegada.Text) AndAlso Not CInt(txtBrutoDeChegada.Text) > 0 Then
                MsgBox(Me.Page, "É necessário informar bruto de chegada maior que zero!")
                Exit Sub
            End If

            If Not String.IsNullOrWhiteSpace(txtLiquido.Text) AndAlso Not CInt(txtLiquido.Text) > 0 Then
                MsgBox(Me.Page, "É necessário informar líquido de chegada maior que zero!")
                Exit Sub
            End If

            Dim objNFDestino As New [Lib].Negocio.NotaFiscalXDestino(CType(Session("objNFConsultaTerceiro" & HID.Value), [Lib].Negocio.NotaFiscal))

            objNFDestino.IUD = "I"

            objNFDestino.Movimento = CDate(txtDataDeChegada.Text)

            If txtBrutoDeChegada.Text.Length = 0 Then txtBrutoDeChegada.Text = "0"
            objNFDestino.PesoBruto = txtBrutoDeChegada.Text

            If txtDesconto.Text.Length = 0 Then txtDesconto.Text = "0"
            objNFDestino.Desconto = txtDesconto.Text

            objNFDestino.PesoLiquido = txtLiquido.Text

            objNFDestino.Sinistro = chkSinistro.Checked

            If txtTarifaFrete.Text.Length = 0 Then txtTarifaFrete.Text = "0"
            objNFDestino.TarifaFrete = txtTarifaFrete.Text

            Dim sqls As New ArrayList
            objNFDestino.SalvarSql(sqls)

            If Banco.GravaBanco(sqls) Then
                MsgBox(Me.Page, "Peso de chegada incluido com sucesso!")
                LimparCampos()
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro!")
        End If
    End Sub

    Protected Sub BtnGravar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnGravar.Click
        Try
            InclusaoPesoDeChegada()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("NotasDeTerceiro", "LEITURA") Then
                Consultar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível consultar a nota")
        End Try
    End Sub

    Protected Sub lnkImportar_Click(sender As Object, e As EventArgs) Handles lnkImportar.Click
        Try
            If Funcoes.VerificaPermissao("NotasDeTerceiro", "LEITURA") Then
                ImportarNotaDeTerceiro()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível consultar a nota, verifique a estrutura do XML. " & verURL & " - " & ex.Message.ToString)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("NotasDeTerceiro", "EXCLUIR") Then
                If grdProdutos.Rows.Count > 0 Then
                    IniciarExclusaoNotaFiscal()
                Else
                    MsgBox(Me.Page, "Consulte a Nota para Exclusão.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível limpar a nota")
        End Try
    End Sub

End Class