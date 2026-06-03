Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Windows.Forms
Imports System.Drawing.Printing
Imports iTextSharp.text.pdf.parser
Imports iTextSharp.text.pdf

Partial Class RegistrosDeICMS
    Inherits BasePage

    Dim Sql As String
    Dim Sql2 As String
    Dim Row As DataRow

    Dim ds As New DataSet
    Dim ds2 As New DataSet
    Dim dss As New DataSet

    Dim SqlArray As New ArrayList
    Dim Empresa() As String

    Dim Descricao As String = ""
    Dim Mensagem As String = ""

    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date

    Dim Processo As Integer = 0
    Dim Livro As Integer = 0
    Dim Folha As Integer = 0
    Dim Vencimento As Date
    Dim CodigoDaReceita As String

    Dim Opcao As String = ""
    Dim EmpresaNome As String = ""
    Dim EmpresaCNPJ As String = ""
    Dim EmpresaInscricao As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Fiscal)
            If Not IsPostBack And IsConnect Then

                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Fiscal.aspx", eTitulo.Info)
                    Exit Sub
                End If

                If Funcoes.VerificaPermissao("RegistrosDeICMS", "ACESSAR") Then
                    CarregarUnidade()
                    VerificaUnidade()
                    CriaOpcoes()
                    HID.Value = Guid.NewGuid().ToString
                    ucRegistrosIcmsAjustaResumo.SetarHID(HID.Value)
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Fiscal.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarUnidade()
        Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & vbCrLf &
              "FROM Clientes C " & vbCrLf &
              "INNER JOIN ClientesXTipos CT " & vbCrLf &
              "ON C.Cliente_Id = CT.Cliente_Id " & vbCrLf &
              "WHERE CT.Tipo_Id = 050 " & vbCrLf &
              "ORDER BY Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            ddlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
        Next

        ddlUnidade.Items.Insert(0, "")
        ddlUnidade.SelectedIndex = 0
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf &
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf &
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf &
              " from Usuarios" & vbCrLf &
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresas()
            DdlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
            CarregarProcesso()
        Next
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaEmpresas()
        If String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            Exit Sub
        End If

        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresa.Items.Clear()

        HttpContext.Current.Session("EmpresaIcms") = ""
        HttpContext.Current.Session("EndEmpresaIcms") = ""
        HttpContext.Current.Session("ProcessoIcms") = 0

        Sql = " SELECT Clientes.Cliente_Id as Codigo," & vbCrLf &
              "        Clientes.Endereco_Id," & vbCrLf &
              "        Clientes.Reduzido," & vbCrLf &
              "        Clientes.Nome," & vbCrLf &
              "        Clientes.Cidade," & vbCrLf &
              "        Clientes.Estado" & vbCrLf &
              "   FROM GruposXEmpresas" & vbCrLf &
              "  INNER JOIN Clientes" & vbCrLf &
              "     ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id" & vbCrLf &
              "    AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
              "  Where GruposXEmpresas.Empresa_Id = '" & ddlUnidade.SelectedValue & "'" & vbCrLf

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(Sql, "Clientes")

        For Each Dr As DataRow In ds.Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
            DdlEmpresa.Items.Add(New ListItem(Descricao, Codigo))
        Next

        DdlEmpresa.Items.Insert(0, "")
        DdlEmpresa.SelectedIndex = 0
    End Sub

    Protected Sub DdlEmpresa_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarProcesso()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlProcesso_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not String.IsNullOrWhiteSpace(DdlProcesso.SelectedValue) Then
                Empresa = DdlEmpresa.SelectedValue.Split("-")
                HttpContext.Current.Session("ProcessoIcms") = DdlProcesso.SelectedValue

                Sql = "SELECT * " & vbCrLf &
                      "  FROM ProcessoRAICMS " & vbCrLf &
                      " WHERE Empresa_Id = '" & Empresa(0) & "'" & vbCrLf &
                      "   AND Processo_Id = " & DdlProcesso.SelectedValue & vbCrLf &
                      " ORDER BY PeriodoFinal DESC" & vbCrLf

                For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                    txtDataInicial.Text = Dr("PeriodoInicial")
                    txtDataFinal.Text = Dr("PeriodoFinal")
                    txtLivro.Text = Dr("Livro")
                    txtFolha.Text = Dr("PaginaInicial")
                    txtCodigoDaReceita.Text = Dr("CodigoDaReceita")
                    txtVenctoDaObrigacao.Text = Dr("VencimentodaObrigacao")

                    'Liberado apenas para Curtume e Química
                    If (Left(Session("ssEmpresa"), 8) = "03189063" Or
                        Left(Session("ssEmpresa"), 8) = "05272759") Then
                        Exit Sub
                    End If

                    lnkNovo.Parent.Visible = False
                    lnkExcluir.Parent.Visible = False
                    txtDataInicial.Enabled = False
                    txtDataFinal.Enabled = False
                    txtLivro.Enabled = False
                    txtCodigoDaReceita.Enabled = False
                    txtFolha.Enabled = False
                    txtVenctoDaObrigacao.Enabled = False
                Next
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarProcesso()
        If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            Exit Sub
        End If

        Empresa = DdlEmpresa.SelectedValue.Split("-")
        HttpContext.Current.Session("EmpresaIcms") = Empresa(0)
        HttpContext.Current.Session("EndEmpresaIcms") = Empresa(1)
        DdlProcesso.Items.Clear()

        Sql = "SELECT * " & vbCrLf &
              " FROM ProcessoRAICMS " & vbCrLf &
              " WHERE Empresa_Id = '" & Empresa(0) & "'" & vbCrLf &
              " ORDER BY PeriodoFinal DESC" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Descricao = "Processo " & Format(Dr("Processo_Id"), "0000") & "      Período  "
            Descricao &= CDate(Dr("PeriodoInicial")).ToString("dd/MM/yyyy") & "  à  "
            Descricao &= CDate(Dr("PeriodoFinal")).ToString("dd/MM/yyyy") & "  Livro  "
            Descricao &= Format(Dr("Livro"), "000") & "  Folha  "
            Descricao &= Format(Dr("PaginaInicial"), "000")

            DdlProcesso.Items.Add(New ListItem(Descricao, Dr("Processo_ID")))
        Next

        DdlProcesso.Items.Insert(0, "")
        DdlProcesso.SelectedIndex = 0
    End Sub

    Protected Sub btnConsultaClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConsultaClientes.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)

        If Not Session("objCliente" & HID.Value) Is Nothing Then

            Dim objCliente As [Lib].Negocio.Cliente = Session("objCliente" & HID.Value)
            Session.Remove("objCliente" & HID.Value)

            'Row.CodigoCliente = objCliente.Codigo
            'Row.EndCliente = objCliente.CodigoEndereco

            'Row.CodigoDestinatario = objCliente.Codigo
            'Row.EndDestinatario = objCliente.CodigoEndereco
            txtNomeCliente.Text = objCliente.Nome & "|" & objCliente.Codigo & "|" & objCliente.CodigoEndereco

        End If

    End Sub

    Private Sub impressaoPDF(ByVal dsPrint As DataSet, ByVal entradaSaida As String, ByVal opcoes As String)

        Try

            If ValidarImpressao() Then

                Dim dt As DataTable = dsPrint.Tables(0) ' A tabela do DataSet onde está a coluna "nota"
                Dim pageNumber As Integer = 1 ' Inicia a contagem de páginas

                If txtFolhaInicial.Text.Length > 0 Then
                    pageNumber = txtFolha.Text
                End If

                Dim rowCount As Integer = 0 ' Contador de linhas para controle da troca de página
                Dim caminhoLivro As String = IO.Path.Combine(Server.MapPath("~"), "LivroFiscal")
                Dim empresa() As String = DdlEmpresa.SelectedValue.Split("-")

                Dim pastaLivro As String = CDate(txtDataInicial.Text).ToString("yyyyMM")
                Dim pathFiles As String = String.Format("{0}\{1}\{2}\{3}", caminhoLivro, empresa(0), opcoes, pastaLivro)
                Dim urlPath As String = String.Format("LivroFiscal/{0}/{1}/{2}", empresa(0), opcoes, pastaLivro)

                If chkZIP.Checked Then

                    'Se a opção de zipar estiver marcado, precisamos limpar os arquivos se houver seleção de paginas
                    If Directory.Exists(String.Format("{0}\{1}\{2}\{3}", caminhoLivro, empresa(0), opcoes, pastaLivro)) Then

                        Directory.Delete(String.Format("{0}\{1}\{2}\{3}", caminhoLivro, empresa(0), opcoes, pastaLivro), True)

                    End If

                End If

                Funcoes.CriarPastaLivroFiscal(String.Format("{0}\{1}\{2}\{3}", caminhoLivro, empresa(0), opcoes, pastaLivro))

                For Each rowPrint As DataRow In dt.Rows

                    rowCount += 1

                    'Linhas por paginas do pdf
                    If rowCount > 24 Then

                        pageNumber += 1
                        Console.WriteLine("Página " & pageNumber)
                        rowCount = 1 ' Reseta o contador de linhas para a nova página

                    End If

                    If txtFolhaInicial.Text.Length > 0 AndAlso CType(txtFolhaInicial.Text, Integer) > 0 Then
                        If pageNumber < CType(txtFolhaInicial.Text, Integer) Then
                            Continue For
                        End If
                    End If

                    If txtFolhaFinal.Text.Length > 0 AndAlso CType(txtFolhaFinal.Text, Integer) > 0 Then
                        If pageNumber > CType(txtFolhaFinal.Text, Integer) Then
                            Exit For
                        End If
                    End If

                    ' Obtem o valor da coluna "nota"
                    Dim objCliente As New [Lib].Negocio.Cliente()
                    Dim cliente As String = rowPrint("Cliente").ToString()
                    Dim inscricao As String = rowPrint("InscEstadual").ToString()
                    Dim serie As String = rowPrint("Serie").ToString()
                    Dim nota As String = rowPrint("NumeroDaNota").ToString()
                    Dim objNFe As New [Lib].Negocio.NotaFiscal()

                    objCliente.Codigo = cliente
                    objCliente.InscricaoEstadual = inscricao
                    objCliente.BuscarClientePelaInscricaoEstadual()
                    objNFe.CodigoEmpresa = empresa(0)
                    objNFe.EnderecoEmpresa = empresa(1)
                    objNFe.CodigoCliente = objCliente.Codigo
                    objNFe.EnderecoCliente = objCliente.CodigoEndereco
                    objNFe.EntradaSaida = IIf(entradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                    objNFe.Serie = serie
                    objNFe.Codigo = nota
                    objNFe = New [Lib].Negocio.NotaFiscal(objNFe)

                    Dim msgNFE As String = ""

                    Dim caminhoArquivo As String

                    If objNFe.ChaveNFE.Length > 0 Then
                        caminhoArquivo = String.Format("{0}\{1}.pdf", pathFiles, objNFe.ChaveNFE)
                    Else
                        caminhoArquivo = String.Format("{0}\{1}_{2}_{3}.pdf", pathFiles, objNFe.CodigoEmpresa, objNFe.CodigoCliente, objNFe.Codigo.ToString())
                    End If

                    If Not File.Exists(caminhoArquivo) Then

                        Dim bytes As Byte() = Nothing

                        If objNFe.Arquivos.Where(Function(x) x.Arquivo.Length > 0 And x.Descricao.ToUpper().Contains(".PDF")).Count > 0 Then

                            Dim arquivo As New [Lib].Negocio.Arquivo
                            arquivo = objNFe.Arquivos.Where(Function(x) x.Arquivo.Length > 0 And x.Descricao.ToUpper().Contains(".PDF")).FirstOrDefault()
                            bytes = arquivo.Arquivo
                        End If

                        If bytes Is Nothing OrElse bytes.Length = 0 Then

                            If DocumentoEletronico.ImprimirNFeDanfe(objNFe, 1, msgNFE) Then
                                bytes = New FilesManager().getFile(String.Format("{0}.pdf", objNFe.ChaveNFE), eTipoDeDocumento.Nota)
                            Else
                                If objNFe.ChaveNFE.Length = 0 Then
                                    msgNFE = String.Format("{0} - Empresa: {1} - Cliente: {2} - Nota Fiscal: {3}", msgNFE, objNFe.CodigoEmpresa, objNFe.CodigoCliente, objNFe.Codigo.ToString())
                                End If
                                MsgBox(Me.Page, msgNFE)
                                Continue For
                            End If

                        End If

                        If bytes IsNot Nothing Then

                            System.IO.File.WriteAllBytes(caminhoArquivo, bytes)

                        End If

                    End If

                    If chkImpressao.Checked Then
                        Funcoes.ImprimirArquivo(Me.Page, caminhoArquivo, objNFe)
                    End If

                    If chkAbrir.Checked Then

                        Dim url As String = HttpContext.Current.Request.Url.AbsoluteUri
                        Dim domainServer As String

                        If HttpContext.Current.Request.Url.Query.Length > 0 Then
                            domainServer = url.Replace(HttpContext.Current.Request.Url.Query, "").Replace(HttpContext.Current.Request.Url.LocalPath, "")
                        Else
                            domainServer = url.Replace(HttpContext.Current.Request.Url.LocalPath, "")
                        End If

                        Dim fileInfo As FileInfo = New FileInfo(caminhoArquivo)
                        Dim bServidor As Boolean

                        If url.ToUpper().Contains("/NGS/") Or url.ToUpper().Contains("/NGSTESTE/") Then
                            bServidor = True
                        End If

                        If bServidor Then
                            If url.ToUpper().Contains("/NGS/") Then
                                Funcoes.AbrirArquivo(Me.Page, String.Format("{0}/ngs/{1}/{2}", domainServer, urlPath, fileInfo.Name))
                            Else
                                Funcoes.AbrirArquivo(Me.Page, String.Format("{0}/ngsTeste/{1}/{2}", domainServer, urlPath, fileInfo.Name))
                            End If
                        Else
                            Funcoes.AbrirArquivo(Me.Page, String.Format("{0}/{1}/{2}", domainServer, urlPath, fileInfo.Name))
                        End If

                    End If

                Next

                If chkZIP.Checked Then

                    Funcoes.DownloadZIP(Me.Page, pathFiles)

                End If

            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Private Sub CriaOpcoes()
        Dim i As Integer = 0
        Dim dsOpcoes As New DataSet
        Dim dtOpcoes As DataTable
        Dim Opcoes As Integer = 12

        dtOpcoes = New DataTable("Opcoes")
        dtOpcoes.Columns.Add("Codigo", Type.GetType("System.String")).DefaultValue = ""
        dtOpcoes.Columns.Add("Descricao", Type.GetType("System.String")).DefaultValue = ""

        While i < Opcoes
            Row = dtOpcoes.NewRow()
            Select Case i
                Case 0
                    Row("Codigo") = "01"
                    Row("Descricao") = "Livro Registro de Entradas - RE -  Modelo P1"
                Case 1
                    Row("Codigo") = "02"
                    Row("Descricao") = "Livro Registro de Saídas - RS - Modelo P2"
                Case 2
                    Row("Codigo") = "03"
                    Row("Descricao") = "Livro Registro de Apuração do ICMS - Ajustes Outros Débitos/Créditos"
                Case 3
                    Row("Codigo") = "04"
                    Row("Descricao") = "Livro Registro de Apuração do ICMS – RAICMS Modelo P9 - Resumo da Apuração do Imposto"
                Case 4
                    Row("Codigo") = "05"
                    Row("Descricao") = "Livro Registro de Apuração do ICMS – RAICMS Modelo P9 - Resumo por Alíquota"
                Case 5
                    Row("Codigo") = "06"
                    Row("Descricao") = "Livro Resumo de Entradas Por U.F."
                Case 6
                    Row("Codigo") = "07"
                    Row("Descricao") = "Livro Resumo de Saidas Por U.F."
                Case 7
                    Row("Codigo") = "08"
                    Row("Descricao") = "Códigos de Emitentes"
                Case 8
                    Row("Codigo") = "09"
                    Row("Descricao") = "Observações Fiscais"
                Case 9
                    Row("Codigo") = "10"
                    Row("Descricao") = "Termos de Abertura e Encerramento"
                Case 10
                    Row("Codigo") = "11"
                    Row("Descricao") = "Registro DAR - Diferencial de Alíquota"
                Case 11
                    Row("Codigo") = "12"
                    Row("Descricao") = "Impressão do Livro Registro de Apuração do ICMS – RAICMS Modelo P9 - Resumo da Apuração do Imposto"
            End Select

            dtOpcoes.Rows.Add(Row)
            i += 1
        End While

        GridOpcoes.DataSource = dtOpcoes
        GridOpcoes.DataBind()
    End Sub

    Private Sub CriaTermos()
        Dim i As Integer = 0
        Dim dsOpcoes As New DataSet
        Dim dtOpcoes As DataTable
        Dim Opcoes As Integer = 3

        dtOpcoes = New DataTable("Opcoes")
        dtOpcoes.Columns.Add("Codigo", Type.GetType("System.String")).DefaultValue = ""
        dtOpcoes.Columns.Add("Descricao", Type.GetType("System.String")).DefaultValue = ""

        While i < Opcoes
            Row = dtOpcoes.NewRow()
            Select Case i
                Case 0
                    Row("Codigo") = "01"
                    Row("Descricao") = "Livro Registro de Entradas"
                Case 1
                    Row("Codigo") = "02"
                    Row("Descricao") = "Livro Registro de Saidas"
                Case 2
                    Row("Codigo") = "03"
                    Row("Descricao") = "Registro de Apuração do ICMS"
            End Select

            dtOpcoes.Rows.Add(Row)
            i += 1
        End While

        GridTermos.DataSource = dtOpcoes
        GridTermos.DataBind()
    End Sub

    Protected Sub GridOpcoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridOpcoes.SelectedIndexChanged
        Try
            If Validar() Then

                GridTermos.Parent.Visible = False

                Select Case GridOpcoes.SelectedRow.Cells(1).Text()
                    Case "01"
                        RegistroDeEntradasModeloP1()
                    Case "02"
                        RegistroDeSaidasModeloP1()
                    Case "03"
                        If Funcoes.VerificaPermissao("RegistrosDeICMS", "GRAVAR") AndAlso Funcoes.VerificaPermissao("RegistrosDeICMS", "ALTERAR") Then
                            ucRegistrosIcmsAjustaResumo.BindGridView()
                            Popup.ConsultaDeRegistrosIcmsAjustaResumo(Me.Page, "objRegistrosIcmsAjustaResumo" & HID.Value)
                        Else
                            MsgBox(Me.Page, "Usuário sem permissão para incluir/alterar registro.")
                        End If
                    Case "04"
                        RegistroDeApuracaoDeIcmsModeloP9(True)
                    Case "05"
                        RegistroDeIcmsPorAliquota()
                    Case "06"
                        ResumoDeEntradasPorUF()
                    Case "07"
                        ResumoDeSaidasPorUF()
                    Case "08"
                        CodigosDeEmitentes()
                    Case "09"
                        ObservacoesFiscais()
                    Case "10"
                        GridTermos.Parent.Visible = True
                        CriaTermos()
                    Case "11"
                        Dim parameters As New Dictionary(Of String, Object)
                        parameters.Add("Ano", CDate(txtDataInicial.Text).Year)
                        parameters.Add("Mes", CDate(txtDataInicial.Text).Month)
                        parameters.Add("Empresa", DdlEmpresa.SelectedValue)

                        ucDarDiferencialDeAliquota.BindGridView(parameters)
                        Popup.ConsultaDarDiferencialDeAliquota(Me.Page, "objDarDiferencialDeAliquota" & HID.Value)
                    Case "12"
                        RegistroDeApuracaoDeIcmsModeloP9(False)
                End Select
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridTermos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Validar() Then
                Select Case GridTermos.SelectedRow.Cells(1).Text()
                    Case "01"
                        Opcao = "01"
                        Termos()
                    Case "02"
                        Opcao = "02"
                        Termos()
                    Case "03"
                        Opcao = "04"
                        Termos()
                End Select
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#Region "Registros Fiscais de Entrada Modelo P1"

    Private Function getDatasetEntradasModeloP1() As DataSet
        Try
            Sql = " SELECT Cliente_Id AS Cliente, Nome, Inscricao " & vbCrLf &
                  "  FROM Clientes " & vbCrLf &
                  " WHERE Cliente_Id = '" & Empresa(0) & "'" & vbCrLf &
                  "   and Endereco_Id = " & Empresa(1) & vbCrLf
            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            For Each Dr As DataRow In ds.Tables(0).Rows
                EmpresaNome = Dr("Nome")
                EmpresaCNPJ = Funcoes.FormatarCpfCnpj(Dr("Cliente"))
                EmpresaInscricao = Dr("Inscricao")
            Next

            Sql = String.Empty
            'Temporaria Controla Notas de Serviço Conjugadas
            If CkConsNotasCompServicos.Checked Then
                Sql = "Select nfxi.Empresa_Id," & vbCrLf &
                       "       nfxi.EndEmpresa_Id," & vbCrLf &
                       "       nfxi.Cliente_Id, nfxi.EndCliente_Id, nfxi.EntradaSaida_Id, nfxi.Serie_Id," & vbCrLf &
                       "       nfxi.Nota_Id," & vbCrLf &
                       "       sum(Case" & vbCrLf &
                       "              when OEx.CodigoFiscal in (1933,2933)" & vbCrLf &
                       "                then 1" & vbCrLf &
                       "                else 0" & vbCrLf &
                       "            end) ItensCfop," & vbCrLf &
                       "       sum(1) nrItens" & vbCrLf &
                       "  into #NotasCompostas" & vbCrLf &
                       "  From notasfiscaisxitens nfxi" & vbCrLf &
                       " Inner join OperacaoXEstado OEx" & vbCrLf &
                       "    on Oex.Codigo_Id = nfxi.OperacaoXEstado" & vbCrLf &
                       " Inner join (SELECT NF.Empresa_Id,      NF.EndEmpresa_Id," & vbCrLf &
                       "                    NF.Cliente_Id,      NF.EndCliente_Id," & vbCrLf &
                       "                    NF.EntradaSaida_Id, NF.Serie_Id,      NF.Nota_Id" & vbCrLf &
                       "               FROM NotasFiscais NF" & vbCrLf &
                       "              INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                       "                 ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                       "                AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                       "                AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                       "                AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                       "                AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                       "                AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                       "                AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                       "              Inner join OperacaoXEstado OE" & vbCrLf &
                       "			     on Oe.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                       "              Inner Join Suboperacoes so" & vbCrLf &
                       "                 on so.Operacao_id     = nfi.Operacao" & vbCrLf &
                       "                and so.SubOperacoes_id = nfi.suboperacao" & vbCrLf &
                       "              WHERE NF.EntradaSaida_Id = 'E'" & vbCrLf &
                       "                And NF.Empresa_Id ='" & Empresa(0) & "' " & vbCrLf &
                       "                AND NF.TipoDeDocumento NOT IN (15)" & vbCrLf &
                       "                AND NF.Finalidade NOT IN (30)" & vbCrLf &
                       "                AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                       "                AND NF.Serie_Id  NOT IN ('D', 'F', 'REC')" & vbCrLf &
                       "                AND Oe.CodigoFiscal  IN (1933,2933)" & vbCrLf &
                       "              )sb" & vbCrLf &
                       "    ON sb.Empresa_Id      = nfxi.Empresa_Id" & vbCrLf &
                       "   AND sb.EndEmpresa_Id   = nfxi.EndEmpresa_Id" & vbCrLf &
                       "   AND sb.Cliente_Id      = nfxi.Cliente_Id" & vbCrLf &
                       "   AND sb.EndCliente_Id   = nfxi.EndCliente_Id" & vbCrLf &
                       "   AND sb.EntradaSaida_Id = nfxi.EntradaSaida_Id" & vbCrLf &
                       "   AND sb.Serie_Id        = nfxi.Serie_Id" & vbCrLf &
                       "   AND sb.Nota_Id         = nfxi.Nota_Id" & vbCrLf &
                       " group by nfxi.Empresa_Id, nfxi.EndEmpresa_Id," & vbCrLf &
                       "          nfxi.Cliente_Id, nfxi.EndCliente_Id," & vbCrLf &
                       "          nfxi.EntradaSaida_Id, nfxi.Serie_Id, nfxi.Nota_Id" & vbCrLf

                If PeriodoFinal > "2011-12-31" Then
                    Sql &= "  having sum(1)  <> sum(Case " & vbCrLf &
                            "                          when OEx.CodigoFiscal in (1933,2933)" & vbCrLf &
                            "                            then 1" & vbCrLf &
                            "                            else 0" & vbCrLf &
                            "                        end)" & vbCrLf
                End If
            End If

            Sql &= " SELECT nxe.Empresa_id, nxe.EndEmpresa_id, nxe.Cliente_id, nxe.EndCliente_Id, nxe.EntradaSaida_id, nxe.Serie_id," & vbCrLf &
                  "        nxe.Nota_id, nxe.Produto_id, nxe.Sequencia_id," & vbCrLf &
                  "        isnull(e.EncargoAgrupador,nxe.Encargo_Id) as Encargo_id," & vbCrLf &
                  "		   oe.sticms," & vbCrLf &
                  "		   oe.stIPI," & vbCrLf &
                  "        oe.stPISCOFINS," & vbCrLf &
                  "	       nxe.Valor," & vbCrLf &
                  "	       nxe.base" & vbCrLf &
                  "   INTO #Encargos" & vbCrLf &
                  "   FROM NotasFiscais NF" & vbCrLf &
                  "  Inner Join NotasFiscaisXItens NFI" & vbCrLf &
                  "     on NF.Empresa_id      = NFI.Empresa_Id" & vbCrLf &
                  "    and NF.EndEmpresa_Id   = NFI.EndEmpresa_ID" & vbCrLf &
                  "    and NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf &
                  "    and NF.EndCliente_Id   = NFI.EndCliente_ID" & vbCrLf &
                  "    and NF.EntradaSaida_id = NFI.EntradaSaida_ID" & vbCrLf &
                  "    and NF.Serie_id        = NFI.Serie_Id" & vbCrLf &
                  "    and NF.Nota_id         = NFI.Nota_Id" & vbCrLf &
                  "  INNER JOIN NotasFiscaisXEncargos nxe" & vbCrLf &
                  "     on NxE.Empresa_id      = NFI.Empresa_Id" & vbCrLf &
                  "    and NxE.EndEmpresa_Id   = NFI.EndEmpresa_ID" & vbCrLf &
                  "    and NxE.Cliente_Id      = NFI.Cliente_Id" & vbCrLf &
                  "    and NxE.EndCliente_Id   = NFI.EndCliente_ID" & vbCrLf &
                  "    and NxE.EntradaSaida_id = NFI.EntradaSaida_ID" & vbCrLf &
                  "    and NxE.Serie_id        = NFI.Serie_Id" & vbCrLf &
                  "    and NxE.Nota_id         = NFI.Nota_Id" & vbCrLf &
                  "	   and NxE.Produto_Id      = NFI.Produto_Id" & vbCrLf &
                  "	   and NxE.Sequencia_id    = NFI.Sequencia_id" & vbCrLf &
                  "  Inner join OperacaoXEstado OE" & vbCrLf &
                  "     on oe.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                  "   Left join Encargos e" & vbCrLf &
                  "     on e.Encargo_id = nxe.Encargo_Id" & vbCrLf &
                  "    and isnull(e.EncargoAgrupador,'') <> ''" & vbCrLf &
                  "  WHERE (isnull(e.EncargoAgrupador,nxe.Encargo_Id) IN ('IPI', 'PRODUTO','FRETES','SEGURO','DESP.ADUANEIRAS','DESCONTOS','ICMS','ICMS SUBSTITUIC'))" & vbCrLf &
                  "    And NF.EntradaSaida_Id = 'E'" & vbCrLf &
                  "    AND NF.TipoDeDocumento NOT IN (15)" & vbCrLf &
                  "    AND NF.Finalidade NOT IN (30)" & vbCrLf &
                  "    And NF.Empresa_Id      = '" & Empresa(0) & "'" & vbCrLf &
                  "    AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf &
                  "    AND NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf

            If txtNomeCliente.Text.Length > 0 Then
                Dim cliente As String() = txtNomeCliente.Text.Split("|")

                Sql &= "AND NF.Cliente_Id = '" & cliente(1) & "'" & vbCrLf
                Sql &= "AND NF.EndCliente_Id = " & cliente(2) & "" & vbCrLf

            End If

            If txtSerie.Text.Length > 0 Then
                Sql &= "AND NF.Serie_Id = '" & txtSerie.Text & "'" & vbCrLf
            End If

            If txtNumeroNota.Text.Length > 0 Then
                Sql &= "AND NF.Nota_Id = '" & txtNumeroNota.Text & "'" & vbCrLf
            End If

            Sql &= "SELECT Movimento, '' As Especie, Serie, NumeroDaNota, CFOP," & vbCrLf &
                   "       DataDaNota, Cliente, CodigoContabil," & vbCrLf &
                   "       CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE ((ValorContabil+ValorIPI+OutrasDespesas+ValorContabilFrete) - ValorDesconto) END AS ValorContabil," & vbCrLf &
                   "       CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE BaseICMS  END AS BaseICMS," & vbCrLf &
                   "       CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE ValorIcms END AS ValorIcms," & vbCrLf &
                   "       CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE BaseIPI   END AS BaseIPI," & vbCrLf &
                   "       CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE ValorIPI  END AS ValorIPI," & vbCrLf &
                   "       ROUND(CASE WHEN BaseIcms = 0 THEN 0 ELSE ValorIcms * 100 / BaseIcms END, 0) AS AliquotaIcms," & vbCrLf &
                   "       ROUND(CASE WHEN Baseipi = 0  THEN 0 ELSE Valoripi * 100  / Baseipi  END, 0) AS AliquotaIPI," & vbCrLf &
                   "       Estado, Consulta.Situacao, Consulta.DescSituacao, Consulta.InscEstadual" & vbCrLf &
                   "  FROM (SELECT NotasFiscais.Nota_Id AS NumeroDaNota, NotasFiscais.Serie_Id AS Serie, NotasFiscais.Cliente_Id AS Cliente, NotasFiscais.Movimento," & vbCrLf &
                   "               NotasFiscais.DataDaNota, OE.CodigoFiscal AS CFOP, NotasFiscais.EstadoDoCliente as Estado,S.Situacao_id as Situacao,S.Descricao as DescSituacao," & vbCrLf &
                   "               SUM(CASE WHEN #Encargos.Encargo_id = 'PRODUTO'   THEN #Encargos.Valor ELSE 0 END) AS ValorContabil," & vbCrLf &
                   "               SUM(CASE WHEN #Encargos.Encargo_id IN ('FRETES','SEGURO') THEN #Encargos.Valor ELSE 0 END) AS ValorContabilFrete," & vbCrLf &
                   "               SUM(CASE WHEN #Encargos.Encargo_id = 'DESCONTOS' THEN #Encargos.Valor ELSE 0 END) AS ValorDesconto," & vbCrLf &
                   "               SUM(CASE WHEN #Encargos.Encargo_id = 'ICMS'      THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and #Encargos.STICMS in(10,60))) AND YEAR(NotasFiscais.Movimento) > 2011 then 0 else #Encargos.Base end ELSE 0 END) AS BaseICMS," & vbCrLf &
                   "               SUM(CASE WHEN #Encargos.Encargo_id = 'ICMS'      THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and #Encargos.STICMS in(10,60))) AND YEAR(NotasFiscais.Movimento) > 2011 then 0 else #Encargos.Valor end ELSE 0 END) AS ValorICMS," & vbCrLf &
                   "               SUM(CASE WHEN #Encargos.Encargo_id = 'IPI'       THEN #Encargos.Base ELSE 0 END) AS BaseIPI," & vbCrLf &
                   "               SUM(CASE WHEN #Encargos.Encargo_id = 'IPI'       THEN #Encargos.Valor ELSE 0 END) AS ValorIPI," & vbCrLf &
                   "               SUM(CASE WHEN #Encargos.Encargo_id IN('DESP.ADUANEIRAS','ICMS SUBSTITUIC') THEN #Encargos.Valor ELSE 0 END) AS OutrasDespesas," & vbCrLf &
                   "               SubOperacoes.GrupoDeContas AS CodigoContabil, C.Inscricao as InscEstadual" & vbCrLf &
                   "          FROM NotasFiscais" & vbCrLf &
                   "         INNER JOIN NotasFiscaisXItens" & vbCrLf &
                   "            ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                   "           AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                   "           AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                   "           AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                   "           AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                   "           AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                   "           AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                   "         Inner join OperacaoXEstado OE" & vbCrLf &
                   "	      on OE.Codigo_Id = NotasFiscaisXItens.OperacaoXEstado" & vbCrLf &
                   "         INNER JOIN #Encargos" & vbCrLf &
                   "           ON NotasFiscaisXItens.Empresa_Id      = #Encargos.Empresa_Id" & vbCrLf &
                   "          AND NotasFiscaisXItens.EndEmpresa_Id   = #Encargos.EndEmpresa_Id" & vbCrLf &
                   "          AND NotasFiscaisXItens.Cliente_Id      = #Encargos.Cliente_Id" & vbCrLf &
                   "          AND NotasFiscaisXItens.EndCliente_Id   = #Encargos.EndCliente_Id" & vbCrLf &
                   "          AND NotasFiscaisXItens.EntradaSaida_Id = #Encargos.EntradaSaida_Id" & vbCrLf &
                   "          AND NotasFiscaisXItens.Serie_Id        = #Encargos.Serie_Id" & vbCrLf &
                   "          AND NotasFiscaisXItens.Nota_Id         = #Encargos.Nota_Id" & vbCrLf &
                   "          AND NotasFiscaisXItens.Sequencia_Id    = #Encargos.Sequencia_Id" & vbCrLf &
                   "          AND NotasFiscaisXItens.Produto_Id      = #Encargos.Produto_Id" & vbCrLf &
                   "        INNER JOIN SubOperacoes" & vbCrLf &
                   "           ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id" & vbCrLf &
                   "          AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                   "        Inner join Situacoes S" & vbCrLf &
                   "           on S.Situacao_id = Notasfiscais.Situacao" & vbCrLf &
                   "        Inner Join Clientes C" & vbCrLf &
                   "           On C.Cliente_Id  = NotasFiscais.Cliente_Id" & vbCrLf &
                   "          and C.Endereco_Id = NotasFiscais.EndCliente_Id" & vbCrLf &
                   "        WHERE (" & vbCrLf &
                   "	            NotasFiscais.EntradaSaida_Id = 'E'" & vbCrLf &
                   "              And NotasFiscais.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf &
                   "              AND NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf &
                   "              AND NotasFiscais.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                   "              AND NotasFiscais.TipoDeDocumento NOT IN (15)" & vbCrLf &
                   "              AND NotasFiscais.Finalidade NOT IN (30)" & vbCrLf &
                   "              AND (OE.CodigoFiscal NOT IN (1933,2933))" & vbCrLf &
                   "              AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf &
                   "		    )" & vbCrLf

            If CkConsNotasCompServicos.Checked Then
                Sql &= "             OR exists (Select  1 " & vbCrLf &
                       "                          from #notascompostas " & vbCrLf &
                       "             		     where #NotasCompostas.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                       "             		       AND #NotasCompostas.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                       "             		       AND #NotasCompostas.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                       "             		       AND #NotasCompostas.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                       "             		       AND #NotasCompostas.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                       "             		       AND #NotasCompostas.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                       "             		       AND #NotasCompostas.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                       "             			)" & vbCrLf
            End If


            Sql &= "        GROUP BY NotasFiscais.Nota_Id, NotasFiscais.Serie_Id, NotasFiscais.Cliente_Id, NotasFiscais.Movimento, NotasFiscais.DataDaNota," & vbCrLf &
                   "                 OE.CodigoFiscal, SubOperacoes.GrupoDeContas, NotasFiscais.EstadoDoCliente,S.Situacao_id,S.Descricao,C.Inscricao" & vbCrLf &
                   "    ) As Consulta" & vbCrLf &
                   " Order By Movimento, Serie, NumeroDaNota" & vbCrLf


            Return Banco.ConsultaDataSet(Sql2 + Sql, "NotasFiscais")

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDatasetSaidasModeloP1() As DataSet
        Try
            Sql = " Select Cliente_Id As Cliente, Nome, Inscricao " & vbCrLf &
                          "  FROM Clientes " & vbCrLf &
                          " WHERE Cliente_Id = '" & Empresa(0) & "'" & vbCrLf &
                          "   and Endereco_Id = " & Empresa(1) & vbCrLf
            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            For Each Dr As DataRow In ds.Tables(0).Rows
                EmpresaNome = Dr("Nome")
                EmpresaCNPJ = Funcoes.FormatarCpfCnpj(Dr("Cliente"))
                EmpresaInscricao = Dr("Inscricao")
            Next

            Sql = "SELECT Movimento, '' As Especie, Serie, NumeroDaNota," & vbCrLf &
                  "       DataDaNota, Cliente, CodigoContabil," & vbCrLf &
                  "       CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE ((ValorContabil+ValorIPI+ValorContabilFrete+OutrasDespesas) - ValorDesconto) END AS ValorContabil," & vbCrLf &
                  "       CFOP," & vbCrLf &
                  "       CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE BaseICMS END  AS BaseICMS," & vbCrLf &
                  "       CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE ValorIcms END AS ValorIcms," & vbCrLf &
                  "       CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE BaseIPI END   AS BaseIPI," & vbCrLf &
                  "       CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE ValorIPI END  AS ValorIPI," & vbCrLf &
                  "       ROUND(CASE WHEN BaseIcms = 0 THEN 0 ELSE ValorIcms * 100 / BaseIcms END, 0) AS AliquotaIcms," & vbCrLf &
                  "       ROUND(CASE WHEN Baseipi = 0  THEN 0 ELSE Valoripi * 100  / Baseipi  END, 0) AS AliquotaIPI," & vbCrLf &
                  "       Estado, Consulta.Situacao, Consulta.DescSituacao, Consulta.InscEstadual" & vbCrLf &
                  " FROM (SELECT NF.Nota_Id AS NumeroDaNota, NF.Serie_Id AS Serie, NF.Cliente_Id AS Cliente, NF.Movimento," & vbCrLf &
                  "              NF.DataDaNota, OE.CodigoFiscal AS CFOP, NF.EstadoDoCliente as Estado,S.Situacao_id as Situacao,S.Descricao as DescSituacao," & vbCrLf &
                  "              SUM(CASE WHEN isnull(e.encargoagrupador, nfe.Encargo_Id) = 'PRODUTO'   THEN nfe.Valor ELSE 0 END) AS ValorContabil," & vbCrLf &
                  "              SUM(CASE WHEN isnull(e.encargoagrupador, nfe.Encargo_Id) IN ('FRETES','SEGURO') THEN nfe.Valor ELSE 0 END) AS ValorContabilFrete," & vbCrLf &
                  "              SUM(CASE WHEN isnull(e.encargoagrupador, nfe.Encargo_Id) = 'DESCONTOS' THEN nfe.Valor ELSE 0 END) AS ValorDesconto," & vbCrLf &
                  "              SUM(CASE WHEN isnull(e.encargoagrupador, nfe.Encargo_Id) = 'ICMS' AND NOT so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN nfe.Base  ELSE 0 END) AS BaseICMS," & vbCrLf &
                  "              SUM(CASE WHEN isnull(e.encargoagrupador, nfe.Encargo_Id) = 'ICMS' AND NOT so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN nfe.Valor ELSE 0 END) AS ValorICMS," & vbCrLf &
                  "              SUM(CASE WHEN isnull(e.encargoagrupador, nfe.Encargo_Id) = 'IPI'       THEN nfe.Base  ELSE 0 END) AS BaseIPI," & vbCrLf &
                  "              SUM(CASE WHEN isnull(e.encargoagrupador, nfe.Encargo_Id) = 'IPI'       THEN nfe.Valor ELSE 0 END) AS ValorIPI," & vbCrLf &
                  "              SUM(CASE WHEN isnull(e.encargoagrupador, nfe.Encargo_Id) IN('DESP.ADUANEIRAS','ICMS SUBSTITUIC') THEN nfe.Valor ELSE 0 END) AS OutrasDespesas," & vbCrLf &
                  "              so.GrupoDeContas AS CodigoContabil, C.Inscricao as InscEstadual" & vbCrLf &
                  "         FROM NotasFiscais NF" & vbCrLf &
                  "        INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                  "           ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                  "          AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                  "          AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                  "          AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                  "          AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                  "          AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                  "          AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                  "        Inner Join OperacaoXEstado OE" & vbCrLf &
                  "           on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                  "        INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
                  "           ON nfi.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
                  "          AND nfi.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
                  "          AND nfi.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
                  "          AND nfi.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
                  "          AND nfi.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
                  "          AND nfi.Serie_Id        = nfe.Serie_Id" & vbCrLf &
                  "          AND nfi.Nota_Id         = nfe.Nota_Id" & vbCrLf &
                  "          AND nfi.Produto_Id      = nfe.Produto_Id" & vbCrLf &
                  "          AND nfi.Sequencia_Id    = nfe.Sequencia_Id" & vbCrLf &
                  "         Left Join Encargos e" & vbCrLf &
                  "           on e.Encargo_id = nfe.Encargo_Id" & vbCrLf &
                  "	         and isnull(e.encargoagrupador,'') <> ''" & vbCrLf &
                  "        INNER JOIN SubOperacoes so" & vbCrLf &
                  "           ON NF.Operacao    = so.Operacao_Id" & vbCrLf &
                  "          AND NF.SubOperacao = so.SubOperacoes_Id" & vbCrLf &
                  "        Inner join Situacoes S" & vbCrLf &
                  "           ON S.Situacao_id = NF.Situacao" & vbCrLf &
                  "        INNER JOIN Clientes C" & vbCrLf &
                  "           on C.Cliente_id  = NF.Cliente_Id" & vbCrLf &
                  "          and C.Endereco_id = NF.EndCliente_Id" & vbCrLf &
                  "        WHERE NF.EntradaSaida_Id = 'S'" & vbCrLf &
                  "          AND NF.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf &
                  "          AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                  "          AND NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                  "          AND OE.CodigoFiscal NOT IN (5933,6933)" & vbCrLf &
                  "          AND NOT (so.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (5949,6949)) " & vbCrLf &
                  "	         AND isnull(e.encargoagrupador, nfe.Encargo_Id)  IN ('IPI', 'PRODUTO','FRETES','SEGURO','DESCONTOS', 'DESP.ADUANEIRAS','ICMS','ICMS SUBSTITUIC')" & vbCrLf &
                  "        GROUP BY NF.Nota_Id, NF.Serie_Id, NF.Cliente_Id, NF.Movimento, NF.DataDaNota," & vbCrLf &
                  "                 OE.CodigoFiscal, so.GrupoDeContas, NF.EstadoDoCliente,S.Situacao_id,S.Descricao,C.Inscricao" & vbCrLf &
                  "      ) AS Consulta" & vbCrLf &
                  " Order By Movimento, Serie, NumeroDaNota" & vbCrLf

            Return Banco.ConsultaDataSet(Sql, "NotasFiscais")

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDatasetSaidasModeloP1Excel() As DataSet
        Try
            Dim Sql As String = "  SELECT Movimento,                                                                                                                                    " & vbCrLf &
                                "         DataDaNota,                                                                                                                                   " & vbCrLf &
                                "		 [dbo].[FormatarCpfCnpj](Cliente) as Destinatario,                                                                                              " & vbCrLf &
                                "		 Case When CFOP IN (1252, 1253, 2252, 2253, 142, 143, 242, 243) then 'NFE'                                                                      " & vbCrLf &
                                "			  When CFOP IN (1352, 1353, 1932, 2352, 2353, 3352, 2932, 162, 163, 262, 263, 5352, 5353, 6352, 6353, 7352, 562, 563, 662, 663) then 'CTRC' " & vbCrLf &
                                "			  When CFOP IN (1302, 1303, 2302, 2303, 152, 153, 252, 253) then 'NFST'                                                                     " & vbCrLf &
                                "			  else 'NF'                                                                                                                                 " & vbCrLf &
                                "	     end as Especie,                                                                                                                                " & vbCrLf &
                                "		 NumeroDaNota, Serie, day(DataDaNota) as Dia, Estado,                                                                                           " & vbCrLf &
                                "		 CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE ((ValorContabil+ValorIPI+ValorContabilFrete+OutrasDespesas) - ValorDesconto) END AS ValorContabil," & vbCrLf &
                                "		 CodigoContabil, CFOP,                                                                                                                          " & vbCrLf &
                                "         CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE BaseICMS END  AS BaseICMS,                                                                       " & vbCrLf &
                                "         ROUND(CASE WHEN BaseIcms = 0 THEN 0 ELSE ValorIcms * 100 / BaseIcms END, 0) AS AliquotaIcms,                                                  " & vbCrLf &
                                "         CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE ValorIcms END AS ValorIcms,                                                                      " & vbCrLf &
                                "		 case when ValorContabil = 0 then ValorContabil else ValorContabil - BaseICMS end as OutrasICMS,		                                        " & vbCrLf &
                                "         CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE BaseIPI END   AS BaseIPI,                                                                        " & vbCrLf &
                                "		 ROUND(CASE WHEN Baseipi = 0  THEN 0 ELSE Valoripi * 100  / Baseipi  END, 0) AS AliquotaIPI,                                                    " & vbCrLf &
                                "         CASE WHEN Consulta.Situacao <> 1 THEN 0 ELSE ValorIPI END  AS ValorIPI,                                                                       " & vbCrLf &
                                "		 case when ValorContabil = 0 then ValorContabil else ValorContabil - BaseIPI end as OutrasIPI,                                                  " & vbCrLf &
                                "		 case when Situacao > 1 then DescSituacao else '' end as Observacao, Consulta.InscEstadual                                                      " & vbCrLf &
                                "  FROM  (SELECT NotasFiscais.Nota_Id AS NumeroDaNota, NotasFiscais.Serie_Id AS Serie, NotasFiscais.Cliente_Id AS Cliente, NotasFiscais.Movimento," & vbCrLf &
                                "                NotasFiscais.DataDaNota, NotasFiscaisXItens.CFOP_Id AS CFOP, NotasFiscais.EstadoDoCliente as Estado,S.Situacao_id as Situacao,S.Descricao as DescSituacao," & vbCrLf &
                                "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'PRODUTO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END)   AS ValorContabil," & vbCrLf &
                                "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id IN ('FRETES','SEGURO') THEN NotasFiscaisXEncargos.Valor ELSE 0 END)    AS ValorContabilFrete," & vbCrLf &
                                "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'DESCONTOS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorDesconto," & vbCrLf &
                                "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id like '%ICMS%' AND NOT SubOperacaoes.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS," & vbCrLf &
                                "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id like '%ICMS%' AND NOT SubOperacaoes.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS," & vbCrLf &
                                "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Base ELSE 0 END)        AS BaseIPI," & vbCrLf &
                                "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END)       AS ValorIPI," & vbCrLf &
                                "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id IN ('DESP.ADUANEIRAS','ICMS SUBSTITUIC') THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS OutrasDespesas," & vbCrLf &
                                "                SubOperacoes.GrupoDeContas AS CodigoContabil, C.Inscricao as InscEstadual" & vbCrLf &
                                "           FROM NotasFiscais" & vbCrLf &
                                "          INNER JOIN NotasFiscaisXItens" & vbCrLf &
                                "             ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                                "            AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                                "            AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                                "            AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                                "            AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                                "            AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                                "            AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                                "          INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
                                "             ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
                                "            AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
                                "            AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
                                "            AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
                                "            AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
                                "            AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
                                "            AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
                                "            AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
                                "            AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_Id" & vbCrLf &
                                "            AND (NotasFiscaisXEncargos.Encargo_Id IN ('IPI', 'PRODUTO','FRETES','SEGURO','DESCONTOS', 'DESP.ADUANEIRAS') or  NotasFiscaisXEncargos.Encargo_Id like '%ICMS%')" & vbCrLf &
                                "          INNER JOIN SubOperacoes" & vbCrLf &
                                "             ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                                "            AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                                "          Inner join Situacoes S" & vbCrLf &
                                "             ON S.Situacao_id = Notasfiscais.Situacao" & vbCrLf &
                                "          INNER JOIN Clientes C" & vbCrLf &
                                "             on C.Cliente_id  = NotasFiscais.Cliente_Id" & vbCrLf &
                                "            and C.Endereco_id = NotasFiscais.EndCliente_Id" & vbCrLf &
                                "          WHERE  NotasFiscais.EntradaSaida_Id = 'S'" & vbCrLf &
                                "            And  NotasFiscais.Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                                "            AND  NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                                "            AND  NotasFiscais.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                                "            AND (NotasFiscaisXItens.CFOP_Id NOT IN (5933,6933))" & vbCrLf &
                                "            AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND NotasFiscaisXItens.CFOP_Id IN (5949,6949)) " & vbCrLf &
                                "          GROUP BY NotasFiscais.Nota_Id, NotasFiscais.Serie_Id, NotasFiscais.Cliente_Id, NotasFiscais.Movimento, NotasFiscais.DataDaNota, " & vbCrLf &
                                "                   NotasFiscaisXItens.CFOP_Id, SubOperacoes.GrupoDeContas, NotasFiscais.EstadoDoCliente,S.Situacao_id,S.Descricao,C.Inscricao) AS Consulta" & vbCrLf &
                                "          Order By Movimento, Serie, NumeroDaNota" & vbCrLf
            Return Banco.ConsultaDataSet(Sql, "NotasFiscais")

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Sub RegistroDeEntradasModeloP1()
        Try
            Empresa = DdlEmpresa.SelectedValue.Split("-")
            PeriodoInicial = CDate(txtDataInicial.Text).ToString("yyyy/MM/dd")
            PeriodoFinal = CDate(txtDataFinal.Text).ToString("yyyy/MM/dd")
            Processo = DdlProcesso.SelectedValue
            Livro = txtLivro.Text
            Folha = txtFolha.Text
            Vencimento = CDate(txtVenctoDaObrigacao.Text).ToString("yyyy/MM/dd")
            CodigoDaReceita = txtCodigoDaReceita.Text

            If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0), Vencimento, CodigoDaReceita) Then

                Dim Periodo As String = "Mês ou Período/Ano : " & PeriodoInicial.ToString("dd/MM/yyyy") & " a " & PeriodoFinal.ToString("dd/MM/yyyy")
                Dim ds As DataSet = getDatasetEntradasModeloP1()

                If CkExcell.Checked Then
                    Dim colunas As New Dictionary(Of String, eTipoCampo)
                    colunas.Add("Movimento", eTipoCampo.Data)
                    colunas.Add("DataDaNota", eTipoCampo.Data)
                    colunas.Add("ValorContabil", eTipoCampo.ValorComTotalizador)
                    colunas.Add("BaseICMS", eTipoCampo.ValorComTotalizador)
                    colunas.Add("AliquotaIcms", eTipoCampo.ValorComTotalizador)
                    colunas.Add("ValorIcms", eTipoCampo.ValorComTotalizador)
                    colunas.Add("OutrasIcms", eTipoCampo.ValorComTotalizador)
                    colunas.Add("BaseIPI", eTipoCampo.ValorComTotalizador)
                    colunas.Add("AliquotaIPI", eTipoCampo.ValorComTotalizador)
                    colunas.Add("ValorIPI", eTipoCampo.ValorComTotalizador)
                    colunas.Add("OutrasIPI", eTipoCampo.ValorComTotalizador)

                    Dim caecalho As String = ""
                    caecalho = "Livro Registro de Entradas" & " - Firma : " & EmpresaNome & " - Inscrição Estadual :" & EmpresaInscricao & " - Livro :" & Format(CInt(Livro), "000")

                    Funcoes.BindExcelOffice(Me.Page, ds, "Livro Registro de Entradas", colunas, False, caecalho)
                Else
                    Dim param As New Dictionary(Of String, Object)
                    param.Add("EmpresaNome", EmpresaNome)
                    param.Add("EmpresaCNPJ", EmpresaCNPJ)
                    param.Add("EmpresaInscricao", EmpresaInscricao)
                    param.Add("Periodo", Periodo)
                    param.Add("Livro", "Livro.: " & Format(CInt(Livro), "000"))
                    param.Add("Pagina", Folha)

                    If chkZIP.Checked OrElse chkImpressao.Checked OrElse chkAbrir.Checked Then
                        impressaoPDF(ds, "E", "01")
                    End If

                    Funcoes.BindReport(Me.Page, ds, "Cr_RegistrodeEntradasModeloP1", eExportType.PDF, param, True)

                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Registros Fiscais de Saidas Modelo P1"

    Private Sub RegistroDeSaidasModeloP1()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        PeriodoInicial = CDate(txtDataInicial.Text).ToString("yyyy/MM/dd")
        PeriodoFinal = CDate(txtDataFinal.Text).ToString("yyyy/MM/dd")
        Processo = DdlProcesso.SelectedValue
        Livro = txtLivro.Text
        Folha = txtFolha.Text
        Vencimento = CDate(txtVenctoDaObrigacao.Text).ToString("yyyy/MM/dd")
        CodigoDaReceita = txtCodigoDaReceita.Text

        Dim Periodo As String = "Mês ou Período/Ano : " & PeriodoInicial.ToString("dd/MM/yyyy") & " a " & PeriodoFinal.ToString("dd/MM/yyyy")

        If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0), Vencimento, CodigoDaReceita) Then
            Try
                Dim ds As New DataSet
                If Not CkExcell.Checked Then
                    ds = getDatasetSaidasModeloP1()

                    Dim param As New Dictionary(Of String, Object)
                    param.Add("EmpresaNome", EmpresaNome)
                    param.Add("EmpresaCNPJ", EmpresaCNPJ)
                    param.Add("EmpresaInscricao", EmpresaInscricao)
                    param.Add("Periodo", Periodo)
                    param.Add("Livro", "Livro.: " & Format(CInt(Livro), "000"))
                    param.Add("Pagina", Folha)

                    If chkZIP.Checked OrElse chkImpressao.Checked OrElse chkAbrir.Checked Then
                        impressaoPDF(ds, "S", "02")
                    End If

                    Funcoes.BindReport(Me.Page, ds, "Cr_RegistroDeSaidasModeloP2", eExportType.PDF, param, True)

                Else

                    ds = getDatasetSaidasModeloP1Excel()
                    Dim colunas As New Dictionary(Of String, eTipoCampo)
                    colunas.Add("Movimento", eTipoCampo.Data)
                    colunas.Add("DataDaNota", eTipoCampo.Data)
                    colunas.Add("ValorContabil", eTipoCampo.ValorComTotalizador)
                    colunas.Add("BaseICMS", eTipoCampo.ValorComTotalizador)
                    colunas.Add("AliquotaIcms", eTipoCampo.ValorComTotalizador)
                    colunas.Add("ValorIcms", eTipoCampo.ValorComTotalizador)
                    colunas.Add("OutrasIcms", eTipoCampo.ValorComTotalizador)
                    colunas.Add("BaseIPI", eTipoCampo.ValorComTotalizador)
                    colunas.Add("AliquotaIPI", eTipoCampo.ValorComTotalizador)
                    colunas.Add("ValorIPI", eTipoCampo.ValorComTotalizador)
                    colunas.Add("OutrasIPI", eTipoCampo.ValorComTotalizador)

                    Funcoes.BindExcelOffice(Me.Page, ds, "Registros de Saídas", colunas, False, "Livro Registro de Saidas - RS - Modelo P2")

                End If


            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        End If
    End Sub

#End Region

#Region "Registro De Apuração De ICMS Modelo P9"

    Private Sub RegistroDeApuracaoDeIcmsModeloP9(ByVal processarRegistros As Boolean)
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        PeriodoInicial = CDate(txtDataInicial.Text).ToString("yyyy/MM/dd")
        PeriodoFinal = CDate(txtDataFinal.Text).ToString("yyyy/MM/dd")
        Processo = DdlProcesso.SelectedValue
        Livro = txtLivro.Text
        Folha = txtFolha.Text
        Vencimento = CDate(txtVenctoDaObrigacao.Text).ToString("yyyy/MM/dd")
        CodigoDaReceita = txtCodigoDaReceita.Text

        'CRIADO PARA LIBERAR APENAS A IMPRESSÃO - FURLAN - 06/11/2024
        If processarRegistros Then
            If Not Funcoes.VerificaAcesso(Empresa(0), Empresa(1), txtDataFinal.Text, "NOTAS FISCAIS") Then
                MsgBox(Me.Page, "Movimento Fiscal já Fechado para esta data")
                Exit Sub
            End If

            If Not Funcoes.VerificaPermissao("RegistrosDeICMS", "GRAVAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para gerar Apuração de Icms.")
                Exit Sub
            End If

        End If

        If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0), Vencimento, CodigoDaReceita) Then
            Try
                Sql = " SELECT Cliente_Id AS Cliente, Nome, Inscricao " & vbCrLf &
                      "  FROM Clientes " & vbCrLf &
                      " WHERE Cliente_Id = '" & Empresa(0) & "'" & vbCrLf &
                      "   and Endereco_Id = " & Empresa(1) & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                For Each Dr As DataRow In ds.Tables(0).Rows
                    EmpresaNome = Dr("Nome")
                    EmpresaCNPJ = Funcoes.FormatarCpfCnpj(Dr("Cliente"))
                    EmpresaInscricao = Dr("Inscricao")
                Next

                Dim Periodo As String = "Mês ou Período/Ano : " & CDate(PeriodoInicial).ToString("dd/MM/yyyy") & " a " & CDate(PeriodoFinal).ToString("dd/MM/yyyy")
                Dim TemNotasCompostasDeServico As Boolean

                'Temporaria Controla Notas de Serviço Conjugadas
                If CkConsNotasCompServicos.Checked Then
                    Sql2 = "Select nfxi.Empresa_Id," & vbCrLf &
                           "       nfxi.EndEmpresa_Id," & vbCrLf &
                           "       nfxi.Cliente_Id, nfxi.EndCliente_Id," & vbCrLf &
                           "	   nfxi.EntradaSaida_Id, nfxi.Serie_Id, nfxi.Nota_Id," & vbCrLf &
                           "       sum(Case" & vbCrLf &
                           "	          when OEx.CodigoFiscal in (1933,2933)" & vbCrLf &
                           "               then 1" & vbCrLf &
                           "               else 0" & vbCrLf &
                           "           end ) ItensCfop," & vbCrLf &
                           "       sum(1) nrItens" & vbCrLf &
                           "  into #NotasCompostas" & vbCrLf &
                           "  From notasfiscaisxitens nfxi" & vbCrLf &
                           " Inner Join OperacaoXEstado OEx" & vbCrLf &
                           "    on OEx.Codigo_Id = nfxi.OperacaoXEstado" & vbCrLf &
                           " Inner join (SELECT NF.Empresa_Id, NF.EndEmpresa_Id," & vbCrLf &
                           "                    NF.Cliente_Id, NF.EndCliente_Id," & vbCrLf &
                           "                    NF.EntradaSaida_Id, NF.Serie_Id, NF.Nota_Id" & vbCrLf &
                           "               FROM NotasFiscais NF" & vbCrLf &
                           "              INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                           "                 ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                           "                AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                           "                AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                           "                AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                           "                AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                           "                AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                           "                AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                           "              Inner Join OperacaoXEstado OE" & vbCrLf &
                           "                 on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                           "              Inner Join Suboperacoes so" & vbCrLf &
                           "                 on so.Operacao_id     = nfi.Operacao" & vbCrLf &
                           "                and so.SubOperacoes_id = nfi.suboperacao" & vbCrLf &
                           "              WHERE NF.EntradaSaida_Id = 'E'" & vbCrLf &
                           "                And NF.Empresa_Id ='" & Empresa(0) & "' " & vbCrLf &
                           "                AND NF.TipoDeDocumento NOT IN (15)" & vbCrLf &
                           "                AND NF.Finalidade NOT IN (30)" & vbCrLf &
                           "                AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                           "                AND NF.Serie_Id  NOT IN ('D', 'F', 'REC')" & vbCrLf &
                           "                AND(nfi.CFOP_Id  IN (1933,2933))" & vbCrLf &
                           "              )sb" & vbCrLf &
                           "    ON sb.Empresa_Id      = nfxi.Empresa_Id" & vbCrLf &
                           "   AND sb.EndEmpresa_Id   = nfxi.EndEmpresa_Id" & vbCrLf &
                           "   AND sb.Cliente_Id      = nfxi.Cliente_Id" & vbCrLf &
                           "   AND sb.EndCliente_Id   = nfxi.EndCliente_Id" & vbCrLf &
                           "   AND sb.EntradaSaida_Id = nfxi.EntradaSaida_Id" & vbCrLf &
                           "   AND sb.Serie_Id        = nfxi.Serie_Id" & vbCrLf &
                           "   AND sb.Nota_Id         = nfxi.Nota_Id" & vbCrLf &
                           " group by nfxi.Empresa_Id, nfxi.EndEmpresa_Id," & vbCrLf &
                           "          nfxi.Cliente_Id, nfxi.EndCliente_Id," & vbCrLf &
                           "          nfxi.EntradaSaida_Id, nfxi.Serie_Id, nfxi.Nota_Id" & vbCrLf

                    If PeriodoFinal > "2011-12-31" Then
                        Sql2 &= "having sum(1)  <> sum(Case" & vbCrLf &
                                "                        when OEx.CodigoFiscal in (1933,2933)" & vbCrLf &
                                "                          then 1" & vbCrLf &
                                "                          else 0" & vbCrLf &
                                "                      end)" & vbCrLf
                    End If


                    Sql = " select * from #NotasCompostas" & vbCrLf

                End If

                'Seta essa variável TemNotasCompostasDeServico para usá-la nos testes com or dessa maneira caso não tenha notas compostas ele não precisará fazer o OR na consulta sql
                TemNotasCompostasDeServico = Banco.ConsultaDataSet(Sql2 + Sql, "NotasCompostas").Tables(0).Rows.Count > 0

                Sql = "SELECT Consulta.CFOP," & vbCrLf &
                      "       Cfop.Descricao," & vbCrLf &
                      "       ((Consulta.Contabil + Consulta.OutrasDespesas + Consulta.ValorContabilFrete + Consulta.ValorIPI) - Consulta.ValorDesconto) AS Contabil," & vbCrLf &
                      "       Consulta.BaseICMS," & vbCrLf &
                      "       Consulta.ValorICMS," & vbCrLf &
                      "       0 AS ValorIPI" & vbCrLf &
                      " FROM  (SELECT OE.CodigoFiscal AS CFOP," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'PRODUTO'   THEN nfe.Valor ELSE 0 END) AS Contabil," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'DESCONTOS' THEN nfe.Valor ELSE 0 END) AS ValorDesconto," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) IN ('FRETES','SEGURO') THEN nfe.Valor ELSE 0 END) AS ValorContabilFrete," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'IPI'       THEN nfe.Valor ELSE 0 END) AS ValorIPI," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) IN('DESP.ADUANEIRAS','ICMS SUBSTITUIC') THEN nfe.Valor ELSE 0 END) AS OutrasDespesas," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS' THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and OE.STICMS in(10,60))) AND YEAR(NF.Movimento) > 2011 then 0 else nfe.Base end ELSE 0 END) AS BaseICMS," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS' THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and OE.STICMS in(10,60))) AND YEAR(NF.Movimento) > 2011 then 0 else nfe.Valor end ELSE 0 END) AS ValorICMS" & vbCrLf &
                      "          FROM NotasFiscais NF" & vbCrLf &
                      "		 INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                      "		    ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                      "		   AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                      "		   AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                      "		   AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                      "		   AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                      "		   AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                      "		   AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                      "         Inner Join OperacaoXEstado OE" & vbCrLf &
                      "		    on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                      "		 INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
                      "		    ON nfi.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
                      "		   AND nfi.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
                      "		   AND nfi.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
                      "		   AND nfi.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
                      "		   AND nfi.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
                      "		   AND nfi.Serie_Id        = nfe.Serie_Id" & vbCrLf &
                      "		   AND nfi.Nota_Id         = nfe.Nota_Id" & vbCrLf &
                      "		   AND nfi.Produto_Id      = nfe.Produto_Id" & vbCrLf &
                      "		   AND nfi.Sequencia_Id    = nfe.Sequencia_Id" & vbCrLf &
                      "          Left Join Encargos E" & vbCrLf &
                      "		    on E.Encargo_id = nfe.Encargo_Id" & vbCrLf &
                      "           and isnull(e.EncargoAgrupador,'') <> ''" & vbCrLf &
                      "         INNER JOIN SubOperacoes SO" & vbCrLf &
                      "            ON NF.Operacao    = SO.Operacao_Id" & vbCrLf &
                      "           AND NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
                      "         WHERE (" & vbCrLf &
                      "		        (" & vbCrLf &
                      "		             NF.EntradaSaida_Id = 'E'" & vbCrLf &
                      "                 And NF.Situacao = 1" & vbCrLf &
                      "                 AND NF.TipoDeDocumento NOT IN (15)" & vbCrLf &
                      "                 AND NF.Finalidade NOT IN (30)" & vbCrLf &
                      "                 And NF.Empresa_Id ='" & Empresa(0) & "' " & vbCrLf &
                      "                 AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                      "                 AND NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                      "                 AND nfi.CFOP_Id NOT IN (1933,2933)" & vbCrLf &
                      "                 AND NOT (SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND nfi.CFOP_Id IN (1949,2949))" & vbCrLf &
                      "                )" & vbCrLf

                If CkConsNotasCompServicos.Checked AndAlso TemNotasCompostasDeServico Then
                    Sql &= "               OR" & vbCrLf &
                           "			   exists (Select  1" & vbCrLf &
                           "		                from #notascompostas" & vbCrLf &
                           "					   Where #NotasCompostas.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                           "                         AND #NotasCompostas.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                           "                         AND #NotasCompostas.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                           "                         AND #NotasCompostas.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                           "                         AND #NotasCompostas.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                           "                         AND #NotasCompostas.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                           "                         AND #NotasCompostas.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                           "                       )" & vbCrLf
                End If

                Sql &= "               )" & vbCrLf &
                        "		   AND isnull(e.encargoagrupador,nfe.Encargo_Id) IN ('IPI', 'PRODUTO','DESCONTOS','FRETES','SEGURO','DESP.ADUANEIRAS','ICMS','ICMS SUBSTITUIC')" & vbCrLf &
                        "     Group BY OE.CodigoFiscal" & vbCrLf &
                        "  ) AS Consulta" & vbCrLf &
                        "  INNER JOIN Cfop" & vbCrLf &
                        "     ON Consulta.CFOP = Cfop.Cfop_Id" & vbCrLf &
                        " GROUP BY Consulta.CFOP, Consulta.Contabil, Consulta.OutrasDespesas,Consulta.ValorContabilFrete,Consulta.ValorDesconto, Consulta.BaseICMS, Consulta.ValorICMS, Consulta.ValorIPI, Cfop.Descricao" & vbCrLf &
                        " Order by Consulta.CFOP" & vbCrLf


                If TemNotasCompostasDeServico Then
                    ds.Merge(Banco.ConsultaDataSet(Sql2 + Sql, "ApuracaoIcmsEntradas"))
                Else
                    ds.Merge(Banco.ConsultaDataSet(Sql, "ApuracaoIcmsEntradas"))
                End If

                Sql = "SELECT (CFOP + '.000') as CFOP," & vbCrLf &
                      "       CASE" & vbCrLf &
                      "          WHEN CFOP = 1 THEN 'Do Estado'" & vbCrLf &
                      "          WHEN CFOP = 2 THEN 'De Outros Estados'" & vbCrLf &
                      "          ELSE 'Do Exterior'" & vbCrLf &
                      "        END AS Descricao," & vbCrLf &
                      "		   ((sum(Contabil)+Sum(OutrasDespesas)+Sum(ValorContabilFrete) + sum(+ValorIPI))- Sum(ValorDesconto)) as Contabil," & vbCrLf &
                      "		   Sum(BaseICMS) as BaseIcms," & vbCrLf &
                      "	       Sum(ValorICMS) as ValorIcms," & vbCrLf &
                      "	       0 AS ValorIPI" & vbCrLf &
                      " FROM  (SELECT LEFT(OE.CodigoFiscal, 1) AS CFOP," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'PRODUTO'  THEN nfe.Valor ELSE 0 END) AS Contabil," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'DESCONTOS'THEN nfe.Valor ELSE 0 END) AS ValorDesconto," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS'     THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and oe.STICMS in(10,60))) AND YEAR(nf.Movimento) > 2011 then 0 else nfe.Base end ELSE 0 END)  AS BaseICMS," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS'     THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and oe.STICMS in(10,60))) AND YEAR(nf.Movimento) > 2011 then 0 else nfe.Valor end ELSE 0 END) AS ValorICMS," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) IN ('FRETES','SEGURO') THEN nfe.Valor ELSE 0 END) AS ValorContabilFrete," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'IPI'      THEN nfe.Valor ELSE 0 END) AS ValorIPI," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) IN('DESP.ADUANEIRAS','ICMS SUBSTITUIC') THEN nfe.Valor ELSE 0 END) AS OutrasDespesas" & vbCrLf &
                      "          FROM NotasFiscais NF" & vbCrLf &
                      "		 INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                      "		    ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                      "		   AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                      "		   AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                      "		   AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                      "		   AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                      "		   AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                      "		   AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                      "         Inner join OperacaoXEstado OE" & vbCrLf &
                      "		    on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                      "		 INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
                      "		    ON nfi.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
                      "		   AND nfi.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
                      "		   AND nfi.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
                      "		   AND nfi.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
                      "		   AND nfi.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
                      "		   AND nfi.Serie_Id        = nfe.Serie_Id" & vbCrLf &
                      "		   AND nfi.Nota_Id         = nfe.Nota_Id" & vbCrLf &
                      "		   AND nfi.Produto_Id      = nfe.Produto_Id" & vbCrLf &
                      "		   AND nfi.Sequencia_Id    = nfe.Sequencia_Id" & vbCrLf &
                      "		  Left Join Encargos e" & vbCrLf &
                      "		    on e.Encargo_id  = nfe.Encargo_Id" & vbCrLf &
                      "           and isnull(e.encargoagrupador,'') <> ''" & vbCrLf &
                      "         INNER JOIN SubOperacoes" & vbCrLf &
                      "            ON NF.Operacao = SubOperacoes.Operacao_Id" & vbCrLf &
                      "           AND NF.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                      "         WHERE (" & vbCrLf &
                      "		       (" & vbCrLf &
                      "		           NF.EntradaSaida_Id = 'E'" & vbCrLf &
                      "               And NF.Situacao = 1" & vbCrLf &
                      "               AND NF.TipoDeDocumento NOT IN (15)" & vbCrLf &
                      "               AND NF.Finalidade NOT IN (30)" & vbCrLf &
                      "               And NF.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf &
                      "               AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                      "               AND NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                      "               AND nfi.CFOP_Id NOT IN (1933,2933)" & vbCrLf &
                      "               AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND nfi.CFOP_Id IN (1949,2949))" & vbCrLf &
                      "			   )" & vbCrLf

                If CkConsNotasCompServicos.Checked AndAlso TemNotasCompostasDeServico Then
                    Sql &= "              OR exists (Select  1" & vbCrLf &
                           "			               from #notascompostas" & vbCrLf &
                           "			  			  where #NotasCompostas.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                           "                         AND #NotasCompostas.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                           "                         AND #NotasCompostas.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                           "                         AND #NotasCompostas.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                           "                         AND #NotasCompostas.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                           "                         AND #NotasCompostas.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                           "                         AND #NotasCompostas.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                           "                         )" & vbCrLf
                End If

                Sql &= "			   )" & vbCrLf &
                       "		   AND isnull(e.encargoagrupador,nfe.Encargo_Id) IN ('IPI','PRODUTO','DESCONTOS','FRETES','SEGURO','DESP.ADUANEIRAS','ICMS','ICMS SUBSTITUIC')" & vbCrLf &
                       "      GROUP BY OE.CodigoFiscal" & vbCrLf &
                       "	 ) AS Consulta" & vbCrLf &
                       " GROUP BY CFOP" & vbCrLf

                If TemNotasCompostasDeServico Then
                    ds.Merge(Banco.ConsultaDataSet(Sql2 + Sql, "ApuracaoIcmsEntradasTotais"))
                Else
                    ds.Merge(Banco.ConsultaDataSet(Sql, "ApuracaoIcmsEntradasTotais"))
                End If


                Sql = "SELECT Consulta.CFOP," & vbCrLf &
                      "       Cfop.Descricao," & vbCrLf &
                      "       ((Consulta.Contabil+Consulta.ValorIPI+Consulta.ValorContabilFrete+Consulta.OutrasDespesas) - Consulta.ValorDesconto) AS Contabil," & vbCrLf &
                      "       Consulta.BaseICMS," & vbCrLf &
                      "       Consulta.ValorICMS," & vbCrLf &
                      "       0 AS ValorIPI" & vbCrLf &
                      "  FROM (SELECT OE.CodigoFiscal AS CFOP," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'PRODUTO'   THEN nfe.Valor ELSE 0 END) AS Contabil," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'DESCONTOS' THEN nfe.Valor ELSE 0 END) AS ValorDesconto," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) IN ('FRETES','SEGURO') THEN nfe.Valor ELSE 0 END) AS ValorContabilFrete," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'IPI'       THEN nfe.Valor ELSE 0 END) AS ValorIPI," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) IN ('DESP.ADUANEIRAS','ICMS SUBSTITUIC') THEN nfe.Valor ELSE 0 END) AS OutrasDespesas," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS' AND NOT so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN nfe.Base  ELSE 0 END) AS BaseICMS," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS' AND NOT so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN nfe.Valor ELSE 0 END) AS ValorICMS" & vbCrLf &
                      "          FROM NotasFiscais NF" & vbCrLf &
                      "         INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                      "            ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                      "           AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                      "           AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                      "           AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                      "           AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                      "           AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                      "           AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                      "         Inner join OperacaoXEstado OE" & vbCrLf &
                      "            on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                      "         INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
                      "            ON nfi.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
                      "           AND nfi.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
                      "           AND nfi.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
                      "           AND nfi.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
                      "           AND nfi.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
                      "           AND nfi.Nota_Id         = nfe.Nota_Id" & vbCrLf &
                      "           AND nfi.Serie_Id        = nfe.Serie_Id" & vbCrLf &
                      "           AND nfi.Produto_Id      = nfe.Produto_Id" & vbCrLf &
                      "           AND nfi.Sequencia_Id    = nfe.Sequencia_Id" & vbCrLf &
                      "          LEFT join Encargos E" & vbCrLf &
                      "            on e.Encargo_id = nfe.Encargo_Id" & vbCrLf &
                      "           and isnull(e.EncargoAgrupador,'') <> ''" & vbCrLf &
                      "         INNER JOIN SubOperacoes so" & vbCrLf &
                      "            ON NF.Operacao    = so.Operacao_Id" & vbCrLf &
                      "           AND NF.SubOperacao = so.SubOperacoes_Id" & vbCrLf &
                      "         WHERE NF.EntradaSaida_Id = 'S'" & vbCrLf &
                      "           AND NF.Situacao   = 1" & vbCrLf &
                      "           AND NF.Empresa_Id ='" & Empresa(0) & "' " & vbCrLf &
                      "           AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                      "           AND NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                      "           AND (OE.CodigoFiscal NOT IN (5933,6933))" & vbCrLf &
                      "           AND NOT (so.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (5949,6949))" & vbCrLf &
                      "           AND isnull(e.EncargoAgrupador,nfe.Encargo_Id) IN ('IPI','PRODUTO','FRETES','SEGURO','DESCONTOS','DESP.ADUANEIRAS', 'ICMS','ICMS SUBSTITUIC')" & vbCrLf &
                      "         Group BY OE.CodigoFiscal" & vbCrLf &
                      "      ) AS Consulta" & vbCrLf &
                      " INNER JOIN Cfop" & vbCrLf &
                      "    ON Consulta.CFOP = Cfop.Cfop_Id" & vbCrLf &
                      " GROUP BY Consulta.CFOP, Consulta.Contabil, Consulta.ValorIPI,Consulta.ValorContabilFrete,Consulta.OutrasDespesas, Consulta.ValorDesconto, Consulta.BaseICMS, Consulta.ValorICMS, Cfop.Descricao" & vbCrLf &
                      " Order by Consulta.CFOP" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql, "ApuracaoIcmsSaidas"))


                Sql = "SELECT (CFOP + '.000') as CFOP, CASE" & vbCrLf &
                      "                                  WHEN CFOP = 5 THEN 'P/ O Estado'" & vbCrLf &
                      "                                  WHEN CFOP = 6 THEN 'P/ Outros Estados'" & vbCrLf &
                      "                                  ELSE 'P/ O Exterior'" & vbCrLf &
                      "                                END AS Descricao," & vbCrLf &
                      "       ((Sum(Contabil)+sum(ValorIPI)+sum(ValorContabilFrete)+sum(OutrasDespesas)) - Sum(ValorDesconto)) as Contabil," & vbCrLf &
                      "       Sum(BaseICMS) as BaseIcms," & vbCrLf &
                      "       Sum(ValorICMS) as ValorIcms," & vbCrLf &
                      "       0 AS ValorIPI" & vbCrLf &
                      "  FROM (SELECT LEFT(OE.CodigoFiscal, 1) AS CFOP," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'PRODUTO'   THEN nfe.Valor ELSE 0 END) AS Contabil," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'DESCONTOS' THEN nfe.Valor ELSE 0 END) AS ValorDesconto," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) IN ('FRETES','SEGURO') THEN nfe.Valor ELSE 0 END) AS ValorContabilFrete," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'IPI'       THEN nfe.Valor ELSE 0 END) AS ValorIPI," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) IN ('DESP.ADUANEIRAS','ICMS SUBSTITUIC') THEN nfe.Valor ELSE 0 END) AS OutrasDespesas," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS' AND NOT so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN nfe.Base  ELSE 0 END) AS BaseICMS," & vbCrLf &
                      "               SUM(CASE WHEN isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS' AND NOT so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN nfe.Valor ELSE 0 END) AS ValorICMS" & vbCrLf &
                      "          FROM NotasFiscais NF" & vbCrLf &
                      "         INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                      "            ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                      "           AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                      "           AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                      "           AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                      "           AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                      "           AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                      "           AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                      "         Inner join OperacaoXEstado OE" & vbCrLf &
                      "            on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                      "         INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
                      "            ON nfi.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
                      "           AND nfi.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
                      "           AND nfi.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
                      "           AND nfi.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
                      "           AND nfi.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
                      "           AND nfi.Serie_Id        = nfe.Serie_Id" & vbCrLf &
                      "           AND nfi.Nota_Id         = nfe.Nota_Id" & vbCrLf &
                      "           AND nfi.Produto_Id      = nfe.Produto_Id" & vbCrLf &
                      "           AND nfi.Sequencia_Id    = nfe.Sequencia_Id" & vbCrLf &
                      "          Left Join Encargos E" & vbCrLf &
                      "            on E.Encargo_id = nfe.Encargo_Id" & vbCrLf &
                      "           and isnull(e.encargoagrupador,'') <> ''" & vbCrLf &
                      "         INNER JOIN SubOperacoes so" & vbCrLf &
                      "            ON NF.Operacao    = so.Operacao_Id" & vbCrLf &
                      "           AND NF.SubOperacao = so.SubOperacoes_Id" & vbCrLf &
                      "         WHERE NF.EntradaSaida_Id = 'S'" & vbCrLf &
                      "           And NF.Situacao = 1" & vbCrLf &
                      "           And NF.Empresa_Id ='" & Empresa(0) & "' " & vbCrLf &
                      "           AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                      "           AND NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                      "           AND (OE.CodigoFiscal NOT IN (5933,6933))" & vbCrLf &
                      "           AND NOT (so.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (5949,6949))" & vbCrLf &
                      "           AND isnull(e.encargoagrupador,nfe.Encargo_Id) IN ('IPI','PRODUTO','FRETES','SEGURO','DESCONTOS','DESP.ADUANEIRAS', 'ICMS', 'ICMS SUBSTITUIC')" & vbCrLf &
                      "         GROUP BY LEFT(OE.CodigoFiscal, 1)" & vbCrLf &
                      "    ) AS Consulta" & vbCrLf &
                      "GROUP BY CFOP" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql, "ApuracaoIcmsSaidasTotais"))

                ' Atualiza Debito do Imposto 
                Sql = " SELECT isnull(SUM(ValorICMS), 0) AS ValorIcms" & vbCrLf &
                      "   FROM (SELECT LEFT(OE.CodigoFiscal, 1) AS CFOP," & vbCrLf &
                      "                SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'PRODUTO' THEN nfe.Valor ELSE 0 END) AS Contabil," & vbCrLf &
                      "                SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS' AND NOT so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN nfe.Base ELSE 0 END) AS BaseICMS," & vbCrLf &
                      "                SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS' AND NOT so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN nfe.Valor ELSE 0 END) AS ValorICMS" & vbCrLf &
                      "           FROM NotasFiscais NF" & vbCrLf &
                      "          INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                      "		     ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                      "            AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                      "            AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                      "            AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                      "            AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                      "            AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                      "            AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                      "          INNER JOIN OperacaoXEstado OE" & vbCrLf &
                      "		     on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                      "          INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
                      "             ON nfi.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
                      "            AND nfi.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
                      "            AND nfi.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
                      "            AND nfi.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
                      "            AND nfi.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
                      "            AND nfi.Serie_Id        = nfe.Serie_Id" & vbCrLf &
                      "            AND nfi.Nota_Id         = nfe.Nota_Id" & vbCrLf &
                      "            AND nfi.Produto_Id      = nfe.Produto_Id" & vbCrLf &
                      "            AND nfi.Sequencia_Id    = nfe.Sequencia_Id" & vbCrLf &
                      "		   LEFT Join Encargos E" & vbCrLf &
                      "		     on E.Encargo_id = nfe.Encargo_Id" & vbCrLf &
                      "            and isnull(e.EncargoAgrupador,'') <> ''" & vbCrLf &
                      "          INNER JOIN SubOperacoes SO" & vbCrLf &
                      "             ON NF.Operacao    = SO.Operacao_Id" & vbCrLf &
                      "            AND NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
                      "          WHERE NF.EntradaSaida_Id = 'S'" & vbCrLf &
                      "            And NF.Situacao = 1" & vbCrLf &
                      "            And NF.Empresa_Id ='" & Empresa(0) & "' " & vbCrLf &
                      "            AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                      "            AND NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                      "            AND (OE.CodigoFiscal NOT IN (5933,6933))" & vbCrLf &
                      "            AND NOT (SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (5949,6949))" & vbCrLf &
                      "			AND isnull(e.EncargoAgrupador,nfe.Encargo_Id) IN ('IPI', 'PRODUTO', 'ICMS')" & vbCrLf &
                      "          GROUP BY LEFT(OE.CodigoFiscal, 1)" & vbCrLf &
                      "		) AS Consulta" & vbCrLf

                dss = Banco.ConsultaDataSet(Sql, "RegistrosFiscais")

                If processarRegistros AndAlso dss.Tables(0).Rows.Count > 0 Then
                    'For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                    For Each Dr As DataRow In dss.Tables(0).Rows
                        Sql = "UPDATE ResumoRAICMS " & vbCrLf &
                              " SET Valor = " & Microsoft.VisualBasic.Replace(Dr("ValorIcms"), ",", ".") & vbCrLf &
                              " WHERE Empresa_Id = '" & Empresa(0) & "' " & vbCrLf &
                              " AND Processo_Id = " & Processo & " AND Codigo_Id = 001" & vbCrLf

                        SqlArray.Add(Sql)
                    Next

                    If Banco.GravaBanco(SqlArray) = False Then
                        'Return HttpContext.Current.Session("ssMessage")
                    End If
                End If

                Sql = " SELECT	isnull(SUM(ValorICMS), 0) AS ValorIcms" & vbCrLf &
                      "   FROM (SELECT LEFT(OE.CodigoFiscal, 1) AS CFOP," & vbCrLf &
                      "                SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'PRODUTO' THEN nfe.Valor ELSE 0 END) AS Contabil," & vbCrLf &
                      "                SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS'    THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and OE.STICMS in(10,60))) AND YEAR(NF.Movimento) > 2011 then 0 else nfe.Base end ELSE 0 END) AS BaseICMS," & vbCrLf &
                      "                SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS'    THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and OE.STICMS in(10,60))) AND YEAR(NF.Movimento) > 2011 then 0 else nfe.Valor end ELSE 0 END) AS ValorICMS" & vbCrLf &
                      "           FROM NotasFiscais NF" & vbCrLf &
                      "		  INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                      "		     ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                      "			AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                      "			AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                      "			AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                      "			AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                      "			AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                      "			AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                      "          Inner Join OperacaoXEstado OE" & vbCrLf &
                      "		     on OE.Codigo_id = nfi.OperacaoXEstado" & vbCrLf &
                      "		  INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
                      "		     ON nfi.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
                      "			AND nfi.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
                      "			AND nfi.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
                      "			AND nfi.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
                      "			AND nfi.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
                      "			AND nfi.Serie_Id        = nfe.Serie_Id" & vbCrLf &
                      "			AND nfi.Nota_Id         = nfe.Nota_Id" & vbCrLf &
                      "			AND nfi.Produto_Id      = nfe.Produto_Id" & vbCrLf &
                      "			AND nfi.Sequencia_Id    = nfe.Sequencia_Id" & vbCrLf &
                      "		   Left Join Encargos E" & vbCrLf &
                      "		     on E.Encargo_id = nfe.Encargo_Id" & vbCrLf &
                      "         and isnull(e.EncargoAgrupador,'') <> ''" & vbCrLf &
                      "       INNER JOIN SubOperacoes so" & vbCrLf &
                      "          ON NF.Operacao    = so.Operacao_Id" & vbCrLf &
                      "         AND NF.SubOperacao = so.SubOperacoes_Id" & vbCrLf &
                      "       WHERE (" & vbCrLf &
                      "		         (" & vbCrLf &
                      "				     NF.EntradaSaida_Id = 'E'" & vbCrLf &
                      "                  And NF.Situacao   = 1" & vbCrLf &
                      "                  AND NF.TipoDeDocumento NOT IN (15)" & vbCrLf &
                      "                  AND NF.Finalidade NOT IN (30)" & vbCrLf &
                      "                  And NF.Empresa_Id ='" & Empresa(0) & "' " & vbCrLf &
                      "                  AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                      "                  AND NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                      "                  AND OE.CodigoFiscal NOT IN (1933,2933)" & vbCrLf &
                      "                  AND NOT (so.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf &
                      "			     )" & vbCrLf

                If CkConsNotasCompServicos.Checked AndAlso TemNotasCompostasDeServico Then
                    Sql &= "                 OR exists (Select  1 " & vbCrLf &
                           "			                  from #notascompostas" & vbCrLf &
                           "						     where #NotasCompostas.Empresa_Id        = nfi.Empresa_Id" & vbCrLf &
                           "                              AND #NotasCompostas.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                           "                              AND #NotasCompostas.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                           "                              AND #NotasCompostas.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                           "                              AND #NotasCompostas.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                           "                              AND #NotasCompostas.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                           "                              AND #NotasCompostas.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                           "                          )" & vbCrLf
                End If

                Sql &= "			   )" & vbCrLf &
                       "			AND isnull(e.EncargoAgrupador,nfe.Encargo_Id) IN ('IPI', 'PRODUTO','ICMS')" & vbCrLf &
                       "       GROUP BY LEFT(OE.CodigoFiscal, 1)" & vbCrLf &
                       "    ) AS Consulta" & vbCrLf

                If TemNotasCompostasDeServico Then
                    dss = Banco.ConsultaDataSet(Sql2 + Sql, "RegistrosFiscais")
                Else
                    dss = Banco.ConsultaDataSet(Sql, "RegistrosFiscais")
                End If


                If processarRegistros AndAlso dss.Tables(0).Rows.Count > 0 Then
                    'For Each Dr As DataRow In Banco.ConsultaDataSet(Sql2 + Sql, "Clientes").Tables(0).Rows
                    For Each Dr As DataRow In dss.Tables(0).Rows
                        Sql = "UPDATE ResumoRAICMS " & vbCrLf &
                              " SET Valor = " & Microsoft.VisualBasic.Replace(Dr("ValorIcms"), ",", ".") & vbCrLf &
                              " WHERE Empresa_Id = '" & Empresa(0) & "' " & vbCrLf &
                              " AND Processo_Id = " & Processo & " AND Codigo_Id = 005" & vbCrLf
                        SqlArray.Add(Sql)
                    Next

                    Banco.GravaBanco(SqlArray)
                End If

                ' Pega Saldo Credor do Periodo Anterior 
                Sql = "SELECT  COALESCE(SUM(Valor), 0) AS Valor" & vbCrLf &
                      " FROM ResumoRAICMS " & vbCrLf &
                      " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo - 1 & " And  Codigo_Id = 14 " & vbCrLf
                dss = Banco.ConsultaDataSet(Sql, "ResumoRAICMS")

                If processarRegistros AndAlso dss.Tables(0).Rows.Count > 0 Then
                    'For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                    For Each Dr As DataRow In dss.Tables(0).Rows
                        Sql = "UPDATE ResumoRAICMS " & vbCrLf &
                            " SET Valor = " & Microsoft.VisualBasic.Replace(Dr("Valor"), ",", ".") & vbCrLf &
                            " WHERE Empresa_Id = '" & Empresa(0) & "' " & vbCrLf &
                            " AND Processo_Id = " & Processo & " AND Codigo_Id = 009" & vbCrLf

                        SqlArray.Add(Sql)
                    Next

                    If Banco.GravaBanco(SqlArray) = False Then
                        'Return HttpContext.Current.Session("ssMessage")
                    End If
                End If

                If processarRegistros Then TotalizaResumo()

                Sql = "SELECT 1 AS Ordem, RR.Codigo_Id as Codigo, DR.Descricao, RR.Valor" & vbCrLf &
                      "  FROM ResumoRAICMS RR" & vbCrLf &
                      " INNER JOIN DescricaoRAICMS DR" & vbCrLf &
                      "    ON DR.Codigo_Id = RR.Codigo_Id" & vbCrLf &
                      " WHERE RR.Empresa_Id  ='" & Empresa(0) & "'" & vbCrLf &
                      "   AND RR.Processo_Id = " & Processo & vbCrLf &
                      "   AND RR.Codigo_Id   < 5" & vbCrLf &
                      " UNION" & vbCrLf &
                      "SELECT 2 AS Ordem, RR.Codigo_Id as Codigo, RIR.Descricao, RIR.Valor" & vbCrLf &
                      "  FROM ResumoItensRAICMS RIR" & vbCrLf &
                      " INNER JOIN ResumoRAICMS RR" & vbCrLf &
                      "    ON RIR.Empresa_Id  = RR.Empresa_Id" & vbCrLf &
                      "   AND RIR.Processo_Id = RR.Processo_Id" & vbCrLf &
                      "   AND RIR.Codigo_Id   = RR.Codigo_Id" & vbCrLf &
                      " WHERE RR.Empresa_Id  ='" & Empresa(0) & "'" & vbCrLf &
                      "   AND RR.Processo_Id = " & Processo & vbCrLf &
                      "   AND RR.Codigo_Id   < 5" & vbCrLf &
                      " ORDER BY 2, 1" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql, "ApuracaoIcmsDebitos"))

                Sql = "SELECT 1 AS Ordem, RR.Codigo_Id as Codigo, DR.Descricao, RR.Valor" & vbCrLf &
                      "  FROM ResumoRAICMS RR" & vbCrLf &
                      " INNER JOIN DescricaoRAICMS DR" & vbCrLf &
                      "    ON DR.Codigo_Id = RR.Codigo_Id" & vbCrLf &
                      " WHERE RR.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf &
                      "   AND RR.Processo_Id = " & Processo & vbCrLf &
                      "   AND RR.Codigo_Id between 5 and 10" & vbCrLf &
                      " UNION" & vbCrLf &
                      "SELECT 2 AS Ordem, RR.Codigo_Id as Codigo, RIR.Descricao, RIR.Valor" & vbCrLf &
                      "  FROM ResumoItensRAICMS RIR" & vbCrLf &
                      " INNER JOIN ResumoRAICMS RR" & vbCrLf &
                      "    ON RIR.Empresa_Id  = RR.Empresa_Id" & vbCrLf &
                      "   AND RIR.Processo_Id = RR.Processo_Id" & vbCrLf &
                      "   AND RIR.Codigo_Id   = RR.Codigo_Id" & vbCrLf &
                      " WHERE RR.Empresa_Id  ='" & Empresa(0) & "'" & vbCrLf &
                      "   AND RR.Processo_Id = " & Processo & vbCrLf &
                      "   AND RR.Codigo_Id between 5 and 10" & vbCrLf &
                      " ORDER BY 2, 1" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql, "ApuracaoIcmsCreditos"))

                Sql = "SELECT 1 AS Ordem, RR.Codigo_Id as Codigo, DR.Descricao, RR.Valor" & vbCrLf &
                      "  FROM ResumoRAICMS RR" & vbCrLf &
                      " INNER JOIN DescricaoRAICMS DR" & vbCrLf &
                      "    ON DR.Codigo_Id = RR.Codigo_Id" & vbCrLf &
                      " WHERE RR.Empresa_Id  ='" & Empresa(0) & "'" & vbCrLf &
                      "   AND RR.Processo_Id = " & Processo & vbCrLf &
                      "   AND RR.Codigo_Id > 10" & vbCrLf &
                      " UNION" & vbCrLf &
                      "SELECT 2 AS Ordem, RR.Codigo_Id as Codigo, RIR.Descricao, RIR.Valor" & vbCrLf &
                      "  FROM ResumoItensRAICMS RIR" & vbCrLf &
                      " INNER JOIN ResumoRAICMS RR" & vbCrLf &
                      "    ON RIR.Empresa_Id  = RR.Empresa_Id" & vbCrLf &
                      "   AND RIR.Processo_Id = RR.Processo_Id" & vbCrLf &
                      "   AND RIR.Codigo_Id   = RR.Codigo_Id" & vbCrLf &
                      " WHERE RR.Empresa_Id  ='" & Empresa(0) & "'" & vbCrLf &
                      "   AND RR.Processo_Id = " & Processo & vbCrLf &
                      "   AND RR.Codigo_Id   > 10" & vbCrLf &
                      " ORDER BY 2, 1" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql, "ApuracaoIcmsSaldos"))

                '----DAR - Diferencial de Alíquota ----------------------------------------------------------------------------

                Sql = "SELECT Dar_Id Dar, Data, Valor " & vbCrLf &
                      "  FROM DarDiferencialDeAliquota D" & vbCrLf &
                      " WHERE YEAR(D.DataReferencia)  = " & CInt(txtDataInicial.Text.Substring(6, 4)) & vbCrLf &
                      "   AND MONTH(D.DataReferencia) = " & CInt(txtDataInicial.Text.Substring(3, 2)) & vbCrLf &
                      "   AND Empresa_Id              ='" & Empresa(0) & "'" & vbCrLf &
                      "   AND EndEmpresa_Id           = " & Empresa(1) & vbCrLf &
                      " ORDER BY Data, Dar"
                ds.Merge(Banco.ConsultaDataSet(Sql, "DarDiferencialDeAliquota"))

                Dim parameters = New Dictionary(Of String, Object)()

                parameters.Add("EmpresaNome", EmpresaNome)
                parameters.Add("EmpresaCNPJ", EmpresaCNPJ)
                parameters.Add("EmpresaInscricao", EmpresaInscricao)
                parameters.Add("Periodo", Periodo)
                parameters.Add("Livro", "Livro.: " & Format(CInt(Livro), "000"))
                parameters.Add("Pagina", Folha)

                Funcoes.BindReport(Me.Page, ds, "Cr_ApuracaoIcms", IIf(CkExcell.Checked, eExportType.ExcelCrystal, eExportType.PDF), parameters, True)

            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        End If
    End Sub

#End Region

#Region "Registros de Icms Por Aliquota"

    Private Sub RegistroDeIcmsPorAliquota()
        Try
            Empresa = DdlEmpresa.SelectedValue.Split("-")
            PeriodoInicial = CDate(txtDataInicial.Text).ToString("yyyy/MM/dd")
            PeriodoFinal = CDate(txtDataFinal.Text).ToString("yyyy/MM/dd")
            Processo = DdlProcesso.SelectedValue
            Livro = txtLivro.Text
            Folha = txtFolha.Text
            Vencimento = CDate(txtVenctoDaObrigacao.Text).ToString("yyyy/MM/dd")
            CodigoDaReceita = txtCodigoDaReceita.Text

            If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0), Vencimento, CodigoDaReceita) Then
                Sql = " SELECT Cliente_Id AS Cliente, Nome, Inscricao " & vbCrLf &
                      "  FROM Clientes " & vbCrLf &
                      " WHERE Cliente_Id = '" & Empresa(0) & "'" & vbCrLf &
                      "   and Endereco_Id = " & Empresa(1) & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                For Each Dr As DataRow In ds.Tables(0).Rows
                    EmpresaNome = Dr("Nome")
                    EmpresaCNPJ = Funcoes.FormatarCpfCnpj(Dr("Cliente"))
                    EmpresaInscricao = Dr("Inscricao")
                Next

                Dim Periodo As String = "Mês ou Período/Ano : " & PeriodoInicial.ToString("dd/MM/yyyy") & " a " & PeriodoFinal.ToString("dd/MM/yyyy")

                'Temporaria Controla Notas de Serviço Conjugadas
                If CkConsNotasCompServicos.Checked Then
                    Sql2 = "  Select nfxi.Empresa_Id," & vbCrLf &
                           "         nfxi.EndEmpresa_Id," & vbCrLf &
                           "         nfxi.Cliente_Id, nfxi.EndCliente_Id," & vbCrLf &
                           "		 nfxi.EntradaSaida_Id, nfxi.Serie_Id, nfxi.Nota_Id," & vbCrLf &
                           "         sum(Case" & vbCrLf &
                           "		       when OEx.CodigoFiscal in (1933,2933)" & vbCrLf &
                           "                 then 1" & vbCrLf &
                           "                 else 0" & vbCrLf &
                           "             end ) ItensCfop," & vbCrLf &
                           "         sum(1) nrItens" & vbCrLf &
                           "  into #NotasCompostas" & vbCrLf &
                           "  From notasfiscaisxitens nfxi" & vbCrLf &
                           "  Inner join OperacaoXEstado OEx" & vbCrLf &
                           "     on OEx.Codigo_Id = nfxi.OperacaoXEstado" & vbCrLf &
                           "  Inner join (SELECT NF.Empresa_Id,      NF.EndEmpresa_Id," & vbCrLf &
                           "                     NF.Cliente_Id,      NF.EndCliente_Id," & vbCrLf &
                           "                     NF.EntradaSaida_Id, NF.Serie_Id,      NF.Nota_Id" & vbCrLf &
                           "                FROM NotasFiscais NF" & vbCrLf &
                           "               INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                           "                  ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                           "                 AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                           "                 AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                           "                 AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                           "                 AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                           "                 AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                           "                 AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                           "			   Inner Join OperacaoXEstado OE" & vbCrLf &
                           "			      on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                           "               Inner Join Suboperacoes so" & vbCrLf &
                           "                  on so.Operacao_id     = nfi.Operacao" & vbCrLf &
                           "                 and so.SubOperacoes_id = nfi.suboperacao" & vbCrLf &
                           "               WHERE NF.EntradaSaida_Id = 'E'" & vbCrLf &
                           "                 And NF.Empresa_Id      ='" & Empresa(0) & "' " & vbCrLf &
                           "                 AND NF.TipoDeDocumento NOT IN (15)" & vbCrLf &
                           "                 AND NF.Finalidade NOT IN (30)" & vbCrLf &
                           "                 AND NF.Movimento  BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                           "                 AND NF.Serie_Id  NOT IN ('D', 'F', 'REC')" & vbCrLf &
                           "                 AND(oe.CodigoFiscal  IN (1933,2933))" & vbCrLf &
                           "              )sb" & vbCrLf &
                           "   ON sb.Empresa_Id      = nfxi.Empresa_Id" & vbCrLf &
                           "  AND sb.EndEmpresa_Id   = nfxi.EndEmpresa_Id" & vbCrLf &
                           "  AND sb.Cliente_Id      = nfxi.Cliente_Id" & vbCrLf &
                           "  AND sb.EndCliente_Id   = nfxi.EndCliente_Id" & vbCrLf &
                           "  AND sb.EntradaSaida_Id = nfxi.EntradaSaida_Id" & vbCrLf &
                           "  AND sb.Serie_Id        = nfxi.Serie_Id" & vbCrLf &
                           "  AND sb.Nota_Id         = nfxi.Nota_Id" & vbCrLf &
                           "group by nfxi.Empresa_Id, nfxi.EndEmpresa_Id," & vbCrLf &
                           "         nfxi.Cliente_Id, nfxi.EndCliente_Id," & vbCrLf &
                           "         nfxi.EntradaSaida_Id, nfxi.Serie_Id, nfxi.Nota_Id" & vbCrLf

                    If PeriodoFinal > "2011-12-31" Then
                        Sql2 &= "having sum(1) <> sum(Case" & vbCrLf &
                                "                       when OEx.CodigoFiscal in (1933,2933)" & vbCrLf &
                                "                         then 1" & vbCrLf &
                                "                         else 0" & vbCrLf &
                                "                     end)" & vbCrLf
                    End If
                End If

                Sql = "SELECT Aliquota, Base, Valor" & vbCrLf &
                       "  FROM (SELECT nfe.Percentual AS Aliquota," & vbCrLf &
                       "               SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS' THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and oe.STICMS in(10,60))) AND YEAR(NF.Movimento) > 2011 then 0 else nfe.Base end ELSE 0 END)  AS Base," & vbCrLf &
                       "               SUM(CASE WHEN isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS' THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and oe.STICMS in(10,60))) AND YEAR(NF.Movimento) > 2011 then 0 else nfe.Valor end ELSE 0 END) AS Valor" & vbCrLf &
                       "          FROM NotasFiscais NF" & vbCrLf &
                       "		 INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                       "		    ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                       "		   AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                       "		   AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                       "		   AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                       "		   AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                       "		   AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                       "		   AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                       "		 Inner join OperacaoXEstado OE" & vbCrLf &
                       "		    on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                       "		 INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
                       "		    ON nfi.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
                       "		   AND nfi.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
                       "		   AND nfi.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
                       "		   AND nfi.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
                       "		   AND nfi.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
                       "		   AND nfi.Serie_Id        = nfe.Serie_Id" & vbCrLf &
                       "		   AND nfi.Nota_Id         = nfe.Nota_Id" & vbCrLf &
                       "		   AND nfi.Produto_Id      = nfe.Produto_Id" & vbCrLf &
                       "		   AND nfi.Sequencia_Id    = nfe.Sequencia_Id" & vbCrLf &
                       "          Left join Encargos E" & vbCrLf &
                       "		    on E.Encargo_id = nfe.Encargo_Id" & vbCrLf &
                       "		   and isnull(e.EncargoAgrupador,'') <> ''" & vbCrLf &
                       "         INNER JOIN SubOperacoes so" & vbCrLf &
                       "            ON NF.Operacao    = so.Operacao_Id" & vbCrLf &
                       "           AND NF.SubOperacao = so.SubOperacoes_Id" & vbCrLf &
                       "         WHERE (" & vbCrLf &
                       "                 (" & vbCrLf &
                       "                      NF.EntradaSaida_Id = 'E'" & vbCrLf &
                       "                  And NF.Situacao = 1" & vbCrLf &
                       "                  AND NF.TipoDeDocumento NOT IN (15)" & vbCrLf &
                       "                  AND NF.Finalidade NOT IN (30)" & vbCrLf &
                       "                  And NF.Empresa_Id ='" & Empresa(0) & "' " & vbCrLf &
                       "                  AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                       "                  AND NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                       "                  AND OE.CodigoFiscal NOT IN (1933,2933)" & vbCrLf &
                       "                  AND NOT (so.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf &
                       "                  )" & vbCrLf

                If CkConsNotasCompServicos.Checked Then
                    Sql &= "                  OR exists (Select  1" & vbCrLf &
                           "			                   from #notascompostas" & vbCrLf &
                           "			                  where #NotasCompostas.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                           "                                AND #NotasCompostas.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                           "                                AND #NotasCompostas.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                           "                                AND #NotasCompostas.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                           "                                AND #NotasCompostas.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                           "                                AND #NotasCompostas.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                           "                                AND #NotasCompostas.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                           "                              )" & vbCrLf
                End If

                Sql &= "                )" & vbCrLf &
                       "	       AND isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf &
                       "         GROUP BY nfe.Percentual" & vbCrLf &
                       "     ) AS Consulta" & vbCrLf &
                       " Where Base > 0" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql2 + Sql, "ApuracaoIcmsPorAliquotaEntradas"))

                Sql = "SELECT Aliquota, Base, Valor" & vbCrLf &
                       "  FROM (SELECT nfe.Percentual AS Aliquota," & vbCrLf &
                       "               SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS' AND NOT so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN nfe.Base ELSE 0 END) AS Base," & vbCrLf &
                       "               SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS' AND NOT so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN nfe.Valor ELSE 0 END) AS Valor" & vbCrLf &
                       "          FROM NotasFiscais nf" & vbCrLf &
                       "		 INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                       "		    ON nf.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                       "		   AND nf.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                       "		   AND nf.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                       "		   AND nf.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                       "		   AND nf.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                       "		   AND nf.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                       "		   AND nf.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                       "         Inner join OperacaoXEstado OE" & vbCrLf &
                       "		    on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                       "		 INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
                       "			ON nfi.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
                       "		   AND nfi.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
                       "		   AND nfi.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
                       "		   AND nfi.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
                       "		   AND nfi.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
                       "		   AND nfi.Serie_Id        = nfe.Serie_Id" & vbCrLf &
                       "		   AND nfi.Nota_Id         = nfe.Nota_Id" & vbCrLf &
                       "		   AND nfi.Produto_Id      = nfe.Produto_Id" & vbCrLf &
                       "		   AND nfi.Sequencia_Id    = nfe.Sequencia_Id" & vbCrLf &
                       "          Left Join Encargos E" & vbCrLf &
                       "		    on E.Encargo_id = nfe.Encargo_Id" & vbCrLf &
                       "           and isnull(E.EncargoAgrupador,'') <> ''" & vbCrLf &
                       "         INNER JOIN SubOperacoes so" & vbCrLf &
                       "            ON nf.Operacao    = so.Operacao_Id" & vbCrLf &
                       "           AND nf.SubOperacao = so.SubOperacoes_Id" & vbCrLf &
                       "         WHERE nf.EntradaSaida_Id = 'S'" & vbCrLf &
                       "           And nf.Situacao = 1" & vbCrLf &
                       "           And nf.Empresa_Id = '" & Empresa(0) & "' " & vbCrLf &
                       "           AND nf.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                       "           AND nf.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                       "           AND oe.CodigoFiscal NOT IN (5933,6933)" & vbCrLf &
                       "           AND NOT (SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND oe.CodigoFiscal IN (5949,6949))" & vbCrLf &
                       "		   AND isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf &
                       "         GROUP BY nfe.Percentual" & vbCrLf &
                       "	  ) AS Consulta" & vbCrLf &
                       " Where Base > 0" & vbCrLf

                ds.Merge(Banco.ConsultaDataSet(Sql, "ApuracaoIcmsPorAliquotaSaidas"))

                Dim param As New Dictionary(Of String, Object)
                param.Add("EmpresaNome", EmpresaNome)
                param.Add("EmpresaCNPJ", EmpresaCNPJ)
                param.Add("EmpresaInscricao", EmpresaInscricao)
                param.Add("Periodo", Periodo)
                param.Add("Livro", "Livro.: " & Format(CInt(Livro), "000"))
                param.Add("Pagina", Folha)

                Funcoes.BindReport(Me.Page, ds, "Cr_ApuracaoIcmsPorAliquota", IIf(CkExcell.Checked, eExportType.ExcelCrystal, eExportType.PDF), param, True)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Resumo de Entradas Por UF"

    Private Sub ResumoDeEntradasPorUF()

        Empresa = DdlEmpresa.SelectedValue.Split("-")
        PeriodoInicial = CDate(txtDataInicial.Text).ToString("yyyy/MM/dd")
        PeriodoFinal = CDate(txtDataFinal.Text).ToString("yyyy/MM/dd")
        Processo = DdlProcesso.SelectedValue
        Livro = txtLivro.Text
        Folha = txtFolha.Text
        Vencimento = CDate(txtVenctoDaObrigacao.Text).ToString("yyyy/MM/dd")
        CodigoDaReceita = txtCodigoDaReceita.Text

        If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0), Vencimento, CodigoDaReceita) Then
            Try
                Sql = " SELECT Cliente_Id AS Cliente, Nome, Inscricao " & vbCrLf &
                      "  FROM Clientes " & vbCrLf &
                      " WHERE Cliente_Id = '" & Empresa(0) & "'" & vbCrLf &
                      "   and Endereco_Id = " & Empresa(1) & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                For Each Dr As DataRow In ds.Tables(0).Rows
                    EmpresaNome = Dr("Nome")
                    EmpresaCNPJ = Funcoes.FormatarCpfCnpj(Dr("Cliente"))
                    EmpresaInscricao = Dr("Inscricao")
                Next

                Dim Periodo As String = "Mês ou Período/Ano : " & PeriodoInicial.ToString("dd/MM/yyyy") & " a " & PeriodoFinal.ToString("dd/MM/yyyy")

                'Temporaria Controla Notas de Serviço Conjugadas
                If CkConsNotasCompServicos.Checked Then
                    Sql2 = "Select nfxi.Empresa_Id," & vbCrLf &
                           "       nfxi.EndEmpresa_Id," & vbCrLf &
                           "       nfxi.Cliente_Id, nfxi.EndCliente_Id," & vbCrLf &
                           "       nfxi.EntradaSaida_Id, nfxi.Serie_Id, nfxi.Nota_Id," & vbCrLf &
                           "       sum(Case" & vbCrLf &
                           "		     when OEx.CodigoFiscal in (1933,2933)" & vbCrLf &
                           "               then 1" & vbCrLf &
                           "               else 0" & vbCrLf &
                           "           end) ItensCfop," & vbCrLf &
                           "       sum(1) nrItens" & vbCrLf &
                           "  into #NotasCompostas" & vbCrLf &
                           "  From notasfiscaisxitens nfxi" & vbCrLf &
                           " Inner join OperacaoXEstado OEx" & vbCrLf &
                           "    on OEx.Codigo_Id = nfxi.OperacaoXEstado" & vbCrLf &
                           " Inner join (SELECT NF.Empresa_Id,      NF.EndEmpresa_Id," & vbCrLf &
                           "                    NF.Cliente_Id,      NF.EndCliente_Id," & vbCrLf &
                           "                    NF.EntradaSaida_Id, NF.Serie_Id,      NF.Nota_Id" & vbCrLf &
                           "               FROM NotasFiscais NF" & vbCrLf &
                           "              INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                           "                 ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                           "                AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                           "                AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                           "                AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                           "                AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                           "                AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                           "                AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                           "			  Inner join OperacaoXEstado OE" & vbCrLf &
                           "			     on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                           "              Inner Join Suboperacoes so" & vbCrLf &
                           "                 on so.Operacao_id     = nfi.Operacao" & vbCrLf &
                           "                and so.SubOperacoes_id = nfi.suboperacao" & vbCrLf &
                           "              WHERE NF.EntradaSaida_Id = 'E'" & vbCrLf &
                           "                And NF.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf &
                           "                AND NF.TipoDeDocumento NOT IN (15)" & vbCrLf &
                           "                AND NF.Finalidade NOT IN (30)" & vbCrLf &
                           "                AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf &
                           "                AND NF.Serie_Id  NOT IN ('D', 'F', 'REC')" & vbCrLf &
                           "                AND OE.CodigoFiscal  IN (1933,2933)" & vbCrLf &
                           "             )sb" & vbCrLf &
                           "     ON sb.Empresa_Id      = nfxi.Empresa_Id" & vbCrLf &
                           "    AND sb.EndEmpresa_Id   = nfxi.EndEmpresa_Id" & vbCrLf &
                           "    AND sb.Cliente_Id      = nfxi.Cliente_Id" & vbCrLf &
                           "    AND sb.EndCliente_Id   = nfxi.EndCliente_Id" & vbCrLf &
                           "    AND sb.EntradaSaida_Id = nfxi.EntradaSaida_Id" & vbCrLf &
                           "    AND sb.Serie_Id        = nfxi.Serie_Id" & vbCrLf &
                           "    AND sb.Nota_Id         = nfxi.Nota_Id" & vbCrLf &
                           "  group by nfxi.Empresa_Id, nfxi.EndEmpresa_Id," & vbCrLf &
                           "           nfxi.Cliente_Id, nfxi.EndCliente_Id," & vbCrLf &
                           "           nfxi.EntradaSaida_Id, nfxi.Serie_Id, nfxi.Nota_Id" & vbCrLf

                    If PeriodoFinal > "2011-12-31" Then
                        Sql2 &= " having sum(1) <> sum(Case" & vbCrLf &
                                "                        when OEx.CodigoFiscal in (1933,2933)" & vbCrLf &
                                "                          then 1" & vbCrLf &
                                "                          else 0" & vbCrLf &
                                "                      end)" & vbCrLf
                    End If
                End If

                Sql = " SELECT Estado," & vbCrLf &
                      "        SUM(((ValorContabil+ValorIPI+ValorContabilFrete+OutrasDespesas) - ValorDesconto)) AS Contabil," & vbCrLf &
                      "        SUM(BaseICMS) AS Base," & vbCrLf &
                      "        SUM(ValorICMS) AS Icms," & vbCrLf &
                      "        CodigoFiscal AS CFOP" & vbCrLf &
                      "   FROM (SELECT NF.EstadoDoCliente AS Estado," & vbCrLf &
                      "                SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS'      THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and OE.STICMS in(10,60))) AND YEAR(NF.Movimento) > 2011 then 0 else nfe.Base end ELSE 0 END) AS BaseICMS," & vbCrLf &
                      "                SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS'      THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and OE.STICMS in(10,60))) AND YEAR(NF.Movimento) > 2011 then 0 else nfe.Valor end ELSE 0 END) AS ValorICMS," & vbCrLf &
                      "                SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'PRODUTO'   THEN nfe.Valor ELSE 0 END) AS ValorContabil," & vbCrLf &
                      "                SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'DESCONTOS' THEN nfe.Valor ELSE 0 END) AS ValorDesconto," & vbCrLf &
                      "                SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'IPI'       THEN nfe.Valor ELSE 0 END) AS ValorIPI," & vbCrLf &
                      "                SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) IN ('FRETES','SEGURO') THEN nfe.Valor ELSE 0 END) AS ValorContabilFrete," & vbCrLf &
                      "                SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) IN ('DESP.ADUANEIRAS','ICMS SUBSTITUIC') THEN nfe.Valor ELSE 0 END) AS OutrasDespesas," & vbCrLf &
                      "                OE.CodigoFiscal" & vbCrLf &
                      "           FROM NotasFiscais NF" & vbCrLf &
                      "		  INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                      "		     ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                      "			AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                      "			AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                      "			AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                      "		    AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                      "			AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                      "			AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                      "          INNER JOIN OperacaoXEstado OE" & vbCrLf &
                      "		     on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                      "		  INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
                      "			 ON nfi.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
                      "			AND nfi.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
                      "			AND nfi.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
                      "			AND nfi.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
                      "			AND nfi.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
                      "			AND nfi.Serie_Id        = nfe.Serie_Id" & vbCrLf &
                      "			AND nfi.Nota_Id         = nfe.Nota_Id" & vbCrLf &
                      "			AND nfi.Produto_Id      = nfe.Produto_Id" & vbCrLf &
                      "			AND nfi.Sequencia_Id    = nfe.Sequencia_Id" & vbCrLf &
                      "           Left Join Encargos E" & vbCrLf &
                      "		     on E.Encargo_id = nfe.Encargo_Id" & vbCrLf &
                      "            and isnull(E.EncargoAgrupador,'') <> ''" & vbCrLf &
                      "          INNER JOIN SubOperacoes so" & vbCrLf &
                      "             ON NF.Operacao    = so.Operacao_Id" & vbCrLf &
                      "            AND NF.SubOperacao = so.SubOperacoes_Id" & vbCrLf &
                      "          WHERE (" & vbCrLf &
                      "		          (" & vbCrLf &
                      "				          NF.EntradaSaida_Id = 'E'" & vbCrLf &
                      "                   AND NF.Situacao = 1" & vbCrLf &
                      "                   AND NF.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf &
                      "                   AND NF.TipoDeDocumento NOT IN (15)" & vbCrLf &
                      "                   AND NF.Finalidade NOT IN (30)" & vbCrLf &
                      "                   AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf &
                      "                   AND NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                      "                   AND OE.CodigoFiscal NOT IN (1933,2933)" & vbCrLf &
                      "                   AND NOT (so.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf &
                      "                  )" & vbCrLf

                If CkConsNotasCompServicos.Checked Then
                    Sql &= "                 OR exists (Select  1 " & vbCrLf &
                           "			                 from #notascompostas" & vbCrLf &
                           "						    where #NotasCompostas.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                           "                              AND #NotasCompostas.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                           "                              AND #NotasCompostas.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                           "                              AND #NotasCompostas.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                           "                              AND #NotasCompostas.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                           "                              AND #NotasCompostas.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                           "                              AND #NotasCompostas.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                           "                           )" & vbCrLf
                End If

                Sql &= "				)" & vbCrLf &
                       "            AND isnull(E.EncargoAgrupador,nfe.Encargo_Id) IN ('IPI', 'PRODUTO','FRETES','SEGURO','DESP.ADUANEIRAS','DESCONTOS','ICMS','ICMS SUBSTITUIC')" & vbCrLf &
                       "          GROUP BY NF.EstadoDoCliente, OE.CodigoFiscal" & vbCrLf &
                       "	  ) AS Consulta" & vbCrLf &
                       "   GROUP BY Estado, CodigoFiscal" & vbCrLf


                ds.Merge(Banco.ConsultaDataSet(Sql2 + Sql, "EntradasPorUF"))

                Dim parameters = New Dictionary(Of String, Object)()

                parameters.Add("EmpresaNome", EmpresaNome)
                parameters.Add("EmpresaCNPJ", EmpresaCNPJ)
                parameters.Add("EmpresaInscricao", EmpresaInscricao)
                parameters.Add("Periodo", Periodo)
                parameters.Add("Livro", "Livro.: " & Format(CInt(Livro), "000"))
                parameters.Add("Pagina", Folha)

                Funcoes.BindReport(Me.Page, ds, IIf(ckConferenciaCfopIcms.Checked, "Cr_ResumoDeEntradasPorUFCFOP", "Cr_ResumoDeEntradasPorUF"), IIf(CkExcell.Checked, eExportType.ExcelCrystal, eExportType.PDF), parameters)
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        End If
    End Sub

#End Region

#Region "Resumo de Saidas Por UF"

    Private Sub ResumoDeSaidasPorUF()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        PeriodoInicial = CDate(txtDataInicial.Text).ToString("yyyy/MM/dd")
        PeriodoFinal = CDate(txtDataFinal.Text).ToString("yyyy/MM/dd")
        Processo = DdlProcesso.SelectedValue
        Livro = txtLivro.Text
        Folha = txtFolha.Text
        Vencimento = CDate(txtVenctoDaObrigacao.Text).ToString("yyyy/MM/dd")
        CodigoDaReceita = txtCodigoDaReceita.Text

        If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0), Vencimento, CodigoDaReceita) Then
            Try
                Sql = "SELECT Cliente_Id AS Cliente, Nome, Inscricao " & vbCrLf &
                      "  FROM Clientes " & vbCrLf &
                      " WHERE Cliente_Id  ='" & Empresa(0) & "'" & vbCrLf &
                      "   AND Endereco_Id = " & Empresa(1) & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                For Each Dr As DataRow In ds.Tables(0).Rows
                    EmpresaNome = Dr("Nome")
                    EmpresaCNPJ = Funcoes.FormatarCpfCnpj(Dr("Cliente"))
                    EmpresaInscricao = Dr("Inscricao")
                Next

                Dim Periodo As String = "Mês ou Período/Ano : " & PeriodoInicial.ToString("dd/MM/yyyy") & " a " & PeriodoFinal.ToString("dd/MM/yyyy")

                Sql = "SELECT Estado," & vbCrLf &
                      "       CASE" & vbCrLf &
                      "         WHEN Consulta.Situacao <> 1" & vbCrLf &
                      "          THEN 0" & vbCrLf &
                      "          ELSE ((ValorContabil+ValorIPI+ValorContabilFrete+OutrasDespesas) - ValorDesconto)" & vbCrLf &
                      "       END AS Contabil," & vbCrLf &
                      "       CASE" & vbCrLf &
                      "         WHEN Consulta.Situacao <> 1" & vbCrLf &
                      "          THEN 0" & vbCrLf &
                      "          ELSE BaseICMS" & vbCrLf &
                      "       END AS Base," & vbCrLf &
                      "       CASE" & vbCrLf &
                      "         WHEN Consulta.Situacao <> 1" & vbCrLf &
                      "           THEN 0" & vbCrLf &
                      "           ELSE ValorIcms" & vbCrLf &
                      "       END AS Icms," & vbCrLf &
                      "       Cfop" & vbCrLf &
                      "  FROM (SELECT NF.EstadoDoCliente AS Estado," & vbCrLf &
                      "               OE.CodigoFiscal AS CFOP," & vbCrLf &
                      "				  S.Situacao_id as Situacao," & vbCrLf &
                      "               SUM(CASE WHEN isnull(E.EncargoAgrupador,NFE.Encargo_Id) = 'PRODUTO'         THEN NFE.Valor ELSE 0 END) AS ValorContabil," & vbCrLf &
                      "               SUM(CASE WHEN isnull(E.EncargoAgrupador,NFE.Encargo_Id) IN ('FRETES','SEGURO') THEN NFE.Valor ELSE 0 END) AS ValorContabilFrete," & vbCrLf &
                      "               SUM(CASE WHEN isnull(E.EncargoAgrupador,NFE.Encargo_Id) = 'DESCONTOS'       THEN NFE.Valor ELSE 0 END) AS ValorDesconto," & vbCrLf &
                      "               SUM(CASE WHEN isnull(E.EncargoAgrupador,NFE.Encargo_Id) = 'ICMS' AND NOT so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN NFE.Base  ELSE 0 END) AS BaseICMS," & vbCrLf &
                      "               SUM(CASE WHEN isnull(E.EncargoAgrupador,NFE.Encargo_Id) = 'ICMS' AND NOT so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') THEN NFE.Valor ELSE 0 END) AS ValorICMS," & vbCrLf &
                      "               SUM(CASE WHEN isnull(E.EncargoAgrupador,NFE.Encargo_Id) = 'IPI'             THEN NFE.Base  ELSE 0 END) AS BaseIPI," & vbCrLf &
                      "               SUM(CASE WHEN isnull(E.EncargoAgrupador,NFE.Encargo_Id) = 'IPI'             THEN NFE.Valor ELSE 0 END) AS ValorIPI," & vbCrLf &
                      "               SUM(CASE WHEN isnull(E.EncargoAgrupador,NFE.Encargo_Id) IN ('DESP.ADUANEIRAS','ICMS SUBSTITUIC') THEN NFE.Valor ELSE 0 END) AS OutrasDespesas" & vbCrLf &
                      "          FROM NotasFiscais NF" & vbCrLf &
                      "         INNER JOIN NotasFiscaisXItens NFI" & vbCrLf &
                      "            ON NF.Empresa_Id      = NFI.Empresa_Id" & vbCrLf &
                      "           AND NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf &
                      "           AND NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf &
                      "           AND NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf &
                      "           AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf &
                      "           AND NF.Serie_Id        = NFI.Serie_Id" & vbCrLf &
                      "           AND NF.Nota_Id         = NFI.Nota_Id" & vbCrLf &
                      "		    Inner Join OperacaoXEstado OE" & vbCrLf &
                      "		       on OE.Codigo_Id = NFI.OperacaoXEstado" & vbCrLf &
                      "         INNER JOIN NotasFiscaisXEncargos NFE" & vbCrLf &
                      "            ON NFI.Empresa_Id      = NFE.Empresa_Id" & vbCrLf &
                      "           AND NFI.EndEmpresa_Id   = NFE.EndEmpresa_Id" & vbCrLf &
                      "           AND NFI.Cliente_Id      = NFE.Cliente_Id" & vbCrLf &
                      "           AND NFI.EndCliente_Id   = NFE.EndCliente_Id" & vbCrLf &
                      "           AND NFI.EntradaSaida_Id = NFE.EntradaSaida_Id" & vbCrLf &
                      "           AND NFI.Serie_Id        = NFE.Serie_Id" & vbCrLf &
                      "           AND NFI.Nota_Id         = NFE.Nota_Id" & vbCrLf &
                      "           AND NFI.Produto_Id      = NFE.Produto_Id" & vbCrLf &
                      "           AND NFI.Sequencia_Id    = NFE.Sequencia_Id" & vbCrLf &
                      "          Left join Encargos E" & vbCrLf &
                      "		       on E.Encargo_id = nfe.Encargo_Id" & vbCrLf &
                      "           and isnull(E.EncargoAgrupador,'') <> ''" & vbCrLf &
                      "         INNER JOIN SubOperacoes SO" & vbCrLf &
                      "            ON NF.Operacao    = SO.Operacao_Id" & vbCrLf &
                      "           AND NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
                      "         Inner join Situacoes S" & vbCrLf &
                      "            ON S.Situacao_id = NF.Situacao" & vbCrLf &
                      "         INNER JOIN Clientes C" & vbCrLf &
                      "            on C.Cliente_id  = NF.Cliente_Id" & vbCrLf &
                      "           and C.Endereco_id = NF.EndCliente_Id" & vbCrLf &
                      "         WHERE  NF.EntradaSaida_Id = 'S'" & vbCrLf &
                      "           And  NF.Empresa_Id ='" & Empresa(0) & "' " & vbCrLf &
                      "           AND  NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                      "           AND  NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                      "           AND (NFI.CFOP_Id NOT IN (5933,6933))" & vbCrLf &
                      "           AND NOT (SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND NFI.CFOP_Id IN (5949,6949))" & vbCrLf &
                      "		      AND isnull(E.EncargoAgrupador,NFE.Encargo_Id) IN ('IPI','DESP.ADUANEIRAS', 'PRODUTO','FRETES','SEGURO','DESCONTOS','ICMS','ICMS SUBSTITUIC')" & vbCrLf &
                      "         GROUP BY NF.EstadoDoCliente, OE.CodigoFiscal, S.Situacao_id" & vbCrLf &
                      "    ) AS Consulta" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql, "SaidasPorUF"))

                Dim parameters = New Dictionary(Of String, Object)()

                parameters.Add("EmpresaNome", EmpresaNome)
                parameters.Add("EmpresaCNPJ", EmpresaCNPJ)
                parameters.Add("EmpresaInscricao", EmpresaInscricao)
                parameters.Add("Periodo", Periodo)
                parameters.Add("Livro", "Livro.: " & Format(CInt(Livro), "000"))
                parameters.Add("Pagina", Folha)

                Funcoes.BindReport(Me.Page, ds, IIf(ckConferenciaCfopIcms.Checked, "Cr_ResumoDeSaidasPorUFCFOP", "Cr_ResumoDeSaidasPorUF"), IIf(CkExcell.Checked, eExportType.ExcelCrystal, eExportType.PDF), parameters)
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        End If
    End Sub

#End Region

#Region "Codigos De Emitentes"

    Private Sub CodigosDeEmitentes()
        Try
            Empresa = DdlEmpresa.SelectedValue.Split("-")
            PeriodoInicial = CDate(txtDataInicial.Text).ToString("yyyy/MM/dd")
            PeriodoFinal = CDate(txtDataFinal.Text).ToString("yyyy/MM/dd")
            Processo = DdlProcesso.SelectedValue
            Livro = txtLivro.Text
            Folha = txtFolha.Text
            Vencimento = CDate(txtVenctoDaObrigacao.Text).ToString("yyyy/MM/dd")
            CodigoDaReceita = txtCodigoDaReceita.Text

            If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0), Vencimento, CodigoDaReceita) Then
                Sql = "SELECT Cliente_Id AS Cliente, Nome, Inscricao " & vbCrLf &
                      "FROM Clientes " & vbCrLf &
                      "WHERE Cliente_Id = '" & Empresa(0) & "'" & vbCrLf &
                      "  And Endereco_id = " & Empresa(1) & vbCrLf

                ds2 = Banco.ConsultaDataSet(Sql, "Clientes")

                For Each Dr As DataRow In ds2.Tables(0).Rows
                    EmpresaNome = Dr("Nome")
                    EmpresaCNPJ = Funcoes.FormatarCpfCnpj(Dr("Cliente"))
                    EmpresaInscricao = Dr("Inscricao")
                Next

                Dim Periodo As String = "Mês ou Período/Ano : " & PeriodoInicial.ToString("dd/MM/yyyy") & " a " & PeriodoFinal.ToString("dd/MM/yyyy")

                Sql = " SELECT  Distinct   Clientes.Cliente_Id as Cliente, Clientes.Inscricao, Clientes.Nome, Clientes.Endereco, Clientes.Cidade, Clientes.Estado" & vbCrLf &
                      " FROM  Clientes INNER JOIN" & vbCrLf &
                      " NotasFiscais ON Clientes.Cliente_Id = NotasFiscais.Cliente_Id" & vbCrLf &
                      " WHERE  NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND " & vbCrLf &
                      " NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                      " order by Clientes.Nome, Clientes.Estado, Clientes.Cidade" & vbCrLf

                ds.Merge(Banco.ConsultaDataSet(Sql, "Clientes"))

                Dim parameters = New Dictionary(Of String, Object)()

                parameters.Add("EmpresaNome", EmpresaNome)
                parameters.Add("EmpresaCNPJ", EmpresaCNPJ)
                parameters.Add("EmpresaInscricao", EmpresaInscricao)
                parameters.Add("Periodo", Periodo)
                parameters.Add("Livro", "Livro.: " & Format(CInt(Livro), "000"))
                parameters.Add("Pagina", Folha)

                Funcoes.BindReport(Me.Page, ds, "Cr_CodigosDeEmitentes", IIf(CkExcell.Checked, eExportType.ExcelCrystal, eExportType.PDF), parameters)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


#End Region

#Region "Observaçoes Fiscais"

    Private Sub ObservacoesFiscais()
        Try
            Empresa = DdlEmpresa.SelectedValue.Split("-")
            PeriodoInicial = CDate(txtDataInicial.Text).ToString("yyyy/MM/dd")
            PeriodoFinal = CDate(txtDataFinal.Text).ToString("yyyy/MM/dd")
            Processo = DdlProcesso.SelectedValue
            Livro = txtLivro.Text
            Folha = txtFolha.Text
            Vencimento = CDate(txtVenctoDaObrigacao.Text).ToString("yyyy/MM/dd")
            CodigoDaReceita = txtCodigoDaReceita.Text

            If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0), Vencimento, CodigoDaReceita) Then

                Sql = "SELECT Cliente_Id AS Cliente, Nome, Inscricao " & vbCrLf &
                      "FROM Clientes " & vbCrLf &
                      "WHERE Cliente_Id = '" & Empresa(0) & "'" & vbCrLf &
                      "   and Endereco_Id = " & Empresa(1) & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                For Each Dr As DataRow In ds.Tables(0).Rows
                    EmpresaNome = Dr("Nome")
                    EmpresaCNPJ = Funcoes.FormatarCpfCnpj(Dr("Cliente"))
                    EmpresaInscricao = Dr("Inscricao")
                Next

                Dim Periodo As String = "Mês ou Período/Ano : " & PeriodoInicial.ToString("dd/MM/yyyy") & " a " & PeriodoFinal.ToString("dd/MM/yyyy")

                Sql = "SELECT  Codigo_Id as Codigo, Descricao FROM ObservacoesFiscais" & vbCrLf &
                      " Order by Codigo_Id" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql, "ObservacoesFiscais"))

                Dim parameters = New Dictionary(Of String, Object)()

                parameters.Add("EmpresaNome", EmpresaNome)
                parameters.Add("EmpresaCNPJ", EmpresaCNPJ)
                parameters.Add("EmpresaInscricao", EmpresaInscricao)
                parameters.Add("Periodo", Periodo)
                parameters.Add("Livro", "Livro.: " & Format(CInt(Livro), "000"))
                parameters.Add("Pagina", Folha)

                Funcoes.BindReport(Me.Page, ds, "Cr_ObservacoesFiscais", IIf(CkExcell.Checked, eExportType.ExcelCrystal, eExportType.PDF), parameters)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Termos"

    Private Sub Termos()
        Try
            Empresa = DdlEmpresa.SelectedValue.Split("-")
            PeriodoInicial = CDate(txtDataInicial.Text).ToString("yyyy/MM/dd")
            PeriodoFinal = CDate(txtDataFinal.Text).ToString("yyyy/MM/dd")
            Processo = DdlProcesso.SelectedValue
            Livro = txtLivro.Text
            Folha = txtFolha.Text
            Vencimento = CDate(txtVenctoDaObrigacao.Text).ToString("yyyy/MM/dd")
            CodigoDaReceita = txtCodigoDaReceita.Text

            If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0), Vencimento, CodigoDaReceita) Then
                Dim Ds As New DataSet
                Dim LIVRO01 As String = Format(Val(Livro), "000")
                Dim mystr As String
                mystr = Format(Val(Folha), "000")
                Folha = mystr

                Sql = "  SELECT Clientes.Cliente_Id as CNPJ, Clientes.Nome as RazaoSocial, Clientes.Endereco, Clientes.Cidade, Clientes.Estado, Clientes.Inscricao, ClientesXEmpresas.RegistroNaJunta, " & vbCrLf &
                      " ClientesXEmpresas.EstadodaJunta, ClientesXEmpresas.DatadeFundacao, ClientesXEmpresas.NomeDoContador, ClientesXEmpresas.CPFContador, " & vbCrLf &
                      " ClientesXEmpresas.CRCContador, ClientesXEmpresas.NomeDoTitular, ClientesXEmpresas.CPFTitular, ClientesXEmpresas.CondicaoTitular, " & vbCrLf &
                      " Estados.Descricao as NomeDoEstado, 1 as Ordem" & vbCrLf &
                      " FROM Clientes INNER JOIN" & vbCrLf &
                      " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf &
                      " Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN" & vbCrLf &
                      " Estados ON ClientesXEmpresas.EstadodaJunta = Estados.Estado_Id" & vbCrLf &
                      " WHERE Clientes.Cliente_Id Like '" & Empresa(0) & "'" & vbCrLf &
                      " Union" & vbCrLf &
                      " SELECT Clientes.Cliente_Id as CNPJ, Clientes.Nome as RazaoSocial, Clientes.Endereco, Clientes.Cidade, Clientes.Estado, Clientes.Inscricao, ClientesXEmpresas.RegistroNaJunta, " & vbCrLf &
                      " ClientesXEmpresas.EstadodaJunta, ClientesXEmpresas.DatadeFundacao, ClientesXEmpresas.NomeDoContador, ClientesXEmpresas.CPFContador, " & vbCrLf &
                      " ClientesXEmpresas.CRCContador, ClientesXEmpresas.NomeDoTitular, ClientesXEmpresas.CPFTitular, ClientesXEmpresas.CondicaoTitular, " & vbCrLf &
                      " Estados.Descricao as NomeDoEstado, 2 as Ordem" & vbCrLf &
                      " FROM Clientes INNER JOIN" & vbCrLf &
                      " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf &
                      " Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN" & vbCrLf &
                      " Estados ON ClientesXEmpresas.EstadodaJunta = Estados.Estado_Id" & vbCrLf &
                      " WHERE Clientes.Cliente_Id Like '" & Empresa(0) & "' and Clientes.Endereco_id = " & Empresa(1) & vbCrLf
                Ds.Merge(Banco.ConsultaDataSet(Sql, "Termos"))

                Dim param As New Dictionary(Of String, Object)
                param.Add("Folha", Folha)
                param.Add("DIAINI", PeriodoInicial.Day.ToString().PadLeft(2, "0"))
                param.Add("MESINI", MonthName(Month(PeriodoInicial), False))
                param.Add("ANOINI", Str(Year(PeriodoInicial)))
                param.Add("DIAFIM", PeriodoFinal.Day.ToString().PadLeft(2, "0"))
                param.Add("MESFIM", MonthName(Month(PeriodoFinal), False))
                param.Add("ANOFIM", Str(Year(PeriodoFinal)))
                param.Add("LIVRO", LIVRO01)
                param.Add("Opcao", Opcao)

                Funcoes.BindReport(Me.Page, Ds, "Cr_TermosRegistrosFiscaisICMS", IIf(CkExcell.Checked, eExportType.ExcelCrystal, eExportType.PDF), param)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Funções"

    Private Function GravarProcesso(ByVal Processo As String, ByVal Livro As String, ByVal Folha As String, ByVal PeriodoIni As String, ByVal PeriodoFim As String, ByVal Cliente As String, ByVal Vencimento As String, ByVal CodigoDaReceita As String) As Boolean
        Try
            If Funcoes.VerificaPermissao("RegistrosDeICMS", "GRAVAR") Then
                Dim strSQL As String

                strSQL = "SELECT * FROM ProcessoRAICMS WHERE Empresa_Id = '" & Cliente & "' And Processo_Id = " & Processo

                Dim dsProcesso As DataSet = Banco.ConsultaDataSet(strSQL, "ProcessoRAICMS")

                If dsProcesso.Tables(0).Rows.Count = 0 Then
                    strSQL = "INSERT INTO ProcessoRAICMS " & vbCrLf &
                             "(Empresa_Id, Processo_Id, PeriodoInicial, PeriodoFinal, Livro, PaginaInicial, Fechado, CodigoDaReceita, VencimentoDaObrigacao) " & vbCrLf &
                             "VALUES ('" & Cliente & "', " & Processo & ", '" & PeriodoIni.ToSqlDate() & "', " & vbCrLf &
                             "'" & PeriodoFim.ToSqlDate() & "', " & Livro & ", " & Folha & ", 'N','" & CodigoDaReceita & "', '" & Vencimento.ToSqlDate() & "')" & vbCrLf
                Else
                    strSQL = "UPDATE ProcessoRAICMS " & vbCrLf &
                             "SET PeriodoInicial = '" & PeriodoIni.ToSqlDate() & "'" & vbCrLf &
                             ", PeriodoFinal = '" & PeriodoFim.ToSqlDate() & "'" & vbCrLf &
                             ", Livro = " & Livro & vbCrLf &
                             ", PaginaInicial = " & Folha & vbCrLf &
                             ", CodigoDaReceita = '" & CodigoDaReceita & "'" & vbCrLf &
                             ", VencimentoDaObrigacao = '" & PeriodoFim.ToSqlDate() & "'" & vbCrLf &
                             " WHERE Empresa_Id = '" & Cliente & "' " & vbCrLf &
                             " AND Processo_Id = " & Processo & vbCrLf
                End If

                Dim alSQL As New ArrayList
                alSQL.Add(strSQL)

                Return Banco.GravaBanco(alSQL)
            Else
                Return True
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function DeletarProcesso(ByVal Processo As String, ByVal Livro As String, ByVal Folha As String, ByVal PeriodoIni As String, ByVal PeriodoFim As String, ByVal Cliente As String, ByVal Vencimento As String, ByVal CodigoDaReceita As String) As Boolean
        Try
            Dim strSQL As String

            strSQL = "SELECT * FROM ProcessoRAICMS WHERE Empresa_Id = '" & Cliente & "' And Processo_Id = " & Processo

            Dim dsProcesso As DataSet = Banco.ConsultaDataSet(strSQL, "ProcessoRAICMS")

            If dsProcesso.Tables(0).Rows.Count > 0 Then
                strSQL = "Delete ProcessoRAICMS WHERE " & vbCrLf &
                         "Empresa_Id= '" & Cliente & "' AND Processo_Id= " & Processo & " " & vbCrLf
            Else
                Mensagem = "Erro, não foram encontrados dados com estes parametros para excluir."
            End If

            Dim alSQL As New ArrayList
            alSQL.Add(strSQL)

            If Banco.GravaBanco(alSQL) = False Then
                Mensagem = "Erro de exclusão na base de dados."
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Else
                Limpar()
            End If
            Return True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Function

    Private Sub GravaResumoIcms()
        Empresa = DdlEmpresa.SelectedValue.Split("-")

        Sql = "INSERT INTO ResumoRAICMS " & vbCrLf &
              "(Empresa_Id, Processo_Id, Codigo_Id, Valor) " & vbCrLf &
              "SELECT '" & Empresa(0) & "', " & Processo & ", " & vbCrLf &
              "Codigo_Id, 0 " & vbCrLf &
              "FROM DescricaoRAICMS" & vbCrLf
        SqlArray.Add(Sql)

        If Banco.GravaBanco(SqlArray) = False Then
            Mensagem = "Erro de gravação na base de dados."
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        End If
    End Sub

    Function Validar() As Boolean
        If String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            MsgBox(Me.Page, "Informe a unidade de negócio.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            MsgBox(Me.Page, "Informe o período inicial.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe o período final.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtLivro.Text) Then
            MsgBox(Me.Page, "Informe o número do livro.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtFolha.Text) Then
            MsgBox(Me.Page, "Informe o número da folha.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtCodigoDaReceita.Text) Then
            MsgBox(Me.Page, "É necessário informar o código da receita.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlProcesso.SelectedValue) AndAlso (chkImpressao.Checked Or chkZIP.Checked Or chkAbrir.Checked) Then
            MsgBox(Me.Page, "É necessário informar o processo.")
            Return False
        End If

        Return True
    End Function

    Function ValidarImpressao() As Boolean
        If txtFolhaInicial.Text.Length > 0 And txtFolhaFinal.Text.Length = 0 Then
            MsgBox(Me.Page, "Informe o número da folha final.")
            Return False
        ElseIf txtFolhaInicial.Text.Length = 0 And txtFolhaFinal.Text.Length > 0 Then
            MsgBox(Me.Page, "Informe o número da folha inicial.")
            Return False
        ElseIf txtFolhaInicial.Text.Length > 0 And txtFolha.Text.Length > 0 AndAlso CType(txtFolhaInicial.Text, Integer) < CType(txtFolha.Text, Integer) Then
            MsgBox(Me.Page, "A folha inicial não pode ser menor que a folha informada.")
            Return False
        ElseIf chkImpressao.Checked = False And chkAbrir.Checked = False And chkZIP.Checked = False Then
            Return False
        End If

        Return True
    End Function

    Private Sub PegarNumeroDoProcesso()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        Processo = 1

        Sql = "SELECT COALESCE(max(Processo_Id), 0) AS Processo  FROM ProcessoRAICMS WHERE Empresa_Id = '" & Empresa(0) & "'"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Processo = Dr("Processo") + 1
        Next
    End Sub

    Private Sub Limpar()
        txtDataInicial.Text = ""
        txtDataFinal.Text = ""
        txtLivro.Text = ""
        txtFolha.Text = ""
        txtVenctoDaObrigacao.Text = ""
        txtCodigoDaReceita.Text = ""
        GridTermos.Parent.Visible = False

        txtNomeCliente.Text = ""
        chkImpressao.Checked = False
        chkZIP.Checked = False
        chkAbrir.Checked = False
        txtFolhaInicial.Text = ""
        txtFolhaFinal.Text = ""
        txtNumeroNota.Text = ""
        txtSerie.Text = ""


        If Funcoes.VerificaPermissao("RegistrosDeICMS", "GRAVAR") Then
            lnkNovo.Parent.Visible = True
        Else
            lnkNovo.Parent.Visible = False
        End If

        If Funcoes.VerificaPermissao("RegistrosDeICMS", "EXCLUIR") Then
            lnkExcluir.Parent.Visible = True
        Else
            lnkExcluir.Parent.Visible = False
        End If

        txtDataInicial.Enabled = True
        txtDataFinal.Enabled = True
        txtLivro.Enabled = True
        txtCodigoDaReceita.Enabled = True
        txtFolha.Enabled = True
        txtVenctoDaObrigacao.Enabled = True
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

#End Region

    Private Sub TotalizaResumo()
        Empresa = DdlEmpresa.SelectedValue.Split("-")

        Dim Total01 As Decimal = 0
        Dim Total02 As Decimal = 0
        Dim Total03 As Decimal = 0
        Dim Total04 As Decimal = 0
        Dim Total05 As Decimal = 0
        Dim Total06 As Decimal = 0
        Dim Total07 As Decimal = 0
        Dim Total08 As Decimal = 0
        Dim Total09 As Decimal = 0
        Dim Total10 As Decimal = 0
        Dim Total11 As Decimal = 0
        Dim Total12 As Decimal = 0
        Dim Total13 As Decimal = 0
        Dim Total14 As Decimal = 0

        Sql = " SELECT 1 AS Ordem, RR.Codigo_Id, DR.Descricao, RR.Valor" & vbCrLf &
              " FROM ResumoRAICMS RR" & vbCrLf &
              " INNER JOIN DescricaoRAICMS DR" & vbCrLf &
              " ON DR.Codigo_Id = RR.Codigo_Id" & vbCrLf &
              " WHERE RR.Empresa_Id = '" & Empresa(0) & "' AND RR.Processo_Id = " & Processo & vbCrLf &
              " UNION" & vbCrLf &
              " SELECT 2 AS Ordem, RR.Codigo_Id, RIR.Descricao, RIR.Valor" & vbCrLf &
              " FROM ResumoItensRAICMS RIR" & vbCrLf &
              " INNER JOIN ResumoRAICMS RR" & vbCrLf &
              " ON RIR.Empresa_Id = RR.Empresa_Id" & vbCrLf &
              " AND RIR.Processo_Id = RR.Processo_Id" & vbCrLf &
              " AND RIR.Codigo_Id = RR.Codigo_Id" & vbCrLf &
              " WHERE RR.Empresa_Id = '" & Empresa(0) & "' AND RR.Processo_Id = " & Processo & vbCrLf &
              " ORDER BY 2, 1" & vbCrLf
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows

            If Dr("Ordem") = 1 And Dr("Codigo_Id") = 2 Then
                Dr("Valor") = 0
            End If
            If Dr("Ordem") = 1 And Dr("Codigo_Id") = 3 Then
                Dr("Valor") = 0
            End If
            If Dr("Ordem") = 1 And Dr("Codigo_Id") = 6 Then
                Dr("Valor") = 0
            End If
            If Dr("Ordem") = 1 And Dr("Codigo_Id") = 7 Then
                Dr("Valor") = 0
            End If
            If Dr("Ordem") = 1 And Dr("Codigo_Id") = 12 Then
                Dr("Valor") = 0
            End If

            Select Case Dr("Codigo_Id")
                Case 1
                    Total01 += Dr("Valor")
                Case 2
                    Total02 += Dr("Valor")
                Case 3
                    Total03 += Dr("Valor")
                Case 4
                    Total04 = Total01 + Total02 + Total03
                    Dr("Valor") = Total04
                Case 5
                    Total05 += Dr("Valor")
                Case 6
                    Total06 += Dr("Valor")
                Case 7
                    Total07 += Dr("Valor")
                Case 8
                    Total08 = Total05 + Total06 + Total07
                    Dr("Valor") = Total08
                Case 9
                    Total09 += Dr("Valor")
                Case 10
                    Total10 = Total08 + Total09
                    Dr("Valor") = Total10
                Case 11
                    If Total04 > Total10 Then
                        Total11 = Total04 - Total10
                        Dr("Valor") = Total11
                    End If
                Case 12
                    Total12 += Dr("Valor")
                Case 13
                    If Total11 > 0 Then
                        Total13 = Total11 - Total12
                        Dr("Valor") = Total13
                    End If
                Case 14
                    If Total10 > Total04 Then
                        Total14 = Total10 - Total04
                        Dr("Valor") = Total14
                    End If
            End Select
        Next

        '---------Atualiza Resumo----------------------------------------
        Sql = "UPDATE ResumoRAICMS SET Valor = " & Microsoft.VisualBasic.Replace(Total02, ",", ".") & " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo & " AND Codigo_Id = 002"
        SqlArray.Add(Sql)
        Sql = "UPDATE ResumoRAICMS SET Valor = " & Microsoft.VisualBasic.Replace(Total03, ",", ".") & " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo & " AND Codigo_Id = 003"
        SqlArray.Add(Sql)
        Sql = "UPDATE ResumoRAICMS SET Valor = " & Microsoft.VisualBasic.Replace(Total04, ",", ".") & " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo & " AND Codigo_Id = 004"
        SqlArray.Add(Sql)
        Sql = "UPDATE ResumoRAICMS SET Valor = " & Microsoft.VisualBasic.Replace(Total06, ",", ".") & " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo & " AND Codigo_Id = 006"
        SqlArray.Add(Sql)
        Sql = "UPDATE ResumoRAICMS SET Valor = " & Microsoft.VisualBasic.Replace(Total07, ",", ".") & " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo & " AND Codigo_Id = 007"
        SqlArray.Add(Sql)
        Sql = "UPDATE ResumoRAICMS SET Valor = " & Microsoft.VisualBasic.Replace(Total08, ",", ".") & " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo & " AND Codigo_Id = 008"
        SqlArray.Add(Sql)
        Sql = "UPDATE ResumoRAICMS SET Valor = " & Microsoft.VisualBasic.Replace(Total09, ",", ".") & " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo & " AND Codigo_Id = 009"
        SqlArray.Add(Sql)
        Sql = "UPDATE ResumoRAICMS SET Valor = " & Microsoft.VisualBasic.Replace(Total10, ",", ".") & " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo & " AND Codigo_Id = 010"
        SqlArray.Add(Sql)
        Sql = "UPDATE ResumoRAICMS SET Valor = " & Microsoft.VisualBasic.Replace(Total11, ",", ".") & " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo & " AND Codigo_Id = 011"
        SqlArray.Add(Sql)
        Sql = "UPDATE ResumoRAICMS SET Valor = " & Microsoft.VisualBasic.Replace(Total12, ",", ".") & " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo & " AND Codigo_Id = 012"
        SqlArray.Add(Sql)
        Sql = "UPDATE ResumoRAICMS SET Valor = " & Microsoft.VisualBasic.Replace(Total13, ",", ".") & " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo & " AND Codigo_Id = 013"
        SqlArray.Add(Sql)
        Sql = "UPDATE ResumoRAICMS SET Valor = " & Microsoft.VisualBasic.Replace(Total14, ",", ".") & " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo & " AND Codigo_Id = 014"
        SqlArray.Add(Sql)

        If Banco.GravaBanco(SqlArray) = False Then
            Mensagem = "Erro na atualização do resumo."
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        End If
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("RegistrosDeICMS", "GRAVAR") Then
                If Validar() Then
                    Empresa = DdlEmpresa.SelectedValue.Split("-")
                    PeriodoInicial = CDate(txtDataInicial.Text).ToString("yyyy/MM/dd")
                    PeriodoFinal = CDate(txtDataFinal.Text).ToString("yyyy/MM/dd")
                    Livro = txtLivro.Text
                    Folha = txtFolha.Text
                    Vencimento = CDate(txtVenctoDaObrigacao.Text).ToString("yyyy/MM/dd")
                    CodigoDaReceita = txtCodigoDaReceita.Text

                    If DdlProcesso.Text = "" Then
                        PegarNumeroDoProcesso()
                    Else
                        Processo = DdlProcesso.SelectedValue
                    End If

                    If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0), Vencimento, CodigoDaReceita) Then
                        If DdlProcesso.Text = "" Then
                            GravaResumoIcms()
                        End If
                        CarregarProcesso()
                        DdlProcesso.SelectedValue = Processo

                        lnkNovo.Parent.Visible = False
                        lnkExcluir.Parent.Visible = False
                        txtDataInicial.Enabled = False
                        txtDataFinal.Enabled = False
                        txtLivro.Enabled = False
                        txtCodigoDaReceita.Enabled = False
                        txtFolha.Enabled = False
                        txtVenctoDaObrigacao.Enabled = False
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            'If Funcoes.VerificaPermissao("RegistrosDeICMS", "EXCLUIR") Then
            '    If Validar() Then
            '        Empresa = DdlEmpresa.SelectedValue.Split("-")
            '        PeriodoInicial = CDate(txtDataInicial.Text).ToString("yyyy/MM/dd")
            '        PeriodoFinal = CDate(txtDataFinal.Text).ToString("yyyy/MM/dd")
            '        Livro = txtLivro.Text
            '        Folha = txtFolha.Text
            '        Vencimento = CDate(txtVenctoDaObrigacao.Text).ToString("yyyy/MM/dd")
            '        CodigoDaReceita = txtCodigoDaReceita.Text

            '        Processo = DdlProcesso.SelectedValue
            '    End If

            '    DeletarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0), Vencimento, CodigoDaReceita)
            'Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            'End If
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
            Funcoes.Ajuda(Me.Page, "RegistrosDeICMS")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class


