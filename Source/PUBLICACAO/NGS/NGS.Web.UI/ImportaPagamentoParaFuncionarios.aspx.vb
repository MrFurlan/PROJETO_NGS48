Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports NGS.Lib.Negocio
Imports System.Data.SqlTypes
Imports System.Drawing
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Uteis
Imports Microsoft.Win32
Imports NGS.Lib.Negocio.Novo

Public Class ImportaPagamentoParaFuncionarios
    Inherits BasePage

    Private Mensagem As String
    Dim SqlArray As New ArrayList
    Dim strSQL As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ImportaPagamentoParaFuncionarios", "ACESSAR") Then

                Limpar()

                LimparCancelamento()

                BuncarUnidadeDeNegocio()

                txtMovimnto.Text = Format(Today, "dd/MM/yyyy")

                ddlTiposDeFolha.DataTextField = "Descricao"
                ddlTiposDeFolha.DataValueField = "Codigo"
                ddlTiposDeFolha.Items.Add(New ListItem("FOLHA MENSAL", "001003011;SALARIOS;PAGTO FOLHA MENSAL REF. MM/AAAA"))
                ddlTiposDeFolha.Items.Add(New ListItem("FOLHA MENSAL - PRO LABORE", "001003011;PROLABORE;PAGTO FOLHA MENSAL – PRO LABORE REF. MM/AAAA"))
                ddlTiposDeFolha.Items.Add(New ListItem("ADIANTAMENTO", "001001019;ADTO SALARIO;PAGTO FOLHA ADIANTAMENTO REF. MM/AAAA"))
                ddlTiposDeFolha.Items.Add(New ListItem("COMPLEMENTAR", "001003011;SALARIOS;PAGTO FOLHA COMPLEMENTAR REF. MM/AAAA"))
                ddlTiposDeFolha.Items.Add(New ListItem("PARTICIPAÇÃO DE LUCROS", "001003011;SALARIOS;PAGTO FOLHA PLR REF. MM/AAAA"))
                ddlTiposDeFolha.Items.Add(New ListItem("13º ADIANTAMENTO", "001003017;ADTO 13 SAL;PAGTO FOLHA 13º ADIANTAMENTO REF."))
                ddlTiposDeFolha.Items.Add(New ListItem("13º INTEGRAL", "001003011;SALARIOS;PAGTO FOLHA 13º INTEGRAL REF. AAAA"))
                Funcoes.InserirLinhaEmBranco(ddlTiposDeFolha)

            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub BuncarUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)

        ddl.Carregar(ddlUnidadeNegocio2, CarregarDDL.Tabela.UnidadeDeNegocio)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio2, ddlEmpresa2)
    End Sub

    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeNegocio2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(ddlEmpresa2, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio2.SelectedValue.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()

        lnkNovo.Parent.Visible = False
        lnkImportar.Parent.Visible = True

        Session.Remove("objFolha" & HID.Value)

        HID.Value = Guid.NewGuid().ToString

        Dim dtFolha As New DataTable("ItemFolha")

        dtFolha.Columns.Add("UnidadeDeNegocio", Type.GetType("System.String"))
        dtFolha.Columns.Add("Empresa", Type.GetType("System.String"))
        dtFolha.Columns.Add("EndEmpresa", Type.GetType("System.String"))
        dtFolha.Columns.Add("Movimento", Type.GetType("System.String"))
        dtFolha.Columns.Add("Carteira", Type.GetType("System.String"))
        dtFolha.Columns.Add("Encargo", Type.GetType("System.String"))
        dtFolha.Columns.Add("Historico", Type.GetType("System.String"))
        dtFolha.Columns.Add("Unidade", Type.GetType("System.String"))
        dtFolha.Columns.Add("CPF", Type.GetType("System.String"))
        dtFolha.Columns.Add("Nome", Type.GetType("System.String"))
        dtFolha.Columns.Add("Banco", Type.GetType("System.String"))
        dtFolha.Columns.Add("Agencia", Type.GetType("System.String"))
        dtFolha.Columns.Add("AgenciaDigito", Type.GetType("System.String"))
        dtFolha.Columns.Add("Conta", Type.GetType("System.String"))
        dtFolha.Columns.Add("ContaDigito", Type.GetType("System.String"))

        dtFolha.Columns.Add("Valor", Type.GetType("System.Decimal"))

        Session("objFolha" & HID.Value) = dtFolha

        ddlTiposDeFolha.SelectedIndex = 0

        rowTotal.Visible = False
        lblTotal.Text = String.Empty

        gridFuncionarios.DataSource = Nothing
        gridFuncionarios.DataBind()

    End Sub

    Private Sub LimparCancelamento()
        lnkCancelar.Parent.Visible = False
        lnkConsultarImportados.Parent.Visible = True
        gridImportados.DataSource = Nothing
        gridImportados.DataBind()
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            CriarTitulos()
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message))
        End Try
    End Sub

    Protected Sub lnkImportar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkImportar.Click
        Try
            Importar()
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message))
        End Try
    End Sub

    Private Sub Importar()

        If String.IsNullOrWhiteSpace(filUpload.PostedFile.FileName) Then
            MsgBox(Me.Page, "Arquivo não foi selecionado.", eTitulo.Info)
            Exit Sub
        End If

        Dim extArquivo As String() = filUpload.PostedFile.FileName.Split(".")

        If extArquivo(1).ToUpper = "XLSX" Then
            'LIBERA
        Else
            MsgBox(Me.Page, "Extensão do Arquivo deve ser XLSX, caso tenha dúvida entre em contato com o Suporte.", eTitulo.Info)
            Exit Sub
        End If

        Dim infoarquivo As New IO.FileInfo(filUpload.PostedFile.FileName)
        If Upload(infoarquivo) Then
            If ValidaDados(infoarquivo) Then
                ValidarArquivo(infoarquivo)
            Else
                MsgBox(Me.Page, Mensagem.ToString, eTitulo.Erro)
            End If
        Else
            MsgBox(Me.Page, Mensagem.ToString, eTitulo.Erro)
        End If
    End Sub

    Private Function Upload(ByVal infoarquivo As Object) As Boolean
        Try

            'Verificamos se tem alguma coisa postada 
            If Not IsNothing(filUpload.PostedFile) Then
                'Pegamos as informacoes do arquivo postado 
                'Definimos onde ele será salvo 
                Dim strCaminho As String = Server.MapPath("Files/") & infoarquivo.Name

                If File.Exists(strCaminho) Then File.Delete(strCaminho)

                'Salvamos o mesmo 
                filUpload.PostedFile.SaveAs(strCaminho)
            Else
                MsgBox(Me.Page, "Selecione um arquivo.", eTitulo.Info)
                Return False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

        Return True
    End Function

    Public Function ValidaDados(ByVal infoarquivo As Object) As Boolean
        Try
            If ddlEmpresa.SelectedIndex = 0 Then
                Mensagem = "Empresa não foi selecionanda."
                Return False
            ElseIf txtCompetencia.Text = "" Then
                Mensagem = "Competência da Folha não foi informada."
                Return False
            ElseIf infoarquivo.Name.ToString.Length = 0 Then
                Mensagem = "Arquivo não foi informado."
                Return False
            ElseIf ddlTiposDeFolha.SelectedIndex = 0 Then
                Mensagem = "Tipo da Folha não foi selecionado."
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Function

    Private Sub ValidarArquivo(ByVal infoarquivo As Object)
        If infoarquivo Is Nothing OrElse String.IsNullOrWhiteSpace(infoarquivo.Name) Then
            MsgBox(Me.Page, "Arquivo inválido ou não foi encontrado.", eTitulo.Erro)
            Exit Sub
        End If

        CType(Session("objFolha" & HID.Value), DataTable).Rows.Clear()

        Dim caminhoArquivoExcel As String = Server.MapPath("Files/") & infoarquivo.Name
        'Dim nomePlanilhaExcel As String = "Planilha1" & "$"
        Dim nomePlanilhaExcel As String = "Relatório de Líquidos" & "$"

        Dim Valor As Decimal = 0
        Dim Quantidade As Integer = 0

        Try

            Dim arrEmpresa() As String
            arrEmpresa = ddlEmpresa.SelectedValue.Split("-")

            Dim arrCarteira() As String
            arrCarteira = ddlTiposDeFolha.SelectedValue.Split(";")

            Dim emp As New [Lib].Negocio.Cliente(arrEmpresa(0), arrEmpresa(1))

            Dim ds As DataSet
            ds = FuncoesStrings.LerExcelParaDataSet(caminhoArquivoExcel)

            SqlArray.Clear()

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                For Each row As DataRow In ds.Tables(0).Rows

                    If IsDBNull(row("CPF")) Then
                        MsgBox(Me.Page, "Quantidade de caracteres informados não pode diferente de 11(onze) dígitos. NOME: " & row("Nome do empregado"), eTitulo.Info)
                        Exit Sub
                    End If

                    If String.IsNullOrWhiteSpace(row("CPF")) OrElse Not row("CPF").ToString.Length = 11 Then
                        MsgBox(Me.Page, "Quantidade de caracteres informados não pode diferente de 11(onze) dígitos. CPF: " & row("CPF") & " - " & row("Nome do empregado"), eTitulo.Info)
                        Exit Sub
                    End If

                    Dim cli As New [Lib].Negocio.Cliente(row("CPF"), 0)

                    If String.IsNullOrWhiteSpace(cli.Nome) Then
                        MsgBox(Me.Page, "Verifique o funcionário no Sistema, não foi encontrado o cadastro. CPF: " & row("CPF") & " - " & row("Nome do empregado"), eTitulo.Info)
                        Exit Sub
                    End If

                    Dim temBanco As Boolean = False
                    Dim codBanco As String = String.Empty
                    Dim codAgencia As String = String.Empty
                    Dim digitoAgencia As String = String.Empty
                    Dim codConta As String = String.Empty
                    Dim digitoConta As String = String.Empty

                    For Each conta In cli.ContasBancarias
                        If conta.CodigoBanco = 237 AndAlso conta.CodigoAgencia > 0 AndAlso conta.ContaCorrente.Length > 0 Then
                            temBanco = True
                            codBanco = conta.CodigoBanco
                            codAgencia = conta.CodigoAgencia
                            digitoAgencia = conta.DigitoAgencia
                            codConta = conta.ContaCorrente
                            digitoConta = conta.DigitoConta

                            Exit For
                        End If
                    Next

                    If IsDBNull(row("VALOR")) Then
                        MsgBox(Me.Page, "Valor não pode ser zeros. NOME: " & row("Nome do empregado"), eTitulo.Info)
                        Exit Sub
                    End If

                    If String.IsNullOrWhiteSpace(row("VALOR")) OrElse CDec(row("VALOR")) = 0 Then
                        MsgBox(Me.Page, "Valor não pode ser zeros. NOME: " & row("Nome do empregado"), eTitulo.Info)
                        Exit Sub
                    End If

                    Dim drF As DataRow = CType(Session("objFolha" & HID.Value), DataTable).NewRow()

                    drF("UnidadeDeNegocio") = ddlUnidadeNegocio.SelectedValue

                    If row("Empresa") = emp.Codigo Then
                        drF("Empresa") = emp.Codigo
                    Else
                        MsgBox(Me.Page, "Empresa selecionada diferente do informado no arquivo.", eTitulo.Info)
                        Exit Sub
                    End If

                    drF("EndEmpresa") = emp.CodigoEndereco
                    drF("Movimento") = txtMovimnto.Text
                    drF("Carteira") = arrCarteira(0)
                    drF("Encargo") = arrCarteira(1)

                    Dim Historico As String = arrCarteira(2)

                    If arrCarteira(2).Contains("MM/AAAA") Then
                        Historico = Historico.Replace("MM/AAAA", CDate(txtCompetencia.Text).ToString("MM/yyyy"))
                    ElseIf arrCarteira(2).Contains("AAAA") Then
                        Historico = Historico.Replace("AAAA", CDate(txtCompetencia.Text).ToString("yyyy"))
                    End If

                    drF("Unidade") = emp.Cidade & "/" & emp.CodigoEstado
                    drF("CPF") = cli.Codigo & "-" & cli.CodigoEndereco
                    drF("Nome") = row("Nome do empregado")
                    drF("Valor") = row("Valor")

                    If codConta.Length > 0 Then
                        drF("Banco") = codBanco
                        drF("Agencia") = codAgencia
                        drF("AgenciaDigito") = digitoAgencia
                        drF("Conta") = codConta
                        drF("ContaDigito") = digitoConta
                    Else
                        If Left(Session("ssEmpresa"), 8) = "62780383" Then
                            'SE FOR HORUS NAO FAZER CRÍTICA
                            drF("Banco") = ""
                            drF("Agencia") = ""
                            drF("AgenciaDigito") = ""
                            drF("Conta") = ""
                            drF("ContaDigito") = ""
                        Else
                            drF("Conta") = "CADASTRAR"
                            drF("Historico") = "CONTA NÃO CADASTRADA. FAVOR CADASTRAR PARA SEGUIR COM A IMPORTAÇÃO."
                        End If
                    End If

                    drF("Historico") = Historico & " - " & row("Nome do empregado")

                    Valor += row("Valor")

                    Quantidade += 1

                    CType(Session("objFolha" & HID.Value), DataTable).Rows.Add(drF)

                Next

                gridFuncionarios.DataSource = CType(Session("objFolha" & HID.Value), DataTable)
                gridFuncionarios.DataBind()

                Dim i As Integer = 0
                While i < gridFuncionarios.Rows.Count

                    If gridFuncionarios.Rows(i).Cells(10).Text = "CADASTRAR" Then
                        gridFuncionarios.Rows(i).Cells(0).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(1).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(2).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(3).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(4).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(5).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(6).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(7).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(8).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(9).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(10).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(11).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(12).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(13).ForeColor = System.Drawing.Color.Red
                        gridFuncionarios.Rows(i).Cells(6).ToolTip = "ATENÇÃO, FAVOR VERIFICAR"
                    End If

                    i += 1
                End While

                rowTotal.Visible = True

                lnkImportar.Parent.Visible = False
                lnkNovo.Parent.Visible = True

                lblTotal.Text = "Total de Funcionários: " & Quantidade & ". Valor total da folha: " & String.Format("{0:N2}", Valor)

            Else
                MsgBox(Me.Page, "Lista não encontrada no arquivo.", eTitulo.Erro)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try

    End Sub

    Private Sub CriarTitulos()

        Dim ListaDeTitulos As New List(Of Titulo)
        Dim tituloNovo As Titulo = Nothing
        Dim arrSQL As New ArrayList()

        Dim faltaConta As Boolean = False

        For Each drItem As DataRow In CType(Session("objFolha" & HID.Value), DataTable).Rows
            If drItem("Conta") = "CADASTRAR" Then

                MsgBox(Me.Page, "Funcionário " & drItem("CPF").ToString.Split("-")(0) & "-" & drItem("Nome") & " não tem conta Cadastrada. Favor ajustar e refazer o processo de importação.", eTitulo.Erro)

                faltaConta = True
                Exit For

            End If
        Next

        If faltaConta Then
            Exit Sub
        End If

        For Each drItem As DataRow In CType(Session("objFolha" & HID.Value), DataTable).Rows
            tituloNovo = New Titulo()

            tituloNovo.IUD = "I"
            tituloNovo.CodigoProvisao = 2
            tituloNovo.CodigoCarteira = drItem("Carteira")
            tituloNovo.Tributo = drItem("Encargo")

            If Left(Session("ssEmpresa"), 8) = "62780383" Then
                tituloNovo.CodigoTipoPgto = 7
            Else
                tituloNovo.CodigoTipoPgto = 3
            End If

            If Left(Session("ssEmpresa"), 8) = "62780383" Then
                tituloNovo.CodigoSituacao = 1 'NORMAL
                tituloNovo.Situacao.Codigo = eSituacao.Normal
            Else
                tituloNovo.CodigoSituacao = 105 'REMESSA BANCARIA FUNCIONARIOS
                tituloNovo.Situacao.Codigo = eSituacao.RemessaBancariaFuncionarios
            End If

            tituloNovo.Lote = 70
            tituloNovo.CodigoMoeda = 1 'REAL
            tituloNovo.Moeda.Classificacao = eTiposMoeda.Oficial
            tituloNovo.Movimento = CDate(drItem("Movimento"))
            tituloNovo.Vencimento = CDate(drItem("Movimento"))
            tituloNovo.Prorrogacao = CDate(drItem("Movimento"))
            tituloNovo.DataMoeda = CDate(drItem("Movimento"))
            tituloNovo.Baixa = CDate(drItem("Movimento"))
            tituloNovo.CodigoIndexador = 99
            tituloNovo.CodigoUnidadeDeNegocio = drItem("UnidadeDeNegocio")
            tituloNovo.CodigoEmpresa = drItem("Empresa")
            tituloNovo.EnderecoEmpresa = drItem("EndEmpresa")

            tituloNovo.CodigoCliente = drItem("CPF").ToString.Split("-")(0)
            tituloNovo.EndCliente = drItem("CPF").ToString.Split("-")(1)

            If drItem("Banco").ToString.Length > 0 Then
                tituloNovo.CodigoBancoCliente = drItem("Banco")
                tituloNovo.CodigoAgenciaCliente = drItem("Agencia")
                tituloNovo.DigitoAgenciaCliente = drItem("AgenciaDigito")
                tituloNovo.ContaCliente = drItem("Conta")
                tituloNovo.DigitoContaCliente = drItem("ContaDigito")
            End If

            tituloNovo.ValorDoDocumento = drItem("Valor")
            tituloNovo.Deducoes = 0
            tituloNovo.Juros = 0
            tituloNovo.Descontos = 0
            tituloNovo.Acrescimos = 0
            tituloNovo.Historico = drItem("Historico")
            tituloNovo.CodigoDigitado = False
            tituloNovo.Cheque = False
            tituloNovo.Slips = False
            tituloNovo.Recibo = False
            tituloNovo.Aviso = False
            tituloNovo.ReciboDeposito = False
            tituloNovo.CodigoDestinatario = drItem("CPF").ToString.Split("-")(0)
            tituloNovo.EndDestinatario = drItem("CPF").ToString.Split("-")(1)
            tituloNovo.UsuarioInclusao = Session("ssNomeUsuario")
            tituloNovo.UsuarioInclusaoData = Date.Now
            tituloNovo.Agrupado = eAgrupamentoFinanceiro.NaoAgrupado

            If Left(Session("ssEmpresa"), 8) = "62780383" Then
                tituloNovo.FolhaDePagamento = True
            Else
                tituloNovo.FolhaDePagamento = False
            End If

            ListaDeTitulos.Add(tituloNovo)
        Next

        If ListaDeTitulos.Count > 0 Then
            Dim Banco As New AcessaBanco()
            Dim Sql As String

            For Each tit In ListaDeTitulos
                Sql = "exec sp_Numerador '" & Session("ssNomeServidor").ToUpper() & "',0,1"

                For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Numerador").Tables(0).Rows
                    tit.Codigo = Dr("Sequencia")
                Next

                tit.SalvarSql(arrSQL, False)
            Next

            If arrSQL.Count > 0 Then
                If Banco.GravaBanco(arrSQL) Then
                    MsgBox(Me.Page, "Títulos gravados com sucesso.", eTitulo.Sucess)
                    Limpar()
                    LimparCancelamento()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            End If
        End If

    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ImportaPagamentoParaFuncionarios")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkCancelar_Click(sender As Object, e As EventArgs) Handles lnkCancelar.Click
        Try
            Dim temCancelamento As Boolean = False

            For Each rowgrid As GridViewRow In gridImportados.Rows
                Dim chkTitulo As CheckBox = CType(rowgrid.FindControl("chkCancelar"), CheckBox)

                If chkTitulo.Checked Then
                    temCancelamento = True
                    Exit For
                End If
            Next

            If temCancelamento Then

                Dim arrSQL As New ArrayList()

                For Each rowgrid As GridViewRow In gridImportados.Rows
                    Dim chkTitulo As CheckBox = CType(rowgrid.FindControl("chkCancelar"), CheckBox)

                    If chkTitulo.Checked Then

                        Dim Tit As New Titulo(rowgrid.Cells(1).Text)

                        Tit.IUD = "C"
                        Tit.UsuarioCancelamento = Session("ssNomeUsuario")
                        Tit.UsuarioCancelamentoData = Date.Now

                        Dim obs As String = Tit.ObservacoesControleInterno
                        If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                            obs = obs & ". Titulo Cancelado no Processo ImportaPagamentoParaFuncionarios em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & Session("ssNomeUsuario")
                        Else
                            obs = "Titulo Cancelado no Processo ImportaPagamentoParaFuncionarios em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & Session("ssNomeUsuario")
                        End If

                        Tit.ObservacoesControleInterno = obs

                        Tit.SalvarSql(arrSQL)

                    End If
                Next

                If arrSQL.Count > 0 Then

                    If Banco.GravaBanco(arrSQL) Then
                        MsgBox(Me.Page, "Título(s) cancelado(s) com sucesso.", eTitulo.Sucess)
                        LimparCancelamento()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If

                End If

            Else
                MsgBox(Me.Page, "Sem registros para cancelamento.", eTitulo.Info)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkConsultarImportados_Click(sender As Object, e As EventArgs) Handles lnkConsultarImportados.Click
        Try
            If ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionanda.", eTitulo.Info)
            Else
                Dim where As String

                If Left(Session("ssEmpresa"), 8) = "62780383" Then
                    where = "Empresa = " & ddlEmpresa2.SelectedValue.Split("-")(0) & " AND EndEmpresa = " & ddlEmpresa2.SelectedValue.Split("-")(1) & " AND Situacao = " & eSituacao.Normal & " AND FolhaDePagamento = 1 AND Provisao = 2 "
                Else
                    where = "Empresa = " & ddlEmpresa2.SelectedValue.Split("-")(0) & " AND EndEmpresa = " & ddlEmpresa2.SelectedValue.Split("-")(1) & " AND Situacao = " & eSituacao.RemessaBancariaFuncionarios & " AND FolhaDePagamento = 0"
                End If

                Dim titulos As New ListTitulo(where)

                If titulos.Count > 0 Then
                    gridImportados.DataSource = titulos

                    lnkCancelar.Parent.Visible = True
                    lnkConsultarImportados.Parent.Visible = False
                Else
                    MsgBox(Me.Page, "Sem registros para consulta.", eTitulo.Info)
                    gridImportados.DataSource = Nothing

                    lnkCancelar.Parent.Visible = False
                    lnkConsultarImportados.Parent.Visible = True
                End If

                gridImportados.DataBind()

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub chkCancelarAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If gridImportados.Rows.Count > 0 Then
                Dim chkCancelarAll As CheckBox = CType(sender, CheckBox)

                For Each rowgrid As GridViewRow In gridImportados.Rows
                    Dim chkCancelar As CheckBox = CType(rowgrid.FindControl("chkCancelar"), CheckBox)
                    chkCancelar.Checked = chkCancelarAll.Checked
                Next
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparImportados_Click(sender As Object, e As EventArgs) Handles lnkLimparImportados.Click
        Try
            LimparCancelamento()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

End Class