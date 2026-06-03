Imports System.IO
Imports System.Data
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Receituario
    Inherits BasePage

    Private objNota As [Lib].Negocio.EmissaoReceituarioNota
    Private objReceitaProduto As [Lib].Negocio.EmissaoReceituarioProduto
    Private ProdutoSelecionado As [Lib].Negocio.EmissaoReceituarioProduto
    Private strJavaScript As String
    Private sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Revenda)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Receituario", "ACESSAR") Then
                    If Not Directory.Exists(Server.MapPath("Receituario")) Then
                        Directory.CreateDirectory(Server.MapPath("Receituario"))
                    End If
                    SessaoVerificaCarregaObjetos()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub SessaoSalvaNota()
        Session("ssNota") = objNota
    End Sub

    Private Sub SessaoRecuperaNota()
        objNota = CType(Session("ssNota"), [Lib].Negocio.EmissaoReceituarioNota)
        If gridItem.SelectedIndex <> -1 Then
            ProdutoSelecionado = objNota.Itens(gridItem.SelectedIndex)
        End If
    End Sub

    Private Sub SessaoVerificaCarregaObjetos()
        '************************************************************************************
        '****************************   Empresa   *******************************************
        '************************************************************************************
        If Not Session("objEmpresaREC") Is Nothing Then
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objEmpresaREC"), [Lib].Negocio.Cliente))
            txtEmpresa.Text = itemEmpresa.Text
            txtCodigoEmpresa.Value = itemEmpresa.Value

            Session.Remove("objEmpresaREC")
        End If
    End Sub

    Private Sub ListarEmpresas()
        Dim objEmpresa As New [Lib].Negocio.Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))
        Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objEmpresa)
        txtEmpresa.Text = itemEmpresa.Text
        txtCodigoEmpresa.Value = itemEmpresa.Value
    End Sub

    Private Function ValidarCamposConsulta() As Boolean
        If txtCodigoEmpresa.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
            Return False
        Else
            Return True
        End If
    End Function

    Private Function ValidarReceitaDoProduto() As Boolean
        If ddlCultura.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Cultura não foi selecionada.")
            Return False
        ElseIf ddlPraga.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Praga não foi selecionada.")
            Return False
        ElseIf ddlTipoDeAplicacao.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Tipo de aplicação não foi selecionada.")
            Return False
        ElseIf txtDosagemRecomendada.Text.Length = 0 OrElse CDec(txtDosagemRecomendada.Text) = 0 Then
            MsgBox(Me.Page, "Dosagem não foi selecionada.")
            Return False
        ElseIf txtNumeroAplicacao.Text.Length = 0 OrElse CInt(txtNumeroAplicacao.Text) = 0 Then
            MsgBox(Me.Page, "Número de aplicações não foi informada.")
            Return False
        ElseIf txtDosagemRecomendada.Text.Length = 0 OrElse CDec(txtDosagemRecomendada.Text) = 0 Then
            MsgBox(Me.Page, "Dosagem recomendada não foi informada.")
        End If

        Return True
    End Function

    Private Sub Limpar()
        LimparReceiraProduto()

        Session.Remove("objReceituario" & HID.Value.ToString)
        HID.Value = Guid.NewGuid().ToString
        ucConsultaEmpresas.SetarHID(HID.Value)

        ImgConsultar.Enabled = True
        ImgLimpar.Enabled = True
        gridNotas.Enabled = True
        radNao.Checked = True
        radSim.Checked = False
        txtNota.Text = String.Empty

        ddlAgronomos.Items.Clear()
        gridNotas.DataBind()
        gridItem.DataBind()

        Session.Remove("objEmpresaREC")
        Session.Remove("objClienteREC")
        Session.Remove("ListaDeNotas")
        Session.Remove("ListaItensDaNota")
        Session.Remove("ReceitaProduto")

        TabReceituario.ActiveTabIndex = 0

        If txtEmpresa.Text = "" Then
            txtDataInicial.Text = Now.Date()
            txtDataFinal.Text = Now.Date()
            ListarEmpresas()
        End If
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            btnEmpresa.Enabled = False
        End If
    End Sub

    Private Sub LimparReceiraProduto()
        gridART.DataBind()
        lblFito.Text = ""
        lblFormulacaoFito.Text = ""
        txtModalidadeDeAplicacao.Text = ""
        lblNomeComercial.Text = ""
        lblNomeTecnico.Text = ""
        lblClasseToxicologica.Text = ""
        lblClasseAmbiental.Text = ""
        lblClasseDeRisco.Text = ""
        txtPrimeiroSocorros.Text = ""
        txtInstrucaoEmbalagem.Text = ""
        txtInstrucaoDeManejo.Text = ""
        txtAdvertenciaMeioAmbiente.Text = ""
        lblLocalDeAplicacao.Text = ""
        ddlCultura.Items.Clear()
        txtAreaPlantada.Text = ""
        ddlPraga.Items.Clear()
        gridDosagem.DataBind()
        ddlTipoDeAplicacao.Items.Clear()
        txtEpocaDeAplicacao.Text = ""
        lblQuantidadeDoProduto.Text = ""
        lblDosagemMinima.Text = ""
        txtDosagemRecomendada.Text = ""
        lblDosagemMaxima.Text = ""
        lblIntervaloDeSeguranca.Text = ""
        lblVazao.Text = ""
        lblAreaTratada.Text = ""
    End Sub

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo") = "Livre"
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me.Page, "objReceituario" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ImgConsultar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If ValidarCamposConsulta() Then
                Dim Empresa() As String = txtCodigoEmpresa.Value.ToString.Split("-")
                Dim ListaDeNotas As New [Lib].Negocio.ListEmissaoReceituarioNota(Empresa(0), Empresa(1), txtDataInicial.Text, txtDataFinal.Text)
                If ListaDeNotas.Count > 0 Then
                    Session("ListaDeNotas") = ListaDeNotas
                    gridNotas.DataSource = ListaDeNotas.ToArray
                    gridNotas.DataBind()
                Else
                    MsgBox(Me.Page, "Período sem movimento.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ImgLimpar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridNotas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            TabReceituario.ActiveTabIndex = 1
            objNota = CType(Session("ListaDeNotas"), [Lib].Negocio.ListEmissaoReceituarioNota)(gridNotas.SelectedIndex)
            If objNota.Itens.Count > 0 Then
                If objNota.Cliente.Complemento.Length = 0 Then
                    MsgBox(Me.Page, "Verifique o cliente da nota fiscal pois está sem o local da aplicação. entre no cadastro de clientes e informe o complemento.")
                Else
                    ddl.Carregar(ddlAgronomos, CarregarDDL.Tabela.ClientesXTipos, "16", True)
                    gridItem.DataSource = objNota.Itens.ToArray
                    gridItem.DataBind()
                    SessaoSalvaNota()
                End If
            Else
                MsgBox(Me.Page, "Itens da nota fiscal com problema.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridItem_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridItem.SelectedIndexChanged
        Try
            SessaoRecuperaNota()

            ProdutoSelecionado = objNota.Itens(gridItem.SelectedIndex)

            If ProdutoSelecionado.Fito.CodigoFito = 0 Then
                MsgBox(Me.Page, "Fito do produto não encontrado, verifique se o código está informado no produto e se foi relacionando com o mesmo na tabela ProdutoXFito. qualquer dúvida entre com o suporte.")
            Else
                lblFito.Text = ProdutoSelecionado.Fito.CodigoIndeaMT
                lblFormulacaoFito.Text = ProdutoSelecionado.Fito.FormulacaoFito.Descricao
                txtModalidadeDeAplicacao.Text = Server.HtmlDecode(ProdutoSelecionado.Fito.ModoAplicacao)
                lblNomeComercial.Text = ProdutoSelecionado.Fito.NomeComercial
                lblNomeTecnico.Text = ProdutoSelecionado.Fito.NomeTecnico
                lblClasseToxicologica.Text = ProdutoSelecionado.Fito.ClasseTox.Descricao
                lblClasseAmbiental.Text = ProdutoSelecionado.Fito.ClasseAmbiental.Descricao
                lblClasseDeRisco.Text = ProdutoSelecionado.Fito.ClasseDeRisco.Descricao
                txtPrimeiroSocorros.Text = ProdutoSelecionado.Fito.PrimeirosSocorros
                txtInstrucaoEmbalagem.Text = ProdutoSelecionado.Fito.DescarteEmbalagem
                txtInstrucaoDeManejo.Text = ProdutoSelecionado.Fito.InstrucoesUso
                txtAdvertenciaMeioAmbiente.Text = ProdutoSelecionado.Fito.MeioAmbiente
                lblLocalDeAplicacao.Text = objNota.Cliente.Complemento

                ddlCultura.DataValueField = "Cultura"
                ddlCultura.DataTextField = "Descricao"
                ddlCultura.DataSource = ProdutoSelecionado.Culturas
                ddlCultura.DataBind()

                Funcoes.InserirLinhaEmBranco(ddlCultura)
                SessaoSalvaNota()
                TabReceituario.ActiveTabIndex = 2
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlCultura_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlCultura.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Cultura não foi selecionada.")
            Else
                SessaoRecuperaNota()
                ProdutoSelecionado.CodigoCultura = ddlCultura.SelectedValue

                If ProdutoSelecionado.Cultura.CodigoIndeaMT = 0 Then
                    ProdutoSelecionado.CodigoCultura = 0
                    ddlCultura.SelectedIndex = 0
                    MsgBox(Me.Page, "Código do indea não está informado na tabela de cultura. entre na tabela e informe o código para prosseguir.")
                Else
                    txtAreaPlantada.Text = ProdutoSelecionado.AreaTotal.ToString("N2")
                    ddlPraga.DataValueField = "Praga"
                    ddlPraga.DataTextField = "NomeComum"
                    ddlPraga.DataSource = ProdutoSelecionado.Pragas
                    ddlPraga.DataBind()
                    Funcoes.InserirLinhaEmBranco(ddlPraga)
                End If
                SessaoSalvaNota()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlPraga_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlPraga.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Praga não foi selecionada.")
            Else
                SessaoRecuperaNota()
                ProdutoSelecionado.CodigoPraga = ddlPraga.SelectedValue

                If ProdutoSelecionado.Praga.CodigoIndea.Length = 0 Then
                    ProdutoSelecionado.CodigoPraga = 0
                    ddlPraga.SelectedIndex = 0
                    MsgBox(Me.Page, "Código do indea não está informado na tabela de praga. entre na tabela e informe o código para prosseguir.")
                Else
                    If ProdutoSelecionado.CodigoCulturaPragaFito > 0 Then
                        txtEpocaDeAplicacao.Text = ProdutoSelecionado.EpocaDeAplicacao
                    End If

                    gridDosagem.DataSource = ProdutoSelecionado.Dosagens.ToArray()
                    gridDosagem.DataBind()
                    ddlTipoDeAplicacao.DataValueField = "FormaDeAplicacao_Id"
                    ddlTipoDeAplicacao.DataTextField = "Descricao"
                    ddlTipoDeAplicacao.DataSource = ProdutoSelecionado.FormasDeAplicacao
                    ddlTipoDeAplicacao.DataBind()
                    Funcoes.InserirLinhaEmBranco(ddlTipoDeAplicacao)
                End If
                SessaoSalvaNota()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlTipoDeAplicacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlTipoDeAplicacao.SelectedIndex = 0 Then
            Else
                SessaoRecuperaNota()
                ProdutoSelecionado.CodigoFormaAplicacao = ddlTipoDeAplicacao.SelectedValue
                txtModalidadeDeAplicacao.Text = Server.HtmlDecode(txtModalidadeDeAplicacao.Text & " " & ProdutoSelecionado.ModoDeAplicacao.Descricao & " " & ProdutoSelecionado.ModoDeAplicacao.Descricao)
                SessaoSalvaNota()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridDosagem_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNota()
            ProdutoSelecionado.CodigoDosagem = ProdutoSelecionado.Dosagens(gridDosagem.SelectedIndex).CodigoDosagem
            lblQuantidadeDoProduto.Text = ProdutoSelecionado.Quantidade
            ProdutoSelecionado.NumeroDeAplicacao = 1
            txtNumeroAplicacao.Text = ProdutoSelecionado.NumeroDeAplicacao
            lblDosagemMinima.Text = ProdutoSelecionado.Dosagens(gridDosagem.SelectedIndex).DosagemMinima
            ProdutoSelecionado.DosagemRecomendada = ProdutoSelecionado.Dosagens(gridDosagem.SelectedIndex).DosagemRecomendada
            txtDosagemRecomendada.Text = ProdutoSelecionado.DosagemRecomendada
            lblDosagemMaxima.Text = ProdutoSelecionado.Dosagens(gridDosagem.SelectedIndex).DosagemMaxima
            lblIntervaloDeSeguranca.Text = ProdutoSelecionado.Dosagens(gridDosagem.SelectedIndex).IntervaloDeSeguranca
            lblVazao.Text = Trim(ProdutoSelecionado.Dosagens(gridDosagem.SelectedIndex).VazaoTerrestre) & Trim(ProdutoSelecionado.Dosagens(gridDosagem.SelectedIndex).VazaoAerea)
            ProdutoSelecionado.AreaTratada = (ProdutoSelecionado.Quantidade / ProdutoSelecionado.DosagemRecomendada) / ProdutoSelecionado.NumeroDeAplicacao
            lblAreaTratada.Text = ProdutoSelecionado.AreaTratada.ToString("N4")
            SessaoSalvaNota()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtNumeroAplicacao_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If txtNumeroAplicacao.Text.Length = 0 OrElse CInt(txtNumeroAplicacao.Text) = 0 Then
                MsgBox(Me.Page, "Número de aplicações não foi informada.")
            Else
                SessaoRecuperaNota()
                ProdutoSelecionado.NumeroDeAplicacao = txtNumeroAplicacao.Text
                ProdutoSelecionado.AreaTratada = (ProdutoSelecionado.Quantidade / ProdutoSelecionado.DosagemRecomendada) / ProdutoSelecionado.NumeroDeAplicacao
                lblAreaTratada.Text = ProdutoSelecionado.AreaTratada.ToString("N4")
                SessaoSalvaNota()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtDosagemRecomendada_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If txtDosagemRecomendada.Text.Length = 0 OrElse CDec(txtDosagemRecomendada.Text) = 0 Then
                MsgBox(Me.Page, "Dosagem recomendada não foi informada.")
            Else
                SessaoRecuperaNota()
                ProdutoSelecionado.DosagemRecomendada = txtDosagemRecomendada.Text
                ProdutoSelecionado.AreaTratada = (ProdutoSelecionado.Quantidade / ProdutoSelecionado.DosagemRecomendada) / ProdutoSelecionado.NumeroDeAplicacao
                lblAreaTratada.Text = ProdutoSelecionado.AreaTratada.ToString("N4")
                SessaoSalvaNota()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgConfirmaReceita_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgConfirmaReceita.Click
        Try
            If Funcoes.VerificaPermissao("Receituario", "GRAVAR") Then
                If ValidarReceitaDoProduto() Then
                    SessaoRecuperaNota()
                    ProdutoSelecionado.Encerrada = True
                    gridItem.DataSource = objNota.Itens.ToArray
                    gridItem.DataBind()
                    SessaoSalvaNota()
                    LimparReceiraProduto()
                    TabReceituario.ActiveTabIndex = 1
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlAgronomos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlAgronomos.SelectedIndex > 0 Then
                Dim Empresa() As String = txtCodigoEmpresa.Value.ToString.Split("-")
                Dim Cliente() As String = ddlAgronomos.SelectedValue.ToString.Split("-")
                SessaoRecuperaNota()
                objNota.CodigoRespTecnico = Cliente(0)
                objNota.EndRespTecnico = Cliente(1)
                SessaoSalvaNota()
                Dim ListaART As New [Lib].Negocio.ListArt(True, Empresa(0), Empresa(1), Cliente(0), Cliente(1), True)
                gridART.DataSource = ListaART.ToArray()
                gridART.DataBind()
            Else
                MsgBox(Me.Page, "Agrônomo não foi selecionado")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridART_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNota()
            objNota.CodigoART = gridART.SelectedRow.Cells(1).Text()
            SessaoSalvaNota()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgEmitirReceituario_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim r As Integer = 0
            SessaoRecuperaNota()
            If ddlAgronomos.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Agrônomo não foi selecionado.")
            ElseIf objNota.CodigoART = 0 Then
                MsgBox(Me.Page, "ART não foi selecionada.")
            Else
                Dim Encerrado As Boolean = True
                For Each row As [Lib].Negocio.EmissaoReceituarioProduto In objNota.Itens
                    If row.Encerrada = False Then
                        Encerrado = False
                    Else
                        r += 1
                    End If
                Next
                If Encerrado = True Or (r = 3 And CInt(gridART.SelectedRow.Cells(4).Text()) = 49) Then
                    If objNota.SalvarReceitas Then
                        SessaoSalvaNota()
                        Relatorio()
                        EmitirReceituario()
                        Limpar()
                    Else
                        MsgBox(Me.Page, objNota.Mensagem.ToString)
                    End If
                Else
                    MsgBox(Me.Page, "Produto sem receita informada. Selecione o Produto pendente e informe a receita.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirReceituario()
        SessaoRecuperaNota()

        Dim NomeArquivo2 As String = "Receituario/ART" & objNota.CodigoART & Funcoes.GeraNomeArquivo & ".html"
        Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
        Dim arquivo As String = NomeArquivo
        Dim linha As String = ""
        Dim strm As StreamWriter = Nothing

        If Dir(arquivo).Length > 0 Then Kill(arquivo)
        linha = "<html>" & vbCrLf
        linha &= "<head>" & vbCrLf
        linha &= "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf
        linha &= "<title>NGS - Anotação de Responsabilidade Técnica</title>" & vbCrLf
        linha &= "<style type='text/css'>" & vbCrLf
        linha &= "H6 {page-break-after:always}"
        linha &= "A.A1 {font-family: Ms serif;font-size: 5pt;line height: 15px}" & vbCrLf
        linha &= "A.A2 {font-family: Ms serif;font-size: 6pt;line height: 15px}" & vbCrLf
        linha &= "A.A3 {font-family: Ms serif;font-size: 7pt;line height: 15px}" & vbCrLf
        linha &= "A.A4 {font-family: Ms serif;font-size: 8pt;line height: 15px}" & vbCrLf
        linha &= "A.A5 {font-family: Ms serif;font-size: 9pt;line height: 15px}" & vbCrLf
        linha &= "A.A6 {font-family: Ms serif;font-size: 10pt;line height: 15px}" & vbCrLf
        linha &= "A.A7 {font-family: Ms serif bold;font-size: 8pt;line height: 15px}" & vbCrLf
        linha &= "A.A8 {font-family: Ms serif bold;font-size: 9pt;line height: 15px}" & vbCrLf
        linha &= "A.A9 {font-family: Ms serif bold;font-size: 10pt;line height: 15px}" & vbCrLf
        linha &= "A.A10 {font-family: Ms serif bold;font-size: 11pt;line height: 15px}" & vbCrLf
        linha &= "A.A11 {font-family: Ms serif bold;font-size: 12pt;line height: 15px}" & vbCrLf
        linha &= "A.A12 {font-family: Ms serif bold;font-size: 13pt;line height: 15px}" & vbCrLf
        linha &= "A.A13 {font-family: Ms serif bold;font-size: 14pt;line height: 15px}" & vbCrLf
        linha &= "</style>" & vbCrLf
        linha &= "</head>" & vbCrLf
        linha &= "<body text=#000000 bgcolor=#FFFFFF>" & vbCrLf

        Dim Receitas As String = ""
        Dim Delimitador As String = ""
        For Each row As [Lib].Negocio.EmissaoReceituarioProduto In objNota.Itens
            If Not Receitas.Contains(row.CodigoReceita) Then
                Receitas += Delimitador & row.CodigoReceita
                Delimitador = "-"
            End If
        Next

        Dim NumeroReceitas() As String = Receitas.Split("-")
        Dim ARTReceita As New Hashtable
        For i As Integer = 0 To NumeroReceitas.GetUpperBound(0)
            If NumeroReceitas(i).ToString.Length > 0 Then
                ARTReceita.Add(i, NumeroReceitas(i))
            End If
        Next

        Dim Abriu As Boolean = False
        Dim Primeiro As Boolean = True

        Dim Empresa() As String = txtCodigoEmpresa.Value.ToString.Split("-")
        Dim objEmpresa As New [Lib].Negocio.Cliente(objNota.CodigoEmpresa, objNota.EndEmpresa)
        Dim Cliente() As String = ddlAgronomos.SelectedValue.ToString.Split("-")
        Dim objCliente As New [Lib].Negocio.Cliente(objNota.CodigoRespTecnico, objNota.EndRespTecnico)

        For j As Integer = 0 To ARTReceita.Count
            For Each row As [Lib].Negocio.EmissaoReceituarioProduto In objNota.Itens
                If ARTReceita(j) = row.CodigoReceita Then
                    If Abriu = False Then
                        Abriu = True

                        If Primeiro = False Then
                            linha &= "<H6>&nbsp;</H6>"
                        End If
                        linha &= "<table border='0' width='100%' cellpadding='0' cellspacing='0'>" & vbCrLf
                        linha &= "<tr>" & vbCrLf
                        linha &= "<td colspan='2'>" & vbCrLf
                        linha &= "<hr />" & vbCrLf
                        linha &= "</td>" & vbCrLf
                        linha &= "</tr>" & vbCrLf
                        linha &= "<tr>" & vbCrLf
                        linha &= "<td align='left'><A Class='A5'><B>" & objEmpresa.Nome & "</B></A></td>" & vbCrLf
                        linha &= "<td align='right'><A Class='A5'><B>CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo) & "</B></A></td>" & vbCrLf
                        linha &= "</tr>" & vbCrLf
                        linha &= "<tr>" & vbCrLf
                        linha &= "<td align='left'><A Class='A5'><B>" & Trim(objEmpresa.Endereco) & "," & objEmpresa.Numero & "</B></A></td>" & vbCrLf
                        linha &= "<td align='right'><A Class='A5'><B>INSCR.EST.: " & objEmpresa.InscricaoEstadual & "</B></A></td>" & vbCrLf
                        linha &= "</tr>" & vbCrLf
                        linha &= "<tr>" & vbCrLf
                        linha &= "<td align='left'><A Class='A5'><B>" & objEmpresa.CEP & " - " & Trim(objEmpresa.Cidade) & "/" & objEmpresa.CodigoEstado & "</B></A></td>" & vbCrLf
                        linha &= "<td align='right'><A Class='A5'><B>FONE: " & Trim(objEmpresa.Telefone) & "</B></A></td>" & vbCrLf
                        linha &= "</tr>" & vbCrLf
                        linha &= "</table>" & vbCrLf

                        linha &= "<table border='0' width='100%' cellpadding='0' cellspacing='0'>" & vbCrLf
                        linha &= "<tr>" & vbCrLf
                        linha &= "<td colspan='3'>" & vbCrLf
                        linha &= "<hr />" & vbCrLf
                        linha &= "</td>" & vbCrLf
                        linha &= "</tr>" & vbCrLf
                        linha &= "<tr>" & vbCrLf
                        linha &= "<td align='left'><A Class='A11'>ART " & row.Nota.CodigoART & "</A></td>" & vbCrLf
                        linha &= "<td align='center'><A Class='A11'>RECEITUÁRIO AGRONÔMICO</A></td>" & vbCrLf
                        linha &= "<td align='right'><A Class='A11'>RECEITA " & row.NumeroReceita & "</A></td>" & vbCrLf
                        linha &= "</tr>" & vbCrLf
                        linha &= "</table>" & vbCrLf

                        linha &= "<table border='0' width='100%'>" & vbCrLf
                        linha &= "<tr>" & vbCrLf
                        linha &= "<td colspan='4'>" & vbCrLf
                        linha &= "<hr />" & vbCrLf
                        linha &= "</td>" & vbCrLf
                        linha &= "</tr>" & vbCrLf
                        linha &= "<tr>" & vbCrLf
                        linha &= "<td><A Class='A4'><B>NOTA: </B>" & row.Nota.Nota & "</A></td>" & vbCrLf
                        linha &= "<td><A Class='A4'><B>CULTURA: </B>" & row.Cultura.Descricao & "(" & row.Cultura.NomeCientifico & ")</A></td>" & vbCrLf
                        linha &= "<td></td>" & vbCrLf
                        linha &= "<td><A Class='A4'><B>DATA RECEITA: </B>" & row.Nota.DataDaNota.ToString("dd/MM/yyyy") & "</A></td>" & vbCrLf
                        linha &= "</tr>" & vbCrLf
                        linha &= "<tr>" & vbCrLf
                        linha &= "<td><A Class='A4'><B>CLIENTE: </B>" & row.Nota.Nome & "</A></td>" & vbCrLf
                        linha &= "<td><A Class='A4'><B>CPF/CNPJ: </B>" & Funcoes.FormatarCpfCnpj(row.Nota.CodigoCliente) & "</A></td>" & vbCrLf
                        linha &= "<td colspan='2'><A Class='A4'><B>CIDADE/UF: </B>" & Trim(row.Nota.Cliente.Cidade) & "/" & row.Nota.Cliente.CodigoEstado & "</A></td>" & vbCrLf
                        linha &= "</tr>" & vbCrLf
                        linha &= "<tr>" & vbCrLf
                        linha &= "<td><A Class='A4'><B>RESP.TÊCNICO: </B>" & objCliente.Nome & "</A></td>" & vbCrLf
                        linha &= "<td><A Class='A4'><B>REG.CREA: </B>" & objCliente.OrgaoRegCategoria & "</A></td>" & vbCrLf
                        linha &= "<td></td>" & vbCrLf
                        linha &= "<td></td>" & vbCrLf
                        linha &= "</tr>" & vbCrLf
                        linha &= "</table>" & vbCrLf
                    End If

                    linha &= "<table border='0' width='100%'>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td colspan='5'>" & vbCrLf
                    linha &= "<hr />" & vbCrLf
                    linha &= "</td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td colspan='5'><A Class='A8'><B>RECOMENDAÇÕES TÊCNICAS</B></A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>Nome Comercial</B></A></td>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>Nome Técnico</B></A></td>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>Classe Toxicológica</B></A></td>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>Quantidade/Dosagem</B></A></td>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>Área Tratada</B></A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'>" & row.Fito.NomeComercial & "</A></td>" & vbCrLf
                    linha &= "<td><A Class='A4'>" & row.Fito.NomeTecnico & "</A></td>" & vbCrLf
                    linha &= "<td><A Class='A4'>" & row.Fito.ClasseTox.Descricao & "</A></td>" & vbCrLf
                    linha &= "<td><A Class='A4'>" & row.Quantidade & " " & row.Unidade & "/" & row.Dosagem.DosagemRecomendada & "</A></td>" & vbCrLf
                    linha &= "<td><A Class='A4'>" & ProdutoSelecionado.AreaTratada.ToString("N4") & " ha</A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>Praga</B></A></td>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>Formulação</B></A></td>" & vbCrLf
                    linha &= "<td colspan='3'><A Class='A4'><B>Vazão</B></A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'>" & row.Praga.NomeComum & "(" & row.Praga.NomeCientifico & ")</A></td>" & vbCrLf
                    linha &= "<td><A Class='A4'>" & row.Fito.FormulacaoFito.Descricao & "</A></td>" & vbCrLf
                    linha &= "<td colspan='3'><A Class='A4'>Aérea:" & row.Dosagem.VazaoAerea & " / Terrestre:" & row.Dosagem.VazaoTerrestre & "</A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>Intervalo de Segurança</B></A></td>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>Tipo de Aplicação</B></A></td>" & vbCrLf
                    linha &= "<td colspan='3'><A Class='A4'><B>Local de Aplicação</B></A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'>" & row.Dosagem.IntervaloDeSeguranca & "</A></td>" & vbCrLf
                    linha &= "<td><A Class='A4'>" & row.FormaDeAplicacao.Descricao & "</A></td>" & vbCrLf
                    linha &= "<td colspan='3'><A Class='A4'>" & row.Nota.Cliente.Complemento & "</A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "</table>" & vbCrLf

                    linha &= "<table border='0' width='100%'>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td>" & vbCrLf
                    linha &= "<hr />" & vbCrLf
                    linha &= "</td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>MODO DE APLICAÇÃO</B></A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'>" & Trim(row.Fito.ModoAplicacao) & " " & Trim(row.ModoDeAplicacao.Descricao) & "</A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td>" & vbCrLf
                    linha &= "<hr />" & vbCrLf
                    linha &= "</td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>ÉPOCA DE APLICAÇÃO</B></A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'>" & Trim(row.EpocaDeAplicacao) & "</A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td>" & vbCrLf
                    linha &= "<hr />" & vbCrLf
                    linha &= "</td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>INSTRUÇÕES DE MANEJO</B></A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'>" & Trim(row.Fito.InstrucoesUso) & "</A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td>" & vbCrLf
                    linha &= "<hr />" & vbCrLf
                    linha &= "</td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>PRIMEIROS SOCORROS NO CASO DE ACIDENTE - </B>" & Trim(row.Fito.PrimeirosSocorros) & "</A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>ADVERTÊNCIAS RELACIONADAS COM A PROTEÇÃO DO MEIO AMBIENTE - </B>" & Trim(row.Fito.MeioAmbiente) & "</A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "<tr>" & vbCrLf
                    linha &= "<td><A Class='A4'><B>INSTRUÇÕES SOBRE A DISPOSIÇÃO FINAL DE SOBRAS DE EMBALAGENS - </B>" & Trim(row.Fito.DescarteEmbalagem) & "</A></td>" & vbCrLf
                    linha &= "</tr>" & vbCrLf
                    linha &= "</table>" & vbCrLf
                End If
            Next

            If Abriu = True Then
                Abriu = False
                linha &= "<table border='0' width='100%'>" & vbCrLf
                linha &= "<tr>" & vbCrLf
                linha &= "<td colspan='2'><A Class='A4'>COMPROMETIMENTO DO TRANSPORTE POR TERCEIROS - Mediante este, assumo as responsabilidades do cumprimento das leis de transporte de produtos perigosos, e afirmo ter sido orientado por esta revenda da correta forma de manuseá-lo, assim a isentando de quaisquer responsabilidade com o descumprimento das mesmas.</A></td>" & vbCrLf
                linha &= "</tr>" & vbCrLf
                linha &= "<tr>" & vbCrLf
                linha &= "<td colspan='2'><A Class='A4'><br />ESTOU CIENTE DAS INFORMAÇÕES TÉCNICAS CONTIDAS NESTA RECEITA E NO SEU ANEXO.</A></td>" & vbCrLf
                linha &= "</tr>" & vbCrLf
                linha &= "<tr>" & vbCrLf
                linha &= "<td colspan='2' align='right'><A Class='A5'><br />PRIMAVERA DO LESTE, " & objNota.DataDaNota.ToLongDateString.ToUpper & ".<br /><br /></A></td>" & vbCrLf
                linha &= "</tr>" & vbCrLf
                linha &= "<tr>" & vbCrLf
                linha &= "<td align='cengter'><A Class='A5'>_______________________________________________</A></td>" & vbCrLf
                linha &= "<td align='cengter'><A Class='A5'>_____________________________________________</A></td>" & vbCrLf
                linha &= "</tr>" & vbCrLf
                linha &= "<tr>" & vbCrLf
                linha &= "<td align='cengter'><A Class='A5'>CLIENTE: " & objNota.Nome & "</A></td>" & vbCrLf
                linha &= "<td align='cengter'><A Class='A5'>ENG. AGRÔNOMO RESP.:" & objCliente.Nome & "</A></td>" & vbCrLf
                linha &= "</tr>" & vbCrLf
                linha &= "<tr>" & vbCrLf
                linha &= "<td align='cengter'><A Class='A5'>CPF/CNPJ: " & Funcoes.FormatarCpfCnpj(objNota.CodigoCliente) & "</A></td>" & vbCrLf
                linha &= "<td align='cengter'><A Class='A5'>CPF: " & Funcoes.FormatarCpfCnpj(objCliente.Codigo) & "</A></td>" & vbCrLf
                linha &= "</tr>" & vbCrLf
                linha &= "</table>" & vbCrLf

                linha &= "<table border='0' width='100%' cellpadding='0' cellspacing='0'>" & vbCrLf
                linha &= "<tr>" & vbCrLf
                linha &= "<td colspan='5'>" & vbCrLf
                linha &= "<hr />" & vbCrLf
                linha &= "</td>" & vbCrLf
                linha &= "<tr>" & vbCrLf
                linha &= "<td><A Class='A4'>1a. VIA USUÁRIO</A></td>" & vbCrLf
                linha &= "<td><A Class='A4'>2a. VIA COMÉRCIO</A></td>" & vbCrLf
                linha &= "<td><A Class='A4'>3a. VIA SECRETARIA DA AGRICULTURA</A></td>" & vbCrLf
                linha &= "<td><A Class='A4'>4a. VIA CREA</A></td>" & vbCrLf
                linha &= "<td><A Class='A4'>5a. ENG. AGRÔNOMO RESP.</A></td>" & vbCrLf
                linha &= "</tr>" & vbCrLf
                linha &= "<tr>" & vbCrLf
                linha &= "<td colspan='5'>" & vbCrLf
                linha &= "<hr />" & vbCrLf
                linha &= "</td>" & vbCrLf
                linha &= "</table>" & vbCrLf
                Primeiro = False
            End If
        Next

        linha &= "</body>" & vbCrLf
        linha &= "</html>"

        Try
            strm = New StreamWriter(arquivo, True)
            strm.WriteLine(linha)
            strm.Close()
            Funcoes.AbrirArquivo(Me.Page, NomeArquivo2)
        Catch ex As Exception
            strm.Close()
            MsgBox(Me.Page, ex.Message.ToString)
        End Try
    End Sub

    Protected Sub Relatorio()
        Dim crpt As New ReportDocument()

        Try
            If Funcoes.VerificaPermissao("Receituario", "RELATORIO") Then
                SessaoRecuperaNota()
                crpt.FileName = Server.MapPath("~/Reports/Cr_ReceituarioAgronomico.rpt")
                crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

                Dim NomeArquivo As String = "Receituario/ART" & objNota.CodigoART & Funcoes.GeraNomeArquivo & ".PDF"
                Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)

                Dim Empresa() As String = txtCodigoEmpresa.Value.ToString.Split("-")
                Dim objEmpresa As New [Lib].Negocio.Cliente(objNota.CodigoEmpresa, objNota.EndEmpresa)
                Dim Cliente() As String = ddlAgronomos.SelectedValue.ToString.Split("-")
                Dim objCliente As New [Lib].Negocio.Cliente(objNota.CodigoRespTecnico, objNota.EndRespTecnico)

                Dim crparametervalues As ParameterValues
                Dim crparameterdiscretevalue As ParameterDiscreteValue
                Dim crparameterfielddefinitions As ParameterFieldDefinitions
                Dim crparameterfielddefinition As ParameterFieldDefinition

                crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

                'Empresa
                crparameterfielddefinition = crparameterfielddefinitions.Item("Empresa")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = objEmpresa.Nome
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("EmpresaEndereco")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = objEmpresa.Endereco
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("EmpresaCEP")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = objEmpresa.CEP
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("EmpresaCidade")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = objEmpresa.Cidade & "/" & objEmpresa.CodigoEstado
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("EmpresaCNPJ")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = objEmpresa.CodigoFormatado
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("EmpresaInscEstadual")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = objEmpresa.InscricaoEstadual
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("EmpresaTelefone")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = objEmpresa.Telefone
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                For Each row As [Lib].Negocio.EmissaoReceituarioProduto In objNota.Itens
                    crparameterfielddefinition = crparameterfielddefinitions.Item("ART")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.Nota.CodigoART
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("Receita")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.NumeroReceita
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    'Nota
                    crparameterfielddefinition = crparameterfielddefinitions.Item("Nota")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.Nota.Nota
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)
                    'Cultura
                    crparameterfielddefinition = crparameterfielddefinitions.Item("Cultura")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.Cultura.Descricao
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    'Data receita
                    crparameterfielddefinition = crparameterfielddefinitions.Item("DataReceita")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.Nota.DataDaNota.ToString("dd/MM/yyyy")
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    'Cliente
                    crparameterfielddefinition = crparameterfielddefinitions.Item("Cliente")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.Nota.Nome
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("ClienteCPFouCNPJ")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = Funcoes.FormatarCpfCnpj(row.Nota.CodigoCliente)
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("ClienteCidade")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = Trim(row.Nota.Cliente.Cidade) & "/" & row.Nota.Cliente.CodigoEstado
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    'Resp. Técnico
                    crparameterfielddefinition = crparameterfielddefinitions.Item("RespTecnico")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = objCliente.Nome
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("CPFouCNPJTecnico")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = Funcoes.FormatarCpfCnpj(objCliente.Codigo)
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)
                    'RECOMENDAÇÕES TÉCNICAS
                    crparameterfielddefinition = crparameterfielddefinitions.Item("NomeComercial")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.Fito.NomeComercial
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("NomeTecnico")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.Fito.NomeTecnico
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("ClasseToxicologica")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.Fito.ClasseTox.Descricao
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("QtdeDosagem")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.Quantidade & " " & row.Unidade & "/" & row.Dosagem.DosagemRecomendada
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("AreaTratada")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = ProdutoSelecionado.AreaTratada.ToString("N4")
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("Praga")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.Praga.NomeComum & "(" & row.Praga.NomeCientifico & ")"
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("Formulacao")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.Fito.FormulacaoFito.Descricao
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("Vazao")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = "Aérea:" & row.Dosagem.VazaoAerea & " / Terrestre:" & row.Dosagem.VazaoTerrestre
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("IntervaloSeguranca")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.Dosagem.IntervaloDeSeguranca
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("TipoAplicacao")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.FormaDeAplicacao.Descricao
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("LocalAplicacao")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = row.Nota.Cliente.Complemento
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("ModoAplicacao")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = Trim(row.Fito.ModoAplicacao) & " " & Trim(row.ModoDeAplicacao.Descricao) &
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("EpocaAplicacao")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = Trim(row.EpocaDeAplicacao) &
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("InstrucaoManejo")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = Trim(row.Fito.InstrucoesUso)
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)
                    'primeiros socorros
                    crparameterfielddefinition = crparameterfielddefinitions.Item("PrimeirosSocorros")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = Trim(row.Fito.PrimeirosSocorros)
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)
                    'Advertências ao Meio Ambiente
                    crparameterfielddefinition = crparameterfielddefinitions.Item("AdvertenciasAmbiente")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = Trim(row.Fito.MeioAmbiente)
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)
                    'Instruções Embalagens
                    crparameterfielddefinition = crparameterfielddefinitions.Item("InstrucoesEmbalagens")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = Trim(row.Fito.DescarteEmbalagem)
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    'CREA
                    crparameterfielddefinition = crparameterfielddefinitions.Item("CREA")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = objCliente.OrgaoRegCategoria
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)
                Next

                ' Ver em PDF
                If Dir(arquivo).Length > 0 Then Kill(arquivo)

                crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)

                If IO.File.Exists(arquivo) Then
                    ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo & "');", True)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("Receituario", "LEITURA") Then
                If ValidarCamposConsulta() Then
                    Dim Empresa() As String = txtCodigoEmpresa.Value.ToString.Split("-")
                    Dim ListaDeNotas As New [Lib].Negocio.ListEmissaoReceituarioNota(Empresa(0), Empresa(1), txtDataInicial.Text, txtDataFinal.Text, radSim.Checked, txtNota.Text)

                    If ListaDeNotas.Count > 0 Then
                        Session("ListaDeNotas") = ListaDeNotas
                        gridNotas.DataSource = ListaDeNotas.ToArray
                        gridNotas.DataBind()
                        If radSim.Checked AndAlso ListaDeNotas.Count > 0 Then
                            objNota = ListaDeNotas(0)
                            SessaoSalvaNota()
                        End If
                    Else
                        MsgBox(Me.Page, "Período sem movimento.")
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objReceituario" & HID.Value.ToString) Is Nothing Then
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objReceituario" & HID.Value.ToString), [Lib].Negocio.Cliente))
            txtEmpresa.Text = itemEmpresa.Text
            txtCodigoEmpresa.Value = itemEmpresa.Value
        End If
    End Sub

    Protected Sub Imprimir()
        Try
            SessaoRecuperaNota()
            ProdutoSelecionado = objNota.Itens(0)
            EmitirReceituario()
            Relatorio()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkImprimir_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Imprimir()
            Limpar()
            MsgBox(Me.Page, "Reimpressão realizada com Sucesso.", eTitulo.Sucess)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridItem_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                If String.IsNullOrWhiteSpace(e.Row.Cells(5).Text) OrElse Convert.ToInt32(e.Row.Cells(5).Text) <= 0 Then
                    Dim lnk As LinkButton = CType(e.Row.FindControl("lnkImprimir"), LinkButton)
                    lnk.Visible = False
                Else
                    Dim btn As ImageButton = CType(e.Row.FindControl("imbSelecionar"), ImageButton)
                    btn.Visible = False
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Receituario")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class