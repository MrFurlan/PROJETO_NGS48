Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Fito
    Inherits BasePage

    Private objFito As [Lib].Negocio.Fito
    Private Mensagem As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Fito", "ACESSAR") Then
                txtNomeComercial.Attributes.Add("onKeyPress", "javascript:if(ValidarTecla(event) == 13){ValidarObservacao(event)}")
                txtNomeTecnico.Attributes.Add("onKeyPress", "javascript:if(ValidarTecla(event) == 13){ValidarObservacao(event)}")
                txtFormulaBruta.Attributes.Add("onKeyPress", "javascript:if(ValidarTecla(event) == 13){ValidarObservacao(event)}")
                txtConcentracao.Attributes.Add("onKeyPress", "javascript:if(ValidarTecla(event) == 13){ValidarObservacao(event)}")
                txtMA.Attributes.Add("onKeyPress", "javascript:if(ValidarTecla(event) == 13){ValidarObservacao(event)}")
                txtOnu.Attributes.Add("onKeyPress", "javascript:if(ValidarTecla(event) == 13){ValidarObservacao(event)}")
                txtModoDeAplicacao.Attributes.Add("onKeyPress", "javascript:if(ValidarTecla(event) == 13){ValidarObservacao(event)}")
                txtUso.Attributes.Add("onKeyPress", "javascript:if(ValidarTecla(event) == 13){ValidarObservacao(event)}")
                txtDescarte.Attributes.Add("onKeyPress", "javascript:if(ValidarTecla(event) == 13){ValidarObservacao(event)}")
                txtPrimeirosSocorros.Attributes.Add("onKeyPress", "javascript:if(ValidarTecla(event) == 13){ValidarObservacao(event)}")
                txtMeioAmbiente.Attributes.Add("onKeyPress", "javascript:if(ValidarTecla(event) == 13){ValidarObservacao(event)}")
                txtIncompatibilidade.Attributes.Add("onKeyPress", "javascript:if(ValidarTecla(event) == 13){ValidarObservacao(event)}")
                CarregarClasseToxicologica()
                carregarClasseAmbiental()
                carregarFormulacaoDoFito()
                carregarIngredianteAtivo()
                carregarClasseDeRisco()
                HabilitarCampos()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuario sem permissao para acessar esse módulo.")
                Exit Sub
            End If
        Else
            If Not Request("__EVENTARGUMENT") Is Nothing Then
                If Request("__EVENTARGUMENT") = "AlterarFito" Then IniciarFito("U")
                If Request("__EVENTARGUMENT") = "ExcluirFito" Then IniciarFito("D")
            End If
        End If
    End Sub

    Protected Sub AtualizarGrid(ByVal Codigo As String, ByVal CodigoFito As String, ByVal Nome As String, ByVal CodigoIndea As String)
        Try
            Dim Lista As New [Lib].Negocio.ListaFito("", IIf(CodigoFito.Length = 0, 0, CodigoFito), Nome, "", CodigoIndea)
            gridFitossanitario.DataSource = Lista.ToArray
            gridFitossanitario.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarClasseToxicologica()
        ddl.Carregar(ddlClasseToxicologica, CarregarDDL.Tabela.ClasseToxicologica, "", True)
    End Sub

    Private Sub carregarClasseAmbiental()
        ddl.Carregar(ddlClasseAmbiental, CarregarDDL.Tabela.ClasseAmbiental, "", True)
    End Sub

    Private Sub carregarFormulacaoDoFito()
        ddl.Carregar(ddlFormulacaoDoFito, CarregarDDL.Tabela.FormulacaoFito, "", True)
    End Sub

    Private Sub carregarIngredianteAtivo()
        ddl.Carregar(ddlIngredienteAtivo, CarregarDDL.Tabela.IngredianteAtivo, "", True)
    End Sub

    Private Sub carregarClasseDeRisco()
        ddl.Carregar(ddlClasseDeRisco, CarregarDDL.Tabela.ClasseDeRisco, "", True)
    End Sub

    Function ValidarCampos() As Boolean
        If txtCodigo.Text.Length = 0 OrElse CInt(txtCodigo.Text) = 0 Then
            Mensagem = "Código fito não foi informado."
            Return False
        ElseIf txtCodigoIndea.Text.Length = 0 Then
            Mensagem = "Código Indea não foi informado."
            Return False
        ElseIf ddlClasseToxicologica.SelectedIndex = 0 Then
            Mensagem = "Classe toxicológica não foi selecionada."
            Return False
        ElseIf ddlClasseAmbiental.SelectedIndex = 0 Then
            Mensagem = "Classe ambiental não foi selecionada."
            Return False
        ElseIf ddlFormulacaoDoFito.SelectedIndex = 0 Then
            Mensagem = "Formulação do fito não foi selecionado."
            Return False
        ElseIf ddlIngredienteAtivo.SelectedIndex = 0 Then
            Mensagem = "Ingrediente ativo não foi selecionado."
            Return False
        ElseIf ddlClasseDeRisco.SelectedIndex = 0 Then
            Mensagem = "Classe de risco não foi selecionada."
            Return False
        ElseIf txtNomeComercial.Text.Length = 0 Then
            Mensagem = "Nome comercial não foi informado."
            Return False
        ElseIf txtNomeTecnico.Text.Length = 0 Then
            Mensagem = "Nome técnico não foi informado."
            Return False
        ElseIf txtConcentracao.Text.Length = 0 Then
            Mensagem = "Concentração não foi informada."
            Return False
        ElseIf txtMA.Text.Length = 0 Then
            Mensagem = "Código MA não foi informado."
            Return False
        ElseIf txtModoDeAplicacao.Text.Length = 0 Then
            Mensagem = "Modo de aplicação não foi informado."
            Return False
        ElseIf txtUso.Text.Length = 0 Then
            Mensagem = "Instruções de uso não foi informado."
            Return False
        ElseIf txtDescarte.Text.Length = 0 Then
            Mensagem = "Descarte da embalagem não foi informada."
            Return False
        ElseIf txtPrimeirosSocorros.Text.Length = 0 Then
            Mensagem = "Primeiros socorros não foi informado."
            Return False
        ElseIf txtMeioAmbiente.Text.Length = 0 Then
            Mensagem = "Precauções com o meio ambiente não foi informado."
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub IniciarFito(ByVal Tipo As String)
        Try
            If ValidarCampos() OrElse Tipo.Equals("D") Then
                RecuperarFito()

                objFito.CodigoFito = txtCodigo.Text
                objFito.CodigoIndeaMT = txtCodigoIndea.Text
                objFito.CodigoClasseTox = ddlClasseToxicologica.SelectedValue
                objFito.CodigoClasseAmbiental = ddlClasseAmbiental.SelectedValue
                objFito.CodigoFormulacaoFito = ddlFormulacaoDoFito.SelectedValue
                objFito.CodigoIA = ddlIngredienteAtivo.SelectedValue
                objFito.CodigoClasseDeRisco = ddlClasseDeRisco.SelectedValue
                objFito.NomeComercial = txtNomeComercial.Text.Replace("'", "")
                objFito.NomeTecnico = txtNomeTecnico.Text.Replace("'", "")
                objFito.FormulaBruta = txtFormulaBruta.Text.Replace("'", "")
                objFito.ConcentracaoING = txtConcentracao.Text.Replace("'", "")
                objFito.UFRestricao = txtUFRestricao.Text
                If radInflamavelSim.Checked Then
                    objFito.Inflamavel = "S"
                Else
                    objFito.Inflamavel = "N"
                End If
                If radCorrosivoSim.Checked Then
                    objFito.Corrosivo = "S"
                Else
                    objFito.Corrosivo = "N"
                End If
                objFito.RegistroMA = txtMA.Text
                objFito.RegistroONU = txtOnu.Text
                objFito.ModoAplicacao = Left(txtModoDeAplicacao.Text.Replace("'", ""), 500)
                objFito.InstrucoesUso = Left(txtUso.Text.Replace("'", ""), 500)
                objFito.DescarteEmbalagem = Left(txtDescarte.Text.Replace("'", ""), 500)
                objFito.PrimeirosSocorros = Left(txtPrimeirosSocorros.Text.Replace("'", ""), 500)
                objFito.MeioAmbiente = Left(txtMeioAmbiente.Text.Replace("'", ""), 500)
                objFito.Incompatibilidade = Left(txtIncompatibilidade.Text.Replace("'", ""), 500)
                objFito.IUD = Tipo

                If objFito.Salvar Then
                    Limpar()
                    Select Case Tipo
                        Case "I"
                            MsgBox(Me.Page, "Registro incluído com Sucesso.", eTitulo.Sucess)
                        Case "U"
                            MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                        Case "D"
                            MsgBox(Me.Page, "Registro excluído com Sucesso.", eTitulo.Sucess)
                    End Select
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                End If
            Else
                MsgBox(Me.Page, Mensagem)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub HabilitarCampos()
        txtCodigo.Enabled = True
        txtCodigoIndea.Enabled = True
        ddlClasseToxicologica.Enabled = True
        ddlClasseAmbiental.Enabled = True
        ddlFormulacaoDoFito.Enabled = True
        ddlIngredienteAtivo.Enabled = True
        ddlClasseDeRisco.Enabled = True
        txtNomeComercial.Enabled = True
        txtNomeTecnico.Enabled = True
        txtFormulaBruta.Enabled = True
        txtConcentracao.Enabled = True
        txtUFRestricao.Enabled = True

        radInflamavelSim.Enabled = True
        radInflamavelNao.Enabled = True
        radCorrosivoSim.Enabled = True
        radCorrosivoNao.Enabled = True

        txtMA.Enabled = True
        txtOnu.Enabled = True
        txtModoDeAplicacao.Enabled = True
        txtUso.Enabled = True
        txtDescarte.Enabled = True
        txtPrimeirosSocorros.Enabled = True
        txtMeioAmbiente.Enabled = True
        txtIncompatibilidade.Enabled = True
    End Sub

    Private Sub Limpar()
        lnkNovo.Parent.Visible = True
        lnkExcluir.Parent.Visible = False
        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        lnkLimpar.Parent.Visible = True

        txtCodigo.Enabled = True
        txtCodigo.Text = ""
        txtCodigoIndea.Text = ""
        ddlClasseToxicologica.SelectedIndex = 0
        ddlClasseAmbiental.SelectedIndex = 0
        ddlFormulacaoDoFito.SelectedIndex = 0
        ddlIngredienteAtivo.SelectedIndex = 0
        ddlClasseDeRisco.SelectedIndex = 0
        txtNomeComercial.Text = ""
        txtNomeTecnico.Text = ""
        txtFormulaBruta.Text = ""
        txtConcentracao.Text = ""
        txtUFRestricao.Text = ""

        radInflamavelSim.Checked = False
        radInflamavelNao.Checked = False
        radCorrosivoSim.Checked = False
        radCorrosivoNao.Checked = False

        txtMA.Text = ""
        txtOnu.Text = ""
        txtModoDeAplicacao.Text = ""
        txtUso.Text = ""
        txtDescarte.Text = ""
        txtPrimeirosSocorros.Text = ""
        txtMeioAmbiente.Text = ""
        txtIncompatibilidade.Text = ""

        gridFitossanitario.DataBind()

        objFito = New [Lib].Negocio.Fito()
        objFito.VerSequencia()
        txtCodigo.Text = objFito.CodigoFito

        SalvarFito()
    End Sub

    Private Sub SalvarFito()
        Session("objFito") = objFito
    End Sub

    Private Sub RecuperarFito()
        objFito = CType(Session("objFito"), [Lib].Negocio.Fito)
    End Sub

    Protected Sub gridFitossanitario_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If gridFitossanitario.Rows.Count > 0 Then
                objFito = New [Lib].Negocio.Fito("", gridFitossanitario.SelectedRow.Cells(1).Text())
                txtCodigo.Text = objFito.CodigoFito
                txtCodigoIndea.Text = objFito.CodigoIndeaMT
                ddlClasseToxicologica.SelectedValue = objFito.CodigoClasseTox
                ddlClasseAmbiental.SelectedValue = objFito.CodigoClasseAmbiental
                ddlFormulacaoDoFito.SelectedValue = objFito.CodigoFormulacaoFito
                ddlIngredienteAtivo.SelectedValue = objFito.CodigoIA
                ddlClasseDeRisco.SelectedValue = objFito.CodigoClasseDeRisco
                txtNomeComercial.Text = objFito.NomeComercial
                txtNomeTecnico.Text = objFito.NomeTecnico
                txtFormulaBruta.Text = objFito.FormulaBruta
                txtConcentracao.Text = objFito.ConcentracaoING
                txtUFRestricao.Text = objFito.UFRestricao
                If objFito.Inflamavel = "S" Then
                    radInflamavelSim.Checked = True
                Else
                    radInflamavelNao.Checked = True
                End If
                If objFito.Corrosivo = "S" Then
                    radCorrosivoSim.Checked = True
                Else
                    radCorrosivoNao.Checked = True
                End If
                txtMA.Text = objFito.RegistroMA
                txtOnu.Text = objFito.RegistroONU
                txtModoDeAplicacao.Text = objFito.ModoAplicacao
                txtUso.Text = objFito.InstrucoesUso
                txtDescarte.Text = objFito.DescarteEmbalagem
                txtPrimeirosSocorros.Text = objFito.PrimeirosSocorros
                txtMeioAmbiente.Text = objFito.MeioAmbiente
                txtIncompatibilidade.Text = objFito.Incompatibilidade

                txtCodigo.Enabled = False
                lnkNovo.Parent.Visible = False
                lnkAtualizar.Parent.Visible = True
                lnkExcluir.Parent.Visible = True

                SalvarFito()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Fito", "GRAVAR") Then
                IniciarFito("I")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            If txtCodigo.Text.Length > 0 Or txtCodigoIndea.Text.Length > 0 Or txtNomeComercial.Text.Length > 0 Then
                AtualizarGrid("", txtCodigo.Text, txtNomeComercial.Text, txtCodigoIndea.Text)
                TabContainer1.ActiveTabIndex = 1
            Else
                MsgBox(Me.Page, "Informe o código do fito, códido do indea ou nome comercial.")
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

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("Fito", "ALTERAR") Then
                IniciarFito("U")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Fito", "EXCLUIR") Then
                IniciarFito("D")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Fito")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class